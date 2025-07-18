import streamlit as st
import pandas as pd
from io import BytesIO

st.set_page_config(page_title="Análisis FIFO - CI vs Minuta", layout="wide")
st.title("📦 Análisis FIFO entre CI y Minuta")

st.sidebar.header("Subir archivo Excel")
excel_file = st.sidebar.file_uploader("Carga tu archivo Excel con hojas 'Minuta' y 'CI'", type=["xlsx"])

def procesar_fifo(df_minuta, df_ci):
    df_minuta["Fecha"] = pd.to_datetime(df_minuta["Fecha"], errors='coerce')
    df_minuta = df_minuta.sort_values(by=["Descripción", "Fecha"]).reset_index(drop=True)
    df_minuta["Saldo Pdte"] = pd.to_numeric(df_minuta["Saldo Pdte"], errors="coerce").fillna(0).astype(int)

    resultado = []

    for idx, row in df_ci.iterrows():
        descripcion = row["DES NO CUSTOM"]
        cantidad = int(row["Delivery Quantity"])
        tracking = row["Tracking Number"]
        documento = row["Document"]

        posibles = df_minuta[df_minuta["Descripción"] == descripcion]

        if posibles.empty:
            resultado.append({
                "Tracking Number": tracking,
                "Document": documento,
                "Descripción": descripcion,
                "Cantidad Usada": cantidad,
                "Delivery": "",
                "Precio Unitario": "",
                "Comentario": "Vaso de maquila"
            })
            continue

        match_completo = posibles[posibles["Saldo Pdte"] >= cantidad]
        if not match_completo.empty:
            selected = match_completo.iloc[0]
            i = selected.name
            df_minuta.at[i, "Saldo Pdte"] -= cantidad

            resultado.append({
                "Tracking Number": tracking,
                "Document": documento,
                "Descripción": descripcion,
                "Cantidad Usada": cantidad,
                "Delivery": selected["Delivery"],
                "Precio Unitario": selected["Precio Unitario"]
            })
        else:
            restante = cantidad
            for i, m_row in posibles.iterrows():
                saldo = df_minuta.at[i, "Saldo Pdte"]
                if saldo <= 0:
                    continue
                usar = min(saldo, restante)
                if usar <= 0:
                    continue

                df_minuta.at[i, "Saldo Pdte"] -= usar
                restante -= usar

                resultado.append({
                    "Tracking Number": tracking,
                    "Document": documento,
                    "Descripción": descripcion,
                    "Cantidad Usada": usar,
                    "Delivery": m_row["Delivery"],
                    "Precio Unitario": m_row["Precio Unitario"],
                    "Comentario": "Fraccionado" if cantidad > usar else ""
                })

                if restante == 0:
                    break

            if restante > 0:
                resultado.append({
                    "Tracking Number": tracking,
                    "Document": documento,
                    "Descripción": descripcion,
                    "Cantidad Usada": restante,
                    "Delivery": "",
                    "Precio Unitario": "",
                    "Comentario": "Vaso de maquila (incompleto)"
                })

    return pd.DataFrame(resultado)

if excel_file:
    xls = pd.ExcelFile(excel_file)
    if "Minuta" in xls.sheet_names and "CI" in xls.sheet_names:
        df_minuta = pd.read_excel(xls, sheet_name="Minuta")
        df_ci = pd.read_excel(xls, sheet_name="CI")

        st.success("Archivo cargado correctamente. Procesando...")
        resultado = procesar_fifo(df_minuta, df_ci)

        st.dataframe(resultado)

        # Descarga del resultado
        output = BytesIO()
        resultado.to_excel(output, index=False)
        st.download_button(
            label="📥 Descargar Resultado en Excel",
            data=output.getvalue(),
            file_name="Resultado_FIFO.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    else:
        st.error("El archivo debe contener las hojas 'Minuta' y 'CI'")
