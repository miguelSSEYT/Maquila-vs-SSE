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
    uso_detallado = []

    for idx, row in df_ci.iterrows():
        descripcion = row["DES NO CUSTOM"]
        cantidad = int(row["Delivery Quantity"])
        tracking = row["Tracking Number"]
        documento = row["Document"]
        material = row["Material"] if "Material" in row else ""

        posibles = df_minuta[df_minuta["Descripción"] == descripcion]

        if posibles.empty:
            resultado.append({
                "Tracking Number": tracking,
                "Document": documento,
                "Material": material,
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
                "Material": material,
                "Descripción": descripcion,
                "Cantidad Usada": cantidad,
                "Delivery": selected["Delivery"],
                "Precio Unitario": selected["Precio Unitario"]
            })
            uso_detallado.append({"Delivery": selected["Delivery"], "Cantidad": cantidad})
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
                    "Material": material,
                    "Descripción": descripcion,
                    "Cantidad Usada": usar,
                    "Delivery": m_row["Delivery"],
                    "Precio Unitario": m_row["Precio Unitario"],
                    "Comentario": "Fraccionado" if cantidad > usar else ""
                })

                uso_detallado.append({"Delivery": m_row["Delivery"], "Cantidad": usar})

                if restante == 0:
                    break

            if restante > 0:
                resultado.append({
                    "Tracking Number": tracking,
                    "Document": documento,
                    "Material": material,
                    "Descripción": descripcion,
                    "Cantidad Usada": restante,
                    "Delivery": "",
                    "Precio Unitario": "",
                    "Comentario": "Vaso de maquila (incompleto)"
                })

    df_resultado = pd.DataFrame(resultado)
    df_minuta_actualizada = df_minuta.copy()
    return df_resultado, df_minuta_actualizada

if excel_file:
    xls = pd.ExcelFile(excel_file)
    if "Minuta" in xls.sheet_names and "CI" in xls.sheet_names:
        df_minuta = pd.read_excel(xls, sheet_name="Minuta")
        df_ci = pd.read_excel(xls, sheet_name="CI")

        st.success("Archivo cargado correctamente. Procesando...")
        resultado, minuta_actualizada = procesar_fifo(df_minuta, df_ci)

        st.subheader("Resultado del Análisis FIFO")
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

        # Mostrar minuta actualizada
        st.subheader("Minuta Actualizada")
        st.dataframe(minuta_actualizada)

        output_minuta = BytesIO()
        minuta_actualizada.to_excel(output_minuta, index=False)
        st.download_button(
            label="📥 Descargar Minuta Actualizada",
            data=output_minuta.getvalue(),
            file_name="Minuta_Actualizada.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    else:
        st.error("El archivo debe contener las hojas 'Minuta' y 'CI'")
