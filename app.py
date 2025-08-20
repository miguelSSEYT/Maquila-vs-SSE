import io
from datetime import date
from typing import Dict, List, Tuple

import pandas as pd
import streamlit as st

st.set_page_config(page_title="Analizador MB52 / COOIS / ZCO41", layout="wide")
st.title("Analizador de Demanda vs Inventario")
st.caption(
    "Sube ZCO41, MB52, COOIS y CrossReference. "
    "La hoja de **todos** los archivos debe llamarse **Sheet1**."
)

# ──────────────────────────────────────────────────────────────────────────────
# Columnas mínimas
REQ_ZCO   = ["Pln.Or Qty", "Estimated Ship Date", "Sales Order", "Material description"]
REQ_MB52  = ["Open Quantity", "Material description"]
REQ_COOIS = ["Order quantity (GMEIN)", "Material description", "Sales Order", "Est. Ship Date"]
REQ_XREF  = ["Custom", "Non Custom"]

def normcols(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df.columns = [str(c).strip() for c in df.columns]
    return df

def need_cols(df: pd.DataFrame, cols: List[str]) -> List[str]:
    return [c for c in cols if c not in df.columns]

def load_sheet1(uploaded_file) -> pd.DataFrame:
    # Estricto: siempre hoja "Sheet1"
    return pd.read_excel(uploaded_file, sheet_name="Sheet1")

def build_xref_map(xref: pd.DataFrame) -> Dict[str, str]:
    xref = normcols(xref)
    xref["Custom"] = xref["Custom"].astype(str).str.strip()
    xref["Non Custom"] = xref["Non Custom"].astype(str).str.strip()
    return dict(zip(xref["Custom"], xref["Non Custom"]))

def map_custom_to_non(series: pd.Series, xref_map: Dict[str, str]) -> pd.Series:
    return series.astype(str).str.strip().map(xref_map)

# ──────────────────────────────────────────────────────────────────────────────
# Preparación de insumos

def prep_mb52(mb52: pd.DataFrame) -> pd.DataFrame:
    mb52 = normcols(mb52)
    missing = need_cols(mb52, REQ_MB52)
    if missing:
        raise ValueError(f"MB52 incompleto, faltan columnas: {missing}")
    mb52 = mb52.rename(columns={"Open Quantity":"OpenQty","Material description":"Non Custom"})
    mb52["OpenQty"] = pd.to_numeric(mb52["OpenQty"], errors="coerce").fillna(0.0)
    mb52 = mb52.groupby("Non Custom", as_index=False)["OpenQty"].sum()
    return mb52

def prep_coois(coois: pd.DataFrame, xref_map: Dict[str, str]):
    coois = normcols(coois)
    missing = need_cols(coois, REQ_COOIS)
    if missing:
        raise ValueError(f"COOIS incompleto, faltan columnas: {missing}")
    coois_use = coois.rename(columns={
        "Order quantity (GMEIN)":"Qty",
        "Material description":"Custom",
        "Est. Ship Date":"EstShip"
    })[["Qty","Custom","Sales Order","EstShip"]]
    coois_use["Qty"] = pd.to_numeric(coois_use["Qty"], errors="coerce").fillna(0.0)
    coois_use["Non Custom"] = map_custom_to_non(coois_use["Custom"], xref_map)
    coois_use["EstShip"] = pd.to_datetime(coois_use["EstShip"], errors="coerce")

    # Demanda agregada por SKU para restarla a MB52
    coois_demand = (coois_use.dropna(subset=["Non Custom"])
                    .groupby("Non Custom", as_index=False)["Qty"].sum()
                    .rename(columns={"Qty":"CooisQty"}))
    # Para alertas
    coois_unmapped = (coois_use[coois_use["Non Custom"].isna()]["Custom"]
                      .drop_duplicates().sort_values().tolist())
    return coois_use, coois_unmapped, coois_demand

def prep_zco41(zco41: pd.DataFrame, xref_map: Dict[str, str]) -> Tuple[pd.DataFrame, List[str]]:
    zco = normcols(zco41)
    missing = need_cols(zco, REQ_ZCO)
    if missing:
        raise ValueError(f"ZCO41 incompleto, faltan columnas: {missing}")
    zco = zco.rename(columns={
        "Pln.Or Qty":"Qty",
        "Estimated Ship Date":"EstShip",
        "Material description":"Custom"
    })[["Qty","EstShip","Sales Order","Custom"]]
    zco["Qty"] = pd.to_numeric(zco["Qty"], errors="coerce").fillna(0.0)
    zco["Non Custom"] = map_custom_to_non(zco["Custom"], xref_map)
    zco["EstShip"] = pd.to_datetime(zco["EstShip"], errors="coerce")

    zco_unmapped = (zco[zco["Non Custom"].isna()]["Custom"]
                    .drop_duplicates().sort_values().tolist())
    return zco, zco_unmapped

# ──────────────────────────────────────────────────────────────────────────────
# Motor de evaluación

def evaluate_orders(zco_lines: pd.DataFrame, inv_after: pd.DataFrame):
    """
    Lógica:
      - Ordeno ZCO41 por fecha y SO.
      - Una orden pasa solo si TODAS sus líneas tienen inventario suficiente.
      - Si pasa, descuento inventario (acumulativo). Si no pasa, no descuento.
    """
    avail_dict = dict(zip(inv_after["Non Custom"], inv_after["Avail"]))
    zco_sorted = zco_lines.sort_values(["EstShip", "Sales Order", "Custom"], kind="mergesort")

    orders_out, lines_out = [], []

    for so, grp in zco_sorted.groupby("Sales Order", sort=False):
        order_ok = True
        checks = []

        for _, r in grp.iterrows():
            sku = r["Non Custom"]
            qty = float(r["Qty"])
            if pd.isna(sku):
                checks.append((r, False, qty, "Sin mapeo a Non Custom"))
                order_ok = False
                continue
            avail = float(avail_dict.get(sku, 0.0))
            if avail >= qty:
                checks.append((r, True, 0.0, "Suficiente inventario, pasar"))
            else:
                falt = max(0.0, qty - avail)
                checks.append((r, False, falt, f"No suficiente; faltan {falt:.0f}"))
                order_ok = False

        if order_ok:
            # Descontar inventario por cada línea
            for (r, ok, falt, note) in checks:
                sku = r["Non Custom"]; qty = float(r["Qty"])
                avail_dict[sku] = avail_dict.get(sku, 0.0) - qty

        orders_out.append({
            "Sales Order": so,
            "Est. Ship min": grp["EstShip"].min(),
            "Total líneas": len(grp),
            "Status": "PASAR" if order_ok else "NO PASAR",
        })

        for (r, ok, falt, note) in checks:
            lines_out.append({
                "Sales Order": r["Sales Order"],
                "Est. Ship": r["EstShip"],
                "Custom": r["Custom"],
                "Non Custom": r["Non Custom"],
                "Qty": float(r["Qty"]),
                "Nota": note,
                "Shortage": float(falt),
                "Resultado línea": "OK" if ok else "Insuficiente",
            })

    orders_df = pd.DataFrame(orders_out).sort_values(["Est. Ship min", "Sales Order"])
    lines_df  = pd.DataFrame(lines_out ).sort_values(["Est. Ship", "Sales Order", "Custom"])
    final_avail_df = pd.DataFrame(
        list(avail_dict.items()), columns=["Non Custom","Avail_after_Approvals"]
    ).sort_values("Non Custom")

    return orders_df, lines_df, final_avail_df

# ──────────────────────────────────────────────────────────────────────────────
# Salidas adicionales

def build_inventario_necesito(orders_df: pd.DataFrame, lines_df: pd.DataFrame, inv_after: pd.DataFrame) -> pd.DataFrame:
    # Faltantes de órdenes NO PASAR (usando columna numérica Shortage)
    det = lines_df.merge(orders_df[["Sales Order","Status"]], on="Sales Order", how="left")
    falt = det[(det["Status"] == "NO PASAR") & (det["Shortage"] > 0) & (det["Non Custom"].notna())]
    falt_group = falt.groupby("Non Custom", as_index=False)["Shortage"].sum().rename(columns={"Shortage":"NeededQty"})

    # Negativos de inventario base (MB52 - COOIS)
    neg = inv_after[inv_after["Avail"] < 0][["Non Custom","Avail"]].copy()
    neg["NeededQty"] = neg["Avail"].abs()
    neg = neg.drop(columns=["Avail"])

    out = (pd.concat([falt_group, neg], ignore_index=True)
           .groupby("Non Custom", as_index=False)["NeededQty"].sum())
    out = out[out["NeededQty"] > 0].sort_values("NeededQty", ascending=False)
    return out

def build_past_due_zco(lines_df: pd.DataFrame, today: pd.Timestamp) -> pd.DataFrame:
    past = lines_df.copy()
    past = past[past["Est. Ship"].notna() & (past["Est. Ship"] < today)]
    return past.sort_values(["Est. Ship","Sales Order","Custom"])

def build_past_due_coois(coois_use: pd.DataFrame, inv_after: pd.DataFrame,
                         xref_map: Dict[str, str], today: pd.Timestamp) -> pd.DataFrame:
    base = coois_use.rename(columns={"EstShip":"Est. Ship"})[["Sales Order","Est. Ship","Custom","Qty"]].copy()
    base["Non Custom"] = map_custom_to_non(base["Custom"], xref_map)
    base["Est. Ship"] = pd.to_datetime(base["Est. Ship"], errors="coerce")
    past = base[base["Est. Ship"].notna() & (base["Est. Ship"] < today)].copy()

    # Disponibilidad agregada por SKU (MB52 - COOIS)
    sku_avail = inv_after[["Non Custom","Avail"]]
    past = past.merge(sku_avail, on="Non Custom", how="left")

    def nota_result(avail):
        if pd.isna(avail):
            avail = 0.0  # si no aparece en MB52, tratamos como 0 disponible
        if avail >= 0:
            return "Suficiente inventario, pasar", "OK", 0.0
        else:
            return f"No suficiente; faltan {abs(int(avail))}", "Insuficiente", float(abs(avail))

    notas, resultados, shortages = zip(*past["Avail"].map(nota_result))
    past["Nota"] = notas
    past["Resultado línea"] = resultados
    past["Shortage"] = shortages

    past = past[["Sales Order","Est. Ship","Custom","Non Custom","Qty","Nota","Resultado línea","Shortage"]]
    return past.sort_values(["Est. Ship","Sales Order","Custom"])

def to_excel_bytes(sheets: Dict[str, pd.DataFrame]) -> bytes:
    bio = io.BytesIO()
    with pd.ExcelWriter(bio, engine="xlsxwriter") as writer:
        for name, df in sheets.items():
            df.to_excel(writer, sheet_name=name, index=False)
    bio.seek(0)
    return bio.read()

# ──────────────────────────────────────────────────────────────────────────────
# UI: carga de archivos (solo .xlsx para evitar xlrd)
with st.sidebar:
    st.header("Carga de archivos (Sheet1)")
    zco41_file = st.file_uploader("ZCO41 (demanda nueva)", type=["xlsx"])
    mb52_file  = st.file_uploader("MB52 (inventario)", type=["xlsx"])
    coois_file = st.file_uploader("COOIS (demanda fija)", type=["xlsx"])
    xref_file  = st.file_uploader("CrossReference (Custom ↔ Non Custom)", type=["xlsx"])

    st.markdown("---")
    today = pd.Timestamp(date.today())
    st.caption(f"Hoy: **{today.date()}**")

if not all([zco41_file, mb52_file, coois_file, xref_file]):
    st.info("Sube los cuatro archivos en el panel lateral para procesar.")
    st.stop()

# ──────────────────────────────────────────────────────────────────────────────
# Procesamiento principal
try:
    zco41_df = load_sheet1(zco41_file); zco41_df = normcols(zco41_df)
    mb52_df  = load_sheet1(mb52_file);  mb52_df  = normcols(mb52_df)
    coois_df = load_sheet1(coois_file); coois_df = normcols(coois_df)
    xref_df  = load_sheet1(xref_file);  xref_df  = normcols(xref_df)
except Exception as e:
    st.error(f"No pude leer 'Sheet1' de alguno de los archivos: {e}")
    st.stop()

# Validación de columnas mínimas
missing_msgs = []
for name, df, req in [("ZCO41", zco41_df, REQ_ZCO), ("MB52", mb52_df, REQ_MB52),
                      ("COOIS", coois_df, REQ_COOIS), ("CrossReference", xref_df, REQ_XREF)]:
    miss = need_cols(df, req)
    if miss:
        missing_msgs.append(f"- {name}: faltan columnas {miss}")
if missing_msgs:
    st.error("Columnas faltantes:\n" + "\n".join(missing_msgs))
    st.stop()

# Mapeo
xref_map = build_xref_map(xref_df)

# MB52
mb52_inv = prep_mb52(mb52_df)

# COOIS
coois_use, coois_unmapped, coois_demand = prep_coois(coois_df, xref_map)

# Inventario base = MB52 - COOIS (agregado por SKU)
inv_after = mb52_inv.merge(coois_demand, on="Non Custom", how="left")
inv_after["CooisQty"] = inv_after["CooisQty"].fillna(0.0)
inv_after["Avail"] = inv_after["OpenQty"] - inv_after["CooisQty"]

# ZCO41
zco_use, zco_unmapped = prep_zco41(zco41_df, xref_map)

# Evaluación de órdenes ZCO41
orders_df, lines_df, final_avail_df = evaluate_orders(zco_use, inv_after)

# Inventario que necesito
inventario_necesito = build_inventario_necesito(orders_df, lines_df, inv_after)

# Past due
past_due_zco   = build_past_due_zco(lines_df, today)
past_due_coois = build_past_due_coois(coois_use, inv_after, xref_map, today)

# ──────────────────────────────────────────────────────────────────────────────
# UI de resultados
tabs = st.tabs([
    "Resumen & Descarga",
    "Detalle_Lineas",
    "Inventario que necesito",
    "Past due ZCO41",
    "Past due COOIS",
    "Alertas de mapeo"
])

with tabs[0]:
    st.subheader("Resumen de órdenes")
    st.dataframe(orders_df, use_container_width=True, hide_index=True)

    excel_bytes = to_excel_bytes({
        "Resumen_Ordenes": orders_df,
        "Detalle_Lineas": lines_df,
        "Inventario que necesito": inventario_necesito,
        "Past due ZCO41": past_due_zco,
        "Past due COOIS": past_due_coois
    })
    st.download_button(
        label="⬇️ Descargar Excel (5 hojas)",
        data=excel_bytes,
        file_name="resultado_final_v2_streamlit.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True
    )

with tabs[1]:
    st.subheader("Detalle por línea (con notas)")
    st.dataframe(lines_df, use_container_width=True, hide_index=True)

with tabs[2]:
    st.subheader("Inventario que necesito")
    st.dataframe(inventario_necesito, use_container_width=True, hide_index=True)

with tabs[3]:
    st.subheader("Past due ZCO41")
    st.dataframe(past_due_zco, use_container_width=True, hide_index=True)

with tabs[4]:
    st.subheader("Past due COOIS")
    st.dataframe(past_due_coois, use_container_width=True, hide_index=True)

with tabs[5]:
    st.subheader("Alertas de mapeo (Custom sin equivalencia)")
    col1, col2 = st.columns(2)
    with col1:
        st.markdown("**ZCO41 (Custom sin mapeo)**")
        st.dataframe(pd.DataFrame({"Custom": zco_unmapped}), use_container_width=True, hide_index=True)
    with col2:
        st.markdown("**COOIS (Custom sin mapeo)**")
        st.dataframe(pd.DataFrame({"Custom": coois_unmapped}), use_container_width=True, hide_index=True)

st.success("Listo. Resultados generados con la lógica acordada.")
