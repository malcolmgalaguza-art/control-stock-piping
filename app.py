import streamlit as st
import pandas as pd
import numpy as np
import re, unicodedata
from io import BytesIO
from datetime import datetime

st.set_page_config(page_title="Control de Stock Piping", page_icon="📦", layout="wide")
CENTROS = [3013, 3117, 3118, 3119, 3120, 3121]

def norm(v):
    s = unicodedata.normalize("NFKD", str(v).strip().upper())
    return re.sub(r"\s+", " ", "".join(c for c in s if not unicodedata.combining(c)))

def sap_clean(v):
    if pd.isna(v): return ""
    s = str(v).strip()
    return s[:-2] if re.fullmatch(r"\d+\.0", s) else s

def num(v):
    if pd.isna(v): return 0.0
    if isinstance(v, (int, float, np.number)): return float(v)
    s = str(v).strip().replace(" ", "")
    if s in ("", "-", "—"): return 0.0
    if "," in s and "." in s:
        s = s.replace(".", "").replace(",", ".") if s.rfind(",") > s.rfind(".") else s.replace(",", "")
    elif "," in s: s = s.replace(",", ".")
    try: return float(s)
    except ValueError: return 0.0

def read_stock(file):
    df = pd.read_excel(file)
    cols = {norm(c): c for c in df.columns}
    needed = ["MATERIAL", "TEXTO BREVE DE MATERIAL", "CENTRO", "LIBRE UTILIZACION"]
    missing = [c for c in needed if c not in cols]
    if missing: raise ValueError(f"Faltan columnas en stock SAP: {missing}")
    df = df.rename(columns={
        cols["MATERIAL"]: "SAP", cols["TEXTO BREVE DE MATERIAL"]: "Descripción",
        cols["CENTRO"]: "Centro", cols["LIBRE UTILIZACION"]: "Cantidad"
    })
    df["SAP"] = df["SAP"].map(sap_clean)
    df["Centro"] = pd.to_numeric(df["Centro"], errors="coerce")
    df["Cantidad"] = pd.to_numeric(df["Cantidad"], errors="coerce").fillna(0)
    df = df[df["Centro"].isin(CENTROS) & (df["SAP"] != "")]
    return df.groupby(["SAP", "Descripción", "Centro"], as_index=False)["Cantidad"].sum()

def find_center(file, sheet):
    sample = pd.read_excel(file, sheet_name=sheet, header=None, nrows=8)
    for v in sample.astype(object).values.ravel():
        if pd.isna(v): continue
        for c in CENTROS:
            if re.search(rf"(?<!\d){c}(?!\d)", str(v)): return c
    return None

def pending_col(columns):
    names = [norm(c) for c in columns]
    for group in [
        ["RESTAN ENVIAR", "RESTA ENVIAR", "RESTAENVIAR", "RESTAN EVIAR"],
        ["SALDO PENDIENTE"], ["CANTIDAD PENDIENTE", "PENDIENTE"]
    ]:
        for i, name in enumerate(names):
            if name in group: return i, name, "CONFIRMAR"
    for i, name in enumerate(names):
        if name in ["ENVIAR", "ENVIAR A INTERACERO", "ENVIAR A TYCSA", "ENVIAR A TAPIGAR"]:
            return i, name, "REVISAR COLUMNA"
    return None, None, "SIN COLUMNA"

def find_header(file, sheet):
    sample = pd.read_excel(file, sheet_name=sheet, header=None, nrows=25)
    for row_idx, row in sample.iterrows():
        vals = [norm(v) if pd.notna(v) else "" for v in row]
        if "SAP" in vals and "DESCRIPCION" in vals:
            _, _, quality = pending_col(vals)
            if quality != "SIN COLUMNA": return row_idx
    return None

def process(stock_file, package_files):
    stock = read_stock(stock_file)
    stock_dict = stock.groupby(["SAP", "Centro"])["Cantidad"].sum().to_dict()
    descs = stock.drop_duplicates("SAP").set_index("SAP")["Descripción"].to_dict()
    rows, issues = [], []
    for file in package_files:
        file.seek(0)
        try: book = pd.ExcelFile(file)
        except Exception as e:
            issues.append({"Archivo": file.name, "Hoja": "", "Motivo": str(e)})
            continue
        for sheet in book.sheet_names:
            if any(x in norm(sheet) for x in ["STOCK", "AVANCE", "RESUMEN", "TABLA6"]): continue
            try:
                file.seek(0)
                header = find_header(file, sheet)
                if header is None: continue
                file.seek(0)
                df = pd.read_excel(file, sheet_name=sheet, header=header)
                df.columns = [str(c).strip() for c in df.columns]
                names = [norm(c) for c in df.columns]
                if "SAP" not in names or "DESCRIPCION" not in names: continue
                sap_idx, desc_idx = names.index("SAP"), names.index("DESCRIPCION")
                p_idx, p_name, quality = pending_col(df.columns)
                if p_idx is None:
                    issues.append({"Archivo": file.name, "Hoja": sheet, "Motivo": "Sin columna pendiente"})
                    continue
                file.seek(0)
                center = find_center(file, sheet)
                if center is None:
                    issues.append({"Archivo": file.name, "Hoja": sheet, "Motivo": "No se detectó centro destino"})
                    continue
                workshop = next((x for x in ["TAPIGAR", "ARMEXAS", "INTERACERO", "TYCSA"] if x in norm(file.name)), "OTRO")
                for _, row in df.iterrows():
                    code = sap_clean(row.iloc[sap_idx])
                    if not code or code.upper() in ["NAN", "SAP"]: continue
                    pending = num(row.iloc[p_idx])
                    description = str(row.iloc[desc_idx]).strip() if pd.notna(row.iloc[desc_idx]) else descs.get(code, "")
                    local_stock = max(float(stock_dict.get((code, center), 0)), 0)
                    pending_pos = max(pending, 0)
                    suggested = min(pending_pos, local_stock) if quality == "CONFIRMAR" else 0
                    rows.append({
                        "Taller": workshop, "Paquete": sheet, "Centro destino": int(center),
                        "SAP": code, "Descripción": description, "Columna pendiente": p_name,
                        "Pendiente original": pending, "Pendiente": pending_pos,
                        "Stock local SAP": local_stock, "Cantidad sugerida enviar local": suggested,
                        "Pendiente luego de stock local": max(pending_pos-suggested, 0),
                        "Validación": quality,
                        "Estado": "REVISAR COLUMNA" if quality != "CONFIRMAR" else ("ENVIAR DESDE STOCK LOCAL" if suggested > 0 else "SIN STOCK LOCAL")
                    })
            except Exception as e:
                issues.append({"Archivo": file.name, "Hoja": sheet, "Motivo": str(e)})
    report = pd.DataFrame(rows)
    if report.empty: raise ValueError("No se encontraron tablas de paquetes válidas.")
    return report, pd.DataFrame(issues), stock

def make_excel(report, issues, workshop, center, only_sendable):
    data = report.copy()
    if workshop != "Todos": data = data[data["Taller"] == workshop]
    if center != "Todos": data = data[data["Centro destino"] == int(center)]
    if only_sendable: data = data[(data["Cantidad sugerida enviar local"] > 0) & (data["Validación"] == "CONFIRMAR")]
    send = data[(data["Cantidad sugerida enviar local"] > 0) & (data["Validación"] == "CONFIRMAR")]
    no_stock = data[(data["Pendiente"] > 0) & (data["Cantidad sugerida enviar local"] <= 0) & (data["Validación"] == "CONFIRMAR")]
    review = data[data["Validación"] != "CONFIRMAR"]
    summary = data.groupby("Taller", as_index=False).agg(
        Filas=("SAP", "count"),
        Materiales_con_stock_local=("Cantidad sugerida enviar local", lambda x: int((x > 0).sum())),
        Pendiente_total=("Pendiente", "sum"),
        Cantidad_sugerida_local=("Cantidad sugerida enviar local", "sum"),
        Pendiente_sin_stock_local=("Pendiente luego de stock local", "sum")
    )
    out = BytesIO()
    with pd.ExcelWriter(out, engine="openpyxl") as writer:
        data.to_excel(writer, sheet_name="Detalle", index=False)
        send.to_excel(writer, sheet_name="Enviar stock local", index=False)
        no_stock.to_excel(writer, sheet_name="Sin stock local", index=False)
        review.to_excel(writer, sheet_name="Revisar columnas", index=False)
        summary.to_excel(writer, sheet_name="Resumen talleres", index=False)
        issues.to_excel(writer, sheet_name="Incidencias", index=False)
        for ws in writer.book.worksheets:
            ws.freeze_panes = "A2"
            ws.auto_filter.ref = ws.dimensions
            for col in ws.columns:
                letter = col[0].column_letter
                width = max((len(str(c.value)) if c.value is not None else 0) for c in col[:100])
                ws.column_dimensions[letter].width = min(width + 2, 42)
    out.seek(0)
    return out.getvalue(), data, send, no_stock, review, summary

st.title("📦 Control de Stock Piping")
st.caption("Consulta de pendientes y materiales que podés enviar desde el stock del mismo centro.")

with st.sidebar:
    st.header("Cargar archivos actualizados")
    stock_file = st.file_uploader("Stock SAP (.xlsx)", type=["xlsx", "xlsm"])
    package_files = st.file_uploader("Archivos de paquetes (.xlsx / .xlsm)", type=["xlsx", "xlsm"], accept_multiple_files=True)
    st.info("Esta versión procesa archivos cargados manualmente. No modifica los originales.")

if stock_file and package_files:
    if st.button("🔄 Procesar y actualizar reporte", type="primary", use_container_width=True):
        with st.spinner("Procesando stock y paquetes..."):
            try:
                report, issues, stock = process(stock_file, package_files)
                st.session_state["report"] = report
                st.session_state["issues"] = issues
                st.session_state["last_run"] = datetime.now().strftime("%d/%m/%Y %H:%M")
                st.success("Proceso terminado. Validá las cantidades antes de preparar despachos.")
            except Exception as e:
                st.error(f"Error al procesar: {e}")

if "report" in st.session_state:
    report = st.session_state["report"]
    issues = st.session_state["issues"]
    st.caption(f"Última actualización: {st.session_state['last_run']}")
    pending = report[report["Pendiente"] > 0]
    sendable = report[(report["Cantidad sugerida enviar local"] > 0) & (report["Validación"] == "CONFIRMAR")]
    a, b, c, d = st.columns(4)
    a.metric("Filas de materiales", len(report))
    b.metric("Con pendiente", len(pending))
    c.metric("Enviables con stock local", len(sendable))
    d.metric("Requieren revisión", int((report["Validación"] != "CONFIRMAR").sum()))

    st.subheader("Filtrar materiales")
    c1, c2, c3 = st.columns(3)
    workshop = c1.selectbox("Taller", ["Todos"] + sorted(report["Taller"].dropna().unique().tolist()))
    center = c2.selectbox("Centro destino", ["Todos"] + [str(c) for c in CENTROS])
    only_sendable = c3.checkbox("Mostrar solo lo que puedo enviar", value=True)

    view = report.copy()
    if workshop != "Todos": view = view[view["Taller"] == workshop]
    if center != "Todos": view = view[view["Centro destino"] == int(center)]
    if only_sendable: view = view[(view["Cantidad sugerida enviar local"] > 0) & (view["Validación"] == "CONFIRMAR")]

    st.subheader("Materiales")
    st.dataframe(view[[
        "Taller", "Paquete", "Centro destino", "SAP", "Descripción",
        "Pendiente", "Stock local SAP", "Cantidad sugerida enviar local",
        "Pendiente luego de stock local", "Validación", "Estado"
    ]], use_container_width=True, hide_index=True)

    xlsx, _, _, _, _, summary = make_excel(report, issues, workshop, center, only_sendable)
    st.download_button("⬇️ Descargar Excel", xlsx, "PLAN_ENVIO_STOCK_LOCAL.xlsx",
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", type="primary")
    with st.expander("Resumen por taller"):
        st.dataframe(summary, use_container_width=True, hide_index=True)
    with st.expander("Incidencias"):
        if issues.empty: st.success("No se registraron incidencias.")
        else: st.dataframe(issues, use_container_width=True, hide_index=True)
else:
    st.info("Cargá el archivo de stock SAP y los archivos de paquetes en la barra lateral, y presioná Procesar.")
