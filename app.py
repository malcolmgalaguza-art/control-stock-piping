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
    # Leer todo el archivo SAP para construir el maestro de descripciones,
    # sin filtrar por centro. El stock para cálculos sí se limita a CENTROS.
    file.seek(0)
    raw = pd.read_excel(file)
    cols = {norm(c): c for c in raw.columns}
    needed = ["MATERIAL", "TEXTO BREVE DE MATERIAL", "CENTRO", "LIBRE UTILIZACION"]
    missing = [c for c in needed if c not in cols]
    if missing:
        raise ValueError(f"Faltan columnas en stock SAP: {missing}")
    raw = raw.rename(columns={
        cols["MATERIAL"]: "SAP", cols["TEXTO BREVE DE MATERIAL"]: "Descripción SAP",
        cols["CENTRO"]: "Centro", cols["LIBRE UTILIZACION"]: "Cantidad"
    })
    raw["SAP"] = raw["SAP"].map(sap_clean)
    raw["Centro"] = pd.to_numeric(raw["Centro"], errors="coerce")
    raw["Cantidad"] = pd.to_numeric(raw["Cantidad"], errors="coerce").fillna(0)
    raw["Descripción SAP"] = raw["Descripción SAP"].fillna("").astype(str).str.strip()

    # Maestro de descripciones completo, sin filtrar centros.
    descs = {}
    for _, r in raw[raw["SAP"] != ""].iterrows():
        code, desc = r["SAP"], r["Descripción SAP"]
        if desc and code not in descs:
            descs[code] = desc

    stock = raw[raw["Centro"].isin(CENTROS) & (raw["SAP"] != "")].copy()
    stock = stock.rename(columns={"Descripción SAP": "Descripción"})
    stock = stock.groupby(["SAP", "Descripción", "Centro"], as_index=False)["Cantidad"].sum()
    return stock, descs

def find_center(file, sheet):
    sample = pd.read_excel(file, sheet_name=sheet, header=None, nrows=8)
    for v in sample.astype(object).values.ravel():
        if pd.isna(v): continue
        for c in CENTROS:
            if re.search(rf"(?<!\d){c}(?!\d)", str(v)): return c
    return None

def pending_col(columns):
    names = [norm(c) for c in columns]
    # Priorizar la columna operativa de pendiente. Algunas planillas tienen
    # "SALDO PENDIENTE" como columna auxiliar vacía/cero y "RESTAN ENVIAR"
    # con los valores reales; por eso el saldo genérico queda como respaldo.
    priorities = [
        {"RESTAN ENVIAR", "RESTA ENVIAR", "RESTAENVIAR", "RESTAN EVIAR",
         "RESTAN ENMVIAR", "RESTAN ENVIA R", "RESTANTE ENVIAR"},
        {"CANTIDAD PENDIENTE", "PENDIENTE"},
        {"SALDO PENDIENTE"},
    ]
    for group in priorities:
        for i, name in enumerate(names):
            if name in group:
                return i, name, "CONFIRMAR"
    # Usar columnas genéricas de envío solo como respaldo y marcarlas para revisión.
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

def allocate_cross_center(report, stock):
    """Asignación global: primero stock propio; después solo excedentes de otros centros.
    El stock asignado a una línea no puede volver a asignarse a otra.
    """
    out = report.copy()
    out["Cantidad sugerida enviar local"] = 0.0
    out["Cantidad sugerida préstamo"] = 0.0
    out["Obras que prestan"] = ""
    out["Pendiente luego de stock local"] = out["Pendiente"]
    out["Pendiente final"] = out["Pendiente"]
    out["Plan de origen"] = ""
    valid_idx = out.index[out["Validación"] == "CONFIRMAR"].tolist()
    if not valid_idx:
        return out, pd.DataFrame(columns=["SAP", "Descripción SAP", "Obra que presta", "Obra que recibe", "Taller", "Paquete", "Cantidad préstamo"])

    # Stock total por SAP y centro. Asignar stock local una sola vez entre todos los paquetes.
    stock_map = stock.groupby(["SAP", "Centro"], as_index=False)["Cantidad"].sum()
    available = {(str(r.SAP), int(r.Centro)): max(float(r.Cantidad), 0.0) for r in stock_map.itertuples(index=False)}
    local_demands = out.loc[valid_idx].copy()
    line_counts = local_demands[local_demands["Pendiente"] > 0].groupby(["Taller", "Paquete", "Centro destino"])["SAP"].count().to_dict()
    local_demands["_lines"] = local_demands.apply(lambda r: line_counts.get((r["Taller"], r["Paquete"], r["Centro destino"]), 9999), axis=1)
    local_demands = local_demands.sort_values(["_lines", "Pendiente", "Taller", "Paquete"], ascending=[True, True, True, True])
    for idx, row in local_demands.iterrows():
        key = (str(row["SAP"]), int(row["Centro destino"]))
        need = max(float(row["Pendiente"]), 0.0)
        take = min(need, available.get(key, 0.0))
        out.at[idx, "Cantidad sugerida enviar local"] = take
        out.at[idx, "Pendiente luego de stock local"] = max(need - take, 0.0)
        available[key] = max(available.get(key, 0.0) - take, 0.0)
        out.at[idx, "Pendiente final"] = max(need - take, 0.0)
        if take > 0:
            out.at[idx, "Plan de origen"] = f"Centro {int(row['Centro destino'])}: {take:g} propio"
        if need > 0 and take >= need:
            out.at[idx, "Estado"] = "CUBIERTO CON STOCK PROPIO"

    # Para cada SAP, los centros que cubrieron sus necesidades locales pueden prestar el excedente.
    # Priorizamos paquetes con menos líneas de material todavía pendientes para favorecer cierres.
    pending_mask = (out["Validación"] == "CONFIRMAR") & (out["Pendiente luego de stock local"] > 0)
    remaining_lines = out[pending_mask].groupby(["Taller", "Paquete", "Centro destino"]) ["SAP"].count().to_dict()
    demands = out.loc[pending_mask].copy()
    if not demands.empty:
        demands["_lineas_pendientes_paquete"] = demands.apply(lambda r: remaining_lines.get((r["Taller"], r["Paquete"], r["Centro destino"]), 9999), axis=1)
        demands = demands.sort_values(["_lineas_pendientes_paquete", "Pendiente luego de stock local", "Taller", "Paquete"], ascending=[True, True, True, True])

    loans = []
    for idx, row in demands.iterrows():
        code = str(row["SAP"])
        dest = int(row["Centro destino"])
        need = float(out.at[idx, "Pendiente luego de stock local"])
        donors = sorted(
            [(center, qty) for (sap, center), qty in available.items() if sap == code and center != dest and qty > 0],
            key=lambda x: (-x[1], x[0])
        )
        loan_sources = []
        for donor, qty in donors:
            if need <= 0: break
            take = min(need, available[(code, donor)])
            if take <= 0: continue
            available[(code, donor)] -= take
            need -= take
            out.at[idx, "Cantidad sugerida préstamo"] += take
            loan_sources.append(f"Centro {donor}: {take:g}")
            loans.append({
                "SAP": code, "Descripción SAP": row["Descripción SAP"],
                "Obra que presta": donor, "Obra que recibe": dest,
                "Taller": row["Taller"], "Paquete": row["Paquete"],
                "Cantidad préstamo": take
            })
        out.at[idx, "Obras que prestan"] = "; ".join(loan_sources)
        out.at[idx, "Pendiente final"] = max(need, 0.0)
        if out.at[idx, "Pendiente final"] <= 0 and out.at[idx, "Pendiente"] > 0:
            out.at[idx, "Estado"] = "CUBIERTO CON STOCK PROPIO Y/O PRÉSTAMO"
        elif out.at[idx, "Cantidad sugerida préstamo"] > 0:
            out.at[idx, "Estado"] = "PARCIALMENTE CUBIERTO CON PRÉSTAMO"
        parts = []
        local_qty = float(out.at[idx, "Cantidad sugerida enviar local"])
        if local_qty > 0: parts.append(f"Centro {dest}: {local_qty:g} propio")
        if loan_sources: parts.extend(loan_sources)
        out.at[idx, "Plan de origen"] = " | ".join(parts)

    loans_df = pd.DataFrame(loans, columns=["SAP", "Descripción SAP", "Obra que presta", "Obra que recibe", "Taller", "Paquete", "Cantidad préstamo"])
    return out, loans_df


def process(stock_file, package_files):
    stock, descs = read_stock(stock_file)
    stock_dict = stock.groupby(["SAP", "Centro"])["Cantidad"].sum().to_dict()
    rows, issues, audit = [], [], []
    skip_tokens = ["STOCK", "AVANCE", "RESUMEN", "TABLA6", "TABLA 6"]
    for file in package_files:
        file.seek(0)
        filename = getattr(file, "name", "archivo_sin_nombre")
        try:
            book = pd.ExcelFile(file)
        except Exception as e:
            issues.append({"Archivo": filename, "Hoja": "", "Motivo": str(e)})
            audit.append({"Archivo": filename, "Hoja": "", "Estado": "ERROR DE LECTURA", "Filas incorporadas": 0, "Detalle": str(e)})
            continue
        for sheet in book.sheet_names:
            if any(x in norm(sheet) for x in skip_tokens):
                audit.append({"Archivo": filename, "Hoja": sheet, "Estado": "OMITIDA (hoja auxiliar)", "Filas incorporadas": 0, "Detalle": "Hoja de stock, avance o resumen; no se procesa como paquete."})
                continue
            try:
                file.seek(0)
                header = find_header(file, sheet)
                if header is None:
                    audit.append({"Archivo": filename, "Hoja": sheet, "Estado": "NO INTERPRETADA", "Filas incorporadas": 0, "Detalle": "No se encontró una fila de encabezados con SAP, Descripción y columna de pendiente reconocida."})
                    issues.append({"Archivo": filename, "Hoja": sheet, "Motivo": "No se detectó encabezado válido (SAP, Descripción y columna de pendiente)."})
                    continue
                file.seek(0)
                df = pd.read_excel(file, sheet_name=sheet, header=header)
                df.columns = [str(c).strip() for c in df.columns]
                names = [norm(c) for c in df.columns]
                if "SAP" not in names or "DESCRIPCION" not in names:
                    audit.append({"Archivo": filename, "Hoja": sheet, "Estado": "NO INTERPRETADA", "Filas incorporadas": 0, "Detalle": "Falta SAP o Descripción."})
                    issues.append({"Archivo": filename, "Hoja": sheet, "Motivo": "Falta SAP o Descripción."})
                    continue
                sap_idx, desc_idx = names.index("SAP"), names.index("DESCRIPCION")
                p_idx, p_name, quality = pending_col(df.columns)
                if p_idx is None:
                    audit.append({"Archivo": filename, "Hoja": sheet, "Estado": "NO INTERPRETADA", "Filas incorporadas": 0, "Detalle": "No se encontró columna de pendiente reconocida."})
                    issues.append({"Archivo": filename, "Hoja": sheet, "Motivo": "Sin columna pendiente"})
                    continue
                file.seek(0)
                center = find_center(file, sheet)
                if center is None:
                    audit.append({"Archivo": filename, "Hoja": sheet, "Estado": "NO INTERPRETADA", "Filas incorporadas": 0, "Detalle": "No se detectó centro destino."})
                    issues.append({"Archivo": filename, "Hoja": sheet, "Motivo": "No se detectó centro destino"})
                    continue
                normalized_filename = norm(filename)
                workshop = next((x for x in ["TAPIGAR", "ARMEXAS", "INTERACERO", "TYCSA"] if x in normalized_filename), "OTRO")
                valid_rows = 0
                for _, row in df.iterrows():
                    code = sap_clean(row.iloc[sap_idx])
                    if not code or code.upper() in ["NAN", "SAP"]: continue
                    pending = num(row.iloc[p_idx])
                    description_workshop = str(row.iloc[desc_idx]).strip() if pd.notna(row.iloc[desc_idx]) else ""
                    # La descripción principal siempre sale del maestro de materiales del stock SAP.
                    # La descripción escrita en el archivo del taller se conserva en otra columna para comparar.
                    description_sap = descs.get(code, "")
                    local_stock = max(float(stock_dict.get((code, center), 0)), 0)
                    pending_pos = max(pending, 0)
                    suggested = min(pending_pos, local_stock) if quality == "CONFIRMAR" else 0
                    rows.append({
                        "Taller": workshop, "Paquete": sheet, "Centro destino": int(center),
                        "SAP": code, "Descripción SAP": description_sap if description_sap else "NO ENCONTRADO EN STOCK SAP", "Descripción taller": description_workshop,
                        "Columna pendiente": p_name,
                        "Pendiente original": pending, "Pendiente": pending_pos,
                        "Stock local SAP": local_stock, "Cantidad sugerida enviar local": suggested,
                        "Pendiente luego de stock local": max(pending_pos-suggested, 0),
                        "Validación": quality,
                        "Estado": "REVISAR COLUMNA" if quality != "CONFIRMAR" else ("ENVIAR DESDE STOCK LOCAL" if suggested > 0 else "SIN STOCK LOCAL")
                    })
                    valid_rows += 1
                if valid_rows:
                    audit.append({"Archivo": filename, "Hoja": sheet, "Estado": "PROCESADA" if quality == "CONFIRMAR" else "PROCESADA CON REVISIÓN", "Filas incorporadas": valid_rows, "Detalle": f"Centro {center}; columna usada: {p_name}; validación: {quality}."})
                else:
                    audit.append({"Archivo": filename, "Hoja": sheet, "Estado": "SIN MATERIALES", "Filas incorporadas": 0, "Detalle": f"Se detectaron encabezados, pero no se encontraron filas con código SAP. Columna: {p_name}."})
                    issues.append({"Archivo": filename, "Hoja": sheet, "Motivo": "No se encontraron filas con código SAP después del encabezado."})
            except Exception as e:
                audit.append({"Archivo": filename, "Hoja": sheet, "Estado": "ERROR", "Filas incorporadas": 0, "Detalle": str(e)})
                issues.append({"Archivo": filename, "Hoja": sheet, "Motivo": str(e)})
    report = pd.DataFrame(rows)
    audit_df = pd.DataFrame(audit)
    if report.empty: raise ValueError("No se encontraron tablas de paquetes válidas. Revisá la hoja Control de lectura.")
    report, loans = allocate_cross_center(report, stock)
    return report, pd.DataFrame(issues), stock, audit_df, loans

def make_excel(report, issues, audit, loans, workshop, center, only_sendable):
    data = report.copy()
    if workshop != "Todos": data = data[data["Taller"] == workshop]
    if center != "Todos": data = data[data["Centro destino"] == int(center)]
    if only_sendable: data = data[(data["Cantidad sugerida enviar local"] > 0) & (data["Validación"] == "CONFIRMAR")]
    send = data[((data["Cantidad sugerida enviar local"] > 0) | (data["Cantidad sugerida préstamo"] > 0)) & (data["Validación"] == "CONFIRMAR")]
    no_stock = data[(data["Pendiente"] > 0) & (data["Pendiente final"] > 0) & (data["Validación"] == "CONFIRMAR")]
    if workshop != "Todos": loans = loans[loans["Taller"] == workshop]
    if center != "Todos": loans = loans[loans["Obra que recibe"] == int(center)]
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
        send.to_excel(writer, sheet_name="Preparar envíos", index=False)
        loans.to_excel(writer, sheet_name="Préstamos entre obras", index=False)
        no_stock.to_excel(writer, sheet_name="Faltante final", index=False)
        review.to_excel(writer, sheet_name="Revisar columnas", index=False)
        summary.to_excel(writer, sheet_name="Resumen talleres", index=False)
        issues.to_excel(writer, sheet_name="Incidencias", index=False)
        audit.to_excel(writer, sheet_name="Control de lectura", index=False)
        for ws in writer.book.worksheets:
            ws.freeze_panes = "A2"
            ws.auto_filter.ref = ws.dimensions
            for col in ws.columns:
                letter = col[0].column_letter
                width = max((len(str(c.value)) if c.value is not None else 0) for c in col[:100])
                ws.column_dimensions[letter].width = min(width + 2, 42)
    out.seek(0)
    return out.getvalue(), data, send, no_stock, review, summary, loans

st.title("📦 Control de Stock Piping")
st.caption("La descripción principal se toma del maestro SAP completo del archivo cargado; la descripción del taller queda separada para comparar.")

with st.sidebar:
    st.header("Cargar archivos actualizados")
    stock_file = st.file_uploader("Stock SAP (.xlsx)", type=["xlsx", "xlsm"])
    package_files = st.file_uploader("Archivos de paquetes (.xlsx / .xlsm)", type=["xlsx", "xlsm"], accept_multiple_files=True)
    st.info("Esta versión procesa archivos cargados manualmente. No modifica los originales.")

if stock_file and package_files:
    if st.button("🔄 Procesar y actualizar reporte", type="primary", use_container_width=True):
        with st.spinner("Procesando stock y paquetes..."):
            try:
                report, issues, stock, audit, loans = process(stock_file, package_files)
                st.session_state["report"] = report
                st.session_state["loans"] = loans
                st.session_state["issues"] = issues
                st.session_state["audit"] = audit
                st.session_state["last_run"] = datetime.now().strftime("%d/%m/%Y %H:%M")
                st.success("Proceso terminado. Validá las cantidades antes de preparar despachos.")
            except Exception as e:
                st.error(f"Error al procesar: {e}")

if "report" in st.session_state:
    report = st.session_state["report"]
    issues = st.session_state["issues"]
    audit = st.session_state["audit"]
    loans = st.session_state.get("loans", pd.DataFrame())
    st.caption(f"Última actualización: {st.session_state['last_run']}")
    pending = report[report["Pendiente"] > 0]
    sendable = report[((report["Cantidad sugerida enviar local"] > 0) | (report["Cantidad sugerida préstamo"] > 0)) & (report["Validación"] == "CONFIRMAR")]
    a, b, c, d = st.columns(4)
    a.metric("Filas de materiales", len(report))
    b.metric("Materiales con pendiente", len(pending))
    c.metric("Materiales enviables", len(sendable))
    d.metric("Filas a revisar", int((report["Validación"] != "CONFIRMAR").sum()))
    st.caption(f"Hojas de paquetes revisadas: {len(audit)} · Hojas con materiales incorporados: {int(audit['Filas incorporadas'].gt(0).sum())}")

    # Dashboard: quantities are shown as material quantities, not as a percentage
    # of package completion, because these source files do not always contain the original total.
    st.header("📊 Tablero de avances y pendientes")
    valid = report[report["Validación"] == "CONFIRMAR"].copy()
    if valid.empty:
        st.info("No hay filas con una columna de pendiente confirmada para graficar.")
    else:
        totals = pd.DataFrame({
            "Concepto": ["Pendiente total", "Cubierto con stock propio", "Préstamo sugerido", "Faltante final"],
            "Cantidad": [valid["Pendiente"].sum(), valid["Cantidad sugerida enviar local"].sum(), valid["Cantidad sugerida préstamo"].sum(), valid["Pendiente final"].sum()]
        }).set_index("Concepto")
        st.subheader("Vista general de cantidades")
        st.bar_chart(totals, y="Cantidad", use_container_width=True)

        workshop_dashboard = valid.groupby("Taller", as_index=True).agg(
            **{"Pendiente total": ("Pendiente", "sum"),
               "Cubierto con stock local": ("Cantidad sugerida enviar local", "sum"),
               "Préstamo sugerido": ("Cantidad sugerida préstamo", "sum"),
               "Pendiente restante": ("Pendiente final", "sum")}
        ).sort_values("Pendiente total", ascending=False)
        st.subheader("Pendientes por taller")
        st.bar_chart(workshop_dashboard, use_container_width=True)

        package_dashboard = valid.groupby(["Taller", "Paquete"], as_index=False).agg(
            Pendiente=("Pendiente", "sum"),
            **{"Stock local sugerido": ("Cantidad sugerida enviar local", "sum"),
               "Préstamo sugerido": ("Cantidad sugerida préstamo", "sum"),
               "Saldo restante": ("Pendiente final", "sum")}
        )
        package_dashboard["Taller / paquete"] = package_dashboard["Taller"] + " / " + package_dashboard["Paquete"].astype(str)
        package_dashboard = package_dashboard.sort_values("Pendiente", ascending=False)
        st.subheader("Paquetes con mayor pendiente")
        st.caption("Se muestran los 15 paquetes con mayor cantidad pendiente. Las cantidades corresponden a unidades de los materiales, no a porcentaje de avance.")
        top_packages = package_dashboard.head(15).set_index("Taller / paquete")[["Pendiente", "Stock local sugerido", "Saldo restante"]]
        st.bar_chart(top_packages, use_container_width=True)
        with st.expander("Ver tabla de avance y pendientes por paquete"):
            st.dataframe(package_dashboard[["Taller", "Paquete", "Pendiente", "Stock local sugerido", "Préstamo sugerido", "Saldo restante"]], use_container_width=True, hide_index=True)
    desc_missing = int(report["Descripción SAP"].eq("NO ENCONTRADO EN STOCK SAP").sum())
    if desc_missing:
        st.warning(f"{desc_missing} filas no encontraron descripción en el archivo SAP cargado. Revisá los códigos en la columna SAP.")

    st.subheader("Filtrar materiales")
    c1, c2, c3 = st.columns(3)
    workshop = c1.selectbox("Taller", ["Todos"] + sorted(report["Taller"].dropna().unique().tolist()))
    center = c2.selectbox("Centro destino", ["Todos"] + [str(c) for c in CENTROS])
    only_sendable = c3.checkbox("Mostrar solo lo que puedo enviar", value=False)

    view = report.copy()
    if workshop != "Todos": view = view[view["Taller"] == workshop]
    if center != "Todos": view = view[view["Centro destino"] == int(center)]
    if only_sendable: view = view[((view["Cantidad sugerida enviar local"] > 0) | (view["Cantidad sugerida préstamo"] > 0)) & (view["Validación"] == "CONFIRMAR")]

    st.subheader("Materiales")
    st.dataframe(view[[
        "Taller", "Paquete", "Centro destino", "SAP", "Descripción SAP", "Descripción taller",
        "Pendiente", "Stock local SAP", "Cantidad sugerida enviar local",
        "Cantidad sugerida préstamo", "Obras que prestan", "Plan de origen", "Pendiente final", "Validación", "Estado"
    ]], use_container_width=True, hide_index=True)

    st.subheader("🔁 Préstamos sugeridos entre obras")
    if loans.empty:
        st.info("No hay préstamos sugeridos con el stock y los paquetes cargados.")
    else:
        st.caption("Estas son propuestas automáticas: revisalas antes de confirmar movimientos físicos.")
        st.dataframe(loans, use_container_width=True, hide_index=True)

    xlsx, _, _, _, _, summary, loans_filtered = make_excel(report, issues, audit, loans, workshop, center, only_sendable)
    st.download_button("⬇️ Descargar Excel de preparación y préstamos", xlsx, "PLAN_ENVIO_PREPARACION_PRESTAMOS.xlsx",
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", type="primary")
    with st.expander("Resumen por taller"):
        st.dataframe(summary, use_container_width=True, hide_index=True)
    with st.expander("Control de lectura de archivos y paquetes", expanded=True):
        st.caption("Se muestran todas las hojas encontradas en los archivos cargados, incluidas las omitidas o no interpretadas.")
        st.dataframe(audit, use_container_width=True, hide_index=True)
    with st.expander("Incidencias"):
        if issues.empty: st.success("No se registraron incidencias.")
        else: st.dataframe(issues, use_container_width=True, hide_index=True)
else:
    st.info("Cargá el archivo de stock SAP y los archivos de paquetes en la barra lateral, y presioná Procesar.")
