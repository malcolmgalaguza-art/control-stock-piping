import streamlit as st
import pandas as pd
import numpy as np
import re, unicodedata
from io import BytesIO
from datetime import datetime

st.set_page_config(page_title="Control de Stock Piping", page_icon="📦", layout="wide")
CENTROS = [3013, 3117, 3118, 3119, 3120, 3121]  # 3115 y 3116 excluidos

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
    # Solo almacén 0008. Cada centro conserva su stock como posible origen de préstamo.
    file.seek(0)
    raw = pd.read_excel(file)
    cols = {norm(c): c for c in raw.columns}
    needed = ["MATERIAL", "TEXTO BREVE DE MATERIAL", "CENTRO", "ALMACEN", "LIBRE UTILIZACION"]
    missing = [c for c in needed if c not in cols]
    if missing:
        raise ValueError(f"Faltan columnas en stock SAP: {missing}")
    raw = raw.rename(columns={
        cols["MATERIAL"]: "SAP", cols["TEXTO BREVE DE MATERIAL"]: "Descripción SAP",
        cols["CENTRO"]: "Centro", cols["ALMACEN"]: "Almacén", cols["LIBRE UTILIZACION"]: "Cantidad"
    })
    raw["SAP"] = raw["SAP"].map(sap_clean)
    raw["Centro"] = pd.to_numeric(raw["Centro"], errors="coerce")
    raw["Almacén"] = pd.to_numeric(raw["Almacén"], errors="coerce")
    raw["Cantidad"] = pd.to_numeric(raw["Cantidad"], errors="coerce").fillna(0)
    raw["Descripción SAP"] = raw["Descripción SAP"].fillna("").astype(str).str.strip()
    descs = {}
    for _, r in raw[raw["SAP"] != ""].iterrows():
        if r["Descripción SAP"] and r["SAP"] not in descs:
            descs[r["SAP"]] = r["Descripción SAP"]
    stock = raw[raw["Centro"].isin(CENTROS) & (raw["Almacén"] == 8) & (raw["SAP"] != "")].copy()
    stock = stock.rename(columns={"Descripción SAP": "Descripción"})
    stock["Origen"] = "SAP 0008"
    stock = stock.groupby(["SAP", "Descripción", "Centro", "Almacén", "Origen"], as_index=False)["Cantidad"].sum()
    return stock, descs

def read_returns(file, descs):
    if file is None:
        return pd.DataFrame(columns=["SAP", "Descripción", "Centro", "Almacén", "Origen", "Cantidad", "Unidad de medida"])
    file.seek(0)
    raw = pd.read_excel(file)
    cols = {norm(c): c for c in raw.columns}
    code_col = next((cols[k] for k in ["CODIGO SAP", "CODIGO", "SAP", "MATERIAL"] if k in cols), None)
    qty_col = next((cols[k] for k in ["CANTIDAD", "CANT", "LIBRE UTILIZACION"] if k in cols), None)
    unit_col = next((cols[k] for k in ["UNIDAD DE MEDIDA", "UNIDAD", "UM"] if k in cols), None)
    if not code_col or not qty_col:
        raise ValueError("El archivo de devoluciones debe tener Código SAP y Cantidad.")
    raw["SAP"] = raw[code_col].map(sap_clean)
    raw["Cantidad"] = raw[qty_col].map(num)
    raw["Unidad de medida"] = raw[unit_col].fillna("").astype(str).str.upper().str.strip() if unit_col else ""
    raw = raw[(raw["SAP"] != "") & (raw["Cantidad"] > 0)]
    out = raw.groupby(["SAP", "Unidad de medida"], as_index=False)["Cantidad"].sum()
    out["Descripción"] = out["SAP"].map(descs).fillna("DEVOLUCIÓN SMM")
    out["Centro"] = 0
    out["Almacén"] = 8
    out["Origen"] = "SMM - DEVOLUCIÓN"
    return out[["SAP", "Descripción", "Centro", "Almacén", "Origen", "Cantidad", "Unidad de medida"]]

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

def allocate_cross_center(report, stock, returns):
    """Asigna stock 0008 de centros como préstamos y devoluciones como stock SMM.
    No consulta ni utiliza almacén 0005. Cada cantidad se reserva una sola vez.
    """
    out = report.copy()
    out["Cantidad sugerida enviar desde SMM"] = 0.0
    out["Cantidad sugerida por préstamo"] = 0.0
    out["Centros que prestan"] = ""
    out["Pendiente final"] = out["Pendiente"]
    out["Plan de abastecimiento"] = ""
    loan_cols = ["SAP", "Descripción SAP", "Tipo de movimiento", "Centro que envía", "Centro destino", "Taller", "Paquete", "Cantidad"]
    valid_idx = out.index[out["Validación"] == "CONFIRMAR"].tolist()
    if not valid_idx:
        return out, pd.DataFrame(columns=loan_cols)

    # Bolsas independientes: devoluciones SMM y stock 0008 por centro.
    smm = returns.groupby("SAP")["Cantidad"].sum().to_dict() if not returns.empty else {}
    by_center = stock.groupby(["SAP", "Centro"])["Cantidad"].sum()
    available = {(str(s), int(c)): max(float(q), 0.0) for (s, c), q in by_center.items()}
    demands = out.loc[valid_idx].copy()
    line_counts = demands[demands["Pendiente"] > 0].groupby(["Taller", "Paquete", "Centro destino"])["SAP"].count().to_dict()
    demands["_lines"] = demands.apply(lambda r: line_counts.get((r["Taller"], r["Paquete"], r["Centro destino"]), 9999), axis=1)
    demands = demands.sort_values(["_lines", "Pendiente", "Taller", "Paquete"], ascending=[True, True, True, True])
    movements = []
    for idx, row in demands.iterrows():
        code, dest = str(row["SAP"]), int(row["Centro destino"])
        need = max(float(row["Pendiente"]), 0.0)
        # Primero usar devoluciones que están físicamente en SMM.
        take_smm = min(need, max(float(smm.get(code, 0)), 0.0))
        if take_smm > 0:
            smm[code] -= take_smm
            need -= take_smm
            out.at[idx, "Cantidad sugerida enviar desde SMM"] += take_smm
            movements.append({"SAP": code, "Descripción SAP": row["Descripción SAP"], "Tipo de movimiento": "ENVÍO DESDE SMM (DEVOLUCIÓN)", "Centro que envía": "SMM", "Centro destino": dest, "Taller": row["Taller"], "Paquete": row["Paquete"], "Cantidad": take_smm})
        # Después, préstamo desde el 0008 de los centros. Prioriza centros distintos del destino
        # y conserva el origen para informar quién presta a quién.
        donors = [(c, q) for (s, c), q in available.items() if s == code and q > 0]
        donors.sort(key=lambda x: (0 if x[0] != dest else 1, -x[1], x[0]))
        sources = []
        for donor, qty in donors:
            if need <= 0: break
            take = min(need, available[(code, donor)])
            if take <= 0: continue
            available[(code, donor)] -= take
            need -= take
            out.at[idx, "Cantidad sugerida por préstamo"] += take
            movement_type = "STOCK DEL CENTRO DESTINO" if donor == dest else "PRÉSTAMO ENTRE CENTROS"
            sources.append(f"{movement_type} - Centro {donor}: {take:g}")
            movements.append({"SAP": code, "Descripción SAP": row["Descripción SAP"], "Tipo de movimiento": movement_type, "Centro que envía": donor, "Centro destino": dest, "Taller": row["Taller"], "Paquete": row["Paquete"], "Cantidad": take})
        out.at[idx, "Centros que prestan"] = "; ".join(sources)
        out.at[idx, "Pendiente final"] = max(need, 0.0)
        plan = []
        if take_smm > 0: plan.append(f"SMM devolución: {take_smm:g}")
        if sources: plan.extend([f"PRÉSTAMO {x}" for x in sources])
        out.at[idx, "Plan de abastecimiento"] = " | ".join(plan)
        if float(out.at[idx, "Pendiente"]) > 0 and need <= 0:
            out.at[idx, "Estado"] = "CUBIERTO CON SMM Y/O PRÉSTAMO"
        elif sources or take_smm > 0:
            out.at[idx, "Estado"] = "PARCIALMENTE CUBIERTO"
        elif float(out.at[idx, "Pendiente"]) > 0:
            out.at[idx, "Estado"] = "SIN STOCK 0008 DISPONIBLE"
    return out, pd.DataFrame(movements, columns=loan_cols)

def process(stock_file, package_files, returns_file=None):
    stock, descs = read_stock(stock_file)
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
                    pending_pos = max(pending, 0)
                    rows.append({
                        "Taller": workshop, "Paquete": sheet, "Centro destino": int(center),
                        "SAP": code, "Descripción SAP": description_sap if description_sap else "NO ENCONTRADO EN STOCK SAP", "Descripción taller": description_workshop,
                        "Columna pendiente": p_name,
                        "Pendiente original": pending, "Pendiente": pending_pos,
                        "Validación": quality,
                        "Estado": "REVISAR COLUMNA" if quality != "CONFIRMAR" else "PENDIENTE DE ASIGNACIÓN"
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
    returns = read_returns(returns_file, descs)
    report, movements = allocate_cross_center(report, stock, returns)
    return report, pd.DataFrame(issues), stock, audit_df, movements, returns

def make_excel(report, issues, audit, movements, stock, returns, workshop, center, only_sendable):
    data = report.copy()
    if workshop != "Todos": data = data[data["Taller"] == workshop]
    if center != "Todos": data = data[data["Centro destino"] == int(center)]
    sendable_mask = ((data["Cantidad sugerida enviar desde SMM"] > 0) | (data["Cantidad sugerida por préstamo"] > 0)) & (data["Validación"] == "CONFIRMAR")
    if only_sendable: data = data[sendable_mask]
    send = data[((data["Cantidad sugerida enviar desde SMM"] > 0) | (data["Cantidad sugerida por préstamo"] > 0)) & (data["Validación"] == "CONFIRMAR")]
    no_stock = data[(data["Pendiente"] > 0) & (data["Pendiente final"] > 0) & (data["Validación"] == "CONFIRMAR")]
    if workshop != "Todos": movements = movements[movements["Taller"] == workshop]
    if center != "Todos": movements = movements[movements["Centro destino"] == int(center)]
    review = data[data["Validación"] != "CONFIRMAR"]
    summary = data.groupby("Taller", as_index=False).agg(
        Filas=("SAP", "count"), Pendiente_total=("Pendiente", "sum"),
        Envio_SMM=("Cantidad sugerida enviar desde SMM", "sum"),
        Prestamos=("Cantidad sugerida por préstamo", "sum"),
        Pendiente_no_cubierto=("Pendiente final", "sum"))
    out = BytesIO()
    with pd.ExcelWriter(out, engine="openpyxl") as writer:
        data.to_excel(writer, sheet_name="Detalle", index=False)
        send.to_excel(writer, sheet_name="Preparar envíos", index=False)
        movements.to_excel(writer, sheet_name="Préstamos y envíos", index=False)
        stock.to_excel(writer, sheet_name="Stock SAP 0008 por centro", index=False)
        returns.to_excel(writer, sheet_name="Devoluciones SMM", index=False)
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
    return out.getvalue(), data, send, no_stock, review, summary

st.title("📦 Control de Stock Piping")
st.caption("Solo almacén 0008 de los centros habilitados. El stock de otros centros se propone como PRÉSTAMO; las devoluciones se consideran stock adicional de SMM.")

with st.sidebar:
    st.header("Cargar archivos actualizados")
    stock_file = st.file_uploader("Stock SAP (.xlsx)", type=["xlsx", "xlsm"])
    returns_file = st.file_uploader("Stock adicional SMM - devoluciones (opcional)", type=["xlsx", "xlsm"])
    package_files = st.file_uploader("Archivos de paquetes (.xlsx / .xlsm)", type=["xlsx", "xlsm"], accept_multiple_files=True)
    st.info("Se utiliza únicamente el almacén 0008. El stock de otros centros se propone como PRÉSTAMO; las devoluciones se consideran stock de SMM.")

if stock_file and package_files:
    if st.button("🔄 Procesar y actualizar reporte", type="primary", use_container_width=True):
        with st.spinner("Procesando stock y paquetes..."):
            try:
                report, issues, stock, audit, movements, returns = process(stock_file, package_files, returns_file)
                st.session_state["report"] = report
                st.session_state["movements"] = movements
                st.session_state["returns"] = returns
                st.session_state["stock"] = stock
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
    movements = st.session_state.get("movements", pd.DataFrame())
    returns = st.session_state.get("returns", pd.DataFrame())
    st.caption(f"Última actualización: {st.session_state['last_run']}")
    pending = report[report["Pendiente"] > 0]
    sendable = report[((report["Cantidad sugerida enviar desde SMM"] > 0) | (report["Cantidad sugerida por préstamo"] > 0)) & (report["Validación"] == "CONFIRMAR")]
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
            "Concepto": ["Pendiente total", "Envío desde SMM", "Préstamos entre centros", "Faltante final"],
            "Cantidad": [valid["Pendiente"].sum(), valid["Cantidad sugerida enviar desde SMM"].sum(), valid["Cantidad sugerida por préstamo"].sum(), valid["Pendiente final"].sum()]
        }).set_index("Concepto")
        st.subheader("Vista general de cantidades")
        st.bar_chart(totals, y="Cantidad", use_container_width=True)

        workshop_dashboard = valid.groupby("Taller", as_index=True).agg(
            **{"Pendiente total": ("Pendiente", "sum"),
               "Envío desde SMM": ("Cantidad sugerida enviar desde SMM", "sum"),
               "Préstamos entre centros": ("Cantidad sugerida por préstamo", "sum"),
               "Pendiente restante": ("Pendiente final", "sum")}
        ).sort_values("Pendiente total", ascending=False)
        st.subheader("Pendientes por taller")
        st.bar_chart(workshop_dashboard, use_container_width=True)

        package_dashboard = valid.groupby(["Taller", "Paquete"], as_index=False).agg(
            Pendiente=("Pendiente", "sum"),
            **{"Envío desde SMM": ("Cantidad sugerida enviar desde SMM", "sum"),
               "Préstamos": ("Cantidad sugerida por préstamo", "sum"),
               "Saldo restante": ("Pendiente final", "sum")}
        )
        package_dashboard["Taller / paquete"] = package_dashboard["Taller"] + " / " + package_dashboard["Paquete"].astype(str)
        package_dashboard = package_dashboard.sort_values("Pendiente", ascending=False)
        st.subheader("Paquetes con mayor pendiente")
        st.caption("Se muestran los 15 paquetes con mayor cantidad pendiente. Las cantidades corresponden a unidades de los materiales, no a porcentaje de avance.")
        top_packages = package_dashboard.head(15).set_index("Taller / paquete")[["Pendiente", "Envío desde SMM", "Préstamos", "Saldo restante"]]
        st.bar_chart(top_packages, use_container_width=True)
        with st.expander("Ver tabla de avance y pendientes por paquete"):
            st.dataframe(package_dashboard[["Taller", "Paquete", "Pendiente", "Envío desde SMM", "Préstamos", "Saldo restante"]], use_container_width=True, hide_index=True)
    desc_missing = int(report["Descripción SAP"].eq("NO ENCONTRADO EN STOCK SAP").sum())
    if desc_missing:
        st.warning(f"{desc_missing} filas no encontraron descripción en el archivo SAP cargado. Revisá los códigos en la columna SAP.")

    st.subheader("Filtrar materiales")
    c1, c2, c3 = st.columns(3)
    workshop = c1.selectbox("Taller", ["Todos"] + sorted(report["Taller"].dropna().unique().tolist()))
    center = c2.selectbox("Centro destino", ["Todos"] + [str(c) for c in CENTROS])
    only_sendable = c3.checkbox("Mostrar solo materiales con envío o préstamo", value=False)

    view = report.copy()
    if workshop != "Todos": view = view[view["Taller"] == workshop]
    if center != "Todos": view = view[view["Centro destino"] == int(center)]
    if only_sendable: view = view[((view["Cantidad sugerida enviar desde SMM"] > 0) | (view["Cantidad sugerida por préstamo"] > 0)) & (view["Validación"] == "CONFIRMAR")]

    st.subheader("Materiales")
    st.dataframe(view[[
        "Taller", "Paquete", "Centro destino", "SAP", "Descripción SAP", "Descripción taller",
        "Pendiente", "Cantidad sugerida enviar desde SMM", "Cantidad sugerida por préstamo",
        "Centros que prestan", "Plan de abastecimiento", "Pendiente final", "Validación", "Estado"
    ]], use_container_width=True, hide_index=True)

    st.subheader("📨 Préstamos entre centros y envíos desde SMM")
    if movements.empty:
        st.info("No hay envíos ni préstamos sugeridos con el stock y los paquetes cargados.")
    else:
        st.caption("Los movimientos entre centros se identifican expresamente como PRÉSTAMO. Las devoluciones se consideran stock disponible en SMM.")
        st.dataframe(movements, use_container_width=True, hide_index=True)

    xlsx, _, _, _, _, summary = make_excel(report, issues, audit, movements, st.session_state["stock"], returns, workshop, center, only_sendable)
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
    st.info("Cargá el stock SAP, los archivos de paquetes y, opcionalmente, el archivo de devoluciones SMM. Se utiliza únicamente el almacén 0008.")
