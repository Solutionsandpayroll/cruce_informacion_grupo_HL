import io
import re
import copy
import tempfile
import os
import json
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
from openpyxl.styles.differential import DifferentialStyle
from openpyxl.formatting.rule import Rule


# ------------------------------------------------------------
# 1. Lectura del archivo Reporte
# ------------------------------------------------------------
def parse_reporte(file_obj) -> pd.DataFrame:
    wb = load_workbook(file_obj, data_only=False)
    sheet_name = next(
        (name for name in ("Datos", "Reporte") if name in wb.sheetnames), wb.sheetnames[0]
    )
    ws = wb[sheet_name]

    codigo_col_idx = None
    for col in range(1, ws.max_column + 1):
        cell_value = ws.cell(row=1, column=col).value
        if cell_value and str(cell_value).strip() == "Código":
            codigo_col_idx = col
            break

    if codigo_col_idx is None:
        raise Exception(f"No se encontró la columna 'Código' en la hoja {sheet_name}")

    codigo_raw = {}
    for row in range(2, ws.max_row + 1):
        cell = ws.cell(row=row, column=codigo_col_idx)
        val = cell.value
        codigo_raw[row] = val if val is not None else ""

    file_obj.seek(0)
    df = pd.read_excel(file_obj, sheet_name=sheet_name, dtype=str, keep_default_na=False)

    if "Código" in df.columns:
        for idx, row in df.iterrows():
            excel_row = idx + 2
            df.at[idx, "Código"] = codigo_raw.get(excel_row, "")

    df["Código"] = df["Código"].replace(["nan", "None", ""], "")
    return df


# ------------------------------------------------------------
# 2. Lectura del archivo TrabajadoresVigentes
# ------------------------------------------------------------
def _extract_riesgo_from_sheet(ws) -> str:
    for row in ws.iter_rows(min_row=1, max_row=30):
        for cell in row:
            if cell.value and str(cell.value).strip().lower().startswith("riesgo"):
                for offset in range(1, 5):
                    neighbour = ws.cell(row=cell.row, column=cell.column + offset)
                    if neighbour.value is not None and str(neighbour.value).strip() not in ("", "nan"):
                        return str(neighbour.value).strip()
    return ""


def parse_trabajadores(file_obj) -> pd.DataFrame:
    xl = pd.ExcelFile(file_obj)
    file_obj.seek(0)
    wb = load_workbook(file_obj, data_only=True)

    records = []
    for sheet in xl.sheet_names:
        ws = wb[sheet]
        riesgo_val = _extract_riesgo_from_sheet(ws)

        df_sheet = pd.read_excel(xl, sheet_name=sheet, header=None)

        header_row_idx = None
        id_col = None
        for idx, row in df_sheet.iterrows():
            for col_idx, val in enumerate(row):
                if pd.notna(val) and "Identificaci" in str(val):
                    header_row_idx = idx
                    id_col = col_idx
                    break
            if header_row_idx is not None:
                break

        if header_row_idx is None:
            continue

        header = df_sheet.iloc[header_row_idx]
        col_map = {}
        for col_idx, val in enumerate(header):
            if pd.notna(val):
                col_map[str(val).strip()] = col_idx

        campos = [
            "Identificación", "Nombre", "Cargo",
            "Inicio Vigencia", "EPS", "AFP", "Salario", "Fecha Nac.",
            "C.COSTO",
        ]

        for data_idx in range(header_row_idx + 1, len(df_sheet)):
            row_data = df_sheet.iloc[data_idx]
            id_val = str(row_data[id_col]) if pd.notna(row_data[id_col]) else ""
            id_val = id_val.strip()
            if not id_val or id_val == "nan" or "NÚMERO" in id_val or id_val == "." or id_val == "0":
                continue

            rec = {"raw_id": id_val, "Riesgo ARL": riesgo_val}
            for field in campos:
                col_idx = col_map.get(field)
                if col_idx is not None:
                    v = row_data[col_idx]
                    rec[field] = str(v).strip() if pd.notna(v) else ""
                else:
                    rec[field] = ""
            records.append(rec)

    if not records:
        raise Exception("No se encontraron datos de trabajadores")

    df = pd.DataFrame(records)

    def split_id(s):
        s = re.sub(r"[^\w\s]", "", s)
        m = re.match(r"^(CC|CE|TI|PA|PEP|PT)\s+(.+)$", s.strip(), re.IGNORECASE)
        if m:
            return m.group(1).upper(), m.group(2).strip()
        digits_only = re.sub(r"[^\d]", "", s)
        if digits_only:
            return "CC", digits_only
        return "CC", s.strip()

    df["Tipo"] = df["raw_id"].apply(lambda x: split_id(x)[0])
    df["ID_Num"] = df["raw_id"].apply(lambda x: split_id(x)[1])
    return df


# ------------------------------------------------------------
# 3. Previsualizaciones para el dashboard
# ------------------------------------------------------------
def build_cruce_preview(df_trab: pd.DataFrame) -> list[dict]:
    preview = []
    for _, row in df_trab.iterrows():
        preview.append({
            "Tipo": row.get("Tipo", ""),
            "Identificación": row.get("ID_Num", ""),
            "Nombre": row.get("Nombre", ""),
            "Cargo": row.get("Cargo", ""),
            "Inicio Vigencia": row.get("Inicio Vigencia", ""),
            "EPS": row.get("EPS", ""),
            "AFP": row.get("AFP", ""),
            "Salario": row.get("Salario", ""),
            "Fecha Nac.": row.get("Fecha Nac.", ""),
            "RIESGO EN ARL": row.get("Riesgo ARL", ""),
            "C. COSTO": "",
            "LIBRA": "",
            "VALIDACION": "",
        })
    return preview


def build_emp_preview(df_rep: pd.DataFrame) -> list[dict]:
    cols_show = [
        "Cédula identificación", "Código", "Apellidos, Nombre",
        "Nombre", "EPS", "AFP", "CCF", "C.COSTO",
        "NIVEL ARL", "Salario Mes",
    ]
    existing = [c for c in cols_show if c in df_rep.columns]
    if not existing:
        existing = list(df_rep.columns[:10])
    return df_rep[existing].head(50).fillna("").to_dict(orient="records")


# ------------------------------------------------------------
# 4. Funciones auxiliares de formato
# ------------------------------------------------------------
def normalize_id(id_val) -> str:
    if pd.isna(id_val):
        return ""
    s = str(id_val).strip()
    return re.sub(r"[^\d]", "", s)


def format_nivel_arl(val) -> str:
    try:
        return f"{int(float(val)):02d}"
    except:
        return ""


def _id_to_int(id_str) -> int | None:
    digits = re.sub(r"[^\d]", "", str(id_str))
    try:
        return int(digits) if digits else None
    except ValueError:
        return None


def format_riesgo_arl(val) -> str:
    """Asegura formato de dos dígitos para el riesgo, ej. '5' → '05'."""
    try:
        return f"{int(float(val)):02d}"
    except:
        return str(val).strip() if val else ""


def format_codigo_emp(val) -> str:
    if pd.isna(val) or val == "":
        return ""
    try:
        num = float(val)
    except (ValueError, TypeError):
        return str(val).strip()

    if num == int(num) and num >= 10**7:
        int_val = int(num)
        if int_val % 10**7 == 0:
            coef = int_val // 10**7
            return f"{coef:05d}E07"

    if num == int(num):
        return str(int(num))
    return str(num)


def extract_cost_code(value) -> str:
    """Extrae solo el número del centro de costo, p.ej. '20136 - BARRICK ETP-HL' → '20136'."""
    if not value or pd.isna(value):
        return ""
    s = str(value).strip()
    if not s or s in ("nan", "None"):
        return ""
    # Buscar primer separador (espacio o " - ")
    m = re.match(r"^(\S+)(?:\s*-\s*.*)?$", s)
    if m:
        return m.group(1)
    return s


def format_date_yyyymmdd(date_str) -> str:
    """Convierte fechas tipo '2025-08-04 00:00:00' a '2025/08/04'."""
    if not date_str or pd.isna(date_str):
        return ""
    s = str(date_str).strip()
    if not s or s in ("nan", "None"):
        return ""
    date_part = s.split(" ")[0]
    return date_part.replace("-", "/")


# ------------------------------------------------------------
# 4.1 Helpers de formato de celdas
# ------------------------------------------------------------
def _copy_cell_format(src_cell, dst_cell):
    if src_cell.font:
        dst_cell.font = copy.copy(src_cell.font)
    if src_cell.alignment:
        dst_cell.alignment = copy.copy(src_cell.alignment)
    if src_cell.number_format:
        dst_cell.number_format = src_cell.number_format
    if src_cell.border:
        dst_cell.border = copy.copy(src_cell.border)


def _apply_thin_border(cell, left=True, right=True, top=False, bottom=False):
    thin = Side(style="thin")
    cell.border = Border(
        left=thin if left else Side(),
        right=thin if right else Side(),
        top=thin if top else Side(),
        bottom=thin if bottom else Side(),
    )


def _apply_green_fill(cell):
    cell.fill = PatternFill(fill_type="solid", fgColor="92D050")


def _apply_no_fill(cell):
    cell.fill = PatternFill(fill_type=None)


def _find_dataframe_column(df: pd.DataFrame, candidates: list[str]) -> str | None:
    lower_map = {str(col).strip().lower(): col for col in df.columns}
    for candidate in candidates:
        candidate_lower = candidate.strip().lower()
        if candidate_lower in lower_map:
            return lower_map[candidate_lower]
    for candidate in candidates:
        candidate_lower = candidate.strip().lower()
        for col in df.columns:
            if candidate_lower in str(col).strip().lower():
                return col
    return None


def _value_is_present(value) -> bool:
    if value is None:
        return False
    if isinstance(value, float) and pd.isna(value):
        return False
    text = str(value).strip()
    return text not in ("", "nan", "None")


def _normalized_text(value) -> str:
    """Normaliza texto y formatea números como dos dígitos (para riesgos)."""
    if not _value_is_present(value):
        return ""
    text = str(value).strip()
    text = re.sub(r"\s+", " ", text)
    if re.fullmatch(r"\d+(?:\.0+)?", text):
        try:
            return f"{int(float(text)):02d}"
        except Exception:
            return text
    return text.upper()


def _selected_columns(df: pd.DataFrame, preferred: list[str], fallback: int = 10) -> list[str]:
    existing = [col for col in preferred if col in df.columns]
    if existing:
        return existing
    return list(df.columns[:fallback])


def _row_from_df(row: pd.Series, columns: list[str]) -> dict:
    result = {}
    for col in columns:
        value = row.get(col, "")
        result[col] = "" if pd.isna(value) else value
    return result


def _first_present_value(row: pd.Series, candidates: list[str]) -> str:
    for candidate in candidates:
        if candidate in row.index:
            value = row.get(candidate, "")
            if _value_is_present(value):
                return str(value).strip()
    return ""


# ------------------------------------------------------------
# Columnas que se almacenan como texto (número_format = "@")
# ------------------------------------------------------------
_TEXT_FORMAT_HEADERS = {
    "identificación", "c.cost", "c. cost", "ccost", "costo",
    "inicio vigencia", "fecha nac", "fecha nac.", "riesgo", "clase de riesgo",
}


def _header_needs_text_format(header: str) -> bool:
    h = header.strip().lower()
    return any(key in h for key in _TEXT_FORMAT_HEADERS)


# ------------------------------------------------------------
# 4.2 Escritura de tablas en la hoja Reporte
# ------------------------------------------------------------
def _write_table_section(
    ws,
    start_row: int,
    title: str,
    headers: list[str],
    rows: list[dict],
) -> int:
    from openpyxl.utils import get_column_letter

    if not headers:
        headers = ["Sin datos"]

    base_font = Font(name="Aptos Narrow", size=12)
    bold_font = Font(name="Aptos Narrow", size=12, bold=True)
    thin = Side(style="thin")

    # ── Título ──────────────────────────────────────────────
    title_cell = ws.cell(row=start_row, column=1)
    title_cell.value = title
    title_cell.font = bold_font
    title_cell.fill = PatternFill(fill_type=None)
    title_cell.alignment = Alignment(horizontal="left", vertical="center")

    # ── Encabezados ─────────────────────────────────────────
    header_row = start_row + 1
    for col_idx, header in enumerate(headers, start=1):
        cell = ws.cell(row=header_row, column=col_idx)
        cell.value = header
        cell.font = bold_font
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = Border(left=thin, right=thin, top=thin, bottom=thin)

    # ── Datos ────────────────────────────────────────────────
    data_start = header_row + 1
    if not rows:
        cell = ws.cell(row=data_start, column=1)
        cell.value = "Sin registros"
        cell.font = base_font
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = Border(left=thin, right=thin, top=thin, bottom=thin)
        return data_start + 2

    for row_offset, row_data in enumerate(rows):
        excel_row = data_start + row_offset
        for col_idx, header in enumerate(headers, start=1):
            cell = ws.cell(row=excel_row, column=col_idx)
            value = row_data.get(header, "")
            cell.value = str(value) if _value_is_present(value) else None
            cell.font = base_font
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            cell.border = Border(left=thin, right=thin, top=thin, bottom=thin)
            # Guardar como texto las columnas que lo requieren
            if _header_needs_text_format(header):
                cell.number_format = "@"

    last_row = data_start + len(rows) - 1

    # Ajustar ancho de columnas
    for col_idx, header in enumerate(headers, start=1):
        width = max(len(str(header)), 14)
        for row_idx in range(data_start, last_row + 1):
            value = ws.cell(row=row_idx, column=col_idx).value
            if value is not None:
                width = max(width, min(len(str(value)), 42))
        width_map = {1: 16, 2: 36, 3: 16, 4: 12, 5: 18, 6: 18}
        ws.column_dimensions[get_column_letter(col_idx)].width = width_map.get(col_idx, min(width + 2, 45))

    return data_start + len(rows) + 2


# ------------------------------------------------------------
# 4.3 Construcción de la hoja "Reporte"
# ------------------------------------------------------------

def _safe_get(row: pd.Series, *candidates: str) -> str:
    """
    Devuelve el primer valor presente buscando los candidatos
    exactamente (case-insensitive) en el índice de la fila.
    """
    idx_lower = {str(k).strip().lower(): k for k in row.index}
    for c in candidates:
        real_key = idx_lower.get(c.strip().lower())
        if real_key is not None:
            v = row.get(real_key, "")
            if _value_is_present(v):
                return str(v).strip()
    return ""


def _normalize_cost_code(raw: str) -> str:
    """
    Extrae el código numérico del centro de costo y elimina decimales.
    '20136 - BARRICK ETP-HL' → '20136'
    '20136.0'                → '20136'
    """
    code = extract_cost_code(raw)
    if code:
        try:
            return str(int(float(code)))
        except (ValueError, TypeError):
            return code.strip()
    return ""


def _build_report_sheet(wb, df_rep: pd.DataFrame, df_trab: pd.DataFrame):
    if "Reporte" in wb.sheetnames:
        wb.remove(wb["Reporte"])
    ws = wb.create_sheet("Reporte")
    ws.sheet_view.showGridLines = False

    rep_df  = df_rep.copy()
    trab_df = df_trab.copy()

    # ── Diagnóstico de columnas disponibles ──────────────────────────────
    print("[Reporte] Columnas rep_df :", list(rep_df.columns))
    print("[Reporte] Columnas trab_df:", list(trab_df.columns))

    # ── Detectar columnas para cruce de IDs ──────────────────────────────
    rep_id_col  = _find_dataframe_column(rep_df,  ["Cédula identificación", "Cédula", "Identificación", "Documento", "ID"])
    trab_id_col = _find_dataframe_column(trab_df, ["ID_Num", "Identificación", "Documento", "Cédula"])
    print(f"[Reporte] rep_id_col={rep_id_col}  trab_id_col={trab_id_col}")

    rep_nombre_col  = _find_dataframe_column(rep_df,  ["Apellidos, Nombre", "Nombre"])
    trab_nombre_col = _find_dataframe_column(trab_df, ["Nombre", "Apellidos, Nombre"])

    rep_inicio_col  = _find_dataframe_column(rep_df,  ["Inicio Vigencia", "Fecha Ingreso", "F. Ingreso"])
    trab_inicio_col = _find_dataframe_column(trab_df, ["Inicio Vigencia"])
    print(f"[Reporte] rep_inicio_col={rep_inicio_col}  trab_inicio_col={trab_inicio_col}")

    rep_costo_col  = _find_dataframe_column(rep_df,  ["C.COSTO", "C. COSTO", "COSTO", "C. Cost", "C.Cost"])
    trab_costo_col = _find_dataframe_column(trab_df, ["C.COSTO", "C. COSTO", "COSTO", "C. Cost", "C.Cost"])
    print(f"[Reporte] rep_costo_col={rep_costo_col}  trab_costo_col={trab_costo_col}")

    risk_rep_col  = _find_dataframe_column(rep_df,  ["NIVEL ARL", "Riesgo ARL", "RIESGO EN ARL"])
    risk_trab_col = _find_dataframe_column(trab_df, ["Riesgo ARL", "RIESGO EN ARL", "LIBRA"])
    print(f"[Reporte] risk_rep_col={risk_rep_col}  risk_trab_col={risk_trab_col}")

    baja_col = _find_dataframe_column(rep_df, ["Fecha Última Baja", "Fecha Ultima Baja", "ULTIMA BAJA"])
    print(f"[Reporte] baja_col={baja_col}")

    # ── Normalizar IDs ────────────────────────────────────────────────────
    rep_df["_id_norm"]  = rep_df[rep_id_col].apply(normalize_id)  if rep_id_col  else ""
    trab_df["_id_norm"] = trab_df[trab_id_col].apply(normalize_id) if trab_id_col else ""

    rep_by_id  = {r["_id_norm"]: r for _, r in rep_df.iterrows()  if _value_is_present(r["_id_norm"])}
    trab_by_id = {r["_id_norm"]: r for _, r in trab_df.iterrows() if _value_is_present(r["_id_norm"])}

    # ── Helpers de extracción ─────────────────────────────────────────────
    def get_id_rep(row):
        return _safe_get(row, "Cédula identificación", "Cédula", "Identificación", "Documento", "ID")

    def get_id_trab(row):
        return _safe_get(row, "ID_Num", "Identificación", "Documento", "Cédula")

    def get_nombre_rep(row):
        return _safe_get(row, "Apellidos, Nombre", "Nombre")

    def get_nombre_trab(row):
        return _safe_get(row, "Nombre", "Apellidos, Nombre")

    def get_inicio_rep(row):
        return format_date_yyyymmdd(_safe_get(row, "Inicio Vigencia", "Fecha Ingreso", "F. Ingreso"))

    def get_inicio_trab(row):
        return format_date_yyyymmdd(_safe_get(row, "Inicio Vigencia"))

    def get_costo_rep(row):
        return _normalize_cost_code(_safe_get(row, "C.COSTO", "C. COSTO", "COSTO", "C. Cost", "C.Cost"))

    def get_costo_trab(row):
        return _normalize_cost_code(_safe_get(row, "C.COSTO", "C. COSTO", "COSTO", "C. Cost", "C.Cost"))

    def get_riesgo_rep(row):
        return format_riesgo_arl(_safe_get(row, "NIVEL ARL", "Riesgo ARL", "RIESGO EN ARL"))

    def get_riesgo_trab(row):
        return format_riesgo_arl(_safe_get(row, "Riesgo ARL", "RIESGO EN ARL", "LIBRA"))

    # ── Tabla 1: ACTIVOS en ARL pero INACTIVOS en Nómina ─────────────────
    table_1_rows = []
    for _, row in rep_df.iterrows():
        if baja_col and _value_is_present(row.get(baja_col, "")):
            id_norm  = row.get("_id_norm", "")
            trab_row = trab_by_id.get(id_norm, pd.Series(dtype=object))
            id_val = get_id_trab(trab_row) or get_id_rep(row)
            nombre = get_nombre_trab(trab_row) or get_nombre_rep(row)
            inicio = get_inicio_trab(trab_row) or get_inicio_rep(row)
            costo  = get_costo_trab(trab_row)  or get_costo_rep(row)
            riesgo = get_riesgo_trab(trab_row) or get_riesgo_rep(row)
            table_1_rows.append({
                "Identificación":      id_val,
                "Nombre":              nombre,
                "Inicio Vigencia":     inicio,
                "C. Cost":             costo,
                "Clase de Riesgo ARL": riesgo,
            })

    # ── Tabla 2: Diferente nivel de riesgo ARL vs Nómina ─────────────────
    table_2_rows = []
    for _, row in rep_df.iterrows():
        id_norm = row.get("_id_norm", "")
        if not _value_is_present(id_norm):
            continue
        trab_row = trab_by_id.get(id_norm)
        if trab_row is None:
            continue
        rep_risk  = _normalized_text(get_riesgo_rep(row))
        trab_risk = _normalized_text(get_riesgo_trab(trab_row))
        if rep_risk and trab_risk and rep_risk != trab_risk:
            inicio = get_inicio_trab(trab_row) or get_inicio_rep(row)
            costo  = get_costo_trab(trab_row)  or get_costo_rep(row)
            table_2_rows.append({
                "Identificación":      get_id_trab(trab_row) or get_id_rep(row),
                "Nombre":              get_nombre_trab(trab_row) or get_nombre_rep(row),
                "Inicio Vigencia":     inicio,
                "C. Cost":             costo,
                "Clase de Riesgo ARL": rep_risk,
                "Riesgo ARL NOMINA":   trab_risk,
            })

    # ── Precalcular dict id → C.COSTO desde rep_df (fuente = EMP col CL C.COSTO) ──
    # La fórmula original en Cruce ARL col K era:
    #   =SI.ERROR(EXTRAE(BUSCARV(B2;EMP!$A:$CK;89;0);1;5);0)
    # Es decir: busca el ID del trabajador en EMP y extrae los primeros chars de C.COSTO.
    # Como rep_df ES la fuente de EMP, cruzamos trab_df con rep_df por ID y tomamos C.COSTO de rep_df.
    cruce_id_to_ccosto: dict[str, str] = {}
    for _, _rp in rep_df.iterrows():
        _id = _rp.get("_id_norm", "")
        if not _value_is_present(_id):
            continue
        _raw = _safe_get(_rp, "C.COSTO", "C. COSTO", "COSTO", "C. Cost", "C.Cost")
        _codigo = _normalize_cost_code(_raw)
        cruce_id_to_ccosto[_id] = _codigo

    # ── Índice de EMP: ID (col A=1) → fila Excel ────────────────────────
    ws_emp = wb["EMP"]
    emp_id_to_row: dict[str, int] = {}
    for r in range(2, ws_emp.max_row + 1):
        id_cell = ws_emp.cell(row=r, column=1).value
        if id_cell is not None:
            emp_id_to_row[normalize_id(str(id_cell))] = r

    def _cruce_costo(id_norm: str) -> str:
        """Obtiene C.COSTO para el ID desde el dict precalculado de trab_df."""
        return cruce_id_to_ccosto.get(id_norm, "")

    def _emp_fecha_ant(id_norm: str) -> str:
        """Lee Fecha Antigüedad de la celda DC (col 107) en EMP para ese ID."""
        r = emp_id_to_row.get(id_norm)
        if r is None:
            return ""
        val = ws_emp.cell(row=r, column=107).value
        return format_date_yyyymmdd(str(val)) if _value_is_present(val) else ""

    # ── Tabla 3: ACTIVOS en ARL, NO están en NÓMINA ──────────────────────
    table_3_rows = []
    for _, row in trab_df.iterrows():
        id_norm = row.get("_id_norm", "")
        if not (_value_is_present(id_norm) and id_norm not in rep_by_id):
            continue
        id_val = get_id_trab(row)
        nombre = get_nombre_trab(row)
        inicio = get_inicio_trab(row)
        costo  = _cruce_costo(id_norm)
        riesgo = get_riesgo_trab(row)
        print(f"[T3] id={id_val} inicio={inicio!r} costo={costo!r} riesgo={riesgo!r}")
        table_3_rows.append({
            "Identificación":    id_val,
            "Nombre":            nombre,
            "Inicio Vigencia":   inicio,
            "C. Cost":           costo,
            "Riesgo ARL NOMINA": riesgo,
        })

    # ── Tabla 4: ACTIVOS en Nómina, NO están en ARL ──────────────────────
    table_4_rows = []
    for _, row in rep_df.iterrows():
        id_norm = row.get("_id_norm", "")
        if not (_value_is_present(id_norm) and id_norm not in trab_by_id):
            continue
        id_val = get_id_rep(row)
        nombre = get_nombre_rep(row)
        inicio = _emp_fecha_ant(id_norm)
        costo  = get_costo_rep(row)
        riesgo = get_riesgo_rep(row)
        print(f"[T4] id={id_val} inicio={inicio!r} costo={costo!r} riesgo={riesgo!r}")
        table_4_rows.append({
            "Identificación":    id_val,
            "Nombre":            nombre,
            "Inicio Vigencia":   inicio,
            "C. Cost":           costo,
            "Riesgo ARL NOMINA": riesgo,
        })

    # ── Escribir tablas: la primera comienza en fila 6 ───────────────────
    current_row = 6
    current_row = _write_table_section(
        ws, current_row,
        "Empleados ACTIVOS en ARL, pero INACTIVOS en Nomina",
        ["Identificación", "Nombre", "Inicio Vigencia", "C. Cost", "Clase de Riesgo ARL"],
        table_1_rows,
    )
    current_row = _write_table_section(
        ws, current_row,
        "Empleados con diferente nivel de riesgo en ARL Vs Nomina",
        ["Identificación", "Nombre", "Inicio Vigencia", "C. Cost", "Clase de Riesgo ARL", "Riesgo ARL NOMINA"],
        table_2_rows,
    )
    current_row = _write_table_section(
        ws, current_row,
        "Empleados ACTIVOS en ARL, no estan en NOMINA",
        ["Identificación", "Nombre", "Inicio Vigencia", "C. Cost", "Riesgo ARL NOMINA"],
        table_3_rows,
    )
    _write_table_section(
        ws, current_row,
        "Empleados ACTIVO en nomina, pero NO ESTA en ARL",
        ["Identificación", "Nombre", "Inicio Vigencia", "C. Cost", "Riesgo ARL NOMINA"],
        table_4_rows,
    )

    ws.freeze_panes = "A1"


# ------------------------------------------------------------
# 5. Generación del archivo INFRA final
# ------------------------------------------------------------
def generate_infra(infra_bytes: bytes, df_rep: pd.DataFrame, df_trab: pd.DataFrame) -> bytes:
    wb = load_workbook(io.BytesIO(infra_bytes), keep_vba=False)

    if "EMP" not in wb.sheetnames:
        raise Exception("La hoja 'EMP' no existe en INFRA")
    if "Cruce ARL" not in wb.sheetnames:
        raise Exception("La hoja 'Cruce ARL' no existe en INFRA")

    cedula_col = _find_dataframe_column(
        df_rep,
        ["Cédula identificación", "Cédula", "Identificación", "Documento", "ID"],
    )
    if cedula_col is None:
        cedula_col = df_rep.columns[0]
    df_rep["_cedula_norm"] = df_rep[cedula_col].apply(normalize_id)
    rep_dict = {row["_cedula_norm"]: row for _, row in df_rep.iterrows() if row["_cedula_norm"]}

    df_trab["_id_norm"] = df_trab["ID_Num"].apply(normalize_id)
    trab_dict = {row["_id_norm"]: row for _, row in df_trab.iterrows() if row["_id_norm"]}

    _fill_cruce_sheet(wb, df_trab, rep_dict)
    _fill_emp_sheet(wb, df_rep, trab_dict)
    _build_report_sheet(wb, df_rep, df_trab)

    # ── FIX 1: Eliminar hojas extra usando lista estática para evitar mutación en iteración ──
    KEEP_SHEETS = {"Cruce ARL", "EMP", "Reporte"}
    for sheet_name in list(wb.sheetnames):          # list() congela los nombres antes de borrar
        if sheet_name not in KEEP_SHEETS:
            del wb[sheet_name]
            print(f"[generate_infra] Hoja eliminada: {sheet_name!r}")

    # Verificar que las tres hojas requeridas existen
    for required in ("Cruce ARL", "EMP", "Reporte"):
        if required not in wb.sheetnames:
            raise Exception(f"La hoja requerida '{required}' no existe después de procesar")

    out = io.BytesIO()
    wb.save(out)
    out.seek(0)
    return out.read()


# ------------------------------------------------------------
# 5.1 Llenar hoja "Cruce ARL"
# ------------------------------------------------------------
def _fill_cruce_sheet(wb, df_trab: pd.DataFrame, rep_dict: dict):
    ws = wb["Cruce ARL"]
    DATA_START = 2
    MAX_DATA_COL = 13

    col_index = {
        "Tipo": 1, "ID_Num": 2, "Nombre": 3, "Cargo": 4,
        "Inicio Vigencia": 5, "EPS": 6, "AFP": 7, "Salario": 8,
        "Fecha Nac.": 9, "Riesgo ARL": 10,
    }

    total = len(df_trab)

    ref_fonts = {}
    ref_alignments = {}
    ref_numfmts = {}
    ref_formulas = {}

    _thin = Side(style="thin")
    _thin_border = Border(left=_thin, right=_thin, top=_thin, bottom=_thin)

    for c in range(1, MAX_DATA_COL + 1):
        src = ws.cell(row=2, column=c)
        ref_fonts[c] = copy.copy(src.font) if src.font else None
        ref_alignments[c] = copy.copy(src.alignment) if src.alignment else None
        ref_numfmts[c] = src.number_format
        if c in (11, 12, 13) and src.value and str(src.value).startswith("="):
            ref_formulas[c] = str(src.value)

    # Normalizar separadores: openpyxl requiere comas (sintaxis internacional)
    ref_formulas = {k: v.replace(";", ",") for k, v in ref_formulas.items()}

    # Columna K (11) = C.COSTO: valor directo desde trab_df (no VLOOKUP).
    # openpyxl no evalúa fórmulas, así que guardamos el valor real.
    ref_formulas.pop(11, None)

    _re_row2 = re.compile(r"([A-Z]+)2\b")

    records = df_trab.reset_index(drop=True)

    for i in range(total):
        excel_row = DATA_START + i
        row = records.iloc[i]

        for c in range(1, MAX_DATA_COL + 1):
            cell = ws.cell(row=excel_row, column=c)
            if ref_fonts[c]:
                cell.font = ref_fonts[c]
            if ref_alignments[c]:
                cell.alignment = ref_alignments[c]
            if ref_numfmts[c]:
                cell.number_format = ref_numfmts[c]
            cell.border = _thin_border

        for field, col in col_index.items():
            val = row.get(field, "")
            cell = ws.cell(row=excel_row, column=col)
            if field == "ID_Num":
                cell.value = _id_to_int(val)
            elif field == "Riesgo ARL":
                cell.value = format_riesgo_arl(val) if val else None
                cell.number_format = "@"
            elif field == "Tipo":
                cell.value = val if val else None
                cell.number_format = "@"
            else:
                cell.value = val if val else None

        # ── C.COSTO → col K (11): fórmula =SI.ERROR(EXTRAE(BUSCARV(B{row};EMP!$A:$CK;89;0);1;5);0) ──
        cell_k = ws.cell(row=excel_row, column=11)
        cell_k.value = f"=IFERROR(MID(VLOOKUP(B{excel_row},EMP!$A:$CK,89,0),1,5),0)"
        cell_k.number_format = "@"

        for col_idx, formula in ref_formulas.items():
            cell = ws.cell(row=excel_row, column=col_idx)
            existing = cell.value
            if not existing or not str(existing).startswith("="):
                cell.value = _re_row2.sub(lambda m: f"{m.group(1)}{excel_row}", formula)

    # Limpiar filas sobrantes
    for r in range(DATA_START + total, ws.max_row + 1):
        for c in range(1, MAX_DATA_COL + 1):
            cell = ws.cell(row=r, column=c)
            cell.value = None
            cell.fill = PatternFill(fill_type=None)

    last_data_row = DATA_START + total - 1
    validacion_range = f"M{DATA_START}:M{last_data_row}"
    green_fill = PatternFill(patternType=None, fgColor="00000000", bgColor="C6EFCE")
    green_font = Font(color="006100")
    dxf_ok = DifferentialStyle(fill=green_fill, font=green_font)
    rule_ok = Rule(type="containsText", operator="containsText", text="OK", dxf=dxf_ok)
    rule_ok.formula = [f'NOT(ISERROR(SEARCH("OK",M{DATA_START})))']
    ws.conditional_formatting._cf_rules.clear()
    ws.conditional_formatting.add(validacion_range, rule_ok)

    print(f"Cruce ARL actualizado: {total} filas, CF verde en {validacion_range}")


# ------------------------------------------------------------
# 5.2 Llenar hoja "EMP"
# ------------------------------------------------------------
def _fill_emp_sheet(wb, df_rep: pd.DataFrame, trab_dict: dict):
    from openpyxl.utils import get_column_letter
    ws = wb["EMP"]
    HEADER_ROW = 1
    DATA_START = 2

    col_c_costo = None
    col_validacion = None
    col_name_to_idx = {}
    for c in range(1, ws.max_column + 1):
        val = ws.cell(row=HEADER_ROW, column=c).value
        if val:
            col_name_to_idx[str(val).strip()] = c
            if str(val).strip() == "C.COSTO":
                col_c_costo = c
            if str(val).strip() == "VALIDACION":
                col_validacion = c

    MAX_DATA_COL = max(col_name_to_idx.values()) if col_name_to_idx else 140

    ref_fonts = {}
    ref_alignments = {}
    ref_numfmts = {}
    ref_borders_lr = {}
    formula_cols: dict[int, str] = {}

    for c in range(1, MAX_DATA_COL + 1):
        src = ws.cell(row=DATA_START, column=c)
        ref_fonts[c] = copy.copy(src.font) if src.font else None
        ref_alignments[c] = copy.copy(src.alignment) if src.alignment else None
        ref_numfmts[c] = src.number_format

        b = src.border
        ref_borders_lr[c] = Border(
            left=copy.copy(b.left) if b.left.style else Side(),
            right=copy.copy(b.right) if b.right.style else Side(),
            top=Side(),
            bottom=Side(),
        )

        if src.value and str(src.value).startswith("="):
            formula_cols[c] = str(src.value).replace(";", ",")

    reporte_to_emp: dict[str, int] = {}
    for col in df_rep.columns:
        col_norm = col.strip()
        if col_norm in col_name_to_idx:
            reporte_to_emp[col] = col_name_to_idx[col_norm]
        else:
            for emp_col, idx in col_name_to_idx.items():
                if emp_col.strip().lower() == col_norm.lower():
                    reporte_to_emp[col] = idx
                    break

    # 94 = NIVEL ARL (se escribe manualmente)
    # 107/108/109 = DC/DD/DE fechas (se escriben manualmente con formato)
    skip_cols = set(formula_cols.keys()) | {94, 107, 108, 109}

    _re_row2 = re.compile(r"([A-Z]+)2\b")

    n_rows = len(df_rep)
    records = df_rep.reset_index(drop=True)

    for i in range(n_rows):
        excel_row = DATA_START + i
        rep_row = records.iloc[i]

        if excel_row != DATA_START:
            for c in range(1, MAX_DATA_COL + 1):
                cell = ws.cell(row=excel_row, column=c)
                if ref_fonts[c]:
                    cell.font = ref_fonts[c]
                if ref_alignments[c]:
                    cell.alignment = ref_alignments[c]
                if ref_numfmts[c]:
                    cell.number_format = ref_numfmts[c]
                cell.border = ref_borders_lr[c]

        for rep_col, col_idx in reporte_to_emp.items():
            if col_idx in skip_cols:
                continue
            cell = ws.cell(row=excel_row, column=col_idx)
            if cell.value and str(cell.value).startswith("="):
                continue
            val = rep_row.get(rep_col)
            cell.value = val if not pd.isna(val) else None

        # ── C.COSTO: valor completo en EMP ───────────────────────────────
        # En EMP se conserva el valor completo (ej. "20136 - BARRICK ETP-HL")
        if col_c_costo is not None:
            ccosto_val = _safe_get(rep_row, "C.COSTO", "C. COSTO", "COSTO", "C. Cost", "C.Cost")
            if ccosto_val:
                cell = ws.cell(row=excel_row, column=col_c_costo)
                cell.value = ccosto_val
                cell.number_format = "@"

        # ── Código (col B = 2) ────────────────────────────────────
        codigo_val = rep_row.get("Código", "")
        if codigo_val and str(codigo_val) not in ("nan", "None", ""):
            cell = ws.cell(row=excel_row, column=2)
            cell.value = format_codigo_emp(codigo_val)
            cell.number_format = "@"

        # ── NIVEL ARL (col 94) ────────────────────────────────────
        nivel_arl = format_nivel_arl(_safe_get(rep_row, "NIVEL ARL", "Riesgo ARL", "RIESGO EN ARL"))
        ws.cell(row=excel_row, column=94).value = nivel_arl

        # ── Fechas: DC (107), DD (108), DE (109) ──────────────────
        val_alta = _safe_get(rep_row, "Fecha Alta", "FECHA ALTA")
        cell_dc = ws.cell(row=excel_row, column=107)
        cell_dc.value = format_date_yyyymmdd(val_alta) or None
        cell_dc.number_format = "@"

        val_baja = _safe_get(rep_row, "Fecha Última Baja", "Fecha Ultima Baja", "FECHA ULTIMA BAJA")
        cell_dd = ws.cell(row=excel_row, column=108)
        cell_dd.value = format_date_yyyymmdd(val_baja) or None
        cell_dd.number_format = "@"

        val_ant = _safe_get(rep_row, "Fecha Antigüedad", "FECHA ANTIGUEDAD")
        if not val_ant:
            val_ant = _safe_get(rep_row, "Inicio Vigencia", "Fecha Ingreso", "F. Ingreso")
        cell_de = ws.cell(row=excel_row, column=109)
        cell_de.value = format_date_yyyymmdd(val_ant) or None
        cell_de.number_format = "@"

        # ── Replicar fórmulas ─────────────────────────────────────
        for col_idx, ref_formula in formula_cols.items():
            cell = ws.cell(row=excel_row, column=col_idx)
            if not cell.value or not str(cell.value).startswith("="):
                cell.value = _re_row2.sub(lambda m: f"{m.group(1)}{excel_row}", ref_formula)

    # Formato condicional en VALIDACION
    if col_validacion:
        last_data_row = DATA_START + n_rows - 1
        val_col_letter = get_column_letter(col_validacion)
        val_range = f"{val_col_letter}{DATA_START}:{val_col_letter}{last_data_row}"
        green_fill = PatternFill(patternType=None, fgColor="00000000", bgColor="C6EFCE")
        green_font = Font(color="006100")
        dxf_ok = DifferentialStyle(fill=green_fill, font=green_font)
        rule_ok = Rule(type="containsText", operator="containsText", text="OK", dxf=dxf_ok)
        rule_ok.formula = [f'NOT(ISERROR(SEARCH("OK",{val_col_letter}{DATA_START})))']
        ws.conditional_formatting._cf_rules.clear()
        ws.conditional_formatting.add(val_range, rule_ok)
        print(f"EMP actualizada: {n_rows} filas, CF verde en {val_range}")
    else:
        print(f"EMP actualizada: {n_rows} filas (columna VALIDACION no encontrada)")


# ------------------------------------------------------------
# 6. Gestión de estado sin base de datos
# ------------------------------------------------------------
_TEMP_DIR = tempfile.gettempdir()
_TOKEN_COOKIE = "infra_token"


def _state_path(token: str) -> str:
    safe = re.sub(r"[^a-zA-Z0-9_-]", "", token)
    return os.path.join(_TEMP_DIR, f"infra_state_{safe}.json")


def _infra_path(token: str) -> str:
    safe = re.sub(r"[^a-zA-Z0-9_-]", "", token)
    return os.path.join(_TEMP_DIR, f"infra_bytes_{safe}.bin")


def save_state(token: str, df_rep: pd.DataFrame, df_trab: pd.DataFrame, infra_bytes: bytes):
    state = {
        "reporte_json": df_rep.to_json(orient="records", date_format="iso", default_handler=str),
        "trabajadores_json": df_trab.to_json(orient="records", date_format="iso", default_handler=str),
    }
    with open(_state_path(token), "w", encoding="utf-8") as f:
        json.dump(state, f)
    with open(_infra_path(token), "wb") as f:
        f.write(infra_bytes)


def load_state(token: str):
    sp = _state_path(token)
    ip = _infra_path(token)
    if not os.path.exists(sp) or not os.path.exists(ip):
        raise FileNotFoundError("Archivos temporales no encontrados")

    with open(sp, "r", encoding="utf-8") as f:
        state = json.load(f)
    with open(ip, "rb") as f:
        infra_bytes = f.read()

    df_rep = pd.read_json(io.StringIO(state["reporte_json"]), orient="records")
    df_trab = pd.read_json(io.StringIO(state["trabajadores_json"]), orient="records")
    return df_rep, df_trab, infra_bytes


def delete_state(token: str):
    for path in (_state_path(token), _infra_path(token)):
        try:
            os.remove(path)
        except OSError:
            pass