import io
import re
import unicodedata
from typing import Dict, Optional, Tuple, List

import pandas as pd
import streamlit as st

# -------------------- Configuración de página --------------------
st.set_page_config(page_title="Consolidar asistencia", layout="wide")
st.title("📊 Consolidar hoja 'asistencia' + limpieza de RUT y fechas")

st.write("""
Sube varios archivos Excel. La app:
1) Busca la hoja **'asistencia'** en cada archivo (sin importar mayúsculas).
2) Limpia encabezados "raros" (promueve la primera fila si corresponde).
3) **Estandariza columnas** (por ejemplo, `RUT`, `Rut`, `rut trabajador` → `rut`).
4) **Normaliza y valida RUT** (formato, dígito verificador).
5) **Parsea fechas** y las normaliza a `AAAA-MM-DD`.
6) Consolida todo en una tabla final y permite descargar el Excel.
""")

uploaded_files = st.file_uploader(
    "Arrastra o selecciona tus archivos (.xls/.xlsx)",
    type=["xls", "xlsx"],
    accept_multiple_files=True
)

# -------------------- Utilidades de texto/columnas --------------------
def strip_accents(s: str) -> str:
    return "".join(
        c for c in unicodedata.normalize("NFD", s)
        if unicodedata.category(c) != "Mn"
    )

def normalize_colname(c: str) -> str:
    c = str(c).strip()
    c = strip_accents(c)
    c = c.lower()
    c = re.sub(r"[\.\-_/]+", " ", c)
    c = re.sub(r"\s+", " ", c).strip()
    return c

# Diccionario de sinónimos -> target estandar
COLUMN_SYNONYMS: Dict[str, str] = {
    # RUT
    "rut": "rut", "rut trabajador": "rut", "rut empleado": "rut",
    "r u t": "rut", "id rut": "rut", "identificacion": "rut",
    "dni": "rut", "documento": "rut",
    # nombre
    "nombre": "nombre", "nombre trabajador": "nombre", "nombre empleado": "nombre",
    "nombres": "nombre", "nombre y apellido": "nombre", "apellido y nombre": "nombre",
    # fecha
    "fecha": "fecha", "dia": "fecha", "día": "fecha", "date": "fecha",
    "fecha asistencia": "fecha", "f asistencia": "fecha", "fec": "fecha",
    # entrada
    "entrada": "entrada", "hora entrada": "entrada", "ingreso": "entrada",
    # salida
    "salida": "salida", "hora salida": "salida", "egreso": "salida",
    # otros útiles
    "empresa": "empresa", "compania": "empresa", "compañia": "empresa",
    "proyecto": "proyecto", "faena": "proyecto", "obra": "proyecto",
    "centro de costo": "centro_costo", "centro costo": "centro_costo", "cc": "centro_costo",
    "turno": "turno", "jornada": "turno",
    "ubicacion": "ubicacion", "sede": "ubicacion", "planta": "ubicacion",
    "observaciones": "observaciones", "obs": "observaciones", "comentarios": "observaciones",
}

TARGET_ORDER = [
    "rut", "nombre", "fecha", "entrada", "salida",
    "empresa", "proyecto", "centro_costo", "turno", "ubicacion", "observaciones", "source_file", "error"
]

def map_columns(df: pd.DataFrame) -> pd.DataFrame:
    original_cols = list(df.columns)
    new_cols: List[str] = []
    for c in original_cols:
        norm = normalize_colname(str(c))
        target = COLUMN_SYNONYMS.get(norm)
        if target:
            new_cols.append(target)
        else:
            # intentar matches parciales
            if any(k in norm for k in ["rut", "r u t"]):
                new_cols.append("rut")
            elif any(k in norm for k in ["nombre", "apellido"]):
                new_cols.append("nombre")
            elif any(k in norm for k in ["fecha", "dia", "día", "date", "fec"]):
                new_cols.append("fecha")
            elif "entrada" in norm or "ingreso" in norm:
                new_cols.append("entrada")
            elif "salida" in norm or "egreso" in norm:
                new_cols.append("salida")
            elif "empresa" in norm or "compania" in norm or "compania" in norm:
                new_cols.append("empresa")
            elif "proyecto" in norm or "faena" in norm or "obra" in norm:
                new_cols.append("proyecto")
            elif "centro" in norm and "costo" in norm:
                new_cols.append("centro_costo")
            elif "turno" in norm or "jornada" in norm:
                new_cols.append("turno")
            elif "ubic" in norm or "sede" in norm or "planta" in norm:
                new_cols.append("ubicacion")
            elif "observa" in norm or "coment" in norm or "obs" in norm:
                new_cols.append("observaciones")
            else:
                new_cols.append(str(c))  # dejar como está
    df.columns = new_cols
    # limpiar columnas vacías
    df = df.dropna(how="all").dropna(axis=1, how="all")
    return df

# -------------------- Validación y normalización de RUT --------------------
def clean_rut_text(r: str) -> str:
    if pd.isna(r):
        return ""
    r = str(r).strip().upper()
    r = r.replace(".", "").replace(" ", "")
    # Cambiar K en DV a K mayúscula (ya lo está). Aceptar guion o no.
    return r

def split_rut(r: str) -> Tuple[str, str]:
    # Devuelve (cuerpo, dv) o ("","") si no válido
    if not r:
        return "", ""
    if "-" in r:
        cuerpo, dv = r.split("-", 1)
    else:
        cuerpo, dv = r[:-1], r[-1:]
    return cuerpo, dv

def compute_dv(cuerpo: str) -> Optional[str]:
    try:
        nums = list(map(int, reversed(cuerpo)))
    except ValueError:
        return None
    factors = [2,3,4,5,6,7]
    s = 0
    for i, n in enumerate(nums):
        s += n * factors[i % len(factors)]
    mod = 11 - (s % 11)
    if mod == 11:
        return "0"
    if mod == 10:
        return "K"
    return str(mod)

def normalize_and_validate_rut(r: str) -> Tuple[str, bool]:
    r = clean_rut_text(r)
    cuerpo, dv = split_rut(r)
    if not cuerpo or not dv:
        return "", False
    if not cuerpo.isdigit():
        return "", False
    dv_calc = compute_dv(cuerpo)
    if dv_calc is None:
        return "", False
    ok = dv_calc == dv
    return f"{int(cuerpo):,}".replace(",", ".") + "-" + dv, ok  # formatear con puntos

# -------------------- Fechas y horas --------------------
def parse_date_series(s: pd.Series) -> pd.Series:
    # Intentos con dayfirst=True para formatos chilenos
    out = pd.to_datetime(s, errors="coerce", dayfirst=True)
    # Si aún hay NaT y los valores parecen ISO pero con desorden, no forzamos más
    return out.dt.date

def parse_time_series(s: pd.Series) -> pd.Series:
    # soporta "08:15", "8:15", "08:15:00", etc.
    t = pd.to_datetime(s, errors="coerce").dt.time
    return t

# -------------------- Lectura por archivo --------------------
def read_asistencia(file, display_name: str) -> pd.DataFrame:
    try:
        xls = pd.ExcelFile(file)
        sheet = next((s for s in xls.sheet_names if s.lower().strip() == "asistencia"), None)
        if sheet is None:
            return pd.DataFrame({"source_file":[display_name], "error":["Hoja 'asistencia' no encontrada"]})
        df = pd.read_excel(file, sheet_name=sheet, dtype=str)
        df = df.dropna(how="all").dropna(axis=1, how="all")

        # Promover primera fila si hay muchas columnas Unnamed
        if any(str(c).startswith("Unnamed") for c in df.columns) and len(df) > 1:
            candidate = df.iloc[0].astype(str).str.len().sum()
            cols_len = pd.Series(df.columns).astype(str).str.len().sum()
            if candidate > cols_len:
                df.columns = df.iloc[0].astype(str).str.strip()
                df = df.iloc[1:].reset_index(drop=True)

        # Estandarizar columnas
        df = map_columns(df)
        # Agregar archivo de origen
        df["source_file"] = display_name
        return df.reset_index(drop=True)
    except Exception as e:
        return pd.DataFrame({"source_file":[display_name], "error":[str(e)]})

# -------------------- App logic --------------------
if uploaded_files:
    dfs = []
    with st.spinner("Procesando archivos..."):
        for uf in uploaded_files:
            dfs.append(read_asistencia(uf, uf.name))

    consolidado = pd.concat(dfs, ignore_index=True, sort=False) if dfs else pd.DataFrame()

    # Sidebar: limpieza y validaciones
    st.sidebar.header("⚙️ Opciones de limpieza")
    drop_empty_rows = st.sidebar.checkbox("Eliminar filas totalmente vacías", value=True)
    drop_dupes = st.sidebar.checkbox("Eliminar duplicados exactos", value=False)
    dedupe_on_keys = st.sidebar.checkbox("Eliminar duplicados por claves (rut + fecha)", value=True)

    validate_rut = st.sidebar.checkbox("Validar y normalizar RUT", value=True)
    parse_dates = st.sidebar.checkbox("Parsear y normalizar fechas", value=True)
    parse_times = st.sidebar.checkbox("Parsear horas de ‘entrada’/‘salida’", value=True)

    if drop_empty_rows and not consolidado.empty:
        consolidado = consolidado.dropna(how="all")

    # Normalización de RUT
    invalid_rut_count = 0
    if validate_rut and not consolidado.empty and "rut" in consolidado.columns:
        norm_list, ok_list = [], []
        for r in consolidado["rut"].fillna(""):
            nr, ok = normalize_and_validate_rut(r)
            norm_list.append(nr)
            ok_list.append(ok)
        consolidado["rut_norm"] = norm_list
        consolidado["rut_valido"] = ok_list
        invalid_rut_count = (~consolidado["rut_valido"]).sum()

        # Reemplazar "rut" por la versión normalizada si existe
        consolidado["rut"] = consolidado["rut_norm"].where(consolidado["rut_norm"] != "", consolidado["rut"])
        consolidado.drop(columns=["rut_norm"], inplace=True)

    # Parseo de fechas
    invalid_date_count = 0
    if parse_dates and not consolidado.empty and "fecha" in consolidado.columns:
        parsed = parse_date_series(consolidado["fecha"])
        invalid_date_count = parsed.isna().sum()
        consolidado["fecha_norm"] = parsed.astype("string")
        # Reemplazar "fecha" por la versión normalizada si existe
        consolidado["fecha"] = consolidado["fecha_norm"].where(~consolidado["fecha_norm"].isna(), consolidado["fecha"])
        consolidado.drop(columns=["fecha_norm"], inplace=True)

    # Parseo de horas
    if parse_times and not consolidado.empty:
        if "entrada" in consolidado.columns:
            consolidado["entrada"] = parse_time_series(consolidado["entrada"]).astype("string")
        if "salida" in consolidado.columns:
            consolidado["salida"] = parse_time_series(consolidado["salida"]).astype("string")

    # Deduplicaciones
    if drop_dupes and not consolidado.empty:
        consolidado = consolidado.drop_duplicates()
    if dedupe_on_keys and not consolidado.empty and all(c in consolidado.columns for c in ["rut", "fecha"]):
        consolidado = consolidado.drop_duplicates(subset=["rut", "fecha"])

    # Orden y visualización
    existing_order = [c for c in TARGET_ORDER if c in consolidado.columns]
    rest = [c for c in consolidado.columns if c not in existing_order]
    consolidado = consolidado[existing_order + rest]

    st.subheader("Resultado consolidado")
    st.dataframe(consolidado, use_container_width=True)

    # Métricas rápidas
    cols = st.columns(3)
    cols[0].metric("Filas", len(consolidado))
    if validate_rut and "rut_valido" in consolidado.columns:
        cols[1].metric("RUT inválidos", int(invalid_rut_count))
    if parse_dates and "fecha" in consolidado.columns:
        cols[2].metric("Fechas no parseadas", int(invalid_date_count))

    # Exportar a Excel
    if not consolidado.empty:
        out = io.BytesIO()
        with pd.ExcelWriter(out, engine="xlsxwriter") as writer:
            consolidado.to_excel(writer, index=False, sheet_name="consolidado")
        st.download_button(
            label="⬇️ Descargar Excel consolidado",
            data=out.getvalue(),
            file_name="consolidado_asistencia.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
else:
    st.info("Sube uno o más archivos para comenzar.")

st.caption("Tip: Si alguna columna no se mapea, revisa sus nombres originales y ajusta el diccionario de sinónimos en el código.")