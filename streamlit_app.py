
import io
import re
import os
import zipfile
import unicodedata
from datetime import datetime

import pandas as pd
import streamlit as st

# Optional/fuzzy
try:
    from rapidfuzz import process, fuzz
    HAVE_RAPIDFUZZ = True
except Exception:
    HAVE_RAPIDFUZZ = False

# Optional PDF parsing
try:
    import pdfplumber
    HAVE_PDFPLUMBER = True
except Exception:
    HAVE_PDFPLUMBER = False


st.set_page_config(page_title="Compras Textiles - Sommet", layout="wide")
st.title("🧵 Compras Textiles | Generador de Requerimientos y POs")

with st.expander("ℹ️ Cómo usar", expanded=True):
    st.markdown("""
    1. **Sube tu APU/BOM** (Excel): hoja con columnas como `CODIGO_PRENDA, PRENDA, Descripción, Unidad, Cantidad Total, P.U, Proveedor`.
    2. **Sube tu Orden de Compra (OC)** en **PDF** o **Excel**.
    3. (Opcional) **Sube un Diccionario de Sinónimos** (CSV) con columnas: `DESCRIPCION_OC, CODIGO_PRENDA, PRENDA`.
    4. Ajusta el **umbral de coincidencia** y normalización.
    5. Revisa el **Mapeo Propuesto**; se aplicará el diccionario primero y luego el fuzzy.
    6. Descarga los **Excel**: Orden_Cliente, Requerimientos, Consolidado y POs por Proveedor.
    """)

# -----------------------
# Helpers
# -----------------------
def normalize_text(s: str) -> str:
    s = str(s).upper().strip()
    s = " ".join(s.split())
    s = "".join(ch for ch in unicodedata.normalize("NFD", s) if unicodedata.category(ch) != "Mn")
    return s

def load_bom_from_excel(file) -> pd.DataFrame:
    df = pd.read_excel(file, sheet_name=0)
    df.columns = [c.strip() for c in df.columns]
    rename_map = {
        'Descripción': 'MATERIAL',
        'Unidad': 'UNIDAD',
        'Cantidad Total': 'CONSUMO_POR_PRENDA',
        'P.U': 'COSTO_UNITARIO',
        'Costo/ITEM': 'COSTO_ITEM',
        'Proveedor': 'PROVEEDOR'
    }
    for k, v in rename_map.items():
        if k in df.columns:
            df[v] = df[k]
    required = ['CODIGO_PRENDA','PRENDA','MATERIAL','UNIDAD','CONSUMO_POR_PRENDA','COSTO_UNITARIO','PROVEEDOR']
    missing = [c for c in required if c not in df.columns]
    if missing:
        st.warning(f"Faltan columnas esperadas en el APU/BOM: {missing}")
    keep = [c for c in required if c in df.columns]
    return df[keep].copy()

def extract_oc_from_pdf(file) -> pd.DataFrame:
    if not HAVE_PDFPLUMBER:
        st.error("pdfplumber no está instalado. Instala dependencias para leer PDF o sube la OC en Excel.")
        return pd.DataFrame()
    rows = []
    with pdfplumber.open(file) as pdf:
        for page in pdf.pages:
            text = page.extract_text() or ""
            for line in (text.splitlines() if text else []):
                if re.match(r"^\s*\d{1,3}\s", line) and re.search(r"\b\d{2}\.\d{2}\.\d{4}\b", line):
                    m_date = re.search(r"\b\d{2}\.\d{2}\.\d{4}\b", line)
                    item_part = line[:m_date.start()].strip()
                    rest = line[m_date.end():].strip()
                    m_item = re.match(r"^\s*(\d+)\s+(.*)$", item_part)
                    if not m_item:
                        continue
                    item_no = int(m_item.group(1))
                    desc = m_item.group(2).strip()
                    date = line[m_date.start():m_date.end()]
                    m_rest = re.search(
                        r"(\d+(?:[.,]\d+)?)\s+([A-Z]+)\s+(\d+(?:[.,]\d+)?)\s+\S+\s+(\d+(?:[.,]\d+)?)\s*%\s+(\d+(?:[.,]\d+)?)",
                        rest
                    )
                    qty = um = unit_price = iva_pct = subtotal = None
                    if m_rest:
                        qty = float(m_rest.group(1).replace(",", ""))
                        um = m_rest.group(2)
                        unit_price = float(m_rest.group(3).replace(",", ""))
                        iva_pct = float(m_rest.group(4).replace(",", ""))
                        subtotal = float(m_rest.group(5).replace(",", ""))
                    rows.append({
                        "ITEM": item_no,
                        "DESCRIPCION_OC": desc,
                        "FECHA_ENTREGA": date,
                        "CANTIDAD_OC": qty,
                        "UM_OC": um,
                        "P_UNIT_OC": unit_price,
                        "IVA_%": iva_pct,
                        "SUBTOTAL_OC": subtotal
                    })
    return pd.DataFrame(rows).sort_values("ITEM").reset_index(drop=True)

def extract_oc_from_excel(file) -> pd.DataFrame:
    df = pd.read_excel(file, sheet_name=0)
    # Esperadas mínimas: ITEM, DESCRIPCION_OC, CANTIDAD_OC
    cols = [c.strip().upper() for c in df.columns]
    df.columns = cols
    alias = {
        'DESCRIPCIÓN': 'DESCRIPCION_OC',
        'DESCRIPCION': 'DESCRIPCION_OC',
        'CANTIDAD': 'CANTIDAD_OC',
        'UM': 'UM_OC',
        'FECHA ENTREGA': 'FECHA_ENTREGA',
        'FECHA_ENTREGA': 'FECHA_ENTREGA',
        'P.U': 'P_UNIT_OC',
        'PRECIO UNITARIO': 'P_UNIT_OC',
        'SUBTOTAL': 'SUBTOTAL_OC'
    }
    for k, v in alias.items():
        if k in df.columns:
            df[v] = df[k]
    must = ['DESCRIPCION_OC','CANTIDAD_OC']
    miss = [c for c in must if c not in df.columns]
    if miss:
        st.warning(f"Faltan columnas mínimas en OC Excel: {miss}")
    keep = ['ITEM','DESCRIPCION_OC','FECHA_ENTREGA','CANTIDAD_OC','UM_OC','P_UNIT_OC','IVA_%','SUBTOTAL_OC']
    keep = [c for c in keep if c in df.columns]
    if 'ITEM' not in keep:
        df['ITEM'] = range(1, len(df)+1)
        keep = ['ITEM'] + [c for c in keep if c != 'ITEM']
    return df[keep].copy()

def apply_dictionary(oc_df: pd.DataFrame, dict_df: pd.DataFrame) -> pd.DataFrame:
    if dict_df is None or dict_df.empty:
        oc_df['CODIGO_PRENDA'] = None
        oc_df['PRENDA'] = None
        return oc_df
    d = dict_df.copy()
    d.columns = [c.strip().upper() for c in d.columns]
    req_cols = ['DESCRIPCION_OC','CODIGO_PRENDA','PRENDA']
    for c in req_cols:
        if c not in d.columns:
            st.warning("El diccionario no tiene columnas requeridas: DESCRIPCION_OC,CODIGO_PRENDA,PRENDA")
            oc_df['CODIGO_PRENDA'] = None
            oc_df['PRENDA'] = None
            return oc_df
    # Normalizamos para join robusto
    d['DESCRIPCION_N'] = d['DESCRIPCION_OC'].apply(normalize_text)
    tmp = oc_df.copy()
    tmp['DESCRIPCION_N'] = tmp['DESCRIPCION_OC'].apply(normalize_text)
    merged = tmp.merge(d[['DESCRIPCION_N','CODIGO_PRENDA','PRENDA']],
                       on='DESCRIPCION_N', how='left')
    merged.drop(columns=['DESCRIPCION_N'], inplace=True)
    return merged

def fuzzy_map(oc_df: pd.DataFrame, catalogo: pd.DataFrame, threshold: float) -> pd.DataFrame:
    if not HAVE_RAPIDFUZZ:
        st.info("rapidfuzz no está instalado, se omitirá el mapeo difuso (fuzzy).")
        oc_df['SUG_CODIGO_PRENDA'] = None
        oc_df['SUG_PRENDA'] = None
        oc_df['COINCIDENCIA'] = None
        return oc_df
    cat = catalogo.copy()
    cat['PRENDA_N'] = cat['PRENDA'].apply(normalize_text)
    choices = cat['PRENDA_N'].tolist()

    sug_code, sug_name, score_list = [], [], []
    for _, row in oc_df.iterrows():
        desc_n = normalize_text(row['DESCRIPCION_OC'])
        if not choices:
            sug_code.append(None); sug_name.append(None); score_list.append(None)
            continue
        match = process.extractOne(desc_n, choices, scorer=fuzz.token_sort_ratio)
        if match:
            best, score, idx = match
            if score/100.0 >= threshold:
                sug_code.append(cat.iloc[idx]['CODIGO_PRENDA'])
                sug_name.append(cat.iloc[idx]['PRENDA'])
                score_list.append(round(score/100.0, 3))
            else:
                sug_code.append(None); sug_name.append(None); score_list.append(round(score/100.0, 3))
        else:
            sug_code.append(None); sug_name.append(None); score_list.append(None)
    oc_df['SUG_CODIGO_PRENDA'] = sug_code
    oc_df['SUG_PRENDA'] = sug_name
    oc_df['COINCIDENCIA'] = score_list
    return oc_df

def expand_requirements(orden_cliente: pd.DataFrame, bom: pd.DataFrame) -> pd.DataFrame:
    ok = orden_cliente.dropna(subset=['CODIGO_PRENDA']).copy()
    if ok.empty:
        return pd.DataFrame()
    req = ok.merge(bom, on=['CODIGO_PRENDA','PRENDA'], how='left')
    req['CANTIDAD'] = pd.to_numeric(req['CANTIDAD'], errors='coerce').fillna(0.0)
    req['CONSUMO_POR_PRENDA'] = pd.to_numeric(req['CONSUMO_POR_PRENDA'], errors='coerce').fillna(0.0)
    req['CANTIDAD_MATERIAL'] = req['CANTIDAD'] * req['CONSUMO_POR_PRENDA']
    return req

def consolidate(req: pd.DataFrame, normalize_materials=True, normalize_suppliers=True) -> pd.DataFrame:
    if req.empty:
        return pd.DataFrame()
    tmp = req.copy()
    tmp['MATERIAL_ORIG'] = tmp['MATERIAL']
    tmp['PROVEEDOR_ORIG'] = tmp['PROVEEDOR']
    if normalize_materials:
        tmp['MATERIAL'] = tmp['MATERIAL'].apply(normalize_text)
    if normalize_suppliers:
        tmp['PROVEEDOR'] = tmp['PROVEEDOR'].apply(normalize_text)
    grp = (tmp.groupby(['PROVEEDOR','MATERIAL','UNIDAD'], as_index=False)
              .agg({'CANTIDAD_MATERIAL':'sum','COSTO_UNITARIO':'max'}))
    grp['COSTO_ESTIMADO'] = grp['CANTIDAD_MATERIAL'] * grp['COSTO_UNITARIO']
    return grp

def to_excel_bytes(sheets: dict) -> bytes:
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        for name, df in sheets.items():
            df.to_excel(writer, sheet_name=name[:31] if name else "Sheet1", index=False)
    return output.getvalue()

def separate_pos_zip(consol: pd.DataFrame) -> bytes:
    # Crear un archivo ZIP con una hoja Excel por proveedor
    buffer = io.BytesIO()
    with zipfile.ZipFile(buffer, "w", zipfile.ZIP_DEFLATED) as zf:
        for prov, sub in consol.groupby('PROVEEDOR'):
            file_bytes = to_excel_bytes({"PO_"+str(prov)[:25]: sub})
            zf.writestr(f"PO_{prov}.xlsx", file_bytes)
    return buffer.getvalue()


# -----------------------
# Sidebar
# -----------------------
st.sidebar.header("Parámetros")
th = st.sidebar.slider("Umbral de coincidencia (fuzzy)", 0.50, 0.95, 0.65, 0.01)
normalize_materials = st.sidebar.checkbox("Normalizar materiales", value=True)
normalize_suppliers = st.sidebar.checkbox("Normalizar proveedores", value=True)

# -----------------------
# Inputs
# -----------------------
apu_file = st.file_uploader("APU/BOM (Excel)", type=["xlsx","xls"], key="apu")
oc_file = st.file_uploader("Orden de Compra (PDF o Excel)", type=["pdf","xlsx","xls"], key="oc")
dict_file = st.file_uploader("Diccionario de Sinónimos (CSV opcional)", type=["csv"], key="dict")

if apu_file is not None:
    bom = load_bom_from_excel(apu_file)
    st.success(f"APU/BOM cargado. Prendas únicas: {bom[['CODIGO_PRENDA','PRENDA']].drop_duplicates().shape[0]}")
    st.dataframe(bom.head(20), use_container_width=True)
else:
    st.stop()

# Catalogo prendas
catalogo = bom[['CODIGO_PRENDA','PRENDA']].drop_duplicates().reset_index(drop=True)

# Leer OC
oc_df = pd.DataFrame()
if oc_file is not None:
    ext = os.path.splitext(oc_file.name)[1].lower()
    if ext == ".pdf":
        oc_df = extract_oc_from_pdf(oc_file)
    else:
        oc_df = extract_oc_from_excel(oc_file)
    if oc_df.empty:
        st.warning("No se pudieron extraer ítems de la OC.")
    else:
        st.subheader("OC extraída")
        st.dataframe(oc_df, use_container_width=True)
else:
    st.info("Sube una Orden de Compra para continuar.")
    st.stop()

# Diccionario
dict_df = None
if dict_file is not None:
    dict_df = pd.read_csv(dict_file)
    st.success(f"Diccionario cargado: {dict_df.shape[0]} equivalencias.")
    st.dataframe(dict_df.head(20), use_container_width=True)

# Aplicar diccionario y fuzzy
oc_dict = apply_dictionary(oc_df, dict_df)

if HAVE_RAPIDFUZZ:
    oc_sug = fuzzy_map(oc_dict, catalogo, th)
else:
    oc_sug = oc_dict.copy()
    oc_sug['SUG_CODIGO_PRENDA'] = None
    oc_sug['SUG_PRENDA'] = None
    oc_sug['COINCIDENCIA'] = None

st.subheader("Mapeo Propuesto (aplica diccionario y luego fuzzy)")
st.dataframe(oc_sug, use_container_width=True)

# Construir Orden_Cliente final (aceptando sugeridos si faltan)
orden_cliente = oc_sug.copy()
mask_missing = orden_cliente['CODIGO_PRENDA'].isna() & orden_cliente['SUG_CODIGO_PRENDA'].notna()
orden_cliente.loc[mask_missing, 'CODIGO_PRENDA'] = orden_cliente.loc[mask_missing, 'SUG_CODIGO_PRENDA']
orden_cliente.loc[mask_missing, 'PRENDA'] = orden_cliente.loc[mask_missing, 'SUG_PRENDA']
orden_cliente = orden_cliente.rename(columns={'CANTIDAD_OC':'CANTIDAD'})

st.subheader("Orden_Cliente (lista para expandir)")
st.dataframe(orden_cliente[['ITEM','DESCRIPCION_OC','CODIGO_PRENDA','PRENDA','CANTIDAD','UM_OC','FECHA_ENTREGA']],
             use_container_width=True)

# Expandir a requerimientos
req = expand_requirements(orden_cliente, bom)
if req.empty:
    st.warning("No hay prendas mapeadas para expandir requerimientos.")
else:
    st.subheader("Requerimientos por Prenda")
    st.dataframe(req, use_container_width=True)

    # Consolidado
    consol = consolidate(req, normalize_materials=normalize_materials, normalize_suppliers=normalize_suppliers)
    st.subheader("Consolidado de Material por Proveedor")
    st.dataframe(consol, use_container_width=True)

    # Descargas
    st.subheader("Descargas")
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    # Excel principal
    excel_bytes = to_excel_bytes({
        "OC_Extraida": oc_df,
        "Mapeo_Propuesto": oc_sug,
        "Orden_Cliente": orden_cliente,
        "Req_por_Prenda": req,
        "Consolidado_Material": consol
    })
    st.download_button("⬇️ Descargar Excel (OC + Requerimientos + Consolidado)",
                       data=excel_bytes,
                       file_name=f"ComprasTextiles_{ts}.xlsx",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    # ZIP con POs por proveedor
    if not consol.empty:
        zip_bytes = separate_pos_zip(consol)
        st.download_button("⬇️ Descargar ZIP con POs por Proveedor",
                           data=zip_bytes,
                           file_name=f"POs_por_Proveedor_{ts}.zip",
                           mime="application/zip")

st.caption("© Sommet Supplies – Generador de Compras Textiles")
