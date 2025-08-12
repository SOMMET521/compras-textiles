# (same app code as earlier essential version, shortened comment here for brevity)
# Full code preserved from previous message
import io, re, os, zipfile, unicodedata
from datetime import datetime
import pandas as pd
import streamlit as st

try:
    from rapidfuzz import process, fuzz
    HAVE_RAPIDFUZZ = True
except Exception:
    HAVE_RAPIDFUZZ = False

try:
    import pdfplumber
    HAVE_PDFPLUMBER = True
except Exception:
    HAVE_PDFPLUMBER = False

st.set_page_config(page_title="Compras Textiles - Sommet", layout="wide")
st.title("🧵 Compras Textiles | Generador de Requerimientos y POs")

with st.expander("ℹ️ Cómo usar", expanded=False):
    st.markdown("""1) Sube APU; 2) Sube OC (PDF/Excel); 3) (Opc.) Diccionario; 4) Ajusta parámetros; 5) Descarga Excel y ZIP de POs.""")

def normalize_text(s: str) -> str:
    s = str(s).upper().strip()
    s = " ".join(s.split())
    s = "".join(ch for ch in unicodedata.normalize("NFD", s) if unicodedata.category(ch) != "Mn")
    return s

def load_bom_from_excel(file):
    df = pd.read_excel(file, sheet_name=0, engine="openpyxl")
    df.columns = [c.strip() for c in df.columns]
    rename_map = {'Descripción':'MATERIAL','Unidad':'UNIDAD','Cantidad Total':'CONSUMO_POR_PRENDA','P.U':'COSTO_UNITARIO','Costo/ITEM':'COSTO_ITEM','Proveedor':'PROVEEDOR'}
    for k,v in rename_map.items():
        if k in df.columns: df[v]=df[k]
    need = ['CODIGO_PRENDA','PRENDA','MATERIAL','UNIDAD','CONSUMO_POR_PRENDA','COSTO_UNITARIO','PROVEEDOR']
    keep = [c for c in need if c in df.columns]
    return df[keep].copy()

def extract_oc_from_pdf(file):
    if not HAVE_PDFPLUMBER: return pd.DataFrame()
    rows=[]; import re, pdfplumber
    with pdfplumber.open(file) as pdf:
        for p in pdf.pages:
            t=p.extract_text() or ""
            for line in t.splitlines():
                if re.match(r"^\s*\d{1,3}\s", line) and re.search(r"\b\d{2}\.\d{2}\.\d{4}\b", line):
                    m=re.search(r"\b\d{2}\.\d{2}\.\d{4}\b", line); item_part=line[:m.start()].strip(); rest=line[m.end():].strip()
                    mi=re.match(r"^\s*(\d+)\s+(.*)$", item_part); 
                    if not mi: continue
                    item_no=int(mi.group(1)); desc=mi.group(2).strip(); date=line[m.start():m.end()]
                    mr=re.search(r"(\d+(?:[.,]\d+)?)\s+([A-Z]+)\s+(\d+(?:[.,]\d+)?)\s+\S+\s+(\d+(?:[.,]\d+)?)\s*%\s+(\d+(?:[.,]\d+)?)", rest)
                    qty=um=unit=iva=sub=None
                    if mr:
                        qty=float(mr.group(1).replace(",","")); um=mr.group(2); unit=float(mr.group(3).replace(",","")); iva=float(mr.group(4).replace(",","")); sub=float(mr.group(5).replace(",",""))
                    rows.append({"ITEM":item_no,"DESCRIPCION_OC":desc,"FECHA_ENTREGA":date,"CANTIDAD_OC":qty,"UM_OC":um,"P_UNIT_OC":unit,"IVA_%":iva,"SUBTOTAL_OC":sub})
    return pd.DataFrame(rows).sort_values("ITEM").reset_index(drop=True)

def extract_oc_from_excel(file):
    df = pd.read_excel(file, sheet_name=0, engine="openpyxl")
    cols=[c.strip().upper() for c in df.columns]; df.columns=cols
    alias={'DESCRIPCIÓN':'DESCRIPCION_OC','DESCRIPCION':'DESCRIPCION_OC','CANTIDAD':'CANTIDAD_OC','UM':'UM_OC','FECHA ENTREGA':'FECHA_ENTREGA','FECHA_ENTREGA':'FECHA_ENTREGA','P.U':'P_UNIT_OC','PRECIO UNITARIO':'P_UNIT_OC','SUBTOTAL':'SUBTOTAL_OC'}
    for k,v in alias.items():
        if k in df.columns: df[v]=df[k]
    keep=['ITEM','DESCRIPCION_OC','FECHA_ENTREGA','CANTIDAD_OC','UM_OC','P_UNIT_OC','IVA_%','SUBTOTAL_OC']; keep=[c for c in keep if c in df.columns]
    if 'ITEM' not in keep:
        df['ITEM']=range(1,len(df)+1); keep=['ITEM']+[c for c in keep if c!='ITEM']
    return df[keep].copy()

def apply_dictionary(oc_df, dict_df):
    if dict_df is None or dict_df.empty:
        oc_df['CODIGO_PRENDA']=None; oc_df['PRENDA']=None; return oc_df
    d=dict_df.copy(); d.columns=[c.strip().upper() for c in d.columns]
    if not set(['DESCRIPCION_OC','CODIGO_PRENDA','PRENDA']).issubset(d.columns):
        oc_df['CODIGO_PRENDA']=None; oc_df['PRENDA']=None; return oc_df
    d['DESCRIPCION_N']=d['DESCRIPCION_OC'].apply(normalize_text)
    tmp=oc_df.copy(); tmp['DESCRIPCION_N']=tmp['DESCRIPCION_OC'].apply(normalize_text)
    m=tmp.merge(d[['DESCRIPCION_N','CODIGO_PRENDA','PRENDA']],on='DESCRIPCION_N',how='left').drop(columns=['DESCRIPCION_N'])
    return m

def fuzzy_map(oc_df, catalogo, threshold):
    if not HAVE_RAPIDFUZZ:
        oc_df['SUG_CODIGO_PRENDA']=None; oc_df['SUG_PRENDA']=None; oc_df['COINCIDENCIA']=None; return oc_df
    from rapidfuzz import process, fuzz
    cat=catalogo.copy(); cat['PRENDA_N']=cat['PRENDA'].apply(normalize_text); choices=cat['PRENDA_N'].tolist()
    sug_code=[]; sug_name=[]; score_list=[]
    for _,row in oc_df.iterrows():
        desc_n=normalize_text(row['DESCRIPCION_OC']); 
        if not choices: sug_code.append(None); sug_name.append(None); score_list.append(None); continue
        match=process.extractOne(desc_n, choices, scorer=fuzz.token_sort_ratio)
        if match:
            best,score,idx=match
            if score/100.0>=threshold:
                sug_code.append(cat.iloc[idx]['CODIGO_PRENDA']); sug_name.append(cat.iloc[idx]['PRENDA']); score_list.append(round(score/100.0,3))
            else:
                sug_code.append(None); sug_name.append(None); score_list.append(round(score/100.0,3))
        else:
            sug_code.append(None); sug_name.append(None); score_list.append(None)
    oc_df['SUG_CODIGO_PRENDA']=sug_code; oc_df['SUG_PRENDA']=sug_name; oc_df['COINCIDENCIA']=score_list; return oc_df

def expand_requirements(orden_cliente, bom):
    ok=orden_cliente.dropna(subset=['CODIGO_PRENDA']).copy()
    if ok.empty: return pd.DataFrame()
    req=ok.merge(bom, on=['CODIGO_PRENDA','PRENDA'], how='left')
    req['CANTIDAD']=pd.to_numeric(req['CANTIDAD'], errors='coerce').fillna(0.0)
    req['CONSUMO_POR_PRENDA']=pd.to_numeric(req['CONSUMO_POR_PRENDA'], errors='coerce').fillna(0.0)
    req['CANTIDAD_MATERIAL']=req['CANTIDAD']*req['CONSUMO_POR_PRENDA']
    return req

def consolidate(req, normalize_materials=True, normalize_suppliers=True):
    if req.empty: return pd.DataFrame()
    tmp=req.copy(); tmp['MATERIAL_ORIG']=tmp['MATERIAL']; tmp['PROVEEDOR_ORIG']=tmp['PROVEEDOR']
    if normalize_materials: tmp['MATERIAL']=tmp['MATERIAL'].apply(normalize_text)
    if normalize_suppliers: tmp['PROVEEDOR']=tmp['PROVEEDOR'].apply(normalize_text)
    grp=(tmp.groupby(['PROVEEDOR','MATERIAL','UNIDAD'], as_index=False)
           .agg({'CANTIDAD_MATERIAL':'sum','COSTO_UNITARIO':'max'}))
    grp['COSTO_ESTIMADO']=grp['CANTIDAD_MATERIAL']*grp['COSTO_UNITARIO']
    return grp

def write_formatted_excel(sheets: dict)->bytes:
    output=io.BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        wb=writer.book
        h=wb.add_format({"bold":True,"bg_color":"#444444","font_color":"white","border":1,"align":"center","valign":"vcenter"})
        money=wb.add_format({"num_format":"$#,##0.00","border":1})
        num=wb.add_format({"num_format":"#,##0.00","border":1})
        txt=wb.add_format({"border":1})
        for name,df in sheets.items():
            sheet=name[:31] if name else "Sheet1"
            df.to_excel(writer, sheet_name=sheet, index=False)
            ws=writer.sheets[sheet]
            for c,_ in enumerate(df.columns): ws.write(0,c,df.columns[c],h)
            for i,col in enumerate(df.columns):
                series=df[col]; width=min(max(12, series.astype(str).map(len).max()+2), 40)
                if pd.api.types.is_numeric_dtype(series):
                    if any(k in col.upper() for k in ["COSTO","PRECIO","P.U","SUBTOTAL","TOTAL"]):
                        ws.set_column(i,i,width,money)
                    else:
                        ws.set_column(i,i,width,num)
                else:
                    ws.set_column(i,i,width,txt)
            ws.freeze_panes(1,0)
    return output.getvalue()

def create_po_excel(prov_name, po_code, consol_sub, params, logo_bytes):
    output=io.BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        wb=writer.book; ws=wb.add_worksheet("PO")
        title=wb.add_format({"bold":True,"font_size":16}); lab=wb.add_format({"bold":True})
        head=wb.add_format({"bold":True,"bg_color":"#444444","font_color":"white","border":1,"align":"center","valign":"vcenter"})
        text=wb.add_format({"border":1}); num=wb.add_format({"num_format":"#,##0.00","border":1}); money=wb.add_format({"num_format":"$#,##0.00","border":1})
        total_l=wb.add_format({"bold":True,"border":1}); total_v=wb.add_format({"bold":True,"num_format":"$#,##0.00","border":1})
        row=0
        if logo_bytes:
            try: ws.insert_image(row,0,"logo.png",{"image_data":io.BytesIO(logo_bytes),"x_scale":0.5,"y_scale":0.5})
            except: pass
        ws.write(row,4,"ORDEN DE COMPRA",title); row+=2
        ws.write(row,0,"Proveedor:",lab); ws.write(row,1,prov_name)
        ws.write(row,3,"PO:",lab); ws.write(row,4,po_code); row+=1
        ws.write(row,0,"Empresa:",lab); ws.write(row,1,params.get("empresa","Sommet Supplies"))
        ws.write(row,3,"Fecha:",lab); ws.write(row,4,datetime.now().strftime("%Y-%m-%d")); row+=1
        ws.write(row,0,"RUC:",lab); ws.write(row,1,params.get("ruc",""))
        ws.write(row,3,"OC Cliente:",lab); ws.write(row,4,params.get("oc_cliente","")); row+=1
        ws.write(row,0,"Dirección:",lab); ws.write(row,1,params.get("direccion",""))
        ws.write(row,3,"Contacto:",lab); ws.write(row,4,params.get("contacto","")); row+=2
        headers=["Material","UM","Cantidad","P.U (USD)","Subtotal (USD)","IVA %","Total (USD)"]
        for c,hv in enumerate(headers): ws.write(row,c,hv,head)
        start=row+1; tot_sub=0.0; iva=float(params.get("iva_default",15.0))
        for _,r in consol_sub.iterrows():
            mat=r.get("MATERIAL",""); um=r.get("UNIDAD",""); qty=float(r.get("CANTIDAD_MATERIAL",0)); pu=float(r.get("COSTO_UNITARIO",0))
            sub=qty*pu; tot=sub*(1+iva/100.0)
            ws.write(start,0,str(mat),text); ws.write(start,1,str(um),text)
            ws.write_number(start,2,qty,num); ws.write_number(start,3,pu,money); ws.write_number(start,4,sub,money)
            ws.write_number(start,5,iva,num); ws.write_number(start,6,tot,money)
            tot_sub+=sub; start+=1
        ws.write(start+1,5,"Subtotal:",total_l); ws.write_number(start+1,6,tot_sub,total_v)
        ws.write(start+2,5,f"IVA {iva:.0f}%:",total_l); ws.write_number(start+2,6,tot_sub*iva/100.0,total_v)
        ws.write(start+3,5,"Total:",total_l); ws.write_number(start+3,6,tot_sub*(1+iva/100.0),total_v)
        row_c=start+5; ws.write(row_c,0,"Condiciones:",lab); ws.merge_range(row_c,1,row_c,6,params.get("condiciones","Pago contra entrega."),text)
        ws.set_column(0,0,44); ws.set_column(1,1,8); ws.set_column(2,2,12); ws.set_column(3,3,14); ws.set_column(4,4,16); ws.set_column(5,5,8); ws.set_column(6,6,16)
    return output.getvalue()

def make_po_code(prefix, year2, oc_cliente, seq): return f"{prefix}-{year2}-OC{oc_cliente}-P{seq:03d}"

def pos_zip_with_format(consol, params, logo_bytes):
    buf=io.BytesIO()
    with zipfile.ZipFile(buf,"w",zipfile.ZIP_DEFLATED) as z:
        seq=1
        for prov, sub in consol.groupby('PROVEEDOR'):
            code=make_po_code(params['prefix'], params['year2'], params['oc_cliente'], seq)
            po_bytes=create_po_excel(prov, code, sub, params, logo_bytes)
            fname=f"PO_{code}_{str(prov).upper().replace(' ','_')}.xlsx"
            z.writestr(fname, po_bytes); seq+=1
    return buf.getvalue()

# Sidebar
st.sidebar.header("Parámetros")
th=st.sidebar.slider("Umbral fuzzy",0.50,0.95,0.65,0.01)
st.sidebar.header("PO & Empresa")
prefix=st.sidebar.text_input("Prefijo PO",value="SOM")
year2=st.sidebar.text_input("Año (2 dígitos)",value=datetime.now().strftime("%y"))
oc_cliente=st.sidebar.text_input("OC Cliente (solo números)",value="4500412496")
iva_default=st.sidebar.number_input("IVA (%)",0.0,100.0,15.0,0.5)
empresa=st.sidebar.text_input("Empresa",value="Sommet Supplies")
ruc=st.sidebar.text_input("RUC",value="1711659613001")
direccion=st.sidebar.text_input("Dirección",value="Av. General Eloy Alfaro E13-224 y calle De Los Nogales")
contacto=st.sidebar.text_input("Contacto (tel/email)",value="0999926669 / felipe@sommet.supplies")
condiciones=st.sidebar.text_area("Condiciones",value="Pago contra entrega.")
logo_file=st.sidebar.file_uploader("Logo (PNG/JPG)",type=["png","jpg","jpeg"])
logo_bytes=logo_file.read() if logo_file else None

params={"prefix":prefix,"year2":year2,"oc_cliente":oc_cliente,"iva_default":iva_default,"empresa":empresa,"ruc":ruc,"direccion":direccion,"contacto":contacto,"condiciones":condiciones}

# Inputs
apu_file=st.file_uploader("APU/BOM (Excel)",type=["xlsx","xls"],key="apu")
oc_file=st.file_uploader("Orden de Compra (PDF o Excel)",type=["pdf","xlsx","xls"],key="oc")
dict_file=st.file_uploader("Diccionario de Sinónimos (CSV opcional)",type=["csv"],key="dict")

if apu_file is None: st.stop()
bom=load_bom_from_excel(apu_file)
st.success(f"APU/BOM cargado. Prendas únicas: {bom[['CODIGO_PRENDA','PRENDA']].drop_duplicates().shape[0]}")
st.dataframe(bom.head(20),use_container_width=True)

catalogo=bom[['CODIGO_PRENDA','PRENDA']].drop_duplicates().reset_index(drop=True)

if oc_file is None: st.info("Sube una OC para continuar."); st.stop()
ext=os.path.splitext(oc_file.name)[1].lower()
oc_df=extract_oc_from_pdf(oc_file) if ext==".pdf" else extract_oc_from_excel(oc_file)
if oc_df.empty: st.warning("No se pudieron extraer ítems de la OC."); st.stop()
st.subheader("OC extraída"); st.dataframe(oc_df,use_container_width=True)

dict_df=None
if dict_file is not None:
    import pandas as pd
    dict_df=pd.read_csv(dict_file); st.success(f"Diccionario cargado: {dict_df.shape[0]} equivalencias.")

oc_dict=apply_dictionary(oc_df, dict_df)
if HAVE_RAPIDFUZZ:
    oc_sug=fuzzy_map(oc_dict, catalogo, th)
else:
    oc_sug=oc_dict.copy(); oc_sug['SUG_CODIGO_PRENDA']=None; oc_sug['SUG_PRENDA']=None; oc_sug['COINCIDENCIA']=None

st.subheader("Mapeo Propuesto"); st.dataframe(oc_sug,use_container_width=True)

orden_cliente=oc_sug.copy()
mask=orden_cliente['CODIGO_PRENDA'].isna() & orden_cliente['SUG_CODIGO_PRENDA'].notna()
orden_cliente.loc[mask,'CODIGO_PRENDA']=orden_cliente.loc[mask,'SUG_CODIGO_PRENDA']
orden_cliente.loc[mask,'PRENDA']=orden_cliente.loc[mask,'SUG_PRENDA']
orden_cliente=orden_cliente.rename(columns={'CANTIDAD_OC':'CANTIDAD'})
st.subheader("Orden_Cliente"); st.dataframe(orden_cliente[['ITEM','DESCRIPCION_OC','CODIGO_PRENDA','PRENDA','CANTIDAD','UM_OC','FECHA_ENTREGA']],use_container_width=True)

req=expand_requirements(orden_cliente, bom)
if req.empty: st.warning("No hay prendas mapeadas para expandir."); st.stop()
st.subheader("Requerimientos por Prenda"); st.dataframe(req,use_container_width=True)

consol=consolidate(req, True, True)
st.subheader("Consolidado por Proveedor"); st.dataframe(consol,use_container_width=True)

excel_bytes=write_formatted_excel({"OC_Extraida":oc_df,"Mapeo_Propuesto":oc_sug,"Orden_Cliente":orden_cliente,"Req_por_Prenda":req,"Consolidado_Material":consol})
st.download_button("⬇️ Excel (OC + Requerimientos + Consolidado)", data=excel_bytes, file_name=f"ComprasTextiles_{params['oc_cliente']}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

zip_bytes=pos_zip_with_format(consol, params, logo_bytes)
st.download_button("⬇️ ZIP con POs por Proveedor", data=zip_bytes, file_name=f"POs_OC{params['oc_cliente']}.zip", mime="application/zip")
