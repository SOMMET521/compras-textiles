
# Compras Textiles · Sommet Supplies

## Requisitos
```bash
pip install -r requirements.txt
```

## Ejecutar
```bash
streamlit run streamlit_app.py
```

## Archivos de entrada
- **APU/BOM (Excel)**: Debe contener columnas (o equivalentes):  
  `CODIGO_PRENDA, PRENDA, Descripción, Unidad, Cantidad Total, P.U, Proveedor`
- **OC (PDF o Excel)**
- **Diccionario de Sinónimos (CSV, opcional)** con columnas:  
  `DESCRIPCION_OC, CODIGO_PRENDA, PRENDA`

## Salidas
- Excel con: `OC_Extraida`, `Mapeo_Propuesto`, `Orden_Cliente`, `Req_por_Prenda`, `Consolidado_Material`
- ZIP con POs por proveedor (un Excel por proveedor)
