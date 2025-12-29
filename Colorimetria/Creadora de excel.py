import pandas as pd
import camelot
from PyPDF2 import PdfReader
from datetime import datetime
import re

# --- ARCHIVOS ---
pdf_path = r"C:\Users\Usuario\Desktop\Uni\Labo II\Colorimetria\00 (24 arquivos mesclados).pdf"
excel_path = r"C:\Users\Usuario\Desktop\Uni\Labo II\Colorimetria\Grillas_por_tiempo.xlsx"

# --- EXTRAER TODAS LAS TABLAS ---
print("📄 Extrayendo tablas del PDF...")
tablas = camelot.read_pdf(pdf_path, pages='all', flavor='lattice')  # usa detección por líneas

# --- LEER LAS FECHAS DE CADA PÁGINA ---
reader = PdfReader(pdf_path)
fechas = []
for pagina in reader.pages:
    texto = pagina.extract_text()
    m = re.search(r"Date\s*:\s*([\d-]+)\s+([\d:]+)", texto)
    if m:
        fecha_str = m.group(1) + " " + m.group(2)
        try:
            fecha = datetime.strptime(fecha_str, "%Y-%m-%d %H:%M:%S")
        except:
            try:
                fecha = datetime.strptime(fecha_str, "%Y-%m-%d %H:%M")
            except:
                fecha = None
        fechas.append(fecha)
    else:
        fechas.append(None)

# --- CALCULAR DIFERENCIAS DE TIEMPO CON LA SEGUNDA GRILLA ---
ref_time = fechas[1]
diffs = []
for f in fechas:
    if f and ref_time:
        dt = round((f - ref_time).total_seconds(), 2)
        nombre = f"{dt:+.1f} seg"
    else:
        nombre = "sin_fecha"
    diffs.append(nombre)

# --- EXPORTAR A EXCEL ---
print("🧮 Guardando en Excel...")
with pd.ExcelWriter(excel_path, engine="openpyxl") as writer:
    for i, (t, nombre) in enumerate(zip(tablas, diffs)):
        df = t.df
        hoja = nombre.replace(":", "_")
        df.to_excel(writer, index=False, sheet_name=hoja[:31])

print(f"✅ Listo: {excel_path}")
