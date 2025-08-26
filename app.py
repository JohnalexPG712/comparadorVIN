import streamlit as st
import pandas as pd
import fitz  # PyMuPDF
import re
import io
from collections import Counter

# --------------------------------------------------------------------------
# Todas las funciones de procesamiento se mantienen sin cambios.
# --------------------------------------------------------------------------

def normalizar_vin(vin):
    return (str(vin).replace(" ", "").replace("\r", "").replace("\n", "").replace("\t", "")).upper() if vin else ""

def validar_vin(vin):
    vin_limpio = normalizar_vin(vin)
    return bool(re.fullmatch(r"[A-HJ-NPR-Z0-9]{17}", vin_limpio))

def es_vin_plausible(vin):
    if len(vin) != 17:
        return False
    if len(re.findall(r"[A-Z]{6,}", vin)) > 0:
        return False
    wmi = vin[:3]
    if len(re.findall(r"[A-Z]", wmi)) < 2:
        return False
    serial_part = vin[6:]
    if len(re.findall(r"[0-9]", serial_part)) < 4:
        return False
    return True

def buscar_vin_flexible(vin, texto_pdf):
    # Escapa el VIN para usarlo en regex y crea un patrón que ignore espacios
    regex_pattern = r"".join(re.escape(char) + r"[\s\r\n]*" for char in vin)
    return bool(re.search(regex_pattern, texto_pdf, re.IGNORECASE))

def leer_excel_vins(excel_file):
    vins_validos = []
    vins_invalidos = []
    
    # Usa un buffer para que el archivo pueda ser leído múltiples veces si es necesario
    excel_file.seek(0)
    extension = excel_file.name.split('.')[-1].lower()
    if extension == 'xls':
        df = pd.read_excel(excel_file, engine='xlrd', header=None, dtype=str, keep_default_na=False)
    else:
        df = pd.read_excel(excel_file, header=None, dtype=str, keep_default_na=False)
    
    start_row = 0
    if len(df.columns) > 1:
        # Intenta detectar si la primera fila es un encabezado
        primer_posible_vin = normalizar_vin(df.iloc[0, 1])
        if not validar_vin(primer_posible_vin):
            start_row = 1
        
        vins_crudos = df.iloc[start_row:, 1].tolist()
        for vin_crudo in vins_crudos:
            vin_normalizado = normalizar_vin(vin_crudo)
            if vin_normalizado:
                if validar_vin(vin_normalizado):
                    vins_validos.append(vin_normalizado)
                else:
                    vins_invalidos.append({"vin": vin_crudo, "length": len(vin_normalizado)})
    return vins_validos, vins_invalidos

def leer_pdf(file):
    texto_completo = ""
    # El archivo subido ya está en bytes, no es necesario file.read()
    file.seek(0)
    doc = fitz.open(stream=file.read(), filetype="pdf")
    for page_num in range(doc.page_count):
        page = doc.load_page(page_num)
        texto_completo += page.get_text() + " "
    doc.close()
    texto_completo = re.sub(r'\s+', ' ', texto_completo).upper()
    return texto_completo

# --------------------------------------------------------------------------
# Interfaz de la aplicación Streamlit
# --------------------------------------------------------------------------

st.set_page_config(page_title="Comparador de VINs", layout="centered")
st.title("🔍 Comparador de VINs: Excel (FMM) vs Documentos PDF")

# NUEVO: Se añade un 'key' a cada file_uploader. Esto permite controlarlos
# y borrarlos mediante programación con el nuevo botón de limpiar.
excel_file = st.file_uploader("1. Sube el archivo Excel (FMM)", type=["xlsx", "xls"], key="excel_uploader")
pdf_files = st.file_uploader("2. Sube los archivos PDF de soporte", type=["pdf"], accept_multiple_files=True, key="pdf_uploader")

# NUEVO: Se usan columnas para colocar los botones uno al lado del otro.
col1, col2 = st.columns([1.5, 2])

with col1:
    procesar = st.button("3. Procesar y Comparar", type="primary")

with col2:
    # NUEVO: Este es el botón de limpiar. Al hacer clic, ejecuta st.rerun() que
    # recarga la página, limpiando así el estado y los archivos.
    # El simple hecho de tener el botón y la recarga es suficiente para limpiar.
    if st.button("🧹 Limpiar y Empezar de Nuevo"):
        st.rerun()

if procesar:
    if not excel_file or not pdf_files:
        st.error("Debes subir un archivo Excel y al menos un archivo PDF.")
    else:
        with st.spinner("Procesando archivos y comparando VINs..."):
            try:
                # ------ Lógica de procesamiento (sin cambios) ------
                vins_validos, vins_invalidos = leer_excel_vins(excel_file)
                textos_pdf = {}
                texto_concatenado_pdf = ""
                for pdf_file in pdf_files:
                    texto = leer_pdf(pdf_file)
                    textos_pdf[pdf_file.name] = texto
                    texto_concatenado_pdf += texto

                vin_excel_count = Counter(vins_validos)
                vin_unicos_excel = set(vin_excel_count.keys())
                vin_repetidos_excel = [vin for vin, count in vin_excel_count.items() if count > 1]

                vin_encontrados_en_pdf = {}
                for vin in vin_unicos_excel:
                    archivos_donde_aparece = []
                    for nombre_archivo, texto in textos_pdf.items():
                        if buscar_vin_flexible(vin, texto):
                            archivos_donde_aparece.append(nombre_archivo)
                    if archivos_donde_aparece:
                        vin_encontrados_en_pdf[vin] = archivos_donde_aparece

                vin_solo_en_excel = sorted(list(vin_unicos_excel - set(vin_encontrados_en_pdf.keys())))

                # ------ Lógica para encontrar VINs solo en PDF (ya existía) ------
                texto_pdf_sin_espacios = re.sub(r'\s', '', texto_concatenado_pdf)
                vin_regex = re.compile(r"[A-HJ-NPR-Z0-9]{17}")
                posibles_vins_en_pdf_crudos = vin_regex.findall(texto_pdf_sin_espacios)
                
                vin_solo_en_pdf = set()
                for vin_posible in posibles_vins_en_pdf_crudos:
                    # Se normaliza y valida cada VIN potencial encontrado en los PDFs
                    if es_vin_plausible(vin_posible) and vin_posible not in vin_unicos_excel:
                        vin_solo_en_pdf.add(vin_posible)
                
                # ------ Presentación de resultados (sin cambios en la lógica) ------
                st.subheader("✅ Resumen de Resultados")
                st.write(f"Total de VINs válidos únicos en Excel: **{len(vin_unicos_excel)}**")
                st.write(f"Total de coincidencias entre Excel y PDF: **{len(vin_encontrados_en_pdf)}**")
                st.write(f"Total de VINs solo en Excel (no encontrados en PDF): **{len(vin_solo_en_excel)}**")
                st.write(f"Total de VINs únicos encontrados solo en PDF: **{len(vin_solo_en_pdf)}**")
                st.write(f"Total de VINs repetidos en Excel: **{len(vin_repetidos_excel)}**")

                resultados = []
                # 1. VINs del Excel
                for vin in sorted(vin_unicos_excel):
                    encontrado_en = vin_encontrados_en_pdf.get(vin)
                    resultados.append({
                        "VIN": vin,
                        "Estado": "✅ Encontrado en PDF" if encontrado_en else "❌ No Encontrado",
                        "Archivos PDF": ", ".join(encontrado_en) if encontrado_en else "N/A",
                        "Repetido en Excel": "Sí" if vin in vin_repetidos_excel else "No"
                    })
                # 2. VINs solo en PDF
                for vin in sorted(vin_solo_en_pdf):
                    resultados.append({
                        "VIN": vin,
                        "Estado": "📄 Solo en PDF",
                        "Archivos PDF": "N/A", # No aplica, se encontró en el conjunto de PDFs
                        "Repetido en Excel": "N/A"
                    })
                # 3. VINs inválidos del Excel
                for item in vins_invalidos:
                    resultados.append({
                        "VIN": item['vin'],
                        "Estado": "⚠️ Formato Inválido",
                        "Archivos PDF": "N/A",
                        "Repetido en Excel": "N/A"
                    })

                df_resultados = pd.DataFrame(resultados)
                df_resultados.reset_index(drop=True, inplace=True)
                df_resultados.index = df_resultados.index + 1
                st.dataframe(df_resultados)

                # Convertir DataFrame a Excel en memoria para descarga
                buffer = io.BytesIO()
                with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                    df_resultados.to_excel(writer, index=True, sheet_name="Resultados")
                
                st.download_button(
                    label="📥 Descargar Resultados en Excel",
                    data=buffer.getvalue(),
                    file_name="Reporte_Comparacion_VIN.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

            except Exception as e:
                st.error(f"Ocurrió un error durante el procesamiento: {e}")
