import streamlit as st
import pandas as pd
import fitz  # PyMuPDF
import re
import io
from collections import Counter

# --------------------------------------------------------------------------
# Lógica de Validación de VIN - ¡MODIFICADO!
# --------------------------------------------------------------------------

# NUEVO: Lista centralizada de prefijos válidos.
# ¡Puedes añadir o quitar prefijos de este conjunto para ajustar la validación!
PREFIJOS_VALIDOS = {"9G5", "9G6", "9G7"}

def normalizar_vin(vin):
    """Elimina espacios y convierte a mayúsculas."""
    return (str(vin).replace(" ", "").replace("\r", "").replace("\n", "").replace("\t", "")).upper() if vin else ""

def es_vin_valido_con_prefijo(vin):
    """
    Realiza una validación completa y estricta del VIN:
    1. Normaliza el string de entrada.
    2. Verifica el formato estándar de 17 caracteres (sin I, O, Q).
    3. Comprueba que comience con un prefijo de la lista PREFIJOS_VALIDOS.
    """
    vin_limpio = normalizar_vin(vin)
    
    # 1. Validación de formato estándar
    if not re.fullmatch(r"[A-HJ-NPR-Z0-9]{17}", vin_limpio):
        return False
        
    # 2. Validación de prefijo (WMI)
    if vin_limpio[:3] not in PREFIJOS_VALIDOS:
        return False
        
    return True

# --------------------------------------------------------------------------
# Funciones de Lectura de Archivos (modificadas para usar la nueva validación)
# --------------------------------------------------------------------------

def buscar_vin_flexible(vin, texto_pdf):
    """Busca un VIN en el texto permitiendo espacios entre sus caracteres."""
    regex_pattern = r"".join(re.escape(char) + r"[\s\r\n]*" for char in vin)
    return bool(re.search(regex_pattern, texto_pdf, re.IGNORECASE))

def leer_excel_vins(excel_file):
    vins_validos = []
    vins_invalidos = []
    
    excel_file.seek(0)
    extension = excel_file.name.split('.')[-1].lower()
    df = pd.read_excel(excel_file, engine='xlrd' if extension == 'xls' else 'openpyxl', header=None, dtype=str, keep_default_na=False)
    
    start_row = 0
    if len(df.columns) > 1:
        primer_posible_vin = df.iloc[0, 1]
        # MODIFICADO: Usa la nueva función para detectar si la primera fila es un encabezado.
        if not es_vin_valido_con_prefijo(primer_posible_vin):
            start_row = 1
        
        vins_crudos = df.iloc[start_row:, 1].tolist()
        for vin_crudo in vins_crudos:
            vin_normalizado = normalizar_vin(vin_crudo)
            if vin_normalizado:
                # MODIFICADO: Usa la nueva función para clasificar los VINs.
                if es_vin_valido_con_prefijo(vin_normalizado):
                    vins_validos.append(vin_normalizado)
                else:
                    vins_invalidos.append({"vin": vin_crudo, "length": len(vin_normalizado)})
    return vins_validos, vins_invalidos

def leer_pdf(file):
    texto_completo = ""
    file.seek(0)
    doc = fitz.open(stream=file.read(), filetype="pdf")
    for page in doc:
        texto_completo += page.get_text() + " "
    doc.close()
    texto_completo = re.sub(r'\s+', ' ', texto_completo).upper()
    return texto_completo

# --------------------------------------------------------------------------
# Interfaz de la aplicación Streamlit
# --------------------------------------------------------------------------

st.set_page_config(page_title="Comparador de VINs", layout="centered")
st.title("🔍 Comparador de VINs: Excel (FMM) vs Documentos PDF")

excel_file = st.file_uploader("1. Sube el archivo Excel (FMM)", type=["xlsx", "xls"], key="excel_uploader")
pdf_files = st.file_uploader("2. Sube los archivos PDF de soporte", type=["pdf"], accept_multiple_files=True, key="pdf_uploader")

col1, col2 = st.columns([1.5, 2])

with col1:
    procesar = st.button("3. Procesar y Comparar", type="primary")

with col2:
    if st.button("🧹 Limpiar y Empezar de Nuevo"):
        st.rerun()

if procesar:
    if not excel_file or not pdf_files:
        st.error("Debes subir un archivo Excel y al menos un archivo PDF.")
    else:
        with st.spinner("Procesando archivos y comparando VINs..."):
            try:
                # --- Extracción y comparación ---
                vins_validos, vins_invalidos = leer_excel_vins(excel_file)
                textos_pdf = {pdf.name: leer_pdf(pdf) for pdf in pdf_files}
                texto_concatenado_pdf = " ".join(textos_pdf.values())
                
                vin_excel_count = Counter(vins_validos)
                vin_unicos_excel = set(vin_excel_count.keys())
                vin_repetidos_excel = {vin for vin, count in vin_excel_count.items() if count > 1}

                vin_encontrados_en_pdf = {}
                for vin in vin_unicos_excel:
                    archivos_donde_aparece = [name for name, texto in textos_pdf.items() if buscar_vin_flexible(vin, texto)]
                    if archivos_donde_aparece:
                        vin_encontrados_en_pdf[vin] = archivos_donde_aparece

                vin_solo_en_excel = sorted(list(vin_unicos_excel - set(vin_encontrados_en_pdf.keys())))

                # --- Búsqueda de VINs solo en PDF con la nueva validación ---
                vin_regex = re.compile(r"[A-HJ-NPR-Z0-9]{17}")
                posibles_vins_en_pdf_crudos = vin_regex.findall(re.sub(r'\s', '', texto_concatenado_pdf))
                
                vin_solo_en_pdf = set()
                for vin_posible in posibles_vins_en_pdf_crudos:
                    # MODIFICADO: Usa la nueva función de validación estricta.
                    if es_vin_valido_con_prefijo(vin_posible) and vin_posible not in vin_unicos_excel:
                        vin_solo_en_pdf.add(vin_posible)
                
                # --- Presentación de resultados ---
                st.subheader("✅ Resumen de Resultados")
                st.write(f"Total de VINs válidos únicos en Excel: **{len(vin_unicos_excel)}**")
                st.write(f"Total de coincidencias entre Excel y PDF: **{len(vin_encontrados_en_pdf)}**")
                st.write(f"Total de VINs solo en Excel (no encontrados en PDF): **{len(vin_solo_en_excel)}**")
                st.write(f"Total de VINs únicos encontrados solo en PDF: **{len(vin_solo_en_pdf)}**")
                st.write(f"Total de VINs repetidos en Excel: **{len(vin_repetidos_excel)}**")

                resultados = []
                for vin in sorted(vin_unicos_excel):
                    encontrado_en = vin_encontrados_en_pdf.get(vin)
                    resultados.append({
                        "VIN": vin,
                        "Estado": "✅ Encontrado en PDF" if encontrado_en else "❌ No Encontrado",
                        "Archivos PDF": ", ".join(encontrado_en) if encontrado_en else "N/A",
                        "Repetido en Excel": "Sí" if vin in vin_repetidos_excel else "No"
                    })
                for vin in sorted(vin_solo_en_pdf):
                    resultados.append({"VIN": vin, "Estado": "📄 Solo en PDF", "Archivos PDF": "N/A", "Repetido en Excel": "N/A"})
                for item in vins_invalidos:
                    resultados.append({"VIN": item['vin'], "Estado": "⚠️ Formato Inválido", "Archivos PDF": "N/A", "Repetido en Excel": "N/A"})

                df_resultados = pd.DataFrame(resultados).set_index(pd.Index(range(1, len(resultados) + 1)))
                st.dataframe(df_resultados)

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
