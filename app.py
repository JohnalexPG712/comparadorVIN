import streamlit as st
import pandas as pd
import fitz  # PyMuPDF
import re
import io
from collections import Counter

# --------------------------------------------------------------------------
# PATRÓN DE DISEÑO DEFINITIVO: REINICIO POR CAMBIO DE CLAVE (KEY)
# --------------------------------------------------------------------------

# 1. Se inicializa un contador en el estado de la sesión.
#    Este contador se usará para generar claves únicas para los widgets.
if 'upload_counter' not in st.session_state:
    st.session_state.upload_counter = 0

# 2. Se define la función callback que incrementará el contador.
def reiniciar_widgets():
    """Incrementa el contador para forzar la recreación de los widgets."""
    st.session_state.upload_counter += 1

# --------------------------------------------------------------------------
# Funciones de backend (sin cambios)
# --------------------------------------------------------------------------

def normalizar_vin(vin):
    return (str(vin).replace(" ", "").replace("\r", "").replace("\n", "").replace("\t", "")).upper() if vin else ""

def tiene_formato_base_vin(vin):
    return bool(re.fullmatch(r"[A-HJ-NPR-Z0-9]{17}", normalizar_vin(vin)))

def aprender_patrones_vin(lista_vins_validos):
    if not lista_vins_validos:
        return set()
    return {vin[:3] for vin in lista_vins_validos}

def crear_validador_dinamico(prefijos_aprendidos):
    def validador(vin):
        vin_limpio = normalizar_vin(vin)
        if not tiene_formato_base_vin(vin_limpio):
            return False
        if not prefijos_aprendidos:
            return True
        return vin_limpio[:3] in prefijos_aprendidos
    return validador

def leer_excel_vins_base(excel_file):
    vins_con_formato_correcto, vins_invalidos = [], []
    excel_file.seek(0)
    extension = excel_file.name.split('.')[-1].lower()
    df = pd.read_excel(excel_file, engine='xlrd' if extension == 'xls' else 'openpyxl', header=None, dtype=str, keep_default_na=False)
    
    start_row = 0
    if len(df.columns) > 1 and not df.empty:
        if not tiene_formato_base_vin(df.iloc[0, 1] if len(df.iloc[0]) > 1 else ""):
            start_row = 1
        vins_crudos = df.iloc[start_row:, 1].tolist()
        for vin_crudo in vins_crudos:
            vin_normalizado = normalizar_vin(vin_crudo)
            if vin_normalizado:
                if tiene_formato_base_vin(vin_normalizado):
                    vins_con_formato_correcto.append(vin_normalizado)
                else:
                    vins_invalidos.append({"vin": vin_crudo})
    return vins_con_formato_correcto, vins_invalidos

def leer_pdf(file):
    texto_completo = ""
    file.seek(0)
    doc = fitz.open(stream=file.read(), filetype="pdf")
    for page in doc:
        texto_completo += page.get_text() + " "
    doc.close()
    return re.sub(r'\s+', ' ', texto_completo).upper()

def buscar_vin_flexible(vin, texto_pdf):
    regex_pattern = r"".join(re.escape(char) + r"[\s\r\n]*" for char in vin)
    return bool(re.search(regex_pattern, texto_pdf, re.IGNORECASE))

# --------------------------------------------------------------------------
# Interfaz de la aplicación Streamlit
# --------------------------------------------------------------------------

st.set_page_config(page_title="Comparador de VINs Adaptativo", layout="centered")
st.title("🔬Comparador de VINs: Excel (FMM) vs Documentos PDF")
st.info("Permite comparar y verificar los VIN entre el Formulario de Movimiento de Mercancías (FMM) y los documentos soporte de las transacciones 329, 401, 422 y 436 (DI, DUTA, Factura o Remisión). Comparador de VINs Adaptativo.")

# 3. Se usan claves dinámicas para los widgets de carga de archivos.
#    Cuando el contador cambie, estas claves cambiarán, forzando un reinicio completo de los widgets.
excel_file = st.file_uploader("1. Sube el archivo Excel (FMM) de referencia", type=["xlsx", "xls"], key=f"excel_uploader_{st.session_state.upload_counter}")
pdf_files = st.file_uploader("2. Sube los archivos PDF de soporte", type=["pdf"], accept_multiple_files=True, key=f"pdf_uploader_{st.session_state.upload_counter}")

col1, col2 = st.columns([1.5, 2])
with col1:
    procesar = st.button("3. Procesar y Verificar", type="primary")
    
    # --- Cambiar el color del botón "Procesar y Verificar" a verde ---
st.markdown("""
    <style>
    /* Botón Procesar y Verificar */
    div.stButton > button:first-child {
        background-color: #1470A8;  /* Azul */
        color: white;
        border-radius: 8px;
        height: 3em;
        width: 100%;
        font-weight: bold;
        border: none;
    }
    div.stButton > button:first-child:hover {
        background-color: #FCFAFA;  /* color blanco */
        color: white;
    }
    </style>
""", unsafe_allow_html=True)


with col2:
    # 4. El botón de limpieza ahora llama a la función que incrementa el contador.
    st.button("🧹 Limpiar Resultados", on_click=reiniciar_widgets)

if procesar:
    if not excel_file or not pdf_files:
        st.error("Debes subir un archivo Excel y al menos un archivo PDF.")
    else:
        with st.spinner("Procesando..."):
            try:
                # ... (El resto del código de procesamiento no cambia)
                vins_excel_formato_base, vins_invalidos_formato = leer_excel_vins_base(excel_file)
                prefijos_aprendidos = aprender_patrones_vin(vins_excel_formato_base)
                
                es_vin_valido = crear_validador_dinamico(prefijos_aprendidos)
                vins_validos_finales = [vin for vin in vins_excel_formato_base if es_vin_valido(vin)]
                vins_con_prefijo_invalido = [vin for vin in vins_excel_formato_base if not es_vin_valido(vin)]
                for vin in vins_con_prefijo_invalido:
                    vins_invalidos_formato.append({"vin": vin})

                textos_pdf = {pdf.name: leer_pdf(pdf) for pdf in pdf_files}
                texto_concatenado_pdf = " ".join(textos_pdf.values())
                
                vin_excel_count = Counter(vins_validos_finales)
                vin_unicos_excel = set(vin_excel_count.keys())
                vin_repetidos_excel = {vin for vin, count in vin_excel_count.items() if count > 1}

                vin_encontrados_en_pdf = {vin: [name for name, texto in textos_pdf.items() if buscar_vin_flexible(vin, texto)] for vin in vin_unicos_excel}
                vin_encontrados_en_pdf = {vin: files for vin, files in vin_encontrados_en_pdf.items() if files}

                vin_solo_en_excel = sorted(list(vin_unicos_excel - set(vin_encontrados_en_pdf.keys())))

                vin_regex = re.compile(r"[A-HJ-NPR-Z0-9]{17}")
                posibles_vins_en_pdf_crudos = vin_regex.findall(re.sub(r'\s', '', texto_concatenado_pdf))
                
                vin_solo_en_pdf = {vin for vin in posibles_vins_en_pdf_crudos if es_vin_valido(vin) and vin not in vin_unicos_excel}
                
                st.subheader("✅ Resumen de Resultados")
                if prefijos_aprendidos:
                    st.write(f"Patrones de VIN aprendidos del Excel: **{', '.join(sorted(prefijos_aprendidos))}**")
                
                st.write(f"Total de VINs válidos únicos en Excel (FMM): **{len(vin_unicos_excel)}**")
                st.write(f"Total de coincidencias (FMM -> PDF): **{len(vin_encontrados_en_pdf)}**")
                st.write(f"Total de VINs solo en Excel (FMM): **{len(vin_solo_en_excel)}**")
                st.write(f"Total de VINs encontrados solo en PDF: **{len(vin_solo_en_pdf)}**")

                resultados = []
                for vin in sorted(vin_unicos_excel):
                    encontrado_en = vin_encontrados_en_pdf.get(vin)
                    resultados.append({"VIN": vin, "Estado": "✅ Encontrado en PDF" if encontrado_en else "❌ No Encontrado", "Archivos PDF": ", ".join(encontrado_en) if encontrado_en else "N/A", "Repetido en Excel": "Sí" if vin in vin_repetidos_excel else "No"})
                for vin in sorted(vin_solo_en_pdf):
                    resultados.append({"VIN": vin, "Estado": "📄 Solo en PDF", "Archivos PDF": "N/A", "Repetido en Excel": "N/A"})
                for item in vins_invalidos_formato:
                    resultados.append({"VIN": item['vin'], "Estado": "⚠️ Formato o Patrón Inválido", "Archivos PDF": "N/A", "Repetido en Excel": "N/A"})

                if resultados:
                    df_resultados = pd.DataFrame(resultados).set_index(pd.Index(range(1, len(resultados) + 1)))
                    st.dataframe(df_resultados)

                    buffer = io.BytesIO()
                    with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                        df_resultados.to_excel(writer, index=True, sheet_name="Resultados")
                    st.download_button(label="📥 Descargar Resultados en Excel", data=buffer.getvalue(), file_name="Reporte_Comparacion_VIN.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
                else:
                    st.info("No se encontraron VINs para mostrar en los resultados.")

            except Exception as e:
                st.error(f"Ocurrió un error durante el procesamiento: {e}")

