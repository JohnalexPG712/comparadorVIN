import streamlit as st
import pandas as pd
import fitz  # PyMuPDF
import re
import io
from collections import Counter

# --------------------------------------------------------------------------
# Funciones de Procesamiento y Validación (con nueva lógica adaptativa)
# --------------------------------------------------------------------------

def normalizar_vin(vin):
    """Elimina espacios y convierte a mayúsculas."""
    return (str(vin).replace(" ", "").replace("\r", "").replace("\n", "").replace("\t", "")).upper() if vin else ""

def tiene_formato_base_vin(vin):
    """Verifica únicamente la estructura base de un VIN: 17 caracteres sin I, O, Q."""
    return bool(re.fullmatch(r"[A-HJ-NPR-Z0-9]{17}", normalizar_vin(vin)))

def aprender_patrones_vin(lista_vins_validos):
    """
    NUEVO: Aprende los prefijos (WMI) de una lista de VINs de referencia.
    Extrae los primeros 3 caracteres de cada VIN y devuelve un conjunto de prefijos únicos.
    """
    if not lista_vins_validos:
        return set()
    prefijos = {vin[:3] for vin in lista_vins_validos}
    return prefijos

def crear_validador_dinamico(prefijos_aprendidos):
    """
    NUEVO: Crea y devuelve una función de validación que utiliza los prefijos aprendidos.
    Esto evita pasar la lista de prefijos a todas las funciones.
    """
    def validador(vin):
        vin_limpio = normalizar_vin(vin)
        if not tiene_formato_base_vin(vin_limpio):
            return False
        if not prefijos_aprendidos: # Si no se aprendió ningún patrón, solo valida el formato base
            return True
        return vin_limpio[:3] in prefijos_aprendidos
    return validador

def leer_excel_vins_base(excel_file):
    """
    Lee el Excel y realiza solo la validación de formato base (17 caracteres).
    No juzga la validez del prefijo todavía.
    """
    vins_con_formato_correcto = []
    vins_invalidos = []
    
    excel_file.seek(0)
    extension = excel_file.name.split('.')[-1].lower()
    df = pd.read_excel(excel_file, engine='xlrd' if extension == 'xls' else 'openpyxl', header=None, dtype=str, keep_default_na=False)
    
    start_row = 0
    if len(df.columns) > 1:
        # Se asume que si la primera celda no parece un VIN, es un encabezado.
        if not tiene_formato_base_vin(df.iloc[0, 1]):
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
st.title("🔬 Comparador de VINs Adaptativo: Excel vs PDF")
st.info("Esta herramienta aprende la estructura de los VINs de tu archivo Excel para realizar una búsqueda más precisa en los PDFs.")

excel_file = st.file_uploader("1. Sube el archivo Excel (FMM) de referencia", type=["xlsx", "xls"], key="excel_uploader")
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
        with st.spinner("Procesando... Aprendiendo patrones y comparando archivos..."):
            try:
                # --- Flujo de Procesamiento Adaptativo ---
                
                # 1. Lectura inicial y validación de formato base del Excel
                vins_excel_formato_base, vins_invalidos_formato = leer_excel_vins_base(excel_file)
                
                # 2. Aprender los patrones (prefijos) de los VINs del Excel
                prefijos_aprendidos = aprender_patrones_vin(vins_excel_formato_base)
                if not prefijos_aprendidos:
                    st.warning("No se pudo aprender ningún patrón de VINs del archivo Excel. La búsqueda en PDF puede ser menos precisa.")
                
                # 3. Crear un validador basado en los patrones aprendidos
                es_vin_valido = crear_validador_dinamico(prefijos_aprendidos)

                # 4. Clasificar finalmente los VINs del Excel
                vins_validos_finales = [vin for vin in vins_excel_formato_base if es_vin_valido(vin)]
                # Los VINs con formato correcto pero prefijo incorrecto se mueven a "inválidos"
                vins_con_prefijo_invalido = [vin for vin in vins_excel_formato_base if not es_vin_valido(vin)]
                for vin in vins_con_prefijo_invalido:
                    vins_invalidos_formato.append({"vin": vin})

                # --- Lógica de Comparación (similar a antes, pero usando el nuevo validador) ---
                textos_pdf = {pdf.name: leer_pdf(pdf) for pdf in pdf_files}
                texto_concatenado_pdf = " ".join(textos_pdf.values())
                
                vin_excel_count = Counter(vins_validos_finales)
                vin_unicos_excel = set(vin_excel_count.keys())
                vin_repetidos_excel = {vin for vin, count in vin_excel_count.items() if count > 1}

                vin_encontrados_en_pdf = {vin: [name for name, texto in textos_pdf.items() if buscar_vin_flexible(vin, texto)] for vin in vin_unicos_excel}
                vin_encontrados_en_pdf = {vin: files for vin, files in vin_encontrados_en_pdf.items() if files}

                vin_solo_en_excel = sorted(list(vin_unicos_excel - set(vin_encontrados_en_pdf.keys())))

                # --- Búsqueda en PDF usando el validador dinámico ---
                vin_regex = re.compile(r"[A-HJ-NPR-Z0-9]{17}")
                posibles_vins_en_pdf_crudos = vin_regex.findall(re.sub(r'\s', '', texto_concatenado_pdf))
                
                vin_solo_en_pdf = set()
                for vin_posible in posibles_vins_en_pdf_crudos:
                    if es_vin_valido(vin_posible) and vin_posible not in vin_unicos_excel:
                        vin_solo_en_pdf.add(vin_posible)
                
                # --- Presentación de Resultados ---
                st.subheader("✅ Resumen de Resultados")
                st.write(f"Patrones de VIN aprendidos del Excel: **{', '.join(sorted(prefijos_aprendidos)) if prefijos_aprendidos else 'Ninguno'}**")
                st.write(f"Total de VINs válidos únicos en Excel: **{len(vin_unicos_excel)}**")
                st.write(f"Total de coincidencias (Excel -> PDF): **{len(vin_encontrados_en_pdf)}**")
                st.write(f"Total de VINs solo en Excel: **{len(vin_solo_en_excel)}**")
                st.write(f"Total de VINs encontrados solo en PDF (con patrón válido): **{len(vin_solo_en_pdf)}**")

                resultados = []
                for vin in sorted(vin_unicos_excel):
                    encontrado_en = vin_encontrados_en_pdf.get(vin)
                    resultados.append({"VIN": vin, "Estado": "✅ Encontrado en PDF" if encontrado_en else "❌ No Encontrado", "Archivos PDF": ", ".join(encontrado_en) if encontrado_en else "N/A", "Repetido en Excel": "Sí" if vin in vin_repetidos_excel else "No"})
                for vin in sorted(vin_solo_en_pdf):
                    resultados.append({"VIN": vin, "Estado": "📄 Solo en PDF", "Archivos PDF": "N/A", "Repetido en Excel": "N/A"})
                for item in vins_invalidos_formato:
                    resultados.append({"VIN": item['vin'], "Estado": "⚠️ Formato o Patrón Inválido", "Archivos PDF": "N/A", "Repetido en Excel": "N/A"})

                df_resultados = pd.DataFrame(resultados).set_index(pd.Index(range(1, len(resultados) + 1)))
                st.dataframe(df_resultados)

                buffer = io.BytesIO()
                with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                    df_resultados.to_excel(writer, index=True, sheet_name="Resultados")
                st.download_button(label="📥 Descargar Resultados en Excel", data=buffer.getvalue(), file_name="Reporte_Comparacion_VIN.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

            except Exception as e:
                st.error(f"Ocurrió un error durante el procesamiento: {e}")
