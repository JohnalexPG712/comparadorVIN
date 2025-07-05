import pandas as pd
import fitz  # PyMuPDF
import re
import os
import glob # Para encontrar archivos por patrón
from collections import Counter

# --- Define la ruta principal de la carpeta aquí ---
# Esta es la carpeta fija donde el script buscará el archivo Excel y todos los PDFs.
CARPETA_PRINCIPAL_FIJA = "E:/Users/Lenovo/Desktop/PYTHON/VIN"
# ----------------------------------------------------


# --- Funciones Auxiliares ---

def normalizar_vin(vin):
    """Normaliza el VIN removiendo espacios y convirtiendo a mayúsculas."""
    return (str(vin).replace(" ", "").replace("\r", "").replace("\n", "").replace("\t", "")).upper() if vin else ""

def validar_vin(vin):
    """Valida si un VIN normalizado cumple el formato estándar de 17 caracteres."""
    vin_limpio = normalizar_vin(vin)
    # JS usa /^[A-HJ-NPR-Z0-9]{17}$/. Equivalente en Python:
    return bool(re.fullmatch(r"[A-HJ-NPR-Z0-9]{17}", vin_limpio))

def es_vin_plausible(vin):
    """Verificación heurística de cadenas plausibles de 17 caracteres encontradas en el PDF (similar a la lógica JS)."""
    if len(vin) != 17:
        return False
    # Verifica si hay demasiadas letras consecutivas (verificación JS)
    if len(re.findall(r"[A-Z]{6,}", vin)) > 0:
        return False
    # Verifica WMI (primeros 3 caracteres) - JS verifica al menos 2 letras
    wmi = vin[:3]
    if len(re.findall(r"[A-Z]", wmi)) < 2:
        return False
    # Verifica el serial (últimos 6 caracteres) - JS verifica los últimos 6, espera al menos 4 dígitos en los últimos 11.
    # La lógica original de JS era "los últimos 11 caracteres tienen al menos 4 dígitos". Mantengamos esa lógica.
    serial_part = vin[6:] # Últimos 11 caracteres
    if len(re.findall(r"[0-9]", serial_part)) < 4:
       return False
    return True

def buscar_vin_flexible(vin, texto_pdf):
    """
    Busca un VIN en el texto PDF, permitiendo espacios o saltos de línea
    entre sus caracteres. Similar a la lógica JavaScript.
    """
    if not vin or not texto_pdf:
        return False
    # Escapa caracteres especiales de regex en el VIN
    escaped_vin = re.escape(vin)
    # Crea regex permitiendo cero o más caracteres de espacio en blanco entre caracteres del VIN
    regex_pattern = r"".join(re.escape(char) + r"\s*" for char in vin)[:-3] # remueve el \s* final
    # Usa re.search para buscar una coincidencia en cualquier parte del texto, sin distinguir mayúsculas/minúsculas
    return bool(re.search(regex_pattern, texto_pdf, re.IGNORECASE))

# --- Funciones de Lógica Principal ---

def leer_excel_vins(excel_path):
    """Lee VINs de la segunda columna (índice 1) de un archivo Excel especificado por ruta."""
    vins_validos = []
    vins_invalidos = []

    try:
        # Usa pandas para leer el archivo Excel
        # header=None para leer sin asumir una fila de encabezado inicialmente
        df = pd.read_excel(excel_path, header=None, dtype=str, keep_default_na=False)

        if df.empty:
            print(f"Advertencia: El archivo Excel '{os.path.basename(excel_path)}' está vacío.")
            return [], []

        # Asume que los VINs están en la segunda columna (índice 1) según row[1] de JS
        # Verifica si la primera celda de la segunda columna parece un VIN
        # Si no, asume que hay una fila de encabezado y la omite.
        start_row = 0
        if len(df.columns) > 1: # Asegura que la segunda columna exista
            primer_posible_vin = normalizar_vin(df.iloc[0, 1])
            if not validar_vin(primer_posible_vin):
                 start_row = 1 # Asume que la primera fila es encabezado si no valida como VIN.

            # Procesa la columna seleccionada a partir de start_row
            vins_crudos = df.iloc[start_row:, 1].tolist()

            for vin_crudo in vins_crudos:
                vin_normalizado = normalizar_vin(vin_crudo)
                if vin_normalizado: # Solo procesa si no está vacío después de la normalización
                    if validar_vin(vin_normalizado):
                        vins_validos.append(vin_normalizado)
                    else:
                        vins_invalidos.append({"vin": vin_crudo, "length": len(vin_normalizado)}) # Almacena el original para el reporte
        else:
             print(f"Advertencia: El archivo Excel '{os.path.basename(excel_path)}' no tiene al menos 2 columnas. No se pudieron leer VINs.")


    except FileNotFoundError:
        # Este error no debería ocurrir si se verificó antes de llamar, pero se mantiene
        print(f"Error: Archivo Excel no encontrado en '{excel_path}'")
        return None, None
    except Exception as e:
        print(f"Error al leer el archivo Excel '{os.path.basename(excel_path)}': {e}")
        return None, None

    return vins_validos, vins_invalidos

def leer_pdfs_en_carpeta(pdf_folder_path):
    """Lee el texto de todos los archivos PDF en la carpeta especificada."""
    textos_pdf = {}
    texto_concatenado_pdf = ""

    # Esta validación ahora es CRUCIAL porque la ruta es fija
    if not os.path.isdir(pdf_folder_path):
        # Este caso se maneja ANTES de llamar a esta función en el main block
        print(f"Error interno: La ruta de carpeta PDF proporcionada no es válida: '{pdf_folder_path}'")
        return None, None # Indica un error irrecuperable en la ruta

    pdf_files = [f for f in os.listdir(pdf_folder_path) if f.lower().endswith('.pdf')]

    if not pdf_files:
        print(f"Advertencia: No se encontraron archivos PDF en la carpeta '{pdf_folder_path}'")
        return {}, "" # Devuelve dic/string vacíos si no hay PDFs pero la carpeta es válida

    print(f"Procesando {len(pdf_files)} archivo(s) PDF dentro de la carpeta '{os.path.basename(pdf_folder_path)}'...")

    for pdf_file in pdf_files:
        file_path = os.path.join(pdf_folder_path, pdf_file)
        try:
            doc = fitz.open(file_path)
            texto_completo = ""
            for page_num in range(doc.page_count):
                page = doc.load_page(page_num)
                texto_completo += page.get_text() + " "
            doc.close()
            textos_pdf[pdf_file] = texto_completo.upper() # Convierte a mayúsculas como en JS
            texto_concatenado_pdf += textos_pdf[pdf_file]
            print(f" - Leído: {pdf_file}")
        except Exception as e:
            print(f"Error al leer el archivo PDF '{pdf_file}': {e}")
            # Almacena texto vacío para este archivo pero continúa
            textos_pdf[pdf_file] = "" # Asegura que el archivo esté en el dict, aunque con texto vacío

    # Normaliza el texto concatenado reemplazando múltiples espacios/saltos de línea con un solo espacio
    texto_concatenado_pdf = re.sub(r'\s+', ' ', texto_concatenado_pdf)

    return textos_pdf, texto_concatenado_pdf

def comparar_vins(excel_path, pdf_folder_path):
    """Realiza la comparación principal de VINs."""
    print("Iniciando comparación...")
    print(f"Archivo Excel (FMM) encontrado: {os.path.basename(excel_path)}")
    print(f"Carpeta de origen (Excel y PDFs): {os.path.abspath(pdf_folder_path)}") # Imprime la carpeta de origen única

    vins_del_excel_validos, vins_del_excel_invalidos = leer_excel_vins(excel_path)
    if vins_del_excel_validos is None: # Maneja errores graves durante la lectura de Excel
         print("No se pudo continuar debido a un error al leer el archivo Excel.")
         return
    if not vins_del_excel_validos and not vins_del_excel_invalidos:
         print("El archivo Excel está vacío o no contiene VINs válidos/inválidos en la segunda columna.")
         # Si no hay VINs en el Excel, no hay nada que comparar contra los PDFs
         return


    textos_pdf, texto_concatenado_pdf = leer_pdfs_en_carpeta(pdf_folder_path)
    # Si leer_pdfs_en_carpeta devolvió None, significa que la ruta fija no era válida (manejado en main)
    if textos_pdf is None:
         return
    if not textos_pdf and not texto_concatenado_pdf: # Maneja caso sin PDFs en la carpeta (pero carpeta válida)
         print("\nNo se encontraron archivos PDF en la carpeta fija. No se realizará comparación de VINs de Excel contra PDFs.")
         # Aún así, podemos reportar VINs inválidos o repetidos del Excel si los hay
         if vins_del_excel_invalidos or any(count > 1 for count in Counter(vins_del_excel_validos).values()):
             print("\nSin embargo, se reportarán los VINs inválidos o repetidos encontrados en el Excel.")
             # Proceder a generar la parte del reporte de Excel solamente si hay algo que reportar
             # Pasamos diccionarios/sets vacíos para los resultados relacionados con PDFs
             generar_reporte(excel_path, pdf_folder_path, vins_del_excel_validos, vins_del_excel_invalidos, {}, set(), [])
         return


    # --- Lógica de Análisis (Traducida de JS) ---

    # Cuenta los VINs en Excel
    vin_excel_count = Counter(vins_del_excel_validos)
    vin_unicos_excel = set(vin_excel_count.keys())
    vin_repetidos_excel = [vin for vin, count in vin_excel_count.items() if count > 1]

    # Encuentra dónde aparece cada VIN único de Excel en los PDFs
    vin_encontrados_en_pdf = {}
    for vin in sorted(list(vin_unicos_excel)): # Itera sobre los VINs únicos ordenados de Excel
        archivos_donde_aparece = []
        for nombre_archivo, texto in textos_pdf.items():
            # Asegura que haya texto para buscar. Evita procesar archivos que dieron error al leer
            if texto and buscar_vin_flexible(vin, texto):
                archivos_donde_aparece.append(nombre_archivo)
        if archivos_donde_aparece:
            vin_encontrados_en_pdf[vin] = archivos_donde_aparece

    # Identifica los VINs que solo están en Excel (los válidos)
    vin_solo_en_excel = sorted(list(vin_unicos_excel - set(vin_encontrados_en_pdf.keys())))

    # Encuentra posibles VINs que solo están en el PDF
    # Elimina todos los espacios del texto concatenado para una búsqueda de patrón VIN estricta
    texto_pdf_sin_espacios = re.sub(r'\s', '', texto_concatenado_pdf)
    # Usa regex con lookarounds para encontrar cadenas de 17 caracteres que no formen parte de secuencias más largas
    # Esto coincide con la regex JS /(?<![A-HJ-NPR-Z0-9])[A-HJ-NPR-Z0-9]{17}(?![A-HJ-NPR-Z0-9])/g
    vin_regex_plausible_pdf = re.compile(r"(?<![A-HJ-NPR-Z0-9])([A-HJ-NPR-Z0-9]{17})(?![A-HJ-NPR-Z0-9])", re.IGNORECASE)
    posibles_vins_en_pdf_crudos = vin_regex_plausible_pdf.findall(texto_pdf_sin_espacios)

    # Filtra VINs plausibles del PDF: deben ser plausibles y NO estar en los VINs únicos de Excel
    vin_solo_en_pdf = set()
    for vin_posible in posibles_vins_en_pdf_crudos:
         vin_normalizado = normalizar_vin(vin_posible) # Asegura la normalización si la regex no capturó completamente
         if vin_normalizado and es_vin_plausible(vin_normalizado) and vin_normalizado not in vin_unicos_excel:
              vin_solo_en_pdf.add(vin_normalizado)
    vin_solo_en_pdf = sorted(list(vin_solo_en_pdf)) # Ordena para una salida consistente

    # --- Generación de Reporte ---
    generar_reporte(
        excel_path, # Pasamos el path del Excel encontrado
        pdf_folder_path, # Pasamos la misma carpeta principal
        vins_del_excel_validos,
        vins_del_excel_invalidos,
        vin_encontrados_en_pdf,
        vin_solo_en_pdf,
        vin_repetidos_excel
    )


def generar_reporte(
    excel_path,
    source_folder_path, # Carpeta donde se encontraron los archivos
    vins_del_excel_validos,
    vins_del_excel_invalidos,
    vin_encontrados_en_pdf, # Diccionario {vin: [archivos]}
    vin_solo_en_pdf, # Lista de vins
    vin_repetidos_excel # Lista de vins
    ):

    vin_unicos_excel = set(vins_del_excel_validos)
    vin_solo_en_excel = sorted(list(vin_unicos_excel - set(vin_encontrados_en_pdf.keys())))


    print("\n" + "="*40)
    print("Resultados de la Comparación de VINs")
    print("="*40)

    print("\n--- Resumen ---")
    print(f"Archivo Excel (FMM) utilizado: {os.path.basename(excel_path)}")
    print(f"Carpeta de origen (Excel y PDFs): {os.path.abspath(source_folder_path)}") # Imprime la carpeta de origen
    print(f"Total de VINs válidos encontrados en Excel): {len(vins_del_excel_validos)}")
    print(f"Total de VINs válidos ÚNICOS encontrados en Excel: {len(vin_unicos_excel)}")
    print(f"Total de VINs válidos únicos de Excel encontrados en al menos un PDF: {len(vin_encontrados_en_pdf)}")
    print(f"Total de coincidencias encontradas (VINs de Excel encontrados en PDF): {len(vin_encontrados_en_pdf)}")
    print(f"VINs válidos repetidos en el Excel: {len(vin_repetidos_excel)}")
    print(f"Número de VINs válidos encontrados Únicamente en los (PDF): {len(vin_solo_en_pdf)}")
    print(f"Número de VINs encontrados Únicamente en el FMM (Excel) y no en los PDFs: {len(vin_solo_en_excel)}")


    if vins_del_excel_invalidos:
        print("\n--- VINs con formato inválido en Excel ---")
        for item in vins_del_excel_invalidos:
            print(f"VIN '{item['vin']}' (longitud: {item['length']}) no cumple el estándar (debe ser 17 caracteres A-HJ-NPR-Z0-9).")

    if vin_solo_en_pdf:
        print("\n--- VINs válidos encontrados Únicamente en los (PDF) ---")
        print(", ".join(vin_solo_en_pdf))

    if vin_repetidos_excel:
        print("\n--- VINs repetidos en el Excel ---")
        print(", ".join(vin_repetidos_excel))

    print("\n--- Detalle por VIN de Excel ---")
    if not vin_unicos_excel:
        print("No se encontraron VINs válidos en el archivo Excel para comparar contra PDFs.")
    else:
        # Reporta el detalle de cada VIN válido único de Excel, ordenado
        for vin in sorted(list(vin_unicos_excel)):
            if vin in vin_encontrados_en_pdf:
                print(f"VIN {vin}: ENCONTRADO en PDF(s): {', '.join(vin_encontrados_en_pdf[vin])}")
            else:
                print(f"VIN {vin}: NO ENCONTRADO en ningún PDF.")

    # --- Prepara datos para la Salida a Excel ---
    resultados_para_excel = []

    # Agrega todos los VINs válidos únicos de Excel
    for vin in sorted(list(vin_unicos_excel)):
        encontrado_en = vin_encontrados_en_pdf.get(vin)
        resultados_para_excel.append({
            "VIN": vin,
            "Origen VIN": "Excel (FMM)",
            "Estado Comparacion": "Encontrado en PDF" if encontrado_en else "No Encontrado en PDF",
            "Archivos PDF Coincidencia": ", ".join(encontrado_en) if encontrado_en else "N/A",
            "Notas": "Repetido en Excel" if vin in vin_repetidos_excel else ""
        })

    # Agrega VINs encontrados solo en PDF
    for vin in sorted(vin_solo_en_pdf):
         resultados_para_excel.append({
            "VIN": vin,
            "Origen VIN": "Solo en PDF",
            "Estado Comparacion": "Solo en PDF (No en Excel)",
            "Archivos PDF Coincidencia": "N/A (identificado por regex, no buscado)", # Aclaración
            "Notas": ""
         })

    # Agrega VINs inválidos del Excel
    for item in vins_del_excel_invalidos:
         resultados_para_excel.append({
             "VIN": item['vin'],
             "Origen VIN": "Excel (FMM)",
             "Estado Comparacion": "Formato Inválido en Excel",
             "Archivos PDF Coincidencia": "N/A",
             "Notas": f"Longitud {item['length']}, no cumple estándar VIN"
         })


    if resultados_para_excel:
        # Crea un DataFrame de pandas y guarda en Excel
        df_resultados = pd.DataFrame(resultados_para_excel)
        output_excel_name = "Reporte_Comparacion_VIN.xlsx"
        # Guarda el Excel en la misma carpeta de origen fija
        output_excel_path = os.path.join(source_folder_path, output_excel_name)
        try:
            df_resultados.to_excel(output_excel_path, index=False)
            print(f"\nReporte detallado guardado en '{output_excel_path}'.")
        except Exception as e:
            print(f"\nError al guardar el archivo Excel de resultados '{output_excel_path}': {e}")
            print(f"Asegúrate de tener permisos de escritura en la carpeta: {source_folder_path}")
    else:
        print("\nNo se generaron resultados para crear un archivo Excel.")


# --- Ejecución Principal con Carpeta de Origen Fija (Busca Excel y PDFs dentro) ---

if __name__ == "__main__":
    print("Iniciando script con carpeta de origen fija...")

    # La ruta de la carpeta principal es la variable fija definida al inicio
    main_folder_path = CARPETA_PRINCIPAL_FIJA

    # --- Validaciones de la ruta fija de la carpeta principal ---
    if not os.path.isdir(main_folder_path):
        print(f"\nError CRÍTICO: La ruta fija especificada para la carpeta principal NO es una carpeta válida o no existe:")
        print(f"'{main_folder_path}'")
        print("Por favor, edita el script y actualiza la variable CARPETA_PRINCIPAL_FIJA con la ruta correcta.")
    else:
        # Busca archivos Excel (.xlsx y .xls) dentro de la carpeta fija
        excel_files_found = glob.glob(os.path.join(main_folder_path, '*.xlsx')) + \
                            glob.glob(os.path.join(main_folder_path, '*.xls'))

        if not excel_files_found:
            print(f"\nError CRÍTICO: No se encontró ningún archivo Excel (.xlsx o .xls) en la carpeta fija:")
            print(f"'{main_folder_path}'")
            print("Por favor, asegúrate de que el archivo Excel (FMM) esté presente en esta carpeta.")
        else:
            # Si se encuentran uno o más archivos Excel, usa el primero (por orden alfabético)
            excel_file_path = excel_files_found[0]
            if len(excel_files_found) > 1:
                 print(f"\nAdvertencia: Se encontraron {len(excel_files_found)} archivos Excel en la carpeta fija. Se utilizará el primero encontrado:")
                 # Opcional: imprimir la lista completa de archivos encontrados
                 # for f in excel_files_found:
                 #      print(f" - {os.path.basename(f)}")
                 print(f"Utilizando: {os.path.basename(excel_file_path)}")

            # La carpeta para buscar PDFs es la misma carpeta principal
            pdf_folder_path = main_folder_path

            # Ejecuta la lógica de comparación con el Excel encontrado y la carpeta principal
            comparar_vins(excel_file_path, pdf_folder_path)

    print("\nScript finalizado.")