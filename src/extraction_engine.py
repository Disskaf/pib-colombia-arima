"""
Módulo de Ingeniería de Datos para el procesamiento, desencriptación,
normalización y empalme en cascada de las series del PIB del DANE (1994-2026).
"""

import os
import io
import re
import glob
import warnings
import unicodedata
from typing import Tuple, List, Optional, Dict
import pandas as pd
import numpy as np
import msoffcrypto

# Desactivación de advertencias de formato para mantener una salida limpia
warnings.filterwarnings("ignore")

# Configuración estática de etiquetas sectoriales normalizadas
ETIQUETAS_SECTORIALES: Dict[str, List[str]] = {
    'construccion': ['construccion'],
    'financiero': [
        'establecimientos financieros, seguros, actividades inmobiliarias y servicios a las empresas',
        'actividades financieras y de seguros',
        'establecimientos financieros'
    ]
}


def normalizar_cadena(texto: any) -> str:
    """
    Normaliza el texto de las celdas mediante la eliminación de acentos,
    mayúsculas, saltos de línea y espacios redundantes.
    """
    if pd.isna(texto):
        return ""
    texto_limpio = str(texto).strip().lower()
    texto_sin_acentos = ''.join(
        c for c in unicodedata.normalize('NFD', texto_limpio)
        if unicodedata.category(c) != 'Mn'
    )
    return re.sub(r'\s+', ' ', texto_sin_acentos)


def desencriptar_archivo_excel(ruta_archivo: str, motor: str) -> Optional[pd.ExcelFile]:
    """
    Abre un archivo Excel. Si se encuentra protegido por el DANE,
    realiza la desencriptación en memoria a través de msoffcrypto.
    """
    try:
        return pd.ExcelFile(ruta_archivo, engine=motor)
    except Exception as error:
        mensaje_error = str(error).lower()
        es_encriptado = any(x in mensaje_error for x in ["encrypted", "encryption", "password"])
        if not es_encriptado:
            raise error

        flujo_desencriptado = io.BytesIO()
        with open(ruta_archivo, "rb") as archivo_fisico:
            archivo_oficina = msoffcrypto.OfficeFile(archivo_fisico)
            try:
                archivo_oficina.load_key(password="VelvetSweatshop")
                archivo_oficina.decrypt(flujo_desencriptado)
            except Exception:
                archivo_fisico.seek(0)
                archivo_oficina.load_key(password="")
                archivo_oficina.decrypt(flujo_desencriptado)
        
        flujo_desencriptado.seek(0)
        return pd.ExcelFile(flujo_desencriptado, engine=motor)


def determinar_hoja_optima(lista_hojas: List[str], nombre_archivo: str) -> str:
    """
    Selecciona la hoja que contiene la matriz real de datos,
    omitiendo índices o menús de navegación.
    """
    nombre_minuscula = nombre_archivo.lower()
    
    # Caso 1: Base 2015 o producción constante moderna
    if "2015" in nombre_minuscula or "produccionconstantes" in nombre_minuscula:
        cuadros_filtrados = [h for h in lista_hojas if "cuadro" in h.lower() and "1" in h]
        if cuadros_filtrados:
            return cuadros_filtrados[0]
            
    # Caso 2: Hojas con palabras clave de ramas de actividad
    ramas_filtradas = [h for h in lista_hojas if any(x in h.lower() for x in ["ramas", "grandes", "abs"])]
    hojas_filtradas = [h for h in r_sheets if not any(x in h.lower() for x in ["var", "anual", "trim", "semest"])] if (r_sheets := ramas_filtradas) else []
    
    if hojas_filtradas:
        return hojas_filtradas[0]
        
    # Caso de descarte: Evitar menús principales
    if len(lista_hojas) > 1 and any(x in lista_hojas[0].lower() for x in ["indice", "índice", "menu", "menú"]):
        return lista_hojas[1]
        
    return lista_hojas[0]


def buscar_coordenadas_sectores(df: pd.DataFrame) -> Tuple[Optional[Tuple[int, int]], Optional[Tuple[int, int]]]:
    """
    Localiza las coordenadas fila-columna de los sectores económicos
    mediante búsquedas exactas y parciales para mitigar la inestabilidad de las plantillas.
    """
    coordenadas_const = []
    coordenadas_fin = []
    filas, columnas = df.shape

    # Búsqueda inicial: Coincidencia exacta
    for r in range(filas):
        for c in range(columnas):
            valor_celda = normalizar_cadena(df.iloc[r, c])
            if valor_celda in ETIQUETAS_SECTORIALES['construccion']:
                coordenadas_const.append((r, c))
            if valor_celda in ETIQUETAS_SECTORIALES['financiero']:
                coordenadas_fin.append((r, c))

    # Cláusula de guarda: Si se obtienen coordenadas exactas, se retornan de inmediato
    if coordenadas_const and coordenadas_fin:
        return coordenadas_const[0], coordenadas_fin[0]

    # Búsqueda secundaria: Coincidencia parcial si falla la exacta
    for r in range(filas):
        for c in range(columnas):
            valor_celda = normalizar_cadena(df.iloc[r, c])
            if not coordenadas_const and any(ext in valor_celda for ext in ETIQUETAS_SECTORIALES['construccion']):
                coordenadas_const.append((r, c))
            if not coordenadas_fin and any(ext in valor_celda for ext in ETIQUETAS_SECTORIALES['financiero']):
                coordenadas_fin.append((r, c))

    coord_const = coordenadas_const[0] if coordenadas_const else None
    coord_fin = coordenadas_fin[0] if coordenadas_fin else None
    return coord_const, coord_fin


def extraer_estructura_horizontal(df: pd.DataFrame, r_const: int, c_const: int, r_fin: int, c_fin: int) -> pd.DataFrame:
    """
    Estrategia de extracción para matrices donde los sectores están en filas
    y los periodos de tiempo se organizan en las columnas.
    """
    fila_trimestres = -1
    for r in range(df.shape[0]):
        valores_fila = df.iloc[r].astype(str).str.strip().str.upper().tolist()
        conteo_trimestres = sum(1 for v in valores_fila if v in ['I', 'II', 'III', 'IV', '1', '2', '3', '4'])
        if conteo_trimestres >= 4:
            fila_trimestres = r
            break

    # Cláusula de guarda si no se localiza la fila temporal
    if fila_trimestres == -1:
        return pd.DataFrame()

    fila_anos = fila_trimestres - 1
    anio_actual = None
    fechas, valores_const, valores_fin = [], [], []
    mapeo_trimestres = {'I': '-03-31', 'II': '-06-30', 'III': '-09-30', 'IV': '-12-31',
                        '1': '-03-31', '2': '-06-30', '3': '-09-30', '4': '-12-31'}

    for col in range(c_const + 1, df.shape[1]):
        valor_anio = str(df.iloc[fila_anos, col]).strip()
        valor_trim = str(df.iloc[fila_trimestres, col]).strip().upper()
        
        busqueda_anio = re.search(r'(19\d{2}|20\d{2})', valor_anio)
        if busqueda_anio:
            anio_actual = busqueda_anio.group(1)
            
        if valor_trim in mapeo_trimestres and anio_actual:
            v_const = pd.to_numeric(df.iloc[r_const, col], errors='coerce')
            v_fin = pd.to_numeric(df.iloc[r_fin, col], errors='coerce')
            
            if pd.notna(v_const) and pd.notna(v_fin):
                fecha_iso = f"{anio_actual}{mapeo_trimestres[valor_trim]}"
                fechas.append(pd.to_datetime(fecha_iso))
                valores_const.append(v_const)
                valores_fin.append(v_fin)

    return pd.DataFrame({'Construccion': valores_const, 'Sector_Financiero': valores_fin}, index=fechas)


def extraer_estructura_vertical(df: pd.DataFrame, r_const: int, c_const: int, r_fin: int, c_fin: int) -> pd.DataFrame:
    """
    Estrategia de extracción para matrices donde los sectores están en columnas
    y los periodos de tiempo se organizan en las filas.
    """
    columna_trimestres = -1
    for col in range(df.shape[1]):
        valores_col = df.iloc[:, col].astype(str).str.strip().str.upper().tolist()
        conteo_trimestres = sum(1 for v in valores_col if v in ['I', 'II', 'III', 'IV', '1', '2', '3', '4'])
        if conteo_trimestres >= 4:
            columna_trimestres = col
            break

    # Cláusula de guarda si no se localiza la columna temporal
    if columna_trimestres == -1:
        return pd.DataFrame()

    columna_anos = columna_trimestres - 1 if columna_trimestres > 0 else 0
    anio_actual = None
    fechas, valores_const, valores_fin = [], [], []
    mapeo_trimestres = {'I': '-03-31', 'II': '-06-30', 'III': '-09-30', 'IV': '-12-31',
                        '1': '-03-31', '2': '-06-30', '3': '-09-30', '4': '-12-31'}

    for row in range(r_const + 1, df.shape[0]):
        valor_anio = str(df.iloc[row, columna_anos]).strip()
        valor_trim = str(df.iloc[row, columna_trimestres]).strip().upper()
        
        busqueda_anio = re.search(r'(19\d{2}|20\d{2})', valor_anio)
        if busqueda_anio:
            anio_actual = busqueda_anio.group(1)
            
        if valor_trim in mapeo_trimestres and anio_actual:
            v_const = pd.to_numeric(df.iloc[row, c_const], errors='coerce')
            v_fin = pd.to_numeric(df.iloc[row, c_fin], errors='coerce')
            
            if pd.notna(v_const) and pd.notna(v_fin):
                fecha_iso = f"{anio_actual}{mapeo_trimestres[valor_trim]}"
                fechas.append(pd.to_datetime(fecha_iso))
                valores_const.append(v_const)
                valores_fin.append(v_fin)

    return pd.DataFrame({'Construccion': valores_const, 'Sector_Financiero': valores_fin}, index=fechas)


def extraer_series_archivo(ruta_archivo: str) -> pd.DataFrame:
    """
    Función orquestadora que integra la carga del archivo,
    identificación de celdas y ejecución del algoritmo de extracción correspondiente.
    """
    nombre_archivo = os.path.basename(ruta_archivo)
    motor = 'openpyxl' if ruta_archivo.endswith('.xlsx') else 'xlrd'
    
    excel_file = desencriptar_archivo_excel(ruta_archivo, motor)
    if not excel_file:
        return pd.DataFrame()
        
    hoja_optima = determinar_hoja_optima(excel_file.sheet_names, nombre_archivo)
    df_crudo = excel_file.parse(hoja_optima, header=None)
    
    coord_const, coord_fin = buscar_coordenadas_sectores(df_crudo)
    if not (coord_const and coord_fin):
        return pd.DataFrame()
        
    r_const, c_const = coord_const
    r_fin, c_fin = coord_fin
    
    # Despacho dinámico de la estrategia de extracción (Horizontal vs Vertical)
    if c_const == c_fin and r_const != r_fin:
        df_final = extraer_estructura_horizontal(df_crudo, r_const, c_const, r_fin, c_fin)
    else:
        df_final = extraer_estructura_vertical(df_crudo, r_const, c_const, r_fin, c_fin)
        
    # Asignación de etiqueta metodológica al dataframe final
    if not df_final.empty:
        nombre_minuscula = nombre_archivo.lower()
        if "2015" in nombre_minuscula or "produccionconstantes" in nombre_minuscula:
            df_final['Base'] = 2015
        elif "1994" in nombre_minuscula:
            df_final['Base'] = 1994
        else:
            df_final['Base'] = 2005
            
    return df_final


def unificar_pib_colombia(ruta_carpeta: str) -> Tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    """
    Ejecuta el pipeline completo: consolidación de archivos y aplicación
    del algoritmo de retropolación sucesiva en cascada para unificar las tres bases metodológicas.
    """
    lista_archivos = glob.glob(os.path.join(ruta_carpeta, "*.xls*"))
    if not lista_archivos:
        raise FileNotFoundError(f"No se detectaron archivos de Excel en la ruta: {ruta_carpeta}")
        
    datos_extraidos = []
    for archivo in lista_archivos:
        df_temp = extraer_series_archivo(archivo)
        if not df_temp.empty:
            datos_extraidos.append(df_temp)
            
    df_consolidado = pd.concat(datos_extraidos)
    df_consolidado = df_consolidado.reset_index().sort_values(by=['Fecha', 'Base'], ascending=[True, False])
    df_consolidado = df_consolidado.drop_duplicates(subset='Fecha', keep='first').set_index('Fecha')
    
    # División por bases originales para fines de validación macroeconómica
    df_base2015 = df_consolidado[df_consolidado['Base'] == 2015][['Construccion', 'Sector_Financiero']].sort_index()
    df_base2005 = df_consolidado[df_consolidado['Base'] == 2005][['Construccion', 'Sector_Financiero']].sort_index()
    df_base1994 = df_consolidado[df_consolidado['Base'] == 1994][['Construccion', 'Sector_Financiero']].sort_index()
    
    # Inicialización de la serie unificada con la base ancla moderna (2015)
    df_pib_unificado = df_base2015.copy()
    
    # --- PASO A: Empalme Base 2005 hacia atrás sobre Base 2015 ---
    tasas_2005 = df_base2005.pct_change() + 1
    fecha_ancla_2015 = df_base2015.index.min()
    fechas_pasado_2005 = df_base2005.index[df_base2005.index < fecha_ancla_2015].sort_values(ascending=False)
    
    for fecha in fechas_pasado_2005:
        fecha_sig_antigua = df_base2005.index[df_base2005.index > fecha].min()
        if pd.notna(fecha_sig_antigua):
            tasa = tasas_2005.loc[fecha_sig_antigua]
            fecha_sig_nueva = df_pib_unificado.index[df_pib_unificado.index > fecha].min()
            if pd.notna(fecha_sig_nueva):
                df_pib_unificado.loc[fecha] = df_pib_unificado.loc[fecha_sig_nueva] / tasa
                
    # --- PASO B: Empalme Base 1994 hacia atrás sobre la serie ya pre-unificada ---
    tasas_1994 = df_base1994.pct_change() + 1
    fecha_ancla_2005 = df_base2005.index.min()
    fechas_pasado_1994 = df_base1994.index[df_base1994.index < fecha_ancla_2005].sort_values(ascending=False)
    
    for fecha in fechas_pasado_1994:
        fecha_sig_antigua = df_base1994.index[df_base1994.index > fecha].min()
        if pd.notna(fecha_sig_antigua):
            tasa = tasas_1994.loc[fecha_sig_antigua]
            fecha_sig_nueva = df_pib_unificado.index[df_pib_unificado.index > fecha].min()
            if pd.notna(fecha_sig_nueva):
                df_pib_unificado.loc[fecha] = df_pib_unificado.loc[fecha_sig_nueva] / tasa
                
    df_pib_unificado = df_pib_unificado.sort_index().asfreq('Q')
    
    return df_pib_unificado, df_base2015, df_base2005, df_base1994
