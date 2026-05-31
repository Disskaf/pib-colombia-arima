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

# Diccionario de etiquetas normalizadas (sin acentos, en minúsculas y limpias)
etiquetas_normalizadas: Dict[str, List[str]] = {
    'construccion': ['construccion'],
    'financiero': [
        'establecimientos financieros, seguros, actividades inmobiliarias y servicios a las empresas',
        'actividades financieras y de seguros',
        'establecimientos financieros'
    ]
}


def normalizar_texto(texto: any) -> str:
    """
    Normaliza el texto de las celdas mediante la eliminación de acentos,
    mayúsculas, saltos de línea y espacios redundantes para garantizar coincidencias exactas.
    """
    if pd.isna(texto):
        return ""
    texto_limpio = str(texto).strip().lower()
    texto_sin_acentos = ''.join(
        c for c in unicodedata.normalize('NFD', texto_limpio)
        if unicodedata.category(c) != 'Mn'
    )
    return re.sub(r'\s+', ' ', texto_sin_acentos)


def desencriptar_y_cargar_excel(ruta_archivo: str, motor: str) -> Optional[pd.ExcelFile]:
    """
    Intenta abrir el archivo Excel. Si está encriptado con la protección estándar
    del DANE (VelvetSweatshop), lo desencripta en memoria usando msoffcrypto.
    """
    try:
        return pd.ExcelFile(ruta_archivo, engine=motor)
    except Exception as e:
        mensaje_error = str(e).lower()
        es_encriptado = any(x in mensaje_error for x in ["encrypted", "encryption", "password"])
        if not es_encriptado:
            raise e

        decrypted_stream = io.BytesIO()
        with open(ruta_archivo, "rb") as f:
            office_file = msoffcrypto.OfficeFile(f)
            try:
                office_file.load_key(password="VelvetSweatshop")
                office_file.decrypt(decrypted_stream)
            except Exception:
                f.seek(0)
                office_file.load_key(password="")
                office_file.decrypt(decrypted_stream)
        
        decrypted_stream.seek(0)
        return pd.ExcelFile(decrypted_stream, engine=motor)


def extraer_series_archivo(ruta_archivo: str) -> pd.DataFrame:
    """
    Identifica la estructura del archivo, gestiona la desencriptación,
    busca la hoja con datos reales (omitiendo índices) y extrae las series.
    """
    nombre_archivo = os.path.basename(ruta_archivo).lower()
    df_temp = pd.DataFrame()

    try:
        motor = 'openpyxl' if ruta_archivo.endswith('.xlsx') else 'xlrd'
        excel_file = desencriptar_y_cargar_excel(ruta_archivo, motor)
        hojas = excel_file.sheet_names

        # Selección de hoja óptima
        hoja_seleccionada = hojas[0]
        if "2015" in nombre_archivo or "produccionconstantes" in nombre_archivo:
            cuadros_1 = [h for h in hojas if "cuadro" in h.lower() and "1" in h]
            if cuadros_1:
                hoja_seleccionada = cuadros_1[0]
        else:
            ramas_sheets = [h for h in hojas if any(x in h.lower() for x in ["ramas", "grandes", "abs"])]
            
            # Corrección del operador morsa para evitar errores de compilación y alcance
            r_sheets = [
                h for h in ramas_sheets 
                if not any(x in h.lower() for x in ["var", "anual", "trim", "semest"])
            ]
            if r_sheets:
                hoja_seleccionada = r_sheets[0]
            elif len(hojas) > 1:
                if any(x in hojas[0].lower() for x in ["indice", "índice", "menu", "menú"]):
                    hoja_seleccionada = hojas[1]

        raw_df = excel_file.parse(hoja_seleccionada, header=None)

        # Localización espacial de las variables
        coords_const = []
        coords_fin = []

        # Coincidencia exacta
        for r in range(raw_df.shape[0]):
            for c in range(raw_df.shape[1]):
                val_norm = normalizar_texto(raw_df.iloc[r, c])
                if val_norm in etiquetas_normalizadas['construccion']:
                    coords_const.append((r, c))
                if val_norm in etiquetas_normalizadas['financiero']:
                    coords_fin.append((r, c))

        # Coincidencia parcial si falla la exacta
        if not coords_const or not coords_fin:
            for r in range(raw_df.shape[0]):
                for c in range(raw_df.shape[1]):
                    val_norm = normalizar_texto(raw_df.iloc[r, c])
                    if not coords_const and any(ext in val_norm for ext in etiquetas_normalizadas['construccion']):
                        coords_const.append((r, c))
                    if not coords_fin and any(ext in val_norm for ext in etiquetas_normalizadas['financiero']):
                        coords_fin.append((r, c))

        if coords_const and coords_fin:
            r_const, c_const = coords_const[0]
            r_fin, c_fin = coords_fin[0]

            fechas_list = []
            vals_const = []
            vals_fin = []

            # CASO A: Estructura Horizontal (Sectores en filas, periodos en columnas)
            if c_const == c_fin and r_const != r_fin:
                row_trimestres = -1
                for r in range(raw_df.shape[0]):
                    row_vals = raw_df.iloc[r].astype(str).str.strip().str.upper().tolist()
                    quarter_count = sum(1 for v in row_vals if v in ['I', 'II', 'III', 'IV', '1', '2', '3', '4'])
                    if quarter_count >= 4:
                        row_trimestres = r
                        break

                if row_trimestres != -1:
                    row_anos = row_trimestres - 1
                    current_year = None

                    for col in range(c_const + 1, raw_df.shape[1]):
                        val_ano = str(raw_df.iloc[row_anos, col]).strip()
                        val_tri = str(raw_df.iloc[row_trimestres, col]).strip().upper()

                        match_year = re.search(r'(19\d{2}|20\d{2})', val_ano)
                        if match_year:
                            current_year = match_year.group(1)

                        if val_tri in ['I', 'II', 'III', 'IV', '1', '2', '3', '4'] and current_year is not None:
                            v_const = pd.to_numeric(raw_df.iloc[r_const, col], errors='coerce')
                            v_fin = pd.to_numeric(raw_df.iloc[r_fin, col], errors='coerce')

                            if pd.notna(v_const) and pd.notna(v_fin):
                                map_q = {'I': '-03-31', 'II': '-06-30', 'III': '-09-30', 'IV': '-12-31',
                                         '1': '-03-31', '2': '-06-30', '3': '-09-30', '4': '-12-31'}
                                fecha_str = f"{current_year}{map_q[val_tri]}"
                                fechas_list.append(pd.to_datetime(fecha_str))
                                vals_const.append(v_const)
                                vals_fin.append(v_fin)

            # CASO B: Estructura Vertical (Sectores en columnas, periodos en filas)
            elif r_const == r_fin and c_const != c_fin:
                col_trimestres = -1
                for col in range(raw_df.shape[1]):
                    col_vals = raw_df.iloc[:, col].astype(str).str.strip().str.upper().tolist()
                    quarter_count = sum(1 for v in col_vals if v in ['I', 'II', 'III', 'IV', '1', '2', '3', '4'])
                    if quarter_count >= 4:
                        col_trimestres = col
                        break

                if col_trimestres != -1:
                    col_anos = col_trimestres - 1 if col_trimestres > 0 else 0
                    current_year = None

                    for row in range(r_const + 1, raw_df.shape[0]):
                        val_ano = str(raw_df.iloc[row, col_anos]).strip()
                        val_tri = str(raw_df.iloc[row, col_trimestres]).strip().upper()

                        match_year = re.search(r'(19\d{2}|20\d{2})', val_ano)
                        if match_year:
                            current_year = match_year.group(1)

                        if val_tri in ['I', 'II', 'III', 'IV', '1', '2', '3', '4'] and current_year is not None:
                            v_const = pd.to_numeric(raw_df.iloc[row, c_const], errors='coerce')
                            v_fin = pd.to_numeric(raw_df.iloc[row, c_fin], errors='coerce')

                            if pd.notna(v_const) and pd.notna(v_fin):
                                map_q = {'I': '-03-31', 'II': '-06-30', 'III': '-09-30', 'IV': '-12-31',
                                         '1': '-03-31', '2': '-06-30', '3': '-09-30', '4': '-12-31'}
                                fecha_str = f"{current_year}{map_q[val_tri]}"
                                fechas_list.append(pd.to_datetime(fecha_str))
                                vals_const.append(v_const)
                                vals_fin.append(v_fin)

            if fechas_list:
                df_temp = pd.DataFrame({
                    'Construccion': vals_const,
                    'Sector_Financiero': vals_fin
                }, index=fechas_list)
                df_temp.index.name = 'Fecha'

                # Clasificar base metodológica explícitamente en tres categorías
                if "2015" in nombre_archivo or "produccionconstantes" in nombre_archivo:
                    df_temp['Base'] = 2015
                elif "1994" in nombre_archivo:
                    df_temp['Base'] = 1994
                else:
                    df_temp['Base'] = 2005

    except Exception as e:
        print(f"⚠️ Error procesando {nombre_archivo}: {str(e)}")

    return df_temp


def unificar_pib_colombia(ruta_carpeta: str) -> Tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    """
    Ejecuta el pipeline de consolidación y el algoritmo de retropolación sucesiva
    con base en la estructura de datos original.
    """
    lista_archivos = glob.glob(os.path.join(ruta_carpeta, "*.xls*"))
    if not lista_archivos:
        raise FileNotFoundError(f"No se detectaron archivos de Excel en la ruta: {ruta_carpeta}")
        
    datos_consolidados = []
    for archivo in lista_archivos:
        datos_parciales = extraer_series_archivo(archivo)
        if not datos_parciales.empty:
            datos_consolidados.append(datos_parciales)

    if not datos_consolidados:
        raise ValueError("No se extrajo información válida. Verifique las rutas o las etiquetas.")

    # Consolidación de matrices
    df_total = pd.concat(datos_consolidados)

    df_total = df_total.reset_index().sort_values(by=['Fecha', 'Base'], ascending=[True, False])
    df_total = df_total.drop_duplicates(subset='Fecha', keep='first').set_index('Fecha')

    # Separación por bases históricas
    df_base2015 = df_total[df_total['Base'] == 2015][['Construccion', 'Sector_Financiero']].sort_index()
    df_base2005 = df_total[df_total['Base'] == 2005][['Construccion', 'Sector_Financiero']].sort_index()
    df_base1994 = df_total[df_total['Base'] == 1994][['Construccion', 'Sector_Financiero']].sort_index()

    # Iniciamos con la base de nivel ancla (2015)
    df_pib = df_base2015.copy()

    # --- PASO A: Empalmar Base 2005 hacia atrás sobre Base 2015 ---
    tasas_2005 = df_base2005.pct_change() + 1
    fecha_ancla_2015 = df_base2015.index.min()
    fechas_pasado_2005 = df_base2005.index[df_base2005.index < fecha_ancla_2015].sort_values(ascending=False)

    for fecha in fechas_pasado_2005:
        fecha_siguiente_antigua = df_base2005.index[df_base2005.index > fecha].min()
        if pd.notna(fecha_siguiente_antigua):
            tasa = tasas_2005.loc[fecha_siguiente_antigua]
            fecha_siguiente_nueva = df_pib.index[df_pib.index > fecha].min()
            if pd.notna(fecha_siguiente_nueva):
                df_pib.loc[fecha] = df_pib.loc[fecha_siguiente_nueva] / tasa

    # --- PASO B: Empalmar Base 1994 hacia atrás sobre la serie ya unificada (2005 empalmada) ---
    tasas_1994 = df_base1994.pct_change() + 1
    fecha_ancla_2005 = df_base2005.index.min()
    fechas_pasado_1994 = df_base1994.index[df_base1994.index < fecha_ancla_2005].sort_values(ascending=False)

    for fecha in fechas_pasado_1994:
        fecha_siguiente_antigua = df_base1994.index[df_base1994.index > fecha].min()
        if pd.notna(fecha_siguiente_antigua):
            tasa = tasas_1994.loc[fecha_siguiente_antigua]
            fecha_siguiente_nueva = df_pib.index[df_pib.index > fecha].min()
            if pd.notna(fecha_siguiente_nueva):
                df_pib.loc[fecha] = df_pib.loc[fecha_siguiente_nueva] / tasa

    df_pib = df_pib.sort_index().asfreq('Q')
    
    return df_pib, df_base2015, df_base2005, df_base1994
```

eof

---

### Pasos definitivos para actualizar tu entorno en Google Colab:

Sigue este protocolo riguroso en tu consola y en Colab para purgar las carpetas antiguas y forzar a Python a cargar la versión unificada sin operador morsa:

#### 1. Sube el código limpio a GitHub:
Reemplaza el contenido de tu archivo local `src/extraction_engine.py` con el código del bloque superior de arriba. Luego, ejecuta en tu terminal para sincronizarlo con la nube:
```bash
git add src/extraction_engine.py
git commit -m "fix: restore legacy extraction engine logic without walrus scope issues"
git push origin main
```

#### 2. Reemplaza por completo la Celda 1 de tu Google Colab:
Para asegurarte de que Colab no conserve ningún archivo temporal huérfano en el disco virtual de la máquina, utiliza esta versión forzada de purga de directorios:
```python
# ==============================================================================
# CELDA 1: Configuración de entorno y limpieza absoluta de caché
# ==============================================================================
import os
import shutil
import sys

# Definir la ruta del proyecto clonado
repo_name = 'pib-colombia-arima'

# 1. Purgar físicamente directorios previos de Colab para romper la caché de disco
for path in [repo_name, 'datos_pib', 'src']:
    if os.path.exists(path):
        shutil.rmtree(path)
        print(f"🧹 Carpeta residual '{path}' eliminada con éxito.")

# 2. Eliminar del registro de Python cualquier versión en caché de tus módulos
for mod in list(sys.modules.keys()):
    if "src" in mod or "extraction_engine" in mod:
        sys.modules.pop(mod, None)
print("🧹 Memoria intermedia de importación de Python purgada.")

# 3. Clonar la última versión limpia de tu repositorio
# REEMPLAZA "tu-usuario" con tu nombre de usuario de GitHub
!git clone https://github.com/tu-usuario/pib-colombia-arima.git

# 4. Cambiar de directorio y configurar el path del sistema
%cd pib-colombia-arima
sys.path.insert(0, os.path.abspath('.'))

# 5. Instalar dependencias limpias
!pip install pmdarima openpyxl xlrd==2.0.1 msoffcrypto-tool
```

#### 3. Corre tu Celda 2 importando la biblioteca:
Ahora puedes llamar al motor en tu celda de Colab de forma elegante:
```python
# ==============================================================================
# CELDA 2: Ejecución del pipeline ETL modularizado
# ==============================================================================
from src import unificar_pib_colombia

# Se pasa la ruta de la carpeta de datos dentro del repositorio clonado
df_pib, df_b2015, df_b2005, df_b1994 = unificar_pib_colombia('./data')

print("\nEstructura final de los datos reales empalmados (Primeros 5 registros):")
display(df_pib.head())
