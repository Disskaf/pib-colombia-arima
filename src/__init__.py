"""
Paquete src: Motor de extracción, unificación y modelado del PIB en Colombia.
Expone la interfaz pública del pipeline para el cuaderno de Google Colab.
"""

# Importación relativa para exponer la función principal en el nivel del paquete
from .extraction_engine import procesar_y_empalmar

# Definición de exportaciones explícitas para evitar fugas de namespace
__all__ = ["procesar_y_empalmar]

# Metadatos del paquete
__author__ = "Fabián"
__version__ = "1.0.0"
