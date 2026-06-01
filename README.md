# Taller Final: Programación aplicada a la economía con R y Python

## Informe Investigativo: Modelado Econométrico Y Complementariedad Metodológica En El Empalme De Series PIB Del DANE (1994-2026)

Institución: Fundación Universitaria del Área Andina
Materia: Inteligencia Artificial para la solución de problemas / Programación Aplicada a la Economía
Profesor: Luz Andrea Sánchez Buitrago
Ciudad: Bogotá D.C., Colombia
Fecha: 31 de mayo de 2026

### Integrantes:
1. Andrés Fabián Sepúlveda Mariño
2. Angie Camila Velásquez
3. Stefania Colorado Tico
4. Jhony Stevan Cárdenas Rodríguez

## 📌 Descripción del Proyecto

Este repositorio contiene la arquitectura de software y el análisis econométrico para la unificación histórica de las series de Producto Interno Bruto (PIB) en Colombia durante el periodo 1994-2026. A través de un enfoque que integra la ingeniería de datos y el modelado estadístico univariado, el proyecto automatiza el procesamiento de los reportes sectoriales de Construcción y Establecimientos Financieros.
La propuesta metodológica clave de este trabajo se centra en el diseño de un algoritmo de retropolación sucesiva en cascada (backcasting), el cual unifica de forma homogénea las bases macroeconómicas de 1994, 2005 y 2015 del DANE; este proceso elimina las distorsiones de escala nominales sin alterar la volatilidad intrínseca de los ciclos económicos de Colombia, lo que permite estimar pronósticos con un menor margen de error.

## 🛠️ Estructura del Repositorio

El diseño de la solución sigue prácticas de desarrollo limpio y modular; el motor de procesamiento se encuentra separado de la interfaz de visualización estadística:

```
pib-colombia-arima/             <-- Directorio raíz del repositorio en GitHub
│
├── data/                       # Carpeta que aloja los 65 archivos XLS/XLSX oficiales del DANE
│
├── src/                        # Paquete modular en Python (Lógica del backend)
│   ├── __init__.py             # Inicialización y exportación pública del paquete
│   └── extraction_engine.py    # Motor ETL (desencriptación en memoria, normalización y empalme)
│
├── pib_colombia_colab.ipynb    # Cuaderno analítico principal (Interfaz frontend para Google Colab)
└── README.md                   # Documentación técnica (este archivo)
```

## 🚀 Guía de Ejecución en Google Colab

El cuaderno `pib_colombia_colab.ipynb` está diseñado para ejecutarse sin requerir cargas manuales de datos o configuraciones locales complejas; la infraestructura virtual se inicializa en dos pasos secuenciales:

## 📈 Fases del Pipeline Analítico

Una vez que el motor de extracción entrega las series temporales homogéneas, el cuaderno pib_colombia_colab.ipynb ejecuta las siguientes etapas econométricas de forma consecutiva:

1. Diagnóstico de Continuidad (Celda 3): Grafica las series unificadas frente a las bases originales utilizando escalas de doble eje Y; esto permite verificar visualmente la consistencia física del empalme del año 2000.
2. Evaluación de Estacionariedad (Celdas 4 y 5): Aplica la prueba aumentada de Dickey-Fuller (ADF) en los niveles de cada sector; justifica de manera formal la necesidad de aplicar diferencias e inspecciona los correlogramas de autocorrelación simple (ACF) y parcial (PACF).
3. Modelado y Proyección ARIMA (Celda 6): Ajusta de manera automática las estructuras óptimas mediante criterios de información de Akaike (AIC) y genera proyecciones puntuales a diez trimestres con intervalos de confianza del 95%.
