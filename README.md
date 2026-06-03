# NDX Sales Calculation Tool

Herramienta de escritorio para extraer y analizar datos de cálculo de ventas de torres Nordex desde archivos Excel.

---

## Uso para el Usuario Final (sin instalar nada)

Si tienes el ejecutable compilado (`.exe`), simplemente:

1. Descomprime la carpeta `NDX_Sales_Calculation.zip`
2. Abre la carpeta `NDX_Sales_Calculation\`
3. Haz doble clic en **`NDX_Sales_Calculation.exe`**

No necesitas instalar Python ni ningún programa adicional.

---

## Cómo Usar la Aplicación

### Paso 1 — Seleccionar el archivo fuente
- Haz clic en **"Seleccionar..."**
- Elige el archivo Excel de entrada (`.xlsx`) con los datos de ventas
- La ruta de salida se sugiere automáticamente

### Paso 2 — Años a procesar
- Selecciona los años que deseas extraer (2026, 2027, 2028, 2029)
- Por defecto están seleccionados 2026, 2027 y 2028

### Paso 3 — Guardar resultado en
- Confirma o cambia la ruta donde se guardará el Excel de salida
- Puedes hacer clic en **"Guardar en..."** para elegir otra ubicación

### Extracción
- Haz clic en **"EXTRAER DATOS"**
- El progreso se muestra en la barra y en el registro de actividad

### Acciones posteriores
- **"Verificar Output"** — Valida el archivo generado
- **"Analizar Gaps"** — Muestra filas sin costo por componente y región

---

## Novedades y Reglas de Negocio (Update)

- **Opcional Fire Detection System**: Extraído como un opcional global aplicable.
- **Relleno Inteligente de Costos**: Las celdas en blanco en el Excel ahora son mapeadas automáticamente a `0`. Las celdas con el texto `n.a.` son procesadas como `N/A`. 
- **Limpieza Automática**: Las torres que se queden en blanco (o con un 0 automático) para ambas monedas en un componente dado serán eliminadas del output final.
- **Restricción TCS**: Los componentes opcionales provenientes del `TCS SC` ahora aplican única y exclusivamente a `Europe`, `Poland`, `Germany` y `Turkey`.
- **Evasión de Encabezados Dinámicos**: Se previno el error donde falsas torres llamadas `"Key"` (por la coincidencia de nombres de los encabezados) se añadían a los datos.
- **Opcionales Canadienses**: Se añadieron los opcionales `Service Lift 60 Hz Canada` y `MV Cables Canada/US- CSA/UL Std` para los proyectos canadienses en la región `Canada`.
- **Expansión Regional de USA**: Los componentes opcionales marcados con región `US` ahora se duplican automáticamente también para las plantas `Arcosa US` y `CS Wind US`.
- **Costes Únicos y Límite de Bloques**: Se ajustó la lógica en TCS SC para ignorar bloques duplicados de referencia o de diferencias (ej. v26.1), garantizando que se extrae un único coste (v26.2). Además, se ampliaron los límites de lectura en TS SC para no dejar atrás datos como los de CS Wind US 2028.
- **Exclusividad para Alemania**: La opción global `Option Anchor Cage: German requirements for NAT AC` está ahora restringida de manera exclusiva para la región `Germany`.

---

## Alternativa: Ejecutar con Python (para usuarios avanzados)

Si tienes Python 3.10+ instalado:

1. Abre una terminal en esta carpeta
2. Ejecuta `setup.bat` para crear el entorno virtual e instalar dependencias
3. Luego haz doble clic en `lanzar_app.bat` para iniciar la app

---

## Estructura del Proyecto

```
ndx_app_sales_calc/
├── src/                ← Código fuente
│   ├── main.py         ← Aplicación principal (interfaz gráfica)
│   └── core/           ← Módulos de lógica y procesamiento
│       ├── extractor.py    ← Extracción de datos desde Excel
│       ├── verifier.py     ← Verificación del output
│       ├── gap_analyzer.py ← Análisis de gaps (filas sin costo)
│       └── paths.py        ← Utilidades de rutas (compatible con .exe)
├── config.json         ← Configuración de hojas y regiones
├── input_source/       ← Coloca aquí el archivo Excel de entrada
├── output/             ← Los resultados se guardan aquí por defecto
├── build/              ← Infraestructura para compilar el .exe
│   ├── ndx_sales.spec  ← Configuración de PyInstaller
│   └── build_exe.bat   ← Script para compilar el ejecutable
├── lanzar_app.bat      ← Lanzador rápido (requiere Python)
├── setup.bat           ← Configura entorno virtual (requiere Python)
└── requirements.txt    ← Dependencias Python
```

---

## Para el Desarrollador: Compilar el Ejecutable

Requisitos previos: Python 3.10+, conexión a internet

```bat
# 1. Crear entorno virtual e instalar dependencias
setup.bat

# 2. Instalar PyInstaller (herramienta de build)
.venv\Scripts\pip install pyinstaller>=6.0

# 3. Compilar el ejecutable
build\build_exe.bat
```

El resultado estará en `dist\NDX_Sales_Calculation\`. 
Comprime esa carpeta en `.zip` y entrega al usuario final.
