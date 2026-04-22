# Sales Calculation - Índice de Archivos

## Archivos Principales

### Documentación
- **QUICKSTART.md** — Guía rápida para usuarios (empezar aquí)
- **README.md** — Documentación completa del proyecto
- **SCRIPTS_GUIDE.md** — Explicación detallada de los 2 scripts
- **INDEX.md** — Este archivo

### Configuración
- **config.json** — Archivo de configuración (rutas, parámetros, años, regiones)
- **requirements.txt** — Dependencias Python (openpyxl, pandas)
- **.gitignore** — Archivos ignorados en Git (venv, outputs, etc.)

### Instalación
- **setup.bat** — Script automático de instalación (Windows)
- **setup.sh** — Script automático de instalación (Linux/Mac)

## Carpetas

### `scripts/`
Contiene los scripts Python:

- **extract_tower_data.py** — Script principal de extracción
  - Lee config.json
  - Abre archivo Excel fuente
  - Extrae de ambas hojas (TS SC v26.1 + TCS_MB SC v26.1)
  - Normaliza datos
  - Genera output en `output/tower_components_extracted.xlsx`
  - Tiempo: ~45 segundos

- **verify_output.py** — Script de verificación (opcional)
  - Valida que valores en output coincidan con source
  - Verifica muestra de torres

### `output/`
Directorio de salida:

- **tower_components_extracted.xlsx** — Archivo Excel generado
  - 133,588 filas de datos
  - Columnas: Component_Category, Component, Key, Brand, Year_Production, Region, Cost_EUR, Cost_USD
  - Datos consolidados de ambas hojas

### `input_source/`
Directorio para archivos de entrada (opcional):

- Carpeta vacía para organización
- Si copias el Excel fuente aquí, actualiza `config.json` para apuntar a esta carpeta
- Útil para no perder archivos

### `.venv/`
Entorno virtual de Python (creado automáticamente por setup.bat/setup.sh):

- Contiene intérprete Python aislado
- Contiene dependencias instaladas
- Se crea en: `.venv/`

## Flujo de Uso

```
1. Usuario descarga proyecto
2. Ejecuta setup.bat (Windows) o setup.sh (Linux/Mac)
3. Edita config.json si es necesario (cambiar ruta del Excel)
4. Ejecuta: python scripts/extract_tower_data.py
5. Obtiene resultado en: output/tower_components_extracted.xlsx
```

## Estructura Típica del Árbol

```
Sales_Calculation/
│
├── QUICKSTART.md              ← Empieza aquí
├── README.md                  ← Documentación completa
├── INDEX.md                   ← Este archivo
│
├── config.json                ← Edita para cambiar ruta/parámetros
├── requirements.txt
├── .gitignore
│
├── setup.bat                  ← Ejecutar primera vez (Windows)
├── setup.sh                   ← Ejecutar primera vez (Linux/Mac)
│
├── .venv/                     ← Entorno virtual (auto-creado)
│   ├── Scripts/
│   ├── Lib/
│   └── ...
│
├── scripts/                   ← Scripts Python
│   ├── extract_tower_data.py  ← Script principal
│   └── verify_output.py       ← Verificación
│
├── output/                    ← Resultado
│   └── tower_components_extracted.xlsx
│
└── input_source/              ← Para guardar archivos (opcional)
    └── .gitkeep
```

## Parámetros en config.json

```json
{
  "input_file_path": "C:\\ruta\\al\\archivo.xlsx",  // Ruta fuente
  "output_file_path": "output/tower_components_extracted.xlsx",  // Ruta salida
  "source_sheets": {
    "ts_sc": "TS SC v26.1",      // Nombre hoja 1
    "tcs_mb": "TCS_MB SC v26.1"   // Nombre hoja 2
  },
  "target_years": [2026, 2027, 2028],  // Años a procesar
  "canonical_regions": [
    "Europe", "Germany", "Turkey", ..., "US"  // Regiones (11 total)
  ]
}
```

## Componentes Extraídos

### De TS SC v26.1 (torres de acero, ~400):
- Tower Shell
- Tower Internals
- Tower Bolts Set
- Anchor cage
- Steel Tower Quality Inspectors
- Option Coating c4/c5
- Option Anchor Cage: German requirements for NAT AC
- Option hybrid tower: MB Monthly Cost Indexation
- Option hybrid tower: white coating of concrete part
- Option hybrid tower: no red stripe on concrete part

### De TCS_MB SC v26.1 (torres de concreto, ~108):
- Tower Shell
- Tower Internals
- Tower Bolts Set
- Foundations
- Concrete Tower Keystones + Internals
- Concrete Tower Logistics
- Concrete Tower C&I
- Option Coating c4/c5
- Steel Tower Quality Inspectors
- Option hybrid tower: white coating of concrete part
- Option hybrid tower: no red stripe on concrete part

## Monedas

- **Cost_EUR** (Currency1) — Euros (siempre presente)
- **Cost_USD** (Currency2) — Dólares Americanos (solo algunas regiones)

## Regiones Canonicales (11 total)

1. Europe
2. Germany
3. Turkey
4. Turkey_DOM
5. Asia
6. China
7. Poland
8. Italy
9. Greece
10. CAN (Canadá)
11. US (USA)

## Comandos Útiles

```bash
# Activar entorno (después del setup)
# Windows:
.venv\Scripts\activate

# Linux/Mac:
source .venv/bin/activate

# Ejecutar extracción
python scripts/extract_tower_data.py

# Verificar valores (opcional)
python scripts/verify_output.py

# Desactivar entorno
deactivate
```

## Versión & Fecha

- **Creado**: Abril 2026
- **Versión**: 1.0
- **Autor**: Nordex Sales Data Pipeline
