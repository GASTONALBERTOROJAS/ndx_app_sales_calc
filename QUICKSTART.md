# Guía Rápida - Sales Calculation

## Inicio Rápido (5 minutos)

### Paso 1: Descargar el proyecto

```bash
cd Sales_Calculation
```

### Paso 2: Configurar (primera vez solamente)

#### Windows:
```bash
setup.bat
```

#### Linux/Mac:
```bash
bash setup.sh
```

Esto crea el entorno virtual (`.venv/`) e instala las dependencias automáticamente.

### Paso 3: Ejecutar la extracción

```bash
# Activar entorno (si no está activado)
# Windows:
.venv\Scripts\activate

# Linux/Mac:
source .venv/bin/activate

# Ejecutar ESTE script (obligatorio):
python scripts/extract_tower_data.py
```

**Nota:** Solo necesitas ejecutar `extract_tower_data.py`. El otro script (`verify_output.py`) es opcional para QA. Ver `SCRIPTS_GUIDE.md` para más detalles.

### Resultado

El archivo generado está en:
```
output/SalesCalc_DDMMYY.xlsx
```

Ejemplo: `SalesCalc_220426.xlsx` (22 de abril de 2026)

---

## Obtener el archivo de entrada

El archivo Excel se descarga desde SharePoint. Ver **`input_source/SOURCE_FILES.md`** para:

- Link a SharePoint
- Ruta en el documento
- Instrucciones de descarga

Una vez descargado, edita `config.json`:

```json
{
  "input_file_path": "C:\\ruta\\donde\\descargaste\\archivo.xlsx",
  ...
}
```

Luego ejecuta el script.

---

## Estructura de carpetas

```
Sales_Calculation/
├── scripts/                    # Scripts Python
├── output/                     # Excel generado aquí
├── input_source/              # (opcional) Archivos de entrada
├── config.json                # Configuración
├── requirements.txt           # Dependencias
├── .gitignore                 # Archivos ignorados en Git
└── README.md                  # Documentación completa
```

---

## Troubleshooting

### "No such file or directory"
Verifica que la ruta en `config.json` sea correcta.

### "ModuleNotFoundError: No module named 'openpyxl'"
Ejecuta: `pip install -r requirements.txt`

### "FileNotFoundError: config.json"
Asegúrate de ejecutar desde la carpeta raíz del proyecto (donde está `config.json`).

---

## Para más detalles

Lee `README.md` para documentación completa sobre qué hace el proyecto y cómo personalizarlo.
