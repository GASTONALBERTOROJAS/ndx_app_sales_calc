# Sales Calculation — Tower Component Cost Extractor

## Descripción

Este proyecto extrae y normaliza datos de costos de componentes de torres eólicas desde archivos Excel de cálculo de ventas. Combina información de dos hojas diferentes:

1. **TS SC v26.1**: Desglose de componentes por región y año (estructura matricial con datos anidados)
2. **TCS_MB SC v26.1**: Tabla plana de costos para torres de concreto/híbridas

El resultado es un **archivo Excel unificado** con todos los costos normalizados, regiones canonicalizadas, y monedas separadas (EUR y USD).

---

## Estructura del Proyecto

```
Sales_Calculation/
├── scripts/                          # Scripts de extracción y verificación
│   ├── extract_tower_data.py        # Script principal de extracción
│   └── verify_output.py             # Script de verificación de valores
├── output/                           # Archivo de salida generado
│   └── tower_components_extracted.xlsx
├── input_source/                     # Carpeta para archivos de entrada (si se desea)
├── config.json                       # Archivo de configuración (rutas, parámetros)
├── requirements.txt                  # Dependencias Python
├── .venv/                            # Entorno virtual (después de instalar)
└── README.md                         # Este archivo
```

---

## Configuración

### Obtener el archivo de entrada

Ver **`input_source/SOURCE_FILES.md`** para:
- Link a SharePoint donde está el archivo
- Ruta en el sitio: `Documents > General > Transfer > Patxi > Sales Calc > 26.1`
- Nombre del archivo exacto
- Instrucciones de descarga

### `config.json`

Archivo central de configuración que define:

- **`input_file_path`**: Ruta al archivo Excel fuente (CAMBIAR después de descargar)
- **`output_file_path`**: Ruta relativa del archivo de salida
- **`source_sheets`**: Nombres de las dos hojas a extraer
- **`target_years`**: Años a procesar (ej: 2026, 2027, 2028)
- **`canonical_regions`**: Lista de 11 regiones estándares

**Editar este archivo para cambiar la ruta del archivo de entrada después de descargarlo.**

---

## Instalación

### 1. Crear y activar entorno virtual

```bash
# Crear .venv
python -m venv .venv

# Activar (Windows)
.venv\Scripts\activate

# Activar (Linux/Mac)
source .venv/bin/activate
```

### 2. Instalar dependencias

```bash
pip install -r requirements.txt
```

---

## Uso

### Ejecutar extracción

```bash
python scripts/extract_tower_data.py
```

El script:
1. Lee la configuración de `config.json`
2. Abre el archivo Excel especificado
3. Extrae datos de ambas hojas
4. Normaliza nombres de componentes, regiones y monedas
5. Genera un archivo Excel con todos los datos unificados

**Salida**: `output/SalesCalc_DDMMYY.xlsx`

Ejemplo: `SalesCalc_220426.xlsx` para la ejecución del 22 de abril de 2026

### Verificar valores (opcional)

```bash
python scripts/verify_output.py
```

Valida que los valores en el Excel de salida coincidan con los del archivo fuente (verifica una muestra de torres).

---

## Datos Extraídos

### Hojas de origen y componentes

#### TS SC v26.1 (400 torres)
- **Torre Shell** (tower shell)
- **Tower Internals** (tower internals)
- **Tower Bolts Set** (Tower bolts set)
- **Anchor cage**
- **Option Coating c4/c5**
- **Steel Tower Quality Inspectors**
- **Option hybrid tower: white coating of concrete part**
- **Option hybrid tower: no red stripe on concrete part**
- **Option hybrid tower: MB Monthly Cost Indexation**
- **Option Anchor Cage: German requirements for NAT AC**

#### TCS_MB SC v26.1 (108 torres)
- **Tower Shell**
- **Tower Internals**
- **Tower Bolts Set**
- **Foundations**
- **Concrete Tower Keystones + Internals**
- **Concrete Tower Logistics**
- **Concrete Tower C&I**
- **Option Coating c4/c5**
- **Steel Tower Quality Inspectors**
- **Option hybrid tower: white coating of concrete part**
- **Option hybrid tower: no red stripe on concrete part**

### Columnas de salida

| Columna | Descripción |
|---|---|
| `Component_Category` | Categoría fija: "Tower" |
| `Component` | Nombre del componente |
| `Key` | Identificador único de la torre |
| `Brand` | Marca (Nx = Nordex, o valor extraído) |
| `Year_Production` | Año de producción (2026, 2027, 2028) |
| `Region` | Región canonicalizada (11 opciones) |
| `Cost_EUR` | Costo en Euros |
| `Cost_USD` | Costo en Dólares Americanos |

### Regiones canonicales

1. Europe
2. Germany
3. Turkey
4. Turkey_DOM (Turquía doméstica)
5. Asia
6. China
7. Poland
8. Italy
9. Greece
10. CAN (Canadá)
11. US (Estados Unidos)

**Nota**: Variantes como "Arcosa US", "Asia EUR", "China USD" se normalizan automáticamente.

---

## Lógica de Extracción

### TS SC v26.1

1. **Lee filas 1-3 y 7** para identificar componentes, años y regiones
2. **Busca patrones** en fila 3: `ComponentNameYearCost_CurrencyXRegion`
3. **Extrae valores** para cada torre (filas 8-407) del archivo
4. **Replica componentes globales** (opcionales sin región) para todas las 11 regiones

### TCS_MB SC v26.1

1. **Lee tabla plana** a partir de fila 85 (encabezado) y 86+ (datos)
2. **Unpivot**: Cada columna de componente genera una fila separada
3. **Mantiene**: Key (columna B), Sourcing/Region (C), Year (D), valores (E-P)

### Combinación

1. **Concatena** los datasets de ambas hojas
2. **Normaliza nombres**: `tower shell` → `Tower Shell`
3. **Normaliza años**: Los convierte todos a `int`
4. **Pivota monedas**: Currency1 → Cost_EUR, Currency2 → Cost_USD

---

## Resultados

- **Total de filas**: 133,588
- **Filas TS SC v26.1**: ~132,400
- **Filas TCS_MB SC v26.1**: 1,188
- **Años**: 2026, 2027, 2028
- **Regiones**: 11 canonicales
- **Sin blancos**: Todos los campos obligatorios están poblados

---

## Personalización

### Cambiar archivo de entrada

Edita `config.json` y actualiza `input_file_path`:

```json
{
  "input_file_path": "C:\\ruta\\nuevo\\archivo.xlsx",
  ...
}
```

### Agregar nuevos años

Modifica `target_years` en `config.json`:

```json
{
  "target_years": [2026, 2027, 2028, 2029],
  ...
}
```

### Agregar nuevas regiones

Extiende `canonical_regions`:

```json
{
  "canonical_regions": [
    "Europe", "Germany", ..., "NewRegion"
  ]
}
```

---

## Notas Técnicas

### Dependencias

- **openpyxl**: Lectura de archivos Excel
- **pandas**: Manipulación de dataframes y pivoting

### Manejo de monedas

- **Cost_Currency1 (EUR)**: Siempre presente en ambas hojas
- **Cost_Currency2 (USD)**: Solo en algunas regiones/componentes de TS SC v26.1
- **Valores n.a.**: Se respetan como texto (no se convierten a números)

### Normalización de regiones

El script mantiene un mapeo automático en `REGION_MAP`:

```python
REGION_MAP = {
    "Europe": "Europe",
    "Asia EUR": "Asia",
    "Arcosa US": "US",
    ...
}
```

Variantes no mapeadas generan un warning en el log.

---

## Troubleshooting

### "FileNotFoundError: [Errno 2] No such file or directory"

Verifica que `config.json` tenga la ruta correcta al archivo Excel:

```bash
# Windows
type config.json | findstr input_file_path
```

### "KeyError: 'TS SC v26.1'"

El nombre de la hoja en el Excel es incorrecto. Verifica en `config.json` que coincida exactamente con el nombre en Excel (mayúsculas, espacios, etc.).

### Años con tipo diferente

Si el script falla al ordenar años, probablemente vienen como strings y ints. Verificar que `config.json` tenga:

```json
"target_years": [2026, 2027, 2028]
```

(sin comillas alrededor de los números)

---

## Autor

Generado automáticamente como parte del pipeline de extracción de datos de Nordex.

## Fecha

Abril 2026
