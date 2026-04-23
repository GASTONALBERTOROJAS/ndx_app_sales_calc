# Sales Calculation Scripts

Flujo de trabajo para extraer y verificar datos de costos de componentes de torres eólicas.

## Scripts

### 01_extract_tower_data.py
**Propósito:** Extrae datos de costos desde Excel fuente y genera archivo unificado.

**Entrada:**
- Archivo Excel configurado en `config.json` → `input_file_path`
- Dos hojas:
  - `TS SC v26.1`: 400 torres de acero (estructura matricial)
  - `TCS_MB SC v26.1`: 108 torres concreto/híbridas (tabla plana)

**Salida:**
- `output/SalesCalc_DDMMYY.xlsx` (ejemplo: `SalesCalc_230426.xlsx`)
- 133,588 filas con columnas:
  - `Component_Category`, `Component`, `Key`, `Brand`
  - `Year_Production`, `Region`
  - `Cost_EUR`, `Cost_USD`

**Ejecutar:**
```bash
python scripts/01_extract_tower_data.py
```

### 02_verify_output.py
**Propósito:** Valida que los valores extraídos coincidan con valores esperados de la fuente.

**Entrada:**
- Archivo generado por 01 (automáticamente detecta `output/SalesCalc_DDMMYY.xlsx`)
- Test cases predefinidos con torre y componente específicos

**Salida:**
- Reporte de verificación: PASS/FAIL para cada check
- Muestra valores esperados vs obtenidos

**Ejecutar:**
```bash
python scripts/02_verify_output.py
```

### 03_analyze_gaps.py
**Propósito:** Analiza filas sin costo (Cost_EUR ni Cost_USD) para diagnosticar datos faltantes.

**Entrada:**
- Archivo generado por 01 (automáticamente detecta `output/SalesCalc_DDMMYY.xlsx`)

**Salida:**
- Distribución de filas sin costo por componente, región y año
- Ejemplos de torres/componentes afectadas

**Ejecutar:**
```bash
python scripts/03_analyze_gaps.py
```

**Notas:**
- Filas sin costo son **legítimas**: opciones no disponibles en ciertas regiones, componentes concreto sin pricing regional
- No son errores

## Configuración

Ver `config.json` en la raíz del proyecto:
- `input_file_path`: Ruta al archivo Excel fuente
- `source_sheets`: Nombres de las dos hojas a extraer
- `target_years`: Años a procesar (2026, 2027, 2028)
- `canonical_regions`: 11 regiones canonicalizadas

## Flujo típico

1. Actualizar `config.json` si cambian rutas o configuración
2. Ejecutar `01_extract_tower_data.py` para generar salida
3. Ejecutar `02_verify_output.py` para validar
