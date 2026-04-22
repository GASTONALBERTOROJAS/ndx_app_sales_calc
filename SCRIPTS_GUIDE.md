# Scripts Guide

## Overview

El proyecto contiene **2 scripts Python**:

1. **`extract_tower_data.py`** — Script PRINCIPAL (obligatorio)
2. **`verify_output.py`** — Script OPCIONAL (para verificación)

---

## 1. `extract_tower_data.py` — EL SCRIPT PRINCIPAL

### ¿Qué hace?

Es el **script de extracción** que:

1. Lee la configuración desde `config.json`
2. Abre el archivo Excel fuente
3. Lee dos hojas:
   - **TS SC v26.1** (400 torres de acero)
   - **TCS_MB SC v26.1** (108 torres de concreto/híbridas)
4. Extrae todos los datos de componentes
5. Normaliza:
   - Nombres de componentes
   - Regiones (a 11 canonicales)
   - Años (2026, 2027, 2028)
   - Monedas (EUR y USD)
6. Genera **`output/tower_components_extracted.xlsx`** con 133,588 filas

### Uso

```bash
python scripts/extract_tower_data.py
```

### Tiempo de ejecución

~45 segundos

### Output

- **Archivo:** `output/SalesCalc_DDMMYY.xlsx` (nombre dinámico con fecha)
- **Ejemplo:** `SalesCalc_220426.xlsx` (22 de abril de 2026)
- **Filas:** 133,588 (datos consolidados de ambas hojas)
- **Columnas:** 8 (Component_Category, Component, Key, Brand, Year_Production, Region, Cost_EUR, Cost_USD)

**Ventaja:** Cada ejecución genera un archivo diferente con la fecha, evitando sobrescribir archivos anteriores.

### ¿Cuándo ejecutar?

**Siempre** — Es el script que tienes que correr normalmente para obtener los datos.

---

## 2. `verify_output.py` — SCRIPT DE VERIFICACIÓN (OPCIONAL)

### ¿Qué hace?

Es un **script de validación** que:

1. Lee el archivo generado: `output/tower_components_extracted.xlsx`
2. Verifica una **muestra de 12 torres** contra valores conocidos
3. Compara:
   - Valores de Cost_EUR vs source
   - Valores de Cost_USD vs source
4. Reporta si todo coincide correctamente

### Uso

```bash
python scripts/verify_output.py
```

### Output

Muestra un reporte:

```
=== VERIFICATION: Source vs Output ===
Tower: Tower N117/3000 Controlled IEC2a TS76 TiT 50Hz NCV
Region: Europe, Year: 2026

OK  Tower Shell: source=255993.398875 -> output=255993.398875
OK  Tower Internals: source=151000 -> output=151000
OK  Anchor cage: source=22339.23818 -> output=22339.23818
...

ALL CHECKS PASSED
```

### ¿Cuándo ejecutar?

**OPCIONAL** — Ejecuta este script si:

- Quieres verificar que la extracción fue correcta
- Recibes un nuevo archivo y quieres validar antes de usar
- Deseas auditar que los valores son precisos
- Eres un usuario técnico que necesita confirmar calidad de datos

### Nota Importante

⚠️ **Este script tiene rutas hardcodeadas** y está configurado para la primera extracción. Si lo quieres usar en el futuro, necesitarías actualizar las rutas internas. Es principalmente una herramienta de **QA/Testing**, no de uso recurrente.

---

## Flujo de Uso Típico

### Usuario Normal (NO técnico):

```
1. setup.bat (Windows) o bash setup.sh (Linux/Mac)
2. python scripts/extract_tower_data.py
3. ✓ Abrir output/tower_components_extracted.xlsx
```

**Script usado:** Solo `extract_tower_data.py`

---

### Usuario Técnico / QA (con verificación):

```
1. setup.bat (Windows) o bash setup.sh (Linux/Mac)
2. python scripts/extract_tower_data.py
3. python scripts/verify_output.py
4. ✓ Revisar reporte
5. ✓ Abrir output/tower_components_extracted.xlsx
```

**Scripts usados:** Ambos

---

## Dependencias

### `extract_tower_data.py`

Requiere:
- `openpyxl` — Lectura de archivos Excel
- `pandas` — Manipulación de dataframes
- `json` — Lectura de config.json
- `re` — Expresiones regulares (normalización)
- `logging` — Logs de proceso

### `verify_output.py`

Requiere:
- `pandas` — Lectura de Excel
- Nada más

Ambas se instalan con: `pip install -r requirements.txt`

---

## Diferencias Clave

| Aspecto | extract_tower_data.py | verify_output.py |
|---|---|---|
| **Objetivo** | Extraer datos | Verificar resultados |
| **Obligatorio** | ✓ SÍ | ✗ NO |
| **Usa config.json** | ✓ SÍ | ✗ NO |
| **Input** | Excel fuente | Excel generado |
| **Output** | Excel con 133,588 filas | Reporte en consola |
| **Tiempo** | ~45 segundos | ~2 segundos |
| **Uso recurrente** | ✓ SÍ | ✗ NO (solo para QA) |
| **Rutas** | Din (config.json) | Hard (rutas fijas) |

---

## Resumen

**En 99% de los casos, solo necesitas correr:**

```bash
python scripts/extract_tower_data.py
```

**El script de verificación (`verify_output.py`) es opcional y está más pensado para:**
- Desarrolladores
- QA / Control de calidad
- Primera validación de la extracción
- Auditoría de datos

No necesitas ejecutarlo para usar el proyecto normalmente.
