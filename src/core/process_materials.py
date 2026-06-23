import os
import pandas as pd
import numpy as np
import logging
from sqlalchemy import create_engine, text
from sqlalchemy.types import String, Numeric, Integer, Float, Text
from datetime import datetime

# Logging setup
logging.basicConfig(level=logging.INFO, format="%(asctime)s [%(levelname)s] %(message)s")
log = logging.getLogger(__name__)

def get_engine():
    """Create SQLAlchemy engine using environment variables."""
    from dotenv import load_dotenv
    load_dotenv()
    
    user = os.getenv("DB_USER", "postgres")
    password = os.getenv("DB_PASS", "admin")
    host = os.getenv("DB_HOST", "localhost")
    port = os.getenv("DB_PORT", "5432")
    db_name = os.getenv("DB_NAME", "sales_calc_db")
    
    conn_str = f"postgresql://{user}:{password}@{host}:{port}/{db_name}"
    return create_engine(conn_str)

def process_materials_table(output_path=None, log_callback=None):
    """
    Reads from 01_ingesta and creates a pivoted table in 03_entrega 
    specifically for the 5 alternative material components.
    """
    def _log(msg: str):
        if log_callback:
            log_callback(msg)
        else:
            log.info(msg)

    engine = get_engine()
    
    # 1. Encontrar la tabla de ingesta más reciente
    try:
        query_tables = """
            SELECT table_name 
            FROM information_schema.tables 
            WHERE table_schema = '01_ingesta' 
              AND table_name LIKE 'salescalc_tower_%'
            ORDER BY table_name DESC LIMIT 1;
        """
        with engine.connect() as conn:
            result = conn.execute(text(query_tables)).fetchone()
            if not result:
                _log("No se encontraron tablas de ingesta en el esquema '01_ingesta'.")
                return
            latest_table = result[0]
            _log(f"Leyendo datos crudos desde 01_ingesta.{latest_table} para Materials...")
            
            df = pd.read_sql_table(latest_table, con=conn, schema='01_ingesta')
    except Exception as e:
        _log(f"Error leyendo base de datos: {e}")
        return

    # 2. Filtrar solo los componentes de materiales
    materials_components = [
        "Steel Plates",
        "Flanges",
        "Conversion",
        "Thereof Damper",
        "Thereof D4K-cable"
    ]
    
    df_mat = df[df["Component_Category"] == "Tower"].copy()
    df_mat = df_mat[df_mat["Component"].isin(materials_components)].copy()
    
    if df_mat.empty:
        _log("No se encontraron registros de materiales en la ingesta.")
        return

    # 3. Calcular el Costo Consolidado (Euro) para cada fila de material
    df_mat["Cost_Currency1"] = pd.to_numeric(df_mat["Cost_Currency1"], errors="coerce").fillna(0)
    df_mat["Cost_Currency2"] = pd.to_numeric(df_mat["Cost_Currency2"], errors="coerce").fillna(0)
    df_mat["Total Cost (Euro)"] = df_mat["Cost_Currency1"] + (df_mat["Cost_Currency2"] / 1.15)
    
    # 4. Formatear y preparar el Full Tower Name
    df_mat["Brand"] = df_mat["Brand"].fillna("Nx").astype(str).str.strip()
    df_mat["Type"] = df_mat["Type"].fillna("").astype(str).str.strip()
    df_mat["Height"] = pd.to_numeric(df_mat["Height"], errors="coerce").fillna(0).astype(int)
    
    def get_full_name(row):
        brand = row["Brand"]
        t_type = row["Type"]
        height = row["Height"]
        # Ignorar alturas 0 en el nombre si las hay
        if height > 0:
            return f"Tower {brand} {t_type} {height}m"
        else:
            return f"Tower {brand} {t_type}".strip()
            
    df_mat["Full Tower Name"] = df_mat.apply(get_full_name, axis=1)
    
    # Category (Misma regla que Power BI)
    def get_category(row):
        t_type = str(row["Type"]).strip().upper()
        h = row["Height"]
        if t_type == "TCS":
            return "Hibrid Towers"
        elif t_type == "TS":
            if h <= 120:
                return "TS <= 120M"
            elif h <= 148:
                return "TS 120M - 148M"
            else:
                return "TS > 148M"
        return "Other"
        
    df_mat["Category"] = df_mat.apply(get_category, axis=1)
    
    # Platform
    df_mat["Platform"] = df_mat["Platform"].fillna("Other").astype(str).str.strip()
    
    # 5. Pivotar los datos para que cada componente sea una columna
    group_cols = [
        "Full Tower Name",
        "Platform",
        "Category",
        "Type",
        "Height",
        "Year_Production",
        "Region"
    ]
    
    # Limpiar N/A
    df_mat = df_mat[df_mat["Total Cost (Euro)"] > 0].copy()
    
    # Pivot
    df_pivot = df_mat.pivot_table(
        index=group_cols,
        columns="Component",
        values="Total Cost (Euro)",
        aggfunc="sum",
        fill_value=0
    ).reset_index()
    
    # Asegurar que todas las columnas existan, incluso si no vinieron datos
    for comp in materials_components:
        if comp not in df_pivot.columns:
            df_pivot[comp] = 0.0
            
    # Opcional: Calcular el subtotal sumando los materiales base
    df_pivot["Total Materials Cost"] = df_pivot["Steel Plates"] + df_pivot["Flanges"] + df_pivot["Conversion"]
            
    # 6. Escribir a SQL esquema 03_entrega
    dtype_mapping = {
        "Full Tower Name": String(255),
        "Platform": Text(),
        "Category": String(50),
        "Type": String(50),
        "Height": Integer(),
        "Year_Production": Integer(),
        "Region": String(100),
        "Steel Plates": Float(),
        "Flanges": Float(),
        "Conversion": Float(),
        "Thereof Damper": Float(),
        "Thereof D4K-cable": Float(),
        "Total Materials Cost": Float()
    }

    try:
        with engine.begin() as conn:
            conn.execute(text("CREATE SCHEMA IF NOT EXISTS \"03_entrega\";"))
            
        df_pivot.to_sql(
            name="tower_materials",
            con=engine,
            schema="03_entrega",
            if_exists="replace",
            index=False,
            dtype=dtype_mapping
        )
        _log(f"Escribiendo {len(df_pivot)} filas procesadas en 03_entrega.tower_materials...")
        _log("Tabla de materiales alternativa creada exitosamente en 03_entrega.")
    except Exception as e:
        _log(f"Error escribiendo en PostgreSQL: {e}")

if __name__ == "__main__":
    process_materials_table()
