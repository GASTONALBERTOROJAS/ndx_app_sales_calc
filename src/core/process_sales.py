import os
import pandas as pd
from sqlalchemy import create_engine, inspect, text
from dotenv import load_dotenv

def process_sales_table():
    load_dotenv()
    
    db_user = os.getenv("DB_USER", "postgres")
    db_pass = os.getenv("DB_PASS", "")
    db_host = os.getenv("DB_HOST", "localhost")
    db_port = os.getenv("DB_PORT", "5432")
    db_name = os.getenv("DB_NAME", "powerbi_reports")
    
    conn_string = f"postgresql://{db_user}:{db_pass}@{db_host}:{db_port}/{db_name}"
    engine = create_engine(conn_string)
    
    # 1. Encontrar la tabla actual en 01_ingesta
    inspector = inspect(engine)
    ingesta_tables = inspector.get_table_names(schema="01_ingesta")
    source_table = None
    for t in ingesta_tables:
        if t.startswith("salescalc_tower_"):
            source_table = t
            break
            
    if not source_table:
        print("ERROR: No se encontró ninguna tabla salescalc_tower_* en 01_ingesta.")
        return
        
    print(f"Leyendo datos crudos desde 01_ingesta.{source_table}...")
    query = f'SELECT * FROM "01_ingesta"."{source_table}"'
    df = pd.read_sql(query, con=engine)
    
    # 2. Eliminar columnas no deseadas para el equipo de ventas
    cols_to_drop = [
        "Type", "Height", "Sections", 
        "Plates Weight net", "Plates Weight gross", "weight flanges"
    ]
    df = df.drop(columns=[c for c in cols_to_drop if c in df.columns], errors="ignore")
    
    # 3. Limpieza y conversión de tipos
    # Las columnas de texto ya son texto (Component_Category, Component, Key, Brand, Region)
    
    # - Year_Production (convertir a integer)
    df["Year_Production"] = pd.to_numeric(df["Year_Production"], errors="coerce").fillna(0).astype(int)
    
    # - Cost_Currency1 y Cost_Currency2 (decimal 10,1, reemplazar N/A por 0)
    def clean_cost(val):
        if pd.isna(val) or str(val).strip() in ["N/A", "None", "nan", "<NA>", ""]:
            return 0.0
        try:
            return float(val)
        except ValueError:
            return 0.0
            
    df["Cost_Currency1"] = df["Cost_Currency1"].apply(clean_cost).round(2)
    df["Cost_Currency2"] = df["Cost_Currency2"].apply(clean_cost).round(2)
    
    # Asegurarnos de usar tipos adecuados para SQL mediante diccionarios si fuera necesario, 
    # pero pandas to_sql infiere float e int automáticamente.
    
    # 4. Crear esquema 03_entrega si no existe
    with engine.begin() as conn:
        conn.execute(text('CREATE SCHEMA IF NOT EXISTS "03_entrega"'))
        
    # 5. Guardar la tabla procesada
    target_table = "salescalc_tower_sales"
    print(f"Escribiendo {len(df)} filas procesadas en 03_entrega.{target_table}...")
    
    from sqlalchemy.types import Numeric, Integer
    dtype_mapping = {
        "Year_Production": Integer(),
        "Cost_Currency1": Numeric(10, 2),
        "Cost_Currency2": Numeric(10, 2)
    }
    
    df.to_sql(
        name=target_table,
        con=engine,
        schema="03_entrega",
        if_exists="replace", # Reemplazará la vista de ventas cada vez que se ejecute
        index=False,
        dtype=dtype_mapping
    )
    
    print("Tabla para el equipo de ventas creada y procesada exitosamente en 03_entrega.")

if __name__ == "__main__":
    process_sales_table()
