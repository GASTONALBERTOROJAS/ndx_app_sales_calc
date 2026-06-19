import os
import pandas as pd
import numpy as np
from sqlalchemy import create_engine, inspect, text
from dotenv import load_dotenv

def process_powerbi_table():
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
    
    # 2. Filtrar Componentes
    df = df[df["Component"].isin(["Tower Shell", "Tower Internals"])].copy()
    
    # 3. Eliminar columnas que no van o impiden agrupar
    cols_to_drop = ["Component", "Component_Category", "Brand"]
    df = df.drop(columns=[c for c in cols_to_drop if c in df.columns], errors="ignore")
    
    # 4. Limpieza y conversión de tipos numéricos ANTES de agrupar
    def clean_num(val):
        if pd.isna(val) or str(val).strip() in ["N/A", "None", "nan", "<NA>", ""]:
            return 0.0
        try:
            return float(val)
        except ValueError:
            return 0.0

    numeric_cols = [
        "Height", "Sections", "Plates Weight net", "Plates Weight gross", 
        "weight flanges", "Year_Production", "Cost_Currency1", "Cost_Currency2"
    ]
    for col in numeric_cols:
        if col in df.columns:
            df[col] = df[col].apply(clean_num)
            
    # Agrupar por las columnas estáticas de la torre para tener 1 fila por torre
    group_cols = [c for c in df.columns if c not in ["Cost_Currency1", "Cost_Currency2"]]
    df = df.groupby(group_cols, as_index=False)[["Cost_Currency1", "Cost_Currency2"]].sum()
    
    # Renombrar columnas
    rename_map = {
        "Key": "Full Tower Name",
        "Type": "Tower Type",
        "Height": "Tower Height"
    }
    df = df.rename(columns=rename_map)
    
    # 5. Limpieza y conversión de tipos numéricos
    def clean_num(val):
        if pd.isna(val) or str(val).strip() in ["N/A", "None", "nan", "<NA>", ""]:
            return 0.0
        try:
            return float(val)
        except ValueError:
            return 0.0

    numeric_cols = [
        "Tower Height", "Sections", "Plates Weight net", "Plates Weight gross", 
        "weight flanges", "Year_Production", "Cost_Currency1", "Cost_Currency2"
    ]
    for col in numeric_cols:
        if col in df.columns:
            df[col] = df[col].apply(clean_num)
            
    # Casteo a int donde corresponda
    df["Tower Height"] = df["Tower Height"].astype(int)
    df["Sections"] = df["Sections"].astype(int)
    df["Year_Production"] = df["Year_Production"].astype(int)
    
    # Los decimales de los pesos y costos se mantienen con su precisión original
    # para que los cálculos de PowerBI cuadren exacto con los del Excel original.
            
    # 6. Lógica Calculada (Replicando el Excel)
    
    # Exchange Rate Applied
    df["Exchange Rate Applied"] = "1 EUR = 1.15 USD"
    
    # Total Cost Consolidated (Euro)
    df["Total Cost Consolidated (Euro)"] = df["Cost_Currency1"] + (df["Cost_Currency2"] / 1.15)
    
    # Total Weight (tons)
    df["Total Weight (tons)"] = df["Plates Weight net"] + df["weight flanges"]
    
    # Categoría
    def get_category(row):
        t_type = str(row["Tower Type"]).strip().upper()
        h = row["Tower Height"]
        if t_type == "TCS":
            return "Híbrida"
        elif t_type == "TS":
            if h <= 120:
                return "Acero <= 120m"
            elif h <= 148:
                return "Acero 120m - 148m"
            else:
                return "Acero > 148m"
        return "Otro"
        
    df["Categoría"] = df.apply(get_category, axis=1)
    
    # Ratios (manejando división por cero)
    df["Ratio EUR/Ton"] = np.where(df["Total Weight (tons)"] > 0, 
                                   df["Total Cost Consolidated (Euro)"] / df["Total Weight (tons)"], 0)
                                   
    df["Ratio EUR/m"] = np.where(df["Tower Height"] > 0, 
                                 df["Total Cost Consolidated (Euro)"] / df["Tower Height"], 0)
                                 
    df["Ratio EUR/Sección"] = np.where(df["Sections"] > 0, 
                                       df["Total Cost Consolidated (Euro)"] / df["Sections"], 0)
    
    # Eliminar filas donde el costo total es 0 (ej. regiones donde vino 'N/A')
    df = df[df["Total Cost Consolidated (Euro)"] > 0].copy()
    
    # 7. Aplicar formato final solicitado (decimal 10,2)
    cols_to_round = [
        "Plates Weight net", "Plates Weight gross", "weight flanges",
        "Total Cost Consolidated (Euro)", "Total Weight (tons)", 
        "Ratio EUR/Ton", "Ratio EUR/m", "Ratio EUR/Sección"
    ]
    for col in cols_to_round:
        if col in df.columns:
            df[col] = df[col].round(2)
            
    # 8. Crear Identificador Único para Power BI
    df["Unique ID"] = df["Full Tower Name"].astype(str) + " | " + df["Region"].astype(str) + " | " + df["Year_Production"].astype(int).astype(str)
    
    # 9. Reordenar columnas y eliminar Cost_Currency1 y Cost_Currency2
    final_cols = [
        "Unique ID",
        "Full Tower Name",
        "Platform",
        "Categoría",
        "Tower Type",
        "Tower Height",
        "Sections",
        "Region",
        "Year_Production",
        "Exchange Rate Applied",
        "Plates Weight net",
        "Plates Weight gross",
        "weight flanges",
        "Total Cost Consolidated (Euro)",
        "Total Weight (tons)",
        "Ratio EUR/Ton",
        "Ratio EUR/m",
        "Ratio EUR/Sección"
    ]
    df = df[final_cols]
    
    # 9. Guardar la tabla
    with engine.begin() as conn:
        conn.execute(text('CREATE SCHEMA IF NOT EXISTS "03_entrega"'))
        
    target_table = "tower_powerbi"
    print(f"Escribiendo {len(df)} filas procesadas en 03_entrega.{target_table}...")
    
    from sqlalchemy.types import Numeric, Integer, Text
    dtype_mapping = {
        "Unique ID": Text(),
        "Platform": Text(),
        "Tower Height": Integer(),
        "Sections": Integer(),
        "Year_Production": Integer(),
        "Plates Weight net": Numeric(10, 2),
        "Plates Weight gross": Numeric(10, 2),
        "weight flanges": Numeric(10, 2),
        "Total Cost Consolidated (Euro)": Numeric(15, 2),
        "Total Weight (tons)": Numeric(10, 2),
        "Ratio EUR/Ton": Numeric(10, 2),
        "Ratio EUR/m": Numeric(10, 2),
        "Ratio EUR/Sección": Numeric(10, 2),
    }
    
    df.to_sql(
        name=target_table,
        con=engine,
        schema="03_entrega",
        if_exists="replace",
        index=False,
        dtype=dtype_mapping
    )
    
    print("Tabla para Power BI creada y procesada exitosamente en 03_entrega.")

if __name__ == "__main__":
    process_powerbi_table()
