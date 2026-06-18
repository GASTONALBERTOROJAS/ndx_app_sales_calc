import os
import pandas as pd
from datetime import datetime
from sqlalchemy import create_engine, inspect, text
from dotenv import load_dotenv

def ingest_to_sql(df: pd.DataFrame, source_filename: str, log_callback=None):
    """
    Ingiere el DataFrame procesado en la base de datos PostgreSQL,
    siguiendo las reglas de generic-sql.md para el esquema 01_ingesta.
    """
    load_dotenv()
    
    def _log(msg):
        if log_callback:
            log_callback(msg)
        else:
            print(msg)
            
    _log("Iniciando proceso de ingesta a SQL en 01_ingesta...")
    
    df = df.copy()
    
    # 1. Convertir TODAS las columnas a TEXT
    # según la regla de "01_ingesta" en generic-sql.md
    for col in df.columns:
        # Forzamos todo a string
        df[col] = df[col].astype(str)
        # Manejo de nulos de pandas convertidos a texto
        df[col] = df[col].replace({"<NA>": None, "nan": None, "None": None})
            
    # 3. Leer credenciales desde entorno o archivo .env
    db_user = os.getenv("DB_USER", "postgres")
    db_pass = os.getenv("DB_PASS", "")
    db_host = os.getenv("DB_HOST", "localhost")
    db_port = os.getenv("DB_PORT", "5432")
    db_name = os.getenv("DB_NAME", "powerbi_reports")
    
    if not db_pass:
        _log("WARNING: No se encontró DB_PASS en el entorno. La conexión podría fallar.")
    
    # String de conexión para PostgreSQL
    conn_string = f"postgresql://{db_user}:{db_pass}@{db_host}:{db_port}/{db_name}"
    
    try:
        engine = create_engine(conn_string)
        
        # 4. Generar nombre de la tabla (source_entity_YYYYMMDD)
        date_str = datetime.now().strftime("%Y%m%d")
        table_name = f"salescalc_tower_{date_str}"
        
        _log(f"Conectando a {db_host}:{db_port}/{db_name} ...")
        
        # 5. Borrar tablas de días anteriores para no acumular histórico
        try:
            inspector = inspect(engine)
            tables = inspector.get_table_names(schema="01_ingesta")
            with engine.begin() as conn:
                for t in tables:
                    if t.startswith("salescalc_tower_") and t != table_name:
                        _log(f"Borrando tabla anterior encontrada: 01_ingesta.{t} ...")
                        conn.execute(text(f'DROP TABLE "01_ingesta"."{t}"'))
        except Exception as drop_e:
            _log(f"No se pudieron borrar tablas anteriores: {drop_e}")

        _log(f"Escribiendo {len(df)} filas en 01_ingesta.{table_name} ...")
        
        # 5. Volcar los datos a PostgreSQL
        df.to_sql(
            name=table_name,
            con=engine,
            schema="01_ingesta",
            if_exists="replace", # O 'append', dependiendo de si se corre varias veces al día
            index=False
        )
        _log(f"Ingesta completada exitosamente en 01_ingesta.{table_name}.")
        
    except Exception as e:
        _log(f"ERROR durante la ingesta a PostgreSQL: {e}")
        raise
