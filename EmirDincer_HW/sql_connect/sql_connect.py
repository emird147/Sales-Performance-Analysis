# Used sqlalchemy due to problem with pyodbc and pandas

import pandas as pd
from sqlalchemy import create_engine
import urllib

username = "emirdincer34"
password = "CCny24018534"
server = "johndroescher.com"
default_db = "Fall_2024"

connection_string = (
    "mssql+pyodbc:///?odbc_connect=" + 
    urllib.parse.quote_plus(
        f"DRIVER={{ODBC Driver 18 for SQL Server}};"
        f"SERVER={server};DATABASE={default_db};"
        f"UID={username};PWD={password};"
        "TrustServerCertificate=yes"
    )
)

engine = create_engine(connection_string)

my_query = """
SELECT TOP 100 * FROM proj_customers
"""

try:
    data = pd.read_sql(my_query, engine)
    print(data)
    
except Exception as e:
    print("An error occurred:", e)
