import pyodbc
import pandas as pd
from tkinter import Tk, filedialog
import os

# Vælg Access-database
root = Tk()
root.withdraw()
db_path = filedialog.askopenfilename(
    title="Vælg Access-database (.accdb eller .mdb)",
    filetypes=[("Access database", "*.accdb *.mdb"), ("Alle filer", "*.*")]
)

if not db_path:
    print("Ingen database valgt. Scriptet afsluttes.")
    exit(1)

# Udfyld med dine query-navne (skal matche i Access)
query_names = ['DDG_ledning_v', 'DDG_knude_v','ddg_opland_v']

# Forbind til Access
conn_str = (
    r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};'
    fr'DBQ={db_path};'
)
conn = pyodbc.connect(conn_str)

# Saml alle dataframes i en liste
all_dfs = []

for name in query_names:
    sql = f"SELECT * FROM [{name}]"
    df = pd.read_sql(sql, conn)
    # Tilføj evt. kolonne med query-navn til identificering
    df["KildeQuery"] = name
    all_dfs.append(df)

conn.close()

# Sammensæt alle DataFrames til én stor DataFrame
df_samlet = pd.concat(all_dfs, ignore_index=True)

# Vælg navn og placering på Excel-fil
excel_path = os.path.join(
    os.path.dirname(db_path),
    "samlet_udtræk_queries.xlsx"
)

# Gem til ét ark
df_samlet.to_excel(excel_path, sheet_name="SamletData", index=False)

print(f"\nAlt data er samlet i arket 'SamletData' i filen:\n{excel_path}")








