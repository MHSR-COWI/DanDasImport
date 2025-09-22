import pyodbc
import pandas as pd
from tkinter import Tk, filedialog
import os

import subprocess
import sys

def install_and_import(package):
    try:
        __import__(package)
        print(f"'{package}' er allerede installeret.")
    except ImportError:
        print(f"'{package}' er ikke installeret. Installerer nu...")
        subprocess.check_call([sys.executable, "-m", "pip", "install", package])
        print(f"'{package}' er nu installeret.")
    finally:
        globals()[package] = __import__(package)

# Tjek og installer pakker
install_and_import('pandas')
install_and_import('openpyxl')

# Nu kan du bruge pandas og openpyxl uden problemer
import pandas as pd



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
all_dfs_prefixed = []

for name in query_names:
    sql = f"SELECT * FROM [{name}]"
    df = pd.read_sql(sql, conn)
    df = df.dropna(how='all')  # fjern tomme rækker
    df = df.add_prefix(f"{name}_")  # tilføj præfiks til kolonner for at undgå overlap
    all_dfs_prefixed.append(df)

conn.close()

# Saml data horisontalt (side om side i kolonner)
df_samlet = pd.concat(all_dfs_prefixed, axis=1)

# Erstat NaN med tomme strenge, så det bliver pænere i Excel
df_samlet.fillna('', inplace=True)

# Vælg navn og placering på Excel-fil
excel_path = os.path.join(
    os.path.dirname(db_path),
    "samlet_udtræk_queries.xlsx"
)

# Gem til ét ark
df_samlet.to_excel(excel_path, sheet_name="SamletData", index=False)

# Gem samlet datafil
df_samlet.to_excel(excel_path, sheet_name="SamletData", index=False)
print(f"\nAlt data er samlet i arket 'SamletData' i filen:\n{excel_path}")

print(f"\nAlt data er samlet i arket 'SamletData' i filen:\n{excel_path}")










