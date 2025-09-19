import pyodbc
import pandas as pd
from tkinter import Tk, filedialog

# --- Stifinder: Vælg Access database ---
root = Tk()
root.withdraw()  # Skjul hovedvinduet
file_path = filedialog.askopenfilename(
    title="Vælg din Access-database (.accdb eller .mdb)",
    filetypes=[("Access Database", "*.accdb *.mdb"), ("Alle filer", "*.*")]
)

if not file_path:
    print("Du har ikke valgt nogen database-fil.")
    exit(1)

print(f"Valgt database: {file_path}")
# ---------------------------------------

# Udskift med dine egne queries
query_names = ['DDG_ledning_v', 'DDG_knude_v','ddg_opland_v']

# Connection string
conn_str = (
    r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};'
    fr'DBQ={file_path};'
)

conn = pyodbc.connect(conn_str)

for name in query_names:
    sql = f"SELECT * FROM [{name}]"
    df = pd.read_sql(sql, conn)
    print(f"----- Data fra {name} -----")
    print(df.head())
    # Gem evt. som CSV:
    # df.to_csv(f"{name}.csv", index=False)

conn.close()