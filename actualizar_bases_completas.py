import os
import pandas as pd
import openpyxl

dir_path = r"c:\Users\Maria Alejandra\OneDrive - VIRTUALSOFT SERVICIOS & SOFTWARE S.A.S\Antigravity\Retiros"
desk_path = r"c:\Users\Maria Alejandra\OneDrive - VIRTUALSOFT SERVICIOS & SOFTWARE S.A.S\Escritorio\Retiros Septiembre.xlsx"
ago_full_path = os.path.join(dir_path, "Analisis automatización Agosto.xlsx")

standard_cols = [
    'Marca', 'Id Usuario', 'Id Retiro', 'Estado Retiro Creado',
    'Fecha Cambio', 'Fecha Cambio Time', 'Fecha Creacion Retiro', 'Fecha Creacion Retiro Time',
    'Fecha Pago', 'Fecha Pago Time', 'Medio De Pago', 'Nivel De Riesgo', 'Tipo Aprobacion',
    'Nombre Usuario Cambio', 'Nombre Usuario Pago', 'Nombre Usuario Rechaza',
    'Nombre Punto Venta', 'Nombre Producto (Pasarela)', 'Nombre Banco', 'Valor Retiros Creados'
]

print("=== 1. Procesando base completa de Agosto ===")
df_ago = pd.read_excel(ago_full_path, sheet_name='Retiros Mensuales')
df_ago = df_ago.dropna(subset=['Marca'])
df_ago['Marca'] = df_ago['Marca'].astype(str).str.strip()
df_ago['Marca'] = df_ago['Marca'].replace({'Paniplay Honduras': 'Catraplay Honduras'})

# Asegurar columnas estándar
for c in standard_cols:
    if c not in df_ago.columns:
        df_ago[c] = None
df_ago = df_ago[standard_cols]

print(f"Total filas Agosto: {len(df_ago)}")
ago_csv = os.path.join(dir_path, "Retiros Agosto_cache.csv")
df_ago.to_csv(ago_csv, index=False)
print(f"[+] Guardado {ago_csv}")

ago_excel = os.path.join(dir_path, "Retiros Agosto.xlsx")
# Para que sea rápido y estándar
with pd.ExcelWriter(ago_excel, engine='openpyxl') as writer:
    df_ago.to_excel(writer, sheet_name='Retiros BD Conexion', index=False)
print(f"[+] Guardado {ago_excel}")

print("\n=== 2. Procesando base completa de Septiembre ===")
df_sep = pd.read_excel(desk_path, sheet_name='Retiros Mensuales')
if 'Marca' not in df_sep.columns or df_sep['Marca'].isna().all():
    df_sep['Marca'] = df_sep['Partner'].astype(str).str.strip() + ' ' + df_sep['Pais'].astype(str).str.strip()
df_sep = df_sep.dropna(subset=['Marca'])
df_sep['Marca'] = df_sep['Marca'].astype(str).str.strip()
df_sep['Marca'] = df_sep['Marca'].replace({'Paniplay Honduras': 'Catraplay Honduras'})

for c in standard_cols:
    if c not in df_sep.columns:
        df_sep[c] = None
df_sep = df_sep[standard_cols]

print(f"Total filas Septiembre: {len(df_sep)}")
sep_csv = os.path.join(dir_path, "Retiros Septiembre_cache.csv")
df_sep.to_csv(sep_csv, index=False)
print(f"[+] Guardado {sep_csv}")

sep_excel = os.path.join(dir_path, "Retiros Septiembre.xlsx")
with pd.ExcelWriter(sep_excel, engine='openpyxl') as writer:
    df_sep.to_excel(writer, sheet_name='Retiros BD Conexion', index=False)
print(f"[+] Guardado {sep_excel}")

# Eliminar resúmenes precalculados antiguos para forzar recálculo limpio con la nueva data completa
for fname in ["Retiros Agosto_summary.json", "Retiros Septiembre_summary.json"]:
    fpath = os.path.join(dir_path, fname)
    if os.path.exists(fpath):
        os.remove(fpath)
        print(f"[+] Eliminado resumen antiguo: {fname}")

print("\n[+] Sincronización de bases completas terminada exitosamente.")
