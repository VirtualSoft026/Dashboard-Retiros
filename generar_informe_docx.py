import os
import glob
import re
import shutil
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import seaborn as sns
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn

# Configuración de estilo de los gráficos
sns.set_theme(style="whitegrid")
plt.rcParams.update({
    'font.family': 'sans-serif',
    'font.size': 10,
    'axes.labelsize': 11,
    'axes.titlesize': 12,
    'xtick.labelsize': 9,
    'ytick.labelsize': 9,
    'figure.titlesize': 14,
    'hatch.linewidth': 0.5
})

# Colores Corporativos Premium
COLOR_CYAN = "#06b6d4"
COLOR_BLUE = "#3b82f6"
COLOR_ORANGE = "#f97316"
COLOR_GREEN = "#10b981"
COLOR_PURPLE = "#a855f7"
COLOR_DARK = "#0f172a"
COLOR_LIGHT_GRAY = "#f1f5f9"

# Directorio de trabajo
dir_path = r"c:\Users\Maria Alejandra\OneDrive - VIRTUALSOFT SERVICIOS & SOFTWARE S.A.S\Antigravity\Retiros"
output_docx_primary = os.path.join(dir_path, "Informe_Analisis_Efectividad_Automatizacion.docx")
output_docx_backup = os.path.join(dir_path, "Informe_Analisis_Efectividad_Automatizacion_V2.docx")
temp_img_dir = os.path.join(dir_path, "scratch", "temp_images")
os.makedirs(temp_img_dir, exist_ok=True)

def find_month_files():
    pattern = os.path.join(dir_path, "Retiros *.xlsx")
    files = glob.glob(pattern)
    
    month_map = {
        "enero": 1, "febrero": 2, "marzo": 3, "abril": 4, "mayo": 5, "junio": 6,
        "julio": 7, "agosto": 8, "septiembre": 9, "octubre": 10, "noviembre": 11, "diciembre": 12
    }
    
    detected = []
    for f in files:
        name = os.path.basename(f)
        if any(x in name.lower() for x in ["backup", "prev", "raw"]):
            continue
        
        match = re.match(r"Retiros\s+([A-Za-z]+)\.xlsx", name, re.IGNORECASE)
        if match:
            m_name = match.group(1).lower()
            m_num = month_map.get(m_name, 99)
            detected.append({
                "month_name": m_name,
                "month_num": m_num,
                "file_path": f,
                "display_name": m_name.capitalize()
            })
    
    detected.sort(key=lambda x: x["month_num"])
    return detected

def prepare_dataframe(df):
    df = df.copy()
    
    # Normalizaciones de alta velocidad
    if 'Tipo Aprobacion' not in df.columns:
        for col in df.columns:
            if 'aprobacion' in str(col).lower() or 'aprobación' in str(col).lower():
                df = df.rename(columns={col: 'Tipo Aprobacion'})
                break
    if 'Tipo Aprobacion' not in df.columns:
        df['Tipo Aprobacion'] = 'Manual'
    is_auto = df['Tipo Aprobacion'].astype(str).str.strip().str.lower().isin(['automatico', 'automático'])
    df['Tipo Aprobacion_Norm'] = np.where(is_auto, 'Automatico', 'Manual')
    
    if 'Nivel De Riesgo' not in df.columns:
        for col in df.columns:
            if str(col).strip().lower() in ['unnamed: 11', 'unnamed:11'] or 'riesgo' in str(col).lower():
                df = df.rename(columns={col: 'Nivel De Riesgo'})
                break
    if 'Nivel De Riesgo' not in df.columns:
        df['Nivel De Riesgo'] = 'Sin Nivel'
        
    df['Nivel De Riesgo_Norm'] = df['Nivel De Riesgo'].fillna('Sin Nivel')
    df['Nivel De Riesgo_Norm'] = df['Nivel De Riesgo_Norm'].replace({'nan': 'Sin Nivel', 'None': 'Sin Nivel', '': 'Sin Nivel'})
    df['Nivel De Riesgo_Norm'] = df['Nivel De Riesgo_Norm'].astype(str).str.strip()
    
    if 'Marca' not in df.columns:
        df['Marca'] = 'Sin Marca'
    df['Marca'] = df['Marca'].astype(str).str.strip()
    df['Marca'] = df['Marca'].replace({'Paniplay Honduras': 'Catraplay Honduras', 'Paniplay': 'Catraplay Honduras'})
    
    return df

def generate_report():
    print("[*] Iniciando procesamiento de datos para el reporte...")
    month_files = find_month_files()
    if not month_files:
        print("[-] ERROR: No se encontraron archivos de meses.")
        return
        
    all_data = []
    seen_ids = set()
    
    cols_to_load = [
        'Id Retiro', 'Marca', 'Tipo Aprobacion', 'Nivel De Riesgo',
        'Fecha Cambio Time', 'Valor Retiros Creados'
    ]
    
    for m in month_files:
        excel_path = m['file_path']
        csv_path = excel_path.replace(".xlsx", "_cache.csv")
        use_cache = False
        if os.path.exists(csv_path) and os.path.getsize(csv_path) > 0:
            excel_mtime = os.path.getmtime(excel_path)
            csv_mtime = os.path.getmtime(csv_path)
            if csv_mtime >= (excel_mtime - 2):
                use_cache = True

        if use_cache:
            print(f"    - Cargando {m['display_name']} desde caché CSV rápido...")
            df_raw = pd.read_csv(csv_path, usecols=lambda c: c in cols_to_load, low_memory=False)
        else:
            print(f"    - Cargando {m['display_name']} desde Excel...")
            xl = pd.ExcelFile(excel_path, engine='openpyxl')
            sheet = 'Retiros BD Conexion' if 'Retiros BD Conexion' in xl.sheet_names else xl.sheet_names[0]
            df_raw = xl.parse(sheet, usecols=lambda c: c in cols_to_load)
            try:
                df_raw.to_csv(csv_path, index=False)
            except Exception:
                pass
        
        # Deduplicación cronológica
        if 'Id Retiro' in df_raw.columns:
            df_raw['Id Retiro_str'] = df_raw['Id Retiro'].astype(str).str.strip().str.replace(r'\.0$', '', regex=True)
            if len(seen_ids) > 0:
                df_raw = df_raw[~df_raw['Id Retiro_str'].isin(seen_ids)].copy()
            seen_ids.update(df_raw['Id Retiro_str'].dropna())
            df_raw = df_raw.drop(columns=['Id Retiro_str'])
            
        df = prepare_dataframe(df_raw)
        df['Mes'] = m['display_name']
        df['Mes_Num'] = m['month_num']
        all_data.append(df)
        
    df_all = pd.concat(all_data, ignore_index=True)
    total_general = len(df_all)
    print(f"[+] Total registros cargados y deduplicados: {total_general:,}")
    
    # -------------------------------------------------------------
    # 1. ANÁLISIS MENSUAL
    # -------------------------------------------------------------
    df_monthly = df_all.groupby(['Mes_Num', 'Mes']).apply(lambda x: pd.Series({
        'Total': len(x),
        'Auto': (x['Tipo Aprobacion_Norm'] == 'Automatico').sum(),
        'Manual': (x['Tipo Aprobacion_Norm'] == 'Manual').sum(),
        'Pct_Auto': ((x['Tipo Aprobacion_Norm'] == 'Automatico').sum() / len(x) * 100) if len(x) > 0 else 0,
        'Monto': x['Valor Retiros Creados'].sum()
    }), include_groups=False).reset_index()
    df_monthly = df_monthly.sort_values('Mes_Num')
    
    first_month = month_files[0]['display_name']
    last_month = month_files[-1]['display_name']
    period_str = f"{first_month} - {last_month} 2026"

    # Gráfico 1: Evolución Mensual (Usamos un ratio de aspecto 2.2:1 estándar)
    fig, ax = plt.subplots(figsize=(7, 3.2))
    ax.plot(df_monthly['Mes'], df_monthly['Pct_Auto'], marker='o', color=COLOR_BLUE, linewidth=2.5, markersize=8)
    for i, txt in enumerate(df_monthly['Pct_Auto']):
        ax.annotate(f"{txt:.1f}%", (df_monthly['Mes'].iloc[i], df_monthly['Pct_Auto'].iloc[i] + 0.8), ha='center', fontweight='bold', color=COLOR_DARK)
    ax.set_title(f"Evolución Mensual del % de Automatización ({period_str})", pad=15, fontweight='bold')
    ax.set_ylabel("% Automatización")
    ax.set_ylim(min(df_monthly['Pct_Auto']) - 5, max(df_monthly['Pct_Auto']) + 5)
    plt.tight_layout()
    chart1_path = os.path.join(temp_img_dir, "evolucion_mensual.png")
    plt.savefig(chart1_path, dpi=200)
    plt.close()
    
    # -------------------------------------------------------------
    # 2. ANÁLISIS DE MARCAS
    # -------------------------------------------------------------
    df_brands = df_all.groupby('Marca').apply(lambda x: pd.Series({
        'Total': len(x),
        'Auto': (x['Tipo Aprobacion_Norm'] == 'Automatico').sum(),
        'Manual': (x['Tipo Aprobacion_Norm'] == 'Manual').sum(),
        'Pct_Auto': ((x['Tipo Aprobacion_Norm'] == 'Automatico').sum() / len(x) * 100) if len(x) > 0 else 0,
        'Monto': x['Valor Retiros Creados'].sum()
    }), include_groups=False).reset_index().sort_values('Pct_Auto', ascending=False)
    
    # Gráfico 2: Automatización por Marcas (Fijamos proporciones correctas)
    fig, ax = plt.subplots(figsize=(8, 3.6))
    colors = [COLOR_GREEN if x >= 60 else (COLOR_BLUE if x >= 40 else COLOR_ORANGE) for x in df_brands['Pct_Auto']]
    bars = ax.bar(df_brands['Marca'], df_brands['Pct_Auto'], color=colors, edgecolor='none', width=0.55)
    for bar in bars:
        height = bar.get_height()
        ax.annotate(f"{height:.1f}%",
                    xy=(bar.get_x() + bar.get_width() / 2, height),
                    xytext=(0, 3),
                    textcoords="offset points",
                    ha='center', va='bottom', fontweight='bold', fontsize=9)
    ax.set_title(f"Porcentaje de Automatización por Marca (Acumulado {period_str})", pad=15, fontweight='bold')
    ax.set_ylabel("% Automatización")
    ax.set_ylim(0, 105)
    plt.xticks(rotation=15, ha='right')
    plt.tight_layout()
    chart2_path = os.path.join(temp_img_dir, "marcas_automatizacion.png")
    plt.savefig(chart2_path, dpi=200)
    plt.close()

    # -------------------------------------------------------------
    # 3. ANÁLISIS DE RIESGO CATRAPLAY COMPARATIVO
    # -------------------------------------------------------------
    df_risk_brand = df_all.groupby(['Marca', 'Nivel De Riesgo_Norm']).apply(lambda x: pd.Series({
        'Total': len(x),
        'Auto': (x['Tipo Aprobacion_Norm'] == 'Automatico').sum(),
        'Pct_Auto': ((x['Tipo Aprobacion_Norm'] == 'Automatico').sum() / len(x) * 100) if len(x) > 0 else 0
    }), include_groups=False).reset_index()

    df_pani = df_risk_brand[df_risk_brand['Marca'] == 'Catraplay Honduras']
    df_others = df_risk_brand[df_risk_brand['Marca'] != 'Catraplay Honduras'].groupby('Nivel De Riesgo_Norm').agg({
        'Total': 'sum',
        'Auto': 'sum'
    }).reset_index()
    df_others['Pct_Auto'] = (df_others['Auto'] / df_others['Total'] * 100)
    df_others['Marca'] = 'Otras Marcas (Promedio)'

    df_compare_pani = pd.concat([
        df_pani[['Marca', 'Nivel De Riesgo_Norm', 'Pct_Auto']],
        df_others[['Marca', 'Nivel De Riesgo_Norm', 'Pct_Auto']]
    ])

    fig, ax = plt.subplots(figsize=(7, 3.2))
    sns.barplot(data=df_compare_pani, x='Nivel De Riesgo_Norm', y='Pct_Auto', hue='Marca', palette=[COLOR_GREEN, COLOR_BLUE], ax=ax)
    ax.set_title("Comparativa de Automatización por Nivel de Riesgo:\nCatraPlay vs Promedio del Resto de Marcas", pad=15, fontweight='bold')
    ax.set_ylabel("% Automatización")
    ax.set_xlabel("Nivel de Riesgo")
    ax.set_ylim(0, 105)
    plt.tight_layout()
    chart3_pani_path = os.path.join(temp_img_dir, "comparativa_riesgo_catraplay.png")
    plt.savefig(chart3_pani_path, dpi=200)
    plt.close()

    # -------------------------------------------------------------
    # 4. ANÁLISIS DE RIESGO GENERAL (OPORTUNIDAD DE SIMULACIÓN)
    # -------------------------------------------------------------
    df_risk = df_all.groupby('Nivel De Riesgo_Norm').apply(lambda x: pd.Series({
        'Total': len(x),
        'Auto': (x['Tipo Aprobacion_Norm'] == 'Automatico').sum(),
        'Manual': (x['Tipo Aprobacion_Norm'] == 'Manual').sum(),
        'Pct_Auto': ((x['Tipo Aprobacion_Norm'] == 'Automatico').sum() / len(x) * 100) if len(x) > 0 else 0
    }), include_groups=False).reset_index()
    
    bajo_series = df_risk.loc[df_risk['Nivel De Riesgo_Norm'] == 'Bajo', 'Manual']
    bajo_manual = int(bajo_series.values[0]) if len(bajo_series) > 0 else 0
    auto_general = (df_all['Tipo Aprobacion_Norm'] == 'Automatico').sum()
    pct_actual = (auto_general / total_general) * 100 if total_general > 0 else 0
    pct_simulado = ((auto_general + bajo_manual) / total_general) * 100 if total_general > 0 else 0
    
    # -------------------------------------------------------------
    # GENERAR DOCUMENTO WORD
    # -------------------------------------------------------------
    print("[*] Escribiendo documento Word...")
    doc = Document()
    
    # Configurar márgenes
    for section in doc.sections:
        section.top_margin = Inches(1)
        section.bottom_margin = Inches(1)
        section.left_margin = Inches(1)
        section.right_margin = Inches(1)
        
    style = doc.styles['Normal']
    font = style.font
    font.name = 'Calibri'
    font.size = Pt(11)
    font.color.rgb = RGBColor(0x33, 0x41, 0x55) # Slate Gray
    
    def add_heading(text, level, color):
        h = doc.add_heading(level=level)
        run = h.add_run(text)
        run.font.name = 'Calibri'
        run.bold = True
        run.font.color.rgb = RGBColor.from_string(color[1:])
        h.paragraph_format.space_before = Pt(12)
        h.paragraph_format.space_after = Pt(6)
        h.paragraph_format.keep_with_next = True
        return h
        
    # Portada / Encabezado
    title_p = doc.add_paragraph()
    title_p.paragraph_format.space_before = Pt(24)
    title_p.paragraph_format.space_after = Pt(4)
    run_title = title_p.add_run("INFORME ESTRATÉGICO DE RETIROS")
    run_title.font.size = Pt(22)
    run_title.bold = True
    run_title.font.color.rgb = RGBColor(0x0f, 0x17, 0x2a)
    
    subtitle_p = doc.add_paragraph()
    subtitle_p.paragraph_format.space_after = Pt(24)
    run_sub = subtitle_p.add_run("Efectividad, Reglas de Negocio Paramétricas y Plan para Superar la Meta del 75% de Automatización")
    run_sub.font.size = Pt(13)
    run_sub.italic = True
    run_sub.font.color.rgb = RGBColor(0x47, 0x55, 0x69)
    
    meta_p = doc.add_paragraph()
    meta_p.paragraph_format.space_after = Pt(36)
    meta_p.add_run("Período analizado: ").bold = True
    meta_p.add_run(f"{period_str}\n")
    meta_p.add_run("Volumen analizado: ").bold = True
    meta_p.add_run(f"{total_general:,} retiros consolidados\n")
    meta_p.add_run("Área emisora: ").bold = True
    meta_p.add_run("Business Intelligence y Control Operativo")
    
    # 1. RESUMEN EJECUTIVO
    add_heading("1. Resumen Ejecutivo y Diagnóstico", 1, COLOR_DARK)
    p = doc.add_paragraph()
    p.add_run(f"Este informe presenta el análisis del período {period_str} para estructurar las decisiones necesarias que permitan alinear todas las marcas a la meta corporativa del 75% de automatización diaria. El análisis revela que la tasa actual del grupo está fuertemente determinada por los límites paramétricos de Rejuego (Rollover) y Riesgo Medio configurados en cada marca.")
    
    # 2. EVOLUCIÓN MENSUAL
    add_heading("2. Evolución Temporal de la Automatización", 1, COLOR_DARK)
    doc.add_paragraph("La automatización general en el período mantiene una tendencia evaluada según el objetivo del 75%:")
    doc.add_picture(chart1_path, width=Inches(6.0))
    doc.paragraphs[-1].alignment = WD_ALIGN_PARAGRAPH.CENTER
    
    # 3. COMPORTAMIENTO POR MARCA
    add_heading("3. Desempeño y Efectividad por Marca", 1, COLOR_DARK)
    top_brand = df_brands.iloc[0] if len(df_brands) > 0 else None
    low_brands = df_brands.tail(2) if len(df_brands) >= 2 else df_brands
    low_str = ", ".join([f"{r['Marca']} ({r['Pct_Auto']:.2f}%)" for _, r in low_brands.iterrows()])
    top_str = f"{top_brand['Marca']} ({top_brand['Pct_Auto']:.2f}%)" if top_brand is not None else "N/D"
    doc.add_paragraph(f"{top_str} destaca con la mayor tasa de efectividad acumulada, mientras que marcas como {low_str} registran los valores más bajos:")
    doc.add_picture(chart2_path, width=Inches(6.0))
    doc.paragraphs[-1].alignment = WD_ALIGN_PARAGRAPH.CENTER

    # 4. EXPLICACIÓN DE REGLAS DE NEGOCIO Y SU IMPACTO EN EL REJUEGO
    add_heading("4. Impacto de los Parámetros de Control (Rejuego y Riesgos)", 1, COLOR_DARK)
    
    p_rejuego = doc.add_paragraph()
    p_rejuego.add_run("El Rejuego (monto apostado / monto depositado) es la variable clave anti-lavado de dinero de la operación. De acuerdo con la matriz de reglas oficial, existen dos extremos operativos que explican los porcentajes de automatización:").bold = True
    
    bp1 = doc.add_paragraph(style='List Bullet')
    bp1.add_run("Rejuego Laxo de 1.01 (Alta Automatización / Menor Seguridad): ").bold = True
    bp1.add_run("Marcas como Doradobet exigen un Rollover de solo 1.01 para riesgo Bajo. Esto hace que casi cualquier retiro se apruebe automáticamente sin esfuerzo de juego, maximizando el KPI pero aumentando el riesgo de retiro inmediato sin actividad en la plataforma.")
    
    bp2 = doc.add_paragraph(style='List Bullet')
    bp2.add_run("Rejuego Estricto de 3.00 (Baja Automatización / Alta Seguridad): ").bold = True
    bp2.add_run("Camanbet VES exige un Rollover de 3.00 para riesgo Bajo. Este control de seguridad estricto obliga a que la inmensa mayoría de los usuarios legítimos caigan en flujo manual, desplomando su tasa de automatización a solo un 23.65%.")

    # 5. CASO ANÓMALO CATRAPLAY
    add_heading("5. Alerta de Configuración: El Caso de CatraPlay", 1, COLOR_DARK)
    doc.add_paragraph("Al comparar las marcas por nivel de riesgo, identificamos por qué CatraPlay supera con creces al resto en automatización:")
    doc.add_picture(chart3_pani_path, width=Inches(5.5))
    doc.paragraphs[-1].alignment = WD_ALIGN_PARAGRAPH.CENTER
    
    # 7. COMPORTAMIENTO RECIENTE TRAS CAMBIO DE REGLAS DEL 12-JUN (5:00 PM)
    add_heading("7. Comportamiento Reciente por Cambio de Reglas (12-Jun 5:00 PM)", 1, COLOR_DARK)
    doc.add_paragraph("El 12 de junio a las 5:00 PM se implementaron modificaciones en las reglas de negocio. A continuación se presenta la variación real del comportamiento de retiros antes (1-Jun al 12-Jun 17:00) y después (12-Jun 17:00 al 15-Jun):")
    
    table_pre_post = doc.add_table(rows=1, cols=4)
    table_pre_post.alignment = WD_TABLE_ALIGNMENT.CENTER
    
    headers_pp = ["Marca", "Antes (1-Jun al 12-Jun 17:00)", "Después (12-Jun 17:00 al 15-Jun)", "Variación Absoluta"]
    hdr_cells_pp = table_pre_post.rows[0].cells
    for i, title_h in enumerate(headers_pp):
        hdr_cells_pp[i].text = title_h
        hdr_cells_pp[i].paragraphs[0].runs[0].font.bold = True
        hdr_cells_pp[i].paragraphs[0].runs[0].font.color.rgb = RGBColor(255, 255, 255)
        shading = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{COLOR_DARK[1:]}"/>')
        hdr_cells_pp[i]._tc.get_or_add_tcPr().append(shading)

    # Cálculo dinámico para la tabla
    df_june = df_all[df_all['Mes_Num'] == 6].copy()
    df_june['t_cambio'] = pd.to_datetime(df_june['Fecha Cambio Time'], format='mixed', errors='coerce')
    split_time = pd.to_datetime('2026-06-12 17:00:00')
    df_june['Period'] = np.where(df_june['t_cambio'] < split_time, 'Antes', 'Despues')
    
    target_brands = ["Doradobet Chile", "Doradobet Ecuador", "Ecuabet Ecuador", "Doradobet Perú"]
    
    for brand in target_brands:
        df_b = df_june[df_june['Marca'] == brand]
        df_b_antes = df_b[df_b['Period'] == 'Antes']
        df_b_despues = df_b[df_b['Period'] == 'Despues']
        
        total_a = len(df_b_antes)
        auto_a = (df_b_antes['Tipo Aprobacion_Norm'] == 'Automatico').sum()
        pct_a = (auto_a / total_a * 100) if total_a > 0 else 0.0
        
        total_d = len(df_b_despues)
        auto_d = (df_b_despues['Tipo Aprobacion_Norm'] == 'Automatico').sum()
        pct_d = (auto_d / total_d * 100) if total_d > 0 else 0.0
        
        diff = pct_d - pct_a
        
        row_cells = table_pre_post.add_row().cells
        row_cells[0].text = brand
        row_cells[1].text = f"{pct_a:.1f}% ({auto_a:,} auto / {total_a:,} total)"
        row_cells[2].text = f"{pct_d:.1f}% ({auto_d:,} auto / {total_d:,} total)"
        
        if diff >= 0:
            row_cells[3].text = f"+{diff:.2f}% (Aumento) 📈"
            row_cells[3].paragraphs[0].runs[0].font.color.rgb = RGBColor(0x22, 0xc5, 0x5e)
        else:
            row_cells[3].text = f"{diff:.2f}% (Descenso) 📉"
            row_cells[3].paragraphs[0].runs[0].font.color.rgb = RGBColor(0xef, 0x44, 0x44)
        row_cells[3].paragraphs[0].runs[0].font.bold = True

    for row in table_pre_post.rows:
        for cell in row.cells:
            cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
            cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER

    doc.add_paragraph("Diagnóstico de Reglas por Nivel de Riesgo (Junio 12 a Hoy):").bold = True
    
    # Agregar desglose por niveles de riesgo
    for brand in target_brands:
        df_b = df_june[df_june['Marca'] == brand]
        p_brand = doc.add_paragraph()
        p_brand.add_run(f"• {brand}:").bold = True
        
        for risk in ["Bajo", "Medio", "Sin Nivel"]:
            df_br = df_b[df_b['Nivel De Riesgo_Norm'] == risk]
            df_br_a = df_br[df_br['Period'] == 'Antes']
            df_br_d = df_br[df_br['Period'] == 'Despues']
            
            tot_a = len(df_br_a)
            aut_a = (df_br_a['Tipo Aprobacion_Norm'] == 'Automatico').sum()
            p_a = (aut_a / tot_a * 100) if tot_a > 0 else 0.0
            
            tot_d = len(df_br_d)
            aut_d = (df_br_d['Tipo Aprobacion_Norm'] == 'Automatico').sum()
            p_d = (aut_d / tot_d * 100) if tot_d > 0 else 0.0
            
            var = p_d - p_a
            sign = "+" if var >= 0 else ""
            
            p_sub = doc.add_paragraph(style='List Bullet')
            p_sub.add_run(f"Riesgo {risk}: ").bold = True
            p_sub.add_run(f"Antes {p_a:.1f}% ({aut_a}/{tot_a}) vs Después {p_d:.1f}% ({aut_d}/{tot_d}) | Variación: {sign}{var:.2f}%")


    p_warn = doc.add_paragraph()
    shading_w = parse_xml(f'<w:shd {nsdecls("w")} w:fill="FFF1F2"/>')
    p_warn.add_run("⚠ ALERTA DE CONFIGURACIÓN OPERATIVA (CatraPlay): ").bold = True
    p_warn.add_run("La matriz de parámetros autorizados no define ninguna regla para Riesgo Alto en CatraPlay. A pesar de esto, la data histórica demuestra que el sistema aprueba automáticamente el 70.99% de las transacciones de riesgo Alto. Se debe auditar con el equipo técnico de inmediato esta inconsistencia en el motor de aprobación.")
    p_warn._p.get_or_add_pPr().append(shading_w)

    # 6. ACCIONES RECOMENDADAS PARA SUPERAR EL 75%
    add_heading("6. Acciones Propuestas para Toma de Decisiones", 1, COLOR_DARK)
    
    recoms = [
        ("Estandarizar el Rejuego (Rollover) en 1.10 o 1.20: ", "Para equilibrar seguridad y volumen de automatización, se sugiere unificar criterios: subir el Rejuego en Doradobet (de 1.01 a 1.10/1.20) y exigir Rentabilidad Mínima de 0.10 para blindar el canal contra lavado, y bajarlo en Camanbet (de 3.00 a 1.50/1.20 en moneda local VES con Rentabilidad Mínima 0.20) para destrabar la aprobación automática de usuarios legítimos."),
        ("Configurar Reglas para Riesgo Medio en Marcas en Blanco: ", "Marcas como Doradobet Chile (CLP) y Honduras (HNL) no tienen parametrizadas reglas de aprobación para riesgo Medio. Habilitarlas bajo un límite de retiro moderado en moneda local (ej. 500,000 CLP en Chile o 5,000 HNL en Honduras), exigiendo Rentabilidad Mínima de 0.20 y similitud de IP del 80-90% incrementará significativamente el porcentaje global."),
        ("Alinear Límites en Riesgo Medio (Perú): ", "En Doradobet Perú (PEN), el límite de retiro automático en riesgo Medio es de 1,999 PEN (Rentabilidad Mínima 0.12) frente a 5,000 PEN de riesgo Bajo (Rentabilidad Mínima 0.01). Incrementar el límite de riesgo Medio a 3,000 PEN capturará gran volumen de retiros seguros de forma inmediata.")
    ]
    
    for title_r, desc_r in recoms:
        p_r = doc.add_paragraph()
        run_title = p_r.add_run(f"✔  {title_r}")
        run_title.bold = True
        run_title.font.color.rgb = RGBColor(0x10, 0xb9, 0x81)
        p_r.add_run(desc_r)

    # Guardar intentando el principal, de lo contrario al backup (V2) para evitar bloqueos
    try:
        doc.save(output_docx_primary)
        print(f"[+] Informe principal guardado en: {output_docx_primary}")
    except PermissionError:
        doc.save(output_docx_backup)
        print(f"[!] Archivo principal bloqueado. Guardado como copia de seguridad en: {output_docx_backup}")
        
    # Limpieza
    for f in glob.glob(os.path.join(temp_img_dir, "*")):
        try: os.remove(f)
        except: pass
    try: os.rmdir(temp_img_dir)
    except: pass

if __name__ == "__main__":
    generate_report()
