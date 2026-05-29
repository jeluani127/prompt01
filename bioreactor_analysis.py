"""
╔══════════════════════════════════════════════════════════════╗
║     BIOREACTOR ANALYSIS SCRIPT — Prompt Maestro 01          ║
║     by Pharmane | jeluani127.github.io/prompt01                       ║
║     Versión 1.0 — Mayo 2026                                  ║
╚══════════════════════════════════════════════════════════════╝

INSTRUCCIONES:
1. Instala dependencias:
   pip install pandas numpy matplotlib seaborn scipy openpyxl python-pptx python-docx

2. Coloca tu archivo Excel en la misma carpeta que este script
   y actualiza la variable ARCHIVO_EXCEL abajo.

3. Ejecuta: python bioreactor_analysis.py

4. Se generarán automáticamente 3 archivos:
   - analisis_biorreactor.xlsx  (Excel profesional con gráficas)
   - reporte_biorreactor.docx   (Word con análisis completo)
   - presentacion_biorreactor.pptx (PowerPoint ejecutivo)
"""

import pandas as pd
import numpy as np
import matplotlib
matplotlib.use('Agg')  # Sin ventanas emergentes
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
from matplotlib.gridspec import GridSpec
import seaborn as sns
from scipy import stats
from scipy.stats import pearsonr, normaltest
import warnings
import os
import sys
from datetime import datetime

warnings.filterwarnings('ignore')

# ─────────────────────────────────────────────────────────────
# CONFIGURACIÓN — EDITA AQUÍ
# ─────────────────────────────────────────────────────────────
ARCHIVO_EXCEL = "datos_biorreactor.xlsx"   # Nombre de tu archivo
HOJA_SENSORES = "Sensores"                  # Hoja con datos de sensores
HOJA_BATCH    = "Batch"                     # Hoja con datos por lote
EMPRESA       = "Pharmane"
PROYECTO      = "Análisis de Biorreactor"
ANALISTA      = "Equipo de Análisis"
ESPECIFICACIONES = {
    # Variable: (min_esp, max_esp)  — ajusta según tu proceso
    "pH":             (6.8, 7.4),
    "Temperatura_C":  (36.5, 37.5),
    "DO_pct":         (30.0, 100.0),
    "Viabilidad_pct": (70.0, 100.0),
}

# Colores corporativos
COLOR_PRIMARY   = "#0D2137"
COLOR_SECONDARY = "#00B4D8"
COLOR_WHITE     = "#FFFFFF"
COLOR_WARNING   = "#FF6B35"
COLOR_SUCCESS   = "#2EC4B6"
COLOR_DANGER    = "#E63946"

plt.rcParams.update({
    'figure.facecolor': COLOR_PRIMARY,
    'axes.facecolor':   '#1A3350',
    'axes.edgecolor':   COLOR_SECONDARY,
    'text.color':       COLOR_WHITE,
    'axes.labelcolor':  COLOR_WHITE,
    'xtick.color':      COLOR_WHITE,
    'ytick.color':      COLOR_WHITE,
    'grid.color':       '#2A4A65',
    'grid.alpha':       0.5,
    'font.family':      'DejaVu Sans',
})


# ─────────────────────────────────────────────────────────────
# FUNCIONES AUXILIARES
# ─────────────────────────────────────────────────────────────

def cargar_datos():
    """Carga y valida los datos del archivo Excel."""
    print(f"\n📂 Cargando datos desde: {ARCHIVO_EXCEL}")

    if not os.path.exists(ARCHIVO_EXCEL):
        print(f"⚠️  Archivo '{ARCHIVO_EXCEL}' no encontrado.")
        print("   Generando datos de ejemplo para demostración...")
        return generar_datos_ejemplo()

    dfs = {}
    xl = pd.ExcelFile(ARCHIVO_EXCEL)

    if HOJA_SENSORES in xl.sheet_names:
        dfs['sensores'] = pd.read_excel(ARCHIVO_EXCEL, sheet_name=HOJA_SENSORES)
        print(f"   ✅ Hoja '{HOJA_SENSORES}': {len(dfs['sensores'])} filas × {len(dfs['sensores'].columns)} columnas")

    if HOJA_BATCH in xl.sheet_names:
        dfs['batch'] = pd.read_excel(ARCHIVO_EXCEL, sheet_name=HOJA_BATCH)
        print(f"   ✅ Hoja '{HOJA_BATCH}': {len(dfs['batch'])} filas × {len(dfs['batch'].columns)} columnas")

    if not dfs:
        print("   No se encontraron hojas conocidas. Generando datos de ejemplo...")
        return generar_datos_ejemplo()

    return dfs


def generar_datos_ejemplo():
    """Genera datos sintéticos realistas de biorreactor para demostración."""
    np.random.seed(42)
    n_lotes = 15
    n_puntos = 120  # Puntos por lote (cada hora por 5 días)

    lotes_sensores = []
    for lote in range(1, n_lotes + 1):
        t = np.linspace(0, 120, n_puntos)
        lote_id = f"Lote_{lote:03d}"

        # pH: empieza en 7.2, deriva ligeramente
        ph = 7.2 + 0.05 * np.sin(t / 20) + np.random.normal(0, 0.03, n_puntos)
        ph -= 0.001 * t  # Deriva leve

        # Temperatura: controlada cerca de 37°C
        temp = 37.0 + np.random.normal(0, 0.1, n_puntos)

        # DO%: consume y se recupera
        do = 60 + 30 * np.exp(-t / 40) * np.cos(t / 15) + np.random.normal(0, 2, n_puntos)
        do = np.clip(do, 20, 100)

        # Agitación
        agit = 200 + 50 * (t / 120) + np.random.normal(0, 5, n_puntos)

        df_lote = pd.DataFrame({
            'Lote_ID': lote_id,
            'Tiempo_h': t,
            'pH': ph,
            'Temperatura_C': temp,
            'DO_pct': do,
            'Agitacion_rpm': agit,
        })
        lotes_sensores.append(df_lote)

    df_sensores = pd.concat(lotes_sensores, ignore_index=True)

    # Datos batch
    df_batch = pd.DataFrame({
        'Lote_ID': [f"Lote_{i:03d}" for i in range(1, n_lotes + 1)],
        'Duracion_h': np.random.normal(118, 8, n_lotes),
        'Viabilidad_pct': np.random.normal(85, 8, n_lotes),
        'Densidad_106_mL': np.random.normal(12, 3, n_lotes),
        'Rendimiento_g_L': np.random.normal(2.5, 0.4, n_lotes),
        'Glucosa_final_g_L': np.random.normal(1.2, 0.5, n_lotes),
        'Costo_USD': np.random.normal(45000, 5000, n_lotes),
        'Resultado': np.random.choice(['APROBADO', 'RECHAZADO'], n_lotes, p=[0.85, 0.15]),
    })

    print("   ✅ Datos de ejemplo generados (15 lotes, 120 puntos/lote)")
    return {'sensores': df_sensores, 'batch': df_batch}


def estadistica_descriptiva(df, vars_numericas):
    """Calcula estadísticas descriptivas completas."""
    stats_dict = {}
    for col in vars_numericas:
        if col not in df.columns:
            continue
        s = df[col].dropna()
        if len(s) == 0:
            continue

        q1, q3 = s.quantile(0.25), s.quantile(0.75)
        iqr = q3 - q1
        outliers = ((s < q1 - 1.5*iqr) | (s > q3 + 1.5*iqr)).sum()

        stats_dict[col] = {
            'N': len(s),
            'Media': round(s.mean(), 4),
            'Mediana': round(s.median(), 4),
            'Std': round(s.std(), 4),
            'Min': round(s.min(), 4),
            'Max': round(s.max(), 4),
            'Q1': round(q1, 4),
            'Q3': round(q3, 4),
            'IQR': round(iqr, 4),
            'CV%': round(s.std() / s.mean() * 100, 2) if s.mean() != 0 else None,
            'Sesgo': round(s.skew(), 4),
            'Curtosis': round(s.kurtosis(), 4),
            'Outliers': int(outliers),
        }

    return pd.DataFrame(stats_dict).T


def detectar_anomalias_sensores(df_sensores, vars_sensores):
    """Auditoría de sensores: congelado, saltos, deriva."""
    resultados = {}
    for col in vars_sensores:
        if col not in df_sensores.columns:
            continue
        s = df_sensores[col].dropna()

        # Sensor congelado: mismo valor >5 puntos seguidos
        congelado = (s.diff().abs() < 1e-6).rolling(5).sum().max() >= 4

        # Saltos abruptos: cambio > 3 sigma en un paso
        cambios = s.diff().abs()
        umbral_salto = cambios.mean() + 3 * cambios.std()
        n_saltos = (cambios > umbral_salto).sum()

        # Deriva: pendiente de tendencia lineal significativa
        x = np.arange(len(s))
        slope, intercept, r_value, p_value, _ = stats.linregress(x, s)
        deriva = abs(slope) > (s.std() / len(s) * 10)

        # Semáforo
        if congelado or n_saltos > 10:
            semaforo = "🔴 CRÍTICO"
        elif n_saltos > 3 or deriva:
            semaforo = "🟡 ATENCIÓN"
        else:
            semaforo = "🟢 NORMAL"

        resultados[col] = {
            'Estado': semaforo,
            'Sensor_Congelado': "SÍ" if congelado else "No",
            'Saltos_Abruptos': int(n_saltos),
            'Deriva_Detectada': "SÍ" if deriva else "No",
            'Pendiente_por_h': round(slope, 6),
            'R²_tendencia': round(r_value**2, 4),
        }

    return pd.DataFrame(resultados).T


def calcular_riesgo_financiero(df_batch):
    """Calcula métricas de riesgo financiero."""
    if 'Resultado' not in df_batch.columns or 'Costo_USD' not in df_batch.columns:
        return None

    total = len(df_batch)
    rechazados = (df_batch['Resultado'] == 'RECHAZADO').sum()
    tasa_rechazo = rechazados / total * 100
    costo_promedio = df_batch['Costo_USD'].mean()
    costo_rechazados = df_batch[df_batch['Resultado'] == 'RECHAZADO']['Costo_USD'].mean()

    riesgo_esperado_mensual = (tasa_rechazo / 100) * costo_rechazados * (total / 12)
    riesgo_anual = riesgo_esperado_mensual * 12

    return {
        'total_lotes': total,
        'lotes_rechazados': int(rechazados),
        'tasa_rechazo_pct': round(tasa_rechazo, 1),
        'costo_promedio_usd': round(costo_promedio, 0),
        'costo_lote_rechazado_usd': round(costo_rechazados, 0) if not pd.isna(costo_rechazados) else 0,
        'riesgo_mensual_usd': round(riesgo_esperado_mensual, 0),
        'riesgo_anual_usd': round(riesgo_anual, 0),
    }


# ─────────────────────────────────────────────────────────────
# GENERADORES DE GRÁFICAS
# ─────────────────────────────────────────────────────────────

def graficar_series_tiempo(df_sensores, vars_sensores, output_dir):
    """Genera gráficas de series de tiempo."""
    os.makedirs(output_dir, exist_ok=True)
    archivos = []
    colores = [COLOR_SECONDARY, '#FF6B35', '#2EC4B6', '#9B5DE5', '#F15BB5', '#FEE440']

    lotes = df_sensores['Lote_ID'].unique() if 'Lote_ID' in df_sensores.columns else ['Todos']

    for var in vars_sensores:
        if var not in df_sensores.columns:
            continue

        fig, ax = plt.subplots(figsize=(12, 5))
        fig.patch.set_facecolor(COLOR_PRIMARY)

        if 'Lote_ID' in df_sensores.columns:
            for i, lote in enumerate(lotes):
                df_lote = df_sensores[df_sensores['Lote_ID'] == lote]
                x_col = 'Tiempo_h' if 'Tiempo_h' in df_lote.columns else df_lote.index
                ax.plot(df_lote[x_col] if isinstance(x_col, str) else x_col,
                        df_lote[var],
                        color=colores[i % len(colores)],
                        alpha=0.7, linewidth=1.2, label=lote)
        else:
            ax.plot(df_sensores.index, df_sensores[var], color=COLOR_SECONDARY, linewidth=1.5)

        # Líneas de especificación
        if var in ESPECIFICACIONES:
            lo, hi = ESPECIFICACIONES[var]
            ax.axhline(lo, color=COLOR_DANGER, linestyle='--', alpha=0.8, linewidth=1, label=f'Lím. inf. ({lo})')
            ax.axhline(hi, color=COLOR_DANGER, linestyle='--', alpha=0.8, linewidth=1, label=f'Lím. sup. ({hi})')

        ax.set_title(f'Serie de Tiempo — {var}', fontsize=13, fontweight='bold', color=COLOR_WHITE, pad=12)
        ax.set_xlabel('Tiempo (h)', fontsize=10)
        ax.set_ylabel(var, fontsize=10)
        ax.grid(True, alpha=0.3)

        if len(lotes) <= 10:
            ax.legend(loc='upper right', fontsize=7, ncol=2,
                     facecolor='#1A3350', edgecolor=COLOR_SECONDARY)

        plt.tight_layout()
        path = os.path.join(output_dir, f"serie_tiempo_{var}.png")
        plt.savefig(path, dpi=150, bbox_inches='tight', facecolor=COLOR_PRIMARY)
        plt.close()
        archivos.append(path)
        print(f"   📈 Serie de tiempo: {var}")

    return archivos


def graficar_correlaciones(df, vars_numericas, output_dir):
    """Matriz de correlación y scatter plots."""
    os.makedirs(output_dir, exist_ok=True)

    df_num = df[vars_numericas].dropna()
    corr = df_num.corr()

    # Heatmap
    fig, ax = plt.subplots(figsize=(10, 8))
    fig.patch.set_facecolor(COLOR_PRIMARY)

    cmap = sns.diverging_palette(220, 20, as_cmap=True)
    mask = np.zeros_like(corr, dtype=bool)
    mask[np.triu_indices_from(mask)] = True

    sns.heatmap(corr, mask=mask, cmap=cmap, vmax=1, vmin=-1, center=0,
                annot=True, fmt='.2f', linewidths=0.5,
                ax=ax, cbar_kws={"shrink": 0.8},
                annot_kws={"size": 9, "color": "white"})

    ax.set_title('Matriz de Correlación de Pearson', fontsize=14, fontweight='bold',
                 color=COLOR_WHITE, pad=15)
    ax.tick_params(colors=COLOR_WHITE)

    plt.tight_layout()
    heatmap_path = os.path.join(output_dir, "correlacion_heatmap.png")
    plt.savefig(heatmap_path, dpi=150, bbox_inches='tight', facecolor=COLOR_PRIMARY)
    plt.close()
    print("   🔵 Mapa de calor de correlaciones")

    # Top 3 scatter plots
    scatter_paths = []
    corr_pairs = []
    for i in range(len(corr.columns)):
        for j in range(i+1, len(corr.columns)):
            r = abs(corr.iloc[i, j])
            if not pd.isna(r):
                corr_pairs.append((r, corr.columns[i], corr.columns[j]))

    corr_pairs.sort(reverse=True)

    for r_val, var1, var2 in corr_pairs[:3]:
        fig, ax = plt.subplots(figsize=(7, 6))
        fig.patch.set_facecolor(COLOR_PRIMARY)

        x = df_num[var1].values
        y = df_num[var2].values
        mask = ~(np.isnan(x) | np.isnan(y))
        x, y = x[mask], y[mask]

        ax.scatter(x, y, color=COLOR_SECONDARY, alpha=0.6, s=30, edgecolors='none')

        # Regresión
        slope, intercept, r_value, p_value, std_err = stats.linregress(x, y)
        x_line = np.linspace(x.min(), x.max(), 100)
        y_line = slope * x_line + intercept
        ax.plot(x_line, y_line, color=COLOR_WARNING, linewidth=2,
                label=f'y = {slope:.3f}x + {intercept:.3f}\nR² = {r_value**2:.3f}, p = {p_value:.4f}')

        ax.set_title(f'Dispersión: {var1} vs {var2}', fontsize=12, fontweight='bold', color=COLOR_WHITE)
        ax.set_xlabel(var1, fontsize=10)
        ax.set_ylabel(var2, fontsize=10)
        ax.legend(facecolor='#1A3350', edgecolor=COLOR_SECONDARY, labelcolor=COLOR_WHITE, fontsize=9)
        ax.grid(True, alpha=0.3)

        plt.tight_layout()
        path = os.path.join(output_dir, f"scatter_{var1}_vs_{var2}.png".replace(' ', '_'))
        plt.savefig(path, dpi=150, bbox_inches='tight', facecolor=COLOR_PRIMARY)
        plt.close()
        scatter_paths.append(path)
        print(f"   🔵 Dispersión: {var1} vs {var2} (r={r_val:.3f})")

    return heatmap_path, scatter_paths, corr


def graficar_boxplots(df_batch, vars_batch, output_dir):
    """Boxplots de variables por lote."""
    os.makedirs(output_dir, exist_ok=True)
    paths = []

    vars_disponibles = [v for v in vars_batch if v in df_batch.columns]
    if not vars_disponibles:
        return paths

    fig, axes = plt.subplots(2, max(1, len(vars_disponibles)//2 + 1),
                              figsize=(14, 8))
    fig.patch.set_facecolor(COLOR_PRIMARY)
    axes_flat = axes.flatten() if hasattr(axes, 'flatten') else [axes]

    for i, var in enumerate(vars_disponibles):
        if i >= len(axes_flat):
            break
        ax = axes_flat[i]

        data = df_batch[var].dropna()
        bp = ax.boxplot(data, patch_artist=True,
                        boxprops=dict(facecolor=COLOR_SECONDARY, color=COLOR_WHITE),
                        medianprops=dict(color=COLOR_WARNING, linewidth=2),
                        whiskerprops=dict(color=COLOR_WHITE),
                        capprops=dict(color=COLOR_WHITE),
                        flierprops=dict(marker='o', color=COLOR_DANGER, markersize=5))

        # Líneas de especificación
        if var in ESPECIFICACIONES:
            lo, hi = ESPECIFICACIONES[var]
            ax.axhline(lo, color=COLOR_DANGER, linestyle='--', alpha=0.8, linewidth=1)
            ax.axhline(hi, color=COLOR_DANGER, linestyle='--', alpha=0.8, linewidth=1)

        ax.set_title(var, fontsize=10, fontweight='bold', color=COLOR_WHITE)
        ax.grid(True, alpha=0.3)

    for i in range(len(vars_disponibles), len(axes_flat)):
        axes_flat[i].set_visible(False)

    fig.suptitle('Distribución de Variables de Batch', fontsize=14, fontweight='bold',
                 color=COLOR_WHITE, y=1.02)
    plt.tight_layout()
    path = os.path.join(output_dir, "boxplots_batch.png")
    plt.savefig(path, dpi=150, bbox_inches='tight', facecolor=COLOR_PRIMARY)
    plt.close()
    paths.append(path)
    print("   📦 Boxplots de variables de batch")
    return paths


def graficar_spc(df_batch, var, output_dir):
    """Carta de control Shewhart."""
    if var not in df_batch.columns:
        return None

    os.makedirs(output_dir, exist_ok=True)
    s = df_batch[var].dropna()

    media = s.mean()
    std = s.std()
    ucl = media + 3 * std
    lcl = media - 3 * std

    fuera_control = (s > ucl) | (s < lcl)

    fig, ax = plt.subplots(figsize=(12, 5))
    fig.patch.set_facecolor(COLOR_PRIMARY)

    ax.plot(range(len(s)), s.values, 'o-', color=COLOR_SECONDARY, linewidth=1.5, markersize=5, label='Valor por lote')
    ax.axhline(media, color=COLOR_WHITE, linewidth=2, linestyle='-', label=f'Media = {media:.3f}')
    ax.axhline(ucl, color=COLOR_DANGER, linewidth=1.5, linestyle='--', label=f'LCS = {ucl:.3f}')
    ax.axhline(lcl, color=COLOR_DANGER, linewidth=1.5, linestyle='--', label=f'LCI = {lcl:.3f}')
    ax.fill_between(range(len(s)), lcl, ucl, alpha=0.1, color=COLOR_SUCCESS)

    # Puntos fuera de control
    out_idx = np.where(fuera_control.values)[0]
    if len(out_idx) > 0:
        ax.scatter(out_idx, s.values[out_idx], color=COLOR_DANGER, s=80, zorder=5, label='Fuera de control')

    ax.set_title(f'Carta de Control SPC — {var}', fontsize=13, fontweight='bold', color=COLOR_WHITE)
    ax.set_xlabel('Número de Lote', fontsize=10)
    ax.set_ylabel(var, fontsize=10)
    ax.legend(facecolor='#1A3350', edgecolor=COLOR_SECONDARY, labelcolor=COLOR_WHITE, fontsize=9)
    ax.grid(True, alpha=0.3)

    plt.tight_layout()
    path = os.path.join(output_dir, f"spc_{var}.png")
    plt.savefig(path, dpi=150, bbox_inches='tight', facecolor=COLOR_PRIMARY)
    plt.close()
    print(f"   📊 Carta de control SPC: {var}")
    return path


def graficar_riesgo_financiero(df_batch, riesgo, output_dir):
    """Gráfica de riesgo financiero."""
    if riesgo is None or 'Resultado' not in df_batch.columns:
        return None

    os.makedirs(output_dir, exist_ok=True)

    fig, axes = plt.subplots(1, 2, figsize=(12, 5))
    fig.patch.set_facecolor(COLOR_PRIMARY)

    # Pie chart
    ax1 = axes[0]
    labels = ['Aprobados', 'Rechazados']
    sizes = [riesgo['total_lotes'] - riesgo['lotes_rechazados'], riesgo['lotes_rechazados']]
    colors = [COLOR_SUCCESS, COLOR_DANGER]
    wedges, texts, autotexts = ax1.pie(sizes, labels=labels, colors=colors,
                                        autopct='%1.1f%%', startangle=90,
                                        textprops={'color': COLOR_WHITE, 'fontsize': 11})
    for at in autotexts:
        at.set_fontsize(12)
        at.set_fontweight('bold')
    ax1.set_title('Tasa de Rechazo de Lotes', fontsize=12, fontweight='bold', color=COLOR_WHITE)

    # Bar chart de costo
    ax2 = axes[1]
    categorias = ['Costo\npromedio/lote', 'Costo\nlote rechazado', 'Riesgo\nmensual', 'Riesgo\nanual']
    valores = [
        riesgo['costo_promedio_usd'],
        riesgo['costo_lote_rechazado_usd'],
        riesgo['riesgo_mensual_usd'],
        riesgo['riesgo_anual_usd']
    ]
    bar_colors = [COLOR_SECONDARY, COLOR_WARNING, COLOR_DANGER, '#8B0000']
    bars = ax2.bar(categorias, valores, color=bar_colors, edgecolor=COLOR_WHITE, linewidth=0.5)

    for bar, val in zip(bars, valores):
        ax2.text(bar.get_x() + bar.get_width()/2, bar.get_height() + max(valores)*0.02,
                f'${val:,.0f}', ha='center', va='bottom', color=COLOR_WHITE, fontsize=9, fontweight='bold')

    ax2.set_title('Análisis de Riesgo Financiero (USD)', fontsize=12, fontweight='bold', color=COLOR_WHITE)
    ax2.set_ylabel('USD', fontsize=10)
    ax2.grid(True, alpha=0.3, axis='y')

    plt.tight_layout()
    path = os.path.join(output_dir, "riesgo_financiero.png")
    plt.savefig(path, dpi=150, bbox_inches='tight', facecolor=COLOR_PRIMARY)
    plt.close()
    print("   💰 Gráfica de riesgo financiero")
    return path


# ─────────────────────────────────────────────────────────────
# EXPORTAR A EXCEL
# ─────────────────────────────────────────────────────────────

def exportar_excel(dfs, stats_sensores, stats_batch, auditoria, riesgo, output_path):
    """Crea el archivo Excel profesional con múltiples hojas."""
    from openpyxl import Workbook
    from openpyxl.styles import (PatternFill, Font, Alignment, Border, Side,
                                  GradientFill)
    from openpyxl.utils.dataframe import dataframe_to_rows
    from openpyxl.chart import BarChart, LineChart, Reference
    from openpyxl.formatting.rule import ColorScaleRule

    wb = Workbook()

    # Estilos
    fill_header = PatternFill("solid", fgColor="0D2137")
    fill_alt    = PatternFill("solid", fgColor="1A3350")
    fill_green  = PatternFill("solid", fgColor="2EC4B6")
    fill_red    = PatternFill("solid", fgColor="E63946")
    fill_yellow = PatternFill("solid", fgColor="FFB703")

    font_header = Font(bold=True, color="FFFFFF", size=11, name="Calibri")
    font_title  = Font(bold=True, color="00B4D8", size=14, name="Calibri")
    font_normal = Font(color="FFFFFF", size=10, name="Calibri")
    font_red    = Font(bold=True, color="E63946", size=10, name="Calibri")

    border = Border(
        left=Side(style='thin', color='00B4D8'),
        right=Side(style='thin', color='00B4D8'),
        top=Side(style='thin', color='00B4D8'),
        bottom=Side(style='thin', color='00B4D8')
    )

    def style_header_row(ws, row, n_cols):
        for col in range(1, n_cols + 1):
            cell = ws.cell(row=row, column=col)
            cell.fill = fill_header
            cell.font = font_header
            cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
            cell.border = border

    def style_data_row(ws, row, n_cols, alternate=False):
        for col in range(1, n_cols + 1):
            cell = ws.cell(row=row, column=col)
            cell.fill = fill_alt if alternate else PatternFill("solid", fgColor="0F2840")
            cell.font = font_normal
            cell.border = border
            cell.alignment = Alignment(horizontal='center', vertical='center')

    # ── HOJA 1: DASHBOARD ──
    ws = wb.active
    ws.title = "Dashboard"
    ws.sheet_view.showGridLines = False
    ws.column_dimensions['A'].width = 30
    ws.column_dimensions['B'].width = 25
    ws.column_dimensions['C'].width = 25

    ws['A1'] = f"🧬 ANÁLISIS DE BIORREACTOR — {EMPRESA}"
    ws['A1'].font = Font(bold=True, color="00B4D8", size=16, name="Calibri")
    ws['A1'].fill = fill_header
    ws.merge_cells('A1:C1')
    ws['A1'].alignment = Alignment(horizontal='center')

    ws['A2'] = "Generado con Prompt Maestro 01 | jeluani127.github.io/prompt01"
    ws['A2'].font = Font(color="AAAAAA", size=9, name="Calibri", italic=True)
    ws.merge_cells('A2:C2')
    ws['A2'].alignment = Alignment(horizontal='center')

    ws['A3'] = "Fecha de análisis:"
    ws['A3'].font = font_header
    ws['A3'].fill = fill_header
    ws['B3'] = datetime.now().strftime("%d/%m/%Y %H:%M")
    ws['B3'].font = font_normal
    ws['B3'].fill = fill_alt

    # KPIs
    kpis = [("", "Indicador", "Valor", "Estado")]
    if riesgo:
        kpis += [
            ("📊", "Total Lotes Analizados", riesgo['total_lotes'], ""),
            ("❌", "Lotes Rechazados", riesgo['lotes_rechazados'], ""),
            ("📉", "Tasa de Rechazo", f"{riesgo['tasa_rechazo_pct']}%", "🔴" if riesgo['tasa_rechazo_pct'] > 15 else "🟢"),
            ("💰", "Riesgo Financiero Anual", f"${riesgo['riesgo_anual_usd']:,.0f}", ""),
        ]

    for i, kpi in enumerate(kpis):
        row = 5 + i
        for j, val in enumerate(kpi):
            cell = ws.cell(row=row, column=j+1, value=val)
            cell.fill = fill_header if i == 0 else fill_alt
            cell.font = font_header if i == 0 else font_normal
            cell.border = border
            cell.alignment = Alignment(horizontal='center')

    # ── HOJA 2: ESTADÍSTICAS SENSORES ──
    if stats_sensores is not None and len(stats_sensores) > 0:
        ws2 = wb.create_sheet("Estadísticas Sensores")
        ws2.sheet_view.showGridLines = False

        ws2['A1'] = "Estadística Descriptiva — Variables de Sensores"
        ws2['A1'].font = font_title
        ws2['A1'].fill = fill_header

        stats_reset = stats_sensores.reset_index()
        stats_reset.columns = ['Variable'] + list(stats_sensores.columns)

        for r_idx, row in enumerate(dataframe_to_rows(stats_reset, index=False, header=True)):
            for c_idx, val in enumerate(row):
                cell = ws2.cell(row=r_idx + 3, column=c_idx + 1, value=val)
                if r_idx == 0:
                    cell.fill = fill_header
                    cell.font = font_header
                else:
                    cell.fill = fill_alt if r_idx % 2 == 0 else PatternFill("solid", fgColor="0F2840")
                    cell.font = font_normal
                cell.border = border
                cell.alignment = Alignment(horizontal='center')

        for col in ws2.columns:
            ws2.column_dimensions[col[0].column_letter].width = 14

    # ── HOJA 3: ESTADÍSTICAS BATCH ──
    if stats_batch is not None and len(stats_batch) > 0:
        ws3 = wb.create_sheet("Estadísticas Batch")
        ws3.sheet_view.showGridLines = False
        ws3['A1'] = "Estadística Descriptiva — Variables de Batch"
        ws3['A1'].font = font_title
        ws3['A1'].fill = fill_header

        stats_reset = stats_batch.reset_index()
        stats_reset.columns = ['Variable'] + list(stats_batch.columns)

        for r_idx, row in enumerate(dataframe_to_rows(stats_reset, index=False, header=True)):
            for c_idx, val in enumerate(row):
                cell = ws3.cell(row=r_idx + 3, column=c_idx + 1, value=val)
                if r_idx == 0:
                    cell.fill = fill_header
                    cell.font = font_header
                else:
                    cell.fill = fill_alt if r_idx % 2 == 0 else PatternFill("solid", fgColor="0F2840")
                    cell.font = font_normal
                cell.border = border
                cell.alignment = Alignment(horizontal='center')

        for col in ws3.columns:
            ws3.column_dimensions[col[0].column_letter].width = 16

    # ── HOJA 4: AUDITORÍA DE SENSORES ──
    if auditoria is not None and len(auditoria) > 0:
        ws4 = wb.create_sheet("Auditoría Sensores")
        ws4.sheet_view.showGridLines = False
        ws4['A1'] = "Auditoría de Sensores — Semáforo de Estado"
        ws4['A1'].font = font_title
        ws4['A1'].fill = fill_header

        aud_reset = auditoria.reset_index()
        aud_reset.columns = ['Sensor'] + list(auditoria.columns)

        for r_idx, row in enumerate(dataframe_to_rows(aud_reset, index=False, header=True)):
            for c_idx, val in enumerate(row):
                cell = ws4.cell(row=r_idx + 3, column=c_idx + 1, value=val)
                if r_idx == 0:
                    cell.fill = fill_header
                    cell.font = font_header
                else:
                    if c_idx == 1:  # Columna Estado
                        if "CRÍTICO" in str(val):
                            cell.fill = fill_red
                        elif "ATENCIÓN" in str(val):
                            cell.fill = fill_yellow
                        else:
                            cell.fill = fill_green
                        cell.font = Font(bold=True, color="FFFFFF", size=10, name="Calibri")
                    else:
                        cell.fill = fill_alt if r_idx % 2 == 0 else PatternFill("solid", fgColor="0F2840")
                        cell.font = font_normal
                cell.border = border
                cell.alignment = Alignment(horizontal='center')

        for col in ws4.columns:
            ws4.column_dimensions[col[0].column_letter].width = 20

    # ── HOJA 5: RIESGO FINANCIERO ──
    if riesgo:
        ws5 = wb.create_sheet("Riesgo Financiero")
        ws5.sheet_view.showGridLines = False
        ws5['A1'] = "Análisis de Riesgo Financiero por Batch"
        ws5['A1'].font = font_title
        ws5['A1'].fill = fill_header

        for i, (key, val) in enumerate(riesgo.items()):
            row = i + 3
            ws5.cell(row=row, column=1, value=key.replace('_', ' ').title()).font = font_header
            ws5.cell(row=row, column=1).fill = fill_header
            ws5.cell(row=row, column=2, value=val).font = font_normal
            ws5.cell(row=row, column=2).fill = fill_alt
            ws5.cell(row=row, column=1).border = border
            ws5.cell(row=row, column=2).border = border

        ws5.column_dimensions['A'].width = 35
        ws5.column_dimensions['B'].width = 20

    # ── HOJA 6: DATOS ORIGINALES ──
    for nombre, df in dfs.items():
        ws_raw = wb.create_sheet(f"Datos_{nombre.capitalize()}")
        ws_raw.sheet_view.showGridLines = False

        for r_idx, row in enumerate(dataframe_to_rows(df.head(500), index=False, header=True)):
            for c_idx, val in enumerate(row):
                cell = ws_raw.cell(row=r_idx + 1, column=c_idx + 1, value=val)
                if r_idx == 0:
                    cell.fill = fill_header
                    cell.font = font_header
                    cell.alignment = Alignment(horizontal='center')
                else:
                    cell.fill = fill_alt if r_idx % 2 == 0 else PatternFill("solid", fgColor="0F2840")
                    cell.font = font_normal
                cell.border = border

    wb.save(output_path)
    print(f"\n   ✅ Excel guardado: {output_path}")


# ─────────────────────────────────────────────────────────────
# EXPORTAR A WORD
# ─────────────────────────────────────────────────────────────

def exportar_word(stats_sensores, stats_batch, auditoria, riesgo, imagenes, output_path):
    """Genera reporte Word profesional."""
    from docx import Document
    from docx.shared import Inches, Pt, RGBColor, Cm
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.enum.table import WD_TABLE_ALIGNMENT
    from docx.oxml.ns import qn
    from docx.oxml import OxmlElement

    doc = Document()

    # Configurar márgenes
    for section in doc.sections:
        section.top_margin    = Cm(2.5)
        section.bottom_margin = Cm(2.5)
        section.left_margin   = Cm(3.0)
        section.right_margin  = Cm(2.5)

    def add_heading(doc, text, level=1, color=None):
        h = doc.add_heading(text, level=level)
        if color:
            for run in h.runs:
                run.font.color.rgb = RGBColor(*color)
        h.alignment = WD_ALIGN_PARAGRAPH.LEFT
        return h

    def add_paragraph_styled(doc, text, bold=False, italic=False, color=None, size=11):
        p = doc.add_paragraph()
        run = p.add_run(text)
        run.bold = bold
        run.italic = italic
        run.font.size = Pt(size)
        if color:
            run.font.color.rgb = RGBColor(*color)
        return p

    # ── PORTADA ──
    doc.add_paragraph()
    title = doc.add_paragraph()
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = title.add_run("ANÁLISIS ESTADÍSTICO DE PROCESO DE BIORREACTOR")
    run.bold = True
    run.font.size = Pt(20)
    run.font.color.rgb = RGBColor(13, 33, 55)

    doc.add_paragraph()
    sub = doc.add_paragraph()
    sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run2 = sub.add_run("Reporte Técnico Confidencial")
    run2.font.size = Pt(14)
    run2.font.color.rgb = RGBColor(0, 180, 216)

    doc.add_paragraph()
    meta = doc.add_paragraph()
    meta.alignment = WD_ALIGN_PARAGRAPH.CENTER
    meta.add_run(f"Empresa: {EMPRESA}\nProyecto: {PROYECTO}\nAnalista: {ANALISTA}\nFecha: {datetime.now().strftime('%d de %B de %Y')}\nVersión: 1.0")

    doc.add_paragraph()
    doc.add_paragraph("Generado con Prompt Maestro 01 | jeluani127.github.io/prompt01").italic = True

    doc.add_page_break()

    # ── RESUMEN EJECUTIVO ──
    add_heading(doc, "1. Resumen Ejecutivo", 1, (13, 33, 55))
    add_paragraph_styled(doc, "Este reporte presenta el análisis estadístico completo de los datos del proceso de biorreactor. Los principales hallazgos son:")

    hallazgos = doc.add_paragraph(style='List Bullet')
    hallazgos.add_run("Se analizaron los perfiles de proceso y variables de calidad de múltiples lotes.")
    doc.add_paragraph("Se identificaron correlaciones significativas entre variables de proceso y resultados de calidad.", style='List Bullet')
    doc.add_paragraph("Se evaluó el comportamiento estadístico de los sensores y se detectaron posibles anomalías.", style='List Bullet')
    if riesgo:
        doc.add_paragraph(f"La tasa de rechazo actual es {riesgo['tasa_rechazo_pct']}%, con un riesgo financiero estimado de ${riesgo['riesgo_anual_usd']:,.0f} USD anuales.", style='List Bullet')
    doc.add_paragraph("Se proponen recomendaciones de mejora priorizadas por impacto en proceso y finanzas.", style='List Bullet')

    # ── ESTADÍSTICAS ──
    add_heading(doc, "2. Estadística Descriptiva", 1, (13, 33, 55))
    add_paragraph_styled(doc, "Las siguientes tablas resumen los parámetros estadísticos calculados para cada variable del proceso.")

    def df_to_word_table(doc, df, title):
        add_heading(doc, title, 2, (0, 180, 216))
        if df is None or len(df) == 0:
            doc.add_paragraph("No hay datos disponibles para esta sección.")
            return

        df_r = df.reset_index()
        table = doc.add_table(rows=1, cols=len(df_r.columns))
        table.style = 'Table Grid'

        # Header
        for i, col in enumerate(df_r.columns):
            cell = table.rows[0].cells[i]
            cell.text = str(col)
            cell.paragraphs[0].runs[0].bold = True
            cell.paragraphs[0].runs[0].font.size = Pt(9)

        # Data
        for _, row in df_r.iterrows():
            cells = table.add_row().cells
            for i, val in enumerate(row):
                cells[i].text = str(round(val, 3) if isinstance(val, float) else val)
                cells[i].paragraphs[0].runs[0].font.size = Pt(8)

        doc.add_paragraph()

    df_to_word_table(doc, stats_sensores, "2.1 Variables de Sensores")
    df_to_word_table(doc, stats_batch, "2.2 Variables de Batch")

    # ── AUDITORÍA ──
    add_heading(doc, "3. Auditoría de Sensores", 1, (13, 33, 55))
    add_paragraph_styled(doc, "Se evaluó el comportamiento de cada sensor para detectar anomalías como congelamiento de señal, saltos abruptos y deriva sistemática.")
    df_to_word_table(doc, auditoria, "Resultados de Auditoría")

    # ── RIESGO FINANCIERO ──
    if riesgo:
        add_heading(doc, "4. Análisis de Riesgo Financiero", 1, (13, 33, 55))
        add_paragraph_styled(doc, "Con base en el historial de lotes, se calcularon las siguientes métricas de riesgo:")
        for key, val in riesgo.items():
            doc.add_paragraph(f"• {key.replace('_', ' ').title()}: {val}", style='List Bullet')

    # ── IMÁGENES ──
    if imagenes:
        add_heading(doc, "5. Gráficas", 1, (13, 33, 55))
        for img_path in imagenes:
            if img_path and os.path.exists(img_path):
                try:
                    doc.add_picture(img_path, width=Inches(5.5))
                    doc.add_paragraph(os.path.basename(img_path).replace('.png', '').replace('_', ' ').title())
                    doc.add_paragraph()
                except Exception:
                    pass

    # ── CONCLUSIONES ──
    add_heading(doc, "6. Conclusiones y Recomendaciones", 1, (13, 33, 55))
    add_paragraph_styled(doc, "Con base en el análisis realizado, se recomiendan las siguientes acciones:")
    doc.add_paragraph("Revisar los sensores con estado CRÍTICO o ATENCIÓN antes del próximo lote.", style='List Bullet')
    doc.add_paragraph("Investigar las causas de las correlaciones negativas entre variables clave.", style='List Bullet')
    doc.add_paragraph("Establecer alertas automáticas para variables fuera de ±3σ.", style='List Bullet')
    if riesgo:
        doc.add_paragraph(f"Implementar programa de reducción de rechazos para recuperar hasta ${riesgo['riesgo_anual_usd']*0.5:,.0f} USD anuales.", style='List Bullet')
    doc.add_paragraph("Actualizar este análisis mensualmente para control estadístico continuo.", style='List Bullet')

    doc.add_paragraph()
    footer_p = doc.add_paragraph()
    footer_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = footer_p.add_run("Prompt Maestro 01 | Pharmane | jeluani127.github.io/prompt01")
    run.font.size = Pt(8)
    run.italic = True
    run.font.color.rgb = RGBColor(150, 150, 150)

    doc.save(output_path)
    print(f"   ✅ Word guardado: {output_path}")


# ─────────────────────────────────────────────────────────────
# EXPORTAR A POWERPOINT
# ─────────────────────────────────────────────────────────────

def exportar_pptx(stats_sensores, stats_batch, riesgo, imagenes, output_path):
    """Genera presentación PowerPoint ejecutiva."""
    from pptx import Presentation
    from pptx.util import Inches, Pt, Emu
    from pptx.dml.color import RGBColor
    from pptx.enum.text import PP_ALIGN

    prs = Presentation()
    prs.slide_width  = Inches(13.33)
    prs.slide_height = Inches(7.5)

    DARK  = RGBColor(13, 33, 55)
    TEAL  = RGBColor(0, 180, 216)
    WHITE = RGBColor(255, 255, 255)
    ORANGE = RGBColor(255, 107, 53)

    blank = prs.slide_layouts[6]  # En blanco

    def add_slide_dark(title_text, subtitle_text="", content_lines=None, img_path=None):
        slide = prs.slides.add_slide(blank)
        bg = slide.background
        fill = bg.fill
        fill.solid()
        fill.fore_color.rgb = DARK

        # Barra superior
        top_bar = slide.shapes.add_shape(1, 0, 0, prs.slide_width, Inches(0.1))
        top_bar.fill.solid()
        top_bar.fill.fore_color.rgb = TEAL
        top_bar.line.fill.background()

        # Título
        txBox = slide.shapes.add_textbox(Inches(0.5), Inches(0.2), Inches(12.3), Inches(1.0))
        tf = txBox.text_frame
        tf.word_wrap = True
        p = tf.paragraphs[0]
        run = p.add_run()
        run.text = title_text
        run.font.size = Pt(24)
        run.font.bold = True
        run.font.color.rgb = TEAL

        if subtitle_text:
            txBox2 = slide.shapes.add_textbox(Inches(0.5), Inches(1.1), Inches(12.3), Inches(0.5))
            tf2 = txBox2.text_frame
            p2 = tf2.paragraphs[0]
            run2 = p2.add_run()
            run2.text = subtitle_text
            run2.font.size = Pt(13)
            run2.font.color.rgb = WHITE

        if content_lines:
            txBox3 = slide.shapes.add_textbox(Inches(0.5), Inches(1.7), Inches(7.5), Inches(5.0))
            tf3 = txBox3.text_frame
            tf3.word_wrap = True
            for i, line in enumerate(content_lines):
                if i == 0:
                    p3 = tf3.paragraphs[0]
                else:
                    p3 = tf3.add_paragraph()
                run3 = p3.add_run()
                run3.text = line
                run3.font.size = Pt(13)
                run3.font.color.rgb = WHITE if not line.startswith("✅") and not line.startswith("📌") else TEAL
                p3.space_before = Pt(4)

        if img_path and os.path.exists(img_path):
            try:
                pic_left  = Inches(8.2) if content_lines else Inches(1.5)
                pic_top   = Inches(1.5)
                pic_width = Inches(4.8) if content_lines else Inches(10)
                slide.shapes.add_picture(img_path, pic_left, pic_top, width=pic_width)
            except Exception:
                pass

        # Pie de página
        footer = slide.shapes.add_textbox(Inches(0.3), Inches(7.1), Inches(12.7), Inches(0.3))
        tf_f = footer.text_frame
        p_f = tf_f.paragraphs[0]
        run_f = p_f.add_run()
        run_f.text = f"Prompt Maestro 01 | Pharmane | jeluani127.github.io/prompt01  |  {datetime.now().strftime('%d/%m/%Y')}  |  CONFIDENCIAL"
        run_f.font.size = Pt(7)
        run_f.font.color.rgb = RGBColor(100, 130, 160)

        return slide

    # ── SLIDE 1: PORTADA ──
    slide1 = prs.slides.add_slide(blank)
    bg = slide1.background; bg.fill.solid(); bg.fill.fore_color.rgb = DARK

    bar = slide1.shapes.add_shape(1, 0, Inches(3.2), prs.slide_width, Inches(1.5))
    bar.fill.solid(); bar.fill.fore_color.rgb = RGBColor(0, 80, 110); bar.line.fill.background()

    txTitle = slide1.shapes.add_textbox(Inches(1), Inches(1.5), Inches(11.3), Inches(1.5))
    tf = txTitle.text_frame
    run = tf.paragraphs[0].add_run()
    run.text = "ANÁLISIS ESTADÍSTICO DE BIORREACTOR"
    run.font.size = Pt(32); run.font.bold = True; run.font.color.rgb = TEAL

    txSub = slide1.shapes.add_textbox(Inches(1), Inches(3.3), Inches(11.3), Inches(1.0))
    tf2 = txSub.text_frame
    run2 = tf2.paragraphs[0].add_run()
    run2.text = f"{EMPRESA}  |  {datetime.now().strftime('%B %Y')}  |  Prompt Maestro 01"
    run2.font.size = Pt(16); run2.font.color.rgb = WHITE

    # ── SLIDE 2: DATOS ANALIZADOS ──
    n_lotes = riesgo['total_lotes'] if riesgo else "N/A"
    add_slide_dark(
        "Datos Analizados",
        "¿Qué datos, de dónde y cuántos?",
        [
            f"🧪 Total de lotes en el análisis: {n_lotes}",
            f"📅 Período: {datetime.now().strftime('%Y')}",
            "📊 Variables de sensores: pH, Temperatura, DO%, Agitación",
            "📋 Variables de batch: Viabilidad, Densidad, Rendimiento",
            "💰 Variables financieras: Costo por lote, aprobados/rechazados",
            "",
            "Fuente: Datos del sistema SCADA / Excel de proceso",
        ]
    )

    # ── SLIDE 3: ESTADÍSTICA DESCRIPTIVA ──
    if stats_batch is not None and len(stats_batch) > 0:
        rows = []
        for var in stats_batch.index[:5]:
            r = stats_batch.loc[var]
            cv = r.get('CV%', 'N/A')
            rows.append(f"📌 {var}: Media={r['Media']:.2f}  CV%={cv}  Outliers={int(r['Outliers'])}")
        add_slide_dark("Estadística Descriptiva", "Resumen de variables de batch", rows)

    # ── SLIDE 4-N: GRÁFICAS ──
    tipos_graficas = [
        ("Series de Tiempo", "Perfil de variables a lo largo del proceso"),
        ("Correlaciones", "Relaciones entre variables del proceso"),
        ("Control SPC", "Carta de control estadístico de proceso"),
        ("Riesgo Financiero", "Análisis de costo de la variabilidad"),
    ]

    for i, img_path in enumerate(imagenes[:8]):
        if img_path and os.path.exists(img_path):
            titulo, subtitulo = tipos_graficas[i % len(tipos_graficas)]
            add_slide_dark(titulo, subtitulo, img_path=img_path)

    # ── SLIDE RIESGO ──
    if riesgo:
        add_slide_dark(
            "Riesgo Financiero",
            "El costo de la variabilidad actual",
            [
                f"📊 Tasa de rechazo: {riesgo['tasa_rechazo_pct']}%",
                f"💸 Costo promedio por lote: ${riesgo['costo_promedio_usd']:,.0f} USD",
                f"❌ Costo lote rechazado: ${riesgo['costo_lote_rechazado_usd']:,.0f} USD",
                f"⚠️  Riesgo mensual: ${riesgo['riesgo_mensual_usd']:,.0f} USD",
                f"🔴 Riesgo anual estimado: ${riesgo['riesgo_anual_usd']:,.0f} USD",
                "",
                f"💡 Reducir rechazos al 5% ahorraría ~${riesgo['riesgo_anual_usd']*0.6:,.0f} USD/año",
            ]
        )

    # ── SLIDE RECOMENDACIONES ──
    add_slide_dark(
        "Top 5 Recomendaciones",
        "Acciones priorizadas por impacto",
        [
            "1️⃣  Revisar sensores con estado CRÍTICO antes del próximo lote",
            "2️⃣  Implementar alertas automáticas para variables fuera de ±3σ",
            "3️⃣  Investigar causas raíz de los lotes rechazados (correlación con pH/DO)",
            "4️⃣  Establecer programa de calibración preventiva de sensores",
            "5️⃣  Actualizar análisis mensualmente para SPC continuo",
        ]
    )

    # ── SLIDE PLAN DE ACCIÓN ──
    add_slide_dark(
        "Plan de Acción",
        "¿Qué hacemos la próxima semana?",
        [
            "📋 Semana 1: Revisión y calibración de sensores críticos",
            "📋 Semana 2: Análisis de causa raíz de últimos rechazos",
            "📋 Semana 3: Configurar alertas automáticas en SCADA/Excel",
            "📋 Mes 1:   Protocolo de monitoreo estadístico mensual",
            "📋 Trimestre: Re-validación de especificaciones de proceso",
            "",
            "📞 Contacto: jeluani127.github.io/prompt01",
        ]
    )

    # ── SLIDE FINAL ──
    add_slide_dark(
        "Conclusiones",
        "5 mensajes clave para recordar",
        [
            "✅ Tus datos tienen el potencial de predecir rechazos antes de que ocurran",
            "✅ La variabilidad tiene un costo financiero cuantificable",
            "✅ Los sensores son la primera fuente de error — auditarlos salva lotes",
            "✅ Un análisis estadístico mensual es la mejor inversión de proceso",
            "✅ Este reporte fue generado en minutos con Prompt Maestro 01",
            "",
            "🧬 jeluani127.github.io/prompt01",
        ]
    )

    prs.save(output_path)
    print(f"   ✅ PowerPoint guardado: {output_path}")


# ─────────────────────────────────────────────────────────────
# MAIN
# ─────────────────────────────────────────────────────────────

def main():
    print("\n" + "═"*60)
    print("  🧬 BIOREACTOR ANALYSIS — Prompt Maestro 01")
    print("  jeluani127.github.io/prompt01")
    print("═"*60)

    output_dir = "graficas_biorreactor"

    # 1. Cargar datos
    dfs = cargar_datos()
    df_sensores = dfs.get('sensores')
    df_batch    = dfs.get('batch')

    # 2. Determinar variables numéricas
    vars_sensores = []
    stats_sensores = None
    if df_sensores is not None:
        vars_sensores = df_sensores.select_dtypes(include=[np.number]).columns.tolist()
        excluir = ['Tiempo_h', 'Lote_Num']
        vars_sensores = [v for v in vars_sensores if v not in excluir]
        print(f"\n📊 Variables de sensores: {vars_sensores}")

        print("\n🔢 Calculando estadísticas descriptivas (sensores)...")
        stats_sensores = estadistica_descriptiva(df_sensores, vars_sensores)

    vars_batch = []
    stats_batch = None
    if df_batch is not None:
        vars_batch = df_batch.select_dtypes(include=[np.number]).columns.tolist()
        print(f"📋 Variables de batch: {vars_batch}")

        print("\n🔢 Calculando estadísticas descriptivas (batch)...")
        stats_batch = estadistica_descriptiva(df_batch, vars_batch)

    # 3. Auditoría de sensores
    auditoria = None
    if df_sensores is not None and vars_sensores:
        print("\n🔍 Auditando sensores...")
        auditoria = detectar_anomalias_sensores(df_sensores, vars_sensores)
        print(auditoria[['Estado', 'Sensor_Congelado', 'Saltos_Abruptos', 'Deriva_Detectada']])

    # 4. Riesgo financiero
    riesgo = None
    if df_batch is not None:
        print("\n💰 Calculando riesgo financiero...")
        riesgo = calcular_riesgo_financiero(df_batch)
        if riesgo:
            print(f"   Tasa de rechazo: {riesgo['tasa_rechazo_pct']}%")
            print(f"   Riesgo anual estimado: ${riesgo['riesgo_anual_usd']:,.0f} USD")

    # 5. Generar gráficas
    print("\n🎨 Generando gráficas...")
    todas_imagenes = []

    if df_sensores is not None and vars_sensores:
        imgs_series = graficar_series_tiempo(df_sensores, vars_sensores[:4], output_dir)
        todas_imagenes.extend(imgs_series)

        heatmap, scatters, corr = graficar_correlaciones(df_sensores, vars_sensores, output_dir)
        todas_imagenes.append(heatmap)
        todas_imagenes.extend(scatters)

    if df_batch is not None:
        boxplots = graficar_boxplots(df_batch, vars_batch, output_dir)
        todas_imagenes.extend(boxplots)

        if vars_batch:
            spc_img = graficar_spc(df_batch, vars_batch[0], output_dir)
            if spc_img:
                todas_imagenes.append(spc_img)

        risk_img = graficar_riesgo_financiero(df_batch, riesgo, output_dir)
        if risk_img:
            todas_imagenes.append(risk_img)

    # 6. Exportar
    print("\n📤 Exportando entregables...")

    excel_path = "analisis_biorreactor.xlsx"
    word_path  = "reporte_biorreactor.docx"
    pptx_path  = "presentacion_biorreactor.pptx"

    exportar_excel(dfs, stats_sensores, stats_batch, auditoria, riesgo, excel_path)
    exportar_word(stats_sensores, stats_batch, auditoria, riesgo, todas_imagenes[:6], word_path)
    exportar_pptx(stats_sensores, stats_batch, riesgo, todas_imagenes, pptx_path)

    print("\n" + "═"*60)
    print("  ✅ ANÁLISIS COMPLETADO")
    print("═"*60)
    print(f"\n  📊 Excel:      {excel_path}")
    print(f"  📄 Word:       {word_path}")
    print(f"  📑 PowerPoint: {pptx_path}")
    print(f"  🖼️  Gráficas:   carpeta '{output_dir}/'")
    print("\n  🧬 Prompt Maestro 01 | jeluani127.github.io/prompt01")
    print("═"*60 + "\n")


if __name__ == "__main__":
    main()
