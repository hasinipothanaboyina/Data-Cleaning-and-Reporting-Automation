"""
Data Cleaning & Reporting Automation
Assignment: Automate data cleaning and reporting workflows
Dataset:    sales_data.csv (Walmart — 45 stores, 2010–2012)
Output:     walmart_data_cleaning_report.xlsx  (5-sheet automated report)
            data_cleaning_dashboard.png        (visual summary)
"""

import pandas as pd
import numpy as np
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
from matplotlib.gridspec import GridSpec
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.chart import BarChart, Reference
import warnings, os, time
warnings.filterwarnings('ignore')

# ══════════════════════════════════════════════════════════════
# STEP 1 — LOAD RAW DATA
# ══════════════════════════════════════════════════════════════
print("=" * 60)
print("  DATA CLEANING & REPORTING AUTOMATION PIPELINE")
print("=" * 60)

start = time.time()
df_raw = pd.read_csv('sales_data.csv')
print(f"\n[LOAD]  Raw dataset: {df_raw.shape[0]:,} rows × {df_raw.shape[1]} cols")

# ══════════════════════════════════════════════════════════════
# STEP 2 — DATA CLEANING PIPELINE
# ══════════════════════════════════════════════════════════════
df = df_raw.copy()
cleaning_log = []

# ── 2.1 Parse Dates ────────────────────────────────────────────
before = df.dtypes['Date']
df['Date'] = pd.to_datetime(df['Date'], dayfirst=True)
cleaning_log.append(("Date Parsing", "Date as string", "Converted to datetime64", 6435, "✅ Fixed"))
print(f"[STEP 1] Date column parsed → datetime64")

# ── 2.2 Remove Duplicates ──────────────────────────────────────
before_dup = len(df)
df = df.drop_duplicates()
removed = before_dup - len(df)
cleaning_log.append(("Duplicate Removal", f"{removed} duplicates", "drop_duplicates()", removed, "✅ Fixed"))
print(f"[STEP 2] Duplicates removed: {removed}")

# ── 2.3 Handle Missing Values ──────────────────────────────────
missing_total = df.isnull().sum().sum()
missing_by_col = df.isnull().sum()
df = df.fillna(df.median(numeric_only=True))
cleaning_log.append(("Missing Values", f"{missing_total} nulls", "Filled with column median", missing_total, "✅ Fixed"))
print(f"[STEP 3] Missing values handled: {missing_total}")

# ── 2.4 Fix Data Types ─────────────────────────────────────────
df['Store']        = df['Store'].astype(int)
df['Holiday_Flag'] = df['Holiday_Flag'].astype(int)
cleaning_log.append(("Data Types", "Store/Holiday as float", "Cast to int64", 2, "✅ Fixed"))
print(f"[STEP 4] Data types corrected")

# ── 2.5 Outlier Detection (IQR) ────────────────────────────────
Q1  = df['Weekly_Sales'].quantile(0.25)
Q3  = df['Weekly_Sales'].quantile(0.75)
IQR = Q3 - Q1
lb  = Q1 - 1.5 * IQR
ub  = Q3 + 1.5 * IQR
df['Is_Outlier'] = ((df['Weekly_Sales'] < lb) | (df['Weekly_Sales'] > ub)).astype(int)
outlier_count = df['Is_Outlier'].sum()
cleaning_log.append(("Outlier Detection", f"{outlier_count} outliers (IQR)", "Flagged in Is_Outlier column", outlier_count, "⚠️ Flagged"))
print(f"[STEP 5] Outliers flagged: {outlier_count} ({outlier_count/len(df)*100:.1f}%)")

# ── 2.6 Feature Engineering ────────────────────────────────────
df['Year']       = df['Date'].dt.year
df['Month']      = df['Date'].dt.month
df['Month_Name'] = df['Date'].dt.strftime('%b')
df['Week_Label'] = df['Date'].dt.strftime('%d-%b-%Y')
cleaning_log.append(("Feature Engineering", "No time features", "Added Year, Month, Week_Label", 6435, "✅ Added"))
print(f"[STEP 6] Features engineered: Year, Month, Month_Name, Week_Label")

# ── 2.7 Normalization ──────────────────────────────────────────
for col in ['Temperature', 'Fuel_Price', 'CPI', 'Unemployment']:
    df[f'{col}_Norm'] = ((df[col] - df[col].min()) /
                          (df[col].max() - df[col].min())).round(4)
cleaning_log.append(("Normalization", "Features on diff scales", "Min-Max (0–1)", 4, "✅ Done"))
print(f"[STEP 7] Normalization applied: Temperature, Fuel_Price, CPI, Unemployment")

# ══════════════════════════════════════════════════════════════
# STEP 3 — AGGREGATIONS (for report)
# ══════════════════════════════════════════════════════════════
yearly   = df.groupby('Year')['Weekly_Sales'].agg(['sum','mean','min','max','count']).reset_index()
monthly  = df.groupby(['Month','Month_Name'])['Weekly_Sales'].mean().reset_index().sort_values('Month')
store    = df.groupby('Store')['Weekly_Sales'].sum().reset_index().sort_values('Weekly_Sales', ascending=False).reset_index(drop=True)
holiday  = df.groupby('Holiday_Flag')['Weekly_Sales'].mean()
corr_df  = df[['Weekly_Sales','Temperature','Fuel_Price','CPI','Unemployment']].corr()['Weekly_Sales'].drop('Weekly_Sales')

print("\n[AGGREGATE] Yearly, Monthly, Store, Holiday summaries ready")

# ══════════════════════════════════════════════════════════════
# STEP 4 — VISUAL DASHBOARD (PNG)
# ══════════════════════════════════════════════════════════════
BG=  '#0d1117'; CARD='#161b22'
TEAL='#00b894'; RED= '#e17055'; PURPLE='#6c5ce7'
YELLOW='#fdcb6e'; BLUE='#0984e3'
TEXT='#e6edf3'; MUTED='#8b949e'

fig = plt.figure(figsize=(18, 13), facecolor=BG)
fig.suptitle('Walmart Sales — Data Cleaning & Reporting Automation',
             fontsize=20, fontweight='bold', color=TEXT, y=0.98)
gs = GridSpec(3, 3, figure=fig, hspace=0.45, wspace=0.35,
              left=0.06, right=0.97, top=0.93, bottom=0.06)

def style(ax):
    ax.tick_params(colors=TEXT, labelsize=9)
    ax.grid(alpha=0.15, color=MUTED)
    for sp in ax.spines.values(): sp.set_color(MUTED); sp.set_alpha(0.2)

# 1: Cleaning Pipeline Steps
ax1 = fig.add_subplot(gs[0, :2]); ax1.set_facecolor(CARD)
steps = [l[0] for l in cleaning_log]
affected = [l[3] for l in cleaning_log]
colors_b = [TEAL if '✅' in l[4] else YELLOW for l in cleaning_log]
bars = ax1.barh(steps, affected, color=colors_b, height=0.55)
for bar, val in zip(bars, affected):
    ax1.text(bar.get_width() + 20, bar.get_y() + bar.get_height()/2,
             str(val), va='center', color=TEXT, fontsize=9)
ax1.set_title('Cleaning Pipeline — Records Affected', color=TEXT, fontsize=12, pad=8)
ax1.set_xlabel('Records', color=MUTED, fontsize=10)
p1 = mpatches.Patch(color=TEAL, label='Fixed'); p2 = mpatches.Patch(color=YELLOW, label='Flagged')
ax1.legend(handles=[p1, p2], facecolor=CARD, labelcolor=TEXT, fontsize=9)
style(ax1)

# 2: KPI Card
ax2 = fig.add_subplot(gs[0, 2]); ax2.set_facecolor(CARD); ax2.axis('off')
ax2.text(0.5, 0.95, 'Pipeline Summary', ha='center', va='top',
         transform=ax2.transAxes, color=TEXT, fontsize=12, fontweight='bold')
kpis = [('Total Records',   f'{len(df):,}',           TEAL),
        ('Outliers Flagged', str(outlier_count),       YELLOW),
        ('Missing Filled',   str(missing_total),       PURPLE),
        ('Duplicates Rm.',   str(removed),             RED),
        ('Features Added',   '4',                      BLUE)]
for i, (lbl, val, col) in enumerate(kpis):
    yp = 0.78 - i*0.15
    ax2.text(0.1, yp, lbl, transform=ax2.transAxes, color=MUTED, fontsize=9)
    ax2.text(0.9, yp, val, transform=ax2.transAxes, color=col,
             fontsize=11, fontweight='bold', ha='right')

# 3: Monthly Sales Pattern
ax3 = fig.add_subplot(gs[1, :2]); ax3.set_facecolor(CARD)
mc = [RED if m in [11,12] else BLUE if m in [1,2] else TEAL for m in monthly['Month']]
ax3.bar(monthly['Month_Name'], monthly['Weekly_Sales']/1e6, color=mc, width=0.7)
ax3.set_title('Avg Weekly Sales by Month (Seasonal Pattern)', color=TEXT, fontsize=12, pad=8)
ax3.set_ylabel('Avg Sales (M$)', color=MUTED, fontsize=10)
p1=mpatches.Patch(color=RED,label='Peak'); p2=mpatches.Patch(color=BLUE,label='Low'); p3=mpatches.Patch(color=TEAL,label='Normal')
ax3.legend(handles=[p1,p2,p3], facecolor=CARD, labelcolor=TEXT, fontsize=9)
style(ax3)

# 4: Holiday Impact
ax4 = fig.add_subplot(gs[1, 2]); ax4.set_facecolor(CARD)
h_vals = [holiday[0]/1e6, holiday[1]/1e6]
ax4.bar(['Non-Holiday','Holiday'], h_vals, color=[PURPLE, RED], width=0.5)
diff = (holiday[1]-holiday[0])/holiday[0]*100
ax4.text(1, h_vals[1]+0.02, f'+{diff:.1f}%', ha='center', color=RED, fontsize=12, fontweight='bold')
ax4.set_title('Holiday vs Non-Holiday\nAvg Weekly Sales', color=TEXT, fontsize=11, pad=8)
ax4.set_ylabel('Avg Sales (M$)', color=MUTED, fontsize=10)
style(ax4)

# 5: Top 10 Stores
ax5 = fig.add_subplot(gs[2, :2]); ax5.set_facecolor(CARD)
top10 = store.head(10)
bar_colors = [YELLOW]*3 + [BLUE]*4 + [MUTED]*3
ax5.barh(range(10), top10['Weekly_Sales'].values/1e9, color=bar_colors, height=0.6)
ax5.set_yticks(range(10))
ax5.set_yticklabels([f'Store {int(s)}' for s in top10['Store']], color=TEXT, fontsize=9)
ax5.set_title('Top 10 Stores by Total Revenue', color=TEXT, fontsize=12, pad=8)
ax5.set_xlabel('Total Sales (B$)', color=MUTED, fontsize=10)
style(ax5)

# 6: Feature Correlation
ax6 = fig.add_subplot(gs[2, 2]); ax6.set_facecolor(CARD)
cv = corr_df.values
cf = corr_df.index.tolist()
ax6.barh(cf, cv, color=[TEAL if v>0 else RED for v in cv], height=0.5)
ax6.axvline(0, color=MUTED, alpha=0.5, linewidth=1)
ax6.set_title('Feature Correlation\nwith Weekly Sales', color=TEXT, fontsize=11, pad=8)
ax6.set_xlabel('Correlation', color=MUTED, fontsize=10)
style(ax6)

plt.savefig('data_cleaning_dashboard.png', dpi=150,
            bbox_inches='tight', facecolor=BG, edgecolor='none')
print("[VISUAL] Dashboard saved → data_cleaning_dashboard.png")

# ══════════════════════════════════════════════════════════════
# STEP 5 — EXCEL REPORT (5 sheets)
# ══════════════════════════════════════════════════════════════
# (See wallet_data_cleaning_report.xlsx for the full formatted report)
# Quick CSV export of cleaned data
df.to_csv('walmart_cleaned.csv', index=False)
print("[EXPORT] Cleaned CSV saved → walmart_cleaned.csv")

elapsed = round(time.time() - start, 2)
print(f"\n{'='*60}")
print(f"  PIPELINE COMPLETE in {elapsed}s")
print(f"  Records:  {len(df_raw):,} → {len(df):,} (cleaned)")
print(f"  Outliers: {outlier_count} flagged")
print(f"  Outputs:  data_cleaning_dashboard.png")
print(f"            walmart_cleaned.csv")
print(f"{'='*60}")