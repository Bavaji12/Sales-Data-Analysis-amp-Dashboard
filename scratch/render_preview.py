import matplotlib.pyplot as plt
import matplotlib.gridspec as gridspec
import pandas as pd
import numpy as np

# Set high DPI and aesthetic styling
plt.rcParams['font.sans-serif'] = 'DejaVu Sans'
plt.rcParams['axes.edgecolor'] = '#CBD5E1'
plt.rcParams['axes.linewidth'] = 0.8

df = pd.read_csv("data/sales_data.csv")

# Create large canvas (1920x1080 proportion)
fig = plt.figure(figsize=(18, 10.5), facecolor='#F8FAFC')
gs = gridspec.GridSpec(3, 3, height_ratios=[0.55, 1.2, 1.2], hspace=0.38, wspace=0.28,
                       left=0.04, right=0.96, top=0.93, bottom=0.05)

# ---------------------------------------------------------
# 1. HEADER BANNER
# ---------------------------------------------------------
header_ax = fig.add_axes([0.04, 0.94, 0.92, 0.05], facecolor='#1B2A4A')
header_ax.text(0.02, 0.65, "EXECUTIVE SALES & PROFITABILITY KPI DASHBOARD", 
               color='white', fontsize=16, fontweight='bold', va='center')
header_ax.text(0.02, 0.25, "FY 2024 Multi-Channel Retail Analytics | 1,000 Transactions | 730 Unique Orders", 
               color='#93C5FD', fontsize=10, va='center')
header_ax.text(0.98, 0.5, "Live Portfolio Demo | Microsoft Excel 365", 
               color='#E2E8F0', fontsize=11, fontweight='bold', ha='right', va='center')
header_ax.axis('off')

# ---------------------------------------------------------
# 2. TOP KPI CARDS (6 CARDS ACROSS ROW 0)
# ---------------------------------------------------------
kpis = [
    ("TOTAL NET SALES", "$1,042,276", "Gross: $1.15M | Disc: 9.7%", "#1E293B", "#2563EB"),
    ("TOTAL ORDERS", "730", "Unique Orders (1.37 items/ord)", "#1E293B", "#0D9488"),
    ("TOTAL PROFIT", "$164,133", "Net Margin Target: $150K (+9.4%)", "#1E293B", "#10B981"),
    ("PROFIT MARGIN", "15.8%", "Benchmark Target: 15.0%", "#1E293B", "#059669"),
    ("AVG ORDER VALUE", "$1,428", "Revenue per unique order", "#1E293B", "#6366F1"),
    ("PEAK MOM GROWTH", "+158.5%", "November Black Friday Surge", "#1E293B", "#D97706")
]

card_w = 0.92 / 6 - 0.012
for i, (title, val, sub, txt_c, accent_c) in enumerate(kpis):
    card_x = 0.04 + i * (card_w + 0.012)
    card_ax = fig.add_axes([card_x, 0.81, card_w, 0.11], facecolor='#FFFFFF')
    card_ax.spines['top'].set_visible(True)
    card_ax.spines['top'].set_color(accent_c)
    card_ax.spines['top'].set_linewidth(3.5)
    for s in ['left', 'right', 'bottom']:
        card_ax.spines[s].set_color('#E2E8F0')
        card_ax.spines[s].set_linewidth(1)
        
    card_ax.text(0.08, 0.80, title, color='#64748B', fontsize=8.5, fontweight='bold', va='center')
    card_ax.text(0.08, 0.45, val, color=txt_c, fontsize=15, fontweight='bold', va='center')
    card_ax.text(0.08, 0.18, sub, color=accent_c, fontsize=7.5, va='center')
    card_ax.set_xticks([])
    card_ax.set_yticks([])

# ---------------------------------------------------------
# 3. CHART 1: Monthly Net Sales Trend (Line Chart)
# ---------------------------------------------------------
ax1 = fig.add_subplot(gs[1, 0], facecolor='white')
df['Month_Num'] = pd.to_datetime(df['Order Date']).dt.month
m_sales = df.groupby('Month_Num')['Net Sales'].sum() / 1000
months = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]

ax1.plot(months, m_sales, marker='o', color='#2563EB', linewidth=2.5, markersize=5, label='Net Sales ($K)')
ax1.fill_between(months, m_sales, color='#DBEAFE', alpha=0.5)
ax1.set_title("Monthly Net Sales Trend ($K)", fontsize=11, fontweight='bold', color='#1E293B', pad=10)
ax1.set_ylabel("Sales ($K)", fontsize=9, color='#64748B')
ax1.grid(True, linestyle='--', alpha=0.5, color='#E2E8F0')
ax1.tick_params(axis='both', labelsize=8)
# Highlight Nov peak
ax1.annotate('Peak: $175K\n(+158% MoM)', xy=(10, m_sales[11]), xytext=(7.8, 155),
             arrowprops=dict(facecolor='#D97706', shrink=0.08, width=1, headwidth=5),
             fontsize=7.5, fontweight='bold', color='#92400E',
             bbox=dict(boxstyle="round,pad=0.3", fc="#FEF3C7", ec="#F59E0B", lw=0.8))

# ---------------------------------------------------------
# 4. CHART 2: Net Sales by Product Category (Bar Chart)
# ---------------------------------------------------------
ax2 = fig.add_subplot(gs[1, 1], facecolor='white')
cat_data = df.groupby('Category')['Net Sales'].sum() / 1000
cat_order = ['Technology', 'Furniture', 'Office Supplies']
cat_vals = [cat_data[c] for c in cat_order]
colors_cat = ['#1E3A8A', '#3B82F6', '#93C5FD']

bars2 = ax2.bar(cat_order, cat_vals, color=colors_cat, width=0.55, edgecolor='#CBD5E1')
ax2.set_title("Net Sales by Category ($K)", fontsize=11, fontweight='bold', color='#1E293B', pad=10)
ax2.set_ylabel("Sales ($K)", fontsize=9, color='#64748B')
ax2.grid(axis='y', linestyle='--', alpha=0.5, color='#E2E8F0')
ax2.tick_params(axis='both', labelsize=8.5)
for bar in bars2:
    y = bar.get_height()
    ax2.text(bar.get_x() + bar.get_width()/2, y + 15, f"${y:,.0f}K", ha='center', fontsize=8, fontweight='bold', color='#1E293B')
ax2.set_ylim(0, max(cat_vals) * 1.18)

# ---------------------------------------------------------
# 5. CHART 3: Profit by Geographic Region (Column Chart)
# ---------------------------------------------------------
ax3 = fig.add_subplot(gs[1, 2], facecolor='white')
reg_profit = df.groupby('Region')['Profit'].sum() / 1000
reg_order = ['East', 'West', 'Central', 'South']
reg_vals = [reg_profit[r] for r in reg_order]
colors_reg = ['#0D9488', '#14B8A6', '#2DD4BF', '#5EEAD4']

bars3 = ax3.bar(reg_order, reg_vals, color=colors_reg, width=0.55, edgecolor='#CBD5E1')
ax3.set_title("Net Operating Profit by Region ($K)", fontsize=11, fontweight='bold', color='#1E293B', pad=10)
ax3.set_ylabel("Profit ($K)", fontsize=9, color='#64748B')
ax3.grid(axis='y', linestyle='--', alpha=0.5, color='#E2E8F0')
ax3.tick_params(axis='both', labelsize=8.5)
for bar in bars3:
    y = bar.get_height()
    ax3.text(bar.get_x() + bar.get_width()/2, y + 1.2, f"${y:,.1f}K", ha='center', fontsize=8, fontweight='bold', color='#0F766E')
ax3.set_ylim(0, max(reg_vals) * 1.18)

# ---------------------------------------------------------
# 6. CHART 4: Sales Channel Mix (Donut Chart)
# ---------------------------------------------------------
ax4 = fig.add_subplot(gs[2, 0], facecolor='white')
chan_sales = df.groupby('Sales Channel')['Net Sales'].sum()
chan_labels = ['Online', 'Wholesale', 'Direct Sales', 'Retail Partner']
chan_vals = [chan_sales[c] for c in chan_labels]
colors_chan = ['#2563EB', '#38BDF8', '#818CF8', '#C7D2FE']

wedges, texts, autotexts = ax4.pie(
    chan_vals, labels=chan_labels, autopct='%1.1f%%', startangle=140,
    colors=colors_chan, pctdistance=0.75, wedgeprops=dict(width=0.45, edgecolor='white')
)
for t in texts:
    t.set_fontsize(8)
    t.set_color('#334155')
for at in autotexts:
    at.set_fontsize(7.5)
    at.set_fontweight('bold')
    at.set_color('white')
ax4.set_title("Sales Distribution by Channel", fontsize=11, fontweight='bold', color='#1E293B', pad=10)

# ---------------------------------------------------------
# 7. CHART 5: Top 10 Products by Net Sales (Horizontal Bar)
# ---------------------------------------------------------
ax5 = fig.add_subplot(gs[2, 1], facecolor='white')
top_p = df.groupby('Product Name')['Net Sales'].sum().sort_values(ascending=True).tail(8) / 1000
p_names = [p[:22] + "..." if len(p) > 22 else p for p in top_p.index]

bars5 = ax5.barh(p_names, top_p.values, color='#475569', height=0.6, edgecolor='#CBD5E1')
# Color Apex Gaming Laptop red to highlight loss leader
bars5[-1].set_color('#EF4444')
bars5[-2].set_color('#10B981') # ProBook Star

ax5.set_title("Top Products by Sales ($K) [Loss-Leader Flagged]", fontsize=10.5, fontweight='bold', color='#1E293B', pad=10)
ax5.set_xlabel("Net Sales ($K)", fontsize=8.5, color='#64748B')
ax5.grid(axis='x', linestyle='--', alpha=0.5, color='#E2E8F0')
ax5.tick_params(axis='both', labelsize=7.5)
for bar in bars5:
    w = bar.get_width()
    ax5.text(w + 1.5, bar.get_y() + bar.get_height()/2, f"${w:,.0f}K", va='center', fontsize=7.5, fontweight='bold', color='#1E293B')
ax5.set_xlim(0, max(top_p.values) * 1.22)

# ---------------------------------------------------------
# 8. CHART 6: Monthly Order Volume (Column Chart)
# ---------------------------------------------------------
ax6 = fig.add_subplot(gs[2, 2], facecolor='white')
m_orders = df.groupby('Month_Num')['Order ID'].nunique()

bars6 = ax6.bar(months, m_orders.values, color='#6366F1', width=0.55, edgecolor='#CBD5E1')
bars6[10].set_color('#D97706') # Nov
bars6[11].set_color('#F59E0B') # Dec

ax6.set_title("Monthly Order Volume (# Unique Orders)", fontsize=11, fontweight='bold', color='#1E293B', pad=10)
ax6.set_ylabel("Order Count", fontsize=9, color='#64748B')
ax6.grid(axis='y', linestyle='--', alpha=0.5, color='#E2E8F0')
ax6.tick_params(axis='both', labelsize=8)
for bar in bars6:
    y = bar.get_height()
    ax6.text(bar.get_x() + bar.get_width()/2, y + 1.8, f"{y}", ha='center', fontsize=7.5, fontweight='bold', color='#312E81')
ax6.set_ylim(0, max(m_orders.values) * 1.18)

# Save high-res screenshot
plt.savefig("screenshots/dashboard-preview.png", dpi=200, bbox_inches='tight')
print("Successfully generated high-res dashboard preview: screenshots/dashboard-preview.png")
