import random
import datetime
import numpy as np
import pandas as pd

# Set fixed seed for perfect reproducibility
np.random.seed(42)
random.seed(42)

# Define Product Catalog with Category, Subcategory, Base Price, Base Margin
PRODUCTS = [
    # Technology
    {"id": "TEC-LAP-001", "name": "ProBook Elite 15.6\"", "category": "Technology", "subcategory": "Laptops", "price": 1199.99, "margin": 0.28},
    {"id": "TEC-LAP-002", "name": "Apex Gaming Laptop 17\"", "category": "Technology", "subcategory": "Laptops", "price": 1499.99, "margin": 0.16}, # High sales, low margin candidate
    {"id": "TEC-TAB-003", "name": "VisionPad Pro 11\"", "category": "Technology", "subcategory": "Phones & Tablets", "price": 649.99, "margin": 0.25},
    {"id": "TEC-TAB-004", "name": "VisionPad Lite 8.5\"", "category": "Technology", "subcategory": "Phones & Tablets", "price": 299.99, "margin": 0.18},
    {"id": "TEC-ACC-005", "name": "UltraWireless Noise-Canceling Headset", "category": "Technology", "subcategory": "Audio & Accessories", "price": 179.99, "margin": 0.38},
    {"id": "TEC-ACC-006", "name": "MultiPort Thunderbolt 4 Dock", "category": "Technology", "subcategory": "Audio & Accessories", "price": 149.99, "margin": 0.34},
    {"id": "TEC-ACC-007", "name": "ErgoMechanical Wireless Keyboard", "category": "Technology", "subcategory": "Audio & Accessories", "price": 109.99, "margin": 0.40},
    {"id": "TEC-ACC-008", "name": "Precision Laser Bluetooth Mouse", "category": "Technology", "subcategory": "Audio & Accessories", "price": 49.99, "margin": 0.45},

    # Furniture
    {"id": "FUR-CHA-001", "name": "ErgoFlex Executive Mesh Chair", "category": "Furniture", "subcategory": "Chairs", "price": 389.99, "margin": 0.15}, # Low margin
    {"id": "FUR-CHA-002", "name": "AeroTask Operator Task Chair", "category": "Furniture", "subcategory": "Chairs", "price": 219.99, "margin": 0.22},
    {"id": "FUR-DSK-003", "name": "ApexMotor Dual-Motor Standing Desk", "category": "Furniture", "subcategory": "Desks & Workstations", "price": 549.99, "margin": 0.20},
    {"id": "FUR-DSK-004", "name": "Metro Compact Corner Workstation", "category": "Furniture", "subcategory": "Desks & Workstations", "price": 329.99, "margin": 0.24},
    {"id": "FUR-BOK-005", "name": "Modular 5-Tier Steel Bookcase", "category": "Furniture", "subcategory": "Bookcases", "price": 189.99, "margin": 0.25},
    {"id": "FUR-LGT-006", "name": "Architect LED Desk Task Lamp", "category": "Furniture", "subcategory": "Lighting", "price": 69.99, "margin": 0.38},

    # Office Supplies
    {"id": "OFF-STR-001", "name": "SteelLock 3-Drawer File Cabinet", "category": "Office Supplies", "subcategory": "Storage & Organizers", "price": 169.99, "margin": 0.32},
    {"id": "OFF-STR-002", "name": "Stackable Heavy-Duty Storage Bins (4-Pk)", "category": "Office Supplies", "subcategory": "Storage & Organizers", "price": 44.99, "margin": 0.45},
    {"id": "OFF-PAP-003", "name": "UltraWhite Premium Multipurpose Paper (Case)", "category": "Office Supplies", "subcategory": "Paper Products", "price": 54.99, "margin": 0.42},
    {"id": "OFF-PAP-004", "name": "EcoRecycled Copy Paper (5 Reams)", "category": "Office Supplies", "subcategory": "Paper Products", "price": 32.99, "margin": 0.40},
    {"id": "OFF-BIN-005", "name": "Heavy-Duty D-Ring Presentation Binder (6-Pk)", "category": "Office Supplies", "subcategory": "Binders & Fasteners", "price": 27.99, "margin": 0.50},
    {"id": "OFF-ART-006", "name": "Professional Presentation Dry Erase Board", "category": "Office Supplies", "subcategory": "Presentation & Art", "price": 89.99, "margin": 0.40},
    {"id": "OFF-DSK-007", "name": "Gel Ergo Mouse Pad & Wrist Rest Combo", "category": "Office Supplies", "subcategory": "Desk Supplies", "price": 21.99, "margin": 0.52},
    {"id": "OFF-DSK-008", "name": "Steel Desktop Mesh Organizer Caddy", "category": "Office Supplies", "subcategory": "Desk Supplies", "price": 18.99, "margin": 0.55}
]

# Regional State Mappings
REGIONS = {
    "West": ["California", "Washington", "Oregon", "Colorado"],
    "East": ["New York", "Pennsylvania", "Massachusetts", "New Jersey"],
    "Central": ["Illinois", "Texas", "Ohio", "Michigan"],
    "South": ["Florida", "Georgia", "North Carolina", "Virginia"]
}

# Regional weights (West & East higher volume)
REGION_WEIGHTS = {"West": 0.33, "East": 0.30, "Central": 0.22, "South": 0.15}

# Channels & Weights
CHANNELS = ["Online", "Direct Sales", "Retail Partner", "Wholesale"]
CHANNEL_WEIGHTS = [0.42, 0.25, 0.20, 0.13]

# Customer Segments & Weights
SEGMENTS = ["Consumer", "Corporate", "Home Office"]
SEGMENT_WEIGHTS = [0.45, 0.35, 0.20]

# Monthly seasonal sales weight (Q4 holiday surge, Q3 back-to-school)
MONTH_WEIGHTS = {
    1: 0.058, 2: 0.062, 3: 0.070, 4: 0.072, 5: 0.075, 6: 0.080,
    7: 0.078, 8: 0.090, 9: 0.092, 10: 0.095, 11: 0.140, 12: 0.158
}

def generate_dataset(num_records=1000):
    rows = []
    
    # 820 unique orders across 1000 rows (approx 1.22 items per order)
    order_id_counter = 1001
    current_order_id = f"ORD-2024-{order_id_counter}"
    
    # Precompute dates based on month weights
    months = list(MONTH_WEIGHTS.keys())
    m_probs_raw = [MONTH_WEIGHTS[m] for m in months]
    total_m_weight = sum(m_probs_raw)
    m_probs = [w / total_m_weight for w in m_probs_raw]
    
    # Pre-select order attributes that stay constant per Order ID
    current_order_date = None
    current_region = None
    current_state = None
    current_channel = None
    current_segment = None
    
    items_in_current_order = 0
    max_items_for_this_order = np.random.choice([1, 2, 3, 4], p=[0.75, 0.18, 0.05, 0.02])
    
    for i in range(num_records):
        if items_in_current_order >= max_items_for_this_order or i == 0:
            # Start new order
            if i > 0:
                order_id_counter += 1
                current_order_id = f"ORD-2024-{order_id_counter}"
            
            # Select month based on seasonality
            month = np.random.choice(months, p=m_probs)
            # Pick valid day
            if month == 2:
                day = random.randint(1, 28)
            elif month in [4, 6, 9, 11]:
                day = random.randint(1, 30)
            else:
                day = random.randint(1, 31)
            current_order_date = datetime.date(2024, month, day)
            
            # Select Region and State
            current_region = np.random.choice(list(REGIONS.keys()), p=[REGION_WEIGHTS[r] for r in REGIONS.keys()])
            current_state = random.choice(REGIONS[current_region])
            
            # Select Channel and Segment
            current_channel = np.random.choice(CHANNELS, p=CHANNEL_WEIGHTS)
            current_segment = np.random.choice(SEGMENTS, p=SEGMENT_WEIGHTS)
            
            items_in_current_order = 0
            max_items_for_this_order = np.random.choice([1, 2, 3, 4], p=[0.75, 0.18, 0.05, 0.02])
            
        items_in_current_order += 1
        
        # Select Product
        # Tech has slightly higher weight, Office Supplies frequent
        prod = random.choice(PRODUCTS)
        
        # Quantity based on channel and category
        if current_channel == "Wholesale":
            qty = random.randint(5, 15)
        elif current_segment == "Corporate":
            qty = random.randint(2, 8)
        else:
            qty = random.randint(1, 4)
            
        # Discount logic
        # Tech laptops occasionally get 10-20% discounts during Q4 or corporate
        # Apex Gaming Laptop gets frequent aggressive discounts (loss leader)
        if prod["id"] == "TEC-LAP-002":
            discount_pct = np.random.choice([0.10, 0.15, 0.20, 0.25], p=[0.2, 0.3, 0.3, 0.2])
        elif prod["id"] == "FUR-CHA-001": # ErgoFlex chair
            discount_pct = np.random.choice([0.05, 0.10, 0.15, 0.20], p=[0.3, 0.3, 0.25, 0.15])
        elif current_channel == "Wholesale":
            discount_pct = np.random.choice([0.10, 0.15, 0.20], p=[0.4, 0.4, 0.2])
        elif current_order_date.month in [11, 12]:
            discount_pct = np.random.choice([0.0, 0.05, 0.10, 0.15, 0.20], p=[0.35, 0.25, 0.20, 0.12, 0.08])
        else:
            discount_pct = np.random.choice([0.0, 0.05, 0.10, 0.15], p=[0.55, 0.25, 0.15, 0.05])
            
        unit_price = round(prod["price"], 2)
        # Unit Cost = Unit Price * (1 - Base Margin)
        unit_cost = round(unit_price * (1.0 - prod["margin"]), 2)
        
        # Derived calculations
        gross_sales = round(qty * unit_price, 2)
        discount_amount = round(gross_sales * discount_pct, 2)
        net_sales = round(gross_sales - discount_amount, 2)
        total_cost = round(qty * unit_cost, 2)
        profit = round(net_sales - total_cost, 2)
        profit_margin = round(profit / net_sales, 4) if net_sales > 0 else 0.0
        
        row = {
            "Order ID": current_order_id,
            "Order Date": current_order_date.strftime("%Y-%m-%d"),
            "Product ID": prod["id"],
            "Product Name": prod["name"],
            "Category": prod["category"],
            "Subcategory": prod["subcategory"],
            "Region": current_region,
            "State": current_state,
            "Sales Channel": current_channel,
            "Customer Segment": current_segment,
            "Quantity": qty,
            "Unit Price": unit_price,
            "Discount Percentage": discount_pct,
            "Unit Cost": unit_cost,
            "Gross Sales": gross_sales,
            "Discount Amount": discount_amount,
            "Net Sales": net_sales,
            "Total Cost": total_cost,
            "Profit": profit,
            "Profit Margin": profit_margin
        }
        rows.append(row)
        
    df = pd.DataFrame(rows)
    return df

df_clean = generate_dataset(1000)

# Sort chronologically by Order Date, then Order ID
df_clean = df_clean.sort_values(by=["Order Date", "Order ID"]).reset_index(drop=True)

# Save Clean Dataset to data/sales_data.csv
df_clean.to_csv("data/sales_data.csv", index=False)
print(f"Generated {len(df_clean)} clean records across {df_clean['Order ID'].nunique()} unique orders.")

# Now create Raw Data with deliberate, documented anomalies for Data_Quality demonstration
df_raw = df_clean.copy()

# 1. Duplicate records (4 identical duplicate rows appended)
dup_rows = df_raw.iloc[[12, 145, 410, 890]].copy()
df_raw = pd.concat([df_raw, dup_rows], ignore_index=True)

# 2. Inconsistent text formatting / leading whitespace
df_raw.loc[25, "Category"] = " Technology "
df_raw.loc[78, "Region"] = "west "
df_raw.loc[190, "Sales Channel"] = "ONLINE"

# 3. Missing non-critical values
df_raw.loc[50, "Customer Segment"] = np.nan
df_raw.loc[312, "Customer Segment"] = None

# 4. Erroneous record (negative quantity typo to be flagged)
df_raw.loc[512, "Quantity"] = -3

# Save Raw Dataset to data/sales_data_raw.csv
df_raw.to_csv("data/sales_data_raw.csv", index=False)
print(f"Generated {len(df_raw)} raw records with 10 deliberate anomalies for Data_Quality audit.")

# Output Key Metrics for analysis
total_net_sales = df_clean["Net Sales"].sum()
total_gross_sales = df_clean["Gross Sales"].sum()
total_discount_amt = df_clean["Discount Amount"].sum()
total_cost = df_clean["Total Cost"].sum()
total_profit = df_clean["Profit"].sum()
profit_margin = total_profit / total_net_sales
total_orders = df_clean["Order ID"].nunique()
total_units = df_clean["Quantity"].sum()
aov = total_net_sales / total_orders
sales_weighted_discount = total_discount_amt / total_gross_sales
simple_avg_discount = df_clean["Discount Percentage"].mean()

print("\n--- EXECUTIVE KPI SUMMARY ---")
print(f"Total Gross Sales: ${total_gross_sales:,.2f}")
print(f"Total Net Sales: ${total_net_sales:,.2f}")
print(f"Total Cost: ${total_cost:,.2f}")
print(f"Total Profit: ${total_profit:,.2f}")
print(f"Profit Margin: {profit_margin:.2%}")
print(f"Total Unique Orders: {total_orders:,}")
print(f"Total Units Sold: {total_units:,}")
print(f"Average Order Value (AOV): ${aov:,.2f}")
print(f"Sales-Weighted Discount: {sales_weighted_discount:.2%}")
print(f"Simple Avg Discount: {simple_avg_discount:.2%}")

print("\n--- CATEGORY BREAKDOWN ---")
cat_summary = df_clean.groupby("Category").agg(
    Net_Sales=("Net Sales", "sum"),
    Profit=("Profit", "sum"),
    Units=("Quantity", "sum")
).reset_index()
cat_summary["Margin%"] = cat_summary["Profit"] / cat_summary["Net_Sales"]
print(cat_summary)

print("\n--- REGIONAL BREAKDOWN ---")
reg_summary = df_clean.groupby("Region").agg(
    Net_Sales=("Net Sales", "sum"),
    Profit=("Profit", "sum"),
    Units=("Quantity", "sum")
).reset_index()
reg_summary["Margin%"] = reg_summary["Profit"] / reg_summary["Net_Sales"]
print(reg_summary)

print("\n--- SALES CHANNEL BREAKDOWN ---")
chan_summary = df_clean.groupby("Sales Channel").agg(
    Net_Sales=("Net Sales", "sum"),
    Profit=("Profit", "sum"),
    Orders=("Order ID", "nunique")
).reset_index()
chan_summary["Share%"] = chan_summary["Net_Sales"] / total_net_sales
chan_summary["AOV"] = chan_summary["Net_Sales"] / chan_summary["Orders"]
print(chan_summary)

print("\n--- TOP 5 PRODUCTS BY NET SALES ---")
top_sales = df_clean.groupby(["Product ID", "Product Name", "Category"]).agg(
    Net_Sales=("Net Sales", "sum"),
    Profit=("Profit", "sum"),
    Units=("Quantity", "sum")
).reset_index().sort_values(by="Net_Sales", ascending=False).head(10)
top_sales["Margin%"] = top_sales["Profit"] / top_sales["Net_Sales"]
print(top_sales[["Product Name", "Category", "Net_Sales", "Profit", "Margin%"]])

print("\n--- MONTHLY BREAKDOWN ---")
df_clean["Month"] = pd.to_datetime(df_clean["Order Date"]).dt.month
monthly_summary = df_clean.groupby("Month").agg(
    Net_Sales=("Net Sales", "sum"),
    Profit=("Profit", "sum"),
    Orders=("Order ID", "nunique")
).reset_index()
monthly_summary["MoM_Growth%"] = monthly_summary["Net_Sales"].pct_change()
print(monthly_summary)
