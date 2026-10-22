import json
import pandas as pd

df = pd.read_csv("data/sales_data.csv")

# Convert dataframe to a JSON records array
records = df.to_dict(orient="records")
data_json = json.dumps(records)

html_content = f"""<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>Executive Sales & Profitability KPI Dashboard | Microsoft Excel Portfolio</title>
  <link rel="preconnect" href="https://fonts.googleapis.com">
  <link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
  <link href="https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@400;500;600;700;800&display=swap" rel="stylesheet">
  <script src="https://cdn.jsdelivr.net/npm/chart.js"></script>
  <style>
    :root {{
      --navy-900: #0f172a;
      --navy-800: #1e293b;
      --navy-700: #334155;
      --blue-600: #2563eb;
      --blue-500: #3b82f6;
      --teal-600: #0d9488;
      --teal-500: #14b8a6;
      --emerald-600: #059669;
      --amber-500: #f59e0b;
      --rose-500: #f43f5e;
      --bg-slate: #f8fafc;
      --card-bg: #ffffff;
      --border-slate: #e2e8f0;
      --text-main: #0f172a;
      --text-muted: #64748b;
    }}

    * {{
      box-sizing: border-box;
      margin: 0;
      padding: 0;
    }}

    body {{
      font-family: 'Plus Jakarta Sans', sans-serif;
      background-color: var(--bg-slate);
      color: var(--text-main);
      line-height: 1.5;
      padding: 24px;
    }}

    .container {{
      max-width: 1440px;
      margin: 0 auto;
    }}

    /* Top Navigation / Header */
    header {{
      background: linear-gradient(135deg, #1e293b 0%, #0f172a 100%);
      color: #ffffff;
      padding: 24px 32px;
      border-radius: 16px;
      box-shadow: 0 10px 25px -5px rgba(15, 23, 42, 0.2);
      margin-bottom: 24px;
      display: flex;
      flex-wrap: wrap;
      justify-content: space-between;
      align-items: center;
      gap: 16px;
    }}

    .brand-title h1 {{
      font-size: 22px;
      font-weight: 800;
      letter-spacing: -0.02em;
      margin-bottom: 4px;
    }}

    .brand-title p {{
      color: #94a3b8;
      font-size: 13px;
    }}

    .header-actions {{
      display: flex;
      gap: 12px;
      align-items: center;
    }}

    .btn {{
      padding: 10px 18px;
      font-size: 13px;
      font-weight: 600;
      border-radius: 8px;
      cursor: pointer;
      text-decoration: none;
      display: inline-flex;
      align-items: center;
      gap: 8px;
      transition: all 0.2s ease;
      border: none;
    }}

    .btn-primary {{
      background-color: var(--blue-600);
      color: white;
    }}

    .btn-primary:hover {{
      background-color: #1d4ed8;
      transform: translateY(-1px);
    }}

    .btn-secondary {{
      background-color: #334155;
      color: #e2e8f0;
    }}

    .btn-secondary:hover {{
      background-color: #475569;
    }}

    /* Slicers / Interactive Filters Section */
    .filter-panel {{
      background: white;
      padding: 18px 24px;
      border-radius: 14px;
      border: 1px solid var(--border-slate);
      box-shadow: 0 2px 6px rgba(0,0,0,0.02);
      margin-bottom: 24px;
      display: grid;
      grid-template-columns: repeat(auto-fit, minmax(220px, 1fr));
      gap: 16px;
      align-items: center;
    }}

    .filter-group {{
      display: flex;
      flex-direction: column;
      gap: 6px;
    }}

    .filter-label {{
      font-size: 11px;
      font-weight: 700;
      text-transform: uppercase;
      letter-spacing: 0.05em;
      color: var(--text-muted);
    }}

    .filter-select {{
      padding: 9px 14px;
      font-size: 13px;
      font-weight: 600;
      color: var(--text-main);
      background-color: #f8fafc;
      border: 1px solid var(--border-slate);
      border-radius: 8px;
      outline: none;
      cursor: pointer;
      transition: all 0.15s ease;
    }}

    .filter-select:focus {{
      border-color: var(--blue-600);
      box-shadow: 0 0 0 3px rgba(37, 99, 235, 0.15);
    }}

    .filter-reset {{
      padding: 9px 16px;
      align-self: flex-end;
      background: #f1f5f9;
      color: #475569;
      border: 1px solid #cbd5e1;
      border-radius: 8px;
      font-size: 12px;
      font-weight: 700;
      cursor: pointer;
    }}

    .filter-reset:hover {{
      background: #e2e8f0;
      color: #0f172a;
    }}

    /* KPI Cards Grid */
    .kpi-grid {{
      display: grid;
      grid-template-columns: repeat(auto-fit, minmax(200px, 1fr));
      gap: 16px;
      margin-bottom: 24px;
    }}

    .kpi-card {{
      background: white;
      padding: 18px 20px;
      border-radius: 12px;
      border: 1px solid var(--border-slate);
      position: relative;
      overflow: hidden;
      box-shadow: 0 2px 4px rgba(0,0,0,0.02);
      transition: transform 0.2s ease, box-shadow 0.2s ease;
    }}

    .kpi-card:hover {{
      transform: translateY(-2px);
      box-shadow: 0 8px 16px rgba(0,0,0,0.05);
    }}

    .kpi-card::before {{
      content: '';
      position: absolute;
      top: 0;
      left: 0;
      right: 0;
      height: 4px;
    }}

    .kpi-card.sales::before {{ background: var(--blue-600); }}
    .kpi-card.orders::before {{ background: var(--teal-600); }}
    .kpi-card.profit::before {{ background: var(--emerald-600); }}
    .kpi-card.margin::before {{ background: #059669; }}
    .kpi-card.aov::before {{ background: #6366f1; }}
    .kpi-card.discount::before {{ background: var(--amber-500); }}

    .kpi-title {{
      font-size: 11px;
      font-weight: 700;
      text-transform: uppercase;
      color: var(--text-muted);
      letter-spacing: 0.05em;
      margin-bottom: 6px;
    }}

    .kpi-value {{
      font-size: 24px;
      font-weight: 800;
      color: var(--navy-900);
      letter-spacing: -0.02em;
      margin-bottom: 4px;
    }}

    .kpi-subtext {{
      font-size: 11px;
      color: var(--text-muted);
      font-weight: 500;
    }}

    /* Charts Grid */
    .charts-grid {{
      display: grid;
      grid-template-columns: repeat(auto-fit, minmax(430px, 1fr));
      gap: 20px;
      margin-bottom: 28px;
    }}

    .chart-card {{
      background: white;
      border-radius: 14px;
      border: 1px solid var(--border-slate);
      padding: 20px;
      box-shadow: 0 2px 6px rgba(0,0,0,0.02);
      display: flex;
      flex-direction: column;
    }}

    .chart-header {{
      display: flex;
      justify-content: space-between;
      align-items: center;
      margin-bottom: 16px;
    }}

    .chart-title {{
      font-size: 14px;
      font-weight: 700;
      color: var(--navy-800);
    }}

    .chart-tag {{
      font-size: 11px;
      font-weight: 600;
      padding: 3px 8px;
      border-radius: 6px;
      background: #eff6ff;
      color: var(--blue-600);
    }}

    .chart-container {{
      position: relative;
      flex: 1;
      min-height: 260px;
      max-height: 300px;
    }}

    /* Callout Banner */
    .callout-box {{
      background: #eff6ff;
      border-left: 4px solid var(--blue-600);
      padding: 16px 20px;
      border-radius: 8px;
      margin-bottom: 24px;
      display: flex;
      align-items: center;
      justify-content: space-between;
      gap: 16px;
    }}

    .callout-text h4 {{
      font-size: 14px;
      font-weight: 700;
      color: #1e3a8a;
      margin-bottom: 2px;
    }}

    .callout-text p {{
      font-size: 12.5px;
      color: #1e40af;
    }}

    footer {{
      text-align: center;
      padding: 24px;
      color: var(--text-muted);
      font-size: 12px;
    }}

    @media (max-width: 768px) {{
      .charts-grid {{
        grid-template-columns: 1fr;
      }}
      header {{
        flex-direction: column;
        align-items: flex-start;
      }}
    }}
  </style>
</head>
<body>

<div class="container">

  <!-- Header -->
  <header>
    <div class="brand-title">
      <h1>Executive Sales & Profitability KPI Dashboard</h1>
      <p>Multi-Channel Retail Performance Analytics | FY 2024 Reconciled Dataset</p>
    </div>
    <div class="header-actions">
      <a href="workbook/Sales_Analysis_Dashboard.xlsx" class="btn btn-primary" download>
        <svg width="16" height="16" fill="currentColor" viewBox="0 0 16 16"><path d="M.5 9.9a.5.5 0 0 1 .5.5v2.5a1 1 0 0 0 1 1h12a1 1 0 0 0 1-1v-2.5a.5.5 0 0 1 1 0v2.5a2 2 0 0 1-2 2H2a2 2 0 0 1-2-2v-2.5a.5.5 0 0 1 .5-.5z"/><path d="M7.646 11.854a.5.5 0 0 0 .708 0l3-3a.5.5 0 0 0-.708-.708L8.5 10.293V1.5a.5.5 0 0 0-1 0v8.793L5.354 8.146a.5.5 0 1 0-.708.708l3 3z"/></svg>
        Download Excel Workbook (.xlsx)
      </a>
      <a href="data/sales_data.csv" class="btn btn-secondary" download>
        <svg width="16" height="16" fill="currentColor" viewBox="0 0 16 16"><path d="M14 14V4.5L9.5 0H4a2 2 0 0 0-2 2v12a2 2 0 0 0 2 2h8a2 2 0 0 0 2-2zM9.5 3A1.5 1.5 0 0 0 11 4.5h2V14a1 1 0 0 1-1 1H4a1 1 0 0 1-1-1V2a1 1 0 0 1 1-1h5.5v2z"/></svg>
        Download CSV
      </a>
    </div>
  </header>

  <!-- Notice Banner -->
  <div class="callout-box">
    <div class="callout-text">
      <h4>⚡ Live Interactive Slicer Demonstration</h4>
      <p>Filter by Category, Region, Channel, or Customer Segment below to watch all 6 KPI cards and 6 visual charts dynamically update in real time.</p>
    </div>
    <span class="chart-tag" id="record-count-badge">1,000 Records Active</span>
  </div>

  <!-- Slicers / Interactive Filters -->
  <div class="filter-panel">
    <div class="filter-group">
      <label class="filter-label" for="categoryFilter">Category</label>
      <select class="filter-select" id="categoryFilter">
        <option value="ALL">All Categories</option>
        <option value="Technology">Technology</option>
        <option value="Furniture">Furniture</option>
        <option value="Office Supplies">Office Supplies</option>
      </select>
    </div>

    <div class="filter-group">
      <label class="filter-label" for="regionFilter">Region</label>
      <select class="filter-select" id="regionFilter">
        <option value="ALL">All Regions</option>
        <option value="West">West</option>
        <option value="East">East</option>
        <option value="Central">Central</option>
        <option value="South">South</option>
      </select>
    </div>

    <div class="filter-group">
      <label class="filter-label" for="channelFilter">Sales Channel</label>
      <select class="filter-select" id="channelFilter">
        <option value="ALL">All Channels</option>
        <option value="Online">Online</option>
        <option value="Wholesale">Wholesale</option>
        <option value="Direct Sales">Direct Sales</option>
        <option value="Retail Partner">Retail Partner</option>
      </select>
    </div>

    <div class="filter-group">
      <label class="filter-label" for="segmentFilter">Customer Segment</label>
      <select class="filter-select" id="segmentFilter">
        <option value="ALL">All Segments</option>
        <option value="Consumer">Consumer</option>
        <option value="Corporate">Corporate</option>
        <option value="Home Office">Home Office</option>
      </select>
    </div>

    <button class="filter-reset" id="resetBtn">Reset All Filters</button>
  </div>

  <!-- 6 Top KPI Cards -->
  <div class="kpi-grid">
    <div class="kpi-card sales">
      <div class="kpi-title">Total Net Sales</div>
      <div class="kpi-value" id="kpi-net-sales">$1,042,276</div>
      <div class="kpi-subtext" id="kpi-gross-sales">Gross: $1.15M | Post-Discount</div>
    </div>

    <div class="kpi-card orders">
      <div class="kpi-title">Total Unique Orders</div>
      <div class="kpi-value" id="kpi-orders">730</div>
      <div class="kpi-subtext" id="kpi-units">4,191 Units | 1.37 items/order</div>
    </div>

    <div class="kpi-card profit">
      <div class="kpi-title">Total Net Profit</div>
      <div class="kpi-value" id="kpi-profit">$164,133</div>
      <div class="kpi-subtext" id="kpi-profit-target">+9.4% above $150K Target</div>
    </div>

    <div class="kpi-card margin">
      <div class="kpi-title">Profit Margin</div>
      <div class="kpi-value" id="kpi-margin">15.8%</div>
      <div class="kpi-subtext">Benchmark: 15.0% Goal</div>
    </div>

    <div class="kpi-card aov">
      <div class="kpi-title">Average Order Value</div>
      <div class="kpi-value" id="kpi-aov">$1,428</div>
      <div class="kpi-subtext">Revenue per unique order</div>
    </div>

    <div class="kpi-card discount">
      <div class="kpi-title">Sales-Weighted Discount</div>
      <div class="kpi-value" id="kpi-discount">9.7%</div>
      <div class="kpi-subtext" id="kpi-discount-dollars">$111,756 Promos Given</div>
    </div>
  </div>

  <!-- 6 Visual Charts -->
  <div class="charts-grid">
    <!-- Chart 1: Monthly Net Sales Trend -->
    <div class="chart-card">
      <div class="chart-header">
        <h3 class="chart-title">Monthly Net Sales Trend ($K)</h3>
        <span class="chart-tag">Seasonality</span>
      </div>
      <div class="chart-container">
        <canvas id="chartMonthlySales"></canvas>
      </div>
    </div>

    <!-- Chart 2: Net Sales by Category -->
    <div class="chart-card">
      <div class="chart-header">
        <h3 class="chart-title">Net Sales by Product Category ($K)</h3>
        <span class="chart-tag">Category Mix</span>
      </div>
      <div class="chart-container">
        <canvas id="chartCategorySales"></canvas>
      </div>
    </div>

    <!-- Chart 3: Profit by Region -->
    <div class="chart-card">
      <div class="chart-header">
        <h3 class="chart-title">Net Operating Profit by Region ($K)</h3>
        <span class="chart-tag">Geography</span>
      </div>
      <div class="chart-container">
        <canvas id="chartRegionalProfit"></canvas>
      </div>
    </div>

    <!-- Chart 4: Sales Channel Mix -->
    <div class="chart-card">
      <div class="chart-header">
        <h3 class="chart-title">Revenue Share by Sales Channel</h3>
        <span class="chart-tag">Distribution</span>
      </div>
      <div class="chart-container">
        <canvas id="chartChannelMix"></canvas>
      </div>
    </div>

    <!-- Chart 5: Top Products by Net Sales -->
    <div class="chart-card">
      <div class="chart-header">
        <h3 class="chart-title">Top Products by Net Sales ($K)</h3>
        <span class="chart-tag">Margin Diagnosis</span>
      </div>
      <div class="chart-container">
        <canvas id="chartTopProducts"></canvas>
      </div>
    </div>

    <!-- Chart 6: Monthly Order Volume -->
    <div class="chart-card">
      <div class="chart-header">
        <h3 class="chart-title">Monthly Unique Order Volume</h3>
        <span class="chart-tag">Volume Flow</span>
      </div>
      <div class="chart-container">
        <canvas id="chartOrderVolume"></canvas>
      </div>
    </div>
  </div>

  <footer>
    <p>Sales Data Analysis & Executive KPI Dashboard | Built with Microsoft Excel, Python & Dynamic Slicers | Portfolio Project</p>
  </footer>

</div>

<!-- Data payload & Dashboard Controller -->
<script>
  const RAW_DATA = {data_json};

  // State elements
  const categoryFilter = document.getElementById('categoryFilter');
  const regionFilter = document.getElementById('regionFilter');
  const channelFilter = document.getElementById('channelFilter');
  const segmentFilter = document.getElementById('segmentFilter');
  const resetBtn = document.getElementById('resetBtn');
  const recordCountBadge = document.getElementById('record-count-badge');

  // Chart instances
  let chart1, chart2, chart3, chart4, chart5, chart6;

  function filterData() {{
    const cat = categoryFilter.value;
    const reg = regionFilter.value;
    const chan = channelFilter.value;
    const seg = segmentFilter.value;

    return RAW_DATA.filter(row => {{
      const matchCat = cat === 'ALL' || row.Category === cat;
      const matchReg = reg === 'ALL' || row.Region === reg;
      const matchChan = chan === 'ALL' || row['Sales Channel'] === chan;
      const matchSeg = seg === 'ALL' || row['Customer Segment'] === seg;
      return matchCat && matchReg && matchChan && matchSeg;
    }});
  }}

  function updateDashboard() {{
    const data = filterData();
    recordCountBadge.textContent = `${{data.length.toLocaleString()}} Records Active`;

    // 1. Calculate KPIs
    const totalGross = data.reduce((sum, r) => sum + r['Gross Sales'], 0);
    const totalDisc = data.reduce((sum, r) => sum + r['Discount Amount'], 0);
    const totalNet = data.reduce((sum, r) => sum + r['Net Sales'], 0);
    const totalCost = data.reduce((sum, r) => sum + r['Total Cost'], 0);
    const totalProfit = data.reduce((sum, r) => sum + r['Profit'], 0);
    const totalUnits = data.reduce((sum, r) => sum + r['Quantity'], 0);
    const uniqueOrders = new Set(data.map(r => r['Order ID'])).size;

    const marginPct = totalNet > 0 ? (totalProfit / totalNet) * 100 : 0;
    const aov = uniqueOrders > 0 ? (totalNet / uniqueOrders) : 0;
    const salesWeightedDiscount = totalGross > 0 ? (totalDisc / totalGross) * 100 : 0;

    // Update DOM
    document.getElementById('kpi-net-sales').textContent = '$' + Math.round(totalNet).toLocaleString();
    document.getElementById('kpi-gross-sales').textContent = `Gross: $${{(totalGross / 1000000).toFixed(2)}}M | Post-Discount`;
    document.getElementById('kpi-orders').textContent = uniqueOrders.toLocaleString();
    document.getElementById('kpi-units').textContent = `${{totalUnits.toLocaleString()}} Units | ${{uniqueOrders > 0 ? (totalUnits / uniqueOrders).toFixed(2) : 0}} items/ord`;
    document.getElementById('kpi-profit').textContent = '$' + Math.round(totalProfit).toLocaleString();
    document.getElementById('kpi-margin').textContent = marginPct.toFixed(1) + '%';
    document.getElementById('kpi-aov').textContent = '$' + Math.round(aov).toLocaleString();
    document.getElementById('kpi-discount').textContent = salesWeightedDiscount.toFixed(1) + '%';
    document.getElementById('kpi-discount-dollars').textContent = `$${{Math.round(totalDisc).toLocaleString()}} Promos Given`;

    // 2. Aggregate Data for Charts
    // Monthly Aggregation
    const months = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];
    const monthlyNet = new Array(12).fill(0);
    const monthlyOrdersMap = Array.from({{length: 12}}, () => new Set());

    data.forEach(r => {{
      const d = new Date(r['Order Date']);
      const m = d.getMonth();
      monthlyNet[m] += r['Net Sales'];
      monthlyOrdersMap[m].add(r['Order ID']);
    }});
    const monthlyOrders = monthlyOrdersMap.map(s => s.size);

    // Category Aggregation
    const catMap = {{}};
    data.forEach(r => {{
      catMap[r.Category] = (catMap[r.Category] || 0) + r['Net Sales'];
    }});
    const catLabels = Object.keys(catMap);
    const catValues = catLabels.map(c => Math.round(catMap[c] / 1000));

    // Regional Profit
    const regMap = {{}};
    data.forEach(r => {{
      regMap[r.Region] = (regMap[r.Region] || 0) + r['Profit'];
    }});
    const regLabels = ["East", "West", "Central", "South"];
    const regValues = regLabels.map(r => Math.round((regMap[r] || 0) / 1000));

    // Channel Mix
    const chanMap = {{}};
    data.forEach(r => {{
      chanMap[r['Sales Channel']] = (chanMap[r['Sales Channel']] || 0) + r['Net Sales'];
    }});
    const chanLabels = Object.keys(chanMap);
    const chanValues = Object.values(chanMap);

    // Top Products
    const prodMap = {{}};
    data.forEach(r => {{
      prodMap[r['Product Name']] = (prodMap[r['Product Name']] || 0) + r['Net Sales'];
    }});
    const sortedProds = Object.entries(prodMap).sort((a,b) => b[1] - a[1]).slice(0, 7);
    const prodLabels = sortedProds.map(p => p[0].length > 20 ? p[0].substring(0, 19) + '...' : p[0]);
    const prodValues = sortedProds.map(p => Math.round(p[1] / 1000));

    // 3. Render or Update Charts
    updateCharts({{
      monthlyNet: monthlyNet.map(v => Math.round(v / 1000)),
      monthlyOrders,
      catLabels,
      catValues,
      regLabels,
      regValues,
      chanLabels,
      chanValues,
      prodLabels,
      prodValues
    }});
  }}

  function updateCharts(agg) {{
    // Chart 1: Monthly Line
    if (chart1) chart1.destroy();
    chart1 = new Chart(document.getElementById('chartMonthlySales'), {{
      type: 'line',
      data: {{
        labels: ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"],
        datasets: [{{
          label: 'Net Sales ($K)',
          data: agg.monthlyNet,
          borderColor: '#2563eb',
          backgroundColor: 'rgba(37, 99, 235, 0.1)',
          fill: true,
          tension: 0.3,
          borderWidth: 2.5,
          pointRadius: 4,
          pointBackgroundColor: '#1d4ed8'
        }}]
      }},
      options: {{
        responsive: true,
        maintainAspectRatio: false,
        plugins: {{ legend: {{ display: false }} }},
        scales: {{ y: {{ beginAtZero: true, grid: {{ color: '#f1f5f9' }} }} }}
      }}
    }});

    // Chart 2: Category Bars
    if (chart2) chart2.destroy();
    chart2 = new Chart(document.getElementById('chartCategorySales'), {{
      type: 'bar',
      data: {{
        labels: agg.catLabels,
        datasets: [{{
          data: agg.catValues,
          backgroundColor: ['#1e3a8a', '#3b82f6', '#93c5fd'],
          borderRadius: 6
        }}]
      }},
      options: {{
        responsive: true,
        maintainAspectRatio: false,
        plugins: {{ legend: {{ display: false }} }},
        scales: {{ y: {{ beginAtZero: true, grid: {{ color: '#f1f5f9' }} }} }}
      }}
    }});

    // Chart 3: Regional Profit
    if (chart3) chart3.destroy();
    chart3 = new Chart(document.getElementById('chartRegionalProfit'), {{
      type: 'bar',
      data: {{
        labels: agg.regLabels,
        datasets: [{{
          data: agg.regValues,
          backgroundColor: ['#0d9488', '#14b8a6', '#2dd4bf', '#5eead4'],
          borderRadius: 6
        }}]
      }},
      options: {{
        responsive: true,
        maintainAspectRatio: false,
        plugins: {{ legend: {{ display: false }} }},
        scales: {{ y: {{ beginAtZero: true, grid: {{ color: '#f1f5f9' }} }} }}
      }}
    }});

    // Chart 4: Channel Donut
    if (chart4) chart4.destroy();
    chart4 = new Chart(document.getElementById('chartChannelMix'), {{
      type: 'doughnut',
      data: {{
        labels: agg.chanLabels,
        datasets: [{{
          data: agg.chanValues,
          backgroundColor: ['#2563eb', '#38bdf8', '#818cf8', '#c7d2fe'],
          borderWidth: 2,
          borderColor: '#ffffff'
        }}]
      }},
      options: {{
        responsive: true,
        maintainAspectRatio: false,
        plugins: {{ legend: {{ position: 'right', labels: {{ boxWidth: 12, font: {{ size: 11 }} }} }} }}
      }}
    }});

    // Chart 5: Top Products Bar
    if (chart5) chart5.destroy();
    chart5 = new Chart(document.getElementById('chartTopProducts'), {{
      type: 'bar',
      data: {{
        labels: agg.prodLabels,
        datasets: [{{
          data: agg.prodValues,
          backgroundColor: agg.prodLabels.map(l => l.includes('Gaming Laptop') ? '#ef4444' : '#475569'),
          borderRadius: 6
        }}]
      }},
      options: {{
        indexAxis: 'y',
        responsive: true,
        maintainAspectRatio: false,
        plugins: {{ legend: {{ display: false }} }},
        scales: {{ x: {{ beginAtZero: true, grid: {{ color: '#f1f5f9' }} }} }}
      }}
    }});

    // Chart 6: Order Volume
    if (chart6) chart6.destroy();
    chart6 = new Chart(document.getElementById('chartOrderVolume'), {{
      type: 'bar',
      data: {{
        labels: ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"],
        datasets: [{{
          data: agg.monthlyOrders,
          backgroundColor: '#6366f1',
          borderRadius: 6
        }}]
      }},
      options: {{
        responsive: true,
        maintainAspectRatio: false,
        plugins: {{ legend: {{ display: false }} }},
        scales: {{ y: {{ beginAtZero: true, grid: {{ color: '#f1f5f9' }} }} }}
      }}
    }});
  }}

  // Event Listeners for Filters
  [categoryFilter, regionFilter, channelFilter, segmentFilter].forEach(el => {{
    el.addEventListener('change', updateDashboard);
  }});

  resetBtn.addEventListener('click', () => {{
    categoryFilter.value = 'ALL';
    regionFilter.value = 'ALL';
    channelFilter.value = 'ALL';
    segmentFilter.value = 'ALL';
    updateDashboard();
  }});

  // Initial load
  updateDashboard();
</script>

</body>
</html>
"""

with open("dashboard.html", "w", encoding="utf-8") as f:
    f.write(html_content)
with open("index.html", "w", encoding="utf-8") as f:
    f.write(html_content)

print("Created live interactive HTML dashboards at dashboard.html and index.html!")
