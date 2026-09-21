import openpyxl
import json
import logging
import datetime

logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

def extract_data(file_path):
    wb = openpyxl.load_workbook(file_path, data_only=True)
    
    ws_data = wb['Dashboard']
    
    last_row = ws_data.max_row
    if last_row < 2:
        return {}
        
    latest_date_val = ws_data.cell(row=last_row, column=1).value
    latest_date = latest_date_val.strftime('%d-%b-%y') if hasattr(latest_date_val, 'strftime') else str(latest_date_val)
    
    etfs = [
        ("GOLDBETA", 2, 3),
        ("HDFCGOLD", 4, 5),
        ("GOLDIETF", 6, 7),
        ("BSLGOLDETF", 8, 9),
        ("LICMFGOLD", 10, 11)
    ]
    
    portfolio_val = 0
    prev_portfolio_val = 0
    first_portfolio_val = 0
    
    table_data = []
    
    # Calculate physical gold 1D return
    prev_row = last_row - 1 if last_row > 2 else 2
    gold_last = ws_data.cell(row=last_row, column=12).value or 0
    gold_prev = ws_data.cell(row=prev_row, column=12).value or 0
    gold_1d_pct = ((gold_last - gold_prev) / gold_prev * 100) if gold_prev else 0
    
    for name, c_close, c_val in etfs:
        curr_price = ws_data.cell(row=last_row, column=c_close).value or 0
        curr_val = ws_data.cell(row=last_row, column=c_val).value or 0
        
        prev_row = last_row - 1 if last_row > 2 else 2
        prev_price = ws_data.cell(row=prev_row, column=c_close).value or 0
        prev_val = ws_data.cell(row=prev_row, column=c_val).value or 0
        first_val = ws_data.cell(row=2, column=c_val).value or 0
        
        portfolio_val += curr_val
        prev_portfolio_val += prev_val
        first_portfolio_val += first_val
        
        c1d = curr_val - prev_val
        c1d_pct = (c1d / prev_val * 100) if prev_val else 0
        
        ws_etf = wb[name]
        nav_last = 0
        nav_prev = 0
        for r in range(ws_etf.max_row, 9, -1):
            d = ws_etf.cell(row=r, column=1).value
            if d:
                d_parsed = d.date() if hasattr(d, 'strftime') else d
                ld_parsed = latest_date_val.date() if hasattr(latest_date_val, 'strftime') else latest_date_val
                if d_parsed == ld_parsed:
                    nav_last = ws_etf.cell(row=r, column=4).value or 0
                    nav_prev = ws_etf.cell(row=r-1, column=4).value or 0
                    break
        nav_change = nav_last - nav_prev
        
        table_data.append({
            "etf": name,
            "price": float(curr_price),
            "prev_price": float(prev_price),
            "val": float(curr_val),
            "1d_change": float(c1d),
            "1d_change_pct": float(c1d_pct),
            "nav_change": float(nav_change),
            "nav_last": float(nav_last),
            "nav_prev": float(nav_prev),
            "tracking_diff": float(c1d_pct - gold_1d_pct)
        })
        
    for t in table_data:
        t["weight"] = (t["val"] / portfolio_val) if portfolio_val else 0
        
    oned_change = portfolio_val - prev_portfolio_val
    change_vs_first = portfolio_val - first_portfolio_val
    physical_gold = ws_data.cell(row=last_row, column=12).value or 0
        
    # 3. Extract Historical Trend Data from Tabular Dashboard
    historical_dates = []
    historical_series = {
        "GOLDBETA": [],
        "HDFCGOLD": [],
        "GOLDIETF": [],
        "BSLGOLDETF": [],
        "LICMFGOLD": []
    }
    
    # Col index in Dashboard:
    col_map = {
        "GOLDBETA": 3,
        "HDFCGOLD": 5,
        "GOLDIETF": 7,
        "BSLGOLDETF": 9,
        "LICMFGOLD": 11
    }
    
    for r in range(2, ws_data.max_row + 1):
        d = ws_data.cell(row=r, column=1).value
        if d:
            if hasattr(d, 'strftime'):
                d_str = d.strftime('%d-%b')
            else:
                d_str = str(d)
            historical_dates.append(d_str)
            
            for etf, col_idx in col_map.items():
                val = ws_data.cell(row=r, column=col_idx).value or 0
                historical_series[etf].append(float(val))
                
    return {
        "latest_date": latest_date,
        "portfolio_val": float(portfolio_val),
        "oned_change": float(oned_change),
        "change_vs_first": float(change_vs_first),
        "physical_gold": float(physical_gold),
        "table_data": table_data,
        "historical_dates": historical_dates,
        "historical_series": historical_series
    }

def generate_html(data):
    colors = ['#f59e0b', '#10b981', '#3b82f6', '#8b5cf6', '#ef4444']
    
    # Format currency helper
    def fmt(val):
        return f"₹{val:,.2f}"
        
    # Format pct helper
    def f_pct(val):
        return f"{val:,.2f}%"
        
    html = f"""<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1.0" />
  <title>Premium Gold ETF Dashboard</title>
  <link href="https://fonts.googleapis.com/css2?family=Outfit:wght@300;400;500;600;700;800&display=swap" rel="stylesheet">
  <script src="https://cdn.jsdelivr.net/npm/chart.js"></script>
  <style>
    :root {{
      --bg-base: #0f0c05;
      --bg-surface: #1a160c;
      --bg-surface-elevated: #292313;
      --text-primary: #fef3c7;
      --text-secondary: #d4d4d8;
      --gold-primary: #fbbf24;
      --gold-secondary: #f59e0b;
      --gold-dark: #b45309;
      --accent-green: #10b981;
      --accent-red: #ef4444;
      --border: rgba(251, 191, 36, 0.15);
      --card-radius: 20px;
    }}
    * {{ box-sizing: border-box; margin: 0; padding: 0; }}
    body {{
      font-family: 'Outfit', sans-serif;
      background-color: var(--bg-base);
      color: var(--text-primary);
      min-height: 100vh;
      padding: 30px;
      background-image: 
        radial-gradient(circle at 15% 50%, rgba(245, 158, 11, 0.08) 0%, transparent 50%),
        radial-gradient(circle at 85% 30%, rgba(251, 191, 36, 0.08) 0%, transparent 50%);
    }}
    .container {{ max-width: 1400px; margin: 0 auto; }}
    header {{
      display: flex; justify-content: space-between; align-items: center;
      margin-bottom: 40px; padding-bottom: 20px;
      border-bottom: 1px solid var(--border);
    }}
    .brand h1 {{
      font-size: 2.2rem; font-weight: 800;
      background: linear-gradient(135deg, #fef3c7, #f59e0b);
      -webkit-background-clip: text; -webkit-text-fill-color: transparent;
    }}
    .brand p {{ color: var(--text-secondary); margin-top: 5px; font-weight: 300; font-size: 1.1rem; }}
    .kpi-grid {{
      display: grid; grid-template-columns: repeat(auto-fit, minmax(280px, 1fr)); gap: 24px; margin-bottom: 40px;
    }}
    .card {{
      background: rgba(26, 22, 12, 0.6);
      backdrop-filter: blur(12px);
      border: 1px solid var(--border);
      border-radius: var(--card-radius);
      padding: 24px;
      transition: transform 0.3s cubic-bezier(0.4, 0, 0.2, 1), box-shadow 0.3s;
    }}
    .card:hover {{
      transform: translateY(-5px);
      box-shadow: 0 10px 30px -10px rgba(245, 158, 11, 0.2);
      border-color: rgba(251, 191, 36, 0.4);
    }}
    .kpi-label {{ font-size: 0.9rem; color: var(--gold-primary); text-transform: uppercase; letter-spacing: 1px; margin-bottom: 10px; font-weight: 600; }}
    .kpi-val {{ font-size: 2.5rem; font-weight: 700; }}
    .charts-grid {{
      display: grid; grid-template-columns: 2fr 1fr; gap: 24px; margin-bottom: 40px;
    }}
    @media (max-width: 900px) {{ .charts-grid {{ grid-template-columns: 1fr; }} }}
    .holdings-table {{
      width: 100%; border-collapse: collapse; margin-top: 20px;
    }}
    .holdings-table th, .holdings-table td {{
      padding: 16px; text-align: left; border-bottom: 1px solid var(--border);
    }}
    .holdings-table th {{ color: var(--gold-secondary); font-weight: 600; font-size: 0.9rem; text-transform: uppercase; letter-spacing: 1px; }}
    .holdings-table tr:hover td {{ background: rgba(251, 191, 36, 0.05); }}
    .pill {{ padding: 6px 12px; border-radius: 20px; font-size: 0.9rem; font-weight: 600; display: inline-block; }}
    .pill.up {{ background: rgba(16, 185, 129, 0.15); color: #34d399; border: 1px solid rgba(16, 185, 129, 0.3); }}
    .pill.down {{ background: rgba(239, 68, 68, 0.15); color: #f87171; border: 1px solid rgba(239, 68, 68, 0.3); }}
  </style>
</head>
<body>
  <div class="container">
    <header>
      <div class="brand">
        <h1>Gold ETF Vault</h1>
        <p>Premium Daily Portfolio Tracker</p>
      </div>
      <div style="text-align: right;">
        <p style="color: var(--text-secondary); font-weight: 300;">Latest Trading Date</p>
        <p style="color: var(--gold-primary); font-weight: 700; font-size: 1.2rem;">{data['latest_date']}</p>
      </div>
    </header>

    <div class="kpi-grid">
      <div class="card">
        <div class="kpi-label">Current Portfolio Value</div>
        <div class="kpi-val">{fmt(data['portfolio_val'])}</div>
      </div>
      <div class="card">
        <div class="kpi-label">1-Day Change</div>
        <div class="kpi-val" style="color: {'#34d399' if data['oned_change'] >= 0 else '#f87171'};">
          {'+' if data['oned_change'] >= 0 else ''}{fmt(data['oned_change']).replace('₹-', '-₹')}
        </div>
      </div>
      <div class="card">
        <div class="kpi-label">Change vs First Day</div>
        <div class="kpi-val" style="color: {'#34d399' if data['change_vs_first'] >= 0 else '#f87171'};">
          {'+' if data['change_vs_first'] >= 0 else ''}{fmt(data['change_vs_first']).replace('₹-', '-₹')}
        </div>
      </div>
      <div class="card" style="border-color: var(--gold-secondary);">
        <div class="kpi-label" style="color: var(--text-primary);">Physical Gold (10g)</div>
        <div class="kpi-val" style="color: var(--gold-secondary);">{fmt(data['physical_gold'])}</div>
      </div>
    </div>

    <div class="charts-grid">
      <div class="card">
        <h3 class="kpi-label" style="font-size: 1.1rem; margin-bottom: 20px;">Holding Value Trend</h3>
        <div style="position: relative; height: 350px; width: 100%;">
          <canvas id="lineChart"></canvas>
        </div>
      </div>

      <div class="card">
        <h3 class="kpi-label" style="font-size: 1.1rem; margin-bottom: 20px;">Capital Allocation</h3>
        <div style="position: relative; height: 350px; width: 100%; display: flex; justify-content: center;">
          <canvas id="donutChart"></canvas>
        </div>
      </div>
    </div>
    
    <div class="card" style="margin-bottom: 40px;">
      <h3 class="kpi-label" style="font-size: 1.1rem;">Detailed Asset Performance</h3>
      <div style="overflow-x: auto;">
      <table class="holdings-table">
        <thead>
          <tr>
            <th>ETF</th>
            <th>Last Traded Price</th>
            <th>Latest Price</th>
            <th>Last NAV</th>
            <th>Present NAV</th>
            <th>NAV Change</th>
            <th>1D Change</th>
            <th>1D Change %</th>
            <th>Tracking Diff</th>
            <th>Holding Value</th>
            <th>Weight %</th>
          </tr>
        </thead>
        <tbody>"""

    labels_pie = []
    data_pie = []
    
    for i, h in enumerate(data['table_data']):
        labels_pie.append(h['etf'])
        data_pie.append(h['weight'] * 100)
        
        c1d_class = "up" if h['1d_change'] >= 0 else "down"
        c1d_sign = "+" if h['1d_change'] >= 0 else ""
        nav_class = "up" if h['nav_change'] >= 0 else "down"
        nav_sign = "+" if h['nav_change'] >= 0 else ""
        td_class = "up" if h['tracking_diff'] >= 0 else "down"
        td_sign = "+" if h['tracking_diff'] >= 0 else ""
        
        html += f"""
          <tr>
            <td style="font-weight: 600; color: {colors[i%len(colors)]}; font-size: 1.1rem;">{h['etf']}</td>
            <td style="color: var(--text-secondary);">{fmt(h['prev_price'])}</td>
            <td style="font-size: 1.05rem; font-weight: 600;">{fmt(h['price'])}</td>
            <td style="color: var(--text-secondary);">{fmt(h['nav_prev'])}</td>
            <td style="font-size: 1.05rem;">{fmt(h['nav_last'])}</td>
            <td><span class="pill {nav_class}" style="background: transparent; border: none; padding: 0;">{nav_sign}{fmt(h['nav_change']).replace('₹-', '-₹')}</span></td>
            <td><span class="pill {c1d_class}">{c1d_sign}{fmt(h['1d_change']).replace('₹-', '-₹')}</span></td>
            <td style="color: {'#34d399' if h['1d_change_pct'] >= 0 else '#f87171'}; font-weight: 600;">{c1d_sign}{f_pct(h['1d_change_pct'])}</td>
            <td><span class="pill {td_class}">{td_sign}{f_pct(h['tracking_diff'])}</span></td>
            <td style="font-weight: 600; font-size: 1.1rem;">{fmt(h['val'])}</td>
            <td style="color: var(--text-secondary);">{f_pct(h['weight']*100)}</td>
          </tr>"""

    html += f"""
        </tbody>
      </table>
      </div>
    </div>
  </div>

  <script>
    Chart.defaults.color = '#d4d4d8';
    Chart.defaults.font.family = "'Outfit', sans-serif";
    Chart.defaults.borderColor = 'rgba(251, 191, 36, 0.1)';
    
    // Donut Chart
    const ctxDonut = document.getElementById('donutChart').getContext('2d');
    new Chart(ctxDonut, {{
      type: 'doughnut',
      data: {{
        labels: {json.dumps(labels_pie)},
        datasets: [{{
          data: {json.dumps(data_pie)},
          backgroundColor: {json.dumps(colors)},
          borderWidth: 0,
          hoverOffset: 15
        }}]
      }},
      options: {{
        responsive: true,
        maintainAspectRatio: false,
        cutout: '75%',
        plugins: {{
          legend: {{ position: 'bottom', labels: {{ padding: 20, font: {{ size: 12, weight: 600 }} }} }},
          tooltip: {{ padding: 15, cornerRadius: 10, titleFont: {{ size: 14 }}, bodyFont: {{ size: 14 }} }}
        }},
        animation: {{ animateScale: true, animateRotate: true }}
      }}
    }});
    
    // Line Chart
    const ctxLine = document.getElementById('lineChart').getContext('2d');
    
    // Create gradients for lines
    const lineColors = ['#f59e0b', '#10b981', '#3b82f6', '#8b5cf6', '#ef4444'];
    const datasets = [];
    const etfNames = {json.dumps(list(data['historical_series'].keys()))};
    const etfData = {json.dumps(list(data['historical_series'].values()))};
    
    for (let i = 0; i < etfNames.length; i++) {{
      datasets.push({{
        label: etfNames[i],
        data: etfData[i],
        borderColor: lineColors[i],
        backgroundColor: lineColors[i] + '20',
        borderWidth: 3,
        pointBackgroundColor: lineColors[i],
        pointBorderColor: '#0f0c05',
        pointBorderWidth: 2,
        pointRadius: 4,
        pointHoverRadius: 6,
        tension: 0.4,
        fill: false
      }});
    }}
    
    new Chart(ctxLine, {{
      type: 'line',
      data: {{
        labels: {json.dumps(data['historical_dates'])},
        datasets: datasets
      }},
      options: {{
        responsive: true,
        maintainAspectRatio: false,
        interaction: {{
          mode: 'index',
          intersect: false,
        }},
        plugins: {{
          legend: {{ position: 'top', labels: {{ usePointStyle: true, padding: 20, font: {{ size: 12 }} }} }},
          tooltip: {{ padding: 15, cornerRadius: 10, titleFont: {{ size: 14 }}, bodyFont: {{ size: 14 }} }}
        }},
        scales: {{
          y: {{
            beginAtZero: false,
            grid: {{ color: 'rgba(251, 191, 36, 0.05)' }}
          }},
          x: {{
            grid: {{ display: false }}
          }}
        }}
      }}
    }});
  </script>
</body>
</html>"""

    with open("index.html", "w", encoding="utf-8") as f:
        f.write(html)
        
    logging.info("HTML Dashboard generated successfully!")

if __name__ == "__main__":
    file_path = "Gold_ETF_Daily_Price_Tracker.xlsx"
    data = extract_data(file_path)
    generate_html(data)
