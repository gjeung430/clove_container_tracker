import pandas as pd
import json
import os
from datetime import datetime

EXCEL_FILE = "All Orders_Rolling_Pre-Orders.xlsx"

SIZE_COLS = [
    'W5', 'W5.5', 'W6', 'W6.5', 'W7', 'W7.5', 'W8', 'W8.5/M7', 'W9/M7.5',
    'W9.5/M8', 'W10/M8.5', 'W10.5/M9', 'W11/M9.5', 'W11.5/M10', 'W12/M10.5',
    'W12.5/M11', 'M11.5', 'M12', 'M13'
]

def get_restocked_sizes(row):
    sizes = []
    for sz in SIZE_COLS:
        val = row.get(sz, 0)
        try:
            if pd.notna(val) and float(val) > 0:
                sizes.append(sz)
        except:
            pass
    return sizes

def load_data():
    prod = pd.read_excel(EXCEL_FILE, sheet_name='Production Schedule', header=1)
    tracker = pd.read_excel(EXCEL_FILE, sheet_name='Container Tracker', header=8)

    tracker.columns = [str(c).strip() for c in tracker.columns]
    tracker['Warehouse Arrival Date'] = pd.to_datetime(tracker['Warehouse Arrival Date'], errors='coerce')
    tracker['ETD'] = pd.to_datetime(tracker['ETD'], errors='coerce')

    prod.columns = [str(c).strip().replace('\n', ' ') for c in prod.columns]
    prod.rename(columns={'SHIPPING  INVOICE NO': 'INVOICE', 'FTY NO': 'CONTAINER'}, inplace=True)
    prod['CONTAINER'] = prod['CONTAINER'].ffill()
    prod['INVOICE'] = prod['INVOICE'].ffill()
    prod['STYLE'] = prod['STYLE'].ffill()

    tracker_2026 = tracker[tracker['Warehouse Arrival Date'] >= '2026-01-01'].copy()
    invoice_map = {}
    for _, row in tracker_2026.iterrows():
        invoice = str(row['Supplier Invoice #']).strip()
        if invoice != 'nan':
            invoice_map[invoice] = {
                'container': str(row['Clove Container #']),
                'etd': row['ETD'].strftime('%b %d, %Y') if pd.notna(row['ETD']) else 'TBD',
                'arrival': row['Warehouse Arrival Date'].strftime('%b %d, %Y') if pd.notna(row['Warehouse Arrival Date']) else None,
                'arrival_raw': row['Warehouse Arrival Date'].strftime('%Y-%m-%d') if pd.notna(row['Warehouse Arrival Date']) else '9999',
            }

    prod['INVOICE_STR'] = prod['INVOICE'].astype(str).str.strip()
    prod_filtered = prod[prod['INVOICE_STR'].isin(invoice_map.keys())].copy()
    prod_filtered = prod_filtered[prod_filtered['Color'].notna() & (prod_filtered['Color'].astype(str).str.strip() != 'nan')]

    container_data = {}
    for _, row in prod_filtered.iterrows():
        inv = row['INVOICE_STR']
        info = invoice_map[inv]
        container = info['container']
        if container not in container_data:
            container_data[container] = {
                'container': container, 'invoice': inv,
                'etd': info['etd'], 'arrival': info['arrival'],
                'arrival_raw': info['arrival_raw'], 'items': []
            }
        sizes = get_restocked_sizes(row)
        if sizes:
            style_val = str(row.get('STYLE', '')).strip()
            if style_val == 'nan':
                style_val = ''
            color_val = str(row['Color']).strip()
            container_data[container]['items'].append({
                'style': style_val, 'color': color_val, 'sizes': sizes
            })

    sorted_containers = sorted(container_data.values(), key=lambda x: x['arrival_raw'])
    return sorted_containers

def generate_html(containers):
    today_str = datetime.now().strftime('%Y-%m-%d')
    containers_json = json.dumps(containers)
    last_updated = datetime.now().strftime('%b %d, %Y %H:%M')

    html = f'''<!DOCTYPE html>
<html>
<head>
<meta charset="UTF-8">
<title>Clove Container Tracker</title>
<style>
  * {{ box-sizing: border-box; }}
  body {{ font-family: Arial, sans-serif; background: #f8f9fa; padding: 24px; margin: 0; }}
  h2 {{ color: #1E3A5F; font-size: 18px; margin-bottom: 4px; text-align: center; }}
  .last-updated {{ text-align: center; font-size: 11px; color: #888; margin-bottom: 20px; }}
  .flow {{ display: flex; flex-direction: column; gap: 0; max-width: 660px; margin: 0 auto; }}
  .step {{ display: flex; align-items: flex-start; gap: 16px; background: white; border-radius: 12px; padding: 16px 20px; box-shadow: 0 2px 8px rgba(0,0,0,0.07); }}
  .step-num {{ width: 36px; height: 36px; border-radius: 50%; display: flex; align-items: center; justify-content: center; font-weight: bold; font-size: 16px; color: white; flex-shrink: 0; margin-top: 2px; }}
  .step-content h3 {{ margin: 0 0 4px 0; font-size: 15px; color: #1a1a2e; }}
  .step-content p {{ margin: 0; font-size: 13px; color: #555; line-height: 1.5; }}
  .step-content ul {{ margin: 6px 0 0 0; padding-left: 18px; font-size: 13px; color: #555; line-height: 1.7; }}
  .arrow {{ text-align: center; font-size: 22px; color: #aaa; margin: 4px 0; }}
  .tag {{ display: inline-block; border-radius: 6px; padding: 2px 8px; font-size: 11px; font-weight: bold; margin-top: 4px; margin-right: 4px; }}
  .green {{ background: #E8F5E9; color: #2E7D32; }}
  .purple {{ background: #F3E5F5; color: #6A1B9A; }}
  .orange {{ background: #FFF3E0; color: #E65100; }}
  .warning-box {{ background: #FFF8E1; border-left: 4px solid #F9A825; border-radius: 8px; padding: 14px 18px; max-width: 660px; margin: 12px auto 0 auto; font-size: 13px; color: #555; line-height: 1.7; }}
  .warning-box strong {{ color: #E65100; }}
  .tip-box {{ background: #E8F5E9; border-left: 4px solid #2E7D32; border-radius: 8px; padding: 14px 18px; max-width: 660px; margin: 10px auto 0 auto; font-size: 13px; color: #555; line-height: 1.7; }}
  .tip-box strong {{ color: #2E7D32; }}
  .restock-section {{ max-width: 660px; margin: 16px auto 0 auto; }}
  .restock-section h3 {{ color: #1E3A5F; font-size: 16px; margin: 0 0 12px 0; }}
  .tab-bar {{ display: flex; gap: 8px; margin-bottom: 12px; }}
  .tab-btn {{ padding: 6px 16px; border-radius: 20px; border: 1.5px solid #ccc; background: white; font-size: 12px; cursor: pointer; font-weight: bold; color: #555; transition: all 0.2s; }}
  .tab-btn.active {{ background: #1E3A5F; color: white; border-color: #1E3A5F; }}
  .container-card {{ background: white; border-radius: 10px; border: 1.5px solid #E0E0E0; margin-bottom: 10px; overflow: hidden; box-shadow: 0 1px 4px rgba(0,0,0,0.06); }}
  .container-card.arrived {{ border-left: 4px solid #2E7D32; }}
  .container-card.upcoming {{ border-left: 4px solid #F9A825; }}
  .container-header {{ display: flex; justify-content: space-between; align-items: center; padding: 12px 16px; cursor: pointer; user-select: none; }}
  .container-header:hover {{ background: #f9f9f9; }}
  .container-title {{ font-weight: bold; font-size: 14px; color: #1a1a2e; }}
  .container-meta {{ font-size: 11px; color: #888; margin-top: 2px; }}
  .container-badge {{ font-size: 11px; font-weight: bold; padding: 2px 8px; border-radius: 10px; }}
  .badge-arrived {{ background: #E8F5E9; color: #2E7D32; }}
  .badge-upcoming {{ background: #FFF3E0; color: #E65100; }}
  .container-body {{ display: none; padding: 0 16px 12px 16px; border-top: 1px solid #f0f0f0; }}
  .container-body.open {{ display: block; }}
  .item-row {{ display: flex; align-items: flex-start; gap: 8px; padding: 6px 0; border-bottom: 1px solid #f5f5f5; font-size: 12px; }}
  .item-row:last-child {{ border-bottom: none; }}
  .item-name {{ min-width: 180px; font-weight: 500; color: #333; }}
  .item-sizes {{ color: #1565C0; font-size: 11px; flex: 1; line-height: 1.6; }}
  .chevron {{ font-size: 12px; color: #aaa; transition: transform 0.2s; }}
  .count-badge {{ background: #EBF3FB; color: #1565C0; border-radius: 10px; padding: 1px 7px; font-size: 11px; font-weight: bold; }}
</style>
</head>
<body>
<h2>📦 Restock Flow</h2>
<p class="last-updated">Last updated: {last_updated}</p>

<div class="flow">
  <div class="step">
    <div class="step-num" style="background:#1565C0;">1</div>
    <div class="step-content">
      <h3>🚢 Container Arrives at Warehouse</h3>
      <p>Shipment arrives from AC.</p>
    </div>
  </div>
  <div class="arrow">↓</div>
  <div class="step" style="border: 2px solid #F9A825; background: #FFFDE7;">
    <div class="step-num" style="background:#F9A825;">!</div>
    <div class="step-content">
      <h3>📋 CX expects Ops to...</h3>
      <ul>
        <li>Update <strong>IPT (Incoming Product Tracker)</strong> date in the shipment tracker</li>
        <li>Remove <strong>pre-order hold</strong> in Shopify for any orders with a pre-order date set</li>
      </ul>
      <span class="tag orange">IPT Update</span>
      <span class="tag orange">Shopify Hold Remove</span>
    </div>
  </div>
  <div class="arrow">↓</div>
  <div class="step">
    <div class="step-num" style="background:#2E7D32;">2</div>
    <div class="step-content">
      <h3>⏱️ Received in Extensiv + NWM</h3>
      <p>Inventory received within <strong>48 hours max</strong>.</p>
      <span class="tag green">Extensiv</span>
      <span class="tag green">NWM</span>
    </div>
  </div>
  <div class="arrow">↓</div>
  <div class="step">
    <div class="step-num" style="background:#AD1457;">3</div>
    <div class="step-content">
      <h3>🛍️ Available in Shopify</h3>
      <p>Inventory available on Shopify <strong>next morning</strong> after receiving.</p>
      <span class="tag purple">Shopify</span>
    </div>
  </div>
</div>

<div class="warning-box">
  <strong>⚠️ NWM Queue Logic</strong><br><br>
  Orders <strong>not canceled in NWM</strong> stay in the <strong>allocated queue</strong> — even if they were canceled in WMS (RS SS list cancels).<br><br>
  These pending orders <strong>hold their spot in line</strong> → new orders go to the <strong>back of the queue</strong>.<br><br>
  <strong>FYI:</strong> WMS canceled orders = RS adds to SS list, but they stay <strong>pending in NWM</strong> unless manually canceled.
</div>

<div class="tip-box">
  <strong>✅ Why Clearing NWM + Shopify Matters</strong><br><br>
  If pending/unresolved orders are sitting in NWM, they take the restocked inventory first — leaving new orders waiting. Keeping NWM and Shopify clean = new orders get fulfilled faster!
</div>

<div class="restock-section">
  <h3>📋 Restock Tracker (2026)</h3>
  <div class="tab-bar">
    <button class="tab-btn active" onclick="filterContainers('all', this)">All</button>
    <button class="tab-btn" onclick="filterContainers('arrived', this)">✅ Arrived</button>
    <button class="tab-btn" onclick="filterContainers('upcoming', this)">🔜 Upcoming</button>
  </div>
  <div id="container-list"></div>
</div>

<script>
const TODAY = '{today_str}';
const containers = {containers_json};

function filterContainers(type, btn) {{
  document.querySelectorAll('.tab-btn').forEach(b => b.classList.remove('active'));
  btn.classList.add('active');
  renderContainers(type);
}}

function renderContainers(type = 'all') {{
  const list = document.getElementById('container-list');
  list.innerHTML = '';
  containers.forEach((c, i) => {{
    const isArrived = c.arrival_raw <= TODAY;
    if (type === 'arrived' && !isArrived) return;
    if (type === 'upcoming' && isArrived) return;
    const card = document.createElement('div');
    card.className = `container-card ${{isArrived ? 'arrived' : 'upcoming'}}`;
    const badgeClass = isArrived ? 'badge-arrived' : 'badge-upcoming';
    const badgeText = isArrived ? '✅ Arrived' : '🔜 Upcoming';
    const itemRows = c.items.map(item => `
      <div class="item-row">
        <div class="item-name">${{item.style ? `<strong>${{item.style}}</strong> — ` : ''}}${{item.color}}</div>
        <div class="item-sizes">${{item.sizes.join(', ')}}</div>
      </div>
    `).join('');
    card.innerHTML = `
      <div class="container-header" onclick="toggleCard(${{i}})">
        <div>
          <div class="container-title">${{c.container}} <span class="count-badge">${{c.items.length}} Colorways</span></div>
          <div class="container-meta">ETD: ${{c.etd}} &nbsp;|&nbsp; Warehouse Arrival: ${{c.arrival || 'TBD'}}</div>
        </div>
        <div style="display:flex;align-items:center;gap:8px;">
          <span class="container-badge ${{badgeClass}}">${{badgeText}}</span>
          <span class="chevron" id="chevron-${{i}}">▼</span>
        </div>
      </div>
      <div class="container-body" id="body-${{i}}">
        <div style="padding-top:8px;">${{itemRows}}</div>
      </div>
    `;
    list.appendChild(card);
  }});
}}

function toggleCard(i) {{
  const body = document.getElementById(`body-${{i}}`);
  const chevron = document.getElementById(`chevron-${{i}}`);
  body.classList.toggle('open');
  chevron.style.transform = body.classList.contains('open') ? 'rotate(180deg)' : '';
}}

renderContainers('all');
</script>
</body>
</html>'''
    return html

def main():
    print("Loading data from Excel...")
    containers = load_data()
    print(f"Found {len(containers)} containers")
    
    print("Generating HTML...")
    html = generate_html(containers)
    
    with open('restock_tracker.html', 'w') as f:
        f.write(html)
    print("HTML saved: restock_tracker.html")

if __name__ == '__main__':
    main()
