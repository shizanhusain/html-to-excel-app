"""
Pharma HTML Report → Excel Converter
S.F. Medical Agency — FoxPro/Tally fixed-width absolute-positioned HTML parser
"""

import io
import re
from flask import Flask, request, jsonify, send_file, render_template
from bs4 import BeautifulSoup
from openpyxl import Workbook

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 32 * 1024 * 1024

NBSP  = '\xa0'
NBHY  = '\u2011'
RSQUO = '\u2019'

def normalize(raw: str) -> str:
    return raw.replace(NBSP, ' ').replace(NBHY, '-').replace(RSQUO, "'").strip()

SKIP_PHRASES = [
    'S.F.MEDICAL', 'PARTY / ITEM', 'PARTY/ ITEM', 'Report For',
    'Company', 'D E S C R I P T I O N', 'Continued..', 'Page No',
    'GSTIN', 'Phone', 'End of Report', 'GRAND TOTAL',
]

_NUM  = re.compile(r'^-?\d+\.?\d*$')
_DASH = re.compile(r'^-$')
_SEP  = re.compile(r'^[-\u2011=\s]+$')

def is_separator(t): return bool(_SEP.match(t))
def is_skip(t):      return any(p in t for p in SKIP_PHRASES)

def is_party(text: str) -> bool:
    if is_separator(text): return False
    tokens = text.split()
    if not tokens: return False
    for tok in tokens:
        if _NUM.match(tok): return False
    alpha = [c for c in text if c.isalpha()]
    if not alpha: return False
    return sum(1 for c in alpha if c.isupper()) / len(alpha) >= 0.85

def parse_item(raw: str) -> dict | None:
    text   = normalize(raw)
    tokens = text.split()
    if not tokens: return None

    trailing, i = [], len(tokens) - 1
    while i >= 0 and len(trailing) < 5:
        tok = tokens[i]
        if _NUM.match(tok) or _DASH.match(tok):
            trailing.insert(0, tok); i -= 1
        else:
            break

    if len(trailing) < 5: return None

    qty_t, free_t, rate_t, amt_t = trailing[0], trailing[1], trailing[2], trailing[3]

    try:
        qty    = float(qty_t)
        free   = None if free_t == '-' else float(free_t)
        rate   = float(rate_t)
        amount = float(amt_t)
    except ValueError:
        return None

    item_name = ' '.join(tokens[:len(tokens) - 5]).strip()
    if not item_name: return None

    return {'item_name': item_name, 'qty': qty, 'free': free,
            'rate': rate, 'amount': amount}

def get_top(div) -> float:
    m = re.search(r'top\s*:\s*([\d.]+)', div.get('style', ''))
    return float(m.group(1)) if m else 0.0

def parse_html(html_content: str) -> list[dict]:
    soup  = BeautifulSoup(html_content, 'html.parser')
    pages = soup.find_all('div', class_='page')

    rows, current_party = [], None

    for page in pages:
        divs = sorted(page.find_all('div', style=True), key=get_top)

        for div in divs:
            raw  = div.get_text()
            text = normalize(raw)

            if not text:           continue
            if is_separator(text): continue
            if is_skip(text):      continue
            if 'TOTAL' in text:    continue

            if is_party(text):
                current_party = text
                continue

            if current_party:
                parsed = parse_item(raw)
                if parsed:
                    rows.append({
                        'Party Name': current_party,
                        'Item Name':  parsed['item_name'],
                        'Qty':        parsed['qty'],
                        'Free':       parsed['free'],
                        'Rate':       parsed['rate'],
                        'Amount':     parsed['amount'],
                    })
    return rows

def build_excel(rows: list[dict]) -> bytes:
    wb = Workbook()
    ws = wb.active
    ws.title = "Sales Report"

    # Plain flat header
    ws.append(['Party Name', 'Item Name', 'Qty', 'Free', 'Rate', 'Amount'])

    # One row per item, party name repeats on every row, no styling
    for rec in rows:
        ws.append([
            rec['Party Name'],
            rec['Item Name'],
            rec['Qty'],
            rec['Free'],   # None = blank cell
            rec['Rate'],
            rec['Amount'],
        ])

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.read()

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/parse', methods=['POST'])
def parse_files():
    files = request.files.getlist('files')
    if not files:
        return jsonify({'error': 'No files uploaded'}), 400

    all_rows, file_stats = [], []

    for f in files:
        if not f.filename: continue
        try:
            html  = f.read().decode('windows-1252', errors='replace')
            rows  = parse_html(html)
            all_rows.extend(rows)
            file_stats.append({'name': f.filename, 'rows': len(rows),
                               'parties': len(set(r['Party Name'] for r in rows)),
                               'status': 'ok'})
        except Exception as e:
            file_stats.append({'name': f.filename, 'rows': 0,
                               'parties': 0, 'status': f'error: {e}'})

    if not all_rows:
        return jsonify({'error': 'No data extracted. Check file format.',
                        'file_stats': file_stats}), 422

    serial = [{'Party Name': r['Party Name'], 'Item Name': r['Item Name'],
               'Qty': r['Qty'], 'Free': r['Free'],
               'Rate': r['Rate'], 'Amount': r['Amount']} for r in all_rows]

    return jsonify({'total_rows': len(serial),
                    'total_parties': len(set(r['Party Name'] for r in serial)),
                    'file_stats': file_stats,
                    'preview': serial[:200],
                    'data': serial})

@app.route('/download', methods=['POST'])
def download():
    payload  = request.get_json(force=True)
    rows     = payload.get('data', [])
    filename = payload.get('filename', 'pharma_report')
    if not rows:
        return jsonify({'error': 'No data to export'}), 400

    xlsx = build_excel(rows)
    buf  = io.BytesIO(xlsx); buf.seek(0)
    safe = re.sub(r'[^\w\-.]', '_', filename)
    return send_file(buf,
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        as_attachment=True, download_name=f'{safe}.xlsx')

if __name__ == '__main__':
    app.run(debug=True, port=5050)
