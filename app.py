"""
Pharma HTML Report → Excel Converter
S.F. Medical Agency — FoxPro/Tally fixed-width absolute-positioned HTML parser

Key insight: HTML has multiple <div class="page"> blocks with position:relative,
so top values RESET per page. Must process each page independently.
Item rows use fixed-width \xa0 padding. Last 5 tokens = Qty, Free, Rate, Amount, %.
"""

import io
import re
from flask import Flask, request, jsonify, send_file, render_template
from bs4 import BeautifulSoup
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 32 * 1024 * 1024

# ── Special characters used in this HTML ──────────────────────────────────────
NBSP  = '\xa0'    # non-breaking space (used as column padding)
NBHY  = '\u2011'  # non-breaking hyphen (used in item names like SP-CLAV)
RSQUO = '\u2019'  # right single quote (SPECIALITY'S)

def normalize(raw: str) -> str:
    return raw.replace(NBSP, ' ').replace(NBHY, '-').replace(RSQUO, "'").strip()

# ── Line classification ────────────────────────────────────────────────────────
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
        if _NUM.match(tok): return False   # numeric token = item row, not party
    alpha = [c for c in text if c.isalpha()]
    if not alpha: return False
    return sum(1 for c in alpha if c.isupper()) / len(alpha) >= 0.85

# ── Item row parser ────────────────────────────────────────────────────────────
def parse_item(raw: str) -> dict | None:
    """
    Each item row ends with 5 tokens: Qty  Free  Rate  Amount  Pct
    Everything before = item name (name + pack size joined).
    '-' in Free slot → stored as None (blank cell in Excel).
    """
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
    # trailing[4] is the % column — not needed

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

# ── HTML parser ────────────────────────────────────────────────────────────────
def get_top(div) -> float:
    m = re.search(r'top\s*:\s*([\d.]+)', div.get('style', ''))
    return float(m.group(1)) if m else 0.0

def parse_html(html_content: str) -> list[dict]:
    """
    CRITICAL: The HTML has 4 <div class="page"> blocks, each with
    position:relative — so top values reset to 0 at each new page.
    We MUST process each page separately (sort by top within the page).
    Flattening all divs across pages (old approach) causes same-top divs
    from different pages to collide and corrupt the output.
    """
    soup  = BeautifulSoup(html_content, 'html.parser')
    pages = soup.find_all('div', class_='page')

    rows, current_party = [], None

    for page in pages:
        divs = sorted(page.find_all('div', style=True), key=get_top)

        for div in divs:
            raw  = div.get_text()
            text = normalize(raw)

            if not text:            continue
            if is_separator(text):  continue
            if is_skip(text):       continue
            if 'TOTAL' in text:     continue   # TOTAL / GRAND TOTAL

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
                        'Free':       parsed['free'],   # None = blank
                        'Rate':       parsed['rate'],
                        'Amount':     parsed['amount'],
                    })
    return rows

# ── Excel builder ──────────────────────────────────────────────────────────────
def build_excel(rows: list[dict]) -> bytes:
    wb = Workbook()
    ws = wb.active
    ws.title = "Sales Report"

    hdr_font   = Font(name='Arial', bold=True, color='FFFFFF', size=11)
    hdr_fill   = PatternFill('solid', start_color='1B4F72')
    party_font = Font(name='Arial', bold=True, color='1B2631', size=10)
    party_fill = PatternFill('solid', start_color='D6EAF8')
    data_font  = Font(name='Arial', size=10)
    alt_fill   = PatternFill('solid', start_color='F8FBFD')
    thin       = Side(style='thin', color='CCCCCC')
    bdr        = Border(left=thin, right=thin, top=thin, bottom=thin)
    C = Alignment(horizontal='center', vertical='center')
    L = Alignment(horizontal='left',   vertical='center')
    R = Alignment(horizontal='right',  vertical='center')

    headers    = ['Party Name', 'Item Name', 'Qty', 'Free', 'Rate', 'Amount']
    col_widths = [30, 38, 10, 10, 12, 14]

    for ci, (h, w) in enumerate(zip(headers, col_widths), 1):
        cell = ws.cell(row=1, column=ci, value=h)
        cell.font = hdr_font; cell.fill = hdr_fill
        cell.border = bdr;    cell.alignment = C
        ws.column_dimensions[get_column_letter(ci)].width = w
    ws.row_dimensions[1].height = 22

    row_num, prev_party, alt = 2, None, False

    for rec in rows:
        party = rec['Party Name']

        if party != prev_party:
            ws.merge_cells(start_row=row_num, start_column=1,
                           end_row=row_num,   end_column=6)
            pc = ws.cell(row=row_num, column=1, value=party)
            pc.font = party_font; pc.fill = party_fill
            pc.border = bdr;      pc.alignment = L
            ws.row_dimensions[row_num].height = 18
            row_num += 1; prev_party = party; alt = False

        fill   = PatternFill('solid', start_color='FFFFFF') if not alt else alt_fill
        free_v = rec['Free']  # None → blank cell
        values = [rec['Party Name'], rec['Item Name'], rec['Qty'],
                  free_v if free_v is not None else '',
                  rec['Rate'], rec['Amount']]
        aligns = [L, L, C, C, R, R]
        fmts   = [None, None, '#,##0', '#,##0', '#,##0.00', '#,##0.00']

        for ci, (v, a, f) in enumerate(zip(values, aligns, fmts), 1):
            cell = ws.cell(row=row_num, column=ci, value=v)
            cell.font = data_font; cell.fill = fill
            cell.border = bdr;    cell.alignment = a
            if f and v != '': cell.number_format = f

        ws.row_dimensions[row_num].height = 16
        row_num += 1; alt = not alt

    ws.freeze_panes = 'A2'
    ws.auto_filter.ref = f'A1:{get_column_letter(len(headers))}1'

    buf = io.BytesIO()
    wb.save(buf); buf.seek(0)
    return buf.read()

# ── Flask routes ───────────────────────────────────────────────────────────────
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
