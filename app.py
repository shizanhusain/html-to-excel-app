"""
Pharma HTML Report → Excel Converter
S.F. Medical Agency — Tally/FoxPro absolute-positioned HTML parser
"""

import io
import re
import json
import pandas as pd
from flask import Flask, request, jsonify, send_file, render_template
from bs4 import BeautifulSoup
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 32 * 1024 * 1024  # 32 MB max upload


# ─────────────────────────────────────────────
#  PARSER
# ─────────────────────────────────────────────

def extract_top(style: str) -> float:
    """Parse 'top:NNpx' from an inline style string. Returns 0 if not found."""
    m = re.search(r'top\s*:\s*([\d.]+)', style or '')
    return float(m.group(1)) if m else 0.0


def extract_left(style: str) -> float:
    """Parse 'left:NNpx' from an inline style string. Returns 0 if not found."""
    m = re.search(r'left\s*:\s*([\d.]+)', style or '')
    return float(m.group(1)) if m else 0.0


def get_divs_sorted(soup: BeautifulSoup):
    """Return all divs with inline style, sorted by (top, left) for visual reading order."""
    divs = []
    for div in soup.find_all('div', style=True):
        text = div.get_text(separator=' ', strip=True)
        if not text:
            continue
        style = div.get('style', '')
        top = extract_top(style)
        left = extract_left(style)
        divs.append({'top': top, 'left': left, 'text': text})

    # Sort by top position first, then left for same-row elements
    divs.sort(key=lambda d: (round(d['top'], 0), d['left']))
    return divs


def group_into_lines(divs, y_tolerance: float = 4.0) -> list[str]:
    """
    Cluster divs that share similar top values into logical lines.
    Within each line, sort by left position and join text fragments.
    y_tolerance: pixel window within which two divs are considered the same line.
    """
    if not divs:
        return []

    lines = []
    current_group = [divs[0]]

    for div in divs[1:]:
        anchor_top = current_group[0]['top']
        if abs(div['top'] - anchor_top) <= y_tolerance:
            current_group.append(div)
        else:
            current_group.sort(key=lambda d: d['left'])
            line_text = '  '.join(d['text'] for d in current_group)
            lines.append(line_text)
            current_group = [div]

    # Flush last group
    if current_group:
        current_group.sort(key=lambda d: d['left'])
        line_text = '  '.join(d['text'] for d in current_group)
        lines.append(line_text)

    return lines


# Numeric token: digits with optional decimal
NUM_RE = re.compile(r'^-?\d+(?:\.\d+)?$')


def is_numeric(token: str) -> bool:
    return bool(NUM_RE.match(token))


def is_party_line(line: str) -> bool:
    """
    Party names are ALL-CAPS lines with NO standalone numbers.
    - Requires ≥88% uppercase alpha characters
    - Rejects any line that contains a standalone numeric token
    - Rejects separator lines (----, ====)
    """
    stripped = line.strip()
    if len(stripped) < 3:
        return False
    if re.match(r'^[-=*_]+$', stripped):
        return False

    # Any standalone number means it's an item row, not a party header
    tokens = stripped.split()
    for tok in tokens:
        if re.match(r'^\d+(?:\.\d+)?$', tok):
            return False

    alpha_chars = [c for c in stripped if c.isalpha()]
    if not alpha_chars:
        return False
    upper_ratio = sum(1 for c in alpha_chars if c.isupper()) / len(alpha_chars)
    if upper_ratio < 0.88:
        return False

    return True


def is_separator(line: str) -> bool:
    return bool(re.match(r'^[\s\-=*_|]+$', line))


def parse_item_line(line: str) -> dict | None:
    """
    Extract (item_name, qty, free, rate, amount) from an item row.

    Strategy:
      1. Tokenise the line.
      2. Find all numeric-or-dash positions.
      3. Take the LAST 4 as: Qty, Free, Rate, Amount.
      4. '-' in the Free slot → 0.
      5. Everything before those 4 tokens = item name.
    """
    tokens = line.split()
    if not tokens:
        return None

    numeric_indices = []
    for i, tok in enumerate(tokens):
        if is_numeric(tok) or tok == '-':
            numeric_indices.append(i)

    if len(numeric_indices) < 4:
        return None

    last4_indices = numeric_indices[-4:]
    qty_idx, free_idx, rate_idx, amt_idx = last4_indices

    try:
        qty      = float(tokens[qty_idx])
        free_tok = tokens[free_idx]
        free     = 0.0 if free_tok == '-' else float(free_tok)
        rate     = float(tokens[rate_idx])
        amount   = float(tokens[amt_idx])
    except ValueError:
        return None

    name_tokens = tokens[:qty_idx]
    item_name = ' '.join(name_tokens).strip()

    if not item_name:
        return None

    return {
        'item_name': item_name,
        'qty':       qty,
        'free':      free,
        'rate':      rate,
        'amount':    amount,
    }


def parse_html(html_content: str) -> list[dict]:
    """
    Main parser: reads the HTML, groups text by vertical position,
    and extracts party → item rows.
    """
    soup = BeautifulSoup(html_content, 'html.parser')
    divs = get_divs_sorted(soup)
    lines = group_into_lines(divs, y_tolerance=5.0)

    rows = []
    current_party = None

    for raw_line in lines:
        line = raw_line.strip()

        if not line:
            continue
        if is_separator(line):
            continue

        # TOTAL lines mark end of a party block — skip them
        if re.search(r'\bTOTAL\b', line, re.IGNORECASE):
            continue

        if is_party_line(line):
            current_party = line.strip()
            continue

        # Attempt to parse as item row
        if current_party:
            parsed = parse_item_line(line)
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


# ─────────────────────────────────────────────
#  EXCEL BUILDER
# ─────────────────────────────────────────────

def build_excel(rows: list[dict]) -> bytes:
    """Build a formatted .xlsx from parsed rows and return as bytes."""
    wb = Workbook()
    ws = wb.active
    ws.title = "Sales Report"

    # Styles
    header_font  = Font(name='Arial', bold=True, color='FFFFFF', size=11)
    header_fill  = PatternFill('solid', start_color='1B4F72')
    party_font   = Font(name='Arial', bold=True, color='1B2631', size=10)
    party_fill   = PatternFill('solid', start_color='D6EAF8')
    data_font    = Font(name='Arial', size=10)
    alt_fill     = PatternFill('solid', start_color='F8FBFD')
    border_side  = Side(style='thin', color='CCCCCC')
    cell_border  = Border(left=border_side, right=border_side,
                          top=border_side, bottom=border_side)
    center_align = Alignment(horizontal='center', vertical='center')
    left_align   = Alignment(horizontal='left',   vertical='center')
    right_align  = Alignment(horizontal='right',  vertical='center')

    # Header row
    headers    = ['Party Name', 'Item Name', 'Qty', 'Free', 'Rate', 'Amount']
    col_widths = [30, 38, 10, 10, 12, 14]

    for col_idx, (hdr, width) in enumerate(zip(headers, col_widths), start=1):
        cell = ws.cell(row=1, column=col_idx, value=hdr)
        cell.font      = header_font
        cell.fill      = header_fill
        cell.border    = cell_border
        cell.alignment = center_align
        ws.column_dimensions[get_column_letter(col_idx)].width = width

    ws.row_dimensions[1].height = 22

    # Data rows
    df = pd.DataFrame(rows)
    row_num    = 2
    prev_party = None
    alt        = False

    for _, record in df.iterrows():
        party = record['Party Name']

        # Insert a party separator row whenever the party changes
        if party != prev_party:
            ws.merge_cells(start_row=row_num, start_column=1,
                           end_row=row_num, end_column=6)
            party_cell = ws.cell(row=row_num, column=1, value=party)
            party_cell.font      = party_font
            party_cell.fill      = party_fill
            party_cell.border    = cell_border
            party_cell.alignment = left_align
            ws.row_dimensions[row_num].height = 18
            row_num    += 1
            prev_party  = party
            alt         = False

        # Item data row
        fill   = PatternFill('solid', start_color='FFFFFF') if not alt else alt_fill
        values = [record['Party Name'], record['Item Name'],
                  record['Qty'], record['Free'], record['Rate'], record['Amount']]
        aligns = [left_align, left_align,
                  center_align, center_align, right_align, right_align]

        for col_idx, (val, align) in enumerate(zip(values, aligns), start=1):
            cell = ws.cell(row=row_num, column=col_idx, value=val)
            cell.font      = data_font
            cell.fill      = fill
            cell.border    = cell_border
            cell.alignment = align
            if col_idx == 5:           # Rate
                cell.number_format = '#,##0.00'
            elif col_idx == 6:         # Amount
                cell.number_format = '#,##0.00'
            elif col_idx in (3, 4):    # Qty, Free
                cell.number_format = '#,##0'

        ws.row_dimensions[row_num].height = 16
        row_num += 1
        alt = not alt

    ws.freeze_panes = 'A2'
    ws.auto_filter.ref = f'A1:{get_column_letter(len(headers))}1'

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.read()


# ─────────────────────────────────────────────
#  ROUTES
# ─────────────────────────────────────────────

@app.route('/')
def index():
    return render_template('index.html')


@app.route('/parse', methods=['POST'])
def parse_files():
    """Accept one or more HTML files, return JSON with parsed rows + stats."""
    files = request.files.getlist('files')
    if not files:
        return jsonify({'error': 'No files uploaded'}), 400

    all_rows   = []
    file_stats = []

    for f in files:
        if not f.filename:
            continue
        try:
            html_content = f.read().decode('utf-8', errors='replace')
            rows = parse_html(html_content)
            all_rows.extend(rows)
            file_stats.append({
                'name':    f.filename,
                'rows':    len(rows),
                'parties': len(set(r['Party Name'] for r in rows)),
                'status':  'ok'
            })
        except Exception as e:
            file_stats.append({
                'name': f.filename, 'rows': 0, 'parties': 0,
                'status': f'error: {str(e)}'
            })

    if not all_rows:
        return jsonify({
            'error': 'No data could be extracted. Please check the file format.',
            'file_stats': file_stats
        }), 422

    return jsonify({
        'total_rows':    len(all_rows),
        'total_parties': len(set(r['Party Name'] for r in all_rows)),
        'file_stats':    file_stats,
        'preview':       all_rows[:200],   # for on-screen table
        'data':          all_rows          # full dataset kept client-side
    })


@app.route('/download', methods=['POST'])
def download():
    """Accept JSON rows in request body, return .xlsx file."""
    payload  = request.get_json(force=True)
    rows     = payload.get('data', [])
    filename = payload.get('filename', 'pharma_report')

    if not rows:
        return jsonify({'error': 'No data to export'}), 400

    xlsx_bytes = build_excel(rows)
    buf = io.BytesIO(xlsx_bytes)
    buf.seek(0)

    safe_name = re.sub(r'[^\w\-.]', '_', filename)
    return send_file(
        buf,
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        as_attachment=True,
        download_name=f'{safe_name}.xlsx'
    )


if __name__ == '__main__':
    app.run(debug=True, port=5050)