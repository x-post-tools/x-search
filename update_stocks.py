import requests
import json
import io
import sys

URL = 'https://www.jpx.co.jp/markets/statistics-equities/misc/tvdivq0000001vg2-att/data_j.xls'
HEADERS = {
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36',
    'Referer': 'https://www.jpx.co.jp/markets/statistics-equities/misc/01.html',
    'Accept-Language': 'ja,en;q=0.9',
}

def parse_openpyxl(content):
    import openpyxl
    wb = openpyxl.load_workbook(io.BytesIO(content))
    ws = wb.active
    stocks = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        if not row or row[1] is None:
            continue
        code = str(row[1]).strip().split('.')[0].zfill(4)
        name = str(row[2]).strip() if row[2] else ''
        if len(code) >= 4 and name and name != 'None':
            stocks.append({'code': code, 'name': name})
    return stocks

def parse_xlrd(content):
    import xlrd
    wb = xlrd.open_workbook(file_contents=content)
    ws = wb.sheet_by_index(0)
    stocks = []
    for i in range(1, ws.nrows):
        row = ws.row_values(i)
        if not row or not row[1]:
            continue
        raw = row[1]
        code = str(int(raw)) if isinstance(raw, float) else str(raw).strip()
        code = code.zfill(4)
        name = str(row[2]).strip() if row[2] else ''
        if len(code) >= 4 and name:
            stocks.append({'code': code, 'name': name})
    return stocks

print('Downloading JPX stock list...')
try:
    resp = requests.get(URL, headers=HEADERS, timeout=60)
    resp.raise_for_status()
    content = resp.content
    print(f'Downloaded {len(content)} bytes')
except Exception as e:
    print(f'Download failed: {e}')
    sys.exit(1)

stocks = None
for parser, label in [(parse_openpyxl, 'openpyxl'), (parse_xlrd, 'xlrd')]:
    try:
        stocks = parser(content)
        print(f'Parsed with {label}: {len(stocks)} stocks')
        break
    except Exception as e:
        print(f'{label} failed: {e}')

if not stocks:
    print('All parsers failed')
    sys.exit(1)

with open('stocks.json', 'w', encoding='utf-8') as f:
    json.dump(stocks, f, ensure_ascii=False, separators=(',', ':'))
print(f'Saved {len(stocks)} stocks to stocks.json')
