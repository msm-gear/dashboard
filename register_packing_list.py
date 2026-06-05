#!/usr/bin/env python3
"""
register_packing_list.py — Convert an HTML packing list to PDF and register it
in the shipment tracker orders.json for the matching order.

Usage:
    python3 register_packing_list.py <html_file> <order_id_or_market_key> [label]

Examples:
    python3 register_packing_list.py /Users/mba/dfw-packing-list.html "DFW BSBL Summer 1"
    python3 register_packing_list.py /Users/mba/dallas-packing-list.html "Dallas BSBL Summer 2" "Dallas BSBL Summer Order #2"

Market key format: "<Market> <Sport> <Season> <Order#>"
  Market: DFW, Dallas, Houston, Tampa, Austin, etc.
  Sport:  BSBL, SOC, 7V7, FSH, SBL, etc.
  Season: Summer, Fall
  Order#: 1, 2, 3, Freshman, etc.
"""

import sys
import json
import shutil
import subprocess
from pathlib import Path
from datetime import datetime

DATA_FILE          = Path(__file__).parent / 'data' / 'orders.json'
PACKING_DOCS_FOLDER = Path(__file__).parent / 'uploads' / 'packing_docs'
CHROME             = '/Applications/Google Chrome.app/Contents/MacOS/Google Chrome'


def find_order(orders, key: str):
    """Match an order by 'Market Sport Season OrderNum' key (case-insensitive)."""
    parts = key.strip().split()
    if len(parts) < 2:
        return None
    market = parts[0]
    rest   = parts[1:]
    for o in orders:
        om = o.get('market', '').lower()
        if om != market.lower():
            continue
        remaining = ' '.join(rest).lower()
        sport  = o.get('sport', '').lower()
        season = o.get('season', '').lower()
        num    = str(o.get('num_for_code', o.get('order_num', ''))).lower()
        if sport in remaining and season in remaining and num in remaining:
            return o
    return None


def html_to_pdf(html_path: Path, pdf_path: Path):
    """Convert HTML to PDF using Chrome headless."""
    result = subprocess.run([
        CHROME,
        '--headless=new',
        '--disable-gpu',
        '--no-sandbox',
        '--run-all-compositor-stages-before-draw',
        f'--print-to-pdf={pdf_path}',
        '--print-to-pdf-no-header',
        f'file://{html_path.resolve()}',
    ], capture_output=True, text=True, timeout=30)
    if not pdf_path.exists():
        raise RuntimeError(f'PDF not created. Chrome output:\n{result.stderr}')
    return pdf_path


def register(html_path: str, order_key: str, label: str = None):
    html_path = Path(html_path).resolve()
    if not html_path.exists():
        print(f'ERROR: HTML file not found: {html_path}')
        sys.exit(1)

    with open(DATA_FILE) as f:
        orders = json.load(f)

    order = find_order(orders, order_key)
    if not order:
        print(f'ERROR: No order matched key "{order_key}"')
        print('Available orders:')
        for o in orders:
            print(f'  {o.get("market")} {o.get("sport")} {o.get("season")} {o.get("order_num")} (id={o.get("id","")[:8]}…)')
        sys.exit(1)

    order_id = order['id']
    if not label:
        m   = order.get('market','')
        sp  = order.get('sport','')
        se  = order.get('season','')
        num = order.get('num_for_code', order.get('order_num',''))
        label = f'{m} {sp} {se} Order #{num}'

    # Build PDF filename from label (matches what user sees)
    pdf_name = label.replace(' ', '-') + '.pdf'
    folder   = PACKING_DOCS_FOLDER / order_id
    folder.mkdir(parents=True, exist_ok=True)
    pdf_path = folder / pdf_name

    # Also keep the HTML alongside for the View link
    html_dest = folder / html_path.name
    shutil.copy2(html_path, html_dest)

    print(f'Converting {html_path.name} → {pdf_name} ...')
    html_to_pdf(html_path, pdf_path)
    size_kb = pdf_path.stat().st_size // 1024
    print(f'PDF created: {pdf_path} ({size_kb} KB)')

    # Register PDF in order (HTML kept as "View" target)
    m   = order.get('market', '')
    sp  = order.get('sport', '')
    se  = order.get('season', '')
    num = order.get('num_for_code', order.get('order_num', ''))
    doc = {
        'filename':      pdf_name,
        'label':         label,
        'download_name': f'{m} {sp} {se} Order # {num} Packing List.pdf',
        'uploaded_at':   datetime.now().isoformat(),
    }
    order.setdefault('packing_docs', [])
    # Replace any existing entry with same filename
    order['packing_docs'] = [d for d in order['packing_docs'] if d['filename'] != pdf_name]
    order['packing_docs'].append(doc)

    with open(DATA_FILE, 'w') as f:
        json.dump(orders, f, indent=2)

    print(f'Registered "{label}" in order {order_id[:8]}…')
    print(f'Download: http://localhost:5003/packing-doc/{order_id}/{pdf_name}')
    return order_id, pdf_name


if __name__ == '__main__':
    if len(sys.argv) < 3:
        print(__doc__)
        sys.exit(1)
    register(sys.argv[1], sys.argv[2], sys.argv[3] if len(sys.argv) > 3 else None)
