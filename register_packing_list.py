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

import os
import sys
import json
import shutil
import subprocess
import tempfile
from pathlib import Path
from datetime import datetime

DATA_FILE          = Path(__file__).parent / 'data' / 'orders.json'
PACKING_DOCS_FOLDER = Path(__file__).parent / 'uploads' / 'packing_docs'
CHROME             = '/Applications/Google Chrome.app/Contents/MacOS/Google Chrome'
# The tracker runs on the MSM server (tracker.msm-hq.com). On the laptop this script makes the PDF
# with Chrome, then hands both files to the server copy of this script to register them there.
SERVER     = 'root@161.35.49.217'
SERVER_KEY = os.path.expanduser('~/.ssh/msm_hetzner')
SERVER_DIR = '/srv/msm/shipment-tracker'
ON_SERVER  = Path(__file__).resolve().parent == Path(SERVER_DIR)


def find_order(orders, key: str):
    """Match an order by 'Market Sport Season OrderNum' key (case-insensitive).
    Tries multi-word market prefixes so 'South Florida BSBL Summer 1' works."""
    parts = key.strip().split()
    if len(parts) < 2:
        return None
    for split_at in range(1, len(parts)):
        market   = ' '.join(parts[:split_at])
        rest_str = ' '.join(parts[split_at:]).lower()
        for o in orders:
            if o.get('market', '').lower() != market.lower():
                continue
            sport  = o.get('sport', '').lower()
            season = o.get('season', '').lower()
            num    = str(o.get('num_for_code', o.get('order_num', ''))).lower()
            if sport in rest_str and season in rest_str and num in rest_str:
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


def register_remote(html_path: str, order_key: str, label: str = None):
    """Laptop: PDF locally (Chrome), then register on the server."""
    html_path = Path(html_path).resolve()
    if not html_path.exists():
        print(f'ERROR: HTML file not found: {html_path}')
        sys.exit(1)
    ssh = ['ssh', '-i', SERVER_KEY, SERVER]
    with tempfile.TemporaryDirectory() as tmp:
        pdf_path = Path(tmp) / 'packing.pdf'
        print(f'Converting {html_path.name} → PDF ...')
        html_to_pdf(html_path, pdf_path)
        remote_tmp = subprocess.run(ssh + ['mktemp -d'], capture_output=True, text=True, check=True).stdout.strip()
        subprocess.run(['scp', '-q', '-i', SERVER_KEY, str(html_path), str(pdf_path), f'{SERVER}:{remote_tmp}/'], check=True)
        cmd = ['sudo', '-u', 'msm', f'{SERVER_DIR}/.venv/bin/python', f'{SERVER_DIR}/register_packing_list.py', '--from-files',
               f'{remote_tmp}/{html_path.name}', f'{remote_tmp}/packing.pdf', order_key] + ([label] if label else [])
        quoted = ' '.join("'" + c.replace("'", "'\\''") + "'" for c in cmd)
        r = subprocess.run(ssh + [f'chmod -R a+r {remote_tmp} && {quoted}; rc=$?; rm -rf {remote_tmp}; exit $rc'])
        sys.exit(r.returncode)


def register(html_path: str, order_key: str, label: str = None, pdf_ready: str = None):
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

    # Build PDF filename — strip chars that break secure_filename (e.g. #)
    import re as _re
    pdf_name = _re.sub(r'[^\w.\-]', '', label.replace(' ', '-')) + '.pdf'
    folder   = PACKING_DOCS_FOLDER / order_id
    folder.mkdir(parents=True, exist_ok=True)
    pdf_path = folder / pdf_name

    # Also keep the HTML alongside for the View link
    html_dest = folder / html_path.name
    shutil.copy2(html_path, html_dest)

    if pdf_ready:                       # PDF already made on the laptop
        shutil.copy2(pdf_ready, pdf_path)
    else:
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
    print(f'Download: https://tracker.msm-hq.com/packing-doc/{order_id}/{pdf_name}')
    return order_id, pdf_name


if __name__ == '__main__':
    if len(sys.argv) >= 5 and sys.argv[1] == '--from-files':     # server side of a laptop run
        register(sys.argv[2], sys.argv[4], sys.argv[5] if len(sys.argv) > 5 else None, pdf_ready=sys.argv[3])
        sys.exit(0)
    if len(sys.argv) < 3:
        print(__doc__)
        sys.exit(1)
    if ON_SERVER or os.environ.get('TRACKER_LOCAL') == '1':
        register(sys.argv[1], sys.argv[2], sys.argv[3] if len(sys.argv) > 3 else None)
    else:
        register_remote(sys.argv[1], sys.argv[2], sys.argv[3] if len(sys.argv) > 3 else None)
