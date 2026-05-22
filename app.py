from __future__ import annotations
from flask import Flask, render_template, request, session, jsonify, redirect, url_for
import openpyxl
import re
import difflib
import tempfile
import os
import json
import uuid
import csv
import requests
import subprocess
import threading
import smtplib
import time
from email.mime.text import MIMEText
from zoneinfo import ZoneInfo
from datetime import datetime, timedelta
from io import StringIO
from dotenv import load_dotenv
from werkzeug.utils import secure_filename
from pathlib import Path
from collections import defaultdict

load_dotenv()

app = Flask(__name__)
app.secret_key = 'msm-shipment-tracker-2026'

UPLOAD_FOLDER = Path(__file__).parent / 'uploads'
DATA_FOLDER = Path(__file__).parent / 'data'
UPLOAD_FOLDER.mkdir(exist_ok=True)
DATA_FOLDER.mkdir(exist_ok=True)

CARRIERS = [
    ('fedex', 'FedEx'),
    ('ups', 'UPS'),
    ('usps', 'USPS'),
    ('dhl', 'DHL'),
    ('fedex-freight', 'FedEx Freight'),
]

SPORTS = [
    ('BSBL', 'Baseball'),
    ('SBL',  'Softball'),
    ('SOC',  'Soccer'),
    ('GLF',  'Golf'),
    ('7V7',  '7V7 Football'),
    ('FSH',  'Freshmen'),
    ('BWL',  'Bowling'),
]

SEASONS = ['Fall', 'Summer', 'Spring']

MARKETS = [
    'Austin',
    'Dallas',
    'DFW',
    'Houston',
    'Kansas City',
    'Las Vegas',
    'Ohio',
    'Orlando',
    'San Antonio',
    'South Florida',
    'South Jersey',
    'Tampa',
    'West Texas',
    'HQ',
]

SHIP_TO = {
    'Austin':        {'name': 'Corey Bailey',    'market': 'Austin',        'phone': '512-653-5591', 'email': 'austin@middleschoolmatchup.com',         'address': '508 Otto Avenue',          'city': 'Georgetown',   'state': 'TX', 'zip': '78626'},
    'Dallas':        {'name': 'Gavin Himel',     'market': 'Dallas',        'phone': '469-912-4187', 'email': 'gavin@middleschoolmatchup.com',           'address': '6249 Edgestone Dr',        'city': 'Frisco',       'state': 'TX', 'zip': '75034'},
    'DFW':           {'name': 'Colin Burns',     'market': 'DFW',           'phone': '682-999-9240', 'email': 'dfwwest@middleschoolmatchup.com',         'address': '837 Landmark Drive',       'city': 'Aledo',        'state': 'TX', 'zip': '76008'},
    'Houston':       {'name': 'Wes Hollas',      'market': 'Houston',       'phone': '713-557-1945', 'email': 'wes@middleschoolmatchup.com',             'address': '6006 Boulder Lake Ct',     'city': 'Kingwood',     'state': 'TX', 'zip': '77345'},
    'Houston SBL':   {'name': 'Wes Hollas',      'market': 'Houston',       'phone': '713-557-1945', 'email': 'wes@middleschoolmatchup.com',             'address': '6006 Boulder Lake Ct',     'city': 'Kingwood',     'state': 'TX', 'zip': '77345'},
    'Houston Soccer':{'name': 'Wes Hollas',      'market': 'Houston',       'phone': '713-557-1945', 'email': 'wes@middleschoolmatchup.com',             'address': '6006 Boulder Lake Ct',     'city': 'Kingwood',     'state': 'TX', 'zip': '77345'},
    'Kansas City':   {'name': 'Jaxon Himel',     'market': 'Kansas City',   'phone': '469-912-4198', 'email': 'jaxonhimel@middleschoolmatchup.com',     'address': '2121 Central Street Apt 401','city': 'Kansas City', 'state': 'MO', 'zip': '64108'},
    'Las Vegas':     {'name': 'Sacha Donovan',   'market': 'Las Vegas',     'phone': '310-483-9399', 'email': 'lasvegas@middleschoolmatchup.com',        'address': '509 Nursery Rd',           'city': 'Spring',       'state': 'TX', 'zip': '77380'},
    'Ohio':          {'name': 'Austin Smith',    'market': 'Ohio',          'phone': '937-838-4340', 'email': 'austin8339@gmail.com',                   'address': '3858 Cuba Rd',             'city': 'Wilmington',   'state': 'OH', 'zip': '45177'},
    'Orlando':       {'name': 'Leisa Morgan',    'market': 'Orlando',       'phone': '214-842-7976', 'email': 'orlando@middleschoolmatchup.com',         'address': '2143 Sunshine Peak Dr',    'city': 'Minneola',     'state': 'FL', 'zip': '34756'},
    'San Antonio':   {'name': 'Paige Kovar',     'market': 'San Antonio',   'phone': '281-468-4838', 'email': 'paige@middleschoolmatchup.com',           'address': '13003 Ferry Cove Lane',    'city': 'Pearland',     'state': 'TX', 'zip': '77584'},
    'South Florida': {'name': 'Tre Ratliff',     'market': 'South Florida', 'phone': '954-245-8873', 'email': 'southflorida@middleschoolmatchup.com',    'address': '450 Martin Rd',            'city': 'Margate',      'state': 'FL', 'zip': '33063'},
    'South Jersey':  {'name': 'Leon Wright',     'market': 'South Jersey',  'phone': '310-483-9399', 'email': 'southjersey@middleschoolmatchup.com',     'address': '146 E Wolfert Station Rd', 'city': 'Mickleton',    'state': 'NJ', 'zip': '08056'},
    'Tampa':         {'name': 'Amanda DeJiulio', 'market': 'Tampa',         'phone': '352-804-4905', 'email': 'tampa@middleschoolmatchup.com',           'address': '5392 Angelonia Terrace',   'city': "Land O' Lakes",'state': 'FL', 'zip': '34639'},
    'West Texas':    {'name': 'Melissa Boland',  'market': 'West Texas',    'phone': '254-707-3258', 'email': 'westtexas@middleschoolmatchup.com',       'address': '6507 Saddle Horn Ln',      'city': 'Midland',      'state': 'TX', 'zip': '79707'},
}


def _ship_to_info(market: str) -> dict:
    """
    Return ship-to contact for a market.
    Resolution order:
      1. Exact case-insensitive match against SHIP_TO keys
      2. Strip trailing sport words (SBL, Soccer, Baseball…) and retry
      3. Fuzzy match (cutoff 0.7) against SHIP_TO keys
    Raises ValueError if no match found, so callers can surface the error.
    """
    _SPORT_SUFFIXES = re.compile(
        r'\s+(sbl|soccer|baseball|bsbl|softball|football|7v7|golf|bowling|freshmen)$',
        re.IGNORECASE,
    )
    m = market.strip()

    # 1. Exact (case-insensitive)
    for key in SHIP_TO:
        if key.lower() == m.lower():
            return SHIP_TO[key]

    # 2. Strip sport suffix and retry
    stripped = _SPORT_SUFFIXES.sub('', m).strip()
    if stripped.lower() != m.lower():
        for key in SHIP_TO:
            if key.lower() == stripped.lower():
                return SHIP_TO[key]

    # 3. Fuzzy match against known keys
    candidates = [k.lower() for k in SHIP_TO]
    close = difflib.get_close_matches(m.lower(), candidates, n=1, cutoff=0.7)
    if not close:
        close = difflib.get_close_matches(stripped.lower(), candidates, n=1, cutoff=0.7)
    if close:
        matched_key = next(k for k in SHIP_TO if k.lower() == close[0])
        return SHIP_TO[matched_key]

    raise ValueError(f"Unknown market '{market}' — no ship-to contact found.")


def _parse_contents_qty(contents: str) -> int:
    """Sum quantities from contents like 'YL-81 YXL-112'."""
    return sum(int(n) for n in re.findall(r'[A-Za-z]+-(\d+)', contents or ''))


MARKET_CODES = {
    'austin':         'AUS',
    'dallas':         'DAL',
    'dfw':            'DFW',
    'houston':        'HOU',
    'houston sbl':    'HOU',
    'houston soccer': 'HOU',
    'kansas city':    'KCY',
    'las vegas':      'LVG',
    'ohio':           'OHI',
    'orlando':        'ORL',
    'san antonio':    'SAT',
    'south florida':  'SFL',
    'south jersey':   'SJY',
    'tampa':          'TPA',
    'west texas':     'WTX',
    'midland':        'WTX',
    'hq':             'HQ',
}

GEAR_CODES = {
    'Jersey': 'JRS',
    'Hoodie': 'HOD',
    'Tee': 'TEE',
    'Cap': 'CAP',
    'Cage Jacket': 'CAG',
    'Windbreaker': 'WNB',
    'Ring': 'RNG',
    'Hypechain': 'HYC',
    'Awards': 'AWD',       # Rings & Hype Chains shipped together
    'Park Banner': 'BAN',
    'Name Banner': 'NBN',
    'Hybrid Pullover': 'HYP',
}

ALL_GEAR_TYPES = list(GEAR_CODES.keys())  # preserves insertion order

# ClickUp custom field IDs for "Order Tracking [NEW]" list
CF_TRACKING        = '3641c929-080f-4b11-9438-7b54aa7ab859'  # 🚚 Tracking # (short_text)
CF_PHASE           = 'e53babec-e5bb-4004-b5a1-b3c559d230d5'  # Production Phase (orderindex: 0=Submitted,3=Shipped)
CF_ORDER_STATUS    = '79de4435-914d-464a-b8b5-423a255eda59'  # Order Status (orderindex: 1=On Schedule)
CF_ORDER_TEXT      = '911a3af5-130a-4d08-a6f1-e71bcfd17803'  # 🛒 Order # (short_text)
CF_ORDER_DATE      = '056d3ca0-dcc7-4f3e-beb2-0a00a93580bd'  # Order Date (date ms)
CF_SHIP_DATE_ACT   = '8c787610-911e-47aa-b5e8-24038d06b835'  # Ship Date [Actual] (date ms)
CF_DELIVERY_DATE   = '17868428-c3c0-42cb-ae8b-c925be3639a7'  # Delivery Date (date ms)
CF_PROOFS_READY    = 'e13b7a70-6632-4c21-8228-799ccda445fe'  # Proofs Ready (date)
CF_PROOFS_APPROVED = 'aeedcc61-3ffc-402a-9f35-b514ac7576ff'  # Proofs Approved (date)
CF_PROD_COMPLETE   = 'e8d1c03f-c3dd-4817-a031-f4d15e082b5b'  # Production Complete (date)
CF_SHIP_DATE       = '5e5082cf-cea1-4c7f-bb93-30f8bdaaebbc'  # Ship Date / estimated (date)
CF_ETA_ZEES        = '7bc80ee9-0961-4974-af7a-f1b0c923c053'  # ETA Provided to Zees (date)
CF_RECEIVING           = '3ca705c5-6fa3-4490-a804-c3bb1844c0d6'  # Receiving Condition (short_text)
CF_DELIVERY_DATE_ACT   = '05c54d04-e367-4cf2-9172-bdba85c1f5c4'  # Delivery Date [Actual] (date ms)
CF_PROOFS_APPROVED_ACT = '78de254b-383c-4150-bbf7-8fe1b13c5d47'  # Proofs Approved [Actual] (date ms)


def count_working_days(start: datetime, end: datetime) -> int:
    """Count Mon–Fri working days from start (exclusive) to end (inclusive)."""
    if not start or not end or end <= start:
        return 0
    count, d = 0, start + timedelta(days=1)
    while d <= end:
        if d.weekday() < 5:
            count += 1
        d += timedelta(days=1)
    return count


def get_market_code(market: str) -> str:
    return MARKET_CODES.get(market.lower().strip(), market[:3].upper())


def _market_dedup_key(market: str) -> str:
    """Return a normalized market code for deduplication, with fuzzy fallback."""
    key = market.lower().strip()
    if key in MARKET_CODES:
        return MARKET_CODES[key]
    # fuzzy-match against known market names to handle typos (e.g. 'allas' → 'dallas')
    close = difflib.get_close_matches(key, MARKET_CODES.keys(), n=1, cutoff=0.6)
    if close:
        return MARKET_CODES[close[0]]
    return key


def to_ms(date_str: str):
    if not date_str:
        return None
    try:
        dt = datetime.strptime(date_str, '%Y-%m-%d')
        return int(dt.timestamp() * 1000)
    except ValueError:
        return None


def date_to_mmddyy(date_str: str) -> str:
    if not date_str:
        return datetime.now().strftime('%m%d%y')
    try:
        return datetime.strptime(date_str, '%Y-%m-%d').strftime('%m%d%y')
    except ValueError:
        return datetime.now().strftime('%m%d%y')


# Gear priority for naming mixed shipments.
# Accessories (rings/chains/banners) take top priority so they can be tracked.
# Jersey/Hoodie/etc. rank below them.
GEAR_PRIORITY = [
    'Ring', 'Hypechain', 'Awards', 'Park Banner', 'Name Banner',
    'Cage Jacket', 'Cap', 'Windbreaker', 'Hybrid Pullover',
    'Jersey', 'Hoodie', 'Tee',
]


def _primary_gear(gear_list: list) -> str:
    """Return the highest-priority gear type from a list."""
    for g in GEAR_PRIORITY:
        if g in gear_list:
            return g
    return gear_list[0] if gear_list else 'Jersey'


def build_shipment_map(rows: list) -> tuple:
    """
    Returns:
      tracking_to_letter  – dict: tracking_num → 'A', 'B', 'C' ...
      tracking_to_primary – dict: tracking_num → primary gear type
      tracking_to_rows    – dict: tracking_num → [rows]

    Assignment goal: every gear type gets at least one shipment letter.
    - Process exclusive (single-gear) shipments first, then shared ones.
    - For each shipment, assign to a gear that has zero trackings assigned yet
      (highest-priority among those). If all gears already have ≥1, assign to
      the one with the fewest (tiebreak by priority).
    """
    order = []
    t2rows = {}
    for r in rows:
        tn = r.get('tracking_number', '')
        if not tn:
            continue
        if tn not in t2rows:
            t2rows[tn] = []
            order.append(tn)
        t2rows[tn].append(r)

    # Process exclusive shipments first so each gear can claim its own slot
    sorted_tns = sorted(order, key=lambda tn: len({r['gear_type'] for r in t2rows[tn]}))

    gear_count = defaultdict(int)
    t2primary  = {}

    for tn in sorted_tns:
        gears     = list({r['gear_type'] for r in t2rows[tn]})
        untracked = [g for g in gears if gear_count[g] == 0]

        if untracked:
            primary = _primary_gear(untracked)
        else:
            min_count  = min(gear_count[g] for g in gears)
            candidates = [g for g in gears if gear_count[g] == min_count]
            primary    = _primary_gear(candidates)

        t2primary[tn] = primary
        gear_count[primary] += 1

    # Letters are per-gear-type; only add A/B/C when a gear has 2+ shipments
    gear_to_tns = defaultdict(list)
    for tn in order:                         # preserve appearance order for letters
        gear_to_tns[t2primary[tn]].append(tn)

    t2letter = {}                            # tracking_num → letter ('' if only one)
    for gear, tns in gear_to_tns.items():
        if len(tns) == 1:
            t2letter[tns[0]] = ''            # single shipment — no letter needed
        else:
            for i, tn in enumerate(tns):
                t2letter[tn] = chr(65 + i)  # A, B, C …

    return t2letter, t2primary, t2rows


def add_working_days(start, n: int):
    """Add n working days (Mon–Fri) to a date or datetime."""
    d = start if not isinstance(start, datetime) else start.date()
    while n > 0:
        d += timedelta(days=1)
        if d.weekday() < 5:
            n -= 1
    return d


def _date_ms(d) -> int:
    """Midnight-local ms timestamp for a date object."""
    return int(datetime(d.year, d.month, d.day).timestamp() * 1000)


# Gear types that go through the proof flow
_PROOF_GEARS = {
    'Jersey':      7,
    'Cage Jacket': 7,
    'Hoodie':      10,
}
# Gear types with direct working-day countdown from order date (no proofs)
_DIRECT_GEARS = {
    'Tee':         8,
    'Cap':         14,
    'Hypechain':   7,
    'Name Banner': 5,
    'Ring':        14,
}


def calc_gear_dates(gear_type: str, order_date_str: str) -> list:
    """Return custom-field dicts for Proofs Ready / Proofs Approved / Production Complete."""
    if not order_date_str:
        return []
    try:
        od = datetime.strptime(order_date_str, '%Y-%m-%d').date()
    except ValueError:
        return []

    prod_complete = None
    fields = []
    if gear_type in _PROOF_GEARS:
        proofs_ready    = add_working_days(od, 2)
        proofs_approved = add_working_days(proofs_ready, 1)
        prod_complete   = add_working_days(proofs_approved, _PROOF_GEARS[gear_type])
        fields = [
            {'id': CF_PROOFS_READY,    'value': _date_ms(proofs_ready)},
            {'id': CF_PROOFS_APPROVED, 'value': _date_ms(proofs_approved)},
            {'id': CF_PROD_COMPLETE,   'value': _date_ms(prod_complete)},
        ]
    elif gear_type in _DIRECT_GEARS:
        prod_complete = add_working_days(od, _DIRECT_GEARS[gear_type])
        fields = [{'id': CF_PROD_COMPLETE, 'value': _date_ms(prod_complete)}]

    if prod_complete:
        ship_date     = prod_complete + timedelta(days=1)
        delivery_date = add_working_days(ship_date, 5)
        fields += [
            {'id': CF_SHIP_DATE,     'value': _date_ms(ship_date)},
            {'id': CF_DELIVERY_DATE, 'value': _date_ms(delivery_date)},
        ]

    return fields


def build_order_code(market: str, sport: str, gear: str, order_date: str, order_num: str) -> str:
    mkt = get_market_code(market)
    gear_code = GEAR_CODES.get(gear, gear[:3].upper())
    date_code = date_to_mmddyy(order_date)
    return f"{mkt}-{sport}-{gear_code}-{date_code}-{order_num}"


def build_task_name(market: str, season: str, sport: str, gear: str, order_num: str) -> str:
    return f"{market} {season} {sport} {gear} {order_num}"


def normalize_gear(raw: str) -> str:
    raw = raw.strip().lower()
    mapping = {
        'jersey': 'Jersey', 'jerseys': 'Jersey',
        'tee': 'Tee', 'tees': 'Tee',
        'hoodie': 'Hoodie', 'hoodies': 'Hoodie',
        'cap': 'Cap', 'caps': 'Cap',
        'windbreaker': 'Windbreaker',
        'jacket': 'Jacket', 'jackets': 'Jacket',
        'cage jacket': 'Cage Jacket', 'cage': 'Cage Jacket',
        'ring': 'Ring', 'rings': 'Ring',
        'hypechain': 'Hypechain', 'hype chain': 'Hypechain', 'hypechainn': 'Hypechain',
        'park banner': 'Park Banner', 'banner': 'Park Banner',
        'name banner': 'Name Banner',
        'hybrid': 'Hybrid Pullover', 'hybrid pullover': 'Hybrid Pullover',
    }
    return mapping.get(raw, raw.capitalize())


def _parse_item_cols(row, col_offset: int):
    """Parse one item entry from a row starting at col_offset (0 or 6)."""
    item = row[col_offset] if len(row) > col_offset else None
    if not item or '#' not in str(item):
        return None

    item_str = str(item).strip()
    hash_idx = item_str.index('#')
    market = item_str[:hash_idx].strip()
    rest   = item_str[hash_idx + 1:].strip()

    m = re.match(r'(\d+)\s+([A-Za-z\s]+?)\s*-\s*(\d+)', rest)
    if not m:
        return None

    def _col(offset):
        idx = col_offset + offset
        return row[idx] if len(row) > idx else None

    tracking_raw = _col(4)
    tracking = ''
    if tracking_raw:
        t = str(tracking_raw).strip()
        if t.endswith('.0'):
            t = t[:-2]
        if t and t.lower() not in ('none', 'nan', ''):
            tracking = t

    return {
        'id':              str(uuid.uuid4()),
        'market':          market,
        'order_num':       m.group(1),
        'gear_type':       normalize_gear(m.group(2).strip()),
        'box_num':         int(m.group(3)),
        'contents':        str(_col(1)).strip() if _col(1) else '',
        'weight':          str(_col(2)).strip() if _col(2) else '',
        'declared_value':  str(_col(3)).strip() if _col(3) else '',
        'tracking_number': tracking,
        'carrier':         'fedex',
    }


def parse_packing_list(filepath: str) -> list:
    wb = openpyxl.load_workbook(filepath)
    ws = wb.active

    # Fill merged cells with their top-left value so every row reads correctly
    for merge_range in list(ws.merged_cells.ranges):
        top_val = ws.cell(merge_range.min_row, merge_range.min_col).value
        ws.unmerge_cells(str(merge_range))
        for row in ws.iter_rows(min_row=merge_range.min_row, max_row=merge_range.max_row,
                                 min_col=merge_range.min_col, max_col=merge_range.max_col):
            for cell in row:
                cell.value = top_val

    raw = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        for col_offset in (0, 6):          # primary cols A-E, secondary cols G-K
            entry = _parse_item_cols(row, col_offset)
            if entry:
                raw.append(entry)

    # Aggregate rows with the same (market, order_num, gear_type, box_num):
    # contents column has one team name per row — join them all.
    seen = {}
    for r in raw:
        key = (_market_dedup_key(r['market']), r['order_num'], r['gear_type'], r['box_num'])
        if key not in seen:
            seen[key] = r
        else:
            existing = seen[key]
            # Prefer row that has a tracking number
            if r['tracking_number'] and not existing['tracking_number']:
                r['contents'] = existing['contents']  # carry over contents before replacing
                seen[key] = r
                existing = r
            # Merge contents: collect unique non-empty team names
            parts = [p.strip() for p in existing['contents'].split(',') if p.strip()
                     and p.strip().lower() != 'see attached']
            new_part = r['contents'].strip()
            if new_part and new_part.lower() != 'see attached' and new_part not in parts:
                parts.append(new_part)
            existing['contents'] = ', '.join(parts)

    result = list(seen.values())
    result.sort(key=lambda r: (r['gear_type'], r['box_num']))
    return result


def _safe_float(v) -> float:
    try:
        return float(str(v).replace(',', '').strip())
    except (ValueError, AttributeError):
        return 0.0


_CST = ZoneInfo('America/Chicago')


def _now_cst() -> datetime:
    return datetime.now(_CST)


def _fmt_cst(dt: datetime = None) -> str:
    """Format a datetime in Chicago time as 'May 14, 2026, 3:45 PM CDT'."""
    d = (dt or _now_cst()).astimezone(_CST)
    hour = d.strftime('%I').lstrip('0') or '12'
    return d.strftime(f'%b {d.day}, %Y, {hour}:%M %p %Z')


def get_settings():
    return {
        'clickup_token':      os.getenv('CLICKUP_TOKEN', ''),
        'clickup_list_id':    os.getenv('CLICKUP_LIST_ID', ''),
        'slack_webhook_url':  os.getenv('SLACK_WEBHOOK_URL', ''),
        'email_app_password': os.getenv('EMAIL_APP_PASSWORD', ''),
        'track17_api_key':    os.getenv('TRACK17_API_KEY', ''),
        'fedex_api_key':      os.getenv('FEDEX_API_KEY', ''),
        'fedex_secret_key':   os.getenv('FEDEX_SECRET_KEY', ''),
    }


# ── Auto-Tracking via FedEx API (direct) + 17track fallback ──────────────────

_last_auto_check = {'time': None, 'error': None}

_17TRACK_REGISTERED = set()   # tracking numbers already registered this session

_fedex_token_cache: dict = {'token': None, 'expires_at': 0.0}


def _get_fedex_token() -> str | None:
    api_key = os.getenv('FEDEX_API_KEY', '')
    secret  = os.getenv('FEDEX_SECRET_KEY', '')
    if not api_key or not secret:
        return None
    now = time.time()
    if _fedex_token_cache['token'] and now < _fedex_token_cache['expires_at']:
        return _fedex_token_cache['token']
    try:
        r = requests.post(
            'https://apis.fedex.com/oauth/token',
            data={'grant_type': 'client_credentials', 'client_id': api_key, 'client_secret': secret},
            headers={'Content-Type': 'application/x-www-form-urlencoded'},
            timeout=10,
        )
        r.raise_for_status()
        d = r.json()
        _fedex_token_cache['token']      = d['access_token']
        _fedex_token_cache['expires_at'] = now + d.get('expires_in', 3600) - 60
        return _fedex_token_cache['token']
    except Exception:
        return None


def _fedex_code_to_stage(code: str, desc: str, country_code: str) -> str:
    code = (code or '').upper()
    desc = (desc or '').upper()
    if code == 'DL':
        return 'Delivered'
    if code == 'OD':
        return 'OutForDelivery'
    if code == 'HL':
        return 'OutForDelivery'   # Hold at location / Ready for pickup
    if code in ('DE', 'DY', 'EX', 'IX', 'SE', 'CA', 'RS', 'PX'):
        return 'Exception'
    if code in ('PU', 'LO'):
        return 'PickedUp'
    if code in ('OC', 'OX', 'PM', 'NP', 'AP'):
        return 'LabelCreated'
    if code == 'CC':
        return 'ClearedCustoms'
    if code == 'CD':
        return 'AtCustoms'
    if code in ('IT', 'AR', 'DP', 'AF', 'FD', 'SF', 'TR', 'PF', 'PL', 'EZ',
                'AA', 'AD', 'EP', 'ME', 'RM', 'SP', 'AX', 'AC'):
        if 'OUT FOR DELIVERY' in desc:
            return 'OutForDelivery'
        if 'CUSTOMS' in desc or 'CLEARANCE' in desc:
            return 'ClearedCustoms' if ('CLEARED' in desc or 'RELEASED' in desc) else 'AtCustoms'
        if country_code == 'US':
            return 'InTransitDom'
        if country_code and country_code != 'US':
            return 'InTransitIntl'
        return 'InTransitDom'
    return 'Pending'


def _fetch_fedex(tracking_number: str) -> dict | None:
    token = _get_fedex_token()
    if not token:
        return None
    try:
        r = requests.post(
            'https://apis.fedex.com/track/v1/trackingnumbers',
            json={
                'includeDetailedScans': True,
                'trackingInfo': [{'trackingNumberInfo': {'trackingNumber': tracking_number}}],
            },
            headers={'Authorization': f'Bearer {token}', 'Content-Type': 'application/json', 'X-locale': 'en_US'},
            timeout=15,
        )
        r.raise_for_status()
        data = r.json()
        track_results = (data.get('output', {})
                             .get('completeTrackResults', [{}])[0]
                             .get('trackResults', [{}])[0])
        latest     = track_results.get('latestStatusDetail', {}) or {}
        events     = track_results.get('events', []) or []
        date_times = track_results.get('dateAndTimes', []) or []

        code = latest.get('code', '') or latest.get('derivedCode', '')

        # Use the latest scan event for description + location when available
        latest_event = events[0] if events else {}
        desc     = latest_event.get('eventDescription') or latest.get('statusByLocale', '') or latest.get('description', '')
        loc_data = latest_event.get('scanLocation') or latest.get('scanLocation', {}) or {}
        city     = loc_data.get('city', '')
        state    = loc_data.get('stateOrProvinceCode', '')
        country  = loc_data.get('countryCode', '')
        location = ', '.join(filter(None, [city, state]))

        # event_time: prefer scan event date, fall back to dateAndTimes entries
        event_time = latest_event.get('date', '')
        if not event_time:
            dt_map = {d['type']: d['dateTime'] for d in date_times}
            event_time = (dt_map.get('ACTUAL_DELIVERY')
                          or dt_map.get('SHIP')
                          or dt_map.get('ACTUAL_PICKUP', ''))

        if not code:
            return None
        stage = _fedex_code_to_stage(code, desc, country)
        return {'tag': stage, 'last_event': desc, 'location': location, 'event_time': event_time}
    except Exception:
        return None


def _17track_tag_to_stage(tag17: str, event_desc: str, location: str, country: str = '') -> str:
    """Map a 17track status tag + event text to our internal stage."""
    d = event_desc.upper()
    l = location.upper()

    if tag17 == 'Delivered':
        return 'Delivered'
    if tag17 in ('Undelivered', 'Alert', 'Expired'):
        return 'Exception'
    if tag17 == 'Pickup':
        return 'PickedUp'
    if tag17 == 'InfoReceived':
        return 'LabelCreated'
    if tag17 == 'NotFound':
        return 'Pending'

    # InTransit — drill into event text for customs / delivery stages
    if tag17 == 'InTransit':
        if 'OUT FOR DELIVERY' in d:
            return 'OutForDelivery'
        if any(w in d for w in ('CLEARED', 'RELEASED FROM CUSTOMS', 'CUSTOMS RELEASED')):
            return 'ClearedCustoms'
        if any(w in d for w in ('CUSTOMS', 'CLEARANCE', 'HELD BY')):
            return 'AtCustoms'
        # Use country code if available, else fall back to location text
        if country == 'US':
            return 'InTransitDom'
        if country and country != 'US':
            return 'InTransitIntl'
        us_states = ('TX', 'CA', 'IL', 'NY', 'FL', 'GA', 'OH', 'WA', 'OR',
                     'TN', 'NJ', 'PA', 'AZ', 'NV', 'CO', 'IN', 'MI', 'NC', 'VA')
        if any(f', {s}' in l or l.endswith(f' {s}') for s in us_states):
            return 'InTransitDom'
        if 'UNITED STATES' in l or ', US' in l:
            return 'InTransitDom'
        return 'InTransitIntl'

    return 'Pending'


def _fetch_17track(tracking_number: str) -> dict:
    """Returns {'tag', 'last_event', 'location'} via 17track API."""
    api_key = os.getenv('TRACK17_API_KEY', '')
    if not api_key:
        return None   # signal: not configured

    headers = {'17token': api_key, 'Content-Type': 'application/json'}

    # Register with FedEx carrier 100003; if already registered on wrong carrier, delete + re-register
    if tracking_number not in _17TRACK_REGISTERED:
        try:
            r = requests.post(
                'https://api.17track.net/track/v2.2/register',
                json=[{'number': tracking_number, 'carrier': 100003}],
                headers=headers,
                timeout=10,
            )
            rejected = r.json().get('data', {}).get('rejected', [])
            if rejected:
                # Already registered (possibly wrong carrier) — delete and retry
                requests.post(
                    'https://api.17track.net/track/v2.2/deletetrack',
                    json=[{'number': tracking_number}],
                    headers=headers,
                    timeout=10,
                )
                time.sleep(1)
                requests.post(
                    'https://api.17track.net/track/v2.2/register',
                    json=[{'number': tracking_number, 'carrier': 100003}],
                    headers=headers,
                    timeout=10,
                )
            _17TRACK_REGISTERED.add(tracking_number)
        except Exception:
            pass

    resp = requests.post(
        'https://api.17track.net/track/v2.2/gettrackinfo',
        json=[{'number': tracking_number}],
        headers=headers,
        timeout=15,
    )
    resp.raise_for_status()
    data = resp.json()

    accepted = data.get('data', {}).get('accepted', [])
    if not accepted:
        return None

    info       = accepted[0]
    track_info = info.get('track_info', {}) or {}
    latest_st  = track_info.get('latest_status', {}) or {}
    latest_ev  = track_info.get('latest_event') or {}

    tag17 = latest_st.get('status') or ''
    if not tag17 or tag17 == 'NotFound':
        return None   # no data yet — preserve existing status

    event      = latest_ev.get('description', '') or ''
    loc        = latest_ev.get('location', '') or ''
    country    = (latest_ev.get('address', {}) or {}).get('country', '') or ''
    event_time = latest_ev.get('time', '') or ''    # ISO datetime from carrier

    stage = _17track_tag_to_stage(tag17, event, loc, country)
    return {'tag': stage, 'last_event': event, 'location': loc, 'event_time': event_time}


def check_all_fedex_statuses(notify_changes: bool = False) -> dict:
    """
    Poll FedEx API (or 17track fallback) for all tracked shipments.
    Update local statuses when changed; send popup + email for forward moves.
    """
    global _last_auto_check

    use_fedex  = bool(os.getenv('FEDEX_API_KEY') and os.getenv('FEDEX_SECRET_KEY'))
    use_17track = bool(os.getenv('TRACK17_API_KEY'))

    if not use_fedex and not use_17track:
        return {'ok': False, 'error': 'No tracking API configured. Add FedEx credentials in Settings.'}

    orders = load_orders()

    all_tns = {}
    for order in orders:
        for box in order.get('packing_list', []):
            tn = box.get('tracking_number', '').strip()
            if tn:
                all_tns[tn] = order
        for m in order.get('manual_shipments', []):
            tn = m.get('tracking_number', '').strip()
            if tn:
                all_tns[tn] = order

    if not all_tns:
        return {'ok': True, 'checked': 0, 'updated': 0}

    results = {}
    errors  = []
    for tn in all_tns:
        try:
            r = None
            if use_fedex:
                r = _fetch_fedex(tn)
            if r is None and use_17track:
                r = _fetch_17track(tn)
            if r is not None:
                results[tn] = r
            time.sleep(0.3)
        except Exception as e:
            errors.append(f'{tn}: {e}')

    # Apply changes
    orders  = load_orders()
    updated = 0
    for order in orders:
        statuses = order.setdefault('tracking_statuses', {})
        # Collect all unique tracking numbers for this order (packing list + manual)
        order_tns = set()
        for box in order.get('packing_list', []):
            tn = box.get('tracking_number', '').strip()
            if tn: order_tns.add(tn)
        for m in order.get('manual_shipments', []):
            tn = m.get('tracking_number', '').strip()
            if tn: order_tns.add(tn)
        for tn in order_tns:
            if tn not in results:
                continue
            new     = results[tn]
            cur     = statuses.get(tn, {})
            new_tag = new['tag']

            # Extract delivery date from carrier event timestamp
            delivered_at = ''
            if new.get('event_time'):
                try:
                    et = new['event_time'][:10]
                    datetime.strptime(et, '%Y-%m-%d')
                    delivered_at = et
                except (ValueError, TypeError):
                    pass

            # Skip if status unchanged AND delivered_at already recorded
            if new_tag == cur.get('tag') and (new_tag != 'Delivered' or cur.get('delivered_at')):
                continue

            note = new['last_event']
            if new['location']:
                note += f" ({new['location']})"
            entry = {
                'tag':        new_tag,
                'label':      _tag_info(new_tag)[0],
                'notes':      note,
                'updated_at': _now_cst().isoformat(),
                'auto':       True,
            }
            if new_tag == 'Delivered':
                entry['delivered_at'] = delivered_at or cur.get('delivered_at', '')
            statuses[tn] = entry
            updated += 1

            if notify_changes and _should_notify_today(new.get('event_time', '')):
                msg  = _build_mac_message(order, tn, new_tag, note)
                label, _ = _tag_info(new_tag)
                subj = (f"Shipment Update — {order.get('market')} {order.get('season')} "
                        f"{order.get('sport')} #{order.get('num_for_code', order.get('order_num'))} · {label}")
                send_mac_notification(msg)
                send_email_notification(subj, msg)

    save_orders(orders)
    _last_auto_check = {'time': _fmt_cst(), 'error': None if not errors else '; '.join(errors)}
    return {'ok': True, 'checked': len(all_tns), 'updated': updated, 'errors': errors}


def _start_fedex_poller():
    """Background thread: poll tracking every 2h (FedEx direct) or 4h (17track fallback)."""
    def loop():
        time.sleep(60)
        while True:
            has_fedex   = bool(os.getenv('FEDEX_API_KEY') and os.getenv('FEDEX_SECRET_KEY'))
            has_17track = bool(os.getenv('TRACK17_API_KEY'))
            if has_fedex or has_17track:
                try:
                    check_all_fedex_statuses(notify_changes=True)
                except Exception:
                    pass
            time.sleep(2 * 3600 if has_fedex else 4 * 3600)

    threading.Thread(target=loop, daemon=True).start()


_start_fedex_poller()


TRACKING_STATUSES = [
    ('Pending',         'Pending',                    'gray',   0),
    ('LabelCreated',    'Label Created',              'purple', 1),
    ('PickedUp',        'Picked Up',                  'purple', 2),
    ('InTransitIntl',   'In Transit — International', 'blue',   3),
    ('AtCustoms',       'At US Customs',              'yellow', 4),
    ('ClearedCustoms',  'Cleared Customs',            'cyan',   5),
    ('InTransitDom',    'In Transit — Domestic',      'blue',   6),
    ('OutForDelivery',  'Out for Delivery',           'cyan',   7),
    ('Delivered',       'Delivered',                  'green',  8),
    ('Exception',       'Exception / Delay',          'red',    -1),
]

_STATUS_STEPS = [t for t in TRACKING_STATUSES if t[3] >= 0]   # ordered steps (no Exception)

_STATUS_MAP = {tag: (label, color) for tag, label, color, _ in TRACKING_STATUSES}


def _tag_info(tag):
    return _STATUS_MAP.get(tag, ('Unknown', 'gray'))


def send_slack_message(text: str) -> str:
    """Post a message to the configured Slack webhook. Returns error string or ''."""
    url = os.getenv('SLACK_WEBHOOK_URL', '')
    if not url:
        return 'Slack webhook not configured'
    try:
        resp = requests.post(url, json={'text': text}, timeout=10)
        if resp.status_code != 200:
            return f'Slack error {resp.status_code}: {resp.text[:100]}'
    except requests.RequestException as e:
        return str(e)
    return ''


def _build_mac_message(order: dict, tn: str, tag: str, notes: str) -> str:
    label, _ = _tag_info(tag)
    emoji     = _status_emoji(tag)
    market    = order.get('market', '')
    season    = order.get('season', '')
    sport     = order.get('sport', '')
    order_num = order.get('num_for_code', order.get('order_num', ''))

    shipments = _collect_shipments(order)
    ship      = next((s for s in shipments if s['tracking_number'] == tn), None)
    carrier   = _carrier_emoji(ship.get('carrier', 'fedex')) if ship else 'FedEx'

    lines = [
        f"{emoji} Shipment Update — {market} {season} {sport}",
        f"Order #{order_num}",
        "",
        f"Tracking: {tn} ({carrier})",
    ]

    if ship and ship['gear_summary']:
        lines.append("Contents:")
        # Group gear types by their notes (same notes = same physical box).
        # Gear with no notes falls back to grouping by box count.
        noted:   dict = {}   # notes_text → [gear, ...]
        unnoted: dict = {}   # box_count  → [gear, ...]
        for g in ship['gear_summary']:
            n = g.get('notes', '').strip()
            if n:
                noted.setdefault(n, {'gears': [], 'boxes': g['boxes']})['gears'].append(g['gear'])
            else:
                unnoted.setdefault(g['boxes'], []).append(g['gear'])
        # Emit noted groups in insertion order, then unnoted groups by count
        for group in noted.values():
            bw = 'box' if group['boxes'] == 1 else 'boxes'
            lines.append(f"  • {', '.join(sorted(group['gears']))} × {group['boxes']} {bw}")
        for count, gears in sorted(unnoted.items()):
            bw = 'box' if count == 1 else 'boxes'
            lines.append(f"  • {', '.join(sorted(gears))} × {count} {bw}")

    lines.append(f"Status: {emoji} {label}")
    if notes:
        lines.append(f"Notes: {notes}")
    lines.append("")
    lines.append(f"Updated {_fmt_cst()}")
    return '\n'.join(lines)


def _show_mac_popup(message: str):
    safe = message.replace('\\', '\\\\').replace('"', '\\"')
    script = f'''
set msg to "{safe}"
set the clipboard to msg
display dialog msg with title "MSM Shipment Tracker" buttons {{"Copy", "OK"}} default button "OK"
-- clicking Copy re-copies in case user cleared clipboard
if button returned of the result is "Copy" then set the clipboard to msg
'''
    subprocess.run(['osascript', '-e', script], timeout=120)


def _should_notify_today(event_time_str: str) -> bool:
    """Return True if a notification should fire for this event.
    - Weekday event → notify only on that same calendar day.
    - Weekend event (Sat/Sun) → notify on the following Monday.
    - No timestamp → always notify (conservative fallback).
    """
    today = _now_cst().date()
    if not event_time_str:
        return True
    try:
        event_date = datetime.fromisoformat(event_time_str[:10]).date()
    except (ValueError, TypeError):
        return True
    wd = event_date.weekday()          # 0=Mon … 6=Sun
    if wd < 5:
        notify_until = event_date      # same weekday only
    else:
        notify_until = event_date + timedelta(days=7 - wd)  # Sat→+2 Mon, Sun→+1 Mon
    return today <= notify_until


def send_mac_notification(message: str):
    """Show a macOS pop-up dialog in a background thread (non-blocking)."""
    threading.Thread(target=_show_mac_popup, args=(message,), daemon=True).start()


def send_email_notification(subject: str, body: str) -> str:
    """Send an email from/to gear@middleschoolmatchup.com. Returns error string or ''."""
    password = os.getenv('EMAIL_APP_PASSWORD', '')
    if not password:
        return 'Email app password not configured'
    sender = 'gear@middleschoolmatchup.com'
    msg = MIMEText(body, 'plain')
    msg['Subject'] = subject
    msg['From']    = f'MSM Shipment Tracker <{sender}>'
    msg['To']      = sender
    try:
        with smtplib.SMTP('smtp.gmail.com', 587) as s:
            s.starttls()
            s.login(sender, password)
            s.send_message(msg)
    except Exception as e:
        return str(e)
    return ''


def _carrier_emoji(carrier: str) -> str:
    return {'fedex': 'FedEx', 'ups': 'UPS', 'usps': 'USPS',
            'dhl': 'DHL', 'fedex-freight': 'FedEx Freight'}.get(carrier.lower(), carrier.upper())


def _status_emoji(tag: str) -> str:
    return {
        'Pending':        '⏳',
        'LabelCreated':   '🏷️',
        'PickedUp':       '📬',
        'InTransitIntl':  '✈️',
        'AtCustoms':      '🛃',
        'ClearedCustoms': '✅',
        'InTransitDom':   '🚛',
        'OutForDelivery': '🚚',
        'Delivered':      '📦',
        'Exception':      '🚨',
    }.get(tag, '📦')


# ── Order Registry ────────────────────────────────────────────────────────────

ORDER_REGISTRY = DATA_FOLDER / 'orders.json'


def _backfill_num_for_code(order: dict) -> None:
    """Ensure num_for_code is set for orders created before the field existed."""
    if not order.get('num_for_code'):
        raw = str(order.get('order_num', '1'))
        _m = re.search(r'(\d+)$', raw.strip())
        order['num_for_code'] = raw.strip() if raw.strip().isdigit() else (_m.group(1) if _m else '1')


def load_orders() -> list:
    if not ORDER_REGISTRY.exists():
        return []
    orders = json.loads(ORDER_REGISTRY.read_text())
    for o in orders:
        _backfill_num_for_code(o)
    return orders


def save_orders(orders: list):
    ORDER_REGISTRY.write_text(json.dumps(orders, indent=2))


def find_order(market: str, order_num: str, sport: str = None):
    """Find a saved order by market, order number, and optionally sport.
    When sport is provided, only returns an order whose sport matches exactly.
    Houston BSBL, Houston SOC, and Houston GLF are treated as separate orders.
    """
    m = market.lower().strip()
    for o in load_orders():
        if o['market'].lower().strip() != m:
            continue
        if str(o['order_num']) != str(order_num):
            continue
        if sport is not None and o.get('sport', '').upper() != sport.upper():
            continue
        return o
    return None


def upsert_order(order: dict):
    orders = load_orders()
    for i, o in enumerate(orders):
        if o['id'] == order['id']:
            orders[i] = order
            save_orders(orders)
            return
    orders.append(order)
    save_orders(orders)


def save_data(rows: list, meta: dict = None) -> str:
    session_id = session.get('data_id') or str(uuid.uuid4())
    session['data_id'] = session_id
    path = DATA_FOLDER / f'{session_id}.json'
    path.write_text(json.dumps({'rows': rows, 'meta': meta or {}}))
    (DATA_FOLDER / 'latest.txt').write_text(session_id)
    return session_id


def load_data() -> tuple:
    session_id = session.get('data_id')
    if not session_id:
        latest = DATA_FOLDER / 'latest.txt'
        if latest.exists():
            session_id = latest.read_text().strip()
            if session_id:
                session['data_id'] = session_id
    if not session_id:
        return [], {}
    path = DATA_FOLDER / f'{session_id}.json'
    if not path.exists():
        return [], {}
    stored = json.loads(path.read_text())
    if isinstance(stored, list):
        return stored, {}
    return stored.get('rows', []), stored.get('meta', {})


def clickup_headers():
    return {
        'Authorization': get_settings()['clickup_token'],
        'Content-Type': 'application/json',
    }


def create_clickup_task(payload: dict):
    list_id = get_settings()['clickup_list_id']
    url = f"https://api.clickup.com/api/v2/list/{list_id}/task"
    return requests.post(url, headers=clickup_headers(), json=payload, timeout=10)


def set_clickup_field(task_id: str, field_id: str, value):
    return requests.post(
        f"https://api.clickup.com/api/v2/task/{task_id}/field/{field_id}",
        headers=clickup_headers(), json={'value': value}, timeout=10,
    )


def update_task_for_shipment(task_id: str, tracking_str: str, ship_date_ms):
    """Update an existing Submitted task to Shipped with tracking info."""
    set_clickup_field(task_id, CF_PHASE,         3)              # Shipped
    set_clickup_field(task_id, CF_TRACKING,       tracking_str)
    if ship_date_ms:
        set_clickup_field(task_id, CF_SHIP_DATE_ACT, ship_date_ms)


# ── Delivered sync helpers ───────────────────────────────────────────────────

def _get_clickup_task(task_id: str):
    resp = requests.get(
        f"https://api.clickup.com/api/v2/task/{task_id}",
        headers=clickup_headers(), timeout=10,
    )
    return resp.json() if resp.status_code == 200 else None


def _get_clickup_subtasks(parent_task_id: str) -> list:
    list_id = get_settings().get('clickup_list_id', '')
    resp = requests.get(
        f"https://api.clickup.com/api/v2/list/{list_id}/task",
        headers=clickup_headers(),
        params={'parent': parent_task_id, 'subtasks': 'true', 'page': 0},
        timeout=10,
    )
    return resp.json().get('tasks', []) if resp.status_code == 200 else []


def _get_task_cf(task: dict, field_id: str):
    for cf in (task.get('custom_fields') or []):
        if cf.get('id') == field_id:
            return cf.get('value')
    return None


def _delivery_order_status(estimated_ms, actual_ms) -> int:
    """Return Order Status orderindex: 0=Ahead, 1=On Schedule, 2=Behind."""
    if estimated_ms is None or actual_ms is None:
        return 1
    est_day = int(estimated_ms) // 86400000
    act_day = int(actual_ms)    // 86400000
    if act_day < est_day:
        return 0   # Ahead — delivered before estimate
    elif act_day > est_day:
        return 2   # Behind — delivered after estimate
    return 1       # On Schedule


def _sync_delivered_to_clickup(order: dict, tn: str, actual_delivery_ms: int) -> list:
    """
    For all gears in this shipment:
      - Set CF_PHASE = 4 (Delivered), CF_DELIVERY_DATE_ACT, CF_ORDER_STATUS on parent task.
      - Find subtasks with CF_TRACKING == tn and apply same updates.
      - If no matching subtask exists, create one under the parent.
    Returns list of error strings (empty = all good).
    """
    errors = []
    if not get_settings().get('clickup_token'):
        return ['ClickUp not configured.']

    gear_tasks = order.get('gear_tasks', {})
    shipments  = _collect_shipments(order)
    ship       = next((s for s in shipments if s['tracking_number'] == tn), None)
    if not ship:
        return [f'Tracking {tn} not found in order.']

    gears = [s['gear'] for s in ship['gear_summary']]

    for gear in gears:
        task_id = gear_tasks.get(gear)
        if not task_id:
            errors.append(f'{gear}: no task ID')
            continue

        # Fetch parent to get estimated delivery date
        parent = _get_clickup_task(task_id)
        estimated_ms = None
        if parent:
            val = _get_task_cf(parent, CF_DELIVERY_DATE)
            if val is not None:
                try:
                    estimated_ms = int(val)
                except (TypeError, ValueError):
                    pass

        order_status = _delivery_order_status(estimated_ms, actual_delivery_ms)

        # Update parent task
        for field_id, value in [
            (CF_PHASE,            4),
            (CF_DELIVERY_DATE_ACT, actual_delivery_ms),
            (CF_ORDER_STATUS,     order_status),
        ]:
            r = set_clickup_field(task_id, field_id, value)
            if r.status_code not in (200, 201):
                errors.append(f'{gear} parent CF {field_id}: {r.status_code}')

        # Update any subtasks that have this tracking number
        subtasks = _get_clickup_subtasks(task_id)
        for sub in subtasks:
            if _get_task_cf(sub, CF_TRACKING) == tn:
                for field_id, value in [
                    (CF_PHASE,            4),
                    (CF_DELIVERY_DATE_ACT, actual_delivery_ms),
                    (CF_ORDER_STATUS,     order_status),
                ]:
                    r = set_clickup_field(sub['id'], field_id, value)
                    if r.status_code not in (200, 201):
                        errors.append(f'Subtask {sub["id"]} CF {field_id}: {r.status_code}')

    return errors


# ── Routes ──────────────────────────────────────────────────────────────────

@app.route('/')
def index():
    settings = get_settings()
    missing = [k for k, v in settings.items() if not v]

    # Collect all manual shipments for the packing list table
    orders_all = load_orders()
    manual_rows = []
    for o in sorted(orders_all, key=lambda x: x.get('created_at', ''), reverse=True):
        tns_map: dict = {}  # tn -> first matching manual entry (for metadata)
        for m in o.get('manual_shipments', []):
            tn = m.get('tracking_number', '').strip()
            if not tn:
                continue
            if tn not in tns_map:
                tns_map[tn] = m
            else:
                # accumulate gear types
                existing = tns_map[tn].get('gear_type', '')
                new_gt = m.get('gear_type', '')
                if new_gt and new_gt not in existing:
                    tns_map[tn] = dict(tns_map[tn])
                    tns_map[tn]['gear_type'] = ', '.join(filter(None, [existing, new_gt]))
        market = o.get('market', '')
        sport  = o.get('sport', '')
        onum   = o.get('num_for_code', o.get('order_num', ''))
        for tn, m in tns_map.items():
            manual_rows.append({
                'order_id':      o['id'],
                'order_label':   f"{market} {sport} #{onum}",
                'tn':            tn,
                'gear_type':     m.get('gear_type', ''),
                'notes':         m.get('notes', ''),
                'weight':        m.get('weight', ''),
                'declared_value': m.get('declared_value', ''),
            })

    return render_template('index.html', missing_settings=missing, manual_rows=manual_rows)


@app.route('/api/update-shipment-meta', methods=['POST'])
def update_shipment_meta():
    """Save weight and declared_value for a manual shipment entry."""
    data    = request.get_json(force=True)
    oid     = data.get('order_id', '').strip()
    tn      = data.get('tracking_number', '').strip()
    weight  = data.get('weight', '').strip()
    declared = data.get('declared_value', '').strip()
    if not oid or not tn:
        return jsonify({'error': 'order_id and tracking_number required'}), 400

    orders = load_orders()
    order = next((o for o in orders if o.get('id') == oid), None)
    if not order:
        return jsonify({'error': 'Order not found'}), 404

    updated = False
    for m in order.get('manual_shipments', []):
        if m.get('tracking_number', '').strip() == tn:
            m['weight']        = weight
            m['declared_value'] = declared
            updated = True

    if not updated:
        return jsonify({'error': 'Tracking number not found in order'}), 404

    upsert_order(order)
    return jsonify({'ok': True})


@app.route('/upload', methods=['POST'])
def upload():
    if 'file' not in request.files:
        return redirect(url_for('index'))
    f = request.files['file']
    if not f.filename:
        return redirect(url_for('index'))

    filename = secure_filename(f.filename)
    filepath = UPLOAD_FOLDER / filename
    f.save(str(filepath))

    try:
        rows = parse_packing_list(str(filepath))
    except Exception as e:
        return render_template('index.html', error=str(e), missing_settings=[])

    session['filename'] = filename
    save_data(rows)
    return redirect(url_for('preview'))


@app.route('/preview')
def preview():
    rows, meta = load_data()
    filename = session.get('filename', '')
    if not rows:
        return redirect(url_for('index'))

    sample = rows[0]
    matched_order = find_order(sample['market'], sample['order_num'])

    # Auto-fill meta from saved order if fields not yet set by user
    if matched_order:
        for key in ('sport', 'season', 'order_date', 'order_sheet_url'):
            if not meta.get(key) and matched_order.get(key):
                meta[key] = matched_order[key]
        save_data(rows, meta)

    # Auto-match: if a tracking number in this packing list already exists in any
    # order's manual_shipments, save the per-box rows to that order's packing_list
    # immediately so Contents page shows real detail without a full ClickUp push.
    tn_to_rows: dict = {}
    for r in rows:
        tn = r.get('tracking_number', '').strip()
        if tn:
            tn_to_rows.setdefault(tn, []).append(r)

    all_orders = load_orders()
    auto_matched: list = []
    for o in all_orders:
        matched_tns = set()
        for m in o.get('manual_shipments', []):
            tn = m.get('tracking_number', '').strip()
            if tn and tn in tn_to_rows and tn not in matched_tns:
                matched_tns.add(tn)
                o.setdefault('packing_list', [])
                o['packing_list'] = [x for x in o['packing_list']
                                     if x.get('tracking_number') != tn]
                o['packing_list'].extend(tn_to_rows[tn])
                auto_matched.append({
                    'order': f"{o.get('market','')} {o.get('sport','')} #{o.get('num_for_code', o.get('order_num',''))}",
                    'tn': tn,
                    'boxes': len(tn_to_rows[tn]),
                })
    if auto_matched:
        save_orders(all_orders)

    # Which gear types already have a ClickUp task saved
    matched_gear_tasks = matched_order.get('gear_tasks', {}) if matched_order else {}

    by_gear = defaultdict(list)
    for r in rows:
        by_gear[r['gear_type']].append(r)

    tracking_letters, _, _ = build_shipment_map(rows)

    settings = get_settings()
    has_settings = all(settings.values())
    today = datetime.now().strftime('%Y-%m-%d')

    return render_template('preview.html',
                           rows=rows,
                           by_gear=dict(by_gear),
                           filename=filename,
                           carriers=CARRIERS,
                           sports=SPORTS,
                           seasons=SEASONS,
                           has_settings=has_settings,
                           meta=meta,
                           today=today,
                           matched_order=matched_order,
                           matched_gear_tasks=matched_gear_tasks,
                           tracking_letters=tracking_letters,
                           auto_matched=auto_matched)


PHASE_LABELS = ['Submitted', 'Proofs', 'Production', 'Shipped', 'Delivered', 'Counted']


@app.route('/orders')
def orders_list():
    orders = load_orders()
    orders_sorted = sorted(orders, key=lambda o: o.get('created_at', ''), reverse=True)
    for o in orders_sorted:
        o['shipments'] = _collect_shipments(o)
        shipped_gears = {g['gear'] for s in o['shipments'] for g in s['gear_summary']}
        o['unshipped_gears'] = [g for g in o.get('gear_tasks', {}).keys() if g not in shipped_gears]

    # Build flat list of unique (order, tn) pairs across all manual shipments
    manual_tns = []
    for o in orders_sorted:
        statuses = o.get('tracking_statuses', {})
        tns_map: dict = {}
        for m in o.get('manual_shipments', []):
            tn = m.get('tracking_number', '').strip()
            if not tn:
                continue
            if statuses.get(tn, {}).get('tag') == 'Delivered':
                continue
            gt = m.get('gear_type', '')
            if tn not in tns_map:
                tns_map[tn] = []
            if gt and gt not in tns_map[tn]:
                tns_map[tn].append(gt)
        market = o.get('market', '')
        sport  = o.get('sport', '')
        onum   = o.get('num_for_code', o.get('order_num', ''))
        for tn, gears in tns_map.items():
            manual_tns.append({
                'order_id':    o['id'],
                'order_label': f"{market} {sport} #{onum}",
                'tn':          tn,
                'gears':       ', '.join(gears) if gears else '',
            })

    return render_template('orders.html', orders=orders_sorted, phase_labels=PHASE_LABELS, manual_tns=manual_tns)


@app.route('/api/order-phases/<order_id>')
def api_order_phases(order_id):
    """Return current CF_PHASE per gear type for an order."""
    orders = load_orders()
    order  = next((o for o in orders if o.get('id') == order_id), None)
    if not order:
        return jsonify({'ok': False, 'error': 'Order not found.'})

    gear_tasks = order.get('gear_tasks', {})
    result = {}
    for gear, task_id in gear_tasks.items():
        t = _get_clickup_task(task_id)
        if t:
            val = _get_task_cf(t, CF_PHASE)
            try:
                result[gear] = int(val) if val is not None else None
            except (TypeError, ValueError):
                result[gear] = None
        else:
            result[gear] = None

    return jsonify({'ok': True, 'phases': result, 'labels': PHASE_LABELS})


@app.route('/push/update-phase', methods=['POST'])
def push_update_phase():
    if not get_settings()['clickup_token']:
        return jsonify({'ok': False, 'error': 'ClickUp not configured.'})

    body            = request.get_json()
    order_id        = body.get('order_id', '')
    phase           = body.get('phase')
    tracking_number = body.get('tracking_number', '').strip()
    proofs_actual   = body.get('proofs_actual', '')   # YYYY-MM-DD, phase 1 only
    rcv_orderindex  = body.get('receiving_orderindex') # phase 5 only
    qty_note        = body.get('qty_note', '').strip()

    orders = load_orders()
    order  = next((o for o in orders if o.get('id') == order_id), None)
    if not order:
        return jsonify({'ok': False, 'error': 'Order not found.'})

    gear_type  = body.get('gear_type', '').strip()
    gear_tasks = order.get('gear_tasks', {})
    if gear_type:
        gear_tasks = {g: tid for g, tid in gear_tasks.items() if g == gear_type}
    elif tracking_number:
        gear_types_in = {
            box.get('gear_type', '') for box in order.get('packing_list', [])
            if box.get('tracking_number', '').strip() == tracking_number
        }
        gear_tasks = {g: tid for g, tid in gear_tasks.items() if g in gear_types_in}

    proofs_actual_ms = None
    if phase == 1 and proofs_actual:
        try:
            dt = datetime.strptime(proofs_actual, '%Y-%m-%d')
            proofs_actual_ms = int(dt.replace(hour=0, minute=0, second=0, microsecond=0).timestamp() * 1000)
        except ValueError:
            pass

    errors  = []
    updated = 0
    for gear, task_id in gear_tasks.items():
        r = set_clickup_field(task_id, CF_PHASE, phase)
        if r.status_code not in (200, 201):
            errors.append(f"{gear}: {r.status_code}")
            continue
        updated += 1

        order_status_val = None
        if proofs_actual_ms is not None:
            set_clickup_field(task_id, CF_PROOFS_APPROVED_ACT, proofs_actual_ms)
            task_data = _get_clickup_task(task_id)
            if task_data:
                try:
                    est_ms = int(_get_task_cf(task_data, CF_PROOFS_APPROVED) or 0) or None
                except (TypeError, ValueError):
                    est_ms = None
                order_status_val = _delivery_order_status(est_ms, proofs_actual_ms)
                set_clickup_field(task_id, CF_ORDER_STATUS, order_status_val)

        if rcv_orderindex is not None:
            set_clickup_field(task_id, CF_RECEIVING, rcv_orderindex)
            if qty_note:
                requests.post(
                    f"https://api.clickup.com/api/v2/task/{task_id}/comment",
                    headers=clickup_headers(), json={"comment_text": f"Qty note: {qty_note}"}, timeout=10,
                )

        # Update matching subtasks
        for sub in _get_clickup_subtasks(task_id):
            if tracking_number and _get_task_cf(sub, CF_TRACKING) != tracking_number:
                continue
            sub_id = sub['id']
            set_clickup_field(sub_id, CF_PHASE, phase)
            if proofs_actual_ms is not None:
                set_clickup_field(sub_id, CF_PROOFS_APPROVED_ACT, proofs_actual_ms)
                if order_status_val is not None:
                    set_clickup_field(sub_id, CF_ORDER_STATUS, order_status_val)
            if rcv_orderindex is not None:
                set_clickup_field(sub_id, CF_RECEIVING, rcv_orderindex)
                if qty_note:
                    requests.post(
                        f"https://api.clickup.com/api/v2/task/{sub_id}/comment",
                        headers=clickup_headers(), json={"comment_text": f"Qty note: {qty_note}"}, timeout=10,
                    )

    if errors:
        return jsonify({'ok': False, 'error': '; '.join(errors), 'updated': updated})
    return jsonify({'ok': True, 'message': f"Updated {updated} task(s) to {PHASE_LABELS[phase]}."})


@app.route('/push/update-ship', methods=['POST'])
def push_update_ship():
    """Mark a gear type as Shipped: set tracking #, ship date actual, phase=3 on task + subtasks."""
    if not get_settings()['clickup_token']:
        return jsonify({'ok': False, 'error': 'ClickUp not configured.'})

    body            = request.get_json()
    order_id        = body.get('order_id', '').strip()
    gear_type       = body.get('gear_type', '').strip()
    tracking_number = body.get('tracking_number', '').strip()
    ship_date       = body.get('ship_date', '').strip()   # YYYY-MM-DD
    notes           = body.get('notes', '').strip()
    declared_value  = body.get('declared_value', '').strip()
    try:
        box_count = max(1, int(body.get('box_count', 1)))
    except (TypeError, ValueError):
        box_count = 1

    if not tracking_number:
        return jsonify({'ok': False, 'error': 'Tracking number is required.'})

    orders = load_orders()
    order  = next((o for o in orders if o.get('id') == order_id), None)
    if not order:
        return jsonify({'ok': False, 'error': 'Order not found.'})

    gear_tasks = order.get('gear_tasks', {})
    if gear_type:
        gear_tasks = {g: tid for g, tid in gear_tasks.items() if g == gear_type}

    ship_date_ms = None
    if ship_date:
        try:
            dt = datetime.strptime(ship_date, '%Y-%m-%d')
            ship_date_ms = int(dt.replace(hour=0, minute=0, second=0, microsecond=0).timestamp() * 1000)
        except ValueError:
            pass

    comment_text = f"📦 Shipped\nTracking: {tracking_number}"
    if ship_date:
        comment_text += f"\nShip Date: {ship_date}"
    if declared_value:
        comment_text += f"\nDeclared Value: ${declared_value}"
    if notes:
        comment_text += f"\nNotes: {notes}"

    errors  = []
    updated = 0
    for gear, task_id in gear_tasks.items():
        r = set_clickup_field(task_id, CF_PHASE, 3)
        if r.status_code not in (200, 201):
            errors.append(f"{gear}: {r.status_code}"); continue

        set_clickup_field(task_id, CF_TRACKING, tracking_number)
        if ship_date_ms:
            set_clickup_field(task_id, CF_SHIP_DATE_ACT, ship_date_ms)

        # Compare actual vs estimated ship date → set Order Status
        order_status_val = None
        if ship_date_ms:
            task_data = _get_clickup_task(task_id)
            if task_data:
                try:
                    est_ms = int(_get_task_cf(task_data, CF_SHIP_DATE) or 0) or None
                except (TypeError, ValueError):
                    est_ms = None
                order_status_val = _ship_order_status(ship_date_ms, est_ms)
                set_clickup_field(task_id, CF_ORDER_STATUS, order_status_val)

        requests.post(
            f"https://api.clickup.com/api/v2/task/{task_id}/comment",
            headers=clickup_headers(), json={"comment_text": comment_text}, timeout=10,
        )
        updated += 1

        for sub in _get_clickup_subtasks(task_id):
            existing_tn = _get_task_cf(sub, CF_TRACKING)
            if existing_tn and existing_tn != tracking_number:
                continue
            sub_id = sub['id']
            set_clickup_field(sub_id, CF_PHASE, 3)
            set_clickup_field(sub_id, CF_TRACKING, tracking_number)
            if ship_date_ms:
                set_clickup_field(sub_id, CF_SHIP_DATE_ACT, ship_date_ms)
            if order_status_val is not None:
                set_clickup_field(sub_id, CF_ORDER_STATUS, order_status_val)
            requests.post(
                f"https://api.clickup.com/api/v2/task/{sub_id}/comment",
                headers=clickup_headers(), json={"comment_text": comment_text}, timeout=10,
            )

    if not errors:
        # Save to order registry so shipment appears in the shipments view
        manual = order.setdefault('manual_shipments', [])
        # Replace existing entry for same gear+tracking or append new
        existing = next((m for m in manual if m['gear_type'] == gear_type
                         and m['tracking_number'] == tracking_number), None)
        entry = {
            'gear_type':       gear_type,
            'tracking_number': tracking_number,
            'ship_date':       ship_date,
            'declared_value':  declared_value,
            'notes':           notes,
            'box_count':       box_count,
            'added_at':        datetime.now().isoformat(),
            'carrier':         'FedEx',
        }
        if existing:
            existing.update(entry)
        else:
            manual.append(entry)
        if not order.get('shipped_at'):
            order['shipped_at'] = datetime.now().isoformat()
        upsert_order(order)

    if errors:
        return jsonify({'ok': False, 'error': '; '.join(errors), 'updated': updated})
    return jsonify({'ok': True, 'message': f"Marked {updated} task(s) as Shipped."})


@app.route('/order-form')
def order_form():
    settings = get_settings()
    has_settings = all(settings.values())
    today = datetime.now().strftime('%Y-%m-%d')
    return render_template('order_form.html',
                           sports=SPORTS,
                           seasons=SEASONS,
                           markets=MARKETS,
                           gear_types=ALL_GEAR_TYPES,
                           has_settings=has_settings,
                           today=today)


@app.route('/fetch-order-sheet', methods=['POST'])
def fetch_order_sheet():
    body = request.get_json()
    url = (body.get('url') or '').strip()
    m = re.search(r'/spreadsheets/d/([a-zA-Z0-9_-]+)', url)
    if not m:
        return jsonify({'ok': False, 'error': 'Invalid Google Sheets URL'})

    sheet_id = m.group(1)

    season = (body.get('season') or '').strip()
    jacket_gear = 'Windbreaker' if season.lower() == 'summer' else 'Cage Jacket'

    SUMMARY_GEAR_MAP = {
        'jerseys': 'Jersey', 'jersey': 'Jersey',
        'hoodies': 'Hoodie', 'hoodie': 'Hoodie',
        'tees': 'Tee', 'tee': 'Tee',
        'jackets': jacket_gear, 'jacket': jacket_gear,
        'cage jackets': 'Cage Jacket', 'cage jacket': 'Cage Jacket',
        'windbreakers': 'Windbreaker', 'windbreaker': 'Windbreaker',
        'caps': 'Cap', 'cap': 'Cap',
        'rings - champs': 'Ring', 'rings - finalist': 'Ring', 'rings': 'Ring', 'ring': 'Ring',
        'hype chains': 'Hypechain', 'hype chain': 'Hypechain',
        'name banner': 'Name Banner', 'name banners': 'Name Banner',
        'park banner pack': 'Park Banner', 'park banner': 'Park Banner',
        'hybrid pullover': 'Hybrid Pullover', 'hybrid pullovers': 'Hybrid Pullover',
    }

    STOP_WORDS = {'total', 'thank you', 'grand total', ''}

    def parse_columns_from_csv(text):
        """
        Layout 1: Order Home Page tab.
        - Gear in col B (idx 1), preceded by 'ORDER SUMMARY' row.
        - Header row has col B = 'Gear' and cols E+ (idx 4+) = order group headers
          e.g. ['BSBL 1', 'FRESHMAN', 'BSBL 2', 'NEED COLORWAYS'].
        - If headers are absent but multiple qty columns exist, returns synthetic
          numeric headers '1', '2', '3'... so user can pick by number.
        Returns list of {'header': str, 'gear_types': [str, ...]} per column,
        or None if layout not found.
        """
        rows = list(csv.reader(StringIO(text)))
        in_summary = False
        col_headers = []
        col_gear = {}
        unnamed_gear_rows = []  # (gear_type, row) when no named headers found

        for row in rows:
            while len(row) < 15:
                row.append('')
            col1 = row[1].strip()
            col1_lower = col1.lower()

            if not in_summary:
                if col1_lower == 'order summary':
                    in_summary = True
                continue

            # Header row: col B = "Gear", cols E+ = order group names
            if col1_lower == 'gear':
                col_headers = [h.strip() for h in row[4:] if h.strip()]
                col_gear = {h: [] for h in col_headers}
                continue

            if col1_lower in STOP_WORDS:
                break
            if col1_lower == '':
                continue

            gear_type = SUMMARY_GEAR_MAP.get(col1_lower)
            if not gear_type:
                continue

            if col_headers:
                for i, header in enumerate(col_headers):
                    col_idx = 4 + i
                    qty = _safe_float(row[col_idx]) if col_idx < len(row) else 0
                    if qty > 0 and gear_type not in col_gear[header]:
                        col_gear[header].append(gear_type)
            else:
                unnamed_gear_rows.append((gear_type, row))

        if col_headers and any(col_gear[h] for h in col_headers):
            return [{'header': h, 'gear_types': col_gear[h]}
                    for h in col_headers if col_gear[h]]

        # Fallback: no named headers — detect multiple qty columns and create
        # synthetic numeric labels so the user can select by order number
        if unnamed_gear_rows:
            col_gear_map = {}
            for gear_type, row in unnamed_gear_rows:
                for col_idx in range(4, min(len(row), 12)):
                    qty = _safe_float(row[col_idx])
                    if qty > 0:
                        col_gear_map.setdefault(col_idx, [])
                        if gear_type not in col_gear_map[col_idx]:
                            col_gear_map[col_idx].append(gear_type)
            active_cols = sorted(col_gear_map.keys())
            if len(active_cols) > 1:
                return [{'header': str(i + 1), 'gear_types': col_gear_map[c]}
                        for i, c in enumerate(active_cols)]

        return None

    _date_re = re.compile(r'^(\d{1,2})[/\-](\d{1,2})[/\-](\d{2,4})$')

    def parse_order_date_from_csv(text):
        """Find the first date-shaped value in col B before the ORDER SUMMARY row."""
        for row in csv.reader(StringIO(text)):
            if len(row) < 2:
                continue
            col1 = row[1].strip()
            if col1.lower() == 'order summary':
                break
            m = _date_re.match(col1)
            if m:
                mo, day, yr = m.group(1), m.group(2), m.group(3)
                if len(yr) == 2:
                    yr = '20' + yr
                return f"{yr}-{mo.zfill(2)}-{day.zfill(2)}"
        return None

    def parse_gear_from_summary_csv(text):
        """
        Layout 1b: Order Home Page with a single qty column.
        - ORDER SUMMARY section, gear in col B (idx 1), qty in col E (idx 4+).
        - GEAR header row has empty E+ cells (no multi-order column headers).
        """
        rows = list(csv.reader(StringIO(text)))
        found = []
        in_summary = False
        for row in rows:
            while len(row) < 10:
                row.append('')
            col1 = row[1].strip()
            col1_lower = col1.lower()
            if not in_summary:
                if col1_lower == 'order summary':
                    in_summary = True
                continue
            if col1_lower in STOP_WORDS:
                break
            if col1_lower in ('gear', ''):
                continue
            gear_type = SUMMARY_GEAR_MAP.get(col1_lower)
            if not gear_type:
                continue
            has_qty = any(_safe_float(v) > 0 for v in row[4:] if v.strip())
            if has_qty and gear_type not in found:
                found.append(gear_type)
        return found

    def parse_gear_fallback(text):
        """Layout 2: simple tab with gear in col A, qty in col B."""
        rows = list(csv.reader(StringIO(text)))
        found = []
        for row in rows:
            if len(row) < 2:
                continue
            gear_type = SUMMARY_GEAR_MAP.get(row[0].strip().lower())
            if gear_type and _safe_float(row[1]) > 0 and gear_type not in found:
                found.append(gear_type)
        return found

    columns_found = None
    gear_found = []
    fetch_error = None

    for tab_name in ['Order Home Page', 'ORDER SUMMARY', 'Order Summary', 'Summary']:
        tab_encoded = requests.utils.quote(tab_name)
        gviz_url = (
            f"https://docs.google.com/spreadsheets/d/{sheet_id}"
            f"/gviz/tq?tqx=out:csv&sheet={tab_encoded}"
        )
        try:
            resp = requests.get(gviz_url, timeout=10, allow_redirects=True)
            if resp.status_code != 200:
                fetch_error = f"HTTP {resp.status_code}"
                continue
            if 'accounts.google.com' in resp.url:
                return jsonify({'ok': False, 'error': 'Sheet is private — share it with view access first.'})

            order_date_val = parse_order_date_from_csv(resp.text)
            columns_found = parse_columns_from_csv(resp.text)
            if columns_found:
                break
            gear_found = parse_gear_from_summary_csv(resp.text)
            if gear_found:
                break
            gear_found = parse_gear_fallback(resp.text)
            if gear_found:
                break
        except Exception as e:
            fetch_error = str(e)
            continue

    if columns_found:
        return jsonify({'ok': True, 'columns': columns_found, 'order_date': order_date_val})

    if gear_found:
        return jsonify({'ok': True, 'gear_types': gear_found, 'order_date': order_date_val})

    msg = 'Could not read order summary.'
    if fetch_error:
        msg += f' ({fetch_error})'
    msg += ' Make sure the sheet is shared with view access.'
    return jsonify({'ok': False, 'error': msg})


@app.route('/save', methods=['POST'])
def save():
    body = request.get_json()
    if body:
        if isinstance(body, list):
            save_data(body)
        else:
            save_data(body.get('rows', []), body.get('meta', {}))
    return jsonify({'ok': True})


# ── Integrations ─────────────────────────────────────────────────────────────

def _ship_order_status(actual_ms, estimated_ms) -> int:  # noqa: E302
    """Compare actual vs estimated ship date → Order Status orderindex."""
    if not actual_ms or not estimated_ms:
        return 1  # On Schedule (no data to compare)
    if actual_ms > estimated_ms:
        return 2  # Behind
    if actual_ms < estimated_ms:
        return 0  # Ahead
    return 1  # On Schedule


def _get_estimated_ship_ms(gear_type: str, order_date: str):
    """Pull the Ship Date ms from the calc_gear_dates result."""
    for f in calc_gear_dates(gear_type, order_date):
        if f['id'] == CF_SHIP_DATE:
            return f['value']
    return None


@app.route('/push/clickup', methods=['POST'])
def push_clickup():
    settings = get_settings()
    if not settings['clickup_token'] or not settings['clickup_list_id']:
        return jsonify({'ok': False, 'error': 'ClickUp not configured. Go to Settings.'})

    rows, meta = load_data()
    if not rows:
        return jsonify({'ok': False, 'error': 'No shipment data loaded.'})

    by_gear = defaultdict(list)
    for r in rows:
        by_gear[r['gear_type']].append(r)

    sample = rows[0]
    market    = sample['market']
    order_num = sample['order_num']

    sport       = meta.get('sport', 'BSBL')
    season      = meta.get('season', 'Fall')
    order_date  = meta.get('order_date', '')
    ship_date   = meta.get('ship_date', '')
    order_sheet = meta.get('order_sheet_url', '')

    order_date_ms    = to_ms(order_date)
    actual_ship_ms   = to_ms(ship_date)

    saved_order      = find_order(market, order_num)
    saved_gear_tasks = saved_order.get('gear_tasks', {}) if saved_order else {}
    new_gear_tasks   = dict(saved_gear_tasks)
    eta_zees_ms      = to_ms(saved_order.get('eta_to_zees', '')) if saved_order else None

    # Resolve numeric suffix used in task names / order codes
    _nfc = saved_order.get('num_for_code', '') if saved_order else ''
    if not _nfc:
        _m2 = re.search(r'(\d+)$', order_num.strip())
        _nfc = order_num.strip() if order_num.strip().isdigit() else (_m2.group(1) if _m2 else '1')
    num_for_code = _nfc

    # Build global shipment map: tracking# → letter, primary gear, rows
    t2letter, t2primary, t2rows = build_shipment_map(rows)

    results          = []
    errors           = []
    parent_task_ids  = {}   # gear_type → clickup task id
    gear_task_names  = {}   # gear_type → task name string
    gear_order_codes = {}   # gear_type → order code string
    gear_order_status = {}  # gear_type → order status orderindex

    # ── PASS 1: create / update one parent task per gear type ────────────────
    for gear_type, boxes in sorted(by_gear.items()):
        gear_trackings = list({b['tracking_number'] for b in boxes if b['tracking_number']})
        tracking_str   = ', '.join(gear_trackings)
        has_tracking   = bool(gear_trackings)

        task_name  = build_task_name(market, season, sport, gear_type, num_for_code)
        order_code = build_order_code(market, sport, gear_type, order_date, num_for_code)
        gear_task_names[gear_type]  = task_name
        gear_order_codes[gear_type] = order_code

        estimated_ship_ms = _get_estimated_ship_ms(gear_type, order_date)
        order_status      = _ship_order_status(
            actual_ship_ms if has_tracking else None, estimated_ship_ms)
        gear_order_status[gear_type] = order_status

        existing_task_id = saved_gear_tasks.get(gear_type)

        if existing_task_id:
            set_clickup_field(existing_task_id, CF_ORDER_STATUS,  order_status)
            set_clickup_field(existing_task_id, CF_TRACKING,      tracking_str)
            if has_tracking and actual_ship_ms:
                set_clickup_field(existing_task_id, CF_SHIP_DATE_ACT, actual_ship_ms)
            parent_task_ids[gear_type] = existing_task_id
            results.append({'gear': gear_type, 'task_id': existing_task_id,
                            'subtasks': [], 'action': 'updated'})
        else:
            desc_lines = [f"Total Boxes: {len(boxes)}"]
            if order_sheet:
                desc_lines.append(f"Order Sheet: {order_sheet}")

            custom_fields = [
                {'id': CF_ORDER_STATUS, 'value': order_status},
                {'id': CF_ORDER_TEXT,   'value': order_code},
                {'id': CF_TRACKING,     'value': tracking_str},
            ]
            if order_date_ms:
                custom_fields.append({'id': CF_ORDER_DATE, 'value': order_date_ms})
            if has_tracking and actual_ship_ms:
                custom_fields.append({'id': CF_SHIP_DATE_ACT, 'value': actual_ship_ms})
            custom_fields.extend(calc_gear_dates(gear_type, order_date))

            resp = create_clickup_task({
                'name': task_name,
                'description': '\n'.join(desc_lines),
                'status': 'to do',
                'custom_fields': custom_fields,
            })
            if resp.status_code not in (200, 201):
                errors.append(f"{gear_type}: {resp.status_code} {resp.text[:120]}")
                continue

            task_id = resp.json().get('id')
            parent_task_ids[gear_type] = task_id
            new_gear_tasks[gear_type]  = task_id
            results.append({'gear': gear_type, 'task_id': task_id,
                            'subtasks': [], 'action': 'created'})

    # ── PASS 2: one subtask per shipment (unique tracking #) ─────────────────
    for tn, letter in t2letter.items():
        primary_gear = t2primary[tn]
        parent_id    = parent_task_ids.get(primary_gear)
        if not parent_id:
            continue

        # Skip if a subtask with this tracking number already exists under this parent
        existing_subs = _get_clickup_subtasks(parent_id)
        if any(_get_task_cf(s, CF_TRACKING) == tn for s in existing_subs):
            continue

        shipment_rows  = t2rows[tn]
        gear_types_in  = sorted({r['gear_type'] for r in shipment_rows})
        contents_desc  = ', '.join(gear_types_in)
        weights        = [r['weight'] for r in shipment_rows if r.get('weight')]
        values         = [r['declared_value'] for r in shipment_rows if r.get('declared_value')]

        # Name and order code mirror the parent + shipment letter appended
        base_name  = gear_task_names.get(primary_gear, f"{market} {season} {sport} {primary_gear} {num_for_code}")
        base_code  = gear_order_codes.get(primary_gear, num_for_code)
        sub_name   = f"{base_name}{letter}"       # e.g. "Dallas Summer BSBL Jersey 1A"
        sub_code   = f"{base_code}{letter}"       # e.g. "DAL-BSBL-JRS-042426-1A"

        # Copy all non-empty parent fields, then override with subtask-specific values.
        # Skip formula/auto-calculated fields that ClickUp computes itself.
        _FORMULA_FIELD_IDS = {
            '08f6fefb-3605-48e4-bdbb-b5e1c765d87a',  # Days Ahead/Behind (formula)
            '43acbd5f-e47a-4230-93fd-cc7f596a731d',  # Total LT (auto)
            '0245e34f-f032-44c5-b016-792b4447c658',  # Shipping Time (auto)
            '9e1380de-7dbf-4c77-afdb-7de496f61e2b',  # Production Time (auto)
        }
        parent_task = _get_clickup_task(parent_id)
        sub_fields_map = {}
        if parent_task:
            for cf in parent_task.get('custom_fields', []):
                val = cf.get('value')
                if val is not None and cf['id'] not in _FORMULA_FIELD_IDS:
                    sub_fields_map[cf['id']] = val

        # Subtask-specific overrides (always set regardless of parent value)
        sub_fields_map[CF_TRACKING]   = tn
        sub_fields_map[CF_ORDER_TEXT] = sub_code

        sub_fields = [{'id': fid, 'value': val} for fid, val in sub_fields_map.items()]

        sresp = create_clickup_task({
            'name': sub_name,
            'parent': parent_id,
            'description': (
                f"Gear: {contents_desc}\n"
                f"Weight: {'; '.join(weights)}\n"
                f"Declared Value: {'; '.join(values)}\n"
                f"Carrier: {shipment_rows[0].get('carrier', 'fedex').upper()}"
            ),
            'custom_fields': sub_fields,
        })
        if sresp.status_code in (200, 201):
            for r in results:
                if r['gear'] == primary_gear:
                    r['subtasks'].append(sresp.json().get('id'))
                    break
        else:
            errors.append(f"  Shipment {order_num}{letter}: {sresp.status_code}")

    # ── Persist to registry ───────────────────────────────────────────────────
    order_record = {
        'market': market, 'sport': sport, 'season': season,
        'order_num': order_num, 'order_date': order_date,
        'order_sheet_url': order_sheet,
        'gear_tasks': new_gear_tasks,
        'packing_list': rows,
        'packing_meta': meta,
        'shipped_at': datetime.now().isoformat(),
    }
    if saved_order:
        saved_order.update(order_record)
        upsert_order(saved_order)
    else:
        order_record['id'] = str(uuid.uuid4())
        order_record['created_at'] = datetime.now().isoformat()
        upsert_order(order_record)

    ok      = len(errors) == 0
    updated = [r for r in results if r.get('action') == 'updated']
    created = [r for r in results if r.get('action') == 'created']
    parts   = []
    if updated:   parts.append(f"Updated {len(updated)} task(s)")
    if created:   parts.append(f"Created {len(created)} task(s)")
    sub_count = sum(len(r['subtasks']) for r in results)
    if sub_count: parts.append(f"{sub_count} shipment subtask(s) added")
    if errors:    parts.append(f"Errors: {'; '.join(errors)}")

    return jsonify({'ok': ok, 'message': '. '.join(parts) or 'Done.', 'results': results})


@app.route('/push/clickup-order-only', methods=['POST'])
def push_clickup_order_only():
    settings = get_settings()
    if not settings['clickup_token'] or not settings['clickup_list_id']:
        return jsonify({'ok': False, 'error': 'ClickUp not configured. Go to Settings.'})

    body = request.get_json()
    if not body:
        return jsonify({'ok': False, 'error': 'No data.'})

    market      = body.get('market', '').strip()
    sport       = body.get('sport', 'BSBL')
    season      = body.get('season', 'Fall')
    order_num   = str(body.get('order_num', '1'))
    order_date  = body.get('order_date', '')
    eta_zees    = body.get('eta_to_zees', '')
    order_sheet = body.get('order_sheet_url', '')
    gear_types  = body.get('gear_types', [])

    if not market or not gear_types:
        return jsonify({'ok': False, 'error': 'Market and at least one gear type are required.'})

    order_date_ms = to_ms(order_date)
    eta_zees_ms   = to_ms(eta_zees)

    # Non-numeric labels (Freshman, BSBL 2…) → use "1" in task names/codes
    _m = re.search(r'(\d+)$', order_num.strip())
    num_for_code = order_num.strip() if order_num.strip().isdigit() else (_m.group(1) if _m else '1')

    saved_order = find_order(market, order_num, sport)
    saved_gear_tasks = saved_order.get('gear_tasks', {}) if saved_order else {}

    results = []
    errors  = []

    for gear_type in gear_types:
        task_name   = build_task_name(market, season, sport, gear_type, num_for_code)
        order_code  = build_order_code(market, sport, gear_type, order_date, num_for_code)
        date_fields = calc_gear_dates(gear_type, order_date)

        existing_task_id = saved_gear_tasks.get(gear_type)

        if existing_task_id:
            # ── UPDATE existing task ──────────────────────────────────────────
            fields_to_set = [
                {'id': CF_ORDER_STATUS, 'value': 1},
            ]
            if order_date_ms:
                fields_to_set.append({'id': CF_ORDER_DATE, 'value': order_date_ms})
            if eta_zees_ms:
                fields_to_set.append({'id': CF_ETA_ZEES, 'value': eta_zees_ms})
            if order_sheet:
                # update description
                requests.put(
                    f"https://api.clickup.com/api/v2/task/{existing_task_id}",
                    headers=clickup_headers(),
                    json={'description': f"Order Sheet: {order_sheet}"},
                    timeout=10,
                )
            fields_to_set.extend(date_fields)
            for f in fields_to_set:
                set_clickup_field(existing_task_id, f['id'], f['value'])
            results.append({'gear': gear_type, 'task_id': existing_task_id, 'action': 'updated'})

        else:
            # ── CREATE new task ───────────────────────────────────────────────
            desc_lines = []
            if order_sheet:
                desc_lines.append(f"Order Sheet: {order_sheet}")

            custom_fields = [
                {'id': CF_PHASE,        'value': 0},
                {'id': CF_ORDER_STATUS, 'value': 1},
                {'id': CF_ORDER_TEXT,   'value': order_code},
            ]
            if order_date_ms:
                custom_fields.append({'id': CF_ORDER_DATE, 'value': order_date_ms})
            if eta_zees_ms:
                custom_fields.append({'id': CF_ETA_ZEES, 'value': eta_zees_ms})
            custom_fields.extend(date_fields)

            resp = create_clickup_task({
                'name': task_name,
                'description': '\n'.join(desc_lines),
                'status': 'to do',
                'custom_fields': custom_fields,
            })
            if resp.status_code in (200, 201):
                results.append({'gear': gear_type, 'task_id': resp.json().get('id'), 'action': 'created'})
            else:
                errors.append(f"{gear_type}: {resp.status_code} {resp.text[:100]}")

    # Persist task IDs to local registry
    all_gear_tasks = dict(saved_gear_tasks)
    all_gear_tasks.update({r['gear']: r['task_id'] for r in results})
    if saved_order:
        saved_order.update({
            'market': market, 'sport': sport, 'season': season,
            'order_num': order_num, 'num_for_code': num_for_code,
            'order_date': order_date,
            'eta_to_zees': eta_zees, 'order_sheet_url': order_sheet,
            'gear_tasks': all_gear_tasks,
        })
        upsert_order(saved_order)
    else:
        upsert_order({
            'id': str(uuid.uuid4()),
            'market': market, 'sport': sport, 'season': season,
            'order_num': order_num, 'num_for_code': num_for_code,
            'order_date': order_date,
            'eta_to_zees': eta_zees, 'order_sheet_url': order_sheet,
            'gear_tasks': all_gear_tasks,
            'created_at': datetime.now().isoformat(),
        })

    created = [r for r in results if r.get('action') == 'created']
    updated = [r for r in results if r.get('action') == 'updated']
    ok  = len(errors) == 0
    parts = []
    if created: parts.append(f"Created {len(created)} task(s)")
    if updated: parts.append(f"Updated {len(updated)} existing task(s)")
    if errors:  parts.append(f"Errors: {'; '.join(errors)}")

    return jsonify({'ok': ok, 'message': '. '.join(parts) or 'Done.'})


def _build_aftership_items(rows: list, meta: dict) -> list:
    """
    Build the list of dicts to send to AfterShip — one per unique tracking number.
    Raises ValueError if market can't be resolved.
    """
    sample     = rows[0]
    market     = sample['market']
    order_num  = sample['order_num']
    sport      = meta.get('sport', 'BSBL')
    order_date = meta.get('order_date', '')

    _so = find_order(market, order_num, sport)
    _nfc = _so.get('num_for_code', '') if _so else ''
    if not _nfc:
        _m2 = re.search(r'(\d+)$', order_num.strip())
        _nfc = order_num.strip() if order_num.strip().isdigit() else (_m2.group(1) if _m2 else '1')
    num_for_code = _nfc

    t2letter, t2primary, t2rows = build_shipment_map(rows)

    contact      = _ship_to_info(market)   # raises ValueError on unknown market
    clean_market = contact['market']

    items = []
    for tn, letter in t2letter.items():
        primary_gear  = t2primary[tn]
        carrier       = t2rows[tn][0].get('carrier', 'fedex')
        shipment_rows = t2rows[tn]

        base_code = build_order_code(market, sport, primary_gear, order_date, num_for_code)
        order_id  = f"{base_code}{letter}"

        gear_qtys = defaultdict(int)
        for r in shipment_rows:
            qty = _parse_contents_qty(r.get('contents', ''))
            if qty:
                gear_qtys[r['gear_type']] += qty
        gear_summary = ', '.join(
            f"{qty} {gear}{'s' if qty != 1 else ''}"
            for gear, qty in sorted(gear_qtys.items()) if qty > 0
        )
        gear_type_list = sorted({r['gear_type'] for r in shipment_rows})
        box_count = len(shipment_rows)
        contents  = gear_summary if gear_summary else ', '.join(gear_type_list)
        note = f"Order {order_id} | {box_count} box{'es' if box_count != 1 else ''} | {contents}" if contents else f"Order {order_id} | {box_count} box{'es' if box_count != 1 else ''}"

        items.append({
            'tracking_number': tn,
            'slug':            carrier,
            'title':           order_id,
            'order_id':        order_id,
            'language':        'en',
            'customer_name':   contact['name'],
            'phone':           contact.get('phone', ''),
            'email':           contact.get('email', ''),
            'role':            f"MSM {clean_market}",
            'note':            note,
            'gear_types':      gear_type_list,
            'primary_gear':    primary_gear,
            'letter':          letter,
            'market':          market,
            'order_num':       num_for_code,
            'ship_date':          meta.get('ship_date', ''),
            'order_date':         order_date,
            'dest_address':       contact.get('address', ''),
            'dest_city':          contact.get('city', ''),
            'dest_state':         contact.get('state', ''),
            'dest_zip':           contact.get('zip', ''),
            'origin_country_iso3':      'CHN',
            'destination_country_iso3': 'USA',
        })

    return items


@app.route('/preview/aftership', methods=['POST'])
def preview_aftership():
    rows, meta = load_data()
    if not rows:
        return jsonify({'ok': False, 'error': 'No shipment data loaded.'})
    # Save latest in case client sent updated data
    body = request.get_json(silent=True)
    if body:
        rows = body.get('rows', rows)
        meta = body.get('meta', meta)
        save_data(rows, meta)

    try:
        items = _build_aftership_items(rows, meta)
    except ValueError as e:
        return jsonify({'ok': False, 'error': str(e)})

    if not items:
        return jsonify({'ok': False, 'error': 'No tracking numbers found in the data.'})

    return jsonify({'ok': True, 'items': items})


@app.route('/api/aftership-csv', methods=['POST'])
def aftership_csv():
    """Build one combined AfterShip CSV from selected (order_id|tracking_number) pairs."""
    items = request.form.getlist('items')  # each value: "order_id|tracking_number"
    if not items:
        return 'No tracking numbers selected', 400

    orders_all = load_orders()
    orders_map = {o['id']: o for o in orders_all}

    HEADERS = [
        'courier','tracking_number','order_id','title','order_path',
        'origin_country','destination_country','custom_1','custom_2',
        'tracking_postal_code','tracking_ship_date','tracking_account_number',
        'tracking_key','tracking_destination_country','order_date','order_number',
        'language','shipment_type','order_promised_delivery_date','shipment_tags',
        'courier_connection_id','origin_state','origin_city','origin_postal_code',
        'origin_address','destination_state','destination_city','destination_postal_code',
        'destination_address','customer_name_1','customer_email_1','customer_role_1',
        'customer_phone_number_1','customer_language_1','customer_name_2','customer_email_2',
        'customer_role_2','customer_phone_number_2','customer_language_2','customer_name_3',
        'customer_email_3','customer_role_3','customer_phone_number_3','customer_language_3',
    ]

    def csv_cell(v):
        v = str(v) if v is not None else ''
        if ',' in v or '"' in v or '\n' in v:
            v = '"' + v.replace('"', '""') + '"'
        return v

    # Group selected items by TN — same TN may come from multiple orders
    tn_order_map: dict = {}   # tn → list of order dicts
    for item in items:
        if '|' not in item:
            continue
        order_id, tn = item.split('|', 1)
        tn = tn.strip()
        order = orders_map.get(order_id)
        if not tn or not order:
            continue
        tn_order_map.setdefault(tn, [])
        if order not in tn_order_map[tn]:
            tn_order_map[tn].append(order)

    rows = []

    for tn, tn_orders in tn_order_map.items():
        # Use the first order for contact/carrier/date (all share the same destination)
        first_order = tn_orders[0]
        market     = first_order.get('market', '')
        order_date = first_order.get('order_date', '')
        carrier    = 'fedex'

        try:
            contact = _ship_to_info(market)
        except ValueError:
            continue
        clean_market = contact['market']

        # Collect ALL manual shipment entries for this TN across ALL selected orders
        all_entries = []
        order_labels = []
        for order in tn_orders:
            entries = [x for x in order.get('manual_shipments', [])
                       if x.get('tracking_number', '').strip() == tn]
            all_entries.extend(entries)
            if entries:
                carrier = entries[0].get('carrier', 'fedex').lower()
                sport      = order.get('sport', 'BSBL')
                order_num  = order.get('num_for_code', order.get('order_num', ''))
                gear_types_o = list(dict.fromkeys(e.get('gear_type', '') for e in entries if e.get('gear_type')))
                primary_gear = gear_types_o[0] if gear_types_o else ''
                order_code_o = build_order_code(market, sport, primary_gear, order_date, order_num)
                order_labels.append(f"Order {order_code_o}: {', '.join(gear_types_o)}")

        if not all_entries:
            continue

        # Use first order's code as the CSV order_id field
        first_sport    = first_order.get('sport', 'BSBL')
        first_order_num = first_order.get('num_for_code', first_order.get('order_num', ''))
        first_entries  = [x for x in first_order.get('manual_shipments', [])
                          if x.get('tracking_number', '').strip() == tn]
        first_gear     = first_entries[0].get('gear_type', '') if first_entries else ''
        order_code     = build_order_code(market, first_sport, first_gear, order_date, first_order_num)

        # Aggregate across all orders
        box_count     = len(all_entries)
        declared_vals = [e.get('declared_value', '').strip() for e in all_entries if e.get('declared_value', '').strip()]
        note_texts    = list(dict.fromkeys(e.get('notes', '').strip() for e in all_entries if e.get('notes', '').strip()))

        # Build note: list each order's contribution
        contents_parts = ' + '.join(order_labels) if order_labels else 'Mixed'
        note = f"{box_count} box{'es' if box_count != 1 else ''} | {contents_parts}"
        if declared_vals:
            total_declared = 0.0
            for v in declared_vals:
                try:
                    total_declared += float(v.replace('$', '').replace(',', ''))
                except ValueError:
                    pass
            note += f" | Declared: ${total_declared:g}"
        if note_texts:
            note += ' | ' + ' | '.join(note_texts)

        order_date_iso = (order_date + 'T00:00:00+00:00') if order_date else ''

        row = [
            carrier, tn, order_code, order_code, '',
            'CHN', 'USA', note, '',
            '', '', '', '', '',
            order_date_iso, order_code,
            'en', 'FedEx International Connect Plus', '', '', '',
            'Shandong', 'Qingdao', '', '',
            contact.get('state', ''), contact.get('city', ''),
            contact.get('zip', ''), contact.get('address', ''),
            contact['name'], contact.get('email', ''), f"MSM {clean_market}",
            contact.get('phone', ''), 'en',
            '', '', '', '', '',
            '', '', '', '', '',
        ]
        rows.append(','.join(csv_cell(c) for c in row))

    if not rows:
        return 'No valid tracking numbers found', 400

    from datetime import date as _date
    filename = f"aftership_{_date.today().strftime('%m%d%y')}.csv"
    csv_content = '\n'.join([','.join(HEADERS)] + rows)
    return csv_content, 200, {
        'Content-Type': 'text/csv',
        'Content-Disposition': f'attachment; filename="{filename}"',
    }




# ── Tracking Dashboard ───────────────────────────────────────────────────────

def _collect_shipments(order: dict) -> list:
    """
    Build shipment list for one order, merging packing_list box data,
    manual_shipments, and any saved tracking_statuses.
    Returns one entry per unique tracking number.
    """
    packing = order.get('packing_list', [])
    saved_statuses = order.get('tracking_statuses', {})

    # Index packing_list boxes by TN for quick lookup
    tn_boxes: dict = {}
    for box in packing:
        tn = box.get('tracking_number', '').strip()
        if tn:
            tn_boxes.setdefault(tn, []).append({
                'box_num':  box.get('box_num', ''),
                'gear':     box.get('gear_type', ''),
                'contents': box.get('contents', ''),
                'weight':   box.get('weight', ''),
            })

    tn_info = {}
    for box in packing:
        tn = box.get('tracking_number', '').strip()
        if not tn:
            continue
        if tn not in tn_info:
            tn_info[tn] = {
                'carrier':    box.get('carrier', 'fedex'),
                'gear_types': {},
                'notes':      '',
            }
        gt = box.get('gear_type', '')
        tn_info[tn]['gear_types'][gt] = tn_info[tn]['gear_types'].get(gt, 0) + 1

    # Merge manual shipments (entered via update panel)
    for m in order.get('manual_shipments', []):
        tn = m.get('tracking_number', '').strip()
        if not tn:
            continue
        if tn not in tn_info:
            tn_info[tn] = {
                'carrier':    m.get('carrier', 'FedEx'),
                'gear_types': {},
                'gear_notes': {},
                'notes':      m.get('notes', ''),
                'ship_date':  m.get('ship_date', ''),
                'added_at':   m.get('added_at', ''),
            }
        gt = m.get('gear_type', '')
        if gt:
            try:
                boxes = max(1, int(m.get('box_count', 1)))
            except (TypeError, ValueError):
                boxes = 1
            tn_info[tn]['gear_types'][gt] = tn_info[tn]['gear_types'].get(gt, 0) + boxes
            tn_info[tn].setdefault('gear_notes', {})[gt] = m.get('notes', '')

    shipments = []
    for tn, info in tn_info.items():
        saved = saved_statuses.get(tn, {})
        tag   = saved.get('tag', 'Pending')
        label, color = _tag_info(tag)
        gear_notes = info.get('gear_notes', {})
        gear_summary = [
            {'gear': g, 'boxes': c, 'notes': gear_notes.get(g, '')}
            for g, c in sorted(info['gear_types'].items())
        ]
        # Build display groups: gear types with the same non-empty notes share a line;
        # gear types with no/different notes each appear on their own line.
        _noted: dict = {}
        gear_groups = []
        _solo = []
        for g in gear_summary:
            n = g['notes'].strip()
            if n:
                if n not in _noted:
                    _noted[n] = {'gears': [], 'boxes': g['boxes']}
                _noted[n]['gears'].append(g['gear'])
            else:
                _solo.append(g)
        for grp in _noted.values():
            bw = 'box' if grp['boxes'] == 1 else 'boxes'
            gear_groups.append(f"{', '.join(sorted(grp['gears']))} × {grp['boxes']} {bw}")
        for g in sorted(_solo, key=lambda x: x['gear']):
            bw = 'box' if g['boxes'] == 1 else 'boxes'
            gear_groups.append(f"{g['gear']} × {g['boxes']} {bw}")
        shipments.append({
            'tracking_number': tn,
            'carrier':         info['carrier'],
            'gear_summary':    gear_summary,   # [{gear, boxes, notes}, ...]
            'gear_groups':     gear_groups,    # grouped display strings for header
            'packing_boxes':   sorted(tn_boxes.get(tn, []),
                                      key=lambda b: (str(b['gear']), int(b['box_num'])
                                                     if str(b['box_num']).isdigit() else 0)),
            'tag':             tag,
            'status_label':    label,
            'status_color':    color,
            'notes':           saved.get('notes', ''),
            'updated_at':      saved.get('updated_at', ''),
            'delivered_at':    saved.get('delivered_at', ''),
        })

    return shipments


def _build_tracking_payload():
    orders = load_orders()
    # Most recently shipped order first
    orders_sorted = sorted(orders, key=lambda o: o.get('shipped_at', ''), reverse=True)
    result_orders = []
    for order in orders_sorted:
        shipments = _collect_shipments(order)
        # Most recently updated shipment first; fall back to shipped_at on order
        shipments.sort(key=lambda s: s.get('updated_at', '') or order.get('shipped_at', ''), reverse=True)
        result_orders.append({
            'id':          order.get('id'),
            'market':      order.get('market', ''),
            'sport':       order.get('sport', ''),
            'season':      order.get('season', ''),
            'order_num':   order.get('num_for_code', order.get('order_num', '')),
            'order_date':  order.get('order_date', ''),
            'shipped_at':  order.get('shipped_at', ''),
            'gear_tasks':  order.get('gear_tasks', {}),
            'shipments':   shipments,
        })
    return {
        'ok':              True,
        'orders':          result_orders,
        'statuses':        [{'tag': t, 'label': l, 'color': c, 'step': s} for t, l, c, s in TRACKING_STATUSES],
        'last_loaded':     _fmt_cst(),
        'last_auto_check': _last_auto_check['time'],
        'auto_check_error': _last_auto_check['error'],
        'has_fedex':       bool((os.getenv('FEDEX_API_KEY') and os.getenv('FEDEX_SECRET_KEY')) or os.getenv('TRACK17_API_KEY')),
    }


def _build_slack_update(order: dict, tn: str, tag: str, notes: str) -> str:
    label, _ = _tag_info(tag)
    emoji     = _status_emoji(tag)
    market    = order.get('market', '')
    season    = order.get('season', '')
    sport     = order.get('sport', '')
    order_num = order.get('num_for_code', order.get('order_num', ''))

    shipments = _collect_shipments(order)
    ship = next((s for s in shipments if s['tracking_number'] == tn), None)

    gear_lines = ''
    if ship:
        carrier = _carrier_emoji(ship.get('carrier', 'fedex'))
        parts   = [f"{s['gear']} × {s['boxes']} box{'es' if s['boxes'] != 1 else ''}"
                   for s in ship['gear_summary']]
        gear_lines = '\n'.join(f"  • {p}" for p in parts)
    else:
        carrier = 'FedEx'

    lines = [
        f"{emoji} *Shipment Update — {market} {season} {sport}*",
        f"Order #{order_num}",
        f"",
        f"*Tracking:* `{tn}` ({carrier})",
        f"*Contents:*",
        gear_lines,
        f"*Status:* {emoji} {label}",
    ]
    if notes:
        lines.append(f"*Notes:* {notes}")
    lines.append(f"\n_Updated {_fmt_cst()}_")
    return '\n'.join(lines)


@app.route('/tracking')
def tracking_dashboard():
    return render_template('tracking.html')


@app.route('/api/tracking-status')
def api_tracking_status():
    return jsonify(_build_tracking_payload())


@app.route('/api/check-tracking', methods=['POST'])
def api_check_tracking():
    result = check_all_fedex_statuses(notify_changes=True)
    return jsonify(result)


@app.route('/api/fetch-delivery-date', methods=['POST'])
def api_fetch_delivery_date():
    body = request.get_json() or {}
    tn   = body.get('tracking_number', '').strip()
    if not tn:
        return jsonify({'ok': False, 'error': 'No tracking number.'})
    try:
        result = _fetch_17track(tn)
    except Exception as e:
        return jsonify({'ok': False, 'error': str(e)})
    if not result:
        return jsonify({'ok': False, 'error': 'No data from carrier.'})

    delivered_at = ''
    if result.get('event_time'):
        try:
            et = result['event_time'][:10]
            datetime.strptime(et, '%Y-%m-%d')
            delivered_at = et
        except (ValueError, TypeError):
            pass

    # Also persist to local order registry
    order_id = body.get('order_id', '')
    if order_id and delivered_at and result.get('tag') == 'Delivered':
        orders = load_orders()
        order  = next((o for o in orders if o.get('id') == order_id), None)
        if order:
            statuses = order.setdefault('tracking_statuses', {})
            cur = statuses.get(tn, {})
            cur['delivered_at'] = delivered_at
            statuses[tn] = cur
            save_orders(orders)

    return jsonify({'ok': True, 'delivered_at': delivered_at,
                    'tag': result.get('tag', ''), 'event': result.get('last_event', '')})


@app.route('/api/update-tracking-status', methods=['POST'])
def update_tracking_status():
    body = request.get_json()
    if not body:
        return jsonify({'ok': False, 'error': 'No data.'})

    order_id = body.get('order_id', '')
    tn       = body.get('tracking_number', '').strip()
    tag      = body.get('tag', 'Pending')
    notes    = body.get('notes', '').strip()
    notify   = body.get('notify', False)   # whether to send Slack

    # Find and update the order
    orders = load_orders()
    order  = next((o for o in orders if o.get('id') == order_id), None)
    if not order:
        return jsonify({'ok': False, 'error': 'Order not found.'})

    statuses = order.setdefault('tracking_statuses', {})
    entry = {
        'tag':        tag,
        'label':      _tag_info(tag)[0],
        'notes':      notes,
        'updated_at': datetime.now().isoformat(),
    }
    if tag == 'Delivered':
        # Preserve existing delivered_at if already set; otherwise default to today
        existing_delivered = statuses.get(tn, {}).get('delivered_at', '')
        entry['delivered_at'] = existing_delivered or datetime.now().strftime('%Y-%m-%d')
    statuses[tn] = entry
    save_orders(orders)

    # ── Update ClickUp ────────────────────────────────────────────────────────
    cu_errors = []
    gear_tasks = order.get('gear_tasks', {})
    if gear_tasks and get_settings()['clickup_token']:
        # Find which gear types are in this shipment
        shipments = _collect_shipments(order)
        ship = next((s for s in shipments if s['tracking_number'] == tn), None)
        gears_in_shipment = [s['gear'] for s in ship['gear_summary']] if ship else []

        if tag == 'Delivered':
            actual_ms = int(datetime.now().replace(
                hour=0, minute=0, second=0, microsecond=0).timestamp() * 1000)
            cu_errors.extend(_sync_delivered_to_clickup(order, tn, actual_ms))

    # ── Notify ────────────────────────────────────────────────────────────────
    email_error = ''
    if notify:
        mac_msg = _build_mac_message(order, tn, tag, notes)
        send_mac_notification(mac_msg)

        label, _ = _tag_info(tag)
        subject  = (f"Shipment Update — {order.get('market')} {order.get('season')} "
                    f"{order.get('sport')} #{order.get('num_for_code', order.get('order_num'))} · {label}")
        email_error = send_email_notification(subject, mac_msg)

        slack_url = os.getenv('SLACK_WEBHOOK_URL', '')
        if slack_url:
            send_slack_message(_build_slack_update(order, tn, tag, notes))

    parts = ['Status saved.']
    if notify and not email_error:
        parts.append('Popup shown · email sent.')
    if notify and email_error:
        parts.append(f'Popup shown · email error: {email_error}')
    if cu_errors:
        parts.append(f'ClickUp: {"; ".join(cu_errors)}')

    return jsonify({'ok': True, 'message': ' '.join(parts)})


@app.route('/api/sync-delivered', methods=['POST'])
def api_sync_delivered():
    body     = request.get_json() or {}
    order_id = body.get('order_id', '').strip()
    tn       = body.get('tracking_number', '').strip()
    date_str = body.get('delivery_date', '').strip()   # YYYY-MM-DD

    orders = load_orders()
    order  = next((o for o in orders if o.get('id') == order_id), None)
    if not order:
        return jsonify({'ok': False, 'error': 'Order not found.'})
    if not tn:
        return jsonify({'ok': False, 'error': 'No tracking number.'})

    if date_str:
        try:
            dt = datetime.strptime(date_str, '%Y-%m-%d')
        except ValueError:
            dt = datetime.now()
    else:
        dt = datetime.now()
    actual_ms = int(dt.replace(hour=0, minute=0, second=0, microsecond=0).timestamp() * 1000)

    errs = _sync_delivered_to_clickup(order, tn, actual_ms)
    if errs:
        return jsonify({'ok': False, 'error': '; '.join(errs)})
    return jsonify({'ok': True, 'message': 'Delivered, Delivery Date [Actual], and Order Status updated in ClickUp.'})


# ── Dashboard ────────────────────────────────────────────────────────────────

_REMINDER_CF_LABELS = {
    CF_PROOFS_READY:    'Proofs Ready',
    CF_PROOFS_APPROVED: 'Proofs Approved',
    CF_SHIP_DATE:       'Ship Date',
    CF_DELIVERY_DATE:   'Delivery Date',
}


@app.route('/api/dashboard-reminders')
def api_dashboard_reminders():
    today = _now_cst().date()
    orders = load_orders()
    reminders = []

    for order in orders:
        gear_tasks = order.get('gear_tasks', {})
        if not gear_tasks:
            continue
        order_date_str = order.get('order_date', '')
        if not order_date_str:
            continue

        order_label = (
            f"{order.get('market')} {order.get('season')} "
            f"{order.get('sport')} #{order.get('num_for_code', order.get('order_num'))}"
        )

        for gear_type, task_id in gear_tasks.items():
            fields = calc_gear_dates(gear_type, order_date_str)
            if not fields:
                continue

            # Fetch live ClickUp actuals for this specific gear task
            task = _get_clickup_task(task_id)
            if not task:
                continue

            def _act(cf_id):
                val = _get_task_cf(task, cf_id)
                try:
                    return int(val) if val is not None else None
                except (TypeError, ValueError):
                    return None

            proofs_approved_act = _act(CF_PROOFS_APPROVED_ACT)
            ship_act            = _act(CF_SHIP_DATE_ACT)
            delivery_act        = _act(CF_DELIVERY_DATE_ACT)

            for f in fields:
                label = _REMINDER_CF_LABELS.get(f['id'])
                if not label:
                    continue

                milestone_date = datetime.fromtimestamp(f['value'] / 1000).date()
                if milestone_date > today:
                    continue

                # Condition per milestone type
                if label in ('Proofs Ready', 'Proofs Approved'):
                    # Only show if proofs haven't been approved yet
                    if proofs_approved_act is not None:
                        continue
                elif label == 'Ship Date':
                    # Only show if not yet shipped
                    if ship_act is not None:
                        continue
                elif label == 'Delivery Date':
                    # Only show if shipped but not yet delivered
                    if ship_act is None or delivery_act is not None:
                        continue

                reminders.append({
                    'order_id':    order.get('id'),
                    'order_label': order_label,
                    'gear_type':   gear_type,
                    'milestone':   label,
                    'date':        milestone_date.isoformat(),
                    'days_diff':   (milestone_date - today).days,
                })

    reminders.sort(key=lambda r: r['days_diff'])
    return jsonify({'ok': True, 'reminders': reminders})


@app.route('/dashboard')
def dashboard():
    return render_template('dashboard.html')


@app.route('/contents')
def contents_lookup():
    orders = load_orders()
    rows = []

    for o in sorted(orders, key=lambda x: x.get('created_at', ''), reverse=True):
        market   = o.get('market', '')
        sport    = o.get('sport', '')
        season   = o.get('season', '')
        onum     = o.get('num_for_code', o.get('order_num', ''))
        label    = f"{market} {sport} #{onum}"
        statuses = o.get('tracking_statuses', {})

        # Packing list boxes (have explicit box numbers)
        for box in o.get('packing_list', []):
            tn = box.get('tracking_number', '').strip()
            if not tn:
                continue
            tag = statuses.get(tn, {}).get('tag', 'Pending')
            status_label, _ = _tag_info(tag)
            rows.append({
                'order_label':  label,
                'market':       market,
                'sport':        sport,
                'gear':         box.get('gear_type', ''),
                'tracking':     tn,
                'box_num':      box.get('box_num', ''),
                'contents':     box.get('contents', ''),
                'weight':       box.get('weight', ''),
                'source':       'packing_list',
                'status':       status_label,
                'status_tag':   tag,
            })

        # Build set of (tn, gear) already covered by packing_list rows
        covered_by_pl = {
            (box.get('tracking_number', '').strip(), box.get('gear_type', ''))
            for box in o.get('packing_list', [])
            if box.get('tracking_number', '').strip() and box.get('gear_type', '')
        }

        # Manual shipments (box numbers not tracked individually)
        seen_tn_gear = set()
        for m in o.get('manual_shipments', []):
            tn = m.get('tracking_number', '').strip()
            gt = m.get('gear_type', '')
            key = (tn, gt)
            if not tn or key in seen_tn_gear or key in covered_by_pl:
                continue
            seen_tn_gear.add(key)
            tag = statuses.get(tn, {}).get('tag', 'Pending')
            status_label, _ = _tag_info(tag)
            boxes = m.get('box_count', 1)
            rows.append({
                'order_label':  label,
                'market':       market,
                'sport':        sport,
                'gear':         gt,
                'tracking':     tn,
                'box_num':      f"{boxes} box{'es' if boxes != 1 else ''}",
                'contents':     m.get('notes', ''),
                'weight':       m.get('weight', ''),
                'source':       'manual',
                'status':       status_label,
                'status_tag':   tag,
            })

    return render_template('contents.html', rows=rows)


# ── Public read-only routes ───────────────────────────────────────────────────

@app.route('/public/orders')
def public_orders():
    orders = load_orders()
    enriched = []
    for o in sorted(orders, key=lambda x: x.get('created_at', ''), reverse=True):
        shipments = _collect_shipments(o)
        enriched.append({**o, 'shipments': shipments})
    return render_template('public_orders.html', orders=enriched)


@app.route('/public/tracking')
def public_tracking():
    return render_template('public_tracking.html')


@app.route('/public/contents')
def public_contents():
    orders = load_orders()
    rows = []
    for o in sorted(orders, key=lambda x: x.get('created_at', ''), reverse=True):
        market   = o.get('market', '')
        sport    = o.get('sport', '')
        onum     = o.get('num_for_code', o.get('order_num', ''))
        label    = f"{market} {sport} #{onum}"
        statuses = o.get('tracking_statuses', {})
        for box in o.get('packing_list', []):
            tn = box.get('tracking_number', '').strip()
            if not tn:
                continue
            tag = statuses.get(tn, {}).get('tag', 'Pending')
            status_label, _ = _tag_info(tag)
            rows.append({
                'order_label': label, 'market': market, 'sport': sport,
                'gear': box.get('gear_type', ''), 'tracking': tn,
                'box_num': box.get('box_num', ''), 'contents': box.get('contents', ''),
                'weight': box.get('weight', ''), 'source': 'packing_list',
                'status': status_label, 'status_tag': tag,
            })
        covered_by_pl = {
            (box.get('tracking_number', '').strip(), box.get('gear_type', ''))
            for box in o.get('packing_list', [])
            if box.get('tracking_number', '').strip() and box.get('gear_type', '')
        }
        seen_tn_gear = set()
        for m in o.get('manual_shipments', []):
            tn = m.get('tracking_number', '').strip()
            gt = m.get('gear_type', '')
            key = (tn, gt)
            if not tn or key in seen_tn_gear or key in covered_by_pl:
                continue
            seen_tn_gear.add(key)
            tag = statuses.get(tn, {}).get('tag', 'Pending')
            status_label, _ = _tag_info(tag)
            boxes = m.get('box_count', 1)
            rows.append({
                'order_label': label, 'market': market, 'sport': sport,
                'gear': gt, 'tracking': tn,
                'box_num': f"{boxes} box{'es' if boxes != 1 else ''}",
                'contents': m.get('notes', ''), 'weight': m.get('weight', ''),
                'source': 'manual', 'status': status_label, 'status_tag': tag,
            })
    return render_template('public_contents.html', rows=rows)


@app.route('/public/dashboard')
def public_dashboard():
    return render_template('public_dashboard.html')


@app.route('/public/calendar')
def public_calendar():
    return render_template('public_calendar.html')


@app.route('/api/calendar-events')
def api_calendar_events():
    """Proxy the activity calendar's event feed so the public site has one origin."""
    try:
        r = requests.get('http://localhost:5004/api/events', timeout=10)
        r.raise_for_status()
        return jsonify(r.json())
    except Exception as e:
        return jsonify({'error': str(e)}), 502


@app.route('/api/parse-packing-for-tn', methods=['POST'])
def api_parse_packing_for_tn():
    """Upload a packing list Excel, return rows (optionally filtered to one TN)."""
    file = request.files.get('file')
    tracking = request.form.get('tracking_number', '').strip()
    if not file:
        return jsonify({'ok': False, 'error': 'No file provided'})
    with tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False) as tmp:
        file.save(tmp.name)
        tmpname = tmp.name
    try:
        rows = parse_packing_list(tmpname)
    except Exception as e:
        os.unlink(tmpname)
        return jsonify({'ok': False, 'error': str(e)})
    os.unlink(tmpname)
    if tracking:
        rows = [r for r in rows if r.get('tracking_number') == tracking]
    return jsonify({'ok': True, 'rows': rows, 'total_all': len(rows)})


@app.route('/api/save-shipment-boxes', methods=['POST'])
def api_save_shipment_boxes():
    """Save per-box packing_list entries for a specific order + tracking number."""
    body = request.get_json()
    order_id = body.get('order_id', '')
    tracking = body.get('tracking_number', '')
    boxes = body.get('boxes', [])

    orders = load_orders()
    order = next((o for o in orders if o['id'] == order_id), None)
    if not order:
        return jsonify({'ok': False, 'error': 'Order not found'})

    order.setdefault('packing_list', [])
    # Remove existing packing_list entries for this TN so we replace them
    order['packing_list'] = [r for r in order['packing_list']
                              if r.get('tracking_number') != tracking]
    for b in boxes:
        order['packing_list'].append({
            'id':              str(uuid.uuid4()),
            'market':          order.get('market', ''),
            'order_num':       order.get('num_for_code', order.get('order_num', '')),
            'gear_type':       b.get('gear_type', ''),
            'box_num':         b.get('box_num', 1),
            'contents':        b.get('contents', ''),
            'weight':          b.get('weight', ''),
            'declared_value':  b.get('declared_value', ''),
            'tracking_number': tracking,
            'carrier':         'fedex',
        })

    save_orders(orders)
    return jsonify({'ok': True, 'saved': len(boxes)})


@app.route('/api/order-sheet-teams')
def api_order_sheet_teams():
    """Fetch team names from an order's Google Sheet for a given gear type."""
    order_id  = request.args.get('order_id', '')
    gear_type = request.args.get('gear_type', '')

    orders = load_orders()
    order = next((o for o in orders if o['id'] == order_id), None)
    if not order:
        return jsonify({'ok': False, 'error': 'Order not found'})

    sheet_url = order.get('order_sheet_url', '')
    if not sheet_url:
        return jsonify({'ok': False, 'error': 'No order sheet URL saved for this order'})

    m = re.search(r'/spreadsheets/d/([a-zA-Z0-9_-]+)', sheet_url)
    if not m:
        return jsonify({'ok': False, 'error': 'Invalid sheet URL'})
    sheet_id = m.group(1)

    gkey = os.getenv('GOOGLE_API_KEY', '')
    if not gkey:
        return jsonify({'ok': False, 'error': 'GOOGLE_API_KEY not set in .env'})

    # Get all tab names + GIDs via Sheets API
    try:
        tab_resp = requests.get(
            f'https://sheets.googleapis.com/v4/spreadsheets/{sheet_id}',
            params={'key': gkey, 'fields': 'sheets.properties'},
            timeout=15,
        )
        tab_resp.raise_for_status()
        all_tabs = [{'name': s['properties']['title'],
                     'gid':  str(s['properties']['sheetId'])}
                    for s in tab_resp.json().get('sheets', [])]
    except Exception as e:
        return jsonify({'ok': False, 'error': f'Could not list sheet tabs: {e}'})

    # Keywords per gear type
    GEAR_KEYWORDS = {
        'Jersey':      ['jersey', 'jerseys'],
        'Hoodie':      ['hoodie', 'hoodies'],
        'Tee':         ['tourney', 'tee', 'tees'],
        'Cap':         ['cap', 'caps'],
        'Windbreaker': ['windbreaker', 'jacket'],
        'Cage Jacket': ['cage jacket', 'jacket'],
        'Ring':        ['ring', 'rings', 'award'],
        'Hypechain':   ['hype chain', 'hypechain'],
        'Name Banner': ['name banner', 'banner'],
        'Park Banner': ['park banner', 'banner'],
    }
    keywords = GEAR_KEYWORDS.get(gear_type, [gear_type.lower()])

    # Filter tabs by gear keyword
    matching = [t for t in all_tabs
                if any(kw in t['name'].lower() for kw in keywords)]
    if not matching:
        return jsonify({'ok': False, 'error': f'No tab found for {gear_type} in this sheet'})

    # Pick best tab: prefer one that also matches order_num or order_num as word
    order_num   = str(order.get('num_for_code', order.get('order_num', ''))).strip()
    raw_order_num = str(order.get('order_num', '')).strip()  # may be "Freshman"

    def tab_score(tab):
        name = tab['name'].lower()
        score = 0
        # Raw order_num word match (e.g. "freshman" in "FRESHMAN - Jerseys Order")
        if raw_order_num and not raw_order_num.isdigit() and raw_order_num.lower() in name:
            score += 10
        # Numeric order match: "# 1", "# 2"
        if order_num.isdigit() and f'# {order_num}' in name:
            score += 5
        return score

    matching.sort(key=tab_score, reverse=True)
    best = matching[0]

    # Read tab by GID
    try:
        csv_resp = requests.get(
            f'https://docs.google.com/spreadsheets/d/{sheet_id}/export',
            params={'format': 'csv', 'gid': best['gid']},
            timeout=20,
        )
        csv_resp.raise_for_status()
    except Exception as e:
        return jsonify({'ok': False, 'error': f'Could not read tab "{best["name"]}": {e}'})

    rows = list(csv.reader(StringIO(csv_resp.text)))

    # Find header row by looking for size codes (YM, YL, YXL, AS, etc.)
    SIZE_CODES = {'ym', 'yl', 'yxl', 'ys', 'as', 'am', 'al', 'axl', 'xs', 's', 'm', 'l', 'xl'}
    header_idx = None
    for i, row in enumerate(rows[:30]):
        cells = [c.strip().lower() for c in row]
        if sum(1 for c in cells if c in SIZE_CODES) >= 2:
            header_idx = i
            break

    if header_idx is None:
        return jsonify({'ok': False, 'error': f'Could not find size header in "{best["name"]}"'})

    headers = rows[header_idx]
    # Prefer the TEAM NAME column (col O); fall back to MIDDLE SCHOOL col (col B, index 1)
    team_col = next(
        (i for i, h in enumerate(headers) if 'TEAM NAME' in h.upper()),
        1  # fallback: col B
    )

    STOP = {'total', 'grand total', 'thank you', '', 'middle school', 'hs/ms', 'team'}
    teams = []
    for row in rows[header_idx + 1:]:
        if len(row) <= team_col:
            continue
        name = row[team_col].strip()
        if not name or name.lower() in STOP:
            continue
        has_qty = any(c.strip().isdigit() and int(c.strip()) > 0
                      for c in row[1:team_col] if c.strip())
        if has_qty:
            teams.append(name)

    if teams:
        return jsonify({'ok': True, 'teams': teams, 'tab': best['name']})
    return jsonify({'ok': False, 'error': f'No teams with quantities in "{best["name"]}"'})


@app.route('/api/dashboard-data')
def api_dashboard_data():
    orders = load_orders()
    result = []

    for order in orders:
        order_date_str = order.get('order_date', '')
        if not order_date_str:
            continue
        try:
            order_date = datetime.strptime(order_date_str, '%Y-%m-%d')
        except ValueError:
            continue

        gear_tasks = order.get('gear_tasks', {})
        if not gear_tasks:
            continue

        first_task_id = next(iter(gear_tasks.values()))
        task = _get_clickup_task(first_task_id)
        if not task:
            continue

        def _ms(cf_id):
            val = _get_task_cf(task, cf_id)
            try:
                return int(val) if val is not None else None
            except (TypeError, ValueError):
                return None

        def _dt(ms):
            try:
                return datetime.fromtimestamp(ms / 1000) if ms is not None else None
            except Exception:
                return None

        def _wdays(ms):
            dt = _dt(ms)
            return count_working_days(order_date, dt) if dt else None

        def _late(act_ms, est_ms):
            a, e = _wdays(act_ms), _wdays(est_ms)
            if a is None or e is None:
                return None
            return a > e

        proofs_act_ms   = _ms(CF_PROOFS_APPROVED_ACT)
        ship_act_ms     = _ms(CF_SHIP_DATE_ACT)
        delivery_act_ms = _ms(CF_DELIVERY_DATE_ACT)
        proofs_est_ms   = _ms(CF_PROOFS_APPROVED)
        ship_est_ms     = _ms(CF_SHIP_DATE)
        delivery_est_ms = _ms(CF_DELIVERY_DATE)

        phase_raw = _get_task_cf(task, CF_PHASE)
        try:
            current_phase = int(phase_raw) if phase_raw is not None else None
        except (TypeError, ValueError):
            current_phase = None

        # Fallback: use local manual_shipments when ClickUp fields are missing
        manual = order.get('manual_shipments', [])
        if manual:
            if ship_act_ms is None:
                dates = sorted(m['ship_date'] for m in manual if m.get('ship_date'))
                if dates:
                    try:
                        ship_act_ms = int(datetime.strptime(dates[0], '%Y-%m-%d').timestamp() * 1000)
                    except ValueError:
                        pass
            if current_phase is None or current_phase < 3:
                current_phase = 3

        if current_phase == 1:
            is_delayed = _late(proofs_act_ms, proofs_est_ms)
        elif current_phase == 3:
            is_delayed = _late(ship_act_ms, ship_est_ms)
        elif current_phase in (4, 5):
            is_delayed = _late(delivery_act_ms, delivery_est_ms)
        else:
            is_delayed = None

        # Count receiving errors across ALL gear tasks (not just the first)
        # Denominator = all gear types ordered; numerator = those with errors
        error_count    = 0
        received_count = len(gear_tasks)   # every gear type ordered counts
        for gear_task_id in gear_tasks.values():
            gt = task if gear_task_id == first_task_id else _get_clickup_task(gear_task_id)
            if not gt:
                continue
            r = _get_task_cf(gt, CF_RECEIVING)
            if r is not None:
                try:
                    ri = int(r) if not isinstance(r, dict) else int(r.get('orderindex', 3))
                    if ri != 3:
                        error_count += 1
                except (TypeError, ValueError):
                    pass

        sport_year = order_date.year if order_date.month >= 4 else order_date.year - 1

        order_status_raw = _get_task_cf(task, CF_ORDER_STATUS)
        try:
            order_status = int(order_status_raw) if not isinstance(order_status_raw, dict) else int(order_status_raw.get('orderindex', -1))
        except (TypeError, ValueError):
            order_status = None

        result.append({
            'id':               order.get('id'),
            'label':            f"{order.get('market')} {order.get('season')} {order.get('sport')} #{order.get('num_for_code', order.get('order_num'))}",
            'market':           order.get('market', ''),
            'season':           order.get('season', ''),
            'sport':            order.get('sport', ''),
            'year':             sport_year,
            'order_date':       order_date_str,
            'days_to_proofs':   _wdays(proofs_act_ms),
            'days_to_ship':     _wdays(ship_act_ms),
            'days_to_delivery': _wdays(delivery_act_ms),
            'is_delayed':       is_delayed,
            'error_count':      error_count,
            'received_count':   received_count,
            'current_phase':    current_phase,
            'order_status':     order_status,
        })

    return jsonify({'ok': True, 'orders': result})


# ── Settings ─────────────────────────────────────────────────────────────────

@app.route('/settings', methods=['GET', 'POST'])
def settings_page():
    env_path = Path(__file__).parent / '.env'
    saved = False

    if request.method == 'POST':
        lines = []
        for key in ('CLICKUP_TOKEN', 'CLICKUP_LIST_ID', 'SLACK_WEBHOOK_URL', 'EMAIL_APP_PASSWORD', 'TRACK17_API_KEY', 'FEDEX_API_KEY', 'FEDEX_SECRET_KEY'):
            val = request.form.get(key, '').strip()
            lines.append(f'{key}={val}')
        env_path.write_text('\n'.join(lines) + '\n')
        load_dotenv(str(env_path), override=True)
        saved = True

    current = get_settings()
    return render_template('settings.html', settings=current, saved=saved)


TOURNEY_LIST_ID = '901109673703'


@app.route('/tourney.ics')
def tourney_ical():
    token = os.getenv('CLICKUP_TOKEN', '')
    headers = {'Authorization': token}
    tasks = []
    page = 0
    while True:
        resp = requests.get(
            f'https://api.clickup.com/api/v2/list/{TOURNEY_LIST_ID}/task',
            headers=headers,
            params={'page': page, 'include_closed': 'true'}
        )
        data = resp.json()
        batch = data.get('tasks', [])
        tasks.extend(batch)
        if data.get('last_page') or not batch:
            break
        page += 1

    lines = [
        'BEGIN:VCALENDAR',
        'VERSION:2.0',
        'PRODID:-//MSM Sports//Tourneys//EN',
        'CALSCALE:GREGORIAN',
        'METHOD:PUBLISH',
        'X-WR-CALNAME:MSM Tourneys',
    ]
    now_stamp = datetime.utcnow().strftime('%Y%m%dT%H%M%SZ')
    for task in tasks:
        due_ms = task.get('due_date')
        if not due_ms:
            continue
        due_dt = datetime.utcfromtimestamp(int(due_ms) / 1000)
        dtstart = due_dt.strftime('%Y%m%d')
        dtend = (due_dt + timedelta(days=1)).strftime('%Y%m%d')
        summary = task.get('name', '').replace('\n', ' ').replace('\r', '')
        uid = f"{task['id']}@clickup.msm"
        url = task.get('url', '')
        lines += [
            'BEGIN:VEVENT',
            f'UID:{uid}',
            f'DTSTART;VALUE=DATE:{dtstart}',
            f'DTEND;VALUE=DATE:{dtend}',
            f'SUMMARY:{summary}',
            f'URL:{url}',
            f'DTSTAMP:{now_stamp}',
            'END:VEVENT',
        ]
    lines.append('END:VCALENDAR')

    ical_text = '\r\n'.join(lines) + '\r\n'
    return ical_text, 200, {
        'Content-Type': 'text/calendar; charset=utf-8',
        'Content-Disposition': 'inline; filename="tourney.ics"',
    }


if __name__ == '__main__':
    app.run(debug=True, port=5003)
