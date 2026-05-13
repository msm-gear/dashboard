from flask import Flask, render_template, request, session, jsonify, redirect, url_for
import openpyxl
import re
import difflib
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
CF_RECEIVING       = '3ca705c5-6fa3-4490-a804-c3bb1844c0d6'  # Receiving Condition (always blank)


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

    # Deduplicate using fuzzy-normalized market key; prefer rows with tracking numbers
    seen = {}
    for r in raw:
        key = (_market_dedup_key(r['market']), r['order_num'], r['gear_type'], r['box_num'])
        if key not in seen or (r['tracking_number'] and not seen[key]['tracking_number']):
            seen[key] = r

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
    """Format a datetime in CST as 'May 14 at 3:45 PM CST'."""
    d = (dt or _now_cst()).astimezone(_CST)
    hour = d.strftime('%I').lstrip('0') or '12'
    return d.strftime(f'%b {d.day} at {hour}:%M %p CST')


def get_settings():
    return {
        'clickup_token':      os.getenv('CLICKUP_TOKEN', ''),
        'clickup_list_id':    os.getenv('CLICKUP_LIST_ID', ''),
        'slack_webhook_url':  os.getenv('SLACK_WEBHOOK_URL', ''),
        'email_app_password': os.getenv('EMAIL_APP_PASSWORD', ''),
        'track17_api_key':    os.getenv('TRACK17_API_KEY', ''),
    }


# ── Auto-Tracking via 17track.net ─────────────────────────────────────────────
# Free account at 17track.net (email only, no credit card).
# API key goes in Settings → 17track API Key.

_last_auto_check = {'time': None, 'error': None}

_17TRACK_REGISTERED = set()   # tracking numbers already registered this session


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

    event   = latest_ev.get('description', '') or ''
    loc     = latest_ev.get('location', '') or ''
    country = (latest_ev.get('address', {}) or {}).get('country', '') or ''

    stage = _17track_tag_to_stage(tag17, event, loc, country)
    return {'tag': stage, 'last_event': event, 'location': loc}


def check_all_fedex_statuses(notify_changes: bool = False) -> dict:
    """
    Poll 17track for all tracked shipments. Update local statuses when changed.
    Auto-sends popup + email for any status that moves forward.
    """
    global _last_auto_check

    if not os.getenv('TRACK17_API_KEY'):
        return {'ok': False, 'error': '17track API key not configured. Add it in Settings.'}

    orders = load_orders()

    all_tns = {}
    for order in orders:
        for box in order.get('packing_list', []):
            tn = box.get('tracking_number', '').strip()
            if tn:
                all_tns[tn] = order

    if not all_tns:
        return {'ok': True, 'checked': 0, 'updated': 0}

    results = {}
    errors  = []
    for tn in all_tns:
        try:
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
        for box in order.get('packing_list', []):
            tn = box.get('tracking_number', '').strip()
            if not tn or tn not in results:
                continue
            new     = results[tn]
            cur     = statuses.get(tn, {})
            new_tag = new['tag']
            if new_tag == cur.get('tag'):
                continue

            note = new['last_event']
            if new['location']:
                note += f" ({new['location']})"
            statuses[tn] = {
                'tag':        new_tag,
                'label':      _tag_info(new_tag)[0],
                'notes':      note,
                'updated_at': _now_cst().isoformat(),
                'auto':       True,
            }
            updated += 1

            if notify_changes:
                msg  = _build_mac_message(order, tn, new_tag, note)
                label, _ = _tag_info(new_tag)
                subj = (f"Shipment Update — {order.get('market')} {order.get('season')} "
                        f"{order.get('sport')} #{order.get('order_num')} · {label}")
                send_mac_notification(msg)
                send_email_notification(subj, msg)

    save_orders(orders)
    _last_auto_check = {'time': _fmt_cst(), 'error': None if not errors else '; '.join(errors)}
    return {'ok': True, 'checked': len(all_tns), 'updated': updated, 'errors': errors}


def _start_fedex_poller():
    """Background thread: check 17track every 4 hours when API key is set."""
    def loop():
        time.sleep(60)
        while True:
            if os.getenv('TRACK17_API_KEY'):
                try:
                    check_all_fedex_statuses(notify_changes=True)
                except Exception:
                    pass
            time.sleep(4 * 3600)

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
    order_num = order.get('order_num', '')

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
        for g in ship['gear_summary']:
            box_word = 'box' if g['boxes'] == 1 else 'boxes'
            lines.append(f"  • {g['gear']} × {g['boxes']} {box_word}")

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


def load_orders() -> list:
    if not ORDER_REGISTRY.exists():
        return []
    return json.loads(ORDER_REGISTRY.read_text())


def save_orders(orders: list):
    ORDER_REGISTRY.write_text(json.dumps(orders, indent=2))


def find_order(market: str, order_num: str):
    """Find a saved order by market name and order number."""
    m = market.lower().strip()
    for o in load_orders():
        if o['market'].lower().strip() == m and str(o['order_num']) == str(order_num):
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
    return session_id


def load_data() -> tuple:
    session_id = session.get('data_id')
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
    return requests.post(
        f"https://api.clickup.com/api/v2/list/{list_id}/task",
        headers=clickup_headers(), json=payload, timeout=10,
    )


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



# ── Routes ──────────────────────────────────────────────────────────────────

@app.route('/')
def index():
    settings = get_settings()
    missing = [k for k, v in settings.items() if not v]
    return render_template('index.html', missing_settings=missing)


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
                           tracking_letters=tracking_letters)


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

    SUMMARY_GEAR_MAP = {
        'jerseys': 'Jersey', 'jersey': 'Jersey',
        'hoodies': 'Hoodie', 'hoodie': 'Hoodie',
        'tees': 'Tee', 'tee': 'Tee',
        'jackets': 'Cage Jacket', 'jacket': 'Cage Jacket', 'cage jackets': 'Cage Jacket',
        'caps': 'Cap', 'cap': 'Cap',
        'rings - champs': 'Ring', 'rings - finalist': 'Ring', 'rings': 'Ring', 'ring': 'Ring',
        'hype chains': 'Hypechain', 'hype chain': 'Hypechain',
        'name banner': 'Name Banner', 'name banners': 'Name Banner',
        'park banner pack': 'Park Banner', 'park banner': 'Park Banner',
        'windbreakers': 'Windbreaker', 'windbreaker': 'Windbreaker',
        'hybrid pullover': 'Hybrid Pullover', 'hybrid pullovers': 'Hybrid Pullover',
    }

    STOP_WORDS = {'total', 'thank you', 'grand total', ''}

    def parse_columns_from_csv(text):
        """
        Layout 1: Order Home Page tab.
        - Gear in col B (idx 1), preceded by 'ORDER SUMMARY' row.
        - Header row has col B = 'Gear' and cols E+ (idx 4+) = order group headers
          e.g. ['BSBL 1', 'FRESHMAN', 'BSBL 2', 'NEED COLORWAYS'].
        Returns list of {'header': str, 'gear_types': [str, ...]} per column,
        or None if layout not found.
        """
        rows = list(csv.reader(StringIO(text)))
        in_summary = False
        col_headers = []
        col_gear = {}

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

            for i, header in enumerate(col_headers):
                col_idx = 4 + i
                qty = _safe_float(row[col_idx]) if col_idx < len(row) else 0
                if qty > 0 and gear_type not in col_gear[header]:
                    col_gear[header].append(gear_type)

        if col_headers and any(col_gear[h] for h in col_headers):
            return [{'header': h, 'gear_types': col_gear[h]}
                    for h in col_headers if col_gear[h]]
        return None

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

            columns_found = parse_columns_from_csv(resp.text)
            if columns_found:
                break
            gear_found = parse_gear_fallback(resp.text)
            if gear_found:
                break
        except Exception as e:
            fetch_error = str(e)
            continue

    if columns_found:
        return jsonify({'ok': True, 'columns': columns_found})

    if gear_found:
        return jsonify({'ok': True, 'gear_types': gear_found})

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

        task_name  = build_task_name(market, season, sport, gear_type, order_num)
        order_code = build_order_code(market, sport, gear_type, order_date, order_num)
        gear_task_names[gear_type]  = task_name
        gear_order_codes[gear_type] = order_code

        estimated_ship_ms = _get_estimated_ship_ms(gear_type, order_date)
        order_status      = _ship_order_status(
            actual_ship_ms if has_tracking else None, estimated_ship_ms)
        gear_order_status[gear_type] = order_status

        existing_task_id = saved_gear_tasks.get(gear_type)

        if existing_task_id:
            set_clickup_field(existing_task_id, CF_PHASE,        3)
            set_clickup_field(existing_task_id, CF_ORDER_STATUS,  order_status)
            set_clickup_field(existing_task_id, CF_TRACKING,      tracking_str)
            set_clickup_field(existing_task_id, CF_RECEIVING,     None)
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
                {'id': CF_PHASE,        'value': 3},
                {'id': CF_ORDER_STATUS, 'value': order_status},
                {'id': CF_ORDER_TEXT,   'value': order_code},
                {'id': CF_TRACKING,     'value': tracking_str},
                {'id': CF_RECEIVING,    'value': None},
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

        shipment_rows  = t2rows[tn]
        gear_types_in  = sorted({r['gear_type'] for r in shipment_rows})
        contents_desc  = ', '.join(gear_types_in)
        weights        = [r['weight'] for r in shipment_rows if r.get('weight')]
        values         = [r['declared_value'] for r in shipment_rows if r.get('declared_value')]

        # Name and order code mirror the parent + shipment letter appended
        base_name  = gear_task_names.get(primary_gear, f"{market} {season} {sport} {primary_gear} {order_num}")
        base_code  = gear_order_codes.get(primary_gear, order_num)
        sub_name   = f"{base_name}{letter}"       # e.g. "Dallas Summer BSBL Jersey 1A"
        sub_code   = f"{base_code}{letter}"       # e.g. "DAL-BSBL-JRS-042426-1A"

        sub_fields = [
            {'id': CF_PHASE,        'value': 3},
            {'id': CF_ORDER_STATUS, 'value': gear_order_status.get(primary_gear, 1)},
            {'id': CF_ORDER_TEXT,   'value': sub_code},
            {'id': CF_TRACKING,     'value': tn},
            {'id': CF_RECEIVING,    'value': None},
        ]
        if order_date_ms:
            sub_fields.append({'id': CF_ORDER_DATE, 'value': order_date_ms})
        if actual_ship_ms:
            sub_fields.append({'id': CF_SHIP_DATE_ACT, 'value': actual_ship_ms})
        if eta_zees_ms:
            sub_fields.append({'id': CF_ETA_ZEES, 'value': eta_zees_ms})
        sub_fields.extend(calc_gear_dates(primary_gear, order_date))

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

    saved_order = find_order(market, order_num)
    saved_gear_tasks = saved_order.get('gear_tasks', {}) if saved_order else {}

    results = []
    errors  = []

    for gear_type in gear_types:
        task_name   = build_task_name(market, season, sport, gear_type, order_num)
        order_code  = build_order_code(market, sport, gear_type, order_date, order_num)
        date_fields = calc_gear_dates(gear_type, order_date)

        existing_task_id = saved_gear_tasks.get(gear_type)

        if existing_task_id:
            # ── UPDATE existing task ──────────────────────────────────────────
            fields_to_set = [
                {'id': CF_ORDER_STATUS, 'value': 1},
                {'id': CF_ORDER_TEXT,   'value': order_code},
                {'id': CF_RECEIVING,    'value': None},
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
                {'id': CF_RECEIVING,    'value': None},
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
            'order_num': order_num, 'order_date': order_date,
            'eta_to_zees': eta_zees, 'order_sheet_url': order_sheet,
            'gear_tasks': all_gear_tasks,
        })
        upsert_order(saved_order)
    else:
        upsert_order({
            'id': str(uuid.uuid4()),
            'market': market, 'sport': sport, 'season': season,
            'order_num': order_num, 'order_date': order_date,
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

    t2letter, t2primary, t2rows = build_shipment_map(rows)

    contact      = _ship_to_info(market)   # raises ValueError on unknown market
    clean_market = contact['market']

    items = []
    for tn, letter in t2letter.items():
        primary_gear  = t2primary[tn]
        carrier       = t2rows[tn][0].get('carrier', 'fedex')
        shipment_rows = t2rows[tn]

        base_code = build_order_code(market, sport, primary_gear, order_date, order_num)
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
        note_parts = [f"{box_count} box{'es' if box_count != 1 else ''}"]
        if gear_summary:
            note_parts.append(gear_summary)
        note = ' | '.join(note_parts)

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
            'order_num':       order_num,
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





# ── Tracking Dashboard ───────────────────────────────────────────────────────

def _collect_shipments(order: dict) -> list:
    """
    Build shipment list for one order, merging packing_list box data
    with any saved tracking_statuses.
    Returns one entry per unique tracking number.
    """
    packing = order.get('packing_list', [])
    saved_statuses = order.get('tracking_statuses', {})

    tn_info = {}
    for box in packing:
        tn = box.get('tracking_number', '').strip()
        if not tn:
            continue
        if tn not in tn_info:
            tn_info[tn] = {
                'carrier':    box.get('carrier', 'fedex'),
                'gear_types': {},   # gear_type → box_count
            }
        gt = box.get('gear_type', '')
        tn_info[tn]['gear_types'][gt] = tn_info[tn]['gear_types'].get(gt, 0) + 1

    shipments = []
    for tn, info in tn_info.items():
        saved = saved_statuses.get(tn, {})
        tag   = saved.get('tag', 'Pending')
        label, color = _tag_info(tag)
        gear_summary = [
            {'gear': g, 'boxes': c}
            for g, c in sorted(info['gear_types'].items())
        ]
        shipments.append({
            'tracking_number': tn,
            'carrier':         info['carrier'],
            'gear_summary':    gear_summary,   # [{gear, boxes}, ...]
            'tag':             tag,
            'status_label':    label,
            'status_color':    color,
            'notes':           saved.get('notes', ''),
            'updated_at':      saved.get('updated_at', ''),
        })

    return shipments


def _build_tracking_payload():
    orders = load_orders()
    result_orders = []
    for order in orders:
        shipments = _collect_shipments(order)
        result_orders.append({
            'id':          order.get('id'),
            'market':      order.get('market', ''),
            'sport':       order.get('sport', ''),
            'season':      order.get('season', ''),
            'order_num':   order.get('order_num', ''),
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
        'has_fedex':       bool(os.getenv('TRACK17_API_KEY')),
    }


def _build_slack_update(order: dict, tn: str, tag: str, notes: str) -> str:
    label, _ = _tag_info(tag)
    emoji     = _status_emoji(tag)
    market    = order.get('market', '')
    season    = order.get('season', '')
    sport     = order.get('sport', '')
    order_num = order.get('order_num', '')

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
    statuses[tn] = {
        'tag':        tag,
        'label':      _tag_info(tag)[0],
        'notes':      notes,
        'updated_at': datetime.now().isoformat(),
    }
    save_orders(orders)

    # ── Update ClickUp ────────────────────────────────────────────────────────
    cu_errors = []
    gear_tasks = order.get('gear_tasks', {})
    if gear_tasks and get_settings()['clickup_token']:
        # Find which gear types are in this shipment
        shipments = _collect_shipments(order)
        ship = next((s for s in shipments if s['tracking_number'] == tn), None)
        gears_in_shipment = [s['gear'] for s in ship['gear_summary']] if ship else []

        for gear in gears_in_shipment:
            task_id = gear_tasks.get(gear)
            if not task_id:
                continue
            if tag == 'Delivered':
                # Mark task complete in ClickUp
                try:
                    r = requests.put(
                        f"https://api.clickup.com/api/v2/task/{task_id}",
                        headers=clickup_headers(),
                        json={'status': 'complete'},
                        timeout=10,
                    )
                    if r.status_code not in (200, 201):
                        cu_errors.append(f"{gear} close: {r.status_code}")
                    # Set actual delivery date to today
                    set_clickup_field(task_id, CF_DELIVERY_DATE,
                                      int(datetime.now().replace(hour=0, minute=0, second=0, microsecond=0).timestamp() * 1000))
                except requests.RequestException as e:
                    cu_errors.append(str(e))
            else:
                # Update order status description note
                try:
                    set_clickup_field(task_id, CF_PHASE, 3)   # keep as Shipped
                except requests.RequestException:
                    pass

    # ── Notify ────────────────────────────────────────────────────────────────
    email_error = ''
    if notify:
        mac_msg = _build_mac_message(order, tn, tag, notes)
        send_mac_notification(mac_msg)

        label, _ = _tag_info(tag)
        subject  = (f"Shipment Update — {order.get('market')} {order.get('season')} "
                    f"{order.get('sport')} #{order.get('order_num')} · {label}")
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


# ── Settings ─────────────────────────────────────────────────────────────────

@app.route('/settings', methods=['GET', 'POST'])
def settings_page():
    env_path = Path(__file__).parent / '.env'
    saved = False

    if request.method == 'POST':
        lines = []
        for key in ('CLICKUP_TOKEN', 'CLICKUP_LIST_ID', 'SLACK_WEBHOOK_URL', 'EMAIL_APP_PASSWORD', 'TRACK17_API_KEY'):
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
