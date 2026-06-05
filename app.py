from __future__ import annotations
from flask import Flask, render_template, request, session, jsonify, redirect, url_for, send_file
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

# Server-side sessions — keeps the cookie tiny (just a session ID)
_SESSION_DIR = Path(__file__).parent / 'data' / 'sessions'
_SESSION_DIR.mkdir(parents=True, exist_ok=True)
app.config['SESSION_TYPE']             = 'filesystem'
app.config['SESSION_FILE_DIR']         = str(_SESSION_DIR)
app.config['SESSION_PERMANENT']        = False
app.config['SESSION_USE_SIGNER']       = True
app.config['SESSION_COOKIE_SAMESITE']  = 'Lax'
app.config['SESSION_COOKIE_NAME']      = 'msm_s'  # new name — forces old bloated cookie to be ignored

from flask_session import Session as FlaskSession
FlaskSession(app)


@app.template_filter('gear_display')
def gear_display_filter(gear: str) -> str:
    return GEAR_DISPLAY.get(gear, gear)

PUBLIC_PASS = 'GearUp@MSM2026'

ORDERS_PASSES = {'GearUp2Sell26!', 'GearUp@MSM2026'}  # Zee | MSM


def pub_authed():
    return session.get('pub_authed') is True


def orders_authed():
    return session.get('orders_authed') is True

@app.after_request
def expire_old_session_cookie(response):
    """Expire the old 'session' cookie so browsers stop sending the bloated payload."""
    if request.cookies.get('session'):
        response.delete_cookie('session', path='/')
    return response


@app.before_request
def block_external_admin():
    # Cloudflare sets CF-Connecting-IP on all tunneled requests
    if not request.headers.get('CF-Connecting-IP'):
        return  # local access — allow everything
    host = request.host.split(':')[0]
    if host == 'admin.msm-hq.com':
        return  # Cloudflare Access handles auth
    # tracker.msm-hq.com — public mirror, all pages open
    if host == 'tracker.msm-hq.com':
        allowed = (
            request.path.startswith('/public/') or
            request.path == '/public/login' or
            request.path == '/public/logout' or
            request.path.startswith('/packing-doc/') or
            request.path.startswith('/api/tracking-status') or
            request.path.startswith('/api/order-sheet') or
            request.path.startswith('/api/shipment-status') or
            request.path.startswith('/api/dashboard-data') or
            request.path.startswith('/api/calendar-events') or
            request.path.startswith('/api/events') or
            request.path.startswith('/api/inventory-data') or
            request.path.startswith('/api/orders-summary') or
            request.path.startswith('/static/')
        )
        if not allowed:
            return 'Not found', 404
        return
    # orders.msm-hq.com — Order Sheet only for Zees
    if host == 'orders.msm-hq.com':
        if request.path == '/':
            return redirect('/public/order-sheet')
        allowed = (
            request.path == '/public/order-sheet' or
            request.path == '/public/orders-login' or
            request.path == '/public/orders-logout' or
            request.path.startswith('/api/order-sheet') or
            request.path.startswith('/static/')
        )
        if not allowed:
            return 'Not found', 404
        return
    return 'Not found', 404

UPLOAD_FOLDER = Path(__file__).parent / 'uploads'
DATA_FOLDER = Path(__file__).parent / 'data'
PACKING_DOCS_FOLDER = UPLOAD_FOLDER / 'packing_docs'
UPLOAD_FOLDER.mkdir(exist_ok=True)
DATA_FOLDER.mkdir(exist_ok=True)
PACKING_DOCS_FOLDER.mkdir(exist_ok=True)

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
    'Cage Jacket': 'CAG',  # Fall only
    'Windbreaker': 'WNB',  # Summer only
    'Ring': 'RNG',
    'Hypechain': 'HYC',
    'Awards': 'AWD',       # Rings & Hype Chains shipped together
    'Park Banner': 'BAN',
    'Name Banner': 'NBN',
    'Hybrid Pullover': 'HYP',
    'Drawstring Bag':  'BAG',
}

ALL_GEAR_TYPES = list(GEAR_CODES.keys())  # preserves insertion order

GEAR_DISPLAY = {
    'Drawstring Bag': 'Drawstring Bags',
}

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
        'windbreaker': 'Windbreaker', 'jacket': 'Windbreaker', 'jackets': 'Windbreaker',
        'cage jacket': 'Cage Jacket', 'cage': 'Cage Jacket',
        'ring': 'Ring', 'rings': 'Ring',
        'hypechain': 'Hypechain', 'hype chain': 'Hypechain', 'hypechainn': 'Hypechain',
        'necklace': 'Hypechain', 'necklaces': 'Hypechain',
        'park banner': 'Park Banner', 'banner': 'Park Banner',
        'parkbanner': 'Park Banner', 'parkbannner': 'Park Banner',
        'name banner': 'Name Banner', 'namebanner': 'Name Banner',
        'hybrid': 'Hybrid Pullover', 'hybrid pullover': 'Hybrid Pullover',
        'bag': 'Drawstring Bag', 'bags': 'Drawstring Bag',
        'drawstring bag': 'Drawstring Bag', 'drawstring bags': 'Drawstring Bag',
    }
    return mapping.get(raw, raw.capitalize())


# Gear keywords that can appear in box contents as "keyword-qty" (e.g. "namebanner-7")
_CONTENT_GEAR_RE = re.compile(r'\b([a-z]+)-(\d+)\b', re.IGNORECASE)
_CONTENT_GEAR_WORDS = {
    'namebanner': 'Name Banner', 'parkbanner': 'Park Banner', 'parkbannner': 'Park Banner',
    'necklace': 'Hypechain', 'necklaces': 'Hypechain',
    'hypechain': 'Hypechain', 'ring': 'Ring', 'rings': 'Ring',
}


def _content_gear_rows(row: dict) -> list:
    """Return synthetic rows for gear types mentioned inside a box's contents field."""
    contents = row.get('contents', '')
    if not contents:
        return []
    extra, seen = [], set()
    for m in _CONTENT_GEAR_RE.finditer(contents):
        word = m.group(1).lower()
        gear_type = _CONTENT_GEAR_WORDS.get(word)
        if gear_type and gear_type != row['gear_type'] and gear_type not in seen:
            seen.add(gear_type)
            extra.append({
                'id':              str(uuid.uuid4()),
                'market':          row['market'],
                'order_num':       row['order_num'],
                'gear_type':       gear_type,
                'box_num':         1,
                'contents':        m.group(0),
                'weight':          row['weight'],
                'declared_value':  row['declared_value'],
                'tracking_number': row['tracking_number'],
                'carrier':         row['carrier'],
            })
    return extra


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


def _parse_item_cols_simple(row, col_offset: int, last_tracking: str):
    """Fallback parser for simple MARKET gear[-N] format (e.g. DFW tee-1).
    Returns (list_of_entries, last_tracking).  Handles '+' combined gear
    (e.g. 'Freshman hoodie+jacket') by emitting one entry per gear type.
    """
    item = row[col_offset] if len(row) > col_offset else None
    if not item:
        return [], last_tracking
    item_str = str(item).strip()
    if not item_str or item_str.lower() in ('items', 'item'):
        return [], last_tracking

    parts = item_str.split(None, 1)  # split on first whitespace
    if len(parts) < 2:
        return [], last_tracking
    market, rest = parts[0], parts[1]

    # Split on '/' to drop secondary label (e.g. "tee/banner" → "tee")
    rest = rest.split('/')[0].strip()

    # Split on '+' to get all combined gear types (e.g. "hoodie+jacket" → ["hoodie", "jacket"])
    gear_parts = [g.strip() for g in rest.split('+') if g.strip()]

    def _col(offset):
        idx = col_offset + offset
        return row[idx] if len(row) > idx else None

    tracking_raw = _col(4)
    tracking = last_tracking
    if tracking_raw:
        t = str(tracking_raw).strip()
        if t.endswith('.0'):
            t = t[:-2]
        if t and t.lower() not in ('none', 'nan', ''):
            tracking = t

    contents      = str(_col(1)).strip() if _col(1) else ''
    weight        = str(_col(2)).strip() if _col(2) else ''
    declared_val  = str(_col(3)).strip() if _col(3) else ''

    entries = []
    for gear_part in gear_parts:
        # Extract trailing box number: "-N" (e.g. "jersey-1") or " N" (e.g. "Jersey 1")
        box_num = 1
        box_m = re.search(r'[- ](\d+)$', gear_part)
        if box_m:
            box_num  = int(box_m.group(1))
            gear_raw = gear_part[:box_m.start()].strip()
        else:
            gear_raw = gear_part

        if not gear_raw:
            continue

        entries.append({
            'id':              str(uuid.uuid4()),
            'market':          market,
            'order_num':       '1',
            'gear_type':       normalize_gear(gear_raw),
            'box_num':         box_num,
            'contents':        contents,
            'weight':          weight,
            'declared_value':  declared_val,
            'tracking_number': tracking,
            'carrier':         'fedex',
        })

    return entries, tracking


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
    last_tn = ''
    for row in ws.iter_rows(min_row=2, values_only=True):
        found = False
        for col_offset in (0, 6):          # primary cols A-E, secondary cols G-K
            entry = _parse_item_cols(row, col_offset)
            if entry:
                if entry['tracking_number']:
                    last_tn = entry['tracking_number']
                raw.append(entry)
                found = True
        # For rows the regular parser skips (no '#'), try the simple fallback
        if not found:
            new_entries, last_tn = _parse_item_cols_simple(row, 0, last_tn)
            raw.extend(new_entries)

    # Scan each row's contents for inline gear mentions (e.g. "namebanner-7", "necklace-30")
    content_extras = []
    for r in raw:
        content_extras.extend(_content_gear_rows(r))
    raw.extend(content_extras)

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
        'ups_client_id':      os.getenv('UPS_CLIENT_ID', ''),
        'ups_client_secret':  os.getenv('UPS_CLIENT_SECRET', ''),
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


def _detect_carrier(tn: str) -> str:
    """Return 'ups', 'usps', or 'fedex' based on tracking number format."""
    tn = tn.strip().upper()
    if tn.startswith('1Z'):
        return 'ups'
    if len(tn) == 22 and tn.startswith('9'):
        return 'usps'
    if tn.startswith('94') or tn.startswith('92') or tn.startswith('93'):
        return 'usps'
    return 'fedex'


# 17track carrier codes
_17TRACK_CARRIER = {'fedex': 100003, 'ups': 100002, 'usps': 100001}


def _fetch_fedex(tracking_number: str) -> dict | None:
    if _detect_carrier(tracking_number) != 'fedex':
        return None  # FedEx API only handles FedEx numbers
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


_ups_token_cache: dict = {'token': None, 'expires_at': 0}


def _get_ups_token() -> str | None:
    client_id     = os.getenv('UPS_CLIENT_ID', '')
    client_secret = os.getenv('UPS_CLIENT_SECRET', '')
    if not client_id or not client_secret:
        return None
    now = time.time()
    if _ups_token_cache['token'] and now < _ups_token_cache['expires_at']:
        return _ups_token_cache['token']
    try:
        r = requests.post(
            'https://onlinetools.ups.com/security/v1/oauth/token',
            data={'grant_type': 'client_credentials'},
            auth=(client_id, client_secret),
            headers={'Content-Type': 'application/x-www-form-urlencoded'},
            timeout=10,
        )
        r.raise_for_status()
        d = r.json()
        _ups_token_cache['token']      = d['access_token']
        _ups_token_cache['expires_at'] = now + int(d.get('expires_in', 14400)) - 60
        return _ups_token_cache['token']
    except Exception:
        return None


def _ups_activity_to_stage(status_type: str, desc: str, country: str) -> str:
    s = (status_type or '').upper()
    d = (desc or '').upper()
    if s == 'D':
        return 'Delivered'
    if s == 'X':
        return 'Exception'
    if s in ('P', 'M'):
        return 'PickedUp'
    if s == 'I':
        if 'OUT FOR DELIVERY' in d:
            return 'OutForDelivery'
        if 'CLEARED CUSTOMS' in d or 'RELEASED' in d:
            return 'ClearedCustoms'
        if 'CUSTOMS' in d or 'CLEARANCE' in d:
            return 'AtCustoms'
        if country == 'US':
            return 'InTransitDom'
        if country and country != 'US':
            return 'InTransitIntl'
        return 'InTransitDom'
    if s in ('O', 'L'):
        return 'LabelCreated'
    return 'Pending'


def _fetch_ups(tracking_number: str) -> dict | None:
    if _detect_carrier(tracking_number) != 'ups':
        return None
    token = _get_ups_token()
    if not token:
        return None
    try:
        r = requests.get(
            f'https://onlinetools.ups.com/api/track/v1/details/{tracking_number}',
            params={'locale': 'en_US', 'returnSignature': 'false'},
            headers={
                'Authorization': f'Bearer {token}',
                'transId':       tracking_number[:16],
                'transactionSrc': 'MSMTracker',
            },
            timeout=15,
        )
        if r.status_code == 404:
            return None
        r.raise_for_status()
        data = r.json()

        shipment  = (data.get('trackResponse', {})
                         .get('shipment', [{}])[0])
        package   = (shipment.get('package', [{}])[0])
        activity  = package.get('activity', [])
        latest    = activity[0] if activity else {}

        status     = latest.get('status', {})
        status_type = status.get('type', '')
        desc       = status.get('description', '')
        loc_data   = latest.get('location', {}).get('address', {})
        city       = loc_data.get('city', '')
        state      = loc_data.get('stateProvince', '')
        country    = loc_data.get('countryCode', '')
        location   = ', '.join(filter(None, [city, state]))
        event_time = latest.get('date', '') + 'T' + latest.get('time', '')

        if not status_type:
            return None

        stage = _ups_activity_to_stage(status_type, desc, country)
        return {'tag': stage, 'last_event': desc, 'location': location, 'event_time': event_time}
    except Exception:
        return None


def _fetch_17track(tracking_number: str) -> dict:
    """Returns {'tag', 'last_event', 'location'} via 17track API."""
    api_key = os.getenv('TRACK17_API_KEY', '')
    if not api_key:
        return None   # signal: not configured

    headers = {'17token': api_key, 'Content-Type': 'application/json'}

    carrier_code = _17TRACK_CARRIER.get(_detect_carrier(tracking_number), 100003)
    reg_key = (tracking_number, carrier_code)
    if reg_key not in _17TRACK_REGISTERED:
        try:
            r = requests.post(
                'https://api.17track.net/track/v2.2/register',
                json=[{'number': tracking_number, 'carrier': carrier_code}],
                headers=headers,
                timeout=10,
            )
            rejected = r.json().get('data', {}).get('rejected', [])
            if rejected:
                # Already registered on a different carrier — delete and re-register
                requests.post(
                    'https://api.17track.net/track/v2.2/deletetrack',
                    json=[{'number': tracking_number}],
                    headers=headers,
                    timeout=10,
                )
                time.sleep(1)
                requests.post(
                    'https://api.17track.net/track/v2.2/register',
                    json=[{'number': tracking_number, 'carrier': carrier_code}],
                    headers=headers,
                    timeout=10,
                )
            _17TRACK_REGISTERED.add(reg_key)
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

    use_fedex   = bool(os.getenv('FEDEX_API_KEY') and os.getenv('FEDEX_SECRET_KEY'))
    use_ups     = bool(os.getenv('UPS_CLIENT_ID') and os.getenv('UPS_CLIENT_SECRET'))
    use_17track = bool(os.getenv('TRACK17_API_KEY'))

    if not use_fedex and not use_ups and not use_17track:
        return {'ok': False, 'error': 'No tracking API configured. Add FedEx or UPS credentials in Settings.'}

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
            carrier = _detect_carrier(tn)
            if carrier == 'ups' and use_ups:
                r = _fetch_ups(tn)
            elif carrier == 'fedex' and use_fedex:
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
                'location':   new.get('location', ''),
                'updated_at': _now_cst().isoformat(),
                'auto':       True,
            }
            if new_tag == 'Delivered':
                entry['delivered_at'] = delivered_at or cur.get('delivered_at', '')
            statuses[tn] = entry
            updated += 1

            # Push Delivered phase to ClickUp when auto-detected
            if new_tag == 'Delivered' and order.get('gear_tasks') and get_settings()['clickup_token']:
                actual_ms = None
                if delivered_at:
                    try:
                        actual_ms = int(datetime.strptime(delivered_at, '%Y-%m-%d').replace(
                            hour=0, minute=0, second=0, microsecond=0).timestamp() * 1000)
                    except (ValueError, TypeError):
                        pass
                if actual_ms is None:
                    actual_ms = int(datetime.now().replace(
                        hour=0, minute=0, second=0, microsecond=0).timestamp() * 1000)
                try:
                    _sync_delivered_to_clickup(order, tn, actual_ms, all_orders=orders)
                except Exception:
                    pass

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
            has_ups     = bool(os.getenv('UPS_CLIENT_ID') and os.getenv('UPS_CLIENT_SECRET'))
            has_17track = bool(os.getenv('TRACK17_API_KEY'))
            if has_fedex or has_ups or has_17track:
                try:
                    check_all_fedex_statuses(notify_changes=True)
                except Exception:
                    pass
            time.sleep(2 * 3600 if has_fedex else 4 * 3600)

    threading.Thread(target=loop, daemon=True).start()


_start_fedex_poller()


# ── ClickUp → local sync ──────────────────────────────────────────────────────

def _parse_clickup_task_name(name: str) -> dict | None:
    """
    Parse a ClickUp task name in the format:
      {market} {season} {sport} {gear} {order_num}
    e.g. "Dallas Fall BSBL Jersey 1"  or  "South Florida Summer SOC Hoodie 2"
    Returns a dict with keys: market, season, sport, gear, order_num, or None.
    """
    name = name.strip()
    # Build sorted-longest-first lists to avoid short prefix matches
    markets_sorted = sorted(MARKETS, key=len, reverse=True)
    seasons_sorted = sorted(SEASONS, key=len, reverse=True)
    sport_codes    = [s[0] for s in SPORTS]  # ['BSBL', 'SBL', 'SOC', ...]
    # Include common aliases (longest first so "Drawstring Bag" beats "Bag")
    _gear_aliases = {'Bags': 'Drawstring Bag', 'Bag': 'Drawstring Bag',
                     'Drawstring Bags': 'Drawstring Bag'}
    gear_names     = sorted(list(ALL_GEAR_TYPES) + list(_gear_aliases.keys()), key=len, reverse=True)

    market = None
    for m in markets_sorted:
        if name.lower().startswith(m.lower()):
            market = m
            name = name[len(m):].strip()
            break
    if not market:
        return None

    season = None
    for s in seasons_sorted:
        if name.lower().startswith(s.lower()):
            season = s
            name = name[len(s):].strip()
            break
    if not season:
        return None

    sport = None
    for code in sport_codes:
        if name.upper().startswith(code.upper()):
            sport = code
            name = name[len(code):].strip()
            break
    if not sport:
        return None

    gear = None
    for g in gear_names:
        if name.lower().startswith(g.lower()):
            gear = g
            name = name[len(g):].strip()
            break
    if not gear:
        return None

    order_num = name.strip() or '1'
    gear = _gear_aliases.get(gear, gear)

    return {'market': market, 'season': season, 'sport': sport,
            'gear': gear, 'order_num': order_num}


def _sync_clickup_tasks() -> dict:
    """
    Fetch all tasks from the ClickUp order-tracking list and upsert any that
    are not already linked in orders.json.  Returns {'added': int, 'skipped': int}.
    """
    settings = get_settings()
    token    = settings.get('clickup_token', '')
    list_id  = settings.get('clickup_list_id', '')
    if not token or not list_id:
        return {'added': 0, 'skipped': 0, 'error': 'ClickUp not configured'}

    headers = {'Authorization': token}

    # Collect every task from the list (paginated)
    all_tasks = []
    page = 0
    while True:
        resp = requests.get(
            f'https://api.clickup.com/api/v2/list/{list_id}/task',
            headers=headers,
            params={'page': page, 'include_closed': 'true', 'subtasks': 'false'},
            timeout=15,
        )
        if resp.status_code != 200:
            return {'added': 0, 'skipped': 0, 'error': f'ClickUp {resp.status_code}'}
        data  = resp.json()
        batch = data.get('tasks', [])
        all_tasks.extend(batch)
        if data.get('last_page') or not batch:
            break
        page += 1

    # Fix any existing manual_shipments that have the wrong carrier stored
    orders = load_orders()
    carrier_fixed = False
    for o in orders:
        for m in o.get('manual_shipments', []):
            tn = m.get('tracking_number', '').strip()
            if not tn:
                continue
            correct = _detect_carrier(tn).upper()
            stored  = m.get('carrier', '').upper()
            if stored != correct:
                m['carrier'] = correct
                carrier_fixed = True
    if carrier_fixed:
        save_orders(orders)

    # Build a set of all task IDs already tracked or explicitly blocked
    orders = load_orders()
    tracked_ids: set = set()
    for o in orders:
        for tid in o.get('gear_tasks', {}).values():
            if tid:
                tracked_ids.add(str(tid))
        for tid in o.get('sync_blocked_task_ids', []):
            if tid:
                tracked_ids.add(str(tid))

    added = 0
    skipped = 0

    # Only consider tasks created on/after the display cutoff (April 2026)
    _cutoff_ms = 1743465600000  # 2026-04-01 00:00 UTC in ms

    for task in all_tasks:
        task_id = str(task.get('id', ''))
        if not task_id or task_id in tracked_ids:
            skipped += 1
            continue

        # Skip old tasks — only sync one-offs created from April 2026 onward
        try:
            date_created_ms = int(task.get('date_created') or 0)
        except (TypeError, ValueError):
            date_created_ms = 0
        if date_created_ms < _cutoff_ms:
            skipped += 1
            continue

        parsed = _parse_clickup_task_name(task.get('name', ''))
        if not parsed:
            skipped += 1
            continue

        market    = parsed['market']
        season    = parsed['season']
        sport     = parsed['sport']
        gear      = parsed['gear']
        order_num = parsed['order_num']

        # Pull useful custom field values from the task
        def _cf(fid):
            return _get_task_cf(task, fid)

        def _ms_to_date(ms_val):
            if not ms_val:
                return ''
            try:
                return datetime.utcfromtimestamp(int(ms_val) / 1000).strftime('%Y-%m-%d')
            except Exception:
                return ''

        order_date = _ms_to_date(_cf(CF_ORDER_DATE))
        eta_zees   = _ms_to_date(_cf(CF_ETA_ZEES))
        tracking   = (_cf(CF_TRACKING) or '').strip()
        ship_date  = _ms_to_date(_cf(CF_SHIP_DATE_ACT))

        # Pull notes/box count from the shipping comment on the task
        cu_notes = ''
        if tracking:
            try:
                cr = requests.get(
                    f'https://api.clickup.com/api/v2/task/{task_id}/comment',
                    headers=headers, timeout=10,
                )
                for c in (cr.json().get('comments', []) if cr.status_code == 200 else []):
                    text = c.get('comment_text', '')
                    if f'Tracking: {tracking}' in text:
                        for line in text.splitlines():
                            if line.startswith('Notes:'):
                                cu_notes = line[len('Notes:'):].strip()
                        break
            except Exception:
                pass

        _m2 = re.search(r'(\d+)$', order_num.strip())
        num_for_code = order_num.strip() if order_num.strip().isdigit() else (_m2.group(1) if _m2 else '1')

        # Reload inside loop so concurrent syncs don't stomp each other
        orders = load_orders()
        existing = find_order(market, order_num, sport)

        def _add_tracking_to_order(order, gear_type, tn, ship_dt, notes=''):
            """Upsert a manual_shipments entry for this gear+TN if not already there."""
            if not tn:
                return
            manual = order.setdefault('manual_shipments', [])
            already = any(
                m.get('tracking_number') == tn and m.get('gear_type') == gear_type
                for m in manual
            )
            if not already:
                # Parse box count from notes if present (e.g. "3 boxes", "2 box")
                bc_match = re.search(r'(\d+)\s*box', notes or '', re.I)
                box_count = int(bc_match.group(1)) if bc_match else 1
                manual.append({
                    'gear_type':       gear_type,
                    'tracking_number': tn,
                    'ship_date':       ship_dt,
                    'declared_value':  '',
                    'notes':           notes,
                    'box_count':       box_count,
                    'added_at':        datetime.now().isoformat(),
                    'carrier':         _detect_carrier(tn).upper(),
                })

        if existing:
            # Never modify existing orders — they're managed by the app
            skipped += 1
        else:
            # Brand-new stub order — safe to pull tracking from ClickUp
            new_order = {
                'id':              str(uuid.uuid4()),
                'market':          market,
                'sport':           sport,
                'season':          season,
                'order_num':       order_num,
                'num_for_code':    num_for_code,
                'order_date':      order_date,
                'eta_to_zees':     eta_zees,
                'order_sheet_url': '',
                'gear_tasks':      {gear: task_id},
                'created_at':      datetime.now().isoformat(),
                'synced_from_clickup': True,
            }
            _add_tracking_to_order(new_order, gear, tracking, ship_date, cu_notes)
            upsert_order(new_order)
            tracked_ids.add(task_id)
            added += 1

    return {'added': added, 'skipped': skipped}


def _start_clickup_sync():
    """Background thread: sync ClickUp tasks into local orders every 30 minutes."""
    def loop():
        time.sleep(90)  # brief startup delay so FedEx poller doesn't compete
        while True:
            try:
                _sync_clickup_tasks()
            except Exception:
                pass
            time.sleep(30 * 60)

    threading.Thread(target=loop, daemon=True).start()


_start_clickup_sync()


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

    if ship and ship.get('gear_groups'):
        lines.append("Contents:")
        for group in ship['gear_groups']:
            lines.append(f"  • {group}")

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


_DISPLAY_CUTOFF = '2026-04-01'

def _order_date_str(o: dict) -> str:
    return (o.get('order_date') or '')[:10]

def load_display_orders() -> list:
    """Load orders filtered to April 2026+. Use this for all display/page routes."""
    return [o for o in load_orders() if _order_date_str(o) >= _DISPLAY_CUTOFF]


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

_clickup_task_cache: dict = {}  # task_id -> (fetched_at, task)
_CLICKUP_CACHE_TTL = 300  # seconds (5 min)

def _get_clickup_task(task_id: str):
    import time
    now = time.time()
    if task_id in _clickup_task_cache:
        fetched_at, task = _clickup_task_cache[task_id]
        if now - fetched_at < _CLICKUP_CACHE_TTL:
            return task
    resp = requests.get(
        f"https://api.clickup.com/api/v2/task/{task_id}",
        headers=clickup_headers(), timeout=10,
    )
    task = resp.json() if resp.status_code == 200 else None
    _clickup_task_cache[task_id] = (now, task)
    return task


def _next_subtask_letter(subtasks: list) -> str:
    """Return the next unused letter suffix (A, B, C…) for a new subtask."""
    used = set()
    for sub in subtasks:
        name = sub.get('name', '')
        if name and name[-1].isupper():
            used.add(name[-1])
    for c in 'ABCDEFGHIJKLMNOPQRSTUVWXYZ':
        if c not in used:
            return c
    return 'A'


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


def _sync_delivered_to_clickup(order: dict, tn: str, actual_delivery_ms: int,
                               all_orders: list = None) -> list:
    """
    For all gears in this shipment:
      - Set CF_PHASE = 4 (Delivered), CF_DELIVERY_DATE_ACT, CF_ORDER_STATUS on parent task.
      - Find subtasks with CF_TRACKING == tn and apply same updates.
    When a gear has no task ID in this order (e.g. Remake order without Hoodie/Ring tasks),
    falls back to searching all_orders for the task ID.
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

    # Build a cross-order task ID lookup for gears missing from this order
    fallback_tasks: dict = {}
    if all_orders:
        for o in all_orders:
            if o.get('id') == order.get('id'):
                continue
            for g, tid in o.get('gear_tasks', {}).items():
                if g not in fallback_tasks:
                    fallback_tasks[g] = tid

    for gear in gears:
        task_id = gear_tasks.get(gear) or fallback_tasks.get(gear)
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

        # Update subtasks that match this TN OR have no TN and are still at Shipped phase
        # (subtasks with empty CF_TRACKING can't be matched by TN but should still be delivered)
        subtasks = _get_clickup_subtasks(task_id)
        for sub in subtasks:
            sub_tn = _get_task_cf(sub, CF_TRACKING) or ''
            sub_phase_val = _get_task_cf(sub, CF_PHASE)
            try:
                sub_phase = int(sub_phase_val) if sub_phase_val is not None else None
            except (TypeError, ValueError):
                sub_phase = None
            # Match by TN, or fall back to any Shipped subtask with no TN
            if sub_tn == tn or (not sub_tn and sub_phase == 3):
                for field_id, value in [
                    (CF_PHASE,            4),
                    (CF_DELIVERY_DATE_ACT, actual_delivery_ms),
                    (CF_ORDER_STATUS,     order_status),
                ]:
                    r = set_clickup_field(sub['id'], field_id, value)
                    if r.status_code not in (200, 201):
                        errors.append(f'Subtask {sub["id"]} CF {field_id}: {r.status_code}')

    # Also update gear tasks in OTHER orders that share this TN
    # (e.g. Tee in FSH #1 when TN 813308 is shared with Dallas #2)
    if all_orders:
        already_updated = set(gear_tasks.values())
        for other_order in all_orders:
            if other_order.get('id') == order.get('id'):
                continue
            tn_in_other = any(
                ms.get('tracking_number', '').strip() == tn
                for ms in other_order.get('manual_shipments', [])
            ) or any(
                row.get('tracking_number', '').strip() == tn
                for row in other_order.get('packing_list', [])
            )
            if not tn_in_other:
                continue
            # Collect gears in the other order that are associated with this TN
            other_shipments = _collect_shipments(other_order)
            other_ship = next((s for s in other_shipments if s['tracking_number'] == tn), None)
            if not other_ship:
                continue
            other_gears = [s['gear'] for s in other_ship['gear_summary']]
            for g in other_gears:
                tid = other_order.get('gear_tasks', {}).get(g)
                if not tid or tid in already_updated:
                    continue
                already_updated.add(tid)
                parent2 = _get_clickup_task(tid)
                est2 = None
                if parent2:
                    val2 = _get_task_cf(parent2, CF_DELIVERY_DATE)
                    try:
                        est2 = int(val2) if val2 is not None else None
                    except (TypeError, ValueError):
                        pass
                os2 = _delivery_order_status(est2, actual_delivery_ms)
                for field_id, value in [
                    (CF_PHASE,             4),
                    (CF_DELIVERY_DATE_ACT, actual_delivery_ms),
                    (CF_ORDER_STATUS,      os2),
                ]:
                    r = set_clickup_field(tid, field_id, value)
                    if r.status_code not in (200, 201):
                        errors.append(f'{other_order.get("market")} {g} CF {field_id}: {r.status_code}')

    return errors


# ── Routes ──────────────────────────────────────────────────────────────────

def _pl_meta_for_tn(tn: str, orders_all: list) -> dict:
    """Compute total weight and declared value for a TN from packing list rows across all orders.
    Deduplicates physical boxes by (order_id, box_num) to avoid double-counting.
    Returns {'weight': '12.5KG', 'declared_value': '13$'} or empty strings if not found.
    """
    def _parse_num(s):
        if not s:
            return 0.0
        try:
            return float(str(s).upper().replace('KG', '').replace('$', '').strip())
        except (ValueError, AttributeError):
            return 0.0

    seen_boxes: dict = {}  # (order_id, box_num) -> row
    for o in orders_all:
        for row in o.get('packing_list', []):
            if row.get('tracking_number', '').strip() != tn:
                continue
            key = (o['id'], str(row.get('box_num', '')))
            if key not in seen_boxes:
                seen_boxes[key] = row

    if not seen_boxes:
        return {'weight': '', 'declared_value': ''}

    total_w  = sum(_parse_num(r.get('weight', ''))        for r in seen_boxes.values())
    total_dv = sum(_parse_num(r.get('declared_value', '')) for r in seen_boxes.values())

    return {
        'weight':        f"{round(total_w, 2)}" if total_w else '',
        'declared_value': f"{int(round(total_dv))}" if total_dv else '',
    }


@app.route('/')
def index():
    settings = get_settings()
    missing = [k for k, v in settings.items() if not v]

    # Collect all manual shipments for the packing list table
    orders_all = load_orders()
    orders_dirty = False

    # Pre-compute PL meta per TN (across all orders) once
    _pl_cache: dict = {}

    def _get_pl_meta(tn):
        if tn not in _pl_cache:
            _pl_cache[tn] = _pl_meta_for_tn(tn, orders_all)
        return _pl_cache[tn]

    # Pass 1: fill missing weight/declared on every manual_shipment entry
    for o in orders_all:
        for ms in o.get('manual_shipments', []):
            tn = ms.get('tracking_number', '').strip()
            if not tn:
                continue
            w  = ms.get('weight', '') or ''
            dv = ms.get('declared_value', '') or ''
            if w and dv:
                continue  # already complete
            pl = _get_pl_meta(tn)
            if pl['weight'] and not w:
                ms['weight'] = pl['weight']
                orders_dirty = True
            if pl['declared_value'] and not dv:
                ms['declared_value'] = pl['declared_value']
                orders_dirty = True

    if orders_dirty:
        for o in orders_all:
            upsert_order(o)
        orders_all = load_orders()  # reload with saved values

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
                existing = tns_map[tn].get('gear_type', '')
                new_gt = m.get('gear_type', '')
                if new_gt and new_gt not in existing:
                    tns_map[tn] = dict(tns_map[tn])
                    tns_map[tn]['gear_type'] = ', '.join(filter(None, [existing, new_gt]))

        market = o.get('market', '')
        sport  = o.get('sport', '')
        onum   = o.get('num_for_code', o.get('order_num', ''))

        for tn, m in tns_map.items():
            def _strip_units(v):
                return str(v or '').upper().replace('KG', '').replace('$', '').strip()

            manual_rows.append({
                'order_id':       o['id'],
                'order_label':    f"{market} {sport} #{onum}",
                'tn':             tn,
                'gear_type':      m.get('gear_type', ''),
                'notes':          m.get('notes', ''),
                'weight':         _strip_units(m.get('weight', '')),
                'declared_value': _strip_units(m.get('declared_value', '')),
                'pl_only':        False,
            })

        # Also add packing-list-only TNs (no manual_shipment entry exists)
        pl_tns_in_order: dict = {}  # tn -> {gear_types, weight, declared}
        for row in o.get('packing_list', []):
            tn = row.get('tracking_number', '').strip()
            if not tn or tn in tns_map:
                continue
            if tn not in pl_tns_in_order:
                pl_tns_in_order[tn] = {'gear_types': set(), 'key': (o['id'], tn)}
            pl_tns_in_order[tn]['gear_types'].add(row.get('gear_type', ''))

        for tn, info in pl_tns_in_order.items():
            pl = _get_pl_meta(tn)
            gear_str = ', '.join(sorted(g for g in info['gear_types'] if g))
            manual_rows.append({
                'order_id':       o['id'],
                'order_label':    f"{market} {sport} #{onum}",
                'tn':             tn,
                'gear_type':      gear_str,
                'notes':          '',
                'weight':         pl.get('weight', ''),
                'declared_value': pl.get('declared_value', ''),
                'pl_only':        True,
            })

    # Collect all packing list rows for display (sorted by order then box)
    all_pl_rows = []
    for o in sorted(orders_all, key=lambda x: x.get('created_at', ''), reverse=True):
        market = o.get('market', '')
        sport  = o.get('sport', '')
        onum   = o.get('num_for_code', o.get('order_num', ''))
        label  = f"{market} {sport} #{onum}"
        for row in o.get('packing_list', []):
            def _su(v):
                return str(v or '').upper().replace('KG', '').replace('$', '').strip()
            all_pl_rows.append({
                'item':     f"{label} {row.get('gear_type','')} box {row.get('box_num','')}",
                'contents': row.get('contents', ''),
                'weight':   _su(row.get('weight', '')),
                'declared': _su(row.get('declared_value', '')),
                'tn':       row.get('tracking_number', ''),
            })

    return render_template('index.html', missing_settings=missing, manual_rows=manual_rows, pl_rows=all_pl_rows)


@app.route('/api/update-shipment-meta', methods=['POST'])
def update_shipment_meta():
    """Save weight, declared_value, and notes for a manual shipment entry."""
    data     = request.get_json(force=True)
    oid      = data.get('order_id', '').strip()
    tn       = data.get('tracking_number', '').strip()
    weight   = data.get('weight', '').strip()
    declared = data.get('declared_value', '').strip()
    notes    = data.get('notes', '')          # optional; None = don't touch
    if not oid or not tn:
        return jsonify({'error': 'order_id and tracking_number required'}), 400

    orders = load_orders()
    order = next((o for o in orders if o.get('id') == oid), None)
    if not order:
        return jsonify({'error': 'Order not found'}), 404

    updated = False
    for m in order.get('manual_shipments', []):
        if m.get('tracking_number', '').strip() == tn:
            m['weight']         = weight
            m['declared_value'] = declared
            if notes is not None:
                m['notes'] = notes
            updated = True

    if not updated:
        return jsonify({'ok': True, 'skipped': True})  # PL-only row, no manual_shipment to update

    upsert_order(order)
    return jsonify({'ok': True})


_ck_notes_cache: dict = {}  # task_id -> {ts, comments}
_CK_NOTES_TTL = 600  # 10 min


@app.route('/api/fetch-shipment-notes')
def api_fetch_shipment_notes():
    """Fetch the Notes: line from the ClickUp shipping comment for a given order+TN."""
    order_id = request.args.get('order_id', '').strip()
    tn       = request.args.get('tn', '').strip()
    if not order_id or not tn:
        return jsonify({'ok': False, 'error': 'order_id and tn required'})

    orders = load_orders()
    order  = next((o for o in orders if o.get('id') == order_id), None)
    if not order:
        return jsonify({'ok': False, 'error': 'Order not found'})

    gear_tasks = order.get('gear_tasks', {})
    if not gear_tasks:
        return jsonify({'ok': True, 'notes': ''})

    cu_token = os.getenv('CLICKUP_TOKEN', '')
    notes_found = ''

    for task_id in gear_tasks.values():
        # Check cache
        cached = _ck_notes_cache.get(task_id)
        if cached and (time.time() - cached['ts']) < _CK_NOTES_TTL:
            comments = cached['comments']
        else:
            try:
                r = requests.get(
                    f'https://api.clickup.com/api/v2/task/{task_id}/comment',
                    headers={'Authorization': cu_token}, timeout=10,
                )
                comments = r.json().get('comments', []) if r.status_code == 200 else []
            except Exception:
                comments = []
            _ck_notes_cache[task_id] = {'ts': time.time(), 'comments': comments}

        for c in comments:
            text = c.get('comment_text', '')
            if f'Tracking: {tn}' not in text:
                continue
            for line in text.splitlines():
                if line.startswith('Notes:'):
                    notes_found = line[len('Notes:'):].strip()
                    break
            if notes_found:
                break
        if notes_found:
            break

    # Persist back to manual_shipments if found and currently empty
    if notes_found:
        changed = False
        for ms in order.get('manual_shipments', []):
            if ms.get('tracking_number', '').strip() == tn and not (ms.get('notes') or ''):
                ms['notes'] = notes_found
                changed = True
        if changed:
            upsert_order(order)

    return jsonify({'ok': True, 'notes': notes_found})


_SUPPLIER_CACHE: dict = {}  # order_id → {ts, supplier}
_SUPPLIER_TTL = 600  # 10 min


def _resolve_supplier(raw: str) -> str:
    """Map a raw 'Supplier:' comment value to 'Team Green' or 'LNX'.
    If the comment explicitly names a China/TG supplier → Team Green.
    Otherwise (e.g. 'LNX') → LNX. Default when called from explicit comment is LNX."""
    if not raw:
        return 'Team Green'
    low = raw.lower()
    if any(k in low for k in ('china', 'qingdao', 'shanghai', 'shenzhen', 'guangzhou',
                               'guangdong', 'beijing', 'hangzhou', 'yiwu', 'ningbo',
                               'fujian', 'zhejiang', 'jiangsu', 'prc', 'team green')):
        return 'Team Green'
    return 'LNX'


def _fetch_supplier_from_clickup(order: dict) -> str:
    """Return 'LNX' if a ClickUp comment explicitly says Supplier: LNX (or non-China).
    Default is 'Team Green' (China) — no comment needed for China orders."""
    gear_tasks = order.get('gear_tasks', {})
    if not gear_tasks:
        return 'Team Green'   # default
    cu_token = os.getenv('CLICKUP_TOKEN', '')
    for task_id in gear_tasks.values():
        cached = _ck_notes_cache.get(f'sup_{task_id}')
        if cached and (time.time() - cached['ts']) < _SUPPLIER_TTL:
            comments = cached['comments']
        else:
            try:
                r = requests.get(
                    f'https://api.clickup.com/api/v2/task/{task_id}/comment',
                    headers={'Authorization': cu_token}, timeout=10,
                )
                comments = r.json().get('comments', []) if r.status_code == 200 else []
            except Exception:
                comments = []
            _ck_notes_cache[f'sup_{task_id}'] = {'ts': time.time(), 'comments': comments}

        for c in reversed(comments):   # latest comment wins
            text = c.get('comment_text', '').strip()
            if text.lower().startswith('supplier'):
                raw = text.split(':', 1)[-1].strip() if ':' in text else text
                return _resolve_supplier(raw)   # explicit override
    return 'Team Green'   # default: China


@app.route('/api/fetch-supplier/<order_id>')
def api_fetch_supplier(order_id):
    """Return the supplier (Team Green or LNX) from ClickUp task comments."""
    cached = _SUPPLIER_CACHE.get(order_id)
    if cached and (time.time() - cached['ts']) < _SUPPLIER_TTL:
        return jsonify({'ok': True, 'supplier': cached['supplier']})

    orders = load_orders()
    order  = next((o for o in orders if o.get('id') == order_id), None)
    if not order:
        return jsonify({'ok': False, 'error': 'Order not found'})

    supplier = order.get('supplier') or _fetch_supplier_from_clickup(order)

    # Persist if found and not already saved
    if supplier and not order.get('supplier'):
        order['supplier'] = supplier
        upsert_order(order)

    _SUPPLIER_CACHE[order_id] = {'ts': time.time(), 'supplier': supplier}
    return jsonify({'ok': True, 'supplier': supplier})


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

    all_orders = load_display_orders()
    auto_matched: list = []
    first_tn_matched_order = None  # track canonical order found via TN match
    for o in all_orders:
        matched_tns = set()
        for m in o.get('manual_shipments', []):
            tn = m.get('tracking_number', '').strip()
            if not tn or tn not in tn_to_rows or tn in matched_tns:
                continue
            matched_tns.add(tn)

            # Save per-box rows to packing_list (replace old rows for this TN)
            o.setdefault('packing_list', [])
            o['packing_list'] = [x for x in o['packing_list']
                                 if x.get('tracking_number') != tn]
            o['packing_list'].extend(tn_to_rows[tn])

            # Fill missing weight / declared_value on manual_shipment entries from pl rows
            pl_rows_for_tn = tn_to_rows[tn]
            for ms in o.get('manual_shipments', []):
                if ms.get('tracking_number', '').strip() != tn:
                    continue
                gear = ms.get('gear_type', '')
                gear_rows = [r for r in pl_rows_for_tn if r.get('gear_type') == gear]
                if not gear_rows:
                    gear_rows = pl_rows_for_tn  # fallback: any row for this TN
                if not ms.get('weight'):
                    weight_val = next((r['weight'] for r in gear_rows if r.get('weight')), '')
                    if weight_val:
                        ms['weight'] = weight_val
                if not ms.get('declared_value'):
                    dv_val = next((r['declared_value'] for r in gear_rows if r.get('declared_value')), '')
                    if dv_val:
                        ms['declared_value'] = dv_val

            if first_tn_matched_order is None:
                first_tn_matched_order = o

            auto_matched.append({
                'order': f"{o.get('market','')} {o.get('sport','')} #{o.get('num_for_code', o.get('order_num',''))}",
                'tn': tn,
                'boxes': len(tn_to_rows[tn]),
            })
    if auto_matched:
        save_orders(all_orders)
        # If market/order_num from packing list didn't match any saved order directly,
        # fill meta from the TN-matched order so the form pre-populates correctly
        if not matched_order and first_tn_matched_order:
            matched_order = first_tn_matched_order
            for key in ('sport', 'season', 'order_date', 'order_sheet_url'):
                if not meta.get(key) and first_tn_matched_order.get(key):
                    meta[key] = first_tn_matched_order[key]
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
                           tracking_letters=tracking_letters,
                           auto_matched=auto_matched)


PHASE_LABELS = ['Submitted', 'Proofs', 'Production', 'Shipped', 'Delivered', 'Counted']


@app.route('/orders')
def orders_list():
    orders = load_display_orders()
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


_orders_summary_cache: dict = {'ts': 0, 'data': None}
_ORDERS_SUMMARY_TTL = 300  # 5 minutes


@app.route('/api/orders-summary')
def api_orders_summary():
    """Return phase + dates for all orders (for the pending summary table)."""
    global _orders_summary_cache
    if _orders_summary_cache['data'] and (time.time() - _orders_summary_cache['ts']) < _ORDERS_SUMMARY_TTL:
        return jsonify(_orders_summary_cache['data'])

    def _ms_to_date(val):
        if not val:
            return None
        try:
            return datetime.fromtimestamp(int(val) / 1000).strftime('%Y-%m-%d')
        except Exception:
            return None

    orders = load_display_orders()
    result = []

    for order in orders:
        gear_tasks = order.get('gear_tasks', {})
        if not gear_tasks:
            continue

        # Fetch first task for shared date fields
        first_task = _get_clickup_task(next(iter(gear_tasks.values())))
        if not first_task:
            continue

        def _cf(cf_id):
            return _get_task_cf(first_task, cf_id)

        proofs_ready_est    = _ms_to_date(_cf(CF_PROOFS_READY))
        proofs_approved_est = _ms_to_date(_cf(CF_PROOFS_APPROVED))
        ship_est            = _ms_to_date(_cf(CF_SHIP_DATE))
        delivery_est        = _ms_to_date(_cf(CF_DELIVERY_DATE))
        proofs_act          = _ms_to_date(_cf(CF_PROOFS_APPROVED_ACT))
        ship_act            = _ms_to_date(_cf(CF_SHIP_DATE_ACT))
        delivery_act        = _ms_to_date(_cf(CF_DELIVERY_DATE_ACT))

        # Per-gear phases
        gear_phases = {}
        for gear, task_id in gear_tasks.items():
            task = first_task if task_id == next(iter(gear_tasks.values())) else _get_clickup_task(task_id)
            if task:
                val = _get_task_cf(task, CF_PHASE)
                try:
                    gear_phases[gear] = int(val) if val is not None else None
                except (TypeError, ValueError):
                    gear_phases[gear] = None
            else:
                gear_phases[gear] = None

        # Skip if all gears are delivered (phase ≥ 4)
        phases = [p for p in gear_phases.values() if p is not None]
        if phases and all(p >= 4 for p in phases):
            continue

        result.append({
            'id':          order.get('id'),
            'market':      order.get('market', ''),
            'season':      order.get('season', ''),
            'sport':       order.get('sport', ''),
            'order_num':   order.get('num_for_code', order.get('order_num', '')),
            'order_date':  order.get('order_date', ''),
            'gear_phases': gear_phases,
            'proofs_ready_est':    proofs_ready_est,
            'proofs_approved_est': proofs_approved_est,
            'ship_est':     ship_est,
            'delivery_est': delivery_est,
            'proofs_act':   proofs_act,
            'ship_act':     ship_act,
            'delivery_act': delivery_act,
        })

    payload = {'ok': True, 'orders': result, 'labels': PHASE_LABELS}
    _orders_summary_cache = {'ts': time.time(), 'data': payload}
    return jsonify(payload)


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
    service_type    = body.get('service_type', '').strip()  # FedEx service type for AfterShip
    try:
        box_count = max(1, int(body.get('box_count', 1)))
    except (TypeError, ValueError):
        box_count = 1
    # Fall back to parsing box count from notes if still at default
    if box_count == 1 and notes:
        bc_m = re.search(r'\b(\d+)\s+box(?:es)?\b', notes, re.IGNORECASE)
        if bc_m:
            box_count = max(1, int(bc_m.group(1)))

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

    _FORMULA_FIELD_IDS = {
        '08f6fefb-3605-48e4-bdbb-b5e1c765d87a',
        '43acbd5f-e47a-4230-93fd-cc7f596a731d',
        '0245e34f-f032-44c5-b016-792b4447c658',
        '9e1380de-7dbf-4c77-afdb-7de496f61e2b',
    }

    errors  = []
    updated = 0
    for gear, task_id in gear_tasks.items():
        r = set_clickup_field(task_id, CF_PHASE, 3)
        if r.status_code not in (200, 201):
            errors.append(f"{gear}: {r.status_code}"); continue

        # Fetch parent BEFORE updating CF_TRACKING so we can detect the old TN for subtask split
        parent_task     = _get_clickup_task(task_id)
        old_parent_tn   = (_get_task_cf(parent_task, CF_TRACKING) if parent_task else None) or None

        set_clickup_field(task_id, CF_TRACKING, tracking_number)
        if ship_date_ms:
            set_clickup_field(task_id, CF_SHIP_DATE_ACT, ship_date_ms)

        order_status_val = None
        if ship_date_ms and parent_task:
            try:
                est_ms = int(_get_task_cf(parent_task, CF_SHIP_DATE) or 0) or None
            except (TypeError, ValueError):
                est_ms = None
            order_status_val = _ship_order_status(ship_date_ms, est_ms)
            set_clickup_field(task_id, CF_ORDER_STATUS, order_status_val)

        requests.post(
            f"https://api.clickup.com/api/v2/task/{task_id}/comment",
            headers=clickup_headers(), json={"comment_text": comment_text}, timeout=10,
        )
        updated += 1

        subtasks = _get_clickup_subtasks(task_id)
        matched_sub = False
        for sub in subtasks:
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
            matched_sub = True

        # Existing subtasks present but none matched this TN → create a new subtask
        if subtasks and not matched_sub:
            letter      = _next_subtask_letter(subtasks)
            parent_name = (parent_task or {}).get('name', '') or \
                          f"{order.get('market','')} {order.get('season','')} {order.get('sport','')} {gear} {order.get('num_for_code', order.get('order_num',''))}"
            parent_code = (_get_task_cf(parent_task, CF_ORDER_TEXT) if parent_task else '') or ''
            sub_name    = f"{parent_name}{letter}"
            sub_code    = f"{parent_code}{letter}" if parent_code else ''

            sub_fields_map = {}
            if parent_task:
                for cf in parent_task.get('custom_fields', []):
                    val = cf.get('value')
                    if val is not None and cf['id'] not in _FORMULA_FIELD_IDS:
                        sub_fields_map[cf['id']] = val
            # Subtask-specific overrides
            sub_fields_map[CF_TRACKING] = tracking_number
            sub_fields_map[CF_PHASE]    = 3
            if ship_date_ms:
                sub_fields_map[CF_SHIP_DATE_ACT] = ship_date_ms
            if order_status_val is not None:
                sub_fields_map[CF_ORDER_STATUS] = order_status_val
            if sub_code:
                sub_fields_map[CF_ORDER_TEXT] = sub_code
            sub_fields_map.pop(CF_RECEIVING, None)  # never set receiving on new subtask

            desc = f"Gear: {gear}\nTracking: {tracking_number}"
            if declared_value:
                desc += f"\nDeclared Value: ${declared_value}"
            if notes:
                desc += f"\nNotes: {notes}"

            sub_cf_ship = [{'id': fid, 'value': val} for fid, val in sub_fields_map.items()]
            sresp = create_clickup_task({
                'name':         sub_name,
                'parent':       task_id,
                'description':  desc,
                'custom_fields': sub_cf_ship,
            })
            if sresp.status_code in (200, 201):
                new_sub_id = sresp.json().get('id', '')
                if new_sub_id:
                    for f in sub_cf_ship:
                        set_clickup_field(new_sub_id, f['id'], f['value'])
                    requests.post(
                        f"https://api.clickup.com/api/v2/task/{new_sub_id}/comment",
                        headers=clickup_headers(), json={"comment_text": comment_text}, timeout=10,
                    )
            else:
                errors.append(f"{gear} new subtask: {sresp.status_code}")

        elif not subtasks:
            # No subtasks yet — if parent had a different TN before this push, split into A + B
            if old_parent_tn and old_parent_tn != tracking_number:
                # Build comment for subtask A using the old shipment's stored data
                old_ms = next((m for m in order.get('manual_shipments', [])
                               if m.get('tracking_number') == old_parent_tn
                               and m.get('gear_type') == gear), None)
                old_comment = f"📦 Shipped\nTracking: {old_parent_tn}"
                if old_ms:
                    if old_ms.get('ship_date'):
                        old_comment += f"\nShip Date: {old_ms['ship_date']}"
                    if old_ms.get('declared_value'):
                        old_comment += f"\nDeclared Value: ${old_ms['declared_value']}"
                    if old_ms.get('notes'):
                        old_comment += f"\nNotes: {old_ms['notes']}"

                # Create subtask A (old TN, parent's existing ship fields)
                # then subtask B (new TN, new ship fields)
                for i, (sub_tn, override_ship_ms, override_status, sub_comment) in enumerate([
                    (old_parent_tn,      None,         None,              old_comment),
                    (tracking_number,   ship_date_ms,  order_status_val,  comment_text),
                ]):
                    letter      = chr(ord('A') + i)
                    parent_name = (parent_task or {}).get('name', '') or \
                                  f"{order.get('market','')} {order.get('season','')} {order.get('sport','')} {gear} {order.get('num_for_code', order.get('order_num',''))}"
                    parent_code = (_get_task_cf(parent_task, CF_ORDER_TEXT) if parent_task else '') or ''
                    sub_name    = f"{parent_name}{letter}"
                    sub_code    = f"{parent_code}{letter}" if parent_code else ''

                    sfm = {}
                    if parent_task:
                        for cf in parent_task.get('custom_fields', []):
                            val = cf.get('value')
                            if val is not None and cf['id'] not in _FORMULA_FIELD_IDS:
                                sfm[cf['id']] = val
                    sfm[CF_TRACKING] = sub_tn
                    sfm[CF_PHASE]    = 3
                    if override_ship_ms:
                        sfm[CF_SHIP_DATE_ACT] = override_ship_ms
                    if override_status is not None:
                        sfm[CF_ORDER_STATUS] = override_status
                    if sub_code:
                        sfm[CF_ORDER_TEXT] = sub_code
                    sfm.pop(CF_RECEIVING, None)

                    sfm_list = [{'id': fid, 'value': val} for fid, val in sfm.items()]
                    sresp = create_clickup_task({
                        'name':          sub_name,
                        'parent':        task_id,
                        'description':   f"Gear: {gear}\nTracking: {sub_tn}",
                        'custom_fields': sfm_list,
                    })
                    if sresp.status_code in (200, 201):
                        new_sub_id = sresp.json().get('id', '')
                        if new_sub_id:
                            for f in sfm_list:
                                set_clickup_field(new_sub_id, f['id'], f['value'])
                            requests.post(
                                f"https://api.clickup.com/api/v2/task/{new_sub_id}/comment",
                                headers=clickup_headers(), json={"comment_text": sub_comment}, timeout=10,
                            )
                    else:
                        errors.append(f"{gear} subtask {letter}: {sresp.status_code}")

    if not errors:
        # Save to order registry so shipment appears in the shipments view
        manual = order.setdefault('manual_shipments', [])
        # Replace existing entry for same gear+tracking or append new
        existing = next((m for m in manual if m['gear_type'] == gear_type
                         and m['tracking_number'] == tracking_number), None)
        detected_carrier = _detect_carrier(tracking_number)
        entry = {
            'gear_type':       gear_type,
            'tracking_number': tracking_number,
            'ship_date':       ship_date,
            'declared_value':  declared_value,
            'notes':           notes,
            'box_count':       box_count,
            'added_at':        datetime.now().isoformat(),
            'carrier':         'UPS' if detected_carrier == 'ups' else 'FedEx',
            'service_type':    service_type,
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

    SUMMARY_GEAR_MAP = {
        'jerseys': 'Jersey', 'jersey': 'Jersey',
        'hoodies': 'Hoodie', 'hoodie': 'Hoodie',
        'tees': 'Tee', 'tee': 'Tee',
        'jackets': 'Windbreaker', 'jacket': 'Windbreaker',
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


def _gear_ship_ms(gear_tns: list, tn_ship_date: dict, fallback_ms) -> int | None:
    """Return the earliest manual ship date for any TN in gear_tns, or fallback_ms."""
    dates = [tn_ship_date[tn] for tn in gear_tns if tn in tn_ship_date]
    return min(dates) if dates else fallback_ms


def _pick_canonical_market(rows: list) -> tuple:
    """
    Return (market, order_num) from the rows, preferring known markets with the
    highest numeric order_num.  Handles simple-parser artifacts like "Dallas#2"
    (market="Dallas", order_num="2") and exceptions like "Freshman" / "Remake"
    which are not in MARKET_CODES and won't override a valid numeric candidate.
    """
    best_m, best_n, best_nval = None, '1', -1
    for r in rows:
        m, n = r['market'].strip(), r['order_num']
        hm = re.match(r'^(.+?)#(\d+)$', m)
        if hm:
            m, n = hm.group(1).strip(), hm.group(2)
        if m.lower() not in MARKET_CODES:
            continue
        try:
            nval = int(n)
        except ValueError:
            nval = 0
        if best_m is None or nval > best_nval:
            best_m, best_n, best_nval = m, n, nval
    if best_m:
        return best_m, best_n
    # Fallback: normalize first row's market (strip #N if present)
    m0, n0 = rows[0]['market'].strip(), rows[0]['order_num']
    hm0 = re.match(r'^(.+?)#(\d+)$', m0)
    if hm0:
        return hm0.group(1).strip(), hm0.group(2)
    return m0, n0


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

    market, order_num = _pick_canonical_market(rows)

    sport       = meta.get('sport', 'BSBL')
    season      = meta.get('season', '') or 'Summer'
    order_date  = meta.get('order_date', '')
    ship_date   = meta.get('ship_date', '')
    order_sheet = meta.get('order_sheet_url', '')

    order_date_ms    = to_ms(order_date)
    actual_ship_ms   = to_ms(ship_date)

    saved_order = find_order(market, order_num)
    if saved_order is None:
        # Fallback: find order by matching any TN from packing list rows to manual_shipments.
        # Packing list market names (e.g. "SOCCER") may differ from the stored order market ("Houston").
        row_tns = {r.get('tracking_number', '').strip() for r in rows if r.get('tracking_number')}
        row_tns.discard('')
        if row_tns:
            for _o in load_orders():
                if any(ms.get('tracking_number', '').strip() in row_tns
                       for ms in _o.get('manual_shipments', [])):
                    saved_order = _o
                    market    = _o.get('market', market)
                    order_num = _o.get('order_num', order_num)
                    if not sport or sport == 'BSBL':
                        sport = _o.get('sport', sport)
                    if not season:
                        season = _o.get('season', season) or 'Summer'
                    if not order_date:
                        order_date     = _o.get('order_date', '')
                        order_date_ms  = to_ms(order_date)
                    break

    saved_gear_tasks = saved_order.get('gear_tasks', {}) if saved_order else {}
    new_gear_tasks   = dict(saved_gear_tasks)
    eta_zees_ms      = to_ms(saved_order.get('eta_to_zees', '')) if saved_order else None

    # Build TN → ship date ms from manually-entered shipments (authoritative over form value)
    tn_ship_date = {}
    for ms_entry in (saved_order.get('manual_shipments', []) if saved_order else []):
        tn  = ms_entry.get('tracking_number', '').strip()
        sd  = ms_entry.get('ship_date', '')
        if tn and sd and tn not in tn_ship_date:
            ms_val = to_ms(sd)
            if ms_val:
                tn_ship_date[tn] = ms_val

    # Resolve numeric suffix used in task names / order codes
    _nfc = saved_order.get('num_for_code', '') if saved_order else ''
    if not _nfc:
        _m2 = re.search(r'(\d+)$', order_num.strip())
        _nfc = order_num.strip() if order_num.strip().isdigit() else (_m2.group(1) if _m2 else '1')
    num_for_code = _nfc

    # Build global shipment map: tracking# → letter, primary gear, rows
    t2letter, t2primary, t2rows = build_shipment_map(rows)

    results           = []
    errors            = []
    parent_task_ids   = {}   # gear_type → clickup task id
    gear_task_names   = {}   # gear_type → task name string
    gear_order_codes  = {}   # gear_type → order code string
    gear_order_status = {}   # gear_type → order status orderindex
    old_parent_tns    = {}   # gear_type → TN the parent had before this push
    parent_subs_before = {}  # gear_type → existing subtasks before this push

    # ── PASS 1: create / update one parent task per gear type ────────────────
    for gear_type, boxes in sorted(by_gear.items()):
        gear_trackings = list({b['tracking_number'] for b in boxes if b['tracking_number']})
        tracking_str   = ', '.join(gear_trackings)
        has_tracking   = bool(gear_trackings)
        owned_tns      = [tn for tn in gear_trackings if t2primary.get(tn) == gear_type]

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
            # Capture state BEFORE any writes so Pass 2 can detect TN changes
            _par = _get_clickup_task(existing_task_id)
            _old_tn_raw = (_get_task_cf(_par, CF_TRACKING) if _par else '') or ''
            old_parent_tns[gear_type]    = _old_tn_raw.split(',')[0].strip()
            parent_subs_before[gear_type] = _get_clickup_subtasks(existing_task_id)

            set_clickup_field(existing_task_id, CF_ORDER_STATUS,  order_status)
            set_clickup_field(existing_task_id, CF_TRACKING,      tracking_str)
            gear_ship_ms = _gear_ship_ms(owned_tns or gear_trackings, tn_ship_date, actual_ship_ms)
            if has_tracking and gear_ship_ms:
                set_clickup_field(existing_task_id, CF_SHIP_DATE_ACT, gear_ship_ms)
                set_clickup_field(existing_task_id, CF_PHASE,         3)
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
            gear_ship_ms = _gear_ship_ms(owned_tns or gear_trackings, tn_ship_date, actual_ship_ms)
            if has_tracking and gear_ship_ms:
                custom_fields.append({'id': CF_SHIP_DATE_ACT, 'value': gear_ship_ms})
                custom_fields.append({'id': CF_PHASE,         'value': 3})
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
            for f in custom_fields:
                set_clickup_field(task_id, f['id'], f['value'])
            parent_task_ids[gear_type] = task_id
            new_gear_tasks[gear_type]  = task_id
            results.append({'gear': gear_type, 'task_id': task_id,
                            'subtasks': [], 'action': 'created'})

    # ── PASS 2: subtasks — same split logic as manual tracking entry ──────────
    _FORMULA_FIELD_IDS = {
        '08f6fefb-3605-48e4-bdbb-b5e1c765d87a',
        '43acbd5f-e47a-4230-93fd-cc7f596a731d',
        '0245e34f-f032-44c5-b016-792b4447c658',
        '9e1380de-7dbf-4c77-afdb-7de496f61e2b',
    }

    # Group packing-list TNs by primary gear (preserving appearance order)
    gear_to_new_tns = defaultdict(list)
    for tn in t2letter:
        pg = t2primary[tn]
        if parent_task_ids.get(pg):
            gear_to_new_tns[pg].append(tn)

    for gear_type, new_tns in gear_to_new_tns.items():
        parent_id     = parent_task_ids[gear_type]
        existing_subs = parent_subs_before.get(gear_type)
        if existing_subs is None:
            existing_subs = _get_clickup_subtasks(parent_id)
        old_tn = old_parent_tns.get(gear_type, '')

        if not existing_subs:
            # No existing subtasks — mirror push_update_ship logic
            if old_tn and old_tn not in new_tns:
                # Parent had a different TN → A=old, B/C/...=new PL TNs
                tns_to_create = [old_tn] + list(new_tns)
            elif len(new_tns) > 1:
                # Multiple TNs in packing list → A, B, ...
                tns_to_create = list(new_tns)
            else:
                # Single TN, no split needed
                tns_to_create = []
            letter_base = 0
        else:
            # Existing subtasks — only create for uncovered TNs
            existing_sub_tns = {_get_task_cf(s, CF_TRACKING) for s in existing_subs}
            tns_to_create = [tn for tn in new_tns if tn not in existing_sub_tns]
            letter_base = None  # use _next_subtask_letter

        # Fetch parent once for field copy
        parent_task_data = _get_clickup_task(parent_id)

        for i, sub_tn in enumerate(tns_to_create):
            if letter_base is not None:
                letter = chr(ord('A') + letter_base + i)
            else:
                letter = _next_subtask_letter(existing_subs)

            shipment_rows = t2rows.get(sub_tn, [])
            gear_types_in = sorted({r['gear_type'] for r in shipment_rows}) if shipment_rows else [gear_type]
            contents_desc = ', '.join(gear_types_in)
            weights       = [r['weight'] for r in shipment_rows if r.get('weight')]
            values        = [r['declared_value'] for r in shipment_rows if r.get('declared_value')]
            carrier       = shipment_rows[0].get('carrier', 'fedex').upper() if shipment_rows else 'FEDEX'

            base_name = gear_task_names.get(gear_type, f"{market} {season} {sport} {gear_type} {num_for_code}")
            base_code = gear_order_codes.get(gear_type, num_for_code)
            sub_name  = f"{base_name}{letter}"
            sub_code  = f"{base_code}{letter}"

            sub_fields_map = {}
            if parent_task_data:
                for cf in parent_task_data.get('custom_fields', []):
                    val = cf.get('value')
                    if val is not None and cf['id'] not in _FORMULA_FIELD_IDS:
                        sub_fields_map[cf['id']] = val

            sub_fields_map[CF_TRACKING]   = sub_tn
            sub_fields_map[CF_ORDER_TEXT] = sub_code
            sub_ship_ms = tn_ship_date.get(sub_tn) or actual_ship_ms
            if shipment_rows and sub_ship_ms:
                sub_fields_map[CF_SHIP_DATE_ACT] = sub_ship_ms
            os_val = gear_order_status.get(gear_type)
            if os_val is not None:
                sub_fields_map[CF_ORDER_STATUS] = os_val
            if sub_ship_ms:
                sub_fields_map[CF_PHASE] = 3

            sub_cf_list = [{'id': fid, 'value': val} for fid, val in sub_fields_map.items()]
            sresp = create_clickup_task({
                'name':   sub_name,
                'parent': parent_id,
                'description': (
                    f"Gear: {contents_desc}\n"
                    f"Weight: {'; '.join(weights)}\n"
                    f"Declared Value: {'; '.join(values)}\n"
                    f"Carrier: {carrier}"
                ),
                'custom_fields': sub_cf_list,
            })
            if sresp.status_code in (200, 201):
                sub_id = sresp.json().get('id')
                for f in sub_cf_list:
                    set_clickup_field(sub_id, f['id'], f['value'])
                for r in results:
                    if r['gear'] == gear_type:
                        r['subtasks'].append(sub_id)
                        break
                if letter_base is None:
                    existing_subs.append({'name': f'{base_name}{letter}'})
            else:
                errors.append(f"  Shipment {order_num}{letter}: {sresp.status_code}")

    # ── PASS 3: push sub-orders (Freshman, Remake, …) ───────────────────────
    # Rows whose market is not a known MARKET_CODES key (and not the parent market
    # after #N stripping) belong to a sub-order stored under the same parent market
    # but with order_num = the label (e.g. market="Dallas", order_num="Freshman").
    sub_label_rows: dict = defaultdict(list)
    for r in rows:
        m = r['market'].strip()
        hm = re.match(r'^(.+?)#(\d+)$', m)
        if hm:
            m = hm.group(1).strip()
        if m.lower() not in MARKET_CODES:
            sub_label_rows[r['market'].strip()].append(r)

    for label, sub_rows in sub_label_rows.items():
        sub_saved = find_order(market, label)
        if not sub_saved or not sub_saved.get('gear_tasks'):
            continue

        sub_by_gear: dict = defaultdict(list)
        for r in sub_rows:
            sub_by_gear[r['gear_type']].append(r)

        for gear_type, boxes in sub_by_gear.items():
            task_id = sub_saved['gear_tasks'].get(gear_type)
            if not task_id:
                continue
            gear_trackings = list({b['tracking_number'] for b in boxes if b['tracking_number']})
            if not gear_trackings:
                continue
            tracking_str = ', '.join(gear_trackings)
            gear_ship_ms = _gear_ship_ms(gear_trackings, tn_ship_date, actual_ship_ms)
            set_clickup_field(task_id, CF_TRACKING, tracking_str)
            if gear_ship_ms:
                set_clickup_field(task_id, CF_SHIP_DATE_ACT, gear_ship_ms)
                set_clickup_field(task_id, CF_PHASE, 3)
            results.append({'gear': f'{label} {gear_type}', 'task_id': task_id,
                            'subtasks': [], 'action': 'updated'})

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
                new_id = resp.json().get('id')
                results.append({'gear': gear_type, 'task_id': new_id, 'action': 'created'})
                # Set each field individually — the create payload silently drops fields
                # that ClickUp can't validate at creation time (date fields, dropdowns, etc.)
                all_create_fields = custom_fields[:]  # includes dates + phase + order_status etc.
                for f in all_create_fields:
                    set_clickup_field(new_id, f['id'], f['value'])
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
        # Physical box count: deduplicate rows sharing the same (box_num, contents) key
        _phys = set()
        for r in shipment_rows:
            c = r.get('contents', '').strip()
            if c.lower() == 'see attached':
                c = ''
            _phys.add((r.get('box_num', ''), c))
        box_count = len(_phys)
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

    # Collect selected TNs (deduplicated) — user may select only one checkbox per TN
    selected_tns: dict = {}   # tn → first selected order (for contact/market info)
    for item in items:
        if '|' not in item:
            continue
        order_id, tn = item.split('|', 1)
        tn = tn.strip()
        order = orders_map.get(order_id)
        if not tn or not order:
            continue
        if tn not in selected_tns:
            selected_tns[tn] = order   # keep first selected order for contact info

    rows = []

    for tn, first_order in selected_tns.items():
        market     = first_order.get('market', '')
        order_date = first_order.get('order_date', '')
        carrier    = 'fedex'

        try:
            contact = _ship_to_info(market)
        except ValueError:
            continue
        clean_market = contact['market']

        # Always collect gear from ALL orders that have this TN — not just selected ones.
        # User uploads one TN to AfterShip; custom fields must reflect everything in that shipment.
        all_entries = []
        order_labels = []
        for order in orders_all:
            entries = [x for x in order.get('manual_shipments', [])
                       if x.get('tracking_number', '').strip() == tn]
            if not entries:
                continue
            all_entries.extend(entries)
            carrier = entries[0].get('carrier', 'fedex').lower()
            sport      = order.get('sport', 'BSBL')
            order_num  = order.get('num_for_code', order.get('order_num', ''))
            gear_types_o = list(dict.fromkeys(e.get('gear_type', '') for e in entries if e.get('gear_type')))
            primary_gear = gear_types_o[0] if gear_types_o else ''
            order_code_o = build_order_code(market, sport, primary_gear, order_date, order_num)
            order_labels.append(f"Order {order_code_o}: {', '.join(gear_types_o)}")

        if not all_entries:
            continue

        # Use first order's code as the CSV order_id field.
        # Prefer ClickUp subtask CF_ORDER_TEXT (has letter suffix like "1B").
        first_sport     = first_order.get('sport', 'BSBL')
        first_order_num = first_order.get('num_for_code', first_order.get('order_num', ''))
        first_entries   = [x for x in first_order.get('manual_shipments', [])
                           if x.get('tracking_number', '').strip() == tn]
        first_gear      = first_entries[0].get('gear_type', '') if first_entries else ''
        order_code      = build_order_code(market, first_sport, first_gear, order_date, first_order_num)
        first_task_id   = first_order.get('gear_tasks', {}).get(first_gear)
        if first_task_id:
            for sub in _get_clickup_subtasks(first_task_id):
                if _get_task_cf(sub, CF_TRACKING) == tn:
                    sub_code = _get_task_cf(sub, CF_ORDER_TEXT)
                    if sub_code:
                        order_code = sub_code
                    break

        # Aggregate across all orders — count physical boxes.
        # Same notes + same box_count → shared physical boxes (use count once).
        # Same notes + different box_counts → distinct boxes per gear (sum them).
        # No notes → each entry contributes its own box_count.
        _notes_groups: dict = {}  # notes str → list of box_counts
        _solo_boxes = 0
        for e in all_entries:
            n = e.get('notes', '').strip()
            bc = int(e.get('box_count') or 1)
            if n:
                _notes_groups.setdefault(n, []).append(bc)
            else:
                _solo_boxes += bc
        box_count = _solo_boxes + sum(
            counts[0] if len(set(counts)) == 1 else sum(counts)
            for counts in _notes_groups.values()
        )
        declared_vals = [e.get('declared_value', '').strip() for e in all_entries if e.get('declared_value', '').strip()]
        note_texts    = list(dict.fromkeys(e.get('notes', '').strip() for e in all_entries if e.get('notes', '').strip()))

        # custom_1: flat, deduplicated gear list across all orders sharing this TN
        all_gear_types = list(dict.fromkeys(
            e.get('gear_type', '') for e in all_entries if e.get('gear_type', '')
        ))
        custom_1 = ', '.join(all_gear_types) if all_gear_types else ''

        # custom_2: full detail note per order
        contents_parts = ' + '.join(order_labels) if order_labels else 'Mixed'
        custom_2 = f"{box_count} box{'es' if box_count != 1 else ''} | {contents_parts}"
        if declared_vals:
            total_declared = 0.0
            for v in declared_vals:
                try:
                    total_declared += float(v.replace('$', '').replace(',', ''))
                except ValueError:
                    pass
            custom_2 += f" | Declared: ${total_declared:g}"
        if note_texts:
            custom_2 += ' | ' + ' | '.join(note_texts)

        order_date_iso = (order_date + 'T00:00:00+00:00') if order_date else ''

        detected = _detect_carrier(tn)
        if detected == 'ups':
            shipment_type  = 'UPS Worldwide Expedited'
            origin_state   = 'Shanghai'
            origin_city    = 'Shanghai'
        else:
            # Use stored service_type from any matching manual_shipment, else default
            stored_service = next(
                (e.get('service_type', '') for e in all_entries if e.get('service_type')),
                ''
            )
            shipment_type  = stored_service or 'FedEx International Connect Plus'
            origin_state   = 'Shandong'
            origin_city    = 'Qingdao'

        row = [
            carrier, tn, order_code, order_code, '',
            'CHN', 'USA', custom_1, custom_2,
            '', '', '', '', '',
            order_date_iso, order_code,
            'en', shipment_type, '', '', '',
            origin_state, origin_city, '', '',
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

    # Build gear_types from packing_list using contents-aware physical box detection:
    # - rows with the same (box_num, non-empty contents) = same physical box → count once
    #   under the highest-priority gear in that box
    # - rows with empty/"see attached" contents each count as their own physical box
    #   (per-gear box numbering assumed when contents are missing)
    for tn, boxes in tn_boxes.items():
        first = next((b for b in packing if b.get('tracking_number', '').strip() == tn), {})
        all_pl_gears = {box['gear'] for box in boxes}
        by_physical_box: dict = {}
        for box in boxes:
            bn = box['box_num']
            c = box.get('contents', '').strip()
            if c.lower() == 'see attached':
                c = ''
            # Empty/"see attached" → box_num is absolute; group by (box_num, '') so
            # all gear types in the same numbered box share one physical box.
            # Non-empty contents → unique contents string identifies the physical box.
            key = (bn, c)
            by_physical_box.setdefault(key, []).append(box['gear'])
        gear_counts: dict = {}
        for gears_in_box in by_physical_box.values():
            primary = _primary_gear(gears_in_box) or gears_in_box[0]
            gear_counts[primary] = gear_counts.get(primary, 0) + 1
        # Non-primary gears in shared boxes get count 0 so they still appear in
        # gear_summary (for sync/delivered logic) but don't inflate box counts.
        for g in all_pl_gears:
            if g not in gear_counts:
                gear_counts[g] = 0
        tn_info[tn] = {
            'carrier':    first.get('carrier', 'fedex'),
            'gear_types': gear_counts,
            'notes':      '',
        }

    # Merge manual shipments (entered via update panel).
    # When packing_list already covers a TN, skip box_count (avoid double-counting);
    # still merge notes, ship_date, and carrier for display.
    for m in order.get('manual_shipments', []):
        tn = m.get('tracking_number', '').strip()
        if not tn:
            continue
        has_pl = tn in tn_boxes  # True only if packing_list covers this TN
        if tn not in tn_info:
            tn_info[tn] = {
                'carrier':    m.get('carrier', 'FedEx'),
                'gear_types': {},
                'gear_notes': {},
                'notes':      m.get('notes', ''),
                'ship_date':  m.get('ship_date', ''),
                'added_at':   m.get('added_at', ''),
            }
        else:
            if not tn_info[tn].get('ship_date') and m.get('ship_date'):
                tn_info[tn]['ship_date'] = m['ship_date']
            if not tn_info[tn].get('added_at') and m.get('added_at'):
                tn_info[tn]['added_at'] = m['added_at']
        gt = m.get('gear_type', '')
        if gt:
            notes_str = m.get('notes', '')
            tn_info[tn].setdefault('gear_notes', {})[gt] = notes_str
            if not has_pl:
                try:
                    boxes = max(1, int(m.get('box_count', 1)))
                except (TypeError, ValueError):
                    boxes = 1
                # If box_count is still the default (1), try parsing from notes
                # e.g. "2 Boxes", "7 boxes", "3 box"
                if boxes == 1 and notes_str:
                    bc_m = re.search(r'\b(\d+)\s+box(?:es)?\b', notes_str, re.IGNORECASE)
                    if bc_m:
                        boxes = max(1, int(bc_m.group(1)))
                tn_info[tn]['gear_types'][gt] = tn_info[tn]['gear_types'].get(gt, 0) + boxes

    # For manual-only TNs (no packing list): gears sharing the same non-empty notes
    # are in the same physical box. Keep box_count on the highest-priority gear only;
    # set the rest to 0 so they appear in gear_summary without inflating box totals.
    pl_tns = set(tn_boxes.keys())
    for tn, info in tn_info.items():
        if tn in pl_tns:
            continue
        gear_notes = info.get('gear_notes', {})
        gear_types = info['gear_types']
        notes_groups: dict = {}
        for g in list(gear_types.keys()):
            n = gear_notes.get(g, '').strip()
            if n:
                notes_groups.setdefault(n, []).append(g)
        for gears_in_group in notes_groups.values():
            if len(gears_in_group) <= 1:
                continue
            primary = _primary_gear(gears_in_group) or gears_in_group[0]
            count = max(gear_types.get(g, 0) for g in gears_in_group)
            for g in gears_in_group:
                gear_types[g] = count if g == primary else 0

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
                    _noted[n] = {'gears': [], 'box_counts': []}
                _noted[n]['box_counts'].append(g['boxes'])
                _noted[n]['gears'].append(g['gear'])
            else:
                _solo.append(g)
        _individual = []  # single-gear groups (one gear per physical box)
        _shared     = []  # multi-gear groups (multiple gear types share a box)
        for grp in _noted.values():
            if any(b == 0 for b in grp['box_counts']):
                total = sum(b for b in grp['box_counts'] if b > 0) or 1
            else:
                total = max(grp['box_counts']) if grp['box_counts'] else 1
            bw   = 'box' if total == 1 else 'boxes'
            line = f"{', '.join(sorted(grp['gears']))} × {total} {bw}"
            if len(grp['gears']) == 1:
                _individual.append(line)
            else:
                _shared.append(line)
        for g in sorted(_solo, key=lambda x: x['gear']):
            if g['boxes'] == 0:
                continue
            bw = 'box' if g['boxes'] == 1 else 'boxes'
            _individual.append(f"{g['gear']} × {g['boxes']} {bw}")
        gear_groups = sorted(_individual) + _shared
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
            'location':        saved.get('location', ''),
            'updated_at':      saved.get('updated_at', ''),
            'delivered_at':    saved.get('delivered_at', ''),
        })

    return shipments


def _build_tracking_payload():
    orders = load_display_orders()
    # Most recently shipped order first
    orders_sorted = sorted(orders, key=lambda o: o.get('shipped_at', ''), reverse=True)
    result_orders = []
    for order in orders_sorted:
        shipments = _collect_shipments(order)
        # Most recently updated shipment first; fall back to shipped_at on order
        shipments.sort(key=lambda s: s.get('updated_at', '') or order.get('shipped_at', ''), reverse=True)
        dest = _ship_to_info(order.get('market', ''))
        result_orders.append({
            'id':           order.get('id'),
            'market':       order.get('market', ''),
            'sport':        order.get('sport', ''),
            'season':       order.get('season', ''),
            'order_num':    order.get('num_for_code', order.get('order_num', '')),
            'order_date':   order.get('order_date', ''),
            'shipped_at':   order.get('shipped_at', ''),
            'gear_tasks':   order.get('gear_tasks', {}),
            'shipments':    shipments,
            'dest_state':   dest.get('state', '') if dest else '',
            'dest_city':    dest.get('city', '') if dest else '',
            'packing_docs': order.get('packing_docs', []),
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
        carrier    = _carrier_emoji(ship.get('carrier', 'fedex'))
        gear_lines = '\n'.join(f"  • {g}" for g in ship.get('gear_groups', []))
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


@app.route('/packing-doc/<order_id>/<filename>')
def serve_packing_doc(order_id, filename):
    safe     = secure_filename(filename)
    filepath = PACKING_DOCS_FOLDER / secure_filename(order_id) / safe
    if not filepath.exists():
        return 'Not found', 404
    # Look up the download_name from orders.json
    orders   = load_orders()
    order    = next((o for o in orders if o.get('id') == order_id), None)
    dl_name  = filename
    if order:
        doc = next((d for d in order.get('packing_docs', []) if d.get('filename') == safe), None)
        if doc and doc.get('download_name'):
            dl_name = doc['download_name']
    as_attachment = request.args.get('dl') == '1'
    return send_file(str(filepath), as_attachment=as_attachment, download_name=dl_name)


@app.route('/api/order/<order_id>/packing-doc', methods=['POST'])
def upload_packing_doc(order_id):
    orders = load_orders()
    order = next((o for o in orders if o.get('id') == order_id), None)
    if not order:
        return jsonify({'ok': False, 'error': 'Order not found'}), 404
    f = request.files.get('file')
    if not f or not f.filename:
        return jsonify({'ok': False, 'error': 'No file provided'}), 400
    label    = request.form.get('label', '').strip()
    filename = secure_filename(f.filename)
    folder   = PACKING_DOCS_FOLDER / secure_filename(order_id)
    folder.mkdir(exist_ok=True)
    f.save(str(folder / filename))
    doc = {'filename': filename, 'label': label or filename, 'uploaded_at': datetime.now().isoformat()}
    order.setdefault('packing_docs', [])
    order['packing_docs'] = [d for d in order['packing_docs'] if d['filename'] != filename]
    order['packing_docs'].append(doc)
    save_orders(orders)
    return jsonify({'ok': True, 'doc': doc})


@app.route('/api/order/<order_id>/packing-doc', methods=['DELETE'])
def delete_packing_doc(order_id):
    body     = request.get_json() or {}
    filename = secure_filename(body.get('filename', ''))
    orders   = load_orders()
    order    = next((o for o in orders if o.get('id') == order_id), None)
    if not order:
        return jsonify({'ok': False, 'error': 'Order not found'}), 404
    filepath = PACKING_DOCS_FOLDER / secure_filename(order_id) / filename
    if filepath.exists():
        filepath.unlink()
    order['packing_docs'] = [d for d in order.get('packing_docs', []) if d['filename'] != filename]
    save_orders(orders)
    return jsonify({'ok': True})


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
            cu_errors.extend(_sync_delivered_to_clickup(order, tn, actual_ms, all_orders=orders))

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

    errs = _sync_delivered_to_clickup(order, tn, actual_ms, all_orders=orders)
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
    orders = load_display_orders()
    reminders = []

    for order in orders:
        order_label = (
            f"{order.get('market')} {order.get('season')} "
            f"{order.get('sport')} #{order.get('num_for_code', order.get('order_num'))}"
        )

        # ── ETA to Zees reminder ─────────────────────────────────────────────
        # Show if any gear is unshipped OR any tracking number is not yet delivered.
        gear_types = set(order.get('gear_tasks', {}).keys())
        shipped_gears = (
            {ms.get('gear_type') for ms in order.get('manual_shipments', [])} |
            {row.get('gear_type') for row in order.get('packing_list', [])}
        )
        all_shipped = bool(gear_types) and gear_types.issubset(shipped_gears)
        tracking_statuses = order.get('tracking_statuses', {})
        all_tn_delivered = (
            bool(tracking_statuses) and
            all(v.get('tag') == 'Delivered' for v in tracking_statuses.values())
        )
        eta_zees_str = order.get('eta_to_zees', '')
        if eta_zees_str and not (all_shipped and all_tn_delivered):
            try:
                eta_date   = datetime.fromisoformat(eta_zees_str).date()
                days_until = (eta_date - today).days
                if days_until <= 7:   # overdue or within 7 days (1 week prior)
                    reminders.append({
                        'order_id':    order.get('id'),
                        'order_label': order_label,
                        'gear_type':   'All Gear',
                        'milestone':   'ETA to Zees',
                        'date':        eta_date.isoformat(),
                        'days_diff':   days_until,
                    })
            except ValueError:
                pass

        gear_tasks = order.get('gear_tasks', {})
        if not gear_tasks:
            continue
        order_date_str = order.get('order_date', '')
        if not order_date_str:
            continue

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

            # Fallback: if CF_SHIP_DATE_ACT not set in ClickUp, check manual_shipments
            if ship_act is None:
                for ms in order.get('manual_shipments', []):
                    if ms.get('gear_type') == gear_type and ms.get('ship_date'):
                        try:
                            sd = datetime.strptime(ms['ship_date'], '%Y-%m-%d')
                            ship_act = int(sd.timestamp() * 1000)
                        except ValueError:
                            pass
                        break

            for f in fields:
                label = _REMINDER_CF_LABELS.get(f['id'])
                if not label:
                    continue

                milestone_date = datetime.fromtimestamp(f['value'] / 1000).date()
                days_diff = (milestone_date - today).days

                # Delivery Date: show from 7 days before until delivered
                # All other milestones: only show if overdue
                if label == 'Delivery Date':
                    if days_diff > 7:
                        continue
                elif milestone_date > today:
                    continue

                # Condition per milestone type — only show if the action hasn't happened yet
                if label in ('Proofs Ready', 'Proofs Approved'):
                    if proofs_approved_act is not None or ship_act is not None:
                        continue
                elif label == 'Ship Date':
                    if ship_act is not None:
                        continue
                elif label == 'Delivery Date':
                    # Skip if not yet shipped, or already delivered
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


_sheet_cache: dict = {}   # sheet_id → {'ts': float, 'gear_types': dict}
_SHEET_CACHE_TTL = 600    # 10 minutes

# Maps substrings in order-sheet item names → gear_task key
_SHEET_GEAR_MAP = [
    ('flex jersey', 'Jersey'),
    ('jersey',      'Jersey'),
    ('hoodie',      'Hoodie'),
    ('tee',         'Tee'),
    ('cap',         'Cap'),
    ('windbreaker', 'Windbreaker'),
    ('jacket',      'Windbreaker'),
    ('ring',        'Ring'),
    ('hype',        'Hypechain'),
    ('necklace',    'Hypechain'),
    ('name banner', 'Name Banner'),
    ('park banner', 'Park Banner'),
]

# Tab name keywords to verify a gear type has a real vendor order (not just an estimate).
# If the matching tab has only 1 numeric row (a single summary), treat qty as 0.
_GEAR_TAB_KW = {
    'Jersey':      ['jersey'],
    'Hoodie':      ['hoodie'],
    'Tee':         ['tee'],
    'Cap':         ['cap'],
    'Windbreaker': ['jacket', 'windbreaker'],
    'Ring':        ['ring'],
    'Hypechain':   ['hype', 'necklace'],
    'Name Banner': ['name banner'],
    'Park Banner': ['park banner'],
}


def _count_numeric_rows(csv_text: str) -> int:
    """Count rows in a CSV that contain at least one cell parseable as a number ≥ 1."""
    count = 0
    for row in csv.reader(StringIO(csv_text)):
        for cell in row:
            try:
                if float(cell.strip().replace(',', '')) >= 1:
                    count += 1
                    break
            except ValueError:
                pass
    return count


def _qty_from_packing_contents(gear_type: str, pl_rows: list) -> int:
    """
    Parse unit qty from packing_list contents fields.
    Handles size breakdowns (YM-82, AS-47...) and explicit counts (rings-78, necklace-30).
    Returns 0 if nothing parseable.
    """
    if not pl_rows:
        return 0
    gear_lower = gear_type.lower()
    total = 0
    seen_contents: set = set()
    for row in pl_rows:
        content = (row.get('contents') or '').strip()
        if not content or content in seen_contents:
            continue
        seen_contents.add(content)
        # Size breakdown: YM-82, AS-47, YL-177 etc — sum all size-N patterns
        size_matches = re.findall(r'\b(?:YXS|YS|YM|YL|YXL|AS|AM|AL|AXL|AXXL|S|M|L|XL|XXL)-(\d+)\b', content, re.I)
        if size_matches:
            total += sum(int(x) for x in size_matches)
            continue
        # Explicit item counts: rings-78, necklace-30, namebanner-7, parkbanner-20
        item_kw = {
            'ring': 'Ring', 'necklace': 'Hypechain', 'hypechain': 'Hypechain',
            'namebanner': 'Name Banner', 'parkbanner': 'Park Banner', 'parkbannner': 'Park Banner',
            'bag': 'Drawstring Bag', 'cap': 'Cap',
        }
        for kw, mapped_gear in item_kw.items():
            if mapped_gear.lower() == gear_lower or kw in gear_lower:
                for m in re.finditer(rf'{re.escape(kw)}s?[\s\-](\d+)', content, re.I):
                    total += int(m.group(1))
    return total


def _ordered_gear_types(sheet_url: str) -> dict | None:
    """Return {gear_type: qty} for confirmed gear orders on the sheet, or None on error.

    A gear type is 'confirmed' when:
      1. The Order Home Page shows qty > 0, AND
      2. The matching gear tab has ≥ 2 numeric rows (real size/team breakdown),
         not just a single summary row (which indicates a registration estimate).
    """
    if not sheet_url:
        return None
    m = re.search(r'/spreadsheets/d/([a-zA-Z0-9_-]+)', sheet_url)
    if not m:
        return None
    sheet_id = m.group(1)

    cached = _sheet_cache.get(sheet_id)
    if cached and (time.time() - cached['ts']) < _SHEET_CACHE_TTL:
        return cached['gear_types'] or None

    gkey = os.getenv('GOOGLE_API_KEY', '')
    if not gkey:
        return None

    try:
        meta = requests.get(
            f'https://sheets.googleapis.com/v4/spreadsheets/{sheet_id}',
            params={'key': gkey, 'fields': 'sheets.properties'},
            timeout=10,
        )
        meta.raise_for_status()
        tabs = meta.json().get('sheets', [])
    except Exception:
        return None

    home_gid   = None
    home_title = ''
    for tab in tabs:
        name = tab['properties']['title'].lower()
        if 'order home' in name or 'home page' in name:
            home_gid   = str(tab['properties']['sheetId'])
            home_title = name
            break
    if home_gid is None and tabs:
        home_gid   = str(tabs[0]['properties']['sheetId'])
        home_title = tabs[0]['properties']['title'].lower()
    if home_gid is None:
        return None

    try:
        csv_resp = requests.get(
            f'https://docs.google.com/spreadsheets/d/{sheet_id}/export',
            params={'format': 'csv', 'gid': home_gid},
            timeout=15,
        )
        csv_resp.raise_for_status()
        if 'accounts.google.com' in csv_resp.url:
            return None
    except Exception:
        return None

    # Parse home page: col B (index 1) = item name, col E (index 4) = qty
    candidate_qty: dict[str, int] = {}
    for row in csv.reader(StringIO(csv_resp.text)):
        if len(row) < 5:
            continue
        name_cell = row[1].strip().lower()
        qty_str   = row[4].strip().replace(',', '').replace('$', '')
        try:
            qty = float(qty_str)
        except ValueError:
            continue
        if qty <= 0:
            continue
        for substr, gear_type in _SHEET_GEAR_MAP:
            if substr in name_cell:
                if gear_type not in candidate_qty:
                    candidate_qty[gear_type] = int(qty)
                break

    if not candidate_qty:
        _sheet_cache[sheet_id] = {'ts': time.time(), 'gear_types': {}}
        return None

    # Verify each candidate against its gear-specific tab.
    # A tab with only 1 numeric row is a summary placeholder, not a real vendor order.
    confirmed: dict[str, int] = {}
    for gear_type, qty in candidate_qty.items():
        keywords = _GEAR_TAB_KW.get(gear_type, [])
        gear_gid = None
        for tab in tabs:
            tname = tab['properties']['title'].lower()
            if tname == home_title:
                continue
            if any(kw in tname for kw in keywords):
                gear_gid = str(tab['properties']['sheetId'])
                break

        if gear_gid is None:
            confirmed[gear_type] = qty
            continue

        try:
            gr = requests.get(
                f'https://docs.google.com/spreadsheets/d/{sheet_id}/export',
                params={'format': 'csv', 'gid': gear_gid},
                timeout=15,
            )
            gr.raise_for_status()
            if _count_numeric_rows(gr.text) >= 2:
                confirmed[gear_type] = qty
        except Exception:
            confirmed[gear_type] = qty  # on fetch error, trust home page

    _sheet_cache[sheet_id] = {'ts': time.time(), 'gear_types': confirmed}
    return confirmed if confirmed else None


_cu_task_cache: dict = {}
_CU_TASK_TTL = 1800  # 30 minutes

_ERROR_KW       = ('remake', 'redo', 're-order', 'reorder', 'replacement')
_CU_AUTO_PREFIXES = ('📦 Shipped', '📬', 'Tracking:', 'Shipped\n', '✅')


def _get_task_error_info(task_id: str) -> dict | None:
    """Return {'is_error': bool, 'reason': str|None} for a ClickUp task.

    'is_error' is True only when the task name contains a remake/redo keyword.
    'reason' is the first non-shipping comment text (first line only).
    Returns None on fetch error.
    """
    if not task_id:
        return None
    cached = _cu_task_cache.get(task_id)
    if cached is not None and (time.time() - cached['ts']) < _CU_TASK_TTL:
        return cached.get('info')

    cu_token = os.getenv('CLICKUP_TOKEN', '')
    if not cu_token:
        return None

    try:
        task_resp = requests.get(
            f'https://api.clickup.com/api/v2/task/{task_id}',
            headers={'Authorization': cu_token},
            timeout=10,
        )
        task_resp.raise_for_status()
        name = task_resp.json().get('name', '')
    except Exception:
        _cu_task_cache[task_id] = {'ts': time.time(), 'info': None}
        return None

    is_error = any(kw in name.lower() for kw in _ERROR_KW)

    reason      = None
    comment_qty = None
    try:
        comm_resp = requests.get(
            f'https://api.clickup.com/api/v2/task/{task_id}/comment',
            headers={'Authorization': cu_token},
            timeout=10,
        )
        comm_resp.raise_for_status()
        for comment in comm_resp.json().get('comments', []):
            text = comment.get('comment_text', '').strip()
            if not text or any(text.startswith(p) for p in _CU_AUTO_PREFIXES):
                continue
            lines = text.split('\n')
            if is_error and reason is None:
                reason = lines[0].rstrip(':').strip()
            # Parse qty only when an explicit qty/unit keyword is present
            for line in lines:
                m = re.search(
                    r'(?:qty[\s\w]*[:\-]\s*|quantity[\s\w]*[:\-]\s*|total[:\s]+)(\d+)'
                    r'|(\d+)\s*(?:pcs?|units?|caps?|bags?|jerseys?|hoodies?|tees?|rings?|pieces?)\b',
                    line, re.I
                )
                if m:
                    val = int(m.group(1) or m.group(2))
                    if val > 0:
                        comment_qty = (comment_qty or 0) + val
            if comment_qty:
                break  # use the first comment that has a qty
    except Exception:
        pass

    info = {'is_error': is_error, 'reason': reason, 'comment_qty': comment_qty, 'name': name}
    _cu_task_cache[task_id] = {'ts': time.time(), 'info': info}
    return info


@app.route('/api/shipment-status')
def api_shipment_status():
    orders = load_display_orders()

    def tag_to_status(tag):
        return 'arrived' if tag == 'Delivered' else 'shipped'

    # market → gear_type → counts + item lists
    # Items: {'order': order_num, 'sport': sport, 'qty': qty, 'tn': tn_or_empty, 'reason': str_or_None}
    def _new_bucket():
        return {
            'arrived': 0, 'shipped': 0, 'in_production': 0, 'error': 0,
            'arrived_items': [], 'shipped_items': [], 'in_production_items': [], 'error_items': [],
        }
    market_gear: dict = defaultdict(lambda: defaultdict(_new_bucket))

    for order in orders:
        market            = order.get('market', 'Unknown')
        order_num         = str(order.get('order_num', '?'))
        sport             = order.get('sport', '')
        tracking_statuses = order.get('tracking_statuses', {})
        gear_tasks        = order.get('gear_tasks', {})
        manual_shipments  = order.get('manual_shipments', [])
        packing_list      = order.get('packing_list', [])
        sheet_url         = order.get('order_sheet_url', '')

        ordered_gear = _ordered_gear_types(sheet_url)  # None = couldn't read sheet

        for gear_type, task_id in gear_tasks.items():
            has_evidence = any(
                e.get('gear_type') == gear_type
                for e in (manual_shipments + packing_list)
            )
            if ordered_gear is not None and gear_type not in ordered_gear and not has_evidence:
                continue

            bucket      = market_gear[market][gear_type]
            task_info   = _get_task_error_info(task_id)
            is_error    = bool(task_info and task_info.get('is_error'))
            reason      = (task_info or {}).get('reason') if is_error else None
            comment_qty = (task_info or {}).get('comment_qty')

            # qty: order sheet → packing contents → comment → 1
            sheet_qty    = (ordered_gear or {}).get(gear_type, 0)
            content_qty  = _qty_from_packing_contents(gear_type, packing_list) if not sheet_qty else 0
            qty = sheet_qty or content_qty or comment_qty or 1

            seen_tns: set[str] = set()
            ship_items: list[dict] = []
            tags: list[str]        = []

            if manual_shipments:
                for ms in manual_shipments:
                    if ms.get('gear_type') == gear_type:
                        tn  = ms.get('tracking_number', '')
                        tag = tracking_statuses.get(tn, {}).get('tag', '')
                        if tn not in seen_tns:
                            seen_tns.add(tn)
                            ship_items.append({'order': order_num, 'sport': sport, 'qty': qty, 'tn': tn})
                            tags.append(tag)

            if not tags:
                for pl in packing_list:
                    if pl.get('gear_type') == gear_type:
                        tn  = pl.get('tracking_number', '')
                        tag = tracking_statuses.get(tn, {}).get('tag', '')
                        if tn not in seen_tns:
                            seen_tns.add(tn)
                            ship_items.append({'order': order_num, 'sport': sport, 'qty': qty, 'tn': tn})
                            tags.append(tag)

            error_qty = comment_qty or qty
            item_base = {'order': order_num, 'sport': sport, 'qty': qty}

            if is_error:
                bucket['error'] += error_qty
                tns = list(dict.fromkeys(i['tn'] for i in ship_items if i['tn']))
                bucket['error_items'].append({
                    'order':      order_num,
                    'sport':      sport,
                    'gear_type':  gear_type,
                    'market':     market,
                    'qty':        error_qty,
                    'reason':     reason,
                    'tns':        tns,
                })
            elif not tags:
                bucket['in_production'] += qty
                bucket['in_production_items'].append(item_base)
            else:
                statuses = {tag_to_status(t) if t else 'shipped' for t in tags}
                if 'arrived' in statuses:
                    bucket['arrived'] += qty
                    bucket['arrived_items'].extend(ship_items)
                else:
                    bucket['shipped'] += qty
                    bucket['shipped_items'].extend(ship_items)

    result = {}
    for market, gear_data in market_gear.items():
        gear_list = []
        m_arrived = m_shipped = m_in_prod = m_error = 0

        for gear_type, counts in sorted(gear_data.items()):
            # bar_total excludes error so bar percentages sum to 100%
            bar_total = counts['arrived'] + counts['shipped'] + counts['in_production']
            all_total = bar_total + counts['error']
            gear_list.append({
                'gear_type':           gear_type,
                'arrived':             counts['arrived'],
                'shipped':             counts['shipped'],
                'in_production':       counts['in_production'],
                'error':               counts['error'],
                'total':               bar_total,
                'arrived_items':       counts['arrived_items'],
                'shipped_items':       counts['shipped_items'],
                'in_production_items': counts['in_production_items'],
                'error_items':         counts['error_items'],
            })
            m_arrived  += counts['arrived']
            m_shipped  += counts['shipped']
            m_in_prod  += counts['in_production']
            m_error    += counts['error']

        def sort_key(g):
            if g['error'] > 0:
                return 0
            if g['in_production'] > 0 and g['arrived'] == 0 and g['shipped'] == 0:
                return 3
            if g['shipped'] > 0 and g['arrived'] == 0 and g['in_production'] == 0:
                return 2
            if g['shipped'] > 0 or g['in_production'] > 0:
                return 1
            return 0

        gear_list.sort(key=sort_key)
        bar_total = m_arrived + m_shipped + m_in_prod   # denominator for bar %
        all_total = bar_total + m_error                 # for error % only
        # Use floor for arrived so it never rounds up to 100% when items are still in transit/production
        arrived_pct = int(m_arrived / bar_total * 100) if bar_total else 0
        shipped_pct = round(m_shipped / bar_total * 100) if bar_total else 0
        in_prod_pct = round(m_in_prod / bar_total * 100) if bar_total else 0
        result[market] = {
            'arrived_pct':       arrived_pct,
            'shipped_pct':       shipped_pct,
            'in_production_pct': in_prod_pct,
            'error_pct':         round(m_error / all_total * 100) if all_total else 0,
            'arrived':           m_arrived,
            'shipped':           m_shipped,
            'in_production':     m_in_prod,
            'error':             m_error,
            'total':             bar_total,
            'gear':              gear_list,
        }

    return jsonify({'ok': True, 'markets': result})


@app.route('/dashboard')
def dashboard():
    return render_template('dashboard.html')


@app.route('/contents')
def contents_lookup():
    orders = load_display_orders()
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
    if not pub_authed():
        return redirect('/public/login?next=/public/orders')
    orders = load_display_orders()
    enriched = []
    for o in sorted(orders, key=lambda x: x.get('created_at', ''), reverse=True):
        shipments = _collect_shipments(o)
        enriched.append({**o, 'shipments': shipments})
    return render_template('public_orders.html', orders=enriched)


@app.route('/public/tracking')
def public_tracking():
    if not pub_authed():
        return redirect('/public/login?next=/public/tracking')
    return render_template('public_tracking.html')


@app.route('/public/contents')
def public_contents():
    if not pub_authed():
        return redirect('/public/login?next=/public/contents')
    orders = load_display_orders()
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
    if not pub_authed():
        return redirect('/public/login?next=/public/dashboard')
    return render_template('public_dashboard.html')


@app.route('/calendar')
def calendar():
    return render_template('calendar.html')

@app.route('/public/calendar')
def public_calendar():
    if not pub_authed():
        return redirect('/public/login?next=/public/calendar')
    return render_template('public_calendar.html')

@app.route('/public/order-sheet')
def public_order_sheet():
    host = request.host.split(':')[0]
    if host == 'orders.msm-hq.com':
        if not orders_authed():
            return redirect('/public/orders-login')
    elif not pub_authed():
        return redirect('/public/login?next=/public/order-sheet')
    return send_file('/Users/mba/order-sheet-tool/index.html')


@app.route('/public/orders-login', methods=['GET', 'POST'])
def public_orders_login():
    error = None
    if request.method == 'POST':
        if request.form.get('password') in ORDERS_PASSES:
            session['orders_authed'] = True
            return redirect('/public/order-sheet')
        error = 'Incorrect password.'
    return render_template('public_orders_login.html', error=error)


@app.route('/public/orders-logout')
def public_orders_logout():
    session.pop('orders_authed', None)
    return redirect('/public/orders-login')


@app.route('/public/login', methods=['GET', 'POST'])
def public_login():
    error = None
    if request.method == 'POST':
        if request.form.get('password') == PUBLIC_PASS:
            session['pub_authed'] = True
            next_url = request.args.get('next', '/public/dashboard')
            return redirect(next_url)
        error = 'Incorrect password.'
    next_url = request.args.get('next', '/public/dashboard')
    return render_template('public_login.html', next_url=next_url, error=error)


@app.route('/public/logout')
def public_logout():
    session.pop('pub_authed', None)
    return redirect('/public/login')


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

    # Fill missing weight / declared_value on matching manual_shipment entries
    for ms in order.get('manual_shipments', []):
        if ms.get('tracking_number', '').strip() != tracking:
            continue
        gear = ms.get('gear_type', '')
        gear_boxes = [b for b in boxes if b.get('gear_type') == gear] or boxes
        if not ms.get('weight'):
            w = next((b['weight'] for b in gear_boxes if b.get('weight')), '')
            if w:
                ms['weight'] = w
        if not ms.get('declared_value'):
            dv = next((b['declared_value'] for b in gear_boxes if b.get('declared_value')), '')
            if dv:
                ms['declared_value'] = dv

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
        'Cage Jacket': ['cage jacket'],
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
    from concurrent.futures import ThreadPoolExecutor
    orders = load_display_orders()
    result = []

    # Pre-fetch all ClickUp tasks in parallel to avoid sequential HTTP calls
    all_task_ids = list({
        tid
        for o in orders
        for tid in o.get('gear_tasks', {}).values()
        if tid
    })
    with ThreadPoolExecutor(max_workers=10) as ex:
        list(ex.map(_get_clickup_task, all_task_ids))

    for order in orders:
        order_date_str = order.get('order_date', '')
        try:
            order_date = datetime.strptime(order_date_str, '%Y-%m-%d') if order_date_str else None
        except ValueError:
            order_date = None

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
            return count_working_days(order_date, dt) if (dt and order_date) else None

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

        sport_year = (order_date.year if order_date.month >= 4 else order_date.year - 1) if order_date else datetime.now().year

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
            'days_ship_to_delivery': (
                round((_dt(delivery_act_ms) - _dt(ship_act_ms)).days)
                if _dt(ship_act_ms) and _dt(delivery_act_ms) else None
            ),
            'is_delayed':       is_delayed,
            'error_count':      error_count,
            'received_count':   received_count,
            'current_phase':    current_phase,
            'order_status':     order_status,
            'remake_qty':       order.get('remake_qty'),
            'order_jersey_qty': order.get('total_jersey_qty'),
        })

    return jsonify({'ok': True, 'orders': result})


# ── Settings ─────────────────────────────────────────────────────────────────

@app.route('/api/sync-clickup', methods=['POST'])
def api_sync_clickup():
    result = _sync_clickup_tasks()
    ok = 'error' not in result
    return jsonify({'ok': ok, **result})


@app.route('/settings', methods=['GET', 'POST'])
def settings_page():
    env_path = Path(__file__).parent / '.env'
    saved = False

    if request.method == 'POST':
        lines = []
        for key in ('CLICKUP_TOKEN', 'CLICKUP_LIST_ID', 'SLACK_WEBHOOK_URL', 'EMAIL_APP_PASSWORD', 'TRACK17_API_KEY', 'FEDEX_API_KEY', 'FEDEX_SECRET_KEY', 'UPS_CLIENT_ID', 'UPS_CLIENT_SECRET'):
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


INVENTORY_SHEET_ID   = '1rJTevdKr-8ONV56D0yvUBpvoaRlzYJZtyLqp_8R_9Js'
INVENTORY_LOG_FILE   = DATA_FOLDER / 'inventory_log.json'
INVENTORY_STOCK_FILE = DATA_FOLDER / 'inventory_stock.json'

_REVERSIBLE_SPORTS = {'Soccer', 'Bowling', '7V7'}


def calc_2026_price(gear_type: str, sport: str, season: str) -> float | None:
    """Return 2026 franchise sell price per unit. None = no known price."""
    gt = gear_type.lower().strip()
    sp = sport.strip()

    if 'jersey' in gt:
        return 18.50 if sp in _REVERSIBLE_SPORTS else 14.50
    if gt == 'hoodies':
        return 20.00
    if gt in ('cage jackets', 'cage jacket'):
        return 21.50
    if gt == 'windbreakers':
        return 21.50
    if gt == 'tees':
        return 11.00 if sp == 'Baseball' else 12.00
    if gt == 'caps':
        return 11.00
    if gt == 'rings':
        return 10.50
    if gt in ('hype chains', 'hype chain'):
        return 12.00
    if gt == 'park banner':
        return 140.00
    return None


def _load_inv_log() -> list:
    if not INVENTORY_LOG_FILE.exists():
        return []
    try:
        return json.loads(INVENTORY_LOG_FILE.read_text())
    except Exception:
        return []


def _save_inv_log(entries: list):
    INVENTORY_LOG_FILE.write_text(json.dumps(entries, indent=2))


def _load_inv_stock() -> dict:
    if not INVENTORY_STOCK_FILE.exists():
        return {}
    try:
        return json.loads(INVENTORY_STOCK_FILE.read_text())
    except Exception:
        return {}


def _save_inv_stock(data: dict):
    INVENTORY_STOCK_FILE.write_text(json.dumps(data, indent=2))


def _item_key(row: dict) -> str:
    return '|'.join([row.get('gear_type',''), row.get('name',''), row.get('sport',''),
                     row.get('code',''), row.get('size',''), row.get('color','')])


@app.route('/inventory')
def inventory():
    return render_template('inventory.html')


@app.route('/public/inventory')
def public_inventory():
    if not pub_authed():
        return redirect('/public/login?next=/public/inventory')
    return render_template('public_inventory.html')


@app.route('/api/inventory-data')
def api_inventory_data():
    api_key = os.getenv('GOOGLE_API_KEY', '')
    if not api_key:
        return jsonify({'ok': False, 'error': 'GOOGLE_API_KEY not configured'}), 500
    try:
        r = requests.get(
            f'https://sheets.googleapis.com/v4/spreadsheets/{INVENTORY_SHEET_ID}/values/HQ%20Inventory!A1:N300',
            params={'key': api_key},
            timeout=15,
        )
        r.raise_for_status()
        data = r.json()
    except Exception as e:
        return jsonify({'ok': False, 'error': str(e)}), 502

    rows = data.get('values', [])
    if len(rows) < 2:
        return jsonify({'ok': True, 'items': []})

    def _f(v):
        try:
            return float(str(v).replace(',', '').strip())
        except Exception:
            return 0.0

    local_stock = _load_inv_stock()

    items = []
    for idx, row in enumerate(rows[1:], start=2):
        while len(row) < 14:
            row.append('')
        gear_type     = row[0]
        sport         = row[2]
        season        = row[4]
        cost_per_unit = _f(row[11])
        p2026         = calc_2026_price(gear_type, sport, season)
        item = {
            'row':           idx,
            'gear_type':     gear_type,
            'name':          row[1],
            'sport':         sport,
            'code':          row[3],
            'season':        season,
            'status':        row[5],
            'color':         row[6],
            'size':          row[7],
            'stock_in':      _f(row[8]),
            'stock_out':     _f(row[9]),
            'current_stock': _f(row[10]),
            'cost_per_unit': cost_per_unit,
            'price_per_unit': p2026,
        }
        key = _item_key(item)
        if key in local_stock:
            saved = local_stock[key]
            item['stock_in']      = saved.get('stock_in', item['stock_in'])
            item['stock_out']     = saved.get('stock_out', item['stock_out'])
            item['current_stock'] = saved.get('current_stock', item['current_stock'])
            item['last_updated']  = saved.get('last_updated', '')
        cs = item['current_stock']
        item['total_cost']  = round(cs * cost_per_unit, 2)
        item['total_price'] = round(cs * p2026, 2) if p2026 is not None else 0
        items.append(item)
    return jsonify({'ok': True, 'items': items})


@app.route('/api/inventory-update', methods=['POST'])
def api_inventory_update():
    body          = request.get_json(force=True)
    item_key      = body.get('item_key')
    new_stock_in  = body.get('stock_in')
    new_stock_out = body.get('stock_out')
    new_current   = body.get('current_stock')
    cst_now       = body.get('cst_now') or _fmt_cst()
    log_entry     = body.get('log_entry', {})

    if not item_key:
        return jsonify({'ok': False, 'error': 'item_key required'}), 400

    stock = _load_inv_stock()
    stock[item_key] = {
        'stock_in':      new_stock_in,
        'stock_out':     new_stock_out,
        'current_stock': new_current,
        'last_updated':  cst_now,
    }
    _save_inv_stock(stock)

    if log_entry:
        entries = _load_inv_log()
        log_entry['id']          = str(uuid.uuid4())
        log_entry['server_time'] = _fmt_cst()
        entries.insert(0, log_entry)
        _save_inv_log(entries)

    return jsonify({'ok': True})


@app.route('/api/inventory-log', methods=['GET'])
def get_inventory_log():
    entries = _load_inv_log()
    limit = request.args.get('limit', type=int)
    if limit:
        entries = entries[:limit]
    return jsonify({'ok': True, 'entries': entries})


@app.route('/api/inventory-log', methods=['POST'])
def add_inventory_log():
    entry = request.get_json(force=True)
    entries = _load_inv_log()
    entry['id']          = str(uuid.uuid4())
    entry['server_time'] = _fmt_cst()
    entries.insert(0, entry)
    _save_inv_log(entries)
    return jsonify({'ok': True, 'id': entry['id']})


@app.route('/packing-verifier')
def packing_verifier():
    orders = load_display_orders()
    return render_template('packing_verifier.html', orders=orders)


@app.route('/api/packing-verifier-data/<order_id>')
def api_packing_verifier_data(order_id):
    orders = load_orders()
    order  = next((o for o in orders if o.get('id') == order_id), None)
    if not order:
        return jsonify({'ok': False, 'error': 'Order not found'}), 404

    gear_tasks       = order.get('gear_tasks', {})
    manual_shipments = order.get('manual_shipments', [])
    tracking_statuses = order.get('tracking_statuses', {})

    # Index shipments by gear_type
    shipped_by_gear = {}
    for ms in manual_shipments:
        g = ms.get('gear_type', '')
        shipped_by_gear.setdefault(g, [])
        shipped_by_gear[g].append(ms)

    gear_checks = []
    for gear in gear_tasks:
        shipments = shipped_by_gear.get(gear, [])
        shipped   = bool(shipments)
        tns       = list(dict.fromkeys(s.get('tracking_number', '') for s in shipments if s.get('tracking_number')))
        boxes     = sum(s.get('box_count', 0) or 0 for s in shipments)
        ship_date = min((s.get('ship_date', '') for s in shipments if s.get('ship_date')), default='')
        tags      = list(dict.fromkeys(
            tracking_statuses.get(tn, {}).get('tag', 'Pending')
            for tn in tns
        ))
        gear_checks.append({
            'gear':             gear,
            'shipped':          shipped,
            'tracking_numbers': tns,
            'box_count':        boxes if shipped else 0,
            'ship_date':        ship_date,
            'delivery_tags':    tags,
        })

    missing     = [g['gear'] for g in gear_checks if not g['shipped']]
    all_shipped = len(missing) == 0 and len(gear_checks) > 0

    return jsonify({
        'ok':          True,
        'all_shipped': all_shipped,
        'missing':     missing,
        'gear_checks': gear_checks,
    })


@app.route('/api/generate-packing-list/<order_id>', methods=['POST'])
def api_generate_packing_list(order_id):
    import subprocess, tempfile, html as html_lib

    orders = load_orders()
    order  = next((o for o in orders if o.get('id') == order_id), None)
    if not order:
        return jsonify({'ok': False, 'error': 'Order not found'}), 404

    market   = order.get('market', '')
    sport    = order.get('sport', 'BSBL')
    season   = order.get('season', 'Summer')
    num      = order.get('num_for_code', order.get('order_num', '1'))
    order_date = order.get('order_date', '')
    manual_shipments  = order.get('manual_shipments', [])
    packing_list_rows = order.get('packing_list', [])
    tracking_statuses = order.get('tracking_statuses', {})
    gear_tasks        = order.get('gear_tasks', {})

    # Collect unique tracking numbers in order of first appearance
    seen_tns = []
    for ms in manual_shipments:
        tn = ms.get('tracking_number', '')
        if tn and tn not in seen_tns:
            seen_tns.append(tn)
    # Also pick up any TNs only in packing_list
    for row in packing_list_rows:
        tn = row.get('tracking_number', '')
        if tn and tn not in seen_tns:
            seen_tns.append(tn)

    # Build summary rows: one row per gear type (ordered)
    ordered_gears = list(gear_tasks.keys()) if gear_tasks else list(
        dict.fromkeys(ms.get('gear_type', '') for ms in manual_shipments)
    )

    shipped_by_gear = {}
    for ms in manual_shipments:
        g = ms.get('gear_type', '')
        shipped_by_gear.setdefault(g, [])
        shipped_by_gear[g].append(ms)

    pl_by_tn = {}
    for row in packing_list_rows:
        tn = row.get('tracking_number', '')
        pl_by_tn.setdefault(tn, [])
        pl_by_tn[tn].append(row)

    def esc(v):
        return html_lib.escape(str(v or ''))

    # ── Order Summary Table rows ──
    summary_rows_html = ''
    for i, gear in enumerate(ordered_gears):
        shipments = shipped_by_gear.get(gear, [])
        shipped   = bool(shipments)
        tns       = list(dict.fromkeys(s.get('tracking_number', '') for s in shipments if s.get('tracking_number')))
        boxes     = sum(s.get('box_count', 0) or 0 for s in shipments)
        ship_dates = sorted(set(s.get('ship_date', '') for s in shipments if s.get('ship_date')))

        # PL contents detail for this gear across all TNs
        pl_gears = [r for r in packing_list_rows if r.get('gear_type') == gear]
        detail = ''
        if pl_gears:
            box_labels = []
            for r in pl_gears:
                bnum = r.get('box_num', '')
                cnt  = r.get('contents', '')
                if bnum and cnt and cnt.lower() != 'see attached':
                    box_labels.append(f"Box {bnum}: {cnt[:60]}{'…' if len(cnt) > 60 else ''}")
                elif bnum:
                    box_labels.append(f"Box {bnum}")
            detail = ' &nbsp;·&nbsp; '.join(box_labels[:4])
        if not detail and shipments:
            notes = shipments[0].get('notes', '')
            detail = esc(notes[:80]) if notes else ''

        # Delivery status from tracking_statuses
        tags = [tracking_statuses.get(tn, {}).get('tag', '') for tn in tns]
        if all(t == 'Delivered' for t in tags if t):
            status_cell = '<span style="background:#dcfce7; color:#166534; border-radius:5px; padding:2px 8px; font-size:11px; font-weight:700;">✓ Delivered</span>'
        elif any(t in ('InTransit', 'OutForDelivery') for t in tags):
            status_cell = '<span style="background:#fef9c3; color:#854d0e; border-radius:5px; padding:2px 8px; font-size:11px; font-weight:700;">⏳ In Transit</span>'
        elif not shipped:
            status_cell = '<span style="background:#fee2e2; color:#991b1b; border-radius:5px; padding:2px 8px; font-size:11px; font-weight:700;">✗ Not Shipped</span>'
        else:
            status_cell = '<span style="background:#dbeafe; color:#1e40af; border-radius:5px; padding:2px 8px; font-size:11px; font-weight:700;">Shipped</span>'

        tn_html = '<br>'.join(esc(tn) for tn in tns) if tns else '—'
        row_bg  = ' background:#fafafa;' if i % 2 == 1 else ''
        ship_str = ', '.join(ship_dates) if ship_dates else '—'

        summary_rows_html += f'''
      <tr style="border-bottom:1px solid #f3f4f6;{row_bg}">
        <td style="padding:9px 14px; font-weight:600; color:#374151;">{esc(gear)}</td>
        <td style="padding:9px 14px; text-align:center; color:#374151;">{esc(str(boxes)) if shipped else '—'}</td>
        <td style="padding:9px 14px; font-size:11px; color:#6b7280;">{detail or '—'}</td>
        <td style="padding:9px 14px; font-size:11px; color:#1a2b6b; font-family:monospace;">{tn_html}</td>
        <td style="padding:9px 14px; font-size:11px; color:#6b7280;">{esc(ship_str)}</td>
        <td style="padding:9px 14px; text-align:center;">{status_cell}</td>
      </tr>'''

    # ── Per-tracking-number sections ──
    tn_sections_html = ''
    for tn in seen_tns:
        pl_rows = pl_by_tn.get(tn, [])
        # Find shipment meta for this TN (weight, declared, ship_date, boxes)
        ms_for_tn = [s for s in manual_shipments if s.get('tracking_number') == tn]
        ts_info   = tracking_statuses.get(tn, {})
        tag       = ts_info.get('tag', '')
        delivered_at = ts_info.get('delivered_at', '')

        if tag == 'Delivered' and delivered_at:
            status_str = f'<strong style="color:#86efac;">✓ Delivered {delivered_at}</strong>'
        elif tag in ('InTransit', 'OutForDelivery'):
            loc = ts_info.get('location', '') or ts_info.get('notes', '')
            status_str = f'<strong style="color:#fcd34d;">⏳ In Transit{" — " + esc(loc) if loc else ""}</strong>'
        else:
            shipped_gears = [s.get('gear_type', '') for s in ms_for_tn]
            status_str = esc(', '.join(g for g in shipped_gears if g) or 'Shipped')

        ship_date_str = ''
        all_ship_dates = sorted(set(s.get('ship_date', '') for s in ms_for_tn if s.get('ship_date')))
        if all_ship_dates:
            ship_date_str = f'Shipped {" / ".join(all_ship_dates)} &nbsp;·&nbsp; '

        # Total physical boxes for this TN
        total_boxes = sum(s.get('box_count', 0) or 0 for s in ms_for_tn)
        if total_boxes == 0 and pl_rows:
            box_nums = set(r.get('box_num') for r in pl_rows if r.get('box_num'))
            total_boxes = len(box_nums)
        weight_vals = list(set(s.get('weight', '') for s in ms_for_tn if s.get('weight')))
        weight_str  = ', '.join(str(w).replace('KG', '').strip() + ' KG' for w in weight_vals if w) or ''

        # Gears in this TN
        gears_in_tn = list(dict.fromkeys(
            s.get('gear_type', '') for s in ms_for_tn if s.get('gear_type')
        ))

        # Box contents summary line from packing list
        contents_summary = ''
        if pl_rows:
            parts = []
            for r in pl_rows:
                bnum = r.get('box_num', '')
                gear = r.get('gear_type', '')
                cnt  = r.get('contents', '')
                if cnt and cnt.lower() not in ('see attached', ''):
                    parts.append(f'<strong>{esc(gear)} Box {bnum}</strong> — {esc(cnt[:120])}')
                elif bnum:
                    parts.append(f'<strong>{esc(gear)} Box {bnum}</strong>')
            contents_summary = ' &nbsp;&nbsp; '.join(parts[:8])

        # Contents table rows (team-level from packing list, grouped by team)
        teams = {}
        for r in pl_rows:
            gear_t = r.get('gear_type', '')
            cnt    = r.get('contents', '')
            bnum   = r.get('box_num', '')
            if cnt and cnt.lower() not in ('see attached', ''):
                for team in cnt.split(','):
                    team = team.strip().title()
                    if team:
                        teams.setdefault(team, {})[gear_t] = bnum

        table_rows_html = ''
        if teams:
            sorted_teams = sorted(teams.keys())
            col_gears    = list(dict.fromkeys(
                r.get('gear_type', '') for r in pl_rows if r.get('gear_type')
            ))
            gear_headers = ''.join(
                f'<th style="padding:9px 14px; text-align:center; font-size:10px; font-weight:700; letter-spacing:1px; color:#9ca3af; text-transform:uppercase; width:20%;">{esc(g)}</th>'
                for g in col_gears
            )
            for ti, team in enumerate(sorted_teams):
                row_bg2 = ' background:#fafafa;' if ti % 2 == 1 else ''
                cols = ''
                for g in col_gears:
                    bnum = teams[team].get(g)
                    if bnum:
                        cols += f'<td style="padding:7px 14px; text-align:center; font-size:11.5px; color:#374151; font-weight:600;">Box {bnum}</td>'
                    else:
                        cols += '<td style="padding:7px 14px; text-align:center; color:#d1d5db; font-size:13px;">—</td>'
                table_rows_html += f'<tr style="border-bottom:1px solid #f3f4f6;{row_bg2}"><td style="padding:7px 14px; color:#374151; font-weight:500;">{esc(team)}</td>{cols}</tr>'

            table_html = f'''
    <table style="width:100%; border-collapse:collapse; font-size:12.5px; border:1px solid #e5e7eb; border-top:none;">
      <thead>
        <tr style="background:#f8fafc; border-bottom:2px solid #e5e7eb;">
          <th style="padding:9px 14px; text-align:left; font-size:10px; font-weight:700; letter-spacing:1px; color:#9ca3af; text-transform:uppercase; width:35%;">School / Team</th>
          {gear_headers}
        </tr>
      </thead>
      <tbody>{table_rows_html}</tbody>
    </table>'''
        elif gears_in_tn:
            table_html = f'<div style="background:#f9fafb; border:1px solid #e5e7eb; border-top:none; border-radius:0 0 10px 10px; padding:12px 14px; font-size:12.5px; color:#374151;">{esc(", ".join(gears_in_tn))}</div>'
        else:
            table_html = ''

        meta_parts = [ship_date_str, status_str]
        if total_boxes:
            meta_parts.append(f'&nbsp;·&nbsp; {total_boxes} box{"es" if total_boxes != 1 else ""}')
        if weight_str:
            meta_parts.append(f'&nbsp;·&nbsp; {esc(weight_str)}')

        contents_bar = ''
        if contents_summary:
            contents_bar = f'<div style="background:#eff6ff; border-left:3px solid #3b82f6; padding:9px 14px; font-size:12px; color:#1e40af; border-right:1px solid #e5e7eb; border-bottom:1px solid #e5e7eb; line-height:2;">{contents_summary}</div>'

        tn_sections_html += f'''
  <div class="tracking-section" style="margin-bottom:32px;">
    <div style="background:#1a2b6b; color:#fff; border-radius:10px 10px 0 0; padding:12px 16px; display:flex; justify-content:space-between; align-items:center;">
      <div>
        <span style="font-size:10px; font-weight:700; letter-spacing:1.5px; opacity:0.7; text-transform:uppercase;">Tracking</span>
        <span style="font-family:monospace; font-size:15px; font-weight:700; margin-left:8px;">{esc(tn)}</span>
      </div>
      <div style="font-size:12px; opacity:0.8; text-align:right;">
        {''.join(meta_parts)}
      </div>
    </div>
    {contents_bar}
    {table_html}
  </div>'''

    # Format order date nicely
    try:
        from datetime import datetime as _dt
        od = _dt.strptime(order_date, '%Y-%m-%d').strftime('%B %-d, %Y') if order_date else ''
    except Exception:
        od = order_date

    # Collect unique ship dates for the header
    all_dates = sorted(set(s.get('ship_date', '') for s in manual_shipments if s.get('ship_date')))
    try:
        from datetime import datetime as _dt2
        date_strs = [_dt2.strptime(d, '%Y-%m-%d').strftime('%B %-d, %Y') for d in all_dates]
    except Exception:
        date_strs = all_dates
    ship_dates_display = ' – '.join(date_strs) if date_strs else ''

    title = f'{esc(market)} {esc(sport)} {esc(season)} Order #{esc(str(num))}'

    html_content = f'''<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>{title} — Packing List</title>
  <script src="https://cdn.tailwindcss.com"></script>
  <style>
    body {{ font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', sans-serif; }}
    @media print {{
      .no-print {{ display: none !important; }}
      body {{ background: white !important; padding: 0 !important; }}
      .doc {{ box-shadow: none !important; border-radius: 0 !important; margin: 0 !important; }}
      .tracking-section {{ page-break-inside: avoid; }}
    }}
    td, th {{ vertical-align: middle; }}
  </style>
</head>
<body class="bg-gray-100 min-h-screen py-8 px-4">

<div class="no-print max-w-5xl mx-auto mb-5 bg-white rounded-xl shadow-sm border border-gray-100 p-4 flex justify-end">
  <button onclick="window.print()"
          class="bg-[#8B1A1A] text-white text-sm font-semibold px-5 py-2.5 rounded-lg hover:bg-[#7a1515] transition">
    Print / Save PDF
  </button>
</div>

<div class="doc max-w-5xl mx-auto bg-white rounded-2xl shadow-lg p-10 mb-8">

  <!-- Header -->
  <div style="display:flex; justify-content:space-between; align-items:flex-start; margin-bottom:22px;">
    <div>
      <div style="font-size:22px; font-weight:800; color:#1a2b6b; letter-spacing:-0.5px;">Middle School Matchup</div>
      <div style="font-size:12px; color:#6b7280; margin-top:3px;">gear@middleschoolmatchup.com</div>
    </div>
    <div style="text-align:right;">
      <div style="font-size:26px; font-weight:800; color:#8B1A1A; letter-spacing:1px;">PACKING LIST</div>
      <div style="margin-top:5px; line-height:1.9;">
        <div style="font-size:13px; font-weight:700; color:#1a2b6b;">{title}</div>
        {'<div style="font-size:12px; color:#6b7280;">Order Date: <strong style="color:#111827;">' + od + '</strong></div>' if od else ''}
        {'<div style="font-size:12px; color:#6b7280;">Ship Date: <strong style="color:#111827;">' + esc(ship_dates_display) + '</strong></div>' if ship_dates_display else ''}
      </div>
    </div>
  </div>

  <div style="border-top:2px solid #1a2b6b; margin-bottom:24px;"></div>

  <!-- Order Summary -->
  <div style="font-size:10px; font-weight:700; letter-spacing:1.5px; color:#9ca3af; text-transform:uppercase; margin-bottom:12px;">Order Summary</div>

  <table style="width:100%; border-collapse:collapse; font-size:12.5px; margin-bottom:24px; border:1px solid #e5e7eb; border-radius:10px; overflow:hidden;">
    <thead>
      <tr style="background:#1a2b6b; color:#fff;">
        <th style="padding:10px 14px; text-align:left; font-weight:700; letter-spacing:0.5px; width:18%;">GEAR</th>
        <th style="padding:10px 14px; text-align:center; font-weight:700; letter-spacing:0.5px; width:8%;">BOXES</th>
        <th style="padding:10px 14px; text-align:left; font-weight:700; letter-spacing:0.5px;">CONTENTS</th>
        <th style="padding:10px 14px; text-align:left; font-weight:700; letter-spacing:0.5px; width:22%;">TRACKING</th>
        <th style="padding:10px 14px; text-align:left; font-weight:700; letter-spacing:0.5px; width:12%;">SHIP DATE</th>
        <th style="padding:10px 14px; text-align:center; font-weight:700; letter-spacing:0.5px; width:10%;">STATUS</th>
      </tr>
    </thead>
    <tbody>{summary_rows_html}</tbody>
  </table>

  <!-- Tracking Sections -->
  {tn_sections_html}

  <!-- Footer -->
  <div style="border-top:1px solid #e5e7eb; padding-top:14px; font-size:11px; color:#9ca3af; text-align:center;">
    Middle School Matchup &nbsp;·&nbsp; gear@middleschoolmatchup.com
  </div>

</div>
</body>
</html>'''

    # Save to temp file
    tmp_path = Path(tempfile.gettempdir()) / f'{order_id}-packing-list.html'
    tmp_path.write_text(html_content, encoding='utf-8')

    # Build order key for register_packing_list.py
    order_key   = f'{market} {sport} {season} {num}'
    label       = f'{market} {sport} {season} Order #{num}'
    script_path = Path(__file__).parent / 'register_packing_list.py'
    venv_python = Path(__file__).parent / '.venv' / 'bin' / 'python3'
    python_exe  = str(venv_python) if venv_python.exists() else 'python3'

    result = subprocess.run(
        [python_exe, str(script_path), str(tmp_path), order_key, label],
        capture_output=True, text=True, timeout=60,
        cwd=str(Path(__file__).parent),
    )

    if result.returncode != 0:
        app.logger.error('register_packing_list failed: %s', result.stderr)
        return jsonify({'ok': False, 'error': result.stderr.strip() or 'PDF generation failed'}), 500

    # Re-read orders to find the registered doc
    orders_fresh = load_orders()
    order_fresh  = next((o for o in orders_fresh if o.get('id') == order_id), None)
    docs         = order_fresh.get('packing_docs', []) if order_fresh else []
    doc          = next((d for d in reversed(docs) if d.get('label') == label), docs[-1] if docs else None)

    if not doc:
        return jsonify({'ok': False, 'error': 'PDF created but not found in registry'}), 500

    pdf_name  = doc['filename']
    view_url  = f'/packing-doc/{order_id}/{pdf_name}'
    dl_url    = f'/packing-doc/{order_id}/{pdf_name}'

    return jsonify({'ok': True, 'view_url': view_url, 'download_url': dl_url, 'filename': pdf_name})


if __name__ == '__main__':
    app.run(debug=True, port=5003)
