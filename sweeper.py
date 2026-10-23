"""
Daily Gear sweeper — finds ClickUp gear tasks (2025 + 2026, open and closed,
parents and subtasks) whose shipping data is out of date and fixes what the
carrier can prove. Runs on the MSM server at 06:00 America/Chicago
(systemd timer msm-gear-sweeper.timer).

What it fixes (only blank / lagging fields — never renames tasks, never
downgrades a phase, never touches Canceled/Critical/Storage statuses):
  - Ship Date [Actual] blank   → carrier pickup date
  - Delivery Date [Actual] blank and every TN delivered → carrier delivery date
  - Phase behind the carrier   → Shipped (picked up) / Delivered (all TNs delivered)
  - Order Status blank on a delivered task → Ahead / On Schedule / Behind
  - Ship Date (estimate) passed, still not shipped, status Ahead/On Schedule → Behind
Things it can't prove (no tracking #, no estimate) are listed in the report only.

Usage:  python sweeper.py [--dry-run] [--no-email]
"""
from __future__ import annotations

import os
import re
import sys
import json
import time
from datetime import datetime, date, timedelta
from zoneinfo import ZoneInfo

os.environ.setdefault('TRACKER_BACKGROUND', '0')   # don't start the app's pollers
os.chdir(os.path.dirname(os.path.abspath(__file__)))

import requests
import app
from app import (CF_TRACKING, CF_PHASE, CF_ORDER_STATUS, CF_SHIP_DATE, CF_SHIP_DATE_ACT,
                 CF_DELIVERY_DATE, CF_DELIVERY_DATE_ACT)

CT = ZoneInfo('America/Chicago')
SINCE_MS = int(datetime(2025, 1, 1, tzinfo=CT).timestamp() * 1000)
STATE_FILE = os.path.join('data', 'sweeper_state.json')
LOG_FILE = os.path.join('data', 'sweeper_log.jsonl')

PHASES = ['Submitted', 'Proofs', 'Production', 'Shipped', 'Delivered', 'Counted']
STATUSES = ['Ahead', 'On Schedule', 'Behind', 'Canceled', 'Critical', 'Storage']
LOCKED_STATUSES = {3, 4, 5}   # Canceled / Critical / Storage — set by hand, leave alone
TN_RE = re.compile(r'^(\d{12,22}|1Z[0-9A-Z]{16})$')

DRY = '--dry-run' in sys.argv
API = 'https://api.clickup.com/api/v2'


def _cu(method, path, **kw):
    """ClickUp call with 429 back-off."""
    for attempt in range(5):
        r = requests.request(method, API + path, headers=app.clickup_headers(), timeout=30, **kw)
        if r.status_code == 429:
            time.sleep(int(r.headers.get('X-RateLimit-Reset', time.time() + 10)) - time.time() + 1
                       if r.headers.get('X-RateLimit-Reset') else 10)
            continue
        r.raise_for_status()
        return r.json()
    r.raise_for_status()


def fetch_tasks() -> list:
    list_id = app.get_settings()['clickup_list_id']
    out, page = [], 0
    while True:
        batch = _cu('GET', f'/list/{list_id}/task', params={
            'page': page, 'include_closed': 'true', 'subtasks': 'true',
            'date_created_gt': SINCE_MS})['tasks']
        out += batch
        if len(batch) < 100:
            return out
        page += 1


def _cf(task, fid):
    for f in task.get('custom_fields') or []:
        if f['id'] == fid:
            v = f.get('value')
            if v in (None, ''):
                return None
            if f.get('type') == 'drop_down':
                if isinstance(v, int):
                    return v
                for o in f['type_config']['options']:   # some API responses give the option id
                    if o['id'] == v:
                        return o['orderindex']
            return v
    return None


def _day(ms) -> date | None:
    return datetime.fromtimestamp(int(ms) / 1000, CT).date() if ms else None


def _noon_ms(d: date) -> int:
    """Mid-day Central so the date reads the same in every US timezone."""
    return int(datetime(d.year, d.month, d.day, 12, tzinfo=CT).timestamp() * 1000)


def _iso_day(s: str) -> date | None:
    if not s:
        return None
    try:
        return datetime.fromisoformat(s.replace('Z', '+00:00')).astimezone(CT).date()
    except ValueError:
        return None


_carrier_cache: dict = {}


def carrier_info(tn: str) -> dict | None:
    """{'picked_up': date|None, 'delivered': date|None, 'source': str} or None if unknown."""
    if tn in _carrier_cache:
        return _carrier_cache[tn]
    info = None
    try:
        if app._detect_carrier(tn) == 'fedex':
            token = app._get_fedex_token()
            r = requests.post(
                'https://apis.fedex.com/track/v1/trackingnumbers',
                json={'includeDetailedScans': False,
                      'trackingInfo': [{'trackingNumberInfo': {'trackingNumber': tn}}]},
                headers={'Authorization': f'Bearer {token}', 'Content-Type': 'application/json'},
                timeout=20)
            tr = (r.json().get('output', {}).get('completeTrackResults', [{}])[0]
                  .get('trackResults', [{}])[0])
            if not tr.get('error'):
                dts = {d['type']: d['dateTime'] for d in tr.get('dateAndTimes', [])}
                code = (tr.get('latestStatusDetail') or {}).get('code', '')
                info = {'picked_up': _iso_day(dts.get('ACTUAL_PICKUP') or dts.get('SHIP')),
                        'delivered': _iso_day(dts.get('ACTUAL_DELIVERY')) if code == 'DL' else None,
                        'source': 'FedEx'}
        else:
            res = app._fetch_17track(tn)   # UPS/USPS: 17track gives latest event only
            if res and res.get('tag') == 'Delivered':
                info = {'picked_up': None, 'delivered': _iso_day(res.get('event_time')),
                        'source': '17track'}
            elif res and res.get('tag') not in (None, '', 'Pending', 'InfoReceived'):
                info = {'picked_up': None, 'delivered': None, 'in_transit': True, 'source': '17track'}
    except Exception as e:
        print(f'  carrier lookup failed for {tn}: {e}')
    _carrier_cache[tn] = info
    return info


def sweep_task(task: dict, today: date) -> tuple[dict, list]:
    """Return (changes {field_id: (value, label)}, attention)."""
    name = task['name']
    phase = _cf(task, CF_PHASE)
    status = _cf(task, CF_ORDER_STATUS)
    ship_est, ship_act = _cf(task, CF_SHIP_DATE), _cf(task, CF_SHIP_DATE_ACT)
    del_est, del_act = _cf(task, CF_DELIVERY_DATE), _cf(task, CF_DELIVERY_DATE_ACT)
    raw_tn = (_cf(task, CF_TRACKING) or '').strip()
    tns = [t.strip().upper() for t in re.split(r'[,\s/;]+', raw_tn) if TN_RE.match(t.strip().upper())]

    changes, attention = {}, []
    shipped_or_later = phase is not None and phase >= 3
    delivered_or_later = phase is not None and phase >= 4

    if tns and (not ship_act or not del_act or not delivered_or_later):
        infos = [carrier_info(t) for t in tns]
        known = [i for i in infos if i]
        pickups = [i['picked_up'] for i in known if i.get('picked_up')]
        deliveries = [i['delivered'] for i in known if i.get('delivered')]
        all_delivered = len(deliveries) == len(tns)
        moving = bool(pickups or deliveries or any(i.get('in_transit') for i in known))
        src = '/'.join(sorted({i['source'] for i in known})) or 'carrier'

        new_ship = None
        if not ship_act and pickups:
            new_ship = min(pickups)
            changes[CF_SHIP_DATE_ACT] = (_noon_ms(new_ship), f'Ship Date [Actual] → {new_ship} ({src} pickup)')
        new_del = None
        if not del_act and all_delivered:
            new_del = max(deliveries)
            changes[CF_DELIVERY_DATE_ACT] = (_noon_ms(new_del), f'Delivery Date [Actual] → {new_del} ({src})')

        if all_delivered and (phase is None or phase < 4):
            changes[CF_PHASE] = (4, f'Phase {PHASES[phase] if phase is not None else "blank"} → Delivered')
            if status not in LOCKED_STATUSES:
                act = new_del or _day(del_act)
                st = app._delivery_order_status(del_est, _noon_ms(act) if act else None)
                if st != status:
                    changes[CF_ORDER_STATUS] = (st, f'Order Status → {STATUSES[st]} (delivered vs estimate)')
        elif moving and (phase is None or phase < 3):
            changes[CF_PHASE] = (3, f'Phase {PHASES[phase] if phase is not None else "blank"} → Shipped')
            act = new_ship or _day(ship_act)
            if status not in LOCKED_STATUSES and act:
                st = app._ship_order_status(_noon_ms(act), ship_est)
                if st != status:
                    changes[CF_ORDER_STATUS] = (st, f'Order Status → {STATUSES[st]} (shipped vs estimate)')

        if not known:
            attention.append(f'tracking {raw_tn} — carrier has no data')
        elif not ship_act and not new_ship and shipped_or_later and not deliveries:
            attention.append('no Ship Date [Actual] and carrier has no pickup date')

    phase_after = changes.get(CF_PHASE, (phase,))[0]
    delivered_after = phase_after is not None and phase_after >= 4

    # Delivered but Order Status blank → compute from the dates we have
    if delivered_after and status is None and CF_ORDER_STATUS not in changes:
        act = del_act or (changes.get(CF_DELIVERY_DATE_ACT) or (None,))[0]
        if act and del_est:
            st = app._delivery_order_status(del_est, act)
            changes[CF_ORDER_STATUS] = (st, f'Order Status → {STATUSES[st]} (delivered vs estimate)')

    # Ship estimate passed and still not shipped → Behind
    if (phase_after is not None and phase_after < 3 and not ship_act
            and CF_SHIP_DATE_ACT not in changes and ship_est and _day(ship_est) < today
            and (status is None or status in (0, 1)) and CF_ORDER_STATUS not in changes):
        changes[CF_ORDER_STATUS] = (2, f'Order Status → Behind (Ship Date {_day(ship_est)} passed, not shipped)')

    # Things the sweeper can't prove — report only
    if delivered_after and not tns:
        missing = [lbl for fid, lbl in ((CF_SHIP_DATE_ACT, 'Ship Date [Actual]'),
                                        (CF_DELIVERY_DATE_ACT, 'Delivery Date [Actual]'))
                   if not _cf(task, fid) and fid not in changes]
        if missing:
            attention.append(f'Delivered, no tracking # — missing {", ".join(missing)}')
    if delivered_after and tns and not del_act and CF_DELIVERY_DATE_ACT not in changes:
        attention.append('Delivered but carrier has no delivery date for every TN')
    if phase_after is not None and phase_after < 3 and not ship_est:
        attention.append(f'{PHASES[phase_after]} with no Ship Date estimate')
    return changes, attention


def main():
    today = datetime.now(CT).date()
    started = datetime.now(CT)
    tasks = fetch_tasks()
    print(f'{started:%Y-%m-%d %H:%M} CT — {len(tasks)} gear tasks since 2025-01-01'
          f'{" (DRY RUN)" if DRY else ""}')

    fixed, attention, errors = [], [], []
    for t in tasks:
        try:
            changes, att = sweep_task(t, today)
        except Exception as e:
            errors.append(f'{t["name"]}: {e}')
            continue
        for a in att:
            attention.append(f'{t["name"]} — {a}')
        if not changes:
            continue
        labels = [lbl for _, lbl in changes.values()]
        print(f'  FIX {t["name"]} ({t["id"]}): ' + '; '.join(labels))
        if DRY:
            fixed.append((t, labels))
            continue
        try:
            # Dates before phase/status so the app's dashboards never see a half-updated task
            order = [CF_SHIP_DATE_ACT, CF_DELIVERY_DATE_ACT, CF_PHASE, CF_ORDER_STATUS]
            for fid in sorted(changes, key=order.index):
                _cu('POST', f'/task/{t["id"]}/field/{fid}', json={'value': changes[fid][0]})
            _cu('POST', f'/task/{t["id"]}/comment', json={
                'comment_text': '🧹 Daily gear sweeper updated this task:\n' + '\n'.join(f'• {l}' for l in labels),
                'notify_all': False})
            fixed.append((t, labels))
        except Exception as e:
            errors.append(f'{t["name"]}: {e}')

    for a in attention:
        print(f'  ATTN {a}')
    for e in errors:
        print(f'  ERR  {e}')
    print(f'Done: {len(fixed)} fixed, {len(attention)} need attention, {len(errors)} errors')

    if DRY:
        return
    try:
        prev = json.load(open(STATE_FILE))
    except Exception:
        prev = {}
    new_attention = sorted(set(attention) - set(prev.get('attention', [])))
    json.dump({'last_run': started.isoformat(), 'attention': sorted(set(attention))},
              open(STATE_FILE, 'w'), indent=1)
    with open(LOG_FILE, 'a') as f:
        f.write(json.dumps({'run': started.isoformat(), 'tasks': len(tasks),
                            'fixed': [{'id': t['id'], 'name': t['name'], 'changes': l} for t, l in fixed],
                            'attention': attention, 'errors': errors}) + '\n')

    if '--no-email' in sys.argv or not (fixed or new_attention or errors):
        return
    lines = [f'Gear sweeper — {started:%a %b %d, %Y %I:%M %p} CT', f'{len(tasks)} ClickUp gear tasks checked (2025 + 2026).', '']
    if fixed:
        lines.append(f'FIXED ({len(fixed)}):')
        for t, labels in fixed:
            lines.append(f'• {t["name"]} — {t["url"]}')
            lines += [f'    {l}' for l in labels]
        lines.append('')
    if new_attention:
        lines.append(f'NEW — needs a human ({len(new_attention)}):')
        lines += [f'• {a}' for a in new_attention]
        lines.append('')
    if errors:
        lines.append(f'ERRORS ({len(errors)}):')
        lines += [f'• {e}' for e in errors]
        lines.append('')
    lines.append(f'Still open from earlier runs: {len(set(attention)) - len(new_attention)} items '
                 f'(see data/sweeper_state.json on the server).')
    err = app.send_email_notification(f'🧹 Gear sweeper: {len(fixed)} fixed, {len(new_attention)} new to check',
                                      '\n'.join(lines))
    if err:
        print(f'Email failed: {err}')


if __name__ == '__main__':
    main()
