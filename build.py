#!/usr/bin/env python3
"""Fetch the latest RAN1 schedule docx from 3GPP FTP and regenerate index.html.

Uses Python stdlib only — no pip install needed.
"""
import io
import json
import re
import sys
import urllib.parse
import urllib.request
import xml.etree.ElementTree as ET
import zipfile
from datetime import datetime, timezone

FTP_DIR = "https://www.3gpp.org/ftp/Meetings_3GPP_SYNC/RAN1/Inbox/Chair_notes/"

NS = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}
NS_W = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'

CATEGORIES = {'6GR', 'R20', 'AI 7/8', 'AI 8', 'AI 9', 'AI/ML', 'Maintenance', 'MNTC'}
HOSTS = {'Hiroki', 'Sorour', 'Sorouri', 'Xiaodong'}
CAT_MAP = {
    '6GR': '6GR', 'R20': 'R20',
    'AI 7/8': 'AI78', 'AI 8': 'AI78', 'AI 9': 'AI78', 'AI/ML': 'AI78',
    'Maintenance': 'MAINT', 'MNTC': 'MAINT',
}
PERIODS_BY_INDEX = {
    0: ('08:30', '10:30'),
    1: ('11:00', '13:00'),
    2: ('14:30', '16:30'),
    3: ('17:00', '19:30'),
}
PAT_WITH_DUR = re.compile(r'^(.+?)\s*\(\s*(\d+)\s*(?:min)?\s*\)?\s*\.?$')

# -------- fetch --------

def fetch(url):
    req = urllib.request.Request(
        url, headers={'User-Agent': 'Mozilla/5.0 (RAN1 schedule builder)'})
    with urllib.request.urlopen(req, timeout=120) as r:
        return r.read()

def find_latest_schedule_url():
    """Scan the FTP directory listing for the latest '...online and offline schedules - vNN.docx'."""
    html = fetch(FTP_DIR).decode('utf-8', errors='replace')
    link_re = re.compile(r'href="(https?://[^"]+\.docx)"', re.IGNORECASE)
    candidates = []
    for m in link_re.finditer(html):
        url = m.group(1)
        decoded = urllib.parse.unquote(url).lower()
        if 'online and offline schedules' in decoded:
            vm = re.search(r'v(\d+)\.docx', decoded)
            ver = int(vm.group(1)) if vm else -1
            candidates.append((ver, url))
    if not candidates:
        return None
    candidates.sort()
    return candidates[-1][1]

# -------- docx parsing --------

def to_min(hhmm):
    h, m = hhmm.split(':')
    return int(h) * 60 + int(m)

def to_hhmm(m):
    return f"{m // 60:02d}:{m % 60:02d}"

def cell_info(cell):
    span = 1
    gs = cell.find('.//w:gridSpan', NS)
    if gs is not None:
        span = int(gs.get(NS_W + 'val', '1'))
    lines = []
    for p in cell.findall('.//w:p', NS):
        runs = [t.text for t in p.findall('.//w:t', NS) if t.text]
        line = ''.join(runs).strip()
        if line:
            lines.append(line)
    return span, lines

def parse_cell(lines, period_start, period_end):
    items = [l.strip() for l in lines if l.strip()]
    sessions = []
    cursor = period_start
    cat = None
    host = None
    pending = [None, 0, False]  # [label, dur, emitted_any]

    def flush(c):
        if pending[0] and not pending[2] and pending[1]:
            label, dur, _ = pending
            end = min(c + dur, period_end)
            if end > c:
                sessions.append({
                    'start': to_hhmm(c), 'end': to_hhmm(end),
                    'title': label, 'ai': label,
                    'category': CAT_MAP.get(label, 'R20'),
                    'host': host,
                })
                c = end
        pending[0] = None; pending[1] = 0; pending[2] = False
        return c

    for item in items:
        for sub in re.split(r'\s+\|\s+', item):
            sub = sub.strip().rstrip('.')
            if not sub:
                continue
            m = PAT_WITH_DUR.match(sub)
            if m:
                label = m.group(1).strip().rstrip('.')
                try:
                    dur = int(m.group(2))
                except ValueError:
                    continue
                if label in HOSTS:
                    host = label
                    continue
                if label in CATEGORIES:
                    cursor = flush(cursor)
                    cat = label
                    pending[0] = label; pending[1] = dur; pending[2] = False
                    continue
                session_label = label.lstrip('.').strip()
                if cursor + dur > period_end:
                    dur = max(0, period_end - cursor)
                if dur <= 0:
                    continue
                sessions.append({
                    'start': to_hhmm(cursor), 'end': to_hhmm(cursor + dur),
                    'title': session_label,
                    'ai': 'AI ' + session_label.split()[0] if session_label else session_label,
                    'category': CAT_MAP.get(cat, 'R20'),
                    'host': host,
                })
                cursor += dur
                if pending[0]:
                    pending[2] = True
            else:
                if sub in CATEGORIES:
                    cursor = flush(cursor); cat = sub
                elif sub in HOSTS:
                    host = sub
    cursor = flush(cursor)
    return sessions

def day_for_col(col, day_to_cols, room_assignment):
    for day, cols in day_to_cols.items():
        if col in cols:
            offset = cols.index(col)
            r = room_assignment(day, offset)
            if r is not None:
                return day, r
    return None

def parse_table(tbl, day_to_cols, room_assignment):
    rows = tbl.findall('w:tr', NS)
    sessions = []
    period_idx = 0
    for row in rows:
        cells = row.findall('w:tc', NS)
        if not cells:
            continue
        _, first_lines = cell_info(cells[0])
        first_text = ' '.join(first_lines)
        if not re.search(r'\d{1,2}:\d{2}', first_text):
            continue
        low = first_text.lower()
        if any(k in low for k in ('break', 'dinner', 'sessions end', 'no exceptions')):
            continue
        if len(cells) == 1:
            continue
        if period_idx not in PERIODS_BY_INDEX:
            period_idx += 1
            continue
        ps, pe = PERIODS_BY_INDEX[period_idx]
        p_start, p_end = to_min(ps), to_min(pe)
        period_idx += 1
        col = 0
        for cell in cells:
            span, lines = cell_info(cell)
            start_col = col
            col += span
            if start_col == 0:
                continue
            dr = day_for_col(start_col, day_to_cols, room_assignment)
            if not dr:
                continue
            day, room = dr
            for s in parse_cell(lines, p_start, p_end):
                s['day'] = day
                s['room'] = room
                sessions.append(s)
    return sessions

def parse_docx(docx_bytes):
    with zipfile.ZipFile(io.BytesIO(docx_bytes)) as z:
        xml = z.read('word/document.xml')
    root = ET.fromstring(xml)
    tables = root.findall('.//w:tbl', NS)
    if len(tables) < 2:
        raise RuntimeError("Document doesn't have the expected tables")

    online_rooms = ['Ballroom B (3F)', 'Ballroom A (3F)', 'Ballroom C (3F)']
    offline_rooms = ['Dalian Ballroom 1 (3F)', 'Shanghai Function room (3F)']
    all_rooms = online_rooms + offline_rooms

    online_day_cols = {'Mon':[1,2,3],'Tue':[4,5,6],'Wed':[7,8,9],'Thu':[10,11,12],'Fri':[13,14,15]}
    offline_day_cols = {'Mon':[1,2,3],'Tue':[4,5],'Wed':[6,7],'Thu':[8,9],'Fri':[10,11]}

    online_sessions = parse_table(
        tables[0], online_day_cols,
        lambda d, o: online_rooms[o] if o < len(online_rooms) else None)
    offline_sessions = parse_table(
        tables[1], offline_day_cols,
        lambda d, o: offline_rooms[min(o, 1)])

    sessions = online_sessions + offline_sessions
    seen = set()
    out = []
    for s in sessions:
        key = (s['day'], s['room'], s['start'], s['title'])
        if key in seen:
            continue
        seen.add(key)
        s['room'] = all_rooms.index(s['room'])
        out.append(s)
    return out, all_rooms

# -------- main --------

def main():
    print("Scanning 3GPP FTP for latest schedule...", file=sys.stderr)
    latest_url = find_latest_schedule_url()
    if not latest_url:
        print("ERROR: no 'online and offline schedules' docx found on 3GPP FTP.", file=sys.stderr)
        sys.exit(1)

    filename = urllib.parse.unquote(latest_url.split('/')[-1])
    meeting_match = re.match(r'^([A-Z0-9#]+)', filename)
    meeting = meeting_match.group(1) if meeting_match else 'RAN1'
    print(f"Latest schedule: {filename}", file=sys.stderr)

    docx_bytes = fetch(latest_url)
    sessions, rooms = parse_docx(docx_bytes)
    print(f"Parsed {len(sessions)} sessions", file=sys.stderr)

    schedule_data = {
        'meeting': meeting,
        'source': filename,
        'sourceUrl': latest_url,
        'generated': datetime.now(timezone.utc).strftime('%Y-%m-%d %H:%M UTC'),
        'rooms': rooms,
        'dayRange': {'start': '08:30', 'end': '19:30'},
        'breaks': [
            {'start': '10:30', 'end': '11:00', 'label': 'Morning Coffee'},
            {'start': '13:00', 'end': '14:30', 'label': 'Lunch'},
            {'start': '16:30', 'end': '17:00', 'label': 'Afternoon Coffee'},
        ],
        'sessions': sessions,
    }

    with open('template.html', 'r', encoding='utf-8') as f:
        template = f.read()

    data_json = json.dumps(schedule_data, ensure_ascii=False, indent=2)
    # Replace the placeholder block.
    new_html, n = re.subn(
        r'/\*INJECT:SCHEDULE_DATA\*/[\s\S]*?/\*END:SCHEDULE_DATA\*/',
        f'/*INJECT:SCHEDULE_DATA*/ {data_json} /*END:SCHEDULE_DATA*/',
        template
    )
    if n != 1:
        print("ERROR: could not find injection markers in template.html", file=sys.stderr)
        sys.exit(1)

    with open('index.html', 'w', encoding='utf-8') as f:
        f.write(new_html)
    print("Wrote index.html", file=sys.stderr)


if __name__ == '__main__':
    main()
