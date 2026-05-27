#!/usr/bin/env python3
"""Fetch the latest RAN1 schedule docx files from 3GPP FTP and regenerate index.html.

Merges three sources:
  - Main schedule (Chair_notes)
  - Hiroki's online + offline schedule (Hiroki_notes)
  - Sorour's online + offline schedule (Sorour_notes)

Also fetches current meeting info (city, dates) from the 3GPP portal.

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

FTP_BASE = "https://www.3gpp.org/ftp/Meetings_3GPP_SYNC/RAN1/Inbox"
FTP_MAIN    = FTP_BASE + "/Chair_notes/"
FTP_HIROKI  = FTP_BASE + "/Hiroki_notes/"
FTP_SOROUR  = FTP_BASE + "/Sorour_notes/"
MEETINGS_URL = "https://www.3gpp.org/dynareport?code=Meetings-R1.htm"

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
PAT_WITH_DUR = re.compile(r'^(.+?)\s*\(\s*~?\s*(\d+)\s*~?\s*(?:min)?\s*\)\s*\.?$')

# Map locations (city or country substring, lowercase) to IANA timezones.
# First matching key in order wins.
LOCATION_TZ = [
    ('china', 'Asia/Shanghai'), ('shanghai', 'Asia/Shanghai'),
    ('beijing', 'Asia/Shanghai'), ('wuhan', 'Asia/Shanghai'),
    ('hangzhou', 'Asia/Shanghai'), ('xi', 'Asia/Shanghai'),
    ('shenzhen', 'Asia/Shanghai'), ('nanjing', 'Asia/Shanghai'),
    ('chongqing', 'Asia/Shanghai'), ('hefei', 'Asia/Shanghai'),
    ('chengdu', 'Asia/Shanghai'), ('changsha', 'Asia/Shanghai'),
    ('qingdao', 'Asia/Shanghai'), ('xiamen', 'Asia/Shanghai'),
    ('zhuhai', 'Asia/Shanghai'),
    ('south korea', 'Asia/Seoul'), ('korea', 'Asia/Seoul'),
    ('seoul', 'Asia/Seoul'), ('incheon', 'Asia/Seoul'),
    ('busan', 'Asia/Seoul'), ('pusan', 'Asia/Seoul'),
    ('jeju', 'Asia/Seoul'), ('gyeongju', 'Asia/Seoul'),
    ('japan', 'Asia/Tokyo'), ('tokyo', 'Asia/Tokyo'),
    ('osaka', 'Asia/Tokyo'), ('yokohama', 'Asia/Tokyo'),
    ('fukuoka', 'Asia/Tokyo'), ('kobe', 'Asia/Tokyo'),
    ('nagoya', 'Asia/Tokyo'), ('miyazaki', 'Asia/Tokyo'),
    ('hong kong', 'Asia/Hong_Kong'), ('taipei', 'Asia/Taipei'),
    ('bengaluru', 'Asia/Kolkata'), ('india', 'Asia/Kolkata'),
    ('athens', 'Europe/Athens'),
    ('toulouse', 'Europe/Paris'), ('paris', 'Europe/Paris'),
    ('sophia-antipolis', 'Europe/Paris'), ('cannes', 'Europe/Paris'),
    ('maastricht', 'Europe/Amsterdam'), ('amsterdam', 'Europe/Amsterdam'),
    ('netherlands', 'Europe/Amsterdam'),
    ('prague', 'Europe/Prague'),
    ('malta', 'Europe/Malta'), ('saint julian', 'Europe/Malta'),
    ('gothenburg', 'Europe/Stockholm'), ('stockholm', 'Europe/Stockholm'),
    ('malmo', 'Europe/Stockholm'),
    ('ljubljana', 'Europe/Ljubljana'),
    ('madrid', 'Europe/Madrid'), ('valencia', 'Europe/Madrid'),
    ('barcelona', 'Europe/Madrid'), ('malaga', 'Europe/Madrid'),
    ('seville', 'Europe/Madrid'),
    ('dresden', 'Europe/Berlin'), ('berlin', 'Europe/Berlin'),
    ('hannover', 'Europe/Berlin'),
    ('lisbon', 'Europe/Lisbon'), ('lisboa', 'Europe/Lisbon'),
    ('sorrento', 'Europe/Rome'), ('turin', 'Europe/Rome'),
    ('helsinki', 'Europe/Helsinki'), ('espoo', 'Europe/Helsinki'),
    ('oulu', 'Europe/Helsinki'),
    ('warsaw', 'Europe/Warsaw'),
    ('riga', 'Europe/Riga'), ('tallinn', 'Europe/Tallinn'),
    ('budapest', 'Europe/Budapest'),
    ('dublin', 'Europe/Dublin'),
    ('belgrade', 'Europe/Belgrade'),
    ('vienna', 'Europe/Vienna'),
    ('calgary', 'America/Edmonton'),
    ('vancouver', 'America/Vancouver'),
    ('montreal', 'America/Toronto'),
    ('reno', 'America/Los_Angeles'),
    ('san francisco', 'America/Los_Angeles'),
    ('san diego', 'America/Los_Angeles'),
    ('los angeles', 'America/Los_Angeles'),
    ('spokane', 'America/Los_Angeles'),
    ('seattle', 'America/Los_Angeles'),
    ('scottsdale', 'America/Phoenix'),
    ('denver', 'America/Denver'),
    ('kansas city', 'America/Chicago'),
    ('new orleans', 'America/Chicago'),
    ('chicago', 'America/Chicago'),
    ('st louis', 'America/Chicago'),
    ('dallas', 'America/Chicago'),
    ('orlando', 'America/New_York'),
    ('jacksonville', 'America/New_York'),
    ('new york', 'America/New_York'),
    ('us', 'America/New_York'), ('united states', 'America/New_York'),
]

# -------- fetch --------

def fetch(url):
    req = urllib.request.Request(
        url, headers={'User-Agent': 'Mozilla/5.0 (RAN1 schedule builder)'})
    with urllib.request.urlopen(req, timeout=120) as r:
        return r.read()

def find_latest_schedule_url(folder_url, keyword='schedule'):
    """Scan a folder for the latest .docx whose name contains the keyword.
    Picks the file with the highest version-like suffix using natural-key sort.
    """
    try:
        html = fetch(folder_url).decode('utf-8', errors='replace')
    except Exception as e:
        print("WARN: failed to fetch %s: %s" % (folder_url, e), file=sys.stderr)
        return None
    link_re = re.compile(r'href="(https?://[^"]+\.docx)"', re.IGNORECASE)
    candidates = []
    for m in link_re.finditer(html):
        url = m.group(1)
        decoded = urllib.parse.unquote(url).lower()
        if keyword in decoded:
            key = tuple(int(t) if t.isdigit() else t
                        for t in re.split(r'(\d+)', decoded))
            candidates.append((key, url))
    if not candidates:
        return None
    candidates.sort()
    return candidates[-1][1]

def location_to_timezone(loc):
    if not loc:
        return 'UTC'
    s = loc.lower()
    for needle, tz in LOCATION_TZ:
        if needle in s:
            return tz
    return 'UTC'

def fetch_meeting_info(target_meeting):
    """Find meeting info matching number `target_meeting` (e.g. '125' or '126-bis')."""
    try:
        html = fetch(MEETINGS_URL).decode('utf-8', errors='replace')
    except Exception as e:
        print("WARN: failed to fetch meeting info: %s" % e, file=sys.stderr)
        return None
    # Each meeting row looks like:
    #   <a name="bmR1-125--2026-05-18">3GPPRAN1#125</a> ... >China</a> ... 2026-05-18 ... 2026-05-22
    # Robust extraction: split on the bm anchor, then for each chunk pull title, city, dates.
    chunks = re.split(r'<a\s+name="bm(R1-[^"]+?)--(\d{4}-\d{2}-\d{2})"', html)
    # chunks[0] is preamble; then for each match we get [slug, date, content]
    results = []
    i = 1
    while i + 1 < len(chunks):
        slug, anchor_date, content = chunks[i], chunks[i+1], chunks[i+2]
        i += 3
        # Title is the first thing after the opening, between > and </a>
        m_title = re.match(r'[^>]*>([^<]*)</a>', content)
        title = m_title.group(1).strip() if m_title else ''
        # City is in the next <a ...>City</a>
        m_city = re.search(r'<a[^>]*>([^<]*)</a>', content[m_title.end():] if m_title else content)
        city = m_city.group(1).strip() if m_city else ''
        # Two dates in YYYY-MM-DD form (use &#8209; or normal -)
        clean_content = content.replace('&#8209;', '-')
        dates = re.findall(r'(\d{4}-\d{2}-\d{2})', clean_content)
        start = dates[0] if dates else anchor_date
        end = dates[1] if len(dates) > 1 else start
        num_match = re.search(r'#(\d+(?:-bis)?)', title)
        meeting_num = num_match.group(1) if num_match else ''
        if not meeting_num:
            continue
        results.append({
            'meetingNumber': meeting_num,
            'title': title,
            'city': city,
            'startDate': start,
            'endDate': end,
            'timezone': location_to_timezone(city),
        })
    for r in results:
        if r['meetingNumber'] == target_meeting:
            return r
    return None

# -------- helpers --------

def to_min(hhmm):
    h, m = hhmm.split(':')
    return int(h) * 60 + int(m)

def to_hhmm(m):
    return "%02d:%02d" % (m // 60, m % 60)

def cell_paragraphs(cell):
    """Return [(text, has_bold, has_italic), ...] for each paragraph in cell."""
    result = []
    for p in cell.findall('.//w:p', NS):
        any_bold = False
        any_italic = False
        parts = []
        for r in p.findall('.//w:r', NS):
            rpr = r.find('w:rPr', NS)
            bold = rpr is not None and rpr.find('w:b', NS) is not None
            italic = rpr is not None and rpr.find('w:i', NS) is not None
            txt = ''.join(t.text for t in r.findall('.//w:t', NS) if t.text)
            if txt and txt.strip():
                if bold: any_bold = True
                if italic: any_italic = True
            parts.append(txt)
        full = ''.join(parts).strip()
        if full:
            result.append((full, any_bold, any_italic))
    return result

def cell_text_lines(cell):
    lines = []
    for p in cell.findall('.//w:p', NS):
        runs = [t.text for t in p.findall('.//w:t', NS) if t.text]
        line = ''.join(runs).strip()
        if line:
            lines.append(line)
    return lines

def cell_span(cell):
    span = 1
    gs = cell.find('.//w:gridSpan', NS)
    if gs is not None:
        span = int(gs.get(NS_W + 'val', '1'))
    return span

def end_align_sessions(sessions, period_end):
    """Shift sessions forward so they end-align with period_end.

    Used for Monday's first time slot where the meeting commences at ~09:00
    (so the 08:30-10:30 slot is only partially filled, and the content
    belongs at the END of the slot, not the start).
    """
    if not sessions:
        return sessions
    last_end = max(to_min(s['end']) for s in sessions)
    shift = period_end - last_end
    if shift > 0:
        for s in sessions:
            s['start'] = to_hhmm(to_min(s['start']) + shift)
            s['end']   = to_hhmm(to_min(s['end']) + shift)
    return sessions

def extract_room_names(docx_bytes):
    """Pull the set of room names mentioned in the main docx file.

    The schedule docx has two header paragraphs that look like:
      "<room1>(3F)<room1>(3F)<room2>(3F)<room2>(3F)...  RAN1#NNN Online Session Schedule"
      "<roomA>(3F)<roomA>(3F)<roomB>(3F)<roomB>(3F)...  RAN1#NNN Offline Session Schedule"
    (Each name is duplicated for layout reasons.)

    Returns (online_rooms, offline_rooms) as ordered lists with duplicates
    removed, or ([], []) if not found. The ORDER reflects the docx text — it
    may or may not match the actual table column order, which is why
    `merge_three` / `main` only uses this for verification + a warning, not as
    the source of truth for the column-to-room mapping.
    """
    try:
        with zipfile.ZipFile(io.BytesIO(docx_bytes)) as z:
            xml = z.read('word/document.xml')
    except Exception:
        return [], []
    root = ET.fromstring(xml)
    body = root.find('.//w:body', NS)
    if body is None:
        return [], []
    online_text = None
    offline_text = None
    for child in body:
        if child.tag.split('}')[-1] != 'p':
            continue
        text = ''.join(t.text for t in child.findall('.//w:t', NS) if t.text)
        if 'Online' in text and 'Schedule' in text and '(3F)' in text:
            online_text = text
        elif 'Offline' in text and 'Schedule' in text and '(3F)' in text:
            offline_text = text
    def parse(text):
        if not text:
            return []
        # Match anything ending in "(<digit>F)" — captures the leading name.
        names = re.findall(r'([A-Z][A-Za-z0-9 ]*?)\s*\(\dF\)', text)
        out = []
        for n in names:
            n = re.sub(r'\s+', ' ', n).strip()
            full = n + ' (3F)'  # normalize spacing
            if full not in out and 'Schedule' not in n and 'RAN' not in n:
                out.append(full)
        return out
    return parse(online_text), parse(offline_text)

def fuzzy_category(label):
    if not label:
        return 'R20'
    s = label.lower()
    if 'maint' in s or 'tei' in s or 'mntc' in s:
        return 'MAINT'
    if s.startswith('r20') or ' r20 ' in s or 'a-iot' in s or 'aiot' in s:
        return 'R20'
    if 'ai/ml' in s or 'ai 7' in s or 'ai 8' in s or 'ai 9' in s or s.startswith('ai '):
        return 'AI78'
    if '6g' in s or 'waveform' in s or 'isac' in s or 'sensing' in s:
        return '6GR'
    return 'R20'

# -------- main parser (existing logic) --------

def parse_main_cell(lines, period_start, period_end):
    # Pre-scan for two Monday-morning idioms:
    #   1. "RAN1#NNN commences at HH:MM on Monday" → shift effective_start
    #   2. "Agenda items 1, 2, 3, 4, 5"             → emit as filler later
    effective_start = period_start
    agenda_line = None
    filtered = []
    for line in lines:
        s = line.strip()
        if not s:
            continue
        m_com = re.search(r'commences\s+at\s+(\d{1,2}):(\d{2})', s, re.IGNORECASE)
        if m_com:
            h = int(m_com.group(1)); mm = int(m_com.group(2))
            new_start = h * 60 + mm
            if new_start > effective_start:
                effective_start = new_start
            continue
        if re.match(r'^agenda\s+items?\b', s, re.IGNORECASE):
            agenda_line = s
            continue
        filtered.append(line)

    items = [l.strip() for l in filtered if l.strip()]
    sessions = []
    cursor = effective_start
    cat = None
    host = None
    pending = [None, 0, False]

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

    # If we detected "Agenda items..." but had no other sessions, emit it as
    # a session spanning the entire effective window.
    if agenda_line and not sessions:
        sessions.append({
            'start': to_hhmm(effective_start),
            'end':   to_hhmm(period_end),
            'title': agenda_line,
            'ai':    'Agenda',
            'category': 'R20',
            'host':  host,
        })
    # Otherwise stash the agenda line on the first session so the table parser
    # can insert it as a leading filler AFTER end-alignment shifts the rest.
    if agenda_line and sessions:
        sessions[0]['_agenda_line'] = agenda_line
        sessions[0]['_effective_start'] = effective_start
    return sessions

def parse_main_table(tbl, day_to_cols, room_assignment, cell_parser):
    rows = tbl.findall('w:tr', NS)
    sessions = []
    period_idx = 0
    for row in rows:
        cells = row.findall('w:tc', NS)
        if not cells:
            continue
        first_lines = cell_text_lines(cells[0])
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
        this_period = period_idx
        period_idx += 1
        col = 0
        for cell in cells:
            span = cell_span(cell)
            lines = cell_text_lines(cell)
            start_col = col
            col += span
            if start_col == 0:
                continue
            dr = None
            for day, cols in day_to_cols.items():
                if start_col in cols:
                    offset = cols.index(start_col)
                    r = room_assignment(day, offset)
                    if r is not None:
                        dr = (day, r)
                    break
            if not dr:
                continue
            day, room = dr
            cell_sessions = list(cell_parser(lines, p_start, p_end))
            # Extract any stashed agenda-line info from parse_main_cell.
            agenda_line = None
            agenda_start = None
            if cell_sessions and '_agenda_line' in cell_sessions[0]:
                agenda_line = cell_sessions[0].pop('_agenda_line')
                agenda_start = cell_sessions[0].pop('_effective_start')
            # Monday's first slot starts late (commences at ~09:00) — end-align
            # so sessions sit at the end of the 08:30-10:30 slot, not the start.
            if day == 'Mon' and this_period == 0:
                end_align_sessions(cell_sessions, p_end)
            # If an "Agenda items..." line was detected and there's a gap
            # between the effective start and the first session, insert a
            # filler session.
            if agenda_line is not None and cell_sessions:
                first_min = min(to_min(s['start']) for s in cell_sessions)
                if first_min > agenda_start:
                    nums = re.findall(r'\d+(?:\.\d+)?', agenda_line)
                    ai_label = ('AI ' + ', '.join(nums)) if nums else 'Agenda'
                    cell_sessions.insert(0, {
                        'start': to_hhmm(agenda_start),
                        'end':   to_hhmm(first_min),
                        'title': agenda_line,
                        'ai':    ai_label,
                        'category': 'R20',
                        'host':  None,
                    })
            for s in cell_sessions:
                s['day'] = day
                s['room'] = room
                sessions.append(s)
    return sessions

def parse_main_docx(docx_bytes):
    with zipfile.ZipFile(io.BytesIO(docx_bytes)) as z:
        xml = z.read('word/document.xml')
    root = ET.fromstring(xml)
    tables = root.findall('.//w:tbl', NS)
    if len(tables) < 2:
        raise RuntimeError("Main document doesn't have the expected tables")
    online_rooms = ['Ballroom B (3F)', 'Ballroom A (3F)', 'Ballroom C (3F)']
    offline_rooms = ['Dalian Ballroom 1 (3F)', 'Shanghai Function room (3F)']
    online_day_cols  = {'Mon':[1,2,3],'Tue':[4,5,6],'Wed':[7,8,9],'Thu':[10,11,12],'Fri':[13,14,15]}
    offline_day_cols = {'Mon':[1,2,3],'Tue':[4,5],'Wed':[6,7],'Thu':[8,9],'Fri':[10,11]}
    online = parse_main_table(
        tables[0], online_day_cols,
        lambda d, o: online_rooms[o] if o < len(online_rooms) else None,
        parse_main_cell)
    offline = parse_main_table(
        tables[1], offline_day_cols,
        lambda d, o: offline_rooms[min(o, 1)],
        parse_main_cell)
    return online + offline

# -------- Sorour parser (similar to main; ignores "-" prefixed sub-bullets) --------

def parse_sorour_cell(lines, period_start, period_end):
    filtered = [l for l in lines if not l.lstrip().startswith('-')]
    return parse_main_cell(filtered, period_start, period_end)

def parse_sorour_docx(docx_bytes):
    with zipfile.ZipFile(io.BytesIO(docx_bytes)) as z:
        xml = z.read('word/document.xml')
    root = ET.fromstring(xml)
    tables = root.findall('.//w:tbl', NS)
    if len(tables) < 2:
        return []
    online_rooms = ['Ballroom B (3F)', 'Ballroom A (3F)', 'Ballroom C (3F)']
    offline_rooms = ['Dalian Ballroom 1 (3F)', 'Shanghai Function room (3F)']
    online_day_cols  = {'Mon':[1,2,3],'Tue':[4,5,6],'Wed':[7,8,9],'Thu':[10,11,12],'Fri':[13,14,15]}
    offline_day_cols = {'Mon':[1,2,3],'Tue':[4,5],'Wed':[6,7],'Thu':[8,9],'Fri':[10,11]}
    online = parse_main_table(
        tables[0], online_day_cols,
        lambda d, o: online_rooms[o] if o < len(online_rooms) else None,
        parse_sorour_cell)
    offline = parse_main_table(
        tables[1], offline_day_cols,
        lambda d, o: offline_rooms[min(o, 1)],
        parse_sorour_cell)
    return online + offline

# -------- Hiroki parser (bold=category, italic=sub-session) --------

def parse_hiroki_cell(cell, period_start, period_end):
    paras = cell_paragraphs(cell)
    sessions = []
    cursor = period_start
    cat_label = None
    for text, is_bold, is_italic in paras:
        low = text.lower()
        if any(k in low for k in ('commences', 'expected to close', 'no exceptions',
                                  'all sessions end', 'tbd', 'to be assigned',
                                  'early dinner')):
            continue
        if cursor >= period_end:
            break
        m = PAT_WITH_DUR.match(text)
        if m:
            label = m.group(1).strip().rstrip('.').strip()
            try:
                dur = int(m.group(2))
            except ValueError:
                continue
            if is_bold and not is_italic:
                cat_label = label
                continue
            else:
                if cursor + dur > period_end:
                    dur = max(0, period_end - cursor)
                if dur <= 0:
                    break
                title = label.lstrip('.').strip()
                ai_label = title
                if not ai_label.lower().startswith('ai'):
                    ai_label = 'AI ' + (title.split()[0] if title else '')
                sessions.append({
                    'start': to_hhmm(cursor),
                    'end':   to_hhmm(cursor + dur),
                    'title': title,
                    'ai':    ai_label,
                    'category': fuzzy_category(cat_label or title),
                    'host':  'Hiroki',
                })
                cursor += dur
        else:
            if is_bold and not is_italic:
                cat_label = text
            else:
                dur = max(0, period_end - cursor)
                if dur > 0:
                    title = text.lstrip('.').strip()
                    ai_label = title
                    if not ai_label.lower().startswith('ai'):
                        ai_label = 'AI ' + (title.split()[0] if title else '')
                    sessions.append({
                        'start': to_hhmm(cursor),
                        'end':   to_hhmm(cursor + dur),
                        'title': title,
                        'ai':    ai_label,
                        'category': fuzzy_category(cat_label or title),
                        'host':  'Hiroki',
                    })
                    cursor = period_end
                break
    return sessions

def parse_hiroki_offline_cell(cell, period_start, period_end):
    text = ' '.join(' '.join(p[0] for p in cell_paragraphs(cell)).split())
    sessions = []
    parts = list(re.finditer(r'(\d{1,2}):(\d{2})\s*[-~]\s*(\d{1,2}):(\d{2})', text))
    for i, m in enumerate(parts):
        sh, sm, eh, em = map(int, m.groups())
        start = sh * 60 + sm
        end = eh * 60 + em
        label_start = m.end()
        label_end = parts[i+1].start() if i + 1 < len(parts) else len(text)
        label = text[label_start:label_end].strip()
        # Strip standalone "(N)" duration markers
        label = re.sub(r'\s*\(\s*\d+\s*\)\s*', ' ', label).strip()
        if not label:
            continue
        if start >= period_end or end <= period_start:
            continue
        start = max(start, period_start)
        end = min(end, period_end)
        if end <= start:
            continue
        ai_label = label
        if not ai_label.lower().startswith('ai'):
            ai_label = 'AI ' + (label.split()[0] if label else '')
        sessions.append({
            'start': to_hhmm(start),
            'end':   to_hhmm(end),
            'title': label,
            'ai':    ai_label,
            'category': fuzzy_category(label),
            'host':  'Hiroki',
        })
    return sessions

def parse_hiroki_table(tbl, offline):
    rows = tbl.findall('w:tr', NS)
    sessions = []
    period_idx = 0
    DAYS = ['Mon', 'Tue', 'Wed', 'Thu', 'Fri']
    for row in rows:
        cells = row.findall('w:tc', NS)
        if not cells:
            continue
        first_lines = cell_text_lines(cells[0])
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
        this_period = period_idx
        period_idx += 1
        day_cells_seen = 0
        for ci, cell in enumerate(cells):
            if ci == 0:
                continue
            if day_cells_seen >= len(DAYS):
                break
            day = DAYS[day_cells_seen]
            day_cells_seen += 1
            if offline:
                day_sessions = parse_hiroki_offline_cell(cell, p_start, p_end)
            else:
                day_sessions = parse_hiroki_cell(cell, p_start, p_end)
            if day == 'Mon' and this_period == 0:
                end_align_sessions(day_sessions, p_end)
            for s in day_sessions:
                s['day'] = day
                if not offline:
                    s['room'] = 'Ballroom C (3F)'
                sessions.append(s)
    return sessions

def parse_hiroki_docx(docx_bytes):
    with zipfile.ZipFile(io.BytesIO(docx_bytes)) as z:
        xml = z.read('word/document.xml')
    root = ET.fromstring(xml)
    tables = root.findall('.//w:tbl', NS)
    if not tables:
        return {'online': [], 'offline': []}
    online = parse_hiroki_table(tables[0], offline=False)
    offline = []
    if len(tables) >= 2:
        offline = parse_hiroki_table(tables[1], offline=True)
    return {'online': online, 'offline': offline}

# -------- merge --------

def merge_three(main_sessions, hiroki_data, sorour_sessions, all_rooms):
    out = []

    # Period intervals used for per-(day, period) override granularity.
    period_intervals = [(to_min(s), to_min(e)) for s, e in PERIODS_BY_INDEX.values()]
    def find_period_idx(s_min):
        for i, (ps, pe) in enumerate(period_intervals):
            if ps <= s_min < pe:
                return i
        return None

    # For each (day, period), which rooms does Sorour/Hiroki actually override?
    # Sorour file's Monday morning row often has gridSpan=3 on the first cell —
    # meaning Sorour has NO Monday-AM data for Ballroom A or C. In those cases
    # we must keep the main file's content rather than blank the slot.
    sorour_a_cover = set()  # (day, period_idx)
    for s in sorour_sessions:
        if s['room'] != 'Ballroom A (3F)':
            continue
        pi = find_period_idx(to_min(s['start']))
        if pi is not None:
            sorour_a_cover.add((s['day'], pi))
    hiroki_c_cover = set()
    for s in hiroki_data['online']:
        pi = find_period_idx(to_min(s['start']))
        if pi is not None:
            hiroki_c_cover.add((s['day'], pi))

    for s in main_sessions:
        sm = to_min(s['start'])
        pi = find_period_idx(sm)
        if s['room'] == 'Ballroom A (3F)' and pi is not None and (s['day'], pi) in sorour_a_cover:
            continue
        if s['room'] == 'Ballroom C (3F)' and pi is not None and (s['day'], pi) in hiroki_c_cover:
            continue
        out.append(s)
    for s in sorour_sessions:
        if s['room'] == 'Ballroom A (3F)':
            out.append(s)
    for s in hiroki_data['online']:
        out.append(s)

    # (period_intervals + find_period_idx already defined above.)
    def find_period(s_min):
        for ps, pe in period_intervals:
            if ps <= s_min < pe:
                return (ps, pe)
        return None
    def by_day_period(sessions):
        d = {}
        for s in sessions:
            sm = to_min(s['start'])
            p = find_period(sm)
            if p is None:
                continue
            d.setdefault((s['day'], p), []).append(s)
        return d

    hiroki_off_idx = by_day_period(hiroki_data['offline'])
    sorour_off_idx = by_day_period([
        s for s in sorour_sessions
        if s['room'] in ('Dalian Ballroom 1 (3F)', 'Shanghai Function room (3F)')
    ])

    keep = []
    for s in out:
        if s['room'] not in ('Dalian Ballroom 1 (3F)', 'Shanghai Function room (3F)'):
            keep.append(s)
            continue
        sm = to_min(s['start'])
        p = find_period(sm)
        host = s.get('host')
        replaced = False
        if p and host == 'Hiroki' and (s['day'], p) in hiroki_off_idx:
            for sub in hiroki_off_idx[(s['day'], p)]:
                ns = dict(sub)
                ns['room'] = s['room']
                ns['day'] = s['day']
                keep.append(ns)
            replaced = True
        elif p and host in ('Sorour', 'Sorouri') and (s['day'], p) in sorour_off_idx:
            for sub in sorour_off_idx[(s['day'], p)]:
                ns = dict(sub)
                ns['room'] = s['room']
                ns['day'] = s['day']
                keep.append(ns)
            replaced = True
        if not replaced:
            keep.append(s)

    seen = set()
    dedup = []
    for s in keep:
        key = (s['day'], s['room'], s['start'], s['end'], s['title'])
        if key in seen:
            continue
        seen.add(key)
        dedup.append(s)

    final = []
    for s in dedup:
        if s['room'] not in all_rooms:
            continue
        s2 = dict(s)
        s2['room'] = all_rooms.index(s['room'])
        final.append(s2)
    return final

# -------- main --------

def main():
    print("Scanning 3GPP FTP for latest schedule files...", file=sys.stderr)
    main_url   = find_latest_schedule_url(FTP_MAIN, 'online and offline schedules')
    hiroki_url = find_latest_schedule_url(FTP_HIROKI, 'schedule')
    sorour_url = find_latest_schedule_url(FTP_SOROUR, 'schedule')

    if not main_url:
        print("ERROR: no main schedule docx found.", file=sys.stderr)
        sys.exit(1)

    main_fn   = urllib.parse.unquote(main_url.split('/')[-1])
    hiroki_fn = urllib.parse.unquote(hiroki_url.split('/')[-1]) if hiroki_url else None
    sorour_fn = urllib.parse.unquote(sorour_url.split('/')[-1]) if sorour_url else None
    print("  main:   %s" % main_fn, file=sys.stderr)
    print("  hiroki: %s" % hiroki_fn, file=sys.stderr)
    print("  sorour: %s" % sorour_fn, file=sys.stderr)

    # Need to be specific: the filename starts "RAN1#NNN...", not just any digit run.
    meeting_match = re.search(r'#(\d+(?:-?bis)?)', main_fn)
    meeting_num = meeting_match.group(1) if meeting_match else '125'
    meeting_label_match = re.match(r'^([A-Z0-9#]+)', main_fn)
    meeting_label = meeting_label_match.group(1) if meeting_label_match else 'RAN1#' + meeting_num

    main_docx_bytes = fetch(main_url)
    main_sessions = parse_main_docx(main_docx_bytes)
    print("  main:   %d sessions" % len(main_sessions), file=sys.stderr)

    hiroki_data = {'online': [], 'offline': []}
    if hiroki_url:
        try:
            hiroki_data = parse_hiroki_docx(fetch(hiroki_url))
            print("  hiroki: %d online + %d offline" % (len(hiroki_data['online']), len(hiroki_data['offline'])), file=sys.stderr)
        except Exception as e:
            print("  hiroki: parse failed: %s" % e, file=sys.stderr)

    sorour_sessions = []
    if sorour_url:
        try:
            sorour_sessions = parse_sorour_docx(fetch(sorour_url))
            print("  sorour: %d sessions" % len(sorour_sessions), file=sys.stderr)
        except Exception as e:
            print("  sorour: parse failed: %s" % e, file=sys.stderr)

    all_rooms = [
        'Ballroom B (3F)', 'Ballroom A (3F)', 'Ballroom C (3F)',
        'Dalian Ballroom 1 (3F)', 'Shanghai Function room (3F)',
    ]

    # Verify the hardcoded room names still match what's in the docx.
    # If they don't (e.g. the next meeting in Maastricht has different rooms),
    # print a prominent warning so the user knows to update `all_rooms` above.
    extracted_online, extracted_offline = extract_room_names(main_docx_bytes)
    hardcoded_online = set(all_rooms[:3])
    hardcoded_offline = set(all_rooms[3:])
    extracted_online_set = set(extracted_online)
    extracted_offline_set = set(extracted_offline)
    if extracted_online and extracted_online_set != hardcoded_online:
        print("", file=sys.stderr)
        print("⚠️  WARNING: docx ONLINE room names changed — please update build.py!", file=sys.stderr)
        print("   docx says:  %s" % extracted_online, file=sys.stderr)
        print("   build.py:   %s" % all_rooms[:3], file=sys.stderr)
        print("", file=sys.stderr)
    if extracted_offline and extracted_offline_set != hardcoded_offline:
        print("", file=sys.stderr)
        print("⚠️  WARNING: docx OFFLINE room names changed — please update build.py!", file=sys.stderr)
        print("   docx says:  %s" % extracted_offline, file=sys.stderr)
        print("   build.py:   %s" % all_rooms[3:], file=sys.stderr)
        print("", file=sys.stderr)
    if extracted_online_set == hardcoded_online and extracted_offline_set == hardcoded_offline:
        print("  room names OK (match docx)", file=sys.stderr)
    merged = merge_three(main_sessions, hiroki_data, sorour_sessions, all_rooms)
    print("  merged: %d sessions" % len(merged), file=sys.stderr)

    print("Fetching meeting info from 3GPP portal...", file=sys.stderr)
    meeting_info = fetch_meeting_info(meeting_num)
    if meeting_info:
        print("  found: %s in %s (%s – %s, %s)" % (meeting_info['title'], meeting_info['city'],
              meeting_info['startDate'], meeting_info['endDate'], meeting_info['timezone']), file=sys.stderr)
    else:
        print("  not found for meeting #%s; using fallback" % meeting_num, file=sys.stderr)
        meeting_info = {
            'meetingNumber': meeting_num,
            'title': meeting_label,
            'city': '',
            'startDate': '',
            'endDate': '',
            'timezone': 'UTC',
        }

    sources = [main_fn] + ([hiroki_fn] if hiroki_fn else []) + ([sorour_fn] if sorour_fn else [])
    source_urls = [main_url] + ([hiroki_url] if hiroki_url else []) + ([sorour_url] if sorour_url else [])

    schedule_data = {
        'meeting': meeting_label,
        'meetingInfo': meeting_info,
        'sources': sources,
        'sourceUrls': source_urls,
        'generated': datetime.now(timezone.utc).strftime('%Y-%m-%d %H:%M UTC'),
        'rooms': all_rooms,
        'dayRange': {'start': '08:30', 'end': '19:30'},
        'breaks': [
            {'start': '10:30', 'end': '11:00', 'label': 'Morning Coffee'},
            {'start': '13:00', 'end': '14:30', 'label': 'Lunch'},
            {'start': '16:30', 'end': '17:00', 'label': 'Afternoon Coffee'},
        ],
        'sessions': merged,
    }

    with open('template.html', 'r', encoding='utf-8') as f:
        template = f.read()

    data_json = json.dumps(schedule_data, ensure_ascii=False, indent=2)
    new_html, n = re.subn(
        r'/\*INJECT:SCHEDULE_DATA\*/[\s\S]*?/\*END:SCHEDULE_DATA\*/',
        '/*INJECT:SCHEDULE_DATA*/ ' + data_json + ' /*END:SCHEDULE_DATA*/',
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
