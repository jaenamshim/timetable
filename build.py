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
# All host name variants seen in the docx sources. 'Sorour i' (with space)
# and 'Sorouri' (no space) are the same person — aliased to canonical 'Sorour'.
HOSTS = {'Hiroki', 'Sorour', 'Sorouri', 'Sorour i', 'Xiaodong'}
HOST_ALIASES = {
    'Sorour i': 'Sorour',
    'Sorouri':  'Sorour',
}
# Each host's primary offline room. When an offline gridSpan>1 cell covers
# both Dalian and Shanghai with multiple host paragraphs (e.g. Fri 14:30
# cell "To be assigned by Sorour / To be assigned by Hiroki"), each host's
# sessions are routed to the room below — matching karlla's behavior.
HOST_TYPICAL_OFFLINE_ROOM = {
    'Hiroki':   'Dalian Ballroom 1 (3F)',
    'Sorour':   'Shanghai Function room (3F)',
    'Xiaodong': 'Dalian Ballroom 1 (3F)',
}

def normalize_host(s):
    """Collapse the 'Sorour i' / 'Sorouri' variants to canonical 'Sorour'.
    No-op for hosts already in canonical form."""
    if not s:
        return s
    s = re.sub(r'\s+', ' ', s.strip())
    return HOST_ALIASES.get(s, s)

CAT_MAP = {
    '6GR': '6GR', 'R20': 'R20',
    'AI 7/8': 'AI78', 'AI 8': 'AI78', 'AI 9': 'AI78', 'AI/ML': 'AI78',
    'Maintenance': 'MAINT', 'MNTC': 'MAINT',
}
# Normalized lookup table — docx labels often have stray spaces ("R 20",
# "M aintenance") that don't match the canonical CATEGORIES set directly.
_CAT_NORM = {re.sub(r'\s+', '', c).upper(): c for c in CATEGORIES}
def normalize_category(s):
    """Return canonical category name (e.g. 'R20') if `s` matches one of
    CATEGORIES up to whitespace and case; else None."""
    if not s:
        return None
    return _CAT_NORM.get(re.sub(r'\s+', '', s).upper())

def extract_ai_label(title):
    """Return 'AI X.Y' style AI item extracted from a session title. Order:
    explicit 'AI X.Y' in the title; leading digit.digit; first-word fallback."""
    if not title:
        return title
    m = re.search(r'\bAI\s+(\d+(?:\.\d+)*(?:/\d+(?:\.\d+)*)?)', title, re.IGNORECASE)
    if m:
        return 'AI ' + m.group(1)
    m = re.search(r'\b(\d+\.\d+(?:\.\d+)*)', title)
    if m:
        return 'AI ' + m.group(1)
    first = title.split()[0] if title.split() else title
    return 'AI ' + first
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
        if not line:
            continue
        # docx authors sometimes glue items together without whitespace,
        # e.g. "MIMO (60)AI/ML (60)" or "10.3.2 Modulation (60) R20 (30)".
        # Insert a newline after any "(NN)" duration marker when followed by
        # a letter (with optional whitespace) — so each duration-bearing item
        # becomes its own line.
        marked = re.sub(r'(\(\d{1,3}\))\s*(?=[A-Za-z])', r'\1\n', line)
        for sl in marked.split('\n'):
            sl = sl.strip()
            if sl:
                lines.append(sl)
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
    # Pre-scan for these idioms:
    #   1. "RAN1#NNN commences at HH:MM on Monday" → shift effective_start
    #   2. "Agenda items 1, 2, 3, 4, 5"             → emit as filler later
    #   3. "expected to close at HH:MM"             → extend effective_end
    effective_start = period_start
    effective_end = period_end
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
        m_close = re.search(r'close[s]?\s+at\s+(\d{1,2}):(\d{2})', s, re.IGNORECASE)
        if m_close:
            h = int(m_close.group(1)); mm = int(m_close.group(2))
            new_end = h * 60 + mm
            if new_end > effective_end:
                effective_end = new_end
            # Don't `continue` — keep the line so its prose ("RAN1#... expected
            # to close at 17:00") is still available as title material.
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
    # Bare labels (no duration, not a category/host/AI number) are accumulated
    # here and prepended to the next emitted session's title — preserves notes
    # like "Sweep" that appear between sessions in main's cells.
    extra_label_parts = []

    def flush(c):
        if pending[0] and not pending[2] and pending[1]:
            label, dur, _ = pending
            end = min(c + dur, effective_end)
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
                    host = normalize_host(label)
                    continue
                canon_cat = normalize_category(label)
                if canon_cat is not None:
                    cursor = flush(cursor)
                    cat = canon_cat
                    pending[0] = canon_cat; pending[1] = dur; pending[2] = False
                    continue
                session_label = label.lstrip('.').strip()
                if cursor + dur > effective_end:
                    dur = max(0, effective_end - cursor)
                if dur <= 0:
                    continue
                sessions.append({
                    'start': to_hhmm(cursor), 'end': to_hhmm(cursor + dur),
                    'title': session_label,
                    'ai': extract_ai_label(session_label),
                    'category': CAT_MAP.get(cat, 'R20'),
                    'host': host,
                })
                cursor += dur
                if pending[0]:
                    pending[2] = True
            else:
                canon_cat = normalize_category(sub)
                if canon_cat is not None:
                    cursor = flush(cursor); cat = canon_cat
                elif sub in HOSTS:
                    host = normalize_host(sub)
                else:
                    # "Sorour (TBD)" / "Hiroki (placeholder)" — host name
                    # followed by a non-numeric parenthetical. PAT_WITH_DUR
                    # rejects these (it wants a duration in the parens), but
                    # the host should still be recognized.
                    m_host = re.match(
                        r'^(Hiroki|Sorour\s?i?|Sorouri|Xiaodong)\s*\(',
                        sub)
                    if m_host:
                        host = normalize_host(m_host.group(1))
                        continue
                    sub_stripped = sub.lstrip('.').strip()
                    if re.match(r'^\d+(\.\d+)+', sub_stripped):
                        # Bare AI item — claim 60 min by default (or whatever
                        # remains if less). If there are more bare items after,
                        # they'll claim their own 60-min slots.
                        cursor = flush(cursor)
                        remaining = effective_end - cursor
                        if remaining > 0:
                            dur = min(60, remaining)
                            # Drop any malformed parenthetical and trailing
                            # garbage from the title — e.g. ".10.5.4.x (80)0)"
                            # → "10.5.4.x". The duration we use is the default,
                            # not whatever junk the docx happened to contain.
                            title = re.sub(r'\s*\(.*$', '', sub_stripped).strip()
                            if not title:
                                title = sub_stripped
                            sessions.append({
                                'start': to_hhmm(cursor),
                                'end':   to_hhmm(cursor + dur),
                                'title': title,
                                'ai':    extract_ai_label(title),
                                'category': CAT_MAP.get(cat, 'R20'),
                                'host':  host,
                            })
                            cursor += dur
                    else:
                        # Plain prose like "Sweep" or "6GR check points" —
                        # buffer for the next session's title or, if no
                        # session follows, the leftover-prose fallback below.
                        extra_label_parts.append(sub)
    cursor = flush(cursor)

    # Trailing-prose fallback: leftover bare labels not yet attached to any
    # session AND there's leftover time in the period — emit one final
    # session with the prose as title (e.g. "6GR check points" 12:00-13:00).
    if extra_label_parts and cursor < effective_end:
        title_parts = list(extra_label_parts)
        # If the current category isn't already in the prose, include it so
        # cells like "Sweep / 6GR" render as title='Sweep 6GR'.
        if cat and not any(cat.lower() in p.lower() for p in extra_label_parts):
            title_parts.append(cat)
        title = ' '.join(title_parts).strip()
        if title.lower() != 'tbd':
            sessions.append({
                'start': to_hhmm(cursor),
                'end':   to_hhmm(effective_end),
                'title': title,
                'ai':    extract_ai_label(title),
                'category': CAT_MAP.get(cat, 'R20'),
                'host':  host,
            })
            cursor = effective_end
        extra_label_parts = []

    # Closing-remarks fallback: cell has only prose (no duration/AI items),
    # e.g. Friday "Any other open issues... expected to close at 17:00".
    # Skip empty-placeholder text ("TBD", "To be assigned by ...") since
    # those don't represent real scheduled sessions.
    if not sessions and extra_label_parts:
        title_parts = list(extra_label_parts)
        # If the current category isn't already in the prose, append it —
        # cells like "Sweep / 6GR" should render title='Sweep 6GR' not just
        # 'Sweep' (matches karlla's behavior on Fri Ballroom B 11:00-13:00).
        if cat and not any(cat.lower() in p.lower() for p in extra_label_parts):
            title_parts.append(cat)
        title = ' '.join(title_parts).strip()
        if title.lower().strip() != 'tbd':
            sessions.append({
                'start': to_hhmm(effective_start),
                'end':   to_hhmm(effective_end),
                'title': title,
                'ai':    'Closing',
                'category': CAT_MAP.get(cat, 'R20'),
                'host':  host,
            })
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

            # Multi-room offline span: if this gridSpan>1 cell covers more
            # than one distinct room (e.g. Fri offline col 10 gs=2 spans
            # Dalian+Shanghai), split the cell's paragraphs at host-name
            # markers and route each segment to that host's typical offline
            # room. Matches karlla: cell "To be assigned by Sorour / To be
            # assigned by Hiroki" → Sorour→Shanghai, Hiroki→Dalian.
            if span > 1:
                cols_covered = list(range(start_col, start_col + span))
                rooms_covered = []
                for c in cols_covered:
                    if c in day_to_cols.get(day, []):
                        off = day_to_cols[day].index(c)
                        r = room_assignment(day, off)
                        if r is not None:
                            rooms_covered.append(r)
                unique_rooms = list(dict.fromkeys(rooms_covered))
                if len(unique_rooms) >= 2:
                    host_pat = re.compile(
                        r'\b(Hiroki|Sorour\s?i?|Sorouri|Xiaodong)\b')
                    # Pair each input line with the host name it references
                    # (if any), then group consecutive same-host lines into
                    # segments. The order in which hosts appear in the cell
                    # is irrelevant — each host's segment is routed to its
                    # typical room, so paragraph order doesn't matter.
                    segments = []  # list of [host_name, [lines...]]
                    for line in lines:
                        m = host_pat.search(line)
                        h = normalize_host(m.group(1)) if m else None
                        if segments and segments[-1][0] == h:
                            segments[-1][1].append(line)
                        else:
                            segments.append([h, [line]])
                    # Only redistribute when at least one segment has a known
                    # host with a typical room covered by this span.
                    routable = any(
                        seg_host and HOST_TYPICAL_OFFLINE_ROOM.get(seg_host)
                        in unique_rooms for seg_host, _ in segments)
                    if routable:
                        for seg_host, seg_lines in segments:
                            tgt = HOST_TYPICAL_OFFLINE_ROOM.get(seg_host)
                            if not (tgt and tgt in unique_rooms):
                                continue
                            seg_sessions = list(
                                cell_parser(seg_lines, p_start, p_end))
                            for s in seg_sessions:
                                s['day'] = day
                                s['room'] = tgt
                                s['span'] = 1
                                if not s.get('host'):
                                    s['host'] = seg_host
                                sessions.append(s)
                        continue  # skip the default single-cell path below

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
            # Clamp span=1 when all columns covered by gridSpan map to the
            # same room (e.g. Mon offline col 2 gs=2 covers cols 2,3 which
            # both clamp to Shanghai). The visual span would otherwise paint
            # into a non-existent neighbor column.
            effective_span = span
            if span > 1:
                cols_covered = list(range(start_col, start_col + span))
                rooms_for_cols = set()
                for c in cols_covered:
                    if c in day_to_cols.get(day, []):
                        off = day_to_cols[day].index(c)
                        r = room_assignment(day, off)
                        if r is not None:
                            rooms_for_cols.add(r)
                if len(rooms_for_cols) <= 1:
                    effective_span = 1
            for s in cell_sessions:
                s['day'] = day
                s['room'] = room
                s['span'] = effective_span
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
        parse_main_docx_room_for_offline,
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
        parse_main_docx_room_for_offline,
        parse_sorour_cell)
    # Sorour file ships a 3rd table that's a Ballroom-A-specific detail
    # schedule (the one karlla1220 uses for Ballroom A) — header layout is
    # Mon=1col, Tue=2cols, Wed=1col, Thu=1col, Fri=1col. Every cell belongs
    # to Ballroom A for its day.
    a_only_extras = []
    if len(tables) >= 3:
        a_cols = {'Mon': [1], 'Tue': [2, 3], 'Wed': [4], 'Thu': [5], 'Fri': [6]}
        a_only_extras = parse_main_table(
            tables[2], a_cols,
            lambda d, o: 'Ballroom A (3F)',
            parse_sorour_cell)
    # Whole Sorour file describes the room Sorour manages (Ballroom A) —
    # default host='Sorour' for any A-room session without an explicit host
    # token in the cell. Sessions where the cell text gave a host (Sorouri,
    # Hiroki, Xiaodong) keep theirs.
    for s in online + offline + a_only_extras:
        if s.get('room') == 'Ballroom A (3F)' and not s.get('host'):
            s['host'] = 'Sorour'
    return online + offline + a_only_extras

def parse_main_docx_room_for_offline(day, offset):
    """Mon offline has 3 cols (Dalian / Shanghai / Sorouri-side); other days
    have 2 (Dalian / Shanghai). The third Mon column is still Shanghai —
    Sorouri's NTN-NR/IoT/NTN sessions happen there during the morning plenary
    period when Ballroom A is otherwise occupied by the joint AM2 session."""
    offline_rooms = ['Dalian Ballroom 1 (3F)', 'Shanghai Function room (3F)']
    return offline_rooms[min(offset, 1)]

# -------- Hiroki parser (bold=category, italic=sub-session) --------

def parse_hiroki_cell(cell, period_start, period_end):
    paras = cell_paragraphs(cell)
    # If a cell has no duration markers at all, treat it as informational
    # (notes/labels) rather than schedulable content. Karlla also skips such
    # cells — e.g. Hiroki Fri Ballroom C 11:00-13:00 lists "6GR check points
    # / 6G waveform / 10.2.1 / 6G ISAC / 10.8.1, 10.8.2" with no (NN)
    # anywhere, and karlla emits nothing for that slot.
    if not any(re.search(r'\(\s*\d+\s*\)', t) for t, _, _ in paras):
        return []
    sessions = []
    cursor = period_start
    cat_label = None
    # Pending "bold-with-duration" block — declares a span of N minutes that
    # the following italic paragraph(s) describe. Each italic-with-its-own-
    # duration consumes from this budget; a comma-list italic (e.g.
    # "10.5.4.3(45), 10.5.4.1(45)") uses the WHOLE remaining budget as one
    # session with the bold label as title and the italic as ai text.
    pending_label = None
    pending_remaining = 0

    def emit(title, ai_text, dur, cat_hint):
        nonlocal cursor
        d = min(dur, period_end - cursor)
        if d <= 0:
            return
        # If ai_text is provided (e.g. the italic AI sub-item), keep it as-is
        # since it's the docx's own listing; otherwise derive from title.
        ai_value = ai_text if ai_text else extract_ai_label(title)
        sessions.append({
            'start': to_hhmm(cursor),
            'end':   to_hhmm(cursor + d),
            'title': title,
            'ai':    ai_value,
            'category': fuzzy_category(cat_hint or title),
            'host':  'Hiroki',
        })
        cursor += d

    def flush_pending():
        nonlocal pending_label, pending_remaining
        if pending_label is not None and pending_remaining > 0:
            emit(pending_label, None, pending_remaining, pending_label)
        pending_label = None
        pending_remaining = 0

    for text, is_bold, is_italic in paras:
        low = text.lower()
        if any(k in low for k in ('commences', 'expected to close', 'no exceptions',
                                  'all sessions end', 'tbd', 'to be assigned',
                                  'early dinner')):
            continue
        if cursor >= period_end:
            break

        if is_bold and not is_italic:
            flush_pending()
            m = PAT_WITH_DUR.match(text)
            if m:
                try:
                    pending_label = m.group(1).strip().rstrip('.').strip()
                    pending_remaining = int(m.group(2))
                except ValueError:
                    pending_label = None
                    pending_remaining = 0
            else:
                cat_label = text
            continue

        # Italic or plain prose (not pure bold).
        m = PAT_WITH_DUR.match(text)
        if m and ',' not in text:
            # Standalone italic-with-duration sub-session. Consumes from any
            # pending bold-block budget.
            try:
                label = m.group(1).strip().rstrip('.').strip()
                dur = int(m.group(2))
            except ValueError:
                continue
            title = label.lstrip('.').strip()
            emit(title, title, dur, pending_label or cat_label or title)
            if pending_label is not None:
                pending_remaining = max(0, pending_remaining - dur)
        else:
            # Multi-item or no-duration italic following a bold-with-dur.
            if pending_label is not None and pending_remaining > 0:
                if ',' in text:
                    # Comma-list italic following a bold-with-dur. Karlla1220
                    # always splits when each item has its own (N), using
                    # those literal durations when they fit in the bold's
                    # pending budget. Only equal-splits when the sum of
                    # individuals would overflow (rare; mostly a transcription
                    # artifact where the bold's stated dur is less than the
                    # sum of its sub-items).
                    items_raw = [p.strip() for p in text.split(',') if p.strip()]
                    item_durs = [re.search(r'\(\s*(\d+)\s*\)', it)
                                 for it in items_raw]
                    each_has_dur = (len(items_raw) >= 2 and
                                    all(m is not None for m in item_durs))
                    if each_has_dur:
                        durs = [int(m.group(1)) for m in item_durs]
                        sum_durs = sum(durs)
                        clean_items = [re.sub(r'\s*\(.*?\)\s*', '', it).strip()
                                       for it in items_raw]
                        clean_items = [c for c in clean_items if c]
                        if clean_items and sum_durs <= pending_remaining:
                            # Use the docx's own per-item durations.
                            for it, d in zip(clean_items, durs):
                                emit(it, None, d, pending_label)
                        elif clean_items:
                            # Overflow — equal-split based on available time.
                            per = pending_remaining // len(clean_items)
                            for it in clean_items:
                                emit(it, None, per, pending_label)
                        else:
                            emit(pending_label, text, pending_remaining,
                                 pending_label)
                    else:
                        # No individual durations (e.g. "Draft LS, 9.3.2,
                        # 9.3.3, 9.3.1") — keep as a single session with the
                        # bold label as title and the comma-list as AI items.
                        emit(pending_label, text, pending_remaining,
                             pending_label)
                else:
                    # Single-item italic (e.g. "10.8.1" or "8.2 R19 ISAC CM")
                    # — italic IS the specific topic, so use it as the title;
                    # let emit() derive the AI label from it.
                    emit(text.lstrip('.').strip(), None,
                         pending_remaining, pending_label)
                pending_label = None
                pending_remaining = 0
            else:
                title = text.lstrip('.').strip()
                emit(title, None, period_end - cursor, cat_label or title)
                break

    flush_pending()
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
        col = 0
        for ci, cell in enumerate(cells):
            span = cell_span(cell)
            start_col = col
            col += span
            if ci == 0:
                continue
            if offline:
                # Offline table: 1 time col + 5 days × 2 rooms = 11 grid columns.
                # Map grid column → (day, room) instead of trusting cell index,
                # because cells can have gridSpan>1 (e.g. when a day's two rooms
                # are merged into one empty cell).
                day_idx = (start_col - 1) // 2
                if day_idx >= len(DAYS):
                    break
                day = DAYS[day_idx]
                room_offset = (start_col - 1) % 2
                room = 'Dalian Ballroom 1 (3F)' if room_offset == 0 else 'Shanghai Function room (3F)'
                day_sessions = parse_hiroki_offline_cell(cell, p_start, p_end)
                if day == 'Mon' and this_period == 0:
                    end_align_sessions(day_sessions, p_end)
                for s in day_sessions:
                    s['day'] = day
                    s['room'] = room
                    if span > 1:
                        s['span'] = span
                    sessions.append(s)
            else:
                if day_cells_seen >= len(DAYS):
                    break
                day = DAYS[day_cells_seen]
                day_cells_seen += 1
                day_sessions = parse_hiroki_cell(cell, p_start, p_end)
                if day == 'Mon' and this_period == 0:
                    end_align_sessions(day_sessions, p_end)
                for s in day_sessions:
                    s['day'] = day
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

    # Insertion order matters for the (day, room, start, end) dedup at the
    # end — the FIRST occurrence of a key wins. Add Hiroki online (Ballroom C)
    # and Sorour Ballroom A sessions BEFORE main_sessions so their detailed
    # titles/AI items take precedence over main's terser generic copies.
    for s in hiroki_data['online']:
        out.append(s)
    for s in sorour_sessions:
        if s['room'] == 'Ballroom A (3F)':
            out.append(s)
    for s in main_sessions:
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

    OFFLINE_ROOMS = ('Dalian Ballroom 1 (3F)', 'Shanghai Function room (3F)')

    # Per (day, period, room), only replace when the corresponding parser
    # actually has content for THAT specific room. Otherwise main Shanghai's
    # "host=Sorour" sessions get wiped because Sorour's file has only Dalian
    # content (or none at all) for the same slot.
    hiroki_replace_rooms = {}  # (day, p) -> set of rooms with Hiroki content
    sorour_replace_rooms = {}
    for (day, p), subs in hiroki_off_idx.items():
        rooms = {sub.get('room') for sub in subs if sub.get('room')}
        if rooms:
            hiroki_replace_rooms[(day, p)] = rooms
    for (day, p), subs in sorour_off_idx.items():
        rooms = {sub.get('room') for sub in subs if sub.get('room')}
        if rooms:
            sorour_replace_rooms[(day, p)] = rooms

    keep = []
    for s in out:
        if s['room'] not in OFFLINE_ROOMS:
            keep.append(s)
            continue
        p = find_period(to_min(s['start']))
        host = s.get('host')
        # Drop main's offline session only if its replacement parser actually
        # has content for THIS room (not just for the same (day, period)).
        if p and host == 'Hiroki' and s['room'] in hiroki_replace_rooms.get((s['day'], p), set()):
            continue
        if p and host in ('Sorour', 'Sorouri') and s['room'] in sorour_replace_rooms.get((s['day'], p), set()):
            continue
        keep.append(s)

    # Add Hiroki/Sorour offline sessions exactly once per (day, period, room)
    # they cover, using parser-assigned rooms.
    added_h = set()
    added_s = set()
    for (day, p), rooms in hiroki_replace_rooms.items():
        for sub in hiroki_off_idx[(day, p)]:
            r = sub.get('room')
            if r not in rooms: continue
            key = (day, p, r, sub.get('start'), sub.get('end'), sub.get('title'))
            if key in added_h: continue
            added_h.add(key)
            ns = dict(sub)
            ns['day'] = day
            keep.append(ns)
    for (day, p), rooms in sorour_replace_rooms.items():
        for sub in sorour_off_idx[(day, p)]:
            r = sub.get('room')
            if r not in rooms: continue
            key = (day, p, r, sub.get('start'), sub.get('end'), sub.get('title'))
            if key in added_s: continue
            added_s.add(key)
            ns = dict(sub)
            ns['day'] = day
            keep.append(ns)

    # Drop main's "category-only" sessions ("AI 8", "R20", "6GR", "Coverage",
    # "Sensing"-style) when more detailed sessions from Hiroki/Sorour cover
    # the same (day, room) time range. These generic blocks appear when main
    # had only a category header for the slot but a parser detail-file has
    # the actual sub-session breakdown — visually they'd just stack on top
    # of the detail and look like duplicates.
    def is_generic_title(t):
        if not t:
            return True
        if t in CATEGORIES:
            return True
        # Single-word categories or fuzzy markers that don't carry topic info.
        low = t.lower()
        if low in ('coverage', 'sensing', 'a-iot phase2', 'a-iot',
                   'sweep', 'mimo', 'iot-ntn', 'ntn-iot', 'ntn-nr', 'nr-ntn',
                   'ntn', 'modulation', 'channel coding'):
            return True
        # "X.Y CategoryWord" — main file's terse summaries like "10.8 Sensing".
        if re.match(r'^\d+(\.\d+)*\s+(sensing|coverage|sweep|mimo|waveform|'
                    r'modulation|channel\s+coding|a-iot|iot-ntn|ntn-iot|'
                    r'ntn-nr|nr-ntn|ntn)$',
                    low):
            return True
        # "Sweep" / "Sweep 6GR" / "Sweep R20" — these are bare-prose artifacts
        # from Sorour Table 0 trailing-prose fallback; the actual AI item
        # for that slot lives in another source.
        if re.match(r'^sweep(\s+\S+)?$', low):
            return True
        return False

    def overlaps(a, b):
        return not (to_min(a['end']) <= to_min(b['start']) or
                    to_min(a['start']) >= to_min(b['end']))

    def is_more_detailed(other, s):
        # `other` covers s's range if its [start,end] overlaps and other has
        # a non-generic title (a real topic, not just a category header).
        return (other is not s and other['day'] == s['day'] and
                other['room'] == s['room'] and overlaps(other, s) and
                not is_generic_title(other.get('title')))

    keep = [s for s in keep
            if not (is_generic_title(s.get('title')) and
                    any(is_more_detailed(o, s) for o in keep))]

    # A second pass: drop no-host sessions that overlap with a hosted session
    # in the same room (Hiroki/Sorour/Xiaodong). Main file often has terse
    # summaries ("10.5.4.x" 60m) that overlap with the detail-file's split
    # versions (e.g. "10.5.4.1" 40m + "10.5.4.3" 40m), and the no-host
    # summary should give way to the hosted detail.
    def has_hosted_overlap(s, all_sessions):
        if s.get('host'):
            return False
        for o in all_sessions:
            if o is s:
                continue
            if (o.get('host') and o['day'] == s['day']
                    and o['room'] == s['room'] and overlaps(o, s)):
                return True
        return False

    keep = [s for s in keep if not has_hosted_overlap(s, keep)]

    seen = set()
    dedup = []
    for s in keep:
        # Dedup on (day, room, start, end) — two sessions can't legitimately
        # occupy the same time slot in the same room, so different titles
        # ('6G waveform' vs '10.2.1 Waveform') for the same slot are the
        # same underlying session as transcribed differently across files.
        # Keep the first occurrence (which is the one from the higher-priority
        # source given our insertion order).
        key = (s['day'], s['room'], s['start'], s['end'])
        if key in seen:
            continue
        seen.add(key)
        dedup.append(s)

    # Ballroom A is the room Sorour manages: any session there without an
    # explicit host (typically because the source was the main docx, which
    # doesn't carry host info) should default to Sorour to match karlla's
    # rendering. Sessions with an explicit host (Sorouri Mon AM2, Hiroki, etc.)
    # keep theirs.
    for s in dedup:
        if s['room'] == 'Ballroom A (3F)' and not s.get('host'):
            s['host'] = 'Sorour'

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
