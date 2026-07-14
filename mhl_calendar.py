#!/usr/bin/env python3
"""MHL (Misconduct Hockey League) のスケジュールを iCalendar (.ics) に変換する。

misconduct.co.jp の「NNth SCHEDULE」ページから、その期の
スケジュールファイル（Excel 由来の HTML, cp932）を自動発見してパースし、
ディビジョン別・イベント別・全体の .ics を出力する。

ファイル名や更新タイミングが不定でも動くよう、
- 期は「/NNth-schedule/」リンクの最大番号で自動判定
- スケジュールファイルはページ内リンクから発見（ファイル名に非依存）
- 試合の分類はファイル名ではなくデータ内の「Division」列で行う
という方針にしている。

依存パッケージなし（Python 3.9+ 標準ライブラリのみ）。
"""

from __future__ import annotations

import argparse
import dataclasses
import hashlib
import json
import re
import sys
import urllib.parse
import urllib.request
from datetime import datetime, timedelta, timezone
from html.parser import HTMLParser
from pathlib import Path

BASE = "https://misconduct.co.jp/"
UA = "Mozilla/5.0 (compatible; mhl-calendar/1.0; +https://misconduct.co.jp/)"
JST = timezone(timedelta(hours=9))

# データの「Division」列に現れる値のうち、実在するディビジョン（=リーグ戦）。
# 表記ゆれに備え小文字で保持し、スラッグ化して .ics ファイル名にする。
KNOWN_DIVISIONS = [
    "Platinum", "Gold", "Silver", "Bronze",
    "Brass", "Copper", "Iron", "Women Gold", "35&Over",
]

# col0 が '-' のプログラム行のうち、カレンダーに載せない雑多な行。
PROGRAM_NOISE = {"時間調整", "MHL開催なし", "MHL 開催なし", ""}


def event_type(title: str) -> str:
    """公式イベント（プログラム）名を「種類」に束ねる。未分類は名前そのまま。"""
    t = title.lower()
    if "pick up" in t:
        return "Pick Up Hockey"
    if "drop in" in t:
        return "Drop in Hockey"
    if "open skate" in t:
        return "Open Skate"
    if "clinic" in t or "クリニック" in title:
        return "クリニック"
    if "エキシビション" in title or "exhibition" in t:
        return "エキシビションゲーム"
    if "親子" in title:
        return "親子スケート"
    if "ハピホケ" in title or "happy hockey" in t:
        return "ハピホケ"
    return title


# ---------------------------------------------------------------------------
# HTML → 10 列グリッド
# ---------------------------------------------------------------------------
class _Cell:
    __slots__ = ("text", "colspan", "cls")

    def __init__(self) -> None:
        self.text = ""
        self.colspan = 1
        self.cls = ""


class _TableParser(HTMLParser):
    """Excel 由来 HTML の <tr>/<td> を colspan を保ったまま行の配列に復元する。"""

    def __init__(self) -> None:
        super().__init__()
        self.rows: list[list[_Cell]] = []
        self._cur: list[_Cell] | None = None
        self._cell: _Cell | None = None

    def handle_starttag(self, tag, attrs):
        if tag == "tr":
            self._cur = []
        elif tag in ("td", "th") and self._cur is not None:
            self._cell = _Cell()
            for k, v in attrs:
                if k == "colspan":
                    try:
                        self._cell.colspan = max(1, int(v))
                    except (TypeError, ValueError):
                        pass
                elif k == "class":
                    self._cell.cls = v or ""

    def handle_endtag(self, tag):
        if tag in ("td", "th") and self._cell is not None and self._cur is not None:
            self._cur.append(self._cell)
            self._cell = None
        elif tag == "tr" and self._cur is not None:
            self.rows.append(self._cur)
            self._cur = None

    def handle_data(self, data):
        if self._cell is not None:
            self._cell.text += data


def _grid(html: str) -> list[list[str]]:
    """行ごとに、colspan を空セルに展開した固定幅のセル配列を返す。"""
    parser = _TableParser()
    parser.feed(html)
    grid: list[list[str]] = []
    for row in parser.rows:
        cells: list[str] = []
        for cell in row:
            text = " ".join(cell.text.split()).replace("\xa0", " ").strip()
            cells.append(text)
            cells.extend([""] * (cell.colspan - 1))  # 結合セルの続きは空扱い
        grid.append(cells)
    return grid


# ---------------------------------------------------------------------------
# パース
# ---------------------------------------------------------------------------
@dataclasses.dataclass
class Event:
    kind: str            # "match" | "program"
    date: str            # YYYY-MM-DD
    start: str           # HH:MM
    end: str             # HH:MM
    title: str
    division: str        # match の場合のみ。program は ""
    number: str          # 試合番号（match のみ）
    note: str            # ※延期 など
    source: str          # 取得元 URL

    @property
    def uid(self) -> str:
        if self.kind == "match" and self.number:
            key = f"{self.division}-{self.number}"
        else:
            h = hashlib.sha1(
                f"{self.date}{self.start}{self.title}".encode()
            ).hexdigest()[:12]
            key = f"prog-{h}"
        return f"mhl-{key}@misconduct.co.jp"


_TIME = re.compile(r"^\d{1,2}:\d{2}$")
_DATE = re.compile(r"(\d{4})/(\d{1,2})/(\d{1,2})")
_NUM = re.compile(r"^\d+$")


def _norm_division(value: str) -> str:
    v = value.strip()
    for d in KNOWN_DIVISIONS:
        if v.lower().replace(" ", "") == d.lower().replace(" ", ""):
            return d
    return v


def div_key(div: str):
    """ディビジョンをリーグの格付け順に並べるためのキー。未知（旧期等）は末尾。"""
    try:
        return (0, KNOWN_DIVISIONS.index(div))
    except ValueError:
        return (1, 0, div)


def parse_schedule(html: str, source: str) -> list[Event]:
    """1 つのスケジュールファイルからイベント一覧を取り出す。"""
    events: list[Event] = []
    current_date: str | None = None

    for cells in _grid(html):
        # 10 列に満たない行は末尾を空で埋める
        c = (cells + [""] * 10)[:10]
        num, start, tilde, end, away, mark1, _vs, _mark2, home, division = c

        # 日付ヘッダ行: col0=='#'、col1 に日付
        if num == "#":
            m = _DATE.search(start) or _DATE.search(away)
            if m:
                y, mo, d = map(int, m.groups())
                current_date = f"{y:04d}-{mo:02d}-{d:02d}"
            continue

        # 時刻が揃っていない行（更新日・見出し・「MHL開催なし」等）は対象外。
        # 区切りは全角チルダ '～'(U+FF5E)/波ダッシュ '〜'(U+301C)/半角 '~' の表記ゆれを許容。
        if not (_TIME.match(start) and _TIME.match(end) and tilde in ("～", "〜", "~")):
            continue
        if current_date is None:
            continue

        note = mark1 if mark1.startswith("※") else ""

        # 試合かイベントかは「ホーム欄(col8)が埋まっているか」で判定する。
        # 試合番号は数字とは限らない（'-' や 'PO14'(プレーオフ)もある）ため、
        # 番号での判定は誤分類の原因になる（例: 「ふらら vs Individuals WB」）。
        if home.strip():  # away/home が揃う → 試合
            gid = num if re.fullmatch(r"[A-Za-z0-9]+", num) else ""  # UID 用（'-'は無効）
            events.append(Event(
                kind="match", date=current_date, start=start, end=end,
                title=f"{away} vs {home}", division=_norm_division(division),
                number=gid, note=note, source=source,
            ))
        else:  # ホームなし → プログラム/イベント
            name = away.strip()
            if name in PROGRAM_NOISE or "時間調整" in name:
                continue
            events.append(Event(
                kind="program", date=current_date, start=start, end=end,
                title=name, division="", number="", note="", source=source,
            ))
    return events


# ---------------------------------------------------------------------------
# レンタルリンク表（rent_YYYYMM.htm）— 黄色セル＝非公式・個人イベント
# ---------------------------------------------------------------------------
_RENT_TITLE = re.compile(r"(\d{4})年\s*(\d{1,2})月")
_RENT_FNAME = re.compile(r"rent_(\d{4})(\d{2})", re.I)


def _bg_classes(html: str, color: str) -> set[str]:
    """CSS から background:<color> のクラス名を集める（クラス名はファイル毎に異なる）。"""
    out: set[str] = set()
    for m in re.finditer(r"\.([\w-]+)\s*\{([^}]*)\}", html):
        if re.search(rf"background:\s*{color}\b", m.group(2), re.I):
            out.add(m.group(1))
    return out


def _yellow_classes(html: str) -> set[str]:
    return _bg_classes(html, "yellow")


def _rent_columns(row: list[_Cell]):
    """行を (開始列, colspan, テキスト, クラス) の並びに展開する。"""
    out = []
    col = 0
    for cell in row:
        text = " ".join(cell.text.split()).replace("\xa0", " ").strip()
        out.append((col, cell.colspan, text, cell.cls))
        col += cell.colspan
    return out


def parse_rent(html: str, source: str) -> list[Event]:
    """レンタル表から公式・非公式イベントを取り出す。

    - 黄色セル = 非公式・個人イベント（kind=rental）
    - 青セル かつ 平日(月〜金) = 公式プログラム（kind=program）。
      青は公式全般だが週末はスケジュールファイルと重複するため、
      スケジュールに無い平日（主に金曜の Pick Up）だけを拾う。
    """
    yellow = _yellow_classes(html)
    blue = _bg_classes(html, "blue")
    if not (yellow or blue):
        return []
    parser = _TableParser()
    parser.feed(html)
    rows = [_rent_columns(r) for r in parser.rows]

    # 年月（タイトル優先、なければファイル名）
    ym = None
    for r in rows:
        for _, _, text, _ in r:
            m = _RENT_TITLE.search(text)
            if m:
                ym = (int(m.group(1)), int(m.group(2)))
                break
        if ym:
            break
    if ym is None:
        m = _RENT_FNAME.search(source)
        if not m:
            return []
        ym = (int(m.group(1)), int(m.group(2)))
    year, month = ym

    # 時刻ヘッダ行を探す（6〜24 の整数ラベルが最も多い行）＋ 列→分 の対応
    def hour_labels(r):
        pts = []
        for col, _, text, _ in r:
            if text.isdigit() and 6 <= int(text) <= 24:
                pts.append((col, int(text)))
        return pts
    header = max(rows, key=lambda r: len(hour_labels(r)), default=[])
    labels = hour_labels(header)
    if len(labels) < 2:
        return []
    labels.sort()
    (c0, h0), (c1, h1) = labels[0], labels[-1]
    per_col = (h1 - h0) * 60 / (c1 - c0)  # 1列あたりの分（通常30）

    def col_to_min(col: float) -> int:
        return round(h0 * 60 + (col - c0) * per_col)

    def hhmm(minutes: int) -> str:
        return f"{minutes // 60}:{minutes % 60:02d}"

    events: list[Event] = []
    for r in rows:
        if not r:
            continue
        day_text = r[0][2]
        if not (day_text.isdigit() and 1 <= int(day_text) <= 31):
            continue
        date = f"{year:04d}-{month:02d}-{int(day_text):02d}"
        weekday = datetime(year, month, int(day_text)).weekday()  # 0=月..4=金
        for col, span, text, cls in r:
            if not text:
                continue
            if cls in yellow:
                events.append(Event(
                    kind="rental", date=date,
                    start=hhmm(col_to_min(col)), end=hhmm(col_to_min(col + span)),
                    title=text, division="", number="", note="", source=source,
                ))
            elif cls in blue and weekday < 5:  # 平日の公式プログラム（金曜Pick Up等）
                events.append(Event(
                    kind="program", date=date,
                    start=hhmm(col_to_min(col)), end=hhmm(col_to_min(col + span)),
                    title=text, division="", number="", note="", source=source,
                ))
    return events


# ---------------------------------------------------------------------------
# 取得・発見
# ---------------------------------------------------------------------------
def fetch(url: str) -> bytes:
    req = urllib.request.Request(url, headers={"User-Agent": UA})
    with urllib.request.urlopen(req, timeout=30) as resp:
        return resp.read()


def fetch_text(url: str, encoding: str = "cp932") -> str:
    return fetch(url).decode(encoding, "replace")


def discover_season_page(base: str = BASE) -> tuple[int, str]:
    """トップページから最新の「NNth-schedule」ページ URL を見つける。"""
    html = fetch(base).decode("utf-8", "replace")
    best: tuple[int, str] | None = None
    for m in re.finditer(r'href="([^"]*?/(\d+)(?:st|nd|rd|th)-schedule/?)"', html, re.I):
        url, n = m.group(1), int(m.group(2))
        if not url.startswith("http"):
            url = urllib.parse.urljoin(base, url)
        if best is None or n > best[0]:
            best = (n, url)
    if best is None:
        raise RuntimeError("スケジュールページを発見できませんでした")
    return best


def _discover_uploads(page_url: str, pattern: str) -> list[str]:
    """ページから uploads 配下の .htm リンクを列挙する（順序維持で重複除去）。"""
    html = fetch(page_url).decode("utf-8", "replace")
    urls = re.findall(pattern, html, re.I)
    seen: set[str] = set()
    out: list[str] = []
    for u in urls:
        if u not in seen:
            seen.add(u)
            out.append(u)
    return out


def discover_schedule_files(season_url: str) -> list[str]:
    """期のページから、スケジュール .htm ファイルのリンクを列挙する。"""
    return _discover_uploads(season_url, r'href="(https://[^"]+?/uploads/[^"]+?\.html?)"')


def discover_rent_files(base: str = BASE) -> list[str]:
    """RENT-A-RINK ページから rent_YYYYMM.htm のリンクを列挙する。"""
    try:
        return _discover_uploads(
            urllib.parse.urljoin(base, "rent-a-rink/"),
            r'href="(https://[^"]+?/uploads/rent_\d+\.html?)"')
    except Exception as exc:
        print(f"  ! rent-a-rink: {exc}", file=sys.stderr)
        return []


def collect_events(base: str = BASE) -> tuple[int, list[Event]]:
    season_no, season_url = discover_season_page(base)
    files = discover_schedule_files(season_url)
    if not files:
        raise RuntimeError(f"{season_url} にスケジュールファイルが見つかりません")

    by_uid: dict[str, Event] = {}
    for url in files:
        try:
            html = fetch_text(url)
        except Exception as exc:  # 個別ファイルの失敗は握りつぶして続行
            print(f"  ! {url}: {exc}", file=sys.stderr)
            continue
        for ev in parse_schedule(html, url):
            # 試合番号ベースの UID で、月別ファイルと部門別ファイルの重複を排除
            by_uid.setdefault(ev.uid, ev)

    # レンタル表の黄色ブロック（非公式・個人イベント）
    for url in discover_rent_files(base):
        try:
            html = fetch_text(url)
        except Exception as exc:
            print(f"  ! {url}: {exc}", file=sys.stderr)
            continue
        for ev in parse_rent(html, url):
            by_uid.setdefault(ev.uid, ev)

    events = sorted(by_uid.values(), key=lambda e: (e.date, _hm_min(e.start), e.number or e.title))
    return season_no, events


# ---------------------------------------------------------------------------
# ICS 出力
# ---------------------------------------------------------------------------
def _hm_min(hm: str) -> int:
    """'H:MM' を分に。文字列ソートだと 8:30 が 16:30 の後になるため数値で並べる用。"""
    try:
        h, m = hm.split(":")
        return int(h) * 60 + int(m)
    except (ValueError, AttributeError):
        return 0


def _dt_utc(date: str, hm: str) -> str:
    y, mo, d = map(int, date.split("-"))
    h, mi = map(int, hm.split(":"))
    local = datetime(y, mo, d, h, mi, tzinfo=JST)
    return local.astimezone(timezone.utc).strftime("%Y%m%dT%H%M%SZ")


def _esc(text: str) -> str:
    return (text.replace("\\", "\\\\").replace(";", r"\;")
                .replace(",", r"\,").replace("\n", r"\n"))


def _fold(line: str) -> str:
    """RFC5545 の 75 オクテット折り返し（UTF-8 バイト単位）。"""
    raw = line.encode("utf-8")
    if len(raw) <= 75:
        return line
    out, cur = [], b""
    for ch in line:
        b = ch.encode("utf-8")
        if len(cur) + len(b) > 75:
            out.append(cur)
            cur = b" " + b  # 継続行は先頭スペース
        else:
            cur += b
    out.append(cur)
    return "\r\n".join(s.decode("utf-8") for s in out)


_ICS_HEADER = [
    "BEGIN:VCALENDAR",
    "VERSION:2.0",
    "PRODID:-//mhl-calendar//misconduct.co.jp//JA",
    "CALSCALE:GREGORIAN",
    "METHOD:PUBLISH",
]


def render_vevent(ev: Event, location: str, dtstamp: str) -> str:
    """1 件の VEVENT ブロックを RFC5545 折り返し済みで返す（末尾改行なし）。"""
    summary = ev.title
    if ev.kind == "match" and ev.division:
        summary = f"{ev.title}（{ev.division}）"
    if ev.note:
        summary = f"⚠{ev.note} {summary}"
    desc_parts = []
    if ev.number:
        desc_parts.append(f"試合 #{ev.number}")
    if ev.division:
        desc_parts.append(f"ディビジョン: {ev.division}")
    if ev.note:
        desc_parts.append(ev.note)
    desc_parts.append(ev.source)
    lines = [
        "BEGIN:VEVENT",
        f"UID:{ev.uid}",
        f"DTSTAMP:{dtstamp}",
        f"DTSTART:{_dt_utc(ev.date, ev.start)}",
        f"DTEND:{_dt_utc(ev.date, ev.end)}",
        f"SUMMARY:{_esc(summary)}",
        f"DESCRIPTION:{_esc(' / '.join(desc_parts))}",
        f"LOCATION:{_esc(location)}",
        "END:VEVENT",
    ]
    return "\r\n".join(_fold(ln) for ln in lines)


def render_ics(events: list[Event], name: str, location: str, dtstamp: str) -> str:
    lines = list(_ICS_HEADER) + [f"X-WR-CALNAME:{_esc(name)}", "X-WR-TIMEZONE:Asia/Tokyo"]
    head = "\r\n".join(_fold(ln) for ln in lines)
    body = "\r\n".join(render_vevent(ev, location, dtstamp) for ev in events)
    parts = [head] + ([body] if body else []) + ["END:VCALENDAR"]
    return "\r\n".join(parts) + "\r\n"


INDEX_CSS = """
:root{color-scheme:light dark;--bg:#fafafa;--fg:#1a1a1a;--card:#fff;--line:#e5e5e5;--accent:#c0392b;--muted:#666}
@media(prefers-color-scheme:dark){:root{--bg:#161616;--fg:#eee;--card:#1f1f1f;--line:#333;--accent:#ff6b5e;--muted:#999}}
*{box-sizing:border-box}
body{margin:0;font-family:system-ui,-apple-system,"Hiragino Sans","Noto Sans JP",sans-serif;background:var(--bg);color:var(--fg);line-height:1.6}
.wrap{max-width:720px;margin:0 auto;padding:2rem 1.2rem 4rem}
h1{font-size:1.5rem;margin:0 0 .3rem}
.sub{color:var(--muted);margin:0 0 2rem;font-size:.9rem}
.card{background:var(--card);border:1px solid var(--line);border-radius:12px;padding:1rem 1.1rem;margin:.6rem 0;display:flex;align-items:center;gap:1rem;flex-wrap:wrap}
.card .name{font-weight:600;flex:1;min-width:9rem}
.card .count{color:var(--muted);font-size:.85rem}
.card .copy{font-size:.85rem;font-weight:600;padding:.45rem .9rem;border-radius:8px;white-space:nowrap;cursor:pointer;font-family:inherit;background:var(--accent);color:#fff;border:1px solid var(--accent)}
.card .copy:hover{opacity:.85}
.card .copy.done{background:transparent;color:var(--accent)}
.sec{font-size:1.05rem;margin:2.2rem 0 .4rem;padding-bottom:.3rem;border-bottom:2px solid var(--accent)}
details{background:var(--card);border:1px solid var(--line);border-radius:12px;margin:.6rem 0;padding:.2rem .4rem}
details>summary{cursor:pointer;font-weight:600;padding:.7rem}
details[open]>summary{border-bottom:1px solid var(--line);margin-bottom:.4rem}
details .card{border:none;border-bottom:1px solid var(--line);border-radius:0;margin:0;background:transparent}
details .card:last-child{border-bottom:none}
.how{margin-top:2.5rem;font-size:.9rem;color:var(--muted)}
.how code{background:var(--line);padding:.1rem .35rem;border-radius:4px;font-size:.85em}
.how ol{padding-left:1.2rem}
.how .note{margin-top:1rem;padding:.8rem 1rem;background:var(--card);border:1px solid var(--line);border-left:3px solid var(--accent);border-radius:8px}
.foot{margin-top:2rem;font-size:.8rem;color:var(--muted)}
.foot a{color:var(--accent)}
.filters{display:flex;flex-wrap:wrap;gap:.5rem .6rem;align-items:center;margin:.4rem 0 1rem}
.filters select,.filters input[type=text]{font:inherit;font-size:.85rem;padding:.4rem .55rem;border:1px solid var(--line);border-radius:8px;background:var(--card);color:var(--fg)}
.filters input[type=text]{min-width:11rem}
.filters label{font-size:.85rem;color:var(--muted);display:inline-flex;align-items:center;gap:.3rem;cursor:pointer}
.filters .reset{font:inherit;font-size:.8rem;padding:.35rem .6rem;border:1px solid var(--line);border-radius:8px;background:transparent;color:var(--muted);cursor:pointer}
#count{margin-left:auto;font-size:.85rem;color:var(--fg);font-weight:600;white-space:nowrap}
.tablewrap{overflow-x:auto;border:1px solid var(--line);border-radius:12px;margin-bottom:1.5rem;max-height:70vh;overflow-y:auto}
table#sched{border-collapse:collapse;width:100%;font-size:.88rem}
#sched th,#sched td{padding:.5rem .7rem;text-align:left;border-bottom:1px solid var(--line);vertical-align:top}
#sched thead th{position:sticky;top:0;background:var(--card);z-index:1;font-size:.78rem;color:var(--muted);white-space:nowrap}
#sched tbody tr:hover{background:rgba(192,57,43,.07)}
#sched .dt,#sched .tm{white-space:nowrap}
#sched td.match{min-width:14rem}
#sched .vs{color:var(--muted);font-size:.8rem;margin:0 .35rem}
#sched .cat{color:var(--muted);font-size:.82rem;white-space:nowrap}
#sched .delay{background:var(--accent);color:#fff;font-size:.7rem;padding:.05rem .35rem;border-radius:4px;margin-right:.3rem;white-space:nowrap}
#sched .rtag{background:var(--line);color:var(--muted);font-size:.68rem;padding:.05rem .3rem;border-radius:4px;margin-right:.3rem}
#sched tr.we .dt{color:var(--accent);font-weight:700}
#sched .empty td{color:var(--muted);text-align:center;padding:1.5rem}
.fpanel{border:1px solid var(--line);border-radius:12px;padding:.8rem 1rem;margin:.4rem 0 1rem;background:var(--card)}
.fgroup{padding:.5rem 0;border-bottom:1px solid var(--line)}
.fgroup:last-child{border-bottom:none}
.flabel{font-size:.78rem;font-weight:700;color:var(--muted);margin-bottom:.4rem;letter-spacing:.02em}
.chips{display:flex;flex-wrap:wrap;gap:.35rem}
.chip{display:inline-flex;align-items:center;gap:.3rem;font-size:.82rem;padding:.3rem .6rem;border:1px solid var(--line);border-radius:999px;cursor:pointer;user-select:none;background:var(--bg)}
.chip:hover{border-color:var(--accent)}
.chip:has(input:checked){background:var(--accent);color:#fff;border-color:var(--accent)}
.chip input{margin:0}
.frow{display:flex;flex-wrap:wrap;gap:.5rem;align-items:center}
.frow .hint{font-size:.78rem;color:var(--muted)}
.frow select,.frow input[type=text]{font:inherit;font-size:.85rem;padding:.35rem .55rem;border:1px solid var(--line);border-radius:8px;background:var(--bg);color:var(--fg)}
.cat{border:1px solid var(--line);border-radius:10px;margin:.5rem 0;background:var(--bg)}
.cat>summary{cursor:pointer;font-weight:700;padding:.6rem .8rem;display:flex;align-items:center;gap:.5rem;list-style:none}
.cat>summary::-webkit-details-marker{display:none}
.cat>summary::before{content:"▸";color:var(--muted);font-weight:400}
.cat[open]>summary::before{content:"▾"}
.cat>summary .count{color:var(--muted);font-weight:400;font-size:.8rem}
.catbody{padding:.2rem .9rem .8rem 1.6rem}
.cat input[type=checkbox]{width:16px;height:16px;cursor:pointer;flex:0 0 auto}
details.tg{margin:.15rem 0;border-bottom:1px solid var(--line)}
details.tg:last-child{border-bottom:none}
details.tg>summary{cursor:pointer;font-size:.85rem;padding:.35rem 0;display:flex;align-items:center;gap:.4rem;list-style:none}
details.tg>summary::-webkit-details-marker{display:none}
details.tg>summary::before{content:"▸";color:var(--muted);font-size:.75rem}
details.tg[open]>summary::before{content:"▾"}
details.tg>summary .nm{font-weight:600}
details.tg>summary .count{color:var(--muted);font-weight:400;font-size:.78rem}
details.tg .divchk{width:15px;height:15px}
details.tg .chips{padding:.4rem 0 .6rem 1.4rem}
.subbox{border:1px solid var(--line);border-radius:12px;padding:.9rem 1rem;background:var(--card);margin:.5rem 0 1rem}
.subbox code#feedurl{display:block;word-break:break-all;font-size:.82rem;background:var(--bg);border:1px solid var(--line);border-radius:8px;padding:.6rem .7rem;color:var(--fg);margin-bottom:.7rem}
.subbtns{display:flex;gap:.6rem;flex-wrap:wrap}
.copy2,.gcopen{font:inherit;font-size:.85rem;font-weight:600;padding:.5rem .9rem;border-radius:8px;cursor:pointer;text-decoration:none;white-space:nowrap}
.copy2{background:var(--accent);color:#fff;border:1px solid var(--accent)}
.copy2:hover{opacity:.85}
.copy2.done{background:transparent;color:var(--accent)}
.gcopen{background:transparent;color:var(--accent);border:1px solid var(--accent)}
.gcopen:hover{background:var(--accent);color:#fff}
.cat>summary .ic,details.tg>summary .ic{font-size:1rem;line-height:1}
.help{border:1px solid var(--line);border-radius:10px;margin:.3rem 0 .8rem;background:var(--card)}
.help>summary{cursor:pointer;font-weight:600;font-size:.9rem;padding:.6rem .8rem;list-style:none;color:var(--accent)}
.help>summary::-webkit-details-marker{display:none}
.helpbody{padding:.2rem 1rem 1rem;font-size:.88rem}
.helpbody ol{padding-left:1.2rem;margin:.5rem 0}
.helpbody .note{margin-top:.6rem;padding:.7rem .8rem;background:var(--bg);border-left:3px solid var(--accent);border-radius:8px;font-size:.82rem;color:var(--muted)}
.helpbody .foot{margin:.7rem 0 0;font-size:.75rem;color:var(--muted)}
.subhint{margin:.6rem 0 0;font-size:.8rem;color:var(--muted)}
@media (max-width:640px){
  .wrap{padding:1rem .7rem 3rem}
  h1{font-size:1.2rem;margin-bottom:.2rem}
  .fpanel{padding:.5rem .6rem}
  .cat>summary{padding:.55rem .6rem;font-size:.92rem}
  .catbody{padding:.2rem .4rem .7rem .8rem}
  .chip{font-size:.8rem;padding:.28rem .55rem}
  .tablewrap{max-height:56vh}
  #sched th,#sched td{padding:.4rem .5rem;font-size:.82rem}
  .sec{font-size:1rem;margin:1.4rem 0 .4rem}
  .frow{gap:.4rem}
  .frow select,.frow input[type=text]{flex:1 1 8rem;min-width:0}
}
"""

# フィルタ＋テーブル描画＋購読URL生成の JS。__DATA__ を試合データ JSON に置換して埋め込む。
# FEED（フィードのベースURL）は別途 `var FEED=...` を前置きして注入する。
INDEX_JS = r"""
var DATA = __DATA__;
if(typeof FEED==='undefined') var FEED='';
var WD = ['日','月','火','水','木','金','土'];
function esc(s){return String(s).replace(/[&<>]/g,function(c){return{'&':'&amp;','<':'&lt;','>':'&gt;'}[c];});}
function wday(d){var p=d.split('-');return new Date(+p[0],+p[1]-1,+p[2]).getDay();}
function fmtDate(d){var p=d.split('-');return p[1]+'/'+p[2]+' ('+WD[wday(d)]+')';}
function qsa(sel){return Array.prototype.slice.call(document.querySelectorAll(sel));}

var cHide=document.getElementById('c-hide'),
    cPast=document.getElementById('c-past'),
    fMonth=document.getElementById('f-month'),
    fKw=document.getElementById('f-kw'),
    tbody=document.getElementById('rows'),
    count=document.getElementById('count'),
    feedEl=document.getElementById('feedurl'),
    allEvents=document.getElementById('all-events'),
    allRent=document.getElementById('all-rent');

// 今日(ローカル=JST想定)を YYYY-MM-DD で。既定は今日以降のみ表示。
var TODAY=(function(){var d=new Date();function p(n){return (n<10?'0':'')+n;}
  return d.getFullYear()+'-'+p(d.getMonth()+1)+'-'+p(d.getDate());})();

// ---- 選択状態（購読対象。月・キーワード・延期は含めない）----
function selection(){
  var teams=qsa('#g-matches .teamchk').filter(function(c){return c.checked;}).map(function(c){return c.value;});
  var et=qsa('#g-events .etchk').filter(function(c){return c.checked;}).map(function(c){return c.value;});
  var etTotal=qsa('#g-events .etchk').length;
  var rl=qsa('#g-rent .rlchk').filter(function(c){return c.checked;}).map(function(c){return c.value;});
  var rlTotal=qsa('#g-rent .rlchk').length;
  return {teams:teams,
          et:et, allEt:(etTotal>0 && et.length===etTotal),
          rl:rl, allRl:(rlTotal>0 && rl.length===rlTotal),
          hide:cHide?cHide.checked:false};
}

function inSubscription(r, s){
  var any = s.teams.length || s.et.length || s.rl.length;
  if(!any) return true;                       // 何も選ばなければ全部
  if(r.k==='p') return s.et.indexOf(r.et)>=0;
  if(r.k==='r') return s.rl.indexOf(r.et)>=0;
  return s.teams.indexOf(r.a)>=0 || s.teams.indexOf(r.h)>=0;
}

function render(){
  var s=selection(), mo=fMonth?fMonth.value:'', kw=fKw?fKw.value.trim().toLowerCase():'';
  var html='', n=0;
  for(var i=0;i<DATA.length;i++){
    var r=DATA[i];
    if(!inSubscription(r,s)) continue;
    if(!(cPast&&cPast.checked) && r.d < TODAY) continue;  // 既定は今日以降
    if(s.hide && r.n) continue;
    if(mo && r.d.slice(0,7)!==mo) continue;
    if(kw && (r.a+' '+r.h).toLowerCase().indexOf(kw)<0) continue;
    n++;
    var badge = r.n ? '<span class="delay">延期</span>' : '';
    var mu = r.k==='m' ? esc(r.a)+'<span class="vs">vs</span>'+esc(r.h) : esc(r.a);
    var cat = r.k==='m' ? esc(r.dv)
            : (r.k==='r' ? '<span class="rtag">個人</span>'+esc(r.et||'') : esc(r.et||'イベント'));
    var we = (wday(r.d)===0||wday(r.d)===6) ? ' class="we"' : '';
    html += '<tr'+we+'><td class="dt">'+fmtDate(r.d)+'</td><td class="tm">'+r.s+'–'+r.e
          + '</td><td class="match">'+badge+mu+'</td><td class="cat">'+cat+'</td></tr>';
  }
  if(!n) html='<tr class="empty"><td colspan="4">該当する試合・イベントがありません</td></tr>';
  tbody.innerHTML=html;
  count.textContent=n+' 件';
  updateFeed(s);
}

function buildFeed(s){
  if(!FEED) return '';
  var base = FEED.replace(/\/+$/,'') + '/calendar.ics';
  var p=[];
  if(s.teams.length) p.push('teams='+s.teams.map(encodeURIComponent).join(','));
  if(s.et.length){ if(s.allEt) p.push('events=1');
                   else p.push('etypes='+s.et.map(encodeURIComponent).join(',')); }
  if(s.rl.length){ if(s.allRl) p.push('rent=1');
                   else p.push('rlabels='+s.rl.map(encodeURIComponent).join(',')); }
  if(s.hide) p.push('hide=1');
  return base + (p.length ? '?'+p.join('&') : '');
}
function updateFeed(s){
  if(!feedEl) return;
  var url=buildFeed(s);
  feedEl.textContent = url || '(フィード未設定)';
  feedEl.dataset.url = url;
}

// ---- 親（全選択）↔子チェックの連動（tri-state）----
function syncParent(parent, sel){
  if(!parent) return;
  var e=qsa(sel), on=e.filter(function(c){return c.checked;}).length;
  parent.checked = on===e.length && e.length>0;
  parent.indeterminate = on>0 && on<e.length;
}
function onChange(t){
  if(t===allEvents){
    qsa('#g-events .etchk').forEach(function(c){c.checked=allEvents.checked;});
  } else if(t===allRent){
    qsa('#g-rent .rlchk').forEach(function(c){c.checked=allRent.checked;});
  } else if(t.classList.contains('etchk')){
    syncParent(allEvents,'#g-events .etchk');
  } else if(t.classList.contains('rlchk')){
    syncParent(allRent,'#g-rent .rlchk');
  }
  render();
}

function bindAll(){
  document.addEventListener('change', function(e){
    var t=e.target;
    if(t && (t===allEvents||t===allRent||t.classList.contains('teamchk')
        ||t.classList.contains('etchk')||t.classList.contains('rlchk')
        ||t===cHide||t===cPast||t===fMonth)) onChange(t);
  });
  // summary 内のチェックボックスは details の開閉を起こさない
  document.addEventListener('click', function(e){
    if(e.target && e.target.matches('summary input[type=checkbox]')) e.stopPropagation();
  }, true);
  if(fKw) fKw.addEventListener('input', render);
  var reset=document.getElementById('f-reset');
  if(reset) reset.addEventListener('click', function(){
    qsa('.fpanel input[type=checkbox]').forEach(function(x){x.checked=false;x.indeterminate=false;});
    if(fMonth)fMonth.value=''; if(fKw)fKw.value='';
    render();
  });
}

function attachCopy(btn, getUrl){
  if(!btn) return;
  btn.addEventListener('click', function(){
    var url=getUrl(), label=btn.textContent;
    if(!url) return;
    function done(){ btn.textContent='✓ コピーしました'; btn.classList.add('done');
      setTimeout(function(){ btn.textContent=label; btn.classList.remove('done'); },1500); }
    if(navigator.clipboard&&navigator.clipboard.writeText){
      navigator.clipboard.writeText(url).then(done,function(){ window.prompt('このURLをコピーしてください',url); });
    } else { window.prompt('このURLをコピーしてください',url); }
  });
}

bindAll();
attachCopy(document.getElementById('feed-copy'), function(){ return feedEl?feedEl.dataset.url:''; });
render();
"""


def write_index(out: Path, specs: list[CalSpec], season_no: int, base_url: str,
                dtstamp: str, feed_url: str = "") -> None:
    """スケジュール表＋絞り込み＋『選択を購読』ページ（index.html）を書き出す。"""
    import html as _h

    # ---- 試合・イベントのテーブル用データ（最小粒度：1行1試合）----
    matches = next((s.events for s in specs if s.category == "overview"), [])
    programs = next((s.events for s in specs if s.category == "events"), [])
    rentals = next((s.events for s in specs if s.category == "rentals"), [])
    rows = []
    for e in matches:
        away, home = (e.title.split(" vs ", 1) + [""])[:2]
        rows.append({"d": e.date, "s": e.start, "e": e.end, "a": away, "h": home,
                     "dv": e.division, "n": 1 if e.note else 0, "k": "m"})
    for e in programs:
        rows.append({"d": e.date, "s": e.start, "e": e.end, "a": e.title, "h": "",
                     "dv": "", "n": 0, "k": "p", "et": event_type(e.title)})
    for e in rentals:
        rows.append({"d": e.date, "s": e.start, "e": e.end, "a": e.title, "h": "",
                     "dv": "", "n": 0, "k": "r", "et": e.title})
    rows.sort(key=lambda r: (r["d"], _hm_min(r["s"])))
    data_json = json.dumps(rows, ensure_ascii=False, separators=(",", ":"))

    def esc(s: str) -> str:
        return _h.escape(s)

    # 階層: 試合 → ディビジョン → チーム ／ 公式イベント → 種類 ／ 非公式・個人イベント → ラベル
    div_names = sorted({r["dv"] for r in rows if r["dv"]}, key=div_key)
    months = sorted({r["d"][:7] for r in rows})
    team_by_div: dict[str, set[str]] = {}
    for r in rows:
        if r["k"] == "m":
            for t in (r["a"], r["h"]):
                if t:
                    team_by_div.setdefault(r["dv"], set()).add(t)
    event_types = sorted({r["et"] for r in rows if r["k"] == "p" and r.get("et")})
    rent_labels = sorted({r["et"] for r in rows if r["k"] == "r" and r.get("et")})

    # 試合ツリー：ディビジョンは見出しのみ（丸ごと購読は使わない）、チームだけ選択可
    match_tree = ""
    for d in div_names:
        teams = sorted(team_by_div.get(d, ()))
        team_chips = "".join(
            f'<label class="chip"><input type="checkbox" class="teamchk" '
            f'data-div="{esc(d)}" value="{esc(t)}">{esc(t)}</label>'
            for t in teams
        )
        match_tree += (
            f'<details class="tg"><summary>'
            f'<span class="nm">{esc(d)}</span><span class="count">{len(teams)}チーム</span>'
            f'</summary><div class="chips">{team_chips}</div></details>'
        )

    event_chips = "".join(
        f'<label class="chip"><input type="checkbox" class="etchk" value="{esc(et)}">{esc(et)}</label>'
        for et in event_types
    )
    rent_chips = "".join(
        f'<label class="chip"><input type="checkbox" class="rlchk" value="{esc(rl)}">{esc(rl)}</label>'
        for rl in rent_labels
    )
    rent_cat = f"""
  <details class="cat">
    <summary><input type="checkbox" id="all-rent"><span class="ic">👥</span><span class="nm">非公式・個人イベント</span>
      <span class="count">{len(rent_labels)}件</span></summary>
    <div class="catbody chips" id="g-rent">{rent_chips}</div>
  </details>""" if rent_labels else ""
    month_opts = "".join(f'<option value="{m}">{m[:4]}/{m[5:7]}</option>' for m in months)

    filter_ui = f"""
<div class="fpanel">
  <details class="cat">
    <summary><span class="ic">🏒</span><span class="nm">試合</span>
      <span class="count">チーム別（{len(div_names)}部門）</span></summary>
    <div class="catbody" id="g-matches">{match_tree}</div>
  </details>
  <details class="cat">
    <summary><input type="checkbox" id="all-events"><span class="ic">📣</span><span class="nm">公式イベント</span>
      <span class="count">{len(event_types)}種類</span></summary>
    <div class="catbody chips" id="g-events">{event_chips}</div>
  </details>{rent_cat}
  <div class="fgroup viewopts"><div class="frow">
    <label class="chip"><input type="checkbox" id="c-past"> 過去も表示<span class="hint">（表のみ）</span></label>
    <label class="chip"><input type="checkbox" id="c-hide"> 延期を除く<span class="hint">（購読も）</span></label>
    <select id="f-month" title="月で絞り込み（表のみ）"><option value="">全期間</option>{month_opts}</select>
    <input type="text" id="f-kw" placeholder="🔍 チーム名など（表のみ）">
    <button type="button" class="reset" id="f-reset">クリア</button>
    <span id="count"></span>
  </div>
</div>
<div class="tablewrap"><table id="sched">
<thead><tr><th>日付</th><th>時刻</th><th>対戦 / 内容</th><th>区分</th></tr></thead>
<tbody id="rows"></tbody></table></div>"""

    # 「選択を購読」パネル。feed_url が未設定なら注意書きを出す。
    if feed_url:
        subscribe_panel = """
<h2 class="sec">📅 選択を購読</h2>
<div class="subbox">
  <code id="feedurl"></code>
  <div class="subbtns">
    <button type="button" class="copy2" id="feed-copy">このURLをコピー</button>
    <a class="gcopen" href="https://calendar.google.com/" target="_blank" rel="noopener">Googleを開く</a>
  </div>
  <p class="subhint">コピー → Googleカレンダー「他のカレンダー ＋ → <b>URLで追加</b>」に貼付。手順は上の <b>❓使い方</b> を参照。</p>
</div>"""
    else:
        subscribe_panel = """
<h2 class="sec">📅 選択を購読（準備中）</h2>
<p class="note">購読フィード（Cloudflare Worker）が未設定です。</p>"""

    feed_js = "var FEED=" + json.dumps(feed_url) + ";\n"
    stamp = f"{dtstamp[:4]}-{dtstamp[4:6]}-{dtstamp[6:8]} {dtstamp[9:11]}:{dtstamp[11:13]} UTC"
    body = f"""<!doctype html><html lang="ja"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>MHL スケジュール（{season_no}期）</title><style>{INDEX_CSS}</style></head>
<body><div class="wrap">
<h1>🏒 MHL スケジュール</h1>
<details class="help"><summary>❓ 使い方・Googleカレンダー登録方法</summary>
<div class="helpbody">
<p>見たい<b>チーム／公式イベント／非公式・個人イベント</b>をチェック → 下の「このURLをコピー」→ Googleカレンダーで購読すると、選んだ内容が1本のURLになり<b>自動更新</b>されます。</p>
<ol>
<li>フィルタで見たいチーム／イベントをチェック</li>
<li>「このURLをコピー」を押す</li>
<li>PCの Google カレンダー左「他のカレンダー ＋」→「<b>URL で追加</b>」に貼り付け</li>
</ol>
<p class="note">購読は Google カレンダーの<b>「URL で追加」</b>で行います。反映は Google 側の都合で数時間〜1日遅れることがあります。延期試合はタイトル先頭に <b>⚠延期</b>。<b>月・キーワード・過去表示は表の閲覧用</b>で購読には反映されません。スマホは一度PCで登録すれば同期されます。</p>
<p class="foot">MHL {season_no}期 / 最終更新 {stamp} / データ元: <a href="{esc(BASE)}">misconduct.co.jp</a>（非公式）</p>
</div></details>
{filter_ui}
{subscribe_panel}
</div>
<script>{feed_js}{INDEX_JS.replace("__DATA__", data_json)}</script>
</body></html>"""
    (out / "index.html").write_text(body, encoding="utf-8")


@dataclasses.dataclass
class CalSpec:
    filename: str
    name: str
    events: list[Event]
    category: str          # "overview" | "division" | "team" | "events"
    division: str = ""     # team/division の所属（index のグループ化用）


def slug(text: str) -> str:
    return re.sub(r"[^a-z0-9]", "", text.lower().replace("&", "and"))


def _team_slug(name: str, used: set[str]) -> str:
    """チーム名から一意で安定なファイル名スラッグを作る（日本語名はハッシュ）。"""
    base = "team-" + slug(name)
    if base == "team-":  # ASCII 成分がない（日本語名など）
        base = "team-" + hashlib.sha1(name.encode()).hexdigest()[:8]
    candidate, i = base, 2
    while candidate in used:  # 万一の衝突回避
        candidate, i = f"{base}-{i}", i + 1
    used.add(candidate)
    return candidate


def _teams(matches: list[Event]) -> dict[str, tuple[str, list[Event]]]:
    """チーム名 → (所属ディビジョン, その試合一覧)。away/home 両方を拾う。"""
    teams: dict[str, list[Event]] = {}
    div_of: dict[str, str] = {}
    for e in matches:
        away, home = e.title.split(" vs ", 1)
        for t in (away.strip(), home.strip()):
            teams.setdefault(t, []).append(e)
            div_of.setdefault(t, e.division)
    return {t: (div_of[t], evs) for t, evs in teams.items()}


def build_calendars(events: list[Event]) -> list[CalSpec]:
    """出力する .ics の仕様一覧を返す。"""
    matches = [e for e in events if e.kind == "match"]
    programs = [e for e in events if e.kind == "program"]
    rentals = [e for e in events if e.kind == "rental"]

    specs: list[CalSpec] = [CalSpec("all.ics", "MHL 全試合", matches, "overview")]

    for div in sorted({e.division for e in matches if e.division}, key=div_key):
        specs.append(CalSpec(
            f"{slug(div)}.ics", f"MHL {div}",
            [e for e in matches if e.division == div], "division", div,
        ))

    used = {s.filename[:-4] for s in specs}
    # 所属ディビジョン→チーム名の順で安定に並べる
    teams = _teams(matches)
    for name in sorted(teams, key=lambda n: (teams[n][0], n)):
        div, evs = teams[name]
        specs.append(CalSpec(
            f"{_team_slug(name, used)}.ics", name, evs, "team", div,
        ))

    if programs:
        specs.append(CalSpec("events.ics", "MHL イベント・クリニック", programs, "events"))
    if rentals:
        specs.append(CalSpec("rentals.ics", "非公式・個人イベント", rentals, "rentals"))
    return specs


def write_feed(out: Path, specs: list[CalSpec], season_no: int,
               location: str, dtstamp: str) -> int:
    """Cloudflare Worker が絞り込みに使う feed.json を書き出す。

    各イベントに、フィルタ用フィールド（dv/a/h/k/n）と、
    そのまま連結できる事前生成済み VEVENT ブロックを持たせる。
    """
    matches = next((s.events for s in specs if s.category == "overview"), [])
    programs = next((s.events for s in specs if s.category == "events"), [])
    rentals = next((s.events for s in specs if s.category == "rentals"), [])
    items = []
    for e in matches:
        away, home = (e.title.split(" vs ", 1) + [""])[:2]
        items.append({"dv": e.division, "a": away, "h": home, "k": "m",
                      "n": 1 if e.note else 0,
                      "ev": render_vevent(e, location, dtstamp)})
    for e in programs:
        items.append({"dv": "", "a": e.title, "h": "", "k": "p", "n": 0,
                      "et": event_type(e.title),
                      "ev": render_vevent(e, location, dtstamp)})
    for e in rentals:
        items.append({"dv": "", "a": e.title, "h": "", "k": "r", "n": 0,
                      "et": e.title,
                      "ev": render_vevent(e, location, dtstamp)})
    feed = {
        "season": season_no,
        "dtstamp": dtstamp,
        "calname": f"MHL {season_no}期",
        "header": _ICS_HEADER + ["X-WR-TIMEZONE:Asia/Tokyo"],
        "events": items,
    }
    (out / "feed.json").write_text(
        json.dumps(feed, ensure_ascii=False, separators=(",", ":")), encoding="utf-8")
    return len(items)


# ---------------------------------------------------------------------------
# CLI
# ---------------------------------------------------------------------------
def main() -> int:
    ap = argparse.ArgumentParser(description="MHL スケジュール → iCalendar 変換")
    ap.add_argument("-o", "--out", default="calendars", help="出力ディレクトリ")
    ap.add_argument("--location", default="MHL TOKYO", help="LOCATION に入れる会場名")
    ap.add_argument("--base", default=BASE, help="サイトのベース URL")
    ap.add_argument("--base-url", default="",
                    help="公開先の URL（例: https://user.github.io/repo）")
    ap.add_argument("--feed-url", default="",
                    help="Cloudflare Worker のベース URL（例: https://mhl.<sub>.workers.dev）。"
                         "index.html の『選択を購読』に使う")
    ap.add_argument("--no-index", action="store_true", help="index.html を生成しない")
    args = ap.parse_args()

    # DTSTAMP は生成時刻。差分ノイズを避けるため分単位に丸める。
    dtstamp = datetime.now(timezone.utc).strftime("%Y%m%dT%H%M00Z")

    print(f"discover season page from {args.base} ...", file=sys.stderr)
    season_no, events = collect_events(args.base)
    print(f"season {season_no}: {len(events)} events", file=sys.stderr)

    out = Path(args.out)
    out.mkdir(parents=True, exist_ok=True)
    specs = build_calendars(events)
    by_cat: dict[str, int] = {}
    for spec in specs:
        ics = render_ics(spec.events, f"{spec.name}（{season_no}期）", args.location, dtstamp)
        (out / spec.filename).write_text(ics, encoding="utf-8")
        by_cat[spec.category] = by_cat.get(spec.category, 0) + 1
    for cat in ("overview", "division", "team", "events"):
        if by_cat.get(cat):
            print(f"  {cat:10} {by_cat[cat]:3d} calendars", file=sys.stderr)
    n_feed = write_feed(out, specs, season_no, args.location, dtstamp)
    print(f"  feed.json  {n_feed:3d} events", file=sys.stderr)
    if not args.no_index:
        write_index(out, specs, season_no, args.base_url, dtstamp, args.feed_url)
        print("  index.html", file=sys.stderr)
    print(f"wrote {len(specs)} calendars to {out}/", file=sys.stderr)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
