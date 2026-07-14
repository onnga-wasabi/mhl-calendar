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


# ---------------------------------------------------------------------------
# HTML → 10 列グリッド
# ---------------------------------------------------------------------------
class _Cell:
    __slots__ = ("text", "colspan")

    def __init__(self) -> None:
        self.text = ""
        self.colspan = 1


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

        if _NUM.match(num):  # 試合番号あり → リーグ戦
            events.append(Event(
                kind="match", date=current_date, start=start, end=end,
                title=f"{away} vs {home}", division=_norm_division(division),
                number=num, note=note, source=source,
            ))
        else:  # col0=='-' → プログラム/イベント
            name = away.strip()
            if name in PROGRAM_NOISE or "時間調整" in name:
                continue
            events.append(Event(
                kind="program", date=current_date, start=start, end=end,
                title=name, division="", number="", note="", source=source,
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


def discover_schedule_files(season_url: str) -> list[str]:
    """期のページから、スケジュール .htm ファイルのリンクを列挙する。"""
    html = fetch(season_url).decode("utf-8", "replace")
    urls = re.findall(r'href="(https://[^"]+?/uploads/[^"]+?\.html?)"', html, re.I)
    # 重複除去（順序維持）
    seen: set[str] = set()
    out: list[str] = []
    for u in urls:
        if u not in seen:
            seen.add(u)
            out.append(u)
    return out


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
    events = sorted(by_uid.values(), key=lambda e: (e.date, e.start, e.number or e.title))
    return season_no, events


# ---------------------------------------------------------------------------
# ICS 出力
# ---------------------------------------------------------------------------
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


def render_ics(events: list[Event], name: str, location: str, dtstamp: str) -> str:
    lines = [
        "BEGIN:VCALENDAR",
        "VERSION:2.0",
        "PRODID:-//mhl-calendar//misconduct.co.jp//JA",
        "CALSCALE:GREGORIAN",
        "METHOD:PUBLISH",
        f"X-WR-CALNAME:{_esc(name)}",
        "X-WR-TIMEZONE:Asia/Tokyo",
    ]
    for ev in events:
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
        lines += [
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
    lines.append("END:VCALENDAR")
    return "\r\n".join(_fold(ln) for ln in lines) + "\r\n"


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
.card .gcal,.card .copy{text-decoration:none;font-size:.82rem;font-weight:600;padding:.45rem .8rem;border-radius:8px;white-space:nowrap;cursor:pointer;font-family:inherit}
.card .gcal{background:var(--accent);color:#fff;border:1px solid var(--accent)}
.card .gcal:hover{opacity:.85}
.card .copy{background:transparent;color:var(--accent);border:1px solid var(--accent)}
.card .copy:hover{background:var(--accent);color:#fff}
.card .copy.done{background:var(--accent);color:#fff}
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
"""


def write_index(out: Path, specs: list[CalSpec], season_no: int, base_url: str,
                dtstamp: str) -> None:
    """購読用リンク一覧ページ（index.html）を書き出す。"""
    import html as _h

    def link(fn: str) -> str:
        return f"{base_url.rstrip('/')}/{fn}" if base_url else fn

    def card(spec: CalSpec) -> str:
        url = link(spec.filename)
        # cid= による「URLで追加（購読）」ディープリンク。ダウンロードは発生しない。
        gcal = "https://calendar.google.com/calendar/render?cid=" + urllib.parse.quote(url, safe="")
        return (
            f'<div class="card"><span class="name">{_h.escape(spec.name)}</span>'
            f'<span class="count">{len(spec.events)} 件</span>'
            f'<a class="gcal" href="{_h.escape(gcal)}" target="_blank" rel="noopener">Googleに追加</a>'
            f'<button class="copy" type="button" data-url="{_h.escape(url)}">URLをコピー</button>'
            f"</div>"
        )

    overview = [s for s in specs if s.category in ("overview", "events")]
    divisions = [s for s in specs if s.category == "division"]
    teams = [s for s in specs if s.category == "team"]

    sections = ["".join(card(s) for s in overview)]

    if divisions:
        sections.append('<h2 class="sec">ディビジョン別</h2>')
        sections.append("".join(card(s) for s in divisions))

    if teams:
        sections.append('<h2 class="sec">チーム別</h2>'
                        '<p class="sub">所属ディビジョンごとに畳んでいます。開いて選んでください。</p>')
        by_div: dict[str, list[CalSpec]] = {}
        for s in teams:
            by_div.setdefault(s.division or "その他", []).append(s)
        for div in sorted(by_div):
            inner = "".join(card(s) for s in sorted(by_div[div], key=lambda s: s.name))
            sections.append(
                f'<details><summary>{_h.escape(div)}'
                f'<span class="count"> {len(by_div[div])} チーム</span></summary>{inner}</details>'
            )

    stamp = f"{dtstamp[:4]}-{dtstamp[4:6]}-{dtstamp[6:8]} {dtstamp[9:11]}:{dtstamp[11:13]} UTC"
    body = f"""<!doctype html><html lang="ja"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>MHL カレンダー購読（{season_no}期）</title><style>{INDEX_CSS}</style></head>
<body><div class="wrap">
<h1>MHL スケジュール カレンダー</h1>
<p class="sub">Misconduct Hockey League {season_no}期 / 最終更新 {stamp}<br>
購読したい単位（全試合・ディビジョン・チーム・イベント）を Google カレンダーに<b>URLで追加</b>してください。自動で最新に追従します。</p>
{''.join(sections)}
<div class="how"><h2 style="font-size:1rem;color:var(--fg)">登録方法（どちらもURL購読＝自動更新）</h2>
<p><b>かんたん:</b>「<b>Googleに追加</b>」を押すと Google カレンダーが開き、そのまま追加できます。</p>
<p><b>手動:</b>「URLをコピー」を押す → Google カレンダー左側「他のカレンダー ＋」→「<b>URLで追加</b>」→ 貼り付けて追加。</p>
<p class="note">※「ダウンロード → インポート」はしないでください。インポートはその時点のコピーで<b>更新されません</b>。
かならず<b>URLで追加（購読）</b>で登録してください。<br>
サイト更新への反映は Google 側の都合で数時間〜1日ほど遅れます。延期の試合はタイトル先頭に <b>⚠※延期</b> が付きます。</p></div>
<p class="foot">データ元: <a href="{_h.escape(BASE)}">misconduct.co.jp</a>
（非公式・個人利用向けの変換ツールです）</p>
</div>
<script>
document.querySelectorAll('button.copy').forEach(function(b){{
  b.addEventListener('click',function(){{
    var url=b.dataset.url, label=b.textContent;
    function done(){{ b.textContent='✓ コピーしました'; b.classList.add('done');
      setTimeout(function(){{ b.textContent=label; b.classList.remove('done'); }},1500); }}
    if(navigator.clipboard&&navigator.clipboard.writeText){{
      navigator.clipboard.writeText(url).then(done,function(){{ window.prompt('このURLをコピーしてください',url); }});
    }} else {{ window.prompt('このURLをコピーしてください',url); }}
  }});
}});
</script>
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

    specs: list[CalSpec] = [CalSpec("all.ics", "MHL 全試合", matches, "overview")]

    for div in sorted({e.division for e in matches if e.division}):
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
    return specs


# ---------------------------------------------------------------------------
# CLI
# ---------------------------------------------------------------------------
def main() -> int:
    ap = argparse.ArgumentParser(description="MHL スケジュール → iCalendar 変換")
    ap.add_argument("-o", "--out", default="calendars", help="出力ディレクトリ")
    ap.add_argument("--location", default="MHL TOKYO", help="LOCATION に入れる会場名")
    ap.add_argument("--base", default=BASE, help="サイトのベース URL")
    ap.add_argument("--base-url", default="",
                    help="公開先の URL（例: https://user.github.io/repo）。index.html の購読リンクに使う")
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
    if not args.no_index:
        write_index(out, specs, season_no, args.base_url, dtstamp)
        print("  index.html", file=sys.stderr)
    print(f"wrote {len(specs)} calendars to {out}/", file=sys.stderr)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
