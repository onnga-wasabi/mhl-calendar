// MHL カレンダー動的フィード（Cloudflare Worker）
//
// GitHub Pages に置かれた feed.json（全試合＋各 VEVENT を事前生成したもの）を読み、
// クエリで絞り込んで iCalendar(.ics) を返す。元サイトには一切アクセスしないので軽い。
//
//   GET /calendar.ics?divs=Gold,Brass&teams=SONIDO,%E3%83%80%E3%82%A4%E3%83%8A%E3%83%A2&events=1&hide=1
//
//   divs   … カンマ区切りのディビジョン名（例 Gold,Brass）
//   teams  … カンマ区切りのチーム名（URLエンコード）。away/home どちらでも一致
//   events … 1 で公式イベントを全種類含める
//   etypes … カンマ区切りのイベント種類（例 クリニック,Drop in Hockey）。events=1 の下位指定
//   rent   … 1 で非公式・個人イベント（レンタル黄色）を全部含める
//   rlabels… カンマ区切りの非公式ラベル（例 WSJ,Next World）。rent=1 の下位指定
//   hide   … 1 で延期の試合を除外
//   （何も指定しなければ全試合＋全イベント＋全非公式）
//
// フィルタの意味はページ側 index.html の inSubscription() と一致させている。

const DEFAULT_FEED = "https://onnga-wasabi.github.io/mhl-calendar/feed.json";

function parseList(v) {
  if (!v) return [];
  return v.split(",").map((s) => decodeURIComponent(s.trim())).filter(Boolean);
}

// 選択内容から見分けやすいカレンダー名を作る（Googleに同名が並ばないように）。
// カテゴリ（チーム/公式/非公式）ごとに要約するので、混在選択でも中身が分かる。
function calName(q, feed) {
  const season = feed.calname || "MHL";
  const teams = parseList(q.get("teams")).concat(parseList(q.get("divs")));
  const et = parseList(q.get("etypes"));
  const rl = parseList(q.get("rlabels"));
  const seg = [];
  if (teams.length) {
    seg.push(teams.length <= 2 ? teams.join("・") : `${teams[0]}他${teams.length - 1}チーム`);
  }
  if (q.get("events") === "1") seg.push("公式イベント");
  else if (et.length) seg.push(et.length <= 2 ? et.join("・") : `公式イベント${et.length}種`);
  if (q.get("rent") === "1") seg.push("非公式イベント");
  else if (rl.length) seg.push(rl.length <= 2 ? rl.join("・") : `非公式${rl.length}件`);
  return seg.length ? `${season} ${seg.join(" + ")}` : `${season} 全部`;
}

function icsEsc(s) {
  return String(s).replace(/([\\;,])/g, "\\$1").replace(/\n/g, "\\n");
}

// RFC5545 の75オクテット折り返し（UTF-8バイト単位）
function foldLine(line) {
  const enc = new TextEncoder();
  if (enc.encode(line).length <= 75) return line;
  const out = [];
  let cur = "";
  for (const ch of line) {
    if (enc.encode(cur + ch).length > 75) {
      out.push(cur);
      cur = " " + ch;
    } else {
      cur += ch;
    }
  }
  out.push(cur);
  return out.join("\r\n");
}

export default {
  async fetch(request, env) {
    const url = new URL(request.url);

    if (!url.pathname.endsWith(".ics")) {
      return new Response(
        "MHL calendar feed.\n" +
          "使い方: /calendar.ics?divs=Gold,Brass&teams=SONIDO&events=1&hide=1\n",
        { headers: { "content-type": "text/plain; charset=utf-8" } }
      );
    }

    const feedUrl = (env && env.FEED_JSON) || DEFAULT_FEED;
    // feed.json をエッジで最大15分キャッシュ（Pages への負荷を抑えつつ追従を速く）
    const res = await fetch(feedUrl, {
      cf: { cacheTtl: 900, cacheEverything: true },
    });
    if (!res.ok) return new Response("feed unavailable", { status: 502 });
    const feed = await res.json();

    const q = url.searchParams;
    const truthy = (k) => q.get(k) === "1" || q.get(k) === "true";
    const divs = new Set(parseList(q.get("divs")));
    const teams = new Set(parseList(q.get("teams")));
    const etypes = new Set(parseList(q.get("etypes")));
    const rlabels = new Set(parseList(q.get("rlabels")));
    const events = truthy("events");
    const rent = truthy("rent");
    const hide = truthy("hide");
    const any = divs.size || teams.size || etypes.size || rlabels.size || events || rent;

    const blocks = [];
    for (const it of feed.events) {
      if (hide && it.n) continue;
      let ok;
      if (!any) ok = true;
      else if (it.k === "p") ok = events || etypes.has(it.et);
      else if (it.k === "r") ok = rent || rlabels.has(it.et);
      else ok = divs.has(it.dv) || teams.has(it.a) || teams.has(it.h);
      if (ok) blocks.push(it.ev);
    }

    const header = feed.header.slice();
    header.push(foldLine("X-WR-CALNAME:" + icsEsc(calName(q, feed))));
    const body =
      header.join("\r\n") +
      "\r\n" +
      (blocks.length ? blocks.join("\r\n") + "\r\n" : "") +
      "END:VCALENDAR\r\n";

    return new Response(body, {
      headers: {
        "content-type": "text/calendar; charset=utf-8",
        "content-disposition": 'inline; filename="mhl.ics"',
        "cache-control": "public, max-age=900",
        "access-control-allow-origin": "*",
      },
    });
  },
};
