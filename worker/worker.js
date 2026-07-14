// MHL カレンダー動的フィード（Cloudflare Worker）
//
// GitHub Pages に置かれた feed.json（全試合＋各 VEVENT を事前生成したもの）を読み、
// クエリで絞り込んで iCalendar(.ics) を返す。元サイトには一切アクセスしないので軽い。
//
//   GET /calendar.ics?divs=Gold,Brass&teams=SONIDO,%E3%83%80%E3%82%A4%E3%83%8A%E3%83%A2&events=1&hide=1
//
//   divs   … カンマ区切りのディビジョン名（例 Gold,Brass）
//   teams  … カンマ区切りのチーム名（URLエンコード）。away/home どちらでも一致
//   events … 1 でイベント/クリニックを含める
//   hide   … 1 で延期の試合を除外
//   （何も指定しなければ全試合＋イベント）
//
// フィルタの意味はページ側 index.html の inSubscription() と一致させている。

const DEFAULT_FEED = "https://onnga-wasabi.github.io/mhl-calendar/feed.json";

function parseList(v) {
  if (!v) return [];
  return v.split(",").map((s) => decodeURIComponent(s.trim())).filter(Boolean);
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
    // feed.json をエッジで最大30分キャッシュ（Pages への負荷を抑える）
    const res = await fetch(feedUrl, {
      cf: { cacheTtl: 1800, cacheEverything: true },
    });
    if (!res.ok) return new Response("feed unavailable", { status: 502 });
    const feed = await res.json();

    const q = url.searchParams;
    const divs = new Set(parseList(q.get("divs")));
    const teams = new Set(parseList(q.get("teams")));
    const events = q.get("events") === "1" || q.get("events") === "true";
    const hide = q.get("hide") === "1" || q.get("hide") === "true";
    const any = divs.size || teams.size || events;

    const blocks = [];
    for (const it of feed.events) {
      if (hide && it.n) continue;
      let ok;
      if (!any) ok = true;
      else if (it.k === "p") ok = events;
      else ok = divs.has(it.dv) || teams.has(it.a) || teams.has(it.h);
      if (ok) blocks.push(it.ev);
    }

    const header = feed.header.slice();
    header.push("X-WR-CALNAME:" + (feed.calname || "MHL"));
    const body =
      header.join("\r\n") +
      "\r\n" +
      (blocks.length ? blocks.join("\r\n") + "\r\n" : "") +
      "END:VCALENDAR\r\n";

    return new Response(body, {
      headers: {
        "content-type": "text/calendar; charset=utf-8",
        "content-disposition": 'inline; filename="mhl.ics"',
        "cache-control": "public, max-age=1800",
        "access-control-allow-origin": "*",
      },
    });
  },
};
