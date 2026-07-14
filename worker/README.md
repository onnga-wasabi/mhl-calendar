# MHL 動的フィード（Cloudflare Worker）

ページでチェックした条件（ディビジョン・チーム・イベント・延期）を、そのまま
**1本の購読URL**にするためのサーバレス。GitHub Pages の `feed.json` を読み、
クエリで絞り込んだ `.ics` を返す。元サイト（misconduct.co.jp）には一切アクセスしない。

```
GET /calendar.ics?divs=Gold,Brass&teams=SONIDO&events=1&hide=1
```

| パラメータ | 意味 |
|---|---|
| `divs`   | カンマ区切りのディビジョン名（例 `Gold,Brass`） |
| `teams`  | カンマ区切りのチーム名（URLエンコード）。home/away どちらでも一致 |
| `events` | `1` でイベント/クリニックを含める |
| `hide`   | `1` で延期の試合を除外 |
| （無指定）| 全試合＋イベント |

## デプロイ手順（あなたの Cloudflare アカウントで）

前提: Node.js が入っていること。`npx` を使うので追加インストールは不要。

```sh
cd worker
npx wrangler login          # ブラウザが開くので Cloudflare にログイン
npx wrangler deploy         # デプロイ。URL が表示される
```

`wrangler deploy` の出力に

```
Published mhl-calendar (x.xx sec)
  https://mhl-calendar.<あなたのサブドメイン>.workers.dev
```

のような **Worker の URL** が出る。これがフィードのベースURL。

### 動作確認

```sh
curl "https://mhl-calendar.<サブドメイン>.workers.dev/calendar.ics?divs=Gold" | head
```

`BEGIN:VCALENDAR` … と Gold の試合が返れば成功。

## ページ側にフィードURLを設定する

Worker のURLを、GitHub リポジトリの **Settings → Secrets and variables → Actions →
Variables** に `FEED_URL` という名前で登録する（値は `https://mhl-calendar.<サブドメイン>.workers.dev`）。

次回のビルド（push か6時間ごと、または Actions を手動実行）で `index.html` の
「選択した内容を購読」パネルが有効になり、チェックした条件のURLが出るようになる。

## メモ

- `feed.json` の場所は `wrangler.toml` の `FEED_JSON`。リポジトリ名/ユーザー名を変えたら合わせる。
- feed.json はエッジで最大30分キャッシュ。無料枠（10万リクエスト/日）で十分。
- Worker は静的な feed.json を読むだけなので、元サイトへの負荷は Actions の6時間ごと取得のみ。
