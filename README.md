# MHL カレンダー連携

[Misconduct Hockey League（misconduct.co.jp）](https://misconduct.co.jp/) の
インラインホッケーのスケジュールを取得し、**iCalendar（.ics）** に変換して
Google カレンダーで購読できるようにするツール。非公式・個人利用向け。

- ディビジョン別に別々の .ics を出力するので、**各自が見たいディビジョンだけ購読**できる
- サイトのファイル名・更新タイミングが不定でも動くよう、**リンクを自動発見**し、
  試合の分類は**データ内の「Division」列**で行う（ファイル名に依存しない）
- 依存パッケージなし（Python 3.9+ 標準ライブラリのみ）

## 出力されるカレンダー

| ファイル | 内容 |
|---|---|
| `all.ics` | 全ディビジョンの全試合 |
| `platinum.ics` / `gold.ics` / `silver.ics` / `bronze.ics` / `brass.ics` / `copper.ics` / `iron.ics` / `womengold.ics` / `35andover.ics` | 各ディビジョンの試合 |
| `team-*.ics` | チーム別（home/away 両方の試合）。日本語名はハッシュ、英字名はチーム名スラッグ |
| `events.ics` | Pick Up Hockey・Drop in Hockey・各種クリニック・エキシビション等の公式プログラム |
| `index.html` | 購読URLの一覧ページ（全試合／ディビジョン別／チーム別／イベント） |

チーム別はチーム数が多いので、`index.html` では所属ディビジョンごとに折りたたんで表示する。

- 開始/終了時刻は日本時間（Asia/Tokyo）で登録される
- 延期の試合はタイトル先頭に `⚠※延期` が付く
- UID は「期＋試合番号」で安定しているので、再取得しても重複せず更新される

## ローカルで試す

```sh
python3 mhl_calendar.py            # ./calendars/ に生成
python3 mhl_calendar.py -o out --location "MHL TOKYO"
```

`.ics` を Google カレンダーの「他のカレンダー → インポート」で読み込めば一回きりの取り込み。
自動追従したい場合は下記の購読方式を使う。

## 自動追従（購読）— GitHub Pages で公開する

`.github/workflows/build.yml` が **6時間ごと**に再生成して GitHub Pages に公開する。
Google カレンダーは公開された URL を定期的に読みに来るので、サイト更新に自動追従する
（※ Google 側の反映は数時間〜1日ほど遅れることがある）。

### セットアップ手順

1. この一式を GitHub リポジトリに push する
   ```sh
   git init && git add -A && git commit -m "MHL calendar"
   gh repo create mhl-calendar --public --source=. --push
   ```
2. リポジトリの **Settings → Pages → Build and deployment → Source** を
   **GitHub Actions** にする
3. Actions タブでワークフローが走り、`https://<ユーザー名>.github.io/mhl-calendar/`
   に一覧ページが公開される
4. その一覧ページで購読したいディビジョンの URL をコピーし、Google カレンダー
   左側「**他のカレンダー ＋ → URLで追加**」に貼り付ける

例）Gold だけ購読したい場合の URL:
`https://<ユーザー名>.github.io/mhl-calendar/gold.ics`

## 仕組み

1. トップページから `/NNth-schedule/` リンクを探し、番号が最大の期を選ぶ
2. その期のページから `wp-content/uploads/*.htm`（cp932）を全部発見して取得
3. Excel 由来の HTML を colspan を考慮して10列グリッドに復元し、
   `試合番号 / 開始 / 〜 / 終了 / Away / vs / Home / Division` を読み取る
4. 月別ファイルと部門別ファイルの重複を「期＋試合番号」で排除
5. ディビジョン別・イベント別・全体の .ics と index.html を出力

## 取り込む範囲

対象は **東京（CXC）の現行シーズン**のスケジュールページ。日付つきの公式イベント
（Drop in Hockey、各種クリニック、Pick Up Hockey、エキシビション、親子スケート等）は
スケジュールファイル内に載っているため `events.ics` に取り込まれる。
`EVENT / CLINIC` ページは各プログラムの説明のみ、トップのニュース欄も
スケジュール内イベントの告知なので、追加で拾う日程データは無い。

以下は別セクションのため現状は対象外（必要なら拡張可能）:
- **HOKKAIDO** 地区のリーグ戦
- **ハピホケ！** セクション独自のスケジュール（東京会場開催分は `events.ics` に入る）

## 注意

- 非公式ツール。サイトの HTML 構造が大きく変わると調整が必要になる場合がある
- 会場名は既定で `MHL TOKYO`。`--location` で変更可
