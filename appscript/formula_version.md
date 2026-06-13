# 数式だけ版（Apps Scriptを使いたくない人向けフォールバック）

スクリプトを使わず、**セルに数式を貼るだけ**でも最小限の2段階ツールが作れます。
自動スコアリングやランキング並べ替えはありませんが、株価とニュースの自動取得は同じように動きます。

> いずれの数式もコピーして対象セルに貼るだけ。`A2` 等は実際のセル位置に合わせてください。

## Stage 1（簡易ランキング表）
新しいシートを作り、A列に銘柄シンボル（`TYO:7203` など）を縦に並べ、隣に以下を入れます。

| 列 | 数式（B2にシンボル、例 `TYO:7203` を入れた場合） |
|---|---|
| 現在値 | `=IFERROR(GOOGLEFINANCE($B2),"—")` |
| 前日比% | `=IFERROR(GOOGLEFINANCE($B2,"changepct"),"—")` |
| 出来高 | `=IFERROR(GOOGLEFINANCE($B2,"volume"),"—")` |
| ニュース件数 | （件数カウントは数式だけでは不可。下のニュース一覧で代用） |

並べ替えは、表を選択して **データ → 範囲を並べ替え → 前日比% 降順** で手動実行。

## Stage 2（個別ウォッチ）
任意のセル（例 `B1`）に銘柄シンボル、`B2` に会社名キーワードを入れて：

```
現在値:   =IFERROR(GOOGLEFINANCE($B$1),"—")
前日比%:  =IFERROR(GOOGLEFINANCE($B$1,"changepct"),"—")
チャート: =IFERROR(SPARKLINE(QUERY(GOOGLEFINANCE($B$1,"price",TODAY()-30,TODAY()),"select Col2")),"—")
```

ニュース見出し（Google News RSS を会社名で自動取得・新しい順20件）：

```
=IFERROR(IMPORTFEED("https://news.google.com/rss/search?q="&ENCODEURL($B$2)&"&hl=ja&gl=JP&ceid=JP:ja","items",TRUE,20),"—")
```

X検索リンク（会社名で）：

```
=HYPERLINK("https://x.com/search?q="&ENCODEURL($B$2)&"&f=live","Xでこの銘柄を検索")
```

中立ソースの見出しを足したいとき（例：NHK経済 / BBCビジネス）：

```
=IMPORTFEED("https://www.nhk.or.jp/rss/news/cat5.xml","items",TRUE,10)
=IMPORTFEED("https://feeds.bbci.co.uk/news/business/rss.xml","items",TRUE,10)
```

## 注意
- `IMPORTFEED` は1シートあたりの数や更新頻度に制限があり、たまに `Loading...` のままになります。少し待つかセルを再入力。
- 自動ランキング・自動スコア・1時間ごと更新が欲しい場合は **Apps Script版（Code.gs）** を使ってください。
