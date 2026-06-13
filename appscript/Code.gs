/**
 * 注目上場企業ウォッチ＆個別深掘り 2段階システム
 * ------------------------------------------------------------
 * このスクリプトは Googleスプレッドシート上で動きます（ターミナル不要）。
 * 使い方は appscript/手順書.md を参照してください。
 *
 * 構成シート:
 *   Watchlist            … 監視銘柄リスト（楽天証券で買える銘柄。自由に行追加可）
 *   ①注目ランキング       … 今注目の銘柄をスコア順に表示（Stage 1）
 *   ③発掘候補（少額×成長） … 予算内で買える小型・成長・IPO銘柄を発掘（Stage 1.5）
 *   ②個別ウォッチ         … 1社を選んで情報を集中表示（Stage 2）
 *   Sources              … 取り込むニュースRSSの一覧
 *
 * 無料・合法の範囲で動かすための方針:
 *   - 株価/出来高/騰落率 … GOOGLEFINANCE（Googleの公式データ・約20分遅延）
 *   - ニュース           … Google News RSS（IMPORTFEED / UrlFetchApp）
 *   - 適時開示           … TDnet（Yanoshin 公開API）
 *   - Yahoo!ファイナンス・株探の画面スクレイピングは規約違反のため行いません。
 *   - 増収増益などの財務成長データは無料では自動取得できないため、IPOの新しさ・
 *     開示/ニュース頻度・株価/出来高モメンタム・手入力メモで「成長らしさ」を近似します。
 */

// ===== 設定（必要なら数値だけ調整してください）=====
var CFG = {
  watchlistSheet: 'Watchlist',
  rankingSheet: '①注目ランキング',
  discoverySheet: '③発掘候補（少額×成長）',
  detailSheet: '②個別ウォッチ',
  suggestSheet: '🔎 新銘柄サジェスト',
  maNewsSheet: '📰 M&A・再編ニュース',
  sourcesSheet: 'Sources',
  newsDays: 7,           // ニュース件数を数える対象日数
  rankingTopN: 30,       // ①注目ランキングの表示件数
  // ①注目スコアの重み（合計が1でなくてOK。相対比較に使います）
  weight: { change: 0.35, volume: 0.20, news: 0.30, disclosure: 0.15 },
  // ③発掘の設定（ここを変えるだけで予算などを調整できます）
  discovery: {
    budgetYen: 50000,    // 1銘柄あたりの予算（円）。この金額で買える銘柄だけを発掘
    unitShares: 100,     // 日本株の売買単位（通常100株）
    ipoRecentDays: 365,  // 上場から何日以内を「IPO（新規上場）」とみなすか
    topN: 30             // 発掘候補の表示件数
  },
  // ③発掘スコアの重み
  wDisc: { growth: 0.30, ipo: 0.20, attention: 0.25, smallcap: 0.15, volspike: 0.10 },
  // 🔎 新銘柄サジェストの設定
  suggest: {
    maxRows: 40,                  // 提案を最大何件出すか
    // TDnetでこの語を含む開示を「きっかけあり」として優先表示
    tdnetNotable: /上方修正|予想.{0,6}修正|増額|株式分割|公開買付|TOB|業務提携|資本提携|新規上場|自己株式取得|ストップ高/,
    // Google Newsで探すクエリと、その「きっかけ」ラベル
    newsQueries: [
      { q: 'ストップ高 東証', label: 'ストップ高' },
      { q: '上方修正 東証', label: '上方修正' },
      { q: '新規上場 グロース', label: '新規上場' }
    ]
  },
  // 📰 M&A・再編ニュースの設定（種別と検索クエリ。自由に追加・編集できます）
  maNews: {
    days: 7,        // 何日以内の記事を対象にするか
    maxRows: 120,   // 表示する最大件数
    categories: [
      { label: '買収',    q: '買収 上場' },
      { label: '子会社化', q: '子会社化 OR 完全子会社化 OR 連結子会社' },
      { label: '分社化',  q: '分社化 OR 会社分割 OR スピンオフ' },
      { label: '経営統合', q: '経営統合 OR 合併' },
      { label: 'TOB',     q: 'TOB OR 株式公開買付' },
      { label: '資本提携', q: '資本業務提携 OR 資本提携' }
    ]
  },
  newsHl: 'ja', newsGl: 'JP', newsCeid: 'JP:ja'
};

// Watchlist の列番号（1始まり）。レイアウト変更時はここだけ直せばよい。
var COL = {
  Code: 1, Name: 2, Market: 3, Symbol: 4, NewsKeyword: 5, XSearchURL: 6,
  ListingDate: 7, KabuMini: 8,                       // ← 入力列（任意）
  Price: 9, ChangePct: 10, Volume: 11, AvgVol20: 12, // ← 以下は自動計算
  NewsCount: 13, Disclosure: 14, VolSpike: 15,
  Unit100Yen: 16, Share1Yen: 17, MinBuyYen: 18,
  AttentionScore: 19, DiscoveryScore: 20, GrowthMemo: 21, UpdatedAt: 22
};
var WL_WIDTH = 22;

// スプレッドシートを開いたときにメニューを追加
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('📈 株ウォッチ')
    .addItem('① 今すぐ全体を更新', 'updateAll')
    .addItem('② 個別ウォッチのニュースを更新', 'refreshDetailNews')
    .addSeparator()
    .addItem('🔎 新銘柄サジェストを更新', 'suggestNewStocks')
    .addItem('➕ チェックした新銘柄をWatchlistに取り込む', 'importCheckedSuggestions')
    .addSeparator()
    .addItem('📰 M&A・再編ニュースを更新', 'updateMaNews')
    .addItem('➕ チェックしたM&A銘柄をWatchlistに取り込む', 'importCheckedMaNews')
    .addSeparator()
    .addItem('🛠 初期セットアップ（最初の1回）', 'setup')
    .addItem('⏰ 自動更新(1時間ごと)をON', 'installHourlyTrigger')
    .addToUi();
}

/** 初期セットアップ：シートを作り、見出し・数式・レイアウトを用意する */
function setup() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  setupWatchlist_(ss);
  setupSources_(ss);
  setupRanking_(ss);
  setupDiscovery_(ss);
  setupSuggest_(ss);
  setupMaNews_(ss);
  setupDetail_(ss);
  SpreadsheetApp.getUi().alert(
    'セットアップ完了。\n\n' +
    '1) 「Watchlist」シートに銘柄を貼り付けてください。\n' +
    '   ・大手も見たい → watchlist_seed.csv\n' +
    '   ・少額×成長を発掘 → growth_smallcap_seed.csv\n' +
    '   （両方つなげて貼ってOK）\n' +
    '2) メニュー「📈 株ウォッチ → ① 今すぐ全体を更新」を実行。\n' +
    '3) 「③発掘候補」シートに予算5万円内で買える小型・成長株が出ます。\n' +
    '4) 「⏰ 自動更新(1時間ごと)をON」で自動化できます。');
}

// ---------- Watchlist ----------
function setupWatchlist_(ss) {
  var sh = ss.getSheetByName(CFG.watchlistSheet) || ss.insertSheet(CFG.watchlistSheet);
  var headers = ['Code', 'Name', 'Market', 'Symbol', 'NewsKeyword', 'XSearchURL',
                 'ListingDate', 'KabuMini',
                 'Price', 'ChangePct', 'Volume', 'AvgVol20', 'NewsCount', 'Disclosure',
                 'VolSpike', 'Unit100Yen', 'Share1Yen', 'MinBuyYen',
                 'AttentionScore', 'DiscoveryScore', 'GrowthMemo(1-5手入力)', 'UpdatedAt'];
  sh.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
  sh.setFrozenRows(1);
  sh.getRange(1, 1, 1, headers.length).setBackground('#1f3864').setFontColor('white');
}

// ---------- Sources（取り込みRSS）----------
function setupSources_(ss) {
  var sh = ss.getSheetByName(CFG.sourcesSheet) || ss.insertSheet(CFG.sourcesSheet);
  sh.clear();
  var rows = [
    ['名称', '区分', 'RSS URL'],
    ['NHK 経済', '公共放送', 'https://www.nhk.or.jp/rss/news/cat5.xml'],
    ['BBC Business', '公共放送', 'https://feeds.bbci.co.uk/news/business/rss.xml'],
    ['Investing.com', '市場データ', 'https://www.investing.com/rss/news.rss'],
    ['ロイター日本語(Google News)', '通信社', 'https://news.google.com/rss/search?q=when:2d%20site:jp.reuters.com&hl=ja&gl=JP&ceid=JP:ja'],
    ['株探(Google News)', '株式専門', 'https://news.google.com/rss/search?q=when:2d%20site:kabutan.jp&hl=ja&gl=JP&ceid=JP:ja'],
    ['TDnet 当日開示', '一次情報', 'https://webapi.yanoshin.jp/webapi/tdnet/list/today.rss']
  ];
  sh.getRange(1, 1, rows.length, 3).setValues(rows);
  sh.getRange('A1:C1').setFontWeight('bold').setBackground('#1f3864').setFontColor('white');
  sh.setFrozenRows(1);
}

// ---------- ①注目ランキング ----------
function setupRanking_(ss) {
  var sh = ss.getSheetByName(CFG.rankingSheet) || ss.insertSheet(CFG.rankingSheet);
  sh.clear();
  sh.getRange('A1').setValue('① 今注目されている銘柄（スコア順）').setFontWeight('bold').setFontSize(14);
  sh.getRange('A2').setValue('更新: メニュー「📈 株ウォッチ → 今すぐ全体を更新」／自動更新は1時間ごと');
  var headers = ['順位', 'Code', 'Name', 'Score', 'ChangePct', 'Volume', 'NewsCount', '開示', '個別を見る', 'X検索'];
  sh.getRange(4, 1, 1, headers.length).setValues([headers]).setFontWeight('bold')
    .setBackground('#2e75b6').setFontColor('white');
  sh.setFrozenRows(4);
}

// ---------- ③発掘候補（少額×成長）----------
function setupDiscovery_(ss) {
  var sh = ss.getSheetByName(CFG.discoverySheet) || ss.insertSheet(CFG.discoverySheet);
  sh.clear();
  sh.getRange('A1').setValue('③ 少額で買える 小型・成長・IPO 発掘候補').setFontWeight('bold').setFontSize(14);
  sh.getRange('A2').setValue(
    '予算 ' + CFG.discovery.budgetYen.toLocaleString() + '円以内で買える銘柄だけを、' +
    '小型・IPO新しさ・話題性・出来高急増・成長メモで採点。予算はコードの CFG.discovery.budgetYen で変更可。');
  sh.getRange('A3').setValue(
    '※「高成長(増収増益)」は無料では自動取得できないため近似値です。最終確認は四季報オンライン無料部分・決算短信(TDnet)・各社IRで。');
  var headers = ['順位', 'Code', 'Name', '市場', '発掘スコア', '株価', '1単元(100株)額', '1株額',
                 '予算内の買い方', 'IPO', '出来高急増', 'ニュース', '個別を見る', 'X検索'];
  sh.getRange(5, 1, 1, headers.length).setValues([headers]).setFontWeight('bold')
    .setBackground('#548235').setFontColor('white');
  sh.setFrozenRows(5);
}

// ---------- 🔎 新銘柄サジェスト ----------
function setupSuggest_(ss) {
  var sh = ss.getSheetByName(CFG.suggestSheet) || ss.insertSheet(CFG.suggestSheet);
  sh.clear();
  sh.getRange('A1').setValue('🔎 新銘柄サジェスト（まだWatchlistに無い注目の新顔）').setFontWeight('bold').setFontSize(14);
  sh.getRange('A2').setValue(
    '更新: メニュー「🔎 新銘柄サジェストを更新」。TDnetの当日開示（上方修正など）とニュース（ストップ高/上方修正/新規上場）から、' +
    'まだ監視していないコードを自動で拾います。');
  sh.getRange('A3').setValue(
    '使い方: 取り込みたい行の「取込」列にチェック → メニュー「➕ チェックした新銘柄をWatchlistに取り込む」。');
  var headers = ['取込', 'Code', 'Name', 'きっかけ', 'ソース', '詳細リンク'];
  sh.getRange(5, 1, 1, headers.length).setValues([headers]).setFontWeight('bold')
    .setBackground('#7030a0').setFontColor('white');
  sh.setFrozenRows(5);
  sh.setColumnWidth(3, 220);
  sh.setColumnWidth(4, 160);
  sh.setColumnWidth(6, 320);
}

// ---------- 📰 M&A・再編ニュース ----------
function setupMaNews_(ss) {
  var sh = ss.getSheetByName(CFG.maNewsSheet) || ss.insertSheet(CFG.maNewsSheet);
  sh.clear();
  sh.getRange('A1').setValue('📰 上場企業の M&A・再編ニュース（買収・子会社化・分社化 ほか）')
    .setFontWeight('bold').setFontSize(14);
  sh.getRange('A2').setValue(
    '更新: メニュー「📰 M&A・再編ニュースを更新」。種別ごとにニュースを検索して新しい順に表示します。' +
    '検索する種別やキーワードは Code.gs の CFG.maNews.categories で編集できます。');
  var headers = ['日付', '種別', '記事（クリックで開く）', 'ソース', '関連コード', 'Watchに追加'];
  sh.getRange(4, 1, 1, headers.length).setValues([headers]).setFontWeight('bold')
    .setBackground('#c55a11').setFontColor('white');
  sh.setFrozenRows(4);
  sh.setColumnWidth(1, 130);
  sh.setColumnWidth(2, 90);
  sh.setColumnWidth(3, 520);
  sh.setColumnWidth(4, 150);
}

// ---------- ②個別ウォッチ ----------
function setupDetail_(ss) {
  var sh = ss.getSheetByName(CFG.detailSheet) || ss.insertSheet(CFG.detailSheet);
  sh.clear();
  sh.getRange('A1').setValue('② 個別企業ウォッチ').setFontWeight('bold').setFontSize(14);
  sh.getRange('A3').setValue('銘柄コードを選択 →').setFontWeight('bold');

  // B3 にウォッチリストの Code から選べるプルダウンを設定
  var rule = SpreadsheetApp.newDataValidation()
    .requireValueInRange(ss.getSheetByName(CFG.watchlistSheet).getRange('A2:A'), true)
    .setAllowInvalid(true).build();
  sh.getRange('B3').setDataValidation(rule);

  var WL = CFG.watchlistSheet;
  // 基本情報（Watchlist から VLOOKUP ＋ GOOGLEFINANCE）。B6 = シンボル。
  var rows = [
    ['会社名',   "=IFERROR(VLOOKUP($B$3,'" + WL + "'!A:H,2,FALSE),\"—\")"],
    ['市場',     "=IFERROR(VLOOKUP($B$3,'" + WL + "'!A:H,3,FALSE),\"—\")"],
    ['シンボル', "=IFERROR(VLOOKUP($B$3,'" + WL + "'!A:H,4,FALSE),\"—\")"],
    ['現在値',   "=IFERROR(GOOGLEFINANCE(B6),\"—\")"],   // B7
    ['前日比%',  "=IFERROR(GOOGLEFINANCE(B6,\"changepct\"),\"—\")"],
    ['出来高',   "=IFERROR(GOOGLEFINANCE(B6,\"volume\"),\"—\")"],
    ['PER',      "=IFERROR(GOOGLEFINANCE(B6,\"pe\"),\"—\")"],
    ['時価総額', "=IFERROR(GOOGLEFINANCE(B6,\"marketcap\"),\"—\")"],
    ['1単元(100株)購入額', "=IFERROR(GOOGLEFINANCE(B6)*" + CFG.discovery.unitShares + ",\"—\")"],
    ['かぶミニ(1株)可', "=IFERROR(VLOOKUP($B$3,'" + WL + "'!A:H,8,FALSE),\"—\")"],
    ['上場日',   "=IFERROR(VLOOKUP($B$3,'" + WL + "'!A:H,7,FALSE),\"—\")"],
    ['30日チャート', "=IFERROR(SPARKLINE(QUERY(GOOGLEFINANCE(B6,\"price\",TODAY()-30,TODAY()),\"select Col2\"),{\"charttype\",\"line\"}),\"—\")"]
  ];
  sh.getRange(4, 1, rows.length, 2).setValues(rows);
  var lastInfoRow = 3 + rows.length; // = 15
  sh.getRange(4, 1, rows.length, 1).setFontWeight('bold');

  // Xへの入口・開示リンク（基本情報の下に配置）
  var r = lastInfoRow + 2; // 17
  sh.getRange('A' + r).setValue('X検索（クリックで開く）').setFontWeight('bold');
  sh.getRange('B' + r).setFormula(
    "=IFERROR(HYPERLINK(VLOOKUP($B$3,'" + WL + "'!A:H,6,FALSE),\"Xでこの銘柄を検索\"),\"—\")");
  sh.getRange('A' + (r + 1)).setValue('適時開示(TDnet)').setFontWeight('bold');
  sh.getRange('B' + (r + 1)).setFormula(
    "=HYPERLINK(\"https://www.release.tdnet.info/inbs/I_main_00.html\",\"TDnet 適時開示閲覧\")");
  sh.getRange('A' + (r + 2)).setValue('EDINET(法定開示)').setFontWeight('bold');
  sh.getRange('B' + (r + 2)).setFormula(
    "=HYPERLINK(\"https://disclosure.edinet-fsa.go.jp/\",\"EDINETで検索\")");

  // ニュース見出し（Google News RSS を会社名で検索して自動表示）
  var nh = r + 4;
  sh.getRange('A' + nh).setValue('▼ この企業のニュース（自動取得・新しい順）').setFontWeight('bold').setFontSize(12);
  var kw = "VLOOKUP($B$3,'" + WL + "'!A:H,5,FALSE)";
  var newsUrl = "\"https://news.google.com/rss/search?q=\"&ENCODEURL(" + kw +
                ")&\"&hl=" + CFG.newsHl + "&gl=" + CFG.newsGl + "&ceid=" + CFG.newsCeid + "\"";
  sh.getRange('A' + (nh + 1)).setFormula(
    "=IFERROR(IMPORTFEED(" + newsUrl + ",\"items\",TRUE,20),\"（銘柄を選ぶとニュースが表示されます）\")");

  sh.getRange('B3').setValue('7203'); // 初期表示
  sh.setColumnWidth(1, 170);
  sh.setColumnWidth(2, 520);
}

/**
 * 全体更新：株価系の数式を入れ、ニュース件数・開示を取得し、各スコアを計算して
 * ①注目ランキングと③発掘候補を並べ替える。メニューまたは1時間トリガーから呼ばれる。
 */
function updateAll() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(CFG.watchlistSheet);
  var last = sh.getLastRow();
  if (last < 2) { SpreadsheetApp.getUi().alert('Watchlist に銘柄を貼り付けてください。'); return; }
  var n = last - 1;

  // 株価系の数式（Price/ChangePct/Volume/AvgVol20）を毎回入れ直す（行追加に追従）
  var priceF = [], chgF = [], volF = [], avgF = [];
  for (var i = 0; i < n; i++) {
    var rr = i + 2;
    priceF.push(['=IFERROR(GOOGLEFINANCE(D' + rr + '),"")']);
    chgF.push(['=IFERROR(GOOGLEFINANCE(D' + rr + ',"changepct"),"")']);
    volF.push(['=IFERROR(GOOGLEFINANCE(D' + rr + ',"volume"),"")']);
    avgF.push(['=IFERROR(AVERAGE(INDEX(GOOGLEFINANCE(D' + rr + ',"volume",TODAY()-30,TODAY()),0,2)),"")']);
  }
  sh.getRange(2, COL.Price, n, 1).setFormulas(priceF);
  sh.getRange(2, COL.ChangePct, n, 1).setFormulas(chgF);
  sh.getRange(2, COL.Volume, n, 1).setFormulas(volF);
  sh.getRange(2, COL.AvgVol20, n, 1).setFormulas(avgF);
  SpreadsheetApp.flush();

  // 全体を読み込み（A..V）
  var data = sh.getRange(2, 1, n, WL_WIDTH).getValues();
  var todaysDisclosures = fetchTodayDisclosures_();
  var budget = CFG.discovery.budgetYen;
  var today = new Date();

  // 1行ずつ外部データ取得＋派生値の計算
  var newsCounts = [], discFlags = [], volSpikes = [], unit100 = [], share1 = [],
      minBuy = [], smallcap = [], ipoScore = [], momentum = [], memoNorm = [];
  for (var j = 0; j < n; j++) {
    var row = data[j];
    var keyword = row[COL.NewsKeyword - 1] || row[COL.Name - 1];
    newsCounts.push(keyword ? fetchNewsCount_(keyword) : 0);
    discFlags.push(hasDisclosure_((row[COL.Name - 1] || '').toString(), todaysDisclosures) ? 1 : 0);

    var price = parseFloat(row[COL.Price - 1]) || 0;
    var vol = parseFloat(row[COL.Volume - 1]) || 0;
    var avg = parseFloat(row[COL.AvgVol20 - 1]) || 0;
    var chg = parseFloat(row[COL.ChangePct - 1]) || 0;
    var u = price * CFG.discovery.unitShares;
    var s1 = price;
    var kabu = isKabuMini_(row[COL.KabuMini - 1]);

    volSpikes.push(avg > 0 ? Math.round((vol / avg) * 100) / 100 : '');
    unit100.push(u || '');
    share1.push(s1 || '');
    minBuy.push(minBuyYen_(u, s1, kabu, budget));

    var market = (row[COL.Market - 1] || '').toString();
    smallcap.push(/グロース|growth/i.test(market) ? 1 : 0);
    ipoScore.push(ipoScore_(row[COL.ListingDate - 1], today));
    momentum.push(chg > 0 ? Math.min(chg / 10, 1) : 0);
    var memo = parseFloat(row[COL.GrowthMemo - 1]);
    memoNorm.push(isNaN(memo) ? 0 : Math.max(0, Math.min(memo / 5, 1)));

    Utilities.sleep(150); // RSSサーバーへの配慮
  }

  // ① 注目スコア（相対ランクで正規化）
  var chgAbs = data.map(function (r) { return Math.abs(parseFloat(r[COL.ChangePct - 1]) || 0); });
  var volArr = data.map(function (r) { return parseFloat(r[COL.Volume - 1]) || 0; });
  var attention = [];
  for (var a = 0; a < n; a++) {
    var sc = CFG.weight.change * pct_(chgAbs, chgAbs[a])
           + CFG.weight.volume * pct_(volArr, volArr[a])
           + CFG.weight.news   * pct_(newsCounts, newsCounts[a])
           + CFG.weight.disclosure * discFlags[a];
    attention.push(Math.round(sc * 1000) / 10);
  }

  // ③ 発掘スコア
  var volSpikeNum = volSpikes.map(function (v) { return parseFloat(v) || 0; });
  var discovery = [];
  for (var d = 0; d < n; d++) {
    var growthProxy = 0.5 * memoNorm[d] + 0.3 * discFlags[d] + 0.2 * momentum[d];
    var attn = 0.6 * pct_(newsCounts, newsCounts[d]) + 0.4 * pct_(volSpikeNum, volSpikeNum[d]);
    var volNorm = Math.min(volSpikeNum[d] / 3, 1); // 3倍で満点
    var score = CFG.wDisc.growth * growthProxy
              + CFG.wDisc.ipo * ipoScore[d]
              + CFG.wDisc.attention * attn
              + CFG.wDisc.smallcap * smallcap[d]
              + CFG.wDisc.volspike * volNorm;
    // 予算で買えない銘柄は発掘候補から除外（スコア0）
    if (!(parseFloat(minBuy[d]) <= budget)) score = 0;
    discovery.push(Math.round(score * 1000) / 10);
  }

  // 書き戻し（GrowthMemo=U は手入力なので触らない）
  writeCol_(sh, COL.NewsCount, newsCounts, n);
  writeCol_(sh, COL.Disclosure, discFlags.map(function (v) { return v ? '◯' : ''; }), n);
  writeCol_(sh, COL.VolSpike, volSpikes, n);
  writeCol_(sh, COL.Unit100Yen, unit100, n);
  writeCol_(sh, COL.Share1Yen, share1, n);
  writeCol_(sh, COL.MinBuyYen, minBuy, n);
  writeCol_(sh, COL.AttentionScore, attention, n);
  writeCol_(sh, COL.DiscoveryScore, discovery, n);
  var stamp = []; for (var t = 0; t < n; t++) stamp.push(new Date());
  sh.getRange(2, COL.UpdatedAt, n, 1).setValues(stamp.map(function (v) { return [v]; }))
    .setNumberFormat('yyyy/MM/dd HH:mm');

  buildRanking_(ss);
  buildDiscovery_(ss);
}

/** 1列ぶんの値を書き込むヘルパー */
function writeCol_(sh, col, arr, n) {
  sh.getRange(2, col, n, 1).setValues(arr.map(function (v) { return [v]; }));
}

/** 相対パーセンタイル（0-1）。配列内でvalが上位なら1に近い */
function pct_(arr, val) {
  var valid = arr.map(function (x) { return parseFloat(x); }).filter(function (x) { return !isNaN(x); });
  if (!valid.length) return 0;
  var below = valid.filter(function (x) { return x < val; }).length;
  return below / valid.length;
}

/** かぶミニ可フラグの判定（○ / ◯ / O / yes / 1 / true を可とみなす） */
function isKabuMini_(v) {
  var s = (v == null ? '' : v).toString().trim().toLowerCase();
  return s === '○' || s === '◯' || s === 'o' || s === 'yes' || s === '1' || s === 'true' || s === '可';
}

/** 予算内で買える最小金額。1単元→無理ならかぶミニ1株→どちらも不可なら単元額を返す */
function minBuyYen_(unit100, share1, kabu, budget) {
  if (!unit100 && !share1) return '';
  if (unit100 && unit100 <= budget) return unit100;
  if (kabu && share1 && share1 <= budget) return share1;
  return unit100 || share1; // 予算外（>budget）
}

/** 上場日から IPO新しさスコア(0-1)。新しいほど1に近い。範囲外/不明は0 */
function ipoScore_(listingDate, today) {
  if (!listingDate) return 0;
  var dt = (listingDate instanceof Date) ? listingDate : new Date(listingDate);
  if (isNaN(dt.getTime())) return 0;
  var days = (today - dt) / (1000 * 60 * 60 * 24);
  if (days < 0 || days > CFG.discovery.ipoRecentDays) return 0;
  return 1 - (days / CFG.discovery.ipoRecentDays);
}

/** ①注目ランキングを並べ替えて出力 */
function buildRanking_(ss) {
  var wl = ss.getSheetByName(CFG.watchlistSheet);
  var last = wl.getLastRow();
  if (last < 2) return;
  var n = last - 1;
  var d = wl.getRange(2, 1, n, WL_WIDTH).getValues();
  var rows = d.map(function (r) {
    return { code: r[COL.Code - 1], name: r[COL.Name - 1], xurl: r[COL.XSearchURL - 1],
             chg: r[COL.ChangePct - 1], vol: r[COL.Volume - 1], news: r[COL.NewsCount - 1],
             disc: r[COL.Disclosure - 1], score: parseFloat(r[COL.AttentionScore - 1]) || 0 };
  });
  rows.sort(function (a, b) { return b.score - a.score; });
  rows = rows.slice(0, CFG.rankingTopN);

  var rk = ss.getSheetByName(CFG.rankingSheet);
  rk.getRange(5, 1, Math.max(rk.getLastRow() - 4, 1), 10).clearContent();
  var detailGid = ss.getSheetByName(CFG.detailSheet).getSheetId();
  var out = rows.map(function (r, i) {
    return [i + 1, r.code, r.name, r.score, r.chg, r.vol, r.news, r.disc ? '◯' : '', r.code, r.xurl];
  });
  if (!out.length) return;
  rk.getRange(5, 1, out.length, 10).setValues(out);
  for (var i = 0; i < out.length; i++) {
    var row = 5 + i;
    rk.getRange(row, 9).setFormula('=HYPERLINK("#gid=' + detailGid + '","▶ ' + out[i][1] + ' を見る")');
    if (out[i][9]) rk.getRange(row, 10).setFormula('=HYPERLINK("' + out[i][9] + '","X検索")');
  }
  rk.getRange(5, 4, out.length, 1).setNumberFormat('0.0');
}

/** ③発掘候補：予算内で買える小型・成長・IPO銘柄を並べ替えて出力 */
function buildDiscovery_(ss) {
  var wl = ss.getSheetByName(CFG.watchlistSheet);
  var last = wl.getLastRow();
  if (last < 2) return;
  var n = last - 1;
  var d = wl.getRange(2, 1, n, WL_WIDTH).getValues();
  var budget = CFG.discovery.budgetYen;
  var today = new Date();

  var rows = [];
  for (var i = 0; i < n; i++) {
    var r = d[i];
    var minB = parseFloat(r[COL.MinBuyYen - 1]);
    var score = parseFloat(r[COL.DiscoveryScore - 1]) || 0;
    var market = (r[COL.Market - 1] || '').toString();
    var ipo = ipoScore_(r[COL.ListingDate - 1], today) > 0;
    var isSmall = /グロース|growth/i.test(market);
    // 予算内 かつ（小型/グロース or IPO新しい）の銘柄だけを対象に
    if (!(minB <= budget) || !(isSmall || ipo)) continue;

    var price = parseFloat(r[COL.Price - 1]) || 0;
    var unit = parseFloat(r[COL.Unit100Yen - 1]) || 0;
    var s1 = parseFloat(r[COL.Share1Yen - 1]) || 0;
    var kabu = isKabuMini_(r[COL.KabuMini - 1]);
    var how = (unit && unit <= budget) ? '1単元(100株)'
            : (kabu && s1 && s1 <= budget) ? '1株(かぶミニ)' : '—';
    var vs = parseFloat(r[COL.VolSpike - 1]);
    rows.push({
      code: r[COL.Code - 1], name: r[COL.Name - 1], market: market, score: score,
      price: price, unit: unit, share1: s1, how: how,
      ipo: ipo ? '◯' : '', volspike: isNaN(vs) ? '' : (vs + '倍'),
      news: r[COL.NewsCount - 1], xurl: r[COL.XSearchURL - 1]
    });
  }
  rows.sort(function (a, b) { return b.score - a.score; });
  rows = rows.slice(0, CFG.discovery.topN);

  var sh = ss.getSheetByName(CFG.discoverySheet);
  sh.getRange(6, 1, Math.max(sh.getLastRow() - 5, 1), 14).clearContent();
  if (!rows.length) {
    sh.getRange(6, 1).setValue('予算内で条件に合う銘柄がありません。Watchlistにグロース/小型/IPO銘柄を追加するか、CFG.discovery.budgetYen を上げてください。');
    return;
  }
  var detailGid = ss.getSheetByName(CFG.detailSheet).getSheetId();
  var out = rows.map(function (r, i) {
    return [i + 1, r.code, r.name, r.market, r.score, r.price, r.unit, r.share1,
            r.how, r.ipo, r.volspike, r.news, r.code, r.xurl];
  });
  sh.getRange(6, 1, out.length, 14).setValues(out);
  for (var k = 0; k < out.length; k++) {
    var row = 6 + k;
    sh.getRange(row, 13).setFormula('=HYPERLINK("#gid=' + detailGid + '","▶ ' + out[k][1] + ' を見る")');
    if (out[k][13]) sh.getRange(row, 14).setFormula('=HYPERLINK("' + out[k][13] + '","X検索")');
  }
  sh.getRange(6, 5, out.length, 1).setNumberFormat('0.0');
  sh.getRange(6, 6, out.length, 3).setNumberFormat('#,##0');
}

/**
 * 🔎 新銘柄サジェスト：TDnet当日開示＋ニュースから、まだWatchlistに無いコードを拾って提案。
 * メニューから手動で実行（外部取得が多いので自動トリガーには入れていない）。
 */
function suggestNewStocks() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(CFG.suggestSheet);
  if (!sh) { setupSuggest_(ss); sh = ss.getSheetByName(CFG.suggestSheet); }

  // 既存のWatchlistコード（除外用）
  var wl = ss.getSheetByName(CFG.watchlistSheet);
  var existing = {};
  if (wl && wl.getLastRow() > 1) {
    wl.getRange(2, COL.Code, wl.getLastRow() - 1, 1).getValues().forEach(function (r) {
      var c = code4_(r[0]); if (c) existing[c] = true;
    });
  }

  var found = {}; // code4 -> { name, reasons:{}, sources:{}, link }
  function add(code, name, reason, source, link) {
    var c = code4_(code);
    if (!c || existing[c]) return;
    if (!found[c]) found[c] = { name: name || '', reasons: {}, sources: {}, link: link || '' };
    if (name && !found[c].name) found[c].name = name;
    if (reason) found[c].reasons[reason] = true;
    if (source) found[c].sources[source] = true;
    if (link && !found[c].link) found[c].link = link;
  }

  // 1) TDnet 当日開示（コード・社名つき）
  var disc = fetchTodayDisclosuresFull_();
  disc.forEach(function (d) {
    var notable = CFG.suggest.tdnetNotable.test(d.title || '');
    var reason = notable ? ('開示: ' + shortTitle_(d.title)) : '適時開示';
    // 高シグナルの開示を優先。ノイズを減らすため、notableのみ採用。
    if (notable) add(d.code, d.name, reason, 'TDnet', d.url);
  });

  // 2) ニュース（コードは <1234> 形式のみ抽出して誤検出を防ぐ）
  CFG.suggest.newsQueries.forEach(function (nq) {
    var items = fetchNewsItems_(nq.q, 30);
    items.forEach(function (it) {
      var text = (it.title || '') + ' ' + (it.desc || '');
      extractCodeNamePairs_(text).forEach(function (p) {
        add(p.code, p.name, nq.label, 'ニュース', it.link);
      });
    });
    Utilities.sleep(200);
  });

  // 出力（TDnet開示きっかけを上に）
  var rows = Object.keys(found).map(function (c) {
    var f = found[c];
    return {
      code: c, name: f.name,
      reason: Object.keys(f.reasons).join(' / '),
      source: Object.keys(f.sources).join(' / '),
      link: f.link,
      pri: f.sources['TDnet'] ? 0 : 1
    };
  });
  rows.sort(function (a, b) { return a.pri - b.pri; });
  rows = rows.slice(0, CFG.suggest.maxRows);

  sh.getRange(6, 1, Math.max(sh.getLastRow() - 5, 1), 6).clearContent();
  sh.getRange(6, 1, Math.max(sh.getLastRow() - 5, 1), 1).removeCheckboxes();
  if (!rows.length) {
    sh.getRange(6, 2).setValue('新しい候補は見つかりませんでした（時間をおいて再実行してください）。');
    return;
  }
  var out = rows.map(function (r) { return [false, r.code, r.name, r.reason, r.source, r.link]; });
  sh.getRange(6, 1, out.length, 6).setValues(out);
  sh.getRange(6, 1, out.length, 1).insertCheckboxes();
  // 詳細リンクをハイパーリンク化
  for (var i = 0; i < out.length; i++) {
    if (out[i][5]) sh.getRange(6 + i, 6).setFormula('=HYPERLINK("' + out[i][5] + '","開く")');
  }
  SpreadsheetApp.getActiveSpreadsheet().toast(out.length + '件の新銘柄候補を表示しました。', '🔎 サジェスト', 5);
}

/** チェックされた提案行をWatchlistに取り込む */
function importCheckedSuggestions() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(CFG.suggestSheet);
  var wl = ss.getSheetByName(CFG.watchlistSheet);
  if (!sh || !wl) return;
  var last = sh.getLastRow();
  if (last < 6) return;
  var vals = sh.getRange(6, 1, last - 5, 3).getValues(); // 取込, Code, Name

  var newRows = [], importedRowIdx = [];
  vals.forEach(function (r, i) {
    if (r[0] === true && r[1]) {
      var code = code4_(r[1]);
      var name = r[2] || code;
      newRows.push([code, name, '', 'TYO:' + code, name,
                    'https://x.com/search?q=%24' + code + '&f=live', '', '']); // A..H
      importedRowIdx.push(i);
    }
  });
  if (!newRows.length) {
    ss.toast('チェックされた行がありません。', '➕ 取り込み', 4);
    return;
  }
  var start = wl.getLastRow() + 1;
  wl.getRange(start, 1, newRows.length, 8).setValues(newRows);
  // 取り込んだ行のチェックを外す
  importedRowIdx.forEach(function (i) { sh.getRange(6 + i, 1).setValue(false); });
  ss.toast(newRows.length + '件をWatchlistに追加しました。市場(C列)は空欄なので、グロース株なら「東証グロース」を入れてください。次に「① 今すぐ全体を更新」を実行。', '➕ 取り込み完了', 8);
}

/**
 * 📰 上場企業の M&A・再編ニュース（買収・子会社化・分社化 ほか）を抽出して一覧化。
 * 種別ごとに Google News RSS を検索し、新しい順に表示。メニューから手動更新。
 */
function updateMaNews() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(CFG.maNewsSheet);
  if (!sh) { setupMaNews_(ss); sh = ss.getSheetByName(CFG.maNewsSheet); }
  var tz = ss.getSpreadsheetTimeZone() || 'Asia/Tokyo';

  var map = {}; // link -> 記事
  CFG.maNews.categories.forEach(function (cat) {
    var q = cat.q + ' when:' + CFG.maNews.days + 'd';
    fetchNewsItems_(q, 40).forEach(function (it) {
      var key = it.link || it.title;
      if (!key) return;
      if (!map[key]) {
        var pairs = extractCodeNamePairs_((it.title || '') + ' ' + (it.desc || ''));
        map[key] = {
          title: cleanNewsTitle_(it.title), link: it.link,
          source: it.source || '', date: parseNewsDate_(it.pubDate),
          labels: {}, code: pairs.length ? pairs[0].code : ''
        };
      }
      map[key].labels[cat.label] = true;
    });
    Utilities.sleep(200);
  });

  var rows = Object.keys(map).map(function (k) {
    var r = map[k];
    return {
      date: r.date, label: Object.keys(r.labels).join(' / '),
      title: r.title, link: r.link, source: r.source, code: r.code
    };
  });
  rows.sort(function (a, b) { return (b.date ? b.date.getTime() : 0) - (a.date ? a.date.getTime() : 0); });
  rows = rows.slice(0, CFG.maNews.maxRows);

  // 既存内容をクリア
  var lastRow = sh.getLastRow();
  if (lastRow >= 5) {
    sh.getRange(5, 1, lastRow - 4, 6).clearContent();
    sh.getRange(5, 6, lastRow - 4, 1).removeCheckboxes();
  }
  if (!rows.length) {
    sh.getRange(5, 1).setValue('該当ニュースが見つかりませんでした（時間をおいて再実行してください）。');
    return;
  }

  // 値・チェックボックスをまとめて書き込み
  var values = rows.map(function (r) {
    return [r.date ? Utilities.formatDate(r.date, tz, 'yyyy/MM/dd HH:mm') : '',
            r.label, '', r.source, r.code, false];
  });
  sh.getRange(5, 1, values.length, 6).setValues(values);
  sh.getRange(5, 6, values.length, 1).insertCheckboxes();
  // 記事タイトルをクリック可能なリンクに（C列）
  for (var i = 0; i < rows.length; i++) {
    var title = rows[i].title || '(無題)';
    var cell = sh.getRange(5 + i, 3);
    if (rows[i].link) {
      cell.setRichTextValue(SpreadsheetApp.newRichTextValue().setText(title).setLinkUrl(rows[i].link).build());
    } else {
      cell.setValue(title);
    }
  }
  ss.toast(rows.length + '件のM&A・再編ニュースを表示しました。', '📰 M&Aニュース', 5);
}

/** 📰シートでチェックした行（関連コードあり）をWatchlistに取り込む */
function importCheckedMaNews() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(CFG.maNewsSheet);
  var wl = ss.getSheetByName(CFG.watchlistSheet);
  if (!sh || !wl || sh.getLastRow() < 5) return;
  var vals = sh.getRange(5, 1, sh.getLastRow() - 4, 6).getValues(); // A..F
  var newRows = [], rowsToUncheck = [];
  vals.forEach(function (r, i) {
    var code = code4_(r[4]); // E 関連コード
    if (r[5] === true && code) {
      newRows.push([code, code, '', 'TYO:' + code, code,
                    'https://x.com/search?q=%24' + code + '&f=live', '', '']);
      rowsToUncheck.push(i);
    }
  });
  if (!newRows.length) {
    ss.toast('チェックされた行（関連コードあり）がありません。', '➕ 取り込み', 4);
    return;
  }
  wl.getRange(wl.getLastRow() + 1, 1, newRows.length, 8).setValues(newRows);
  rowsToUncheck.forEach(function (i) { sh.getRange(5 + i, 6).setValue(false); });
  ss.toast(newRows.length + '件をWatchlistに追加しました。市場(C列)や名称を整えて「① 今すぐ全体を更新」を実行してください。', '➕ 取り込み完了', 8);
}

/** Google News の日付文字列をDateに（失敗時null） */
function parseNewsDate_(s) {
  if (!s) return null;
  var d = new Date(s);
  return isNaN(d.getTime()) ? null : d;
}

/** Google Newsのタイトル末尾の " - 媒体名" を整える（ソース列が別にあるため軽く整形） */
function cleanNewsTitle_(t) {
  return (t || '').toString().replace(/\s+/g, ' ').trim();
}

/** ②個別ウォッチのニュースだけを更新したいとき（数式の再計算を促す） */
function refreshDetailNews() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(CFG.detailSheet);
  var cur = sh.getRange('B3').getValue();
  sh.getRange('B3').setValue('');
  SpreadsheetApp.flush();
  sh.getRange('B3').setValue(cur); // 再設定で IMPORTFEED を再取得
}

/** B3(②の銘柄選択) が変わったらニュースを取り直す */
function onEdit(e) {
  if (!e || !e.range) return;
  var sh = e.range.getSheet();
  if (sh.getName() === CFG.detailSheet && e.range.getA1Notation() === 'B3') {
    // IMPORTFEED は B3 参照なので自動再計算される。ここでは特別な処理は不要。
  }
}

// ===== RSS / API ヘルパー =====

/** Google News RSS を検索し、対象日数内の記事件数を数える */
function fetchNewsCount_(keyword) {
  try {
    var url = 'https://news.google.com/rss/search?q=' + encodeURIComponent(keyword) +
              '%20when:' + CFG.newsDays + 'd&hl=' + CFG.newsHl + '&gl=' + CFG.newsGl +
              '&ceid=' + CFG.newsCeid;
    var res = UrlFetchApp.fetch(url, { muteHttpExceptions: true, followRedirects: true });
    if (res.getResponseCode() !== 200) return 0;
    var root = XmlService.parse(res.getContentText()).getRootElement();
    var channel = root.getChild('channel');
    if (!channel) return 0;
    return channel.getChildren('item').length;
  } catch (err) {
    return 0;
  }
}

/** 証券コードを4桁(または4文字)に正規化。TDnetの5桁(末尾0)は先頭4文字にする */
function code4_(v) {
  if (v == null) return '';
  var s = v.toString().trim().toUpperCase();
  var m = s.match(/[0-9][0-9A-Z]{3}[0-9A-Z]?/); // 4〜5文字の英数コード
  if (!m) return '';
  var c = m[0];
  if (c.length === 5) c = c.substring(0, 4); // 5桁は先頭4文字（TDnet形式）
  return c;
}

/** 開示タイトルを短く（一覧表示用） */
function shortTitle_(t) {
  t = (t || '').toString().replace(/\s+/g, ' ').trim();
  return t.length > 28 ? t.substring(0, 28) + '…' : t;
}

/** TDnet当日開示を コード・社名・タイトル・URL つきで取得 */
function fetchTodayDisclosuresFull_() {
  try {
    var url = 'https://webapi.yanoshin.jp/webapi/tdnet/list/today.json?limit=1000';
    var res = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    if (res.getResponseCode() !== 200) return [];
    var items = (JSON.parse(res.getContentText()).items) || [];
    return items.map(function (it) {
      var t = it.Tdnet || it.tdnet || {};
      return {
        code: (t.company_code || t.companyCode || '').toString(),
        name: (t.company_name || t.companyName || '').toString(),
        title: (t.title || '').toString(),
        url: (t.document_url || t.url || '').toString()
      };
    }).filter(function (d) { return d.code; });
  } catch (err) {
    return [];
  }
}

/** Google News RSS の記事（title/link/desc）を取得 */
function fetchNewsItems_(query, max) {
  try {
    var url = 'https://news.google.com/rss/search?q=' + encodeURIComponent(query) +
              '%20when:3d&hl=' + CFG.newsHl + '&gl=' + CFG.newsGl + '&ceid=' + CFG.newsCeid;
    var res = UrlFetchApp.fetch(url, { muteHttpExceptions: true, followRedirects: true });
    if (res.getResponseCode() !== 200) return [];
    var channel = XmlService.parse(res.getContentText()).getRootElement().getChild('channel');
    if (!channel) return [];
    var items = channel.getChildren('item');
    var out = [];
    for (var i = 0; i < items.length && i < (max || 30); i++) {
      out.push({
        title: items[i].getChildText('title') || '',
        link: items[i].getChildText('link') || '',
        desc: items[i].getChildText('description') || '',
        pubDate: items[i].getChildText('pubDate') || '',
        source: items[i].getChildText('source') || ''
      });
    }
    return out;
  } catch (err) {
    return [];
  }
}

/** テキストから <1234> 形式の証券コードと直前の社名を抽出（誤検出を抑制） */
function extractCodeNamePairs_(text) {
  var pairs = [], seen = {};
  // 全角/半角の山かっこ・かぎかっこ内の4文字コード（例: ＜7203＞ <7203> ［7203］）
  var re = /([^\s　<＜\[［(（]{0,16}?)\s*[<＜\[［]([0-9][0-9A-Za-z]{3})[>＞\]］]/g;
  var m;
  while ((m = re.exec(text)) !== null) {
    var code = code4_(m[2]);
    if (!code || seen[code]) continue;
    seen[code] = true;
    var name = (m[1] || '').replace(/[（）()【】「」、。:：]/g, '').trim();
    pairs.push({ code: code, name: name });
  }
  return pairs;
}

/** TDnet（当日開示）の会社名一覧を取得 */
function fetchTodayDisclosures_() {
  try {
    var url = 'https://webapi.yanoshin.jp/webapi/tdnet/list/today.json?limit=1000';
    var res = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    if (res.getResponseCode() !== 200) return [];
    var json = JSON.parse(res.getContentText());
    var items = json.items || [];
    return items.map(function (it) {
      var t = it.Tdnet || it.tdnet || {};
      return (t.company_name || t.companyName || '').toString();
    }).filter(String);
  } catch (err) {
    return [];
  }
}

/** 会社名が当日開示リストに含まれるか（部分一致） */
function hasDisclosure_(name, list) {
  if (!name || !list || !list.length) return false;
  var key = name.replace(/\s|株式会社|（株）|HD|ホールディングス|グループ/g, '');
  if (key.length < 2) key = name;
  for (var i = 0; i < list.length; i++) {
    if (list[i] && list[i].indexOf(key) >= 0) return true;
  }
  return false;
}

/** 1時間ごとの自動更新トリガーを設置（重複は削除してから） */
function installHourlyTrigger() {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'updateAll') ScriptApp.deleteTrigger(triggers[i]);
  }
  ScriptApp.newTrigger('updateAll').timeBased().everyHours(1).create();
  SpreadsheetApp.getUi().alert('自動更新を設定しました（1時間ごとに updateAll を実行）。');
}
