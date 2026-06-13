/**
 * 注目上場企業ウォッチ＆個別深掘り 2段階システム
 * ------------------------------------------------------------
 * このスクリプトは Googleスプレッドシート上で動きます（ターミナル不要）。
 * 使い方は appscript/手順書.md を参照してください。
 *
 * 構成シート:
 *   Watchlist      … 監視銘柄リスト（楽天証券で買える銘柄。自由に行追加可）
 *   ①注目ランキング … 今注目の銘柄をスコア順に表示（Stage 1）
 *   ②個別ウォッチ   … 1社を選んで情報を集中表示（Stage 2）
 *   Sources        … 取り込むニュースRSSの一覧
 *
 * 無料・合法の範囲で動かすための方針:
 *   - 株価/出来高/騰落率 … GOOGLEFINANCE（Googleの公式データ・約20分遅延）
 *   - ニュース           … Google News RSS（IMPORTFEED / UrlFetchApp）
 *   - 適時開示           … TDnet（Yanoshin 公開API）
 *   - Yahoo!ファイナンス・株探の画面スクレイピングは規約違反のため行いません。
 */

// ===== 設定（必要なら数値だけ調整してください）=====
var CFG = {
  watchlistSheet: 'Watchlist',
  rankingSheet: '①注目ランキング',
  detailSheet: '②個別ウォッチ',
  sourcesSheet: 'Sources',
  newsDays: 7,           // ニュース件数を数える対象日数
  rankingTopN: 30,       // ランキング上位の表示件数
  // 合成スコアの重み（合計が1でなくてOK。相対比較に使います）
  weight: { change: 0.35, volume: 0.20, news: 0.30, disclosure: 0.15 },
  newsHl: 'ja', newsGl: 'JP', newsCeid: 'JP:ja'
};

// スプレッドシートを開いたときにメニューを追加
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('📈 株ウォッチ')
    .addItem('① 今すぐ全体を更新', 'updateAll')
    .addItem('② 個別ウォッチのニュースを更新', 'refreshDetailNews')
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
  setupDetail_(ss);
  SpreadsheetApp.getUi().alert(
    'セットアップ完了。\n\n' +
    '1) 「Watchlist」シートに銘柄を貼り付けてください（watchlist_seed.csv の中身）。\n' +
    '2) メニュー「📈 株ウォッチ → ① 今すぐ全体を更新」を実行。\n' +
    '3) 「⏰ 自動更新(1時間ごと)をON」で自動化できます。');
}

// ---------- Watchlist ----------
function setupWatchlist_(ss) {
  var sh = ss.getSheetByName(CFG.watchlistSheet) || ss.insertSheet(CFG.watchlistSheet);
  var headers = ['Code', 'Name', 'Market', 'Symbol', 'NewsKeyword', 'XSearchURL',
                 'Price', 'ChangePct', 'Volume', 'NewsCount', 'Disclosure', 'Score', 'UpdatedAt'];
  sh.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
  sh.setFrozenRows(1);
  sh.getRange('A1:M1').setBackground('#1f3864').setFontColor('white');
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

// ---------- ②個別ウォッチ ----------
function setupDetail_(ss) {
  var sh = ss.getSheetByName(CFG.detailSheet) || ss.insertSheet(CFG.detailSheet);
  sh.clear();
  sh.getRange('A1').setValue('② 個別企業ウォッチ').setFontWeight('bold').setFontSize(14);
  sh.getRange('A3').setValue('銘柄コードを選択 →').setFontWeight('bold');

  // B3 にウォッチリストの Code から選べるプルダウンを設定
  var wl = "='" + CFG.watchlistSheet + "'!A2:A";
  var rule = SpreadsheetApp.newDataValidation()
    .requireValueInRange(SpreadsheetApp.getActiveSpreadsheet()
      .getSheetByName(CFG.watchlistSheet).getRange('A2:A'), true)
    .setAllowInvalid(true).build();
  sh.getRange('B3').setDataValidation(rule);

  var WL = CFG.watchlistSheet;
  // 基本情報（Watchlist から VLOOKUP ＋ GOOGLEFINANCE）
  var rows = [
    ['会社名',   "=IFERROR(VLOOKUP($B$3,'" + WL + "'!A:F,2,FALSE),\"—\")"],
    ['市場',     "=IFERROR(VLOOKUP($B$3,'" + WL + "'!A:F,3,FALSE),\"—\")"],
    ['シンボル', "=IFERROR(VLOOKUP($B$3,'" + WL + "'!A:F,4,FALSE),\"—\")"],
    ['現在値',   "=IFERROR(GOOGLEFINANCE(B6),\"—\")"],   // B6 = シンボル
    ['前日比%',  "=IFERROR(GOOGLEFINANCE(B6,\"changepct\"),\"—\")"],
    ['出来高',   "=IFERROR(GOOGLEFINANCE(B6,\"volume\"),\"—\")"],
    ['PER',      "=IFERROR(GOOGLEFINANCE(B6,\"pe\"),\"—\")"],
    ['時価総額', "=IFERROR(GOOGLEFINANCE(B6,\"marketcap\"),\"—\")"],
    ['30日チャート', "=IFERROR(SPARKLINE(QUERY(GOOGLEFINANCE(B6,\"price\",TODAY()-30,TODAY()),\"select Col2\"),{\"charttype\",\"line\"}),\"—\")"]
  ];
  sh.getRange(4, 1, rows.length, 2).setValues(rows);
  sh.getRange('A4:A12').setFontWeight('bold');

  // Xへの入口
  sh.getRange('A14').setValue('X検索（クリックで開く）').setFontWeight('bold');
  sh.getRange('B14').setFormula(
    "=IFERROR(HYPERLINK(VLOOKUP($B$3,'" + WL + "'!A:F,6,FALSE),\"Xでこの銘柄を検索\"),\"—\")");

  // 公式開示・IR
  sh.getRange('A15').setValue('適時開示(TDnet)').setFontWeight('bold');
  sh.getRange('B15').setFormula(
    "=HYPERLINK(\"https://www.release.tdnet.info/inbs/I_main_00.html\",\"TDnet 適時開示閲覧\")");
  sh.getRange('A16').setValue('EDINET(法定開示)').setFontWeight('bold');
  sh.getRange('B16').setFormula(
    "=HYPERLINK(\"https://disclosure.edinet-fsa.go.jp/\",\"EDINETで検索\")");

  // ニュース見出し（Google News RSS を会社名で検索して自動表示）
  sh.getRange('A18').setValue('▼ この企業のニュース（自動取得・新しい順）').setFontWeight('bold').setFontSize(12);
  var kw = "VLOOKUP($B$3,'" + WL + "'!A:F,5,FALSE)";
  var newsUrl = "\"https://news.google.com/rss/search?q=\"&ENCODEURL(" + kw +
                ")&\"&hl=" + CFG.newsHl + "&gl=" + CFG.newsGl + "&ceid=" + CFG.newsCeid + "\"";
  sh.getRange('A19').setFormula(
    "=IFERROR(IMPORTFEED(" + newsUrl + ",\"items\",TRUE,20),\"（銘柄を選ぶとニュースが表示されます）\")");

  sh.getRange('B3').setValue('7203'); // 初期表示
  sh.setColumnWidth(1, 160);
  sh.setColumnWidth(2, 520);
}

/**
 * 全体更新：株価系の数式を入れ、ニュース件数・開示を取得し、スコアを計算して
 * ランキングを並べ替える。メニューまたは1時間トリガーから呼ばれる。
 */
function updateAll() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(CFG.watchlistSheet);
  var last = sh.getLastRow();
  if (last < 2) { SpreadsheetApp.getUi().alert('Watchlist に銘柄を貼り付けてください。'); return; }
  var n = last - 1;

  // 株価系の数式（G:I）を毎回入れ直す（行追加に追従）
  var symRange = sh.getRange(2, 4, n, 1); // D列 Symbol
  var priceF = [], chgF = [], volF = [];
  for (var i = 0; i < n; i++) {
    var r = i + 2;
    priceF.push(['=IFERROR(GOOGLEFINANCE(D' + r + '),\"\")']);
    chgF.push(['=IFERROR(GOOGLEFINANCE(D' + r + ',\"changepct\"),\"\")']);
    volF.push(['=IFERROR(GOOGLEFINANCE(D' + r + ',\"volume\"),\"\")']);
  }
  sh.getRange(2, 7, n, 1).setFormulas(priceF); // G Price
  sh.getRange(2, 8, n, 1).setFormulas(chgF);   // H ChangePct
  sh.getRange(2, 9, n, 1).setFormulas(volF);   // I Volume
  SpreadsheetApp.flush();

  // データ読み込み
  var data = sh.getRange(2, 1, n, 11).getValues(); // A..K
  var todaysDisclosures = fetchTodayDisclosures_(); // 会社名の配列

  var newsCounts = [], discFlags = [];
  for (var j = 0; j < n; j++) {
    var keyword = data[j][4] || data[j][1];
    newsCounts.push(keyword ? fetchNewsCount_(keyword) : 0);
    var name = (data[j][1] || '').toString();
    discFlags.push(hasDisclosure_(name, todaysDisclosures) ? 1 : 0);
    Utilities.sleep(150); // RSSサーバーへの配慮
  }

  // スコア計算（相対ランクで正規化）
  var chg = data.map(function (r) { return Math.abs(parseFloat(r[7]) || 0); });
  var vol = data.map(function (r) { return parseFloat(r[8]) || 0; });
  var scores = [];
  for (var k = 0; k < n; k++) {
    var s = CFG.weight.change * pct_(chg, chg[k])
          + CFG.weight.volume * pct_(vol, vol[k])
          + CFG.weight.news   * pct_(newsCounts, newsCounts[k])
          + CFG.weight.disclosure * discFlags[k];
    scores.push([Math.round(s * 1000) / 10]); // 0-100目安
  }

  // 書き戻し（J NewsCount, K Disclosure, L Score, M UpdatedAt）
  sh.getRange(2, 10, n, 1).setValues(newsCounts.map(function (v) { return [v]; }));
  sh.getRange(2, 11, n, 1).setValues(discFlags.map(function (v) { return [v ? '◯' : '']; }));
  sh.getRange(2, 12, n, 1).setValues(scores);
  var now = new Date();
  sh.getRange(2, 13, n, 1).setValue(now).setNumberFormat('yyyy/MM/dd HH:mm');

  buildRanking_(ss);
}

/** 相対パーセンタイル（0-1）。配列内でvalが上位なら1に近い */
function pct_(arr, val) {
  var valid = arr.filter(function (x) { return !isNaN(x); });
  if (!valid.length) return 0;
  var below = valid.filter(function (x) { return x < val; }).length;
  return below / valid.length;
}

/** Watchlist のスコアを読み、①注目ランキングに並べ替えて出力 */
function buildRanking_(ss) {
  var wl = ss.getSheetByName(CFG.watchlistSheet);
  var last = wl.getLastRow();
  if (last < 2) return;
  var n = last - 1;
  var d = wl.getRange(2, 1, n, 12).getValues(); // A..L
  var rows = d.map(function (r) {
    return { code: r[0], name: r[1], xurl: r[5], chg: r[7], vol: r[8],
             news: r[9], disc: r[10], score: parseFloat(r[11]) || 0 };
  });
  rows.sort(function (a, b) { return b.score - a.score; });
  rows = rows.slice(0, CFG.rankingTopN);

  var rk = ss.getSheetByName(CFG.rankingSheet);
  rk.getRange(5, 1, Math.max(rk.getLastRow() - 4, 1), 10).clearContent();
  var out = rows.map(function (r, i) {
    return [i + 1, r.code, r.name, r.score, r.chg, r.vol, r.news, r.disc ? '◯' : '',
            r.code, r.xurl];
  });
  if (out.length) {
    rk.getRange(5, 1, out.length, 10).setValues(out);
    // 「個別を見る」列(I)を②へジャンプするリンクに、X列(J)をリンクにする
    for (var i = 0; i < out.length; i++) {
      var row = 5 + i;
      rk.getRange(row, 9).setFormula(
        '=HYPERLINK("#gid=' + ss.getSheetByName(CFG.detailSheet).getSheetId() +
        '","▶ ' + out[i][1] + ' を見る")');
      if (out[i][9]) {
        rk.getRange(row, 10).setFormula('=HYPERLINK("' + out[i][9] + '","X検索")');
      }
    }
    rk.getRange(5, 4, out.length, 1).setNumberFormat('0.0');
  }
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
