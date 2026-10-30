// ═══════════════════════════════════════════
//  Sieg 日報 GAS
//  1. 新しいスプレッドシートを作成し、IDを下に貼る
//  2. Teams → Workflows →「Webhook 要求を受信したらチャットに投稿する」
//     テンプレートで日報グループ宛のフローを作成し、URLを下に貼る（使わない場合は空文字のままでOK）
//  3. Discordへ自動投稿する場合は、プロジェクトの設定 → スクリプト プロパティに
//     DISCORD_WEBHOOK として Discord の Webhook URL を登録する（コードには書かない）
//     獲得報告（ポータルの獲得報告フォーム）は、獲得報告チャンネルの Webhook URL を
//     DISCORD_GAIN_WEBHOOK として登録する
//  4. setup() を1回実行（ヘッダー作成）
//  5. デプロイ → ウェブアプリ →「全員（匿名含む）」で公開
//  6. デプロイURLを nippou-form の NIPPOU_API に貼る
// ═══════════════════════════════════════════

var SHEET_ID = '1ffkmHDt_sz1ZQ3ZB_fXIaJeFhuW0wSI7bqwMpC9Z33g';
// コードの版。/exec をブラウザで開くと表示されるので、再デプロイが反映されたかの確認に使う。
var CODE_VERSION = '2026-10-31';
var SHEET_NAME = '日報';
// Teams自動投稿を使わない場合は空文字のままでOK（コピペ運用ならシート保存のみで十分）
var TEAMS_WEBHOOK_URL = '';

function setup() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) sheet = ss.insertSheet(SHEET_NAME);
  sheet.getRange(1, 1, 1, 14).setValues([[
    '送信日時', '入店日', '名前', 'キャリア', '稼働場所',
    'アプローチ合計', 'アプローチ内訳', '獲得合計', '獲得内訳',
    'MNP目標', 'MNP通算進捗', '進捗率', '所感', '送信ID'
  ]]);
  sheet.setFrozenRows(1);
}

function getSheet_() {
  return SpreadsheetApp.openById(SHEET_ID).getSheetByName(SHEET_NAME);
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// 内訳配列 [{label, count}] → "キャッチ6 / 店内2"（0件は省略）
function fmtBreakdown_(arr) {
  if (!arr || !arr.length) return '';
  return arr
    .filter(function(x) { return Number(x.count) > 0; })
    .map(function(x) { return x.label + Number(x.count); })
    .join(' / ');
}

function sumBreakdown_(arr) {
  if (!arr || !arr.length) return 0;
  return arr.reduce(function(s, x) { return s + (Number(x.count) || 0); }, 0);
}

// MNP として数える商品（MNP目標の進捗と今月のMNPはこの合計）。
// ポータルの sieg-portal/sieg-products.js で区分が「MNP」の商品と同じにしておく
var MNP_LABELS = ['MNP', 'SMNP', 'YMNP', 'auMNP', 'UQMNP'];

// 内訳配列 [{label, count}] から MNP の件数を合計する
function mnpCount_(arr) {
  if (!arr || !arr.length) return 0;
  return arr.reduce(function(s, x) {
    return s + (MNP_LABELS.indexOf(String(x.label)) >= 0 ? (Number(x.count) || 0) : 0);
  }, 0);
}

// 日付を 'yyyy-MM-dd' に揃える（シートの入店日セルは Date、送信値は文字列で来る）
function ymd_(v, tz) {
  if (Object.prototype.toString.call(v) === '[object Date]') return isNaN(v.getTime()) ? '' : Utilities.formatDate(v, tz, 'yyyy-MM-dd');
  var m = String(v || '').match(/^(\d{4})[-\/](\d{1,2})[-\/](\d{1,2})/);
  if (!m) return '';
  return m[1] + '-' + ('0' + m[2]).slice(-2) + '-' + ('0' + m[3]).slice(-2);
}

// 獲得内訳 "SMNP3 / YMNP1 / 光新規1" から MNP の件数を合計する（過去の行もこの列から数える）
function mnpFromBreakdown_(s) {
  var parts = String(s || '').split(' / '), n = 0;
  for (var i = 0; i < parts.length; i++) {
    var m = parts[i].match(/^(.*\D)(\d+)$/);
    if (m && MNP_LABELS.indexOf(m[1]) >= 0) n += Number(m[2]);
  }
  return n;
}

// 同じ名前の、入店日と同じ月の集計（今回分を含めるため、追記後に呼ぶ）
//   gain   : 獲得合計
//   mnp    : MNP合計
//   days   : 稼働日数（同じ入店日に複数回送っていても1日と数える）
//   mnpAvg : 1稼働日あたりのMNP（小数第1位まで）
function monthlyStats_(name, dateYmd) {
  var sheet = getSheet_();
  var tz = sheet.getParent().getSpreadsheetTimeZone();
  var month = (dateYmd || Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd')).slice(0, 7);
  var data = sheet.getDataRange().getValues();
  var gain = 0, mnp = 0, days = {};
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][2]) !== name) continue;
    var d = ymd_(data[i][1], tz);
    if (d.slice(0, 7) !== month) continue;
    gain += Number(data[i][7]) || 0;
    mnp += mnpFromBreakdown_(data[i][8]);
    days[d] = true;
  }
  var dayCount = Object.keys(days).length;
  return {
    gain: gain,
    mnp: mnp,
    days: dayCount,
    mnpAvg: dayCount ? Math.round(mnp / dayCount * 10) / 10 : 0
  };
}

// 進捗バー "■■■□□□□□□□"
function bar_(pct) {
  var filled = Math.max(0, Math.min(10, Math.round(pct / 10)));
  return Array(filled + 1).join('■') + Array(10 - filled + 1).join('□');
}

// ═══ Discord送信 ═══
// スクリプトプロパティ「DISCORD_WEBHOOK」にURLが設定されていれば #日報 へ自動投稿
//
// meta には doPost で算出済みの値を渡す：
//   { dateStr, pct, cumulative, prev, mnpCount, target, monthly, monthlyMnp, workDays, mnpAvg,
//     approachTotal, gainTotal }
// 渡っていない項目は p から計算し直すので、meta が dateStr だけでも動作する
// （ただし monthly / monthlyMnp / workDays / mnpAvg はシート由来のため、渡さないと 0 表示になる）
//
// 戻り値：URL未設定なら null、投稿を試みたら成功/失敗の true/false（doPost の応答に使う）
function sendToDiscord_(copyText, p, meta) {
  var url = PropertiesService.getScriptProperties().getProperty('DISCORD_WEBHOOK');
  if (!url) return null;

  var m = meta || {};

  var target        = Number(m.target) || Number(p.target) || 0;
  var prev          = Number(m.prev) || Number(p.prev) || 0;
  var mnpCount      = m.mnpCount != null ? (Number(m.mnpCount) || 0) : mnpCount_(p.gains);
  var cumulative    = m.cumulative != null ? (Number(m.cumulative) || 0) : (prev + mnpCount);
  var approachTotal = m.approachTotal != null ? (Number(m.approachTotal) || 0) : sumBreakdown_(p.approaches);
  var gainTotal     = m.gainTotal != null ? (Number(m.gainTotal) || 0) : sumBreakdown_(p.gains);
  var monthly       = Number(m.monthly) || 0;
  var monthlyMnp    = Number(m.monthlyMnp) || 0;
  var workDays      = Number(m.workDays) || 0;
  var mnpAvg        = Number(m.mnpAvg) || 0;
  // doPost の pct は target 未設定のとき '' なので、ここでは数値に寄せる
  var pct           = target > 0 ? (Number(m.pct) || Math.round(cumulative / target * 100)) : 0;

  var hasTarget = target > 0;
  var achieved  = hasTarget && pct >= 100;

  // ─── 進捗に応じたカラー ───
  var color = 0x8A93A6;                              // target未設定 / 50%未満
  if (hasTarget && pct >= 100)     color = 0x4CD97B; // 目標達成
  else if (hasTarget && pct >= 50) color = 0xE8A33D; // 50%以上100%未満

  // ─── fields ───
  var withTotal = function(total, breakdown) {
    return '**' + total + '件**' + (breakdown ? '\n' + breakdown : '');
  };

  var fields = [
    { name: '📍 稼働場所', value: String(p.place || '—'),   inline: true },
    { name: '📶 キャリア', value: String(p.carrier || '—'), inline: true },
    { name: '🚶 アプローチ', value: withTotal(approachTotal, fmtBreakdown_(p.approaches)) },
    { name: '🎯 獲得',       value: withTotal(gainTotal, fmtBreakdown_(p.gains)) }
  ];

  // MNP進捗：target 未設定のときはフィールドごと省略
  if (hasTarget) {
    fields.push({
      name: '📊 MNP進捗',
      value: bar_(pct) + '　' + pct + '%\n'
           + '通算 ' + cumulative + '/' + target + '件'
           + '（前日 ' + prev + '件 ＋ 本日MNP ' + mnpCount + '件）'
    });
  }

  fields.push({
    name: '📈 今月通算',
    value: '獲得 **' + monthly + '件** ／ MNP **' + monthlyMnp + '件**\n'
         + '稼働 ' + workDays + '日 ・ MNP平均 **' + mnpAvg + '件/日**'
  });

  // 所感：入力があるときだけ追加（Discordのフィールド上限1024文字に丸める）
  var note = String(p.note || '').trim();
  if (note) {
    fields.push({ name: '💬 所感', value: note.length > 1000 ? note.slice(0, 1000) + '…' : note });
  }

  try {
    var res = UrlFetchApp.fetch(url, {
      method: 'post',
      contentType: 'application/json',
      muteHttpExceptions: true,
      payload: JSON.stringify({
        // 投稿者の名前を、Webhook の名前（Bot）ではなく日報を送った人の名前にする
        username: String(p.name || '日報').slice(0, 80),
        allowed_mentions: { parse: [] },
        embeds: [{
          title: '📋 日報｜' + String(p.name || '') + '　' + (m.dateStr || '') + (achieved ? ' 🎉' : ''),
          color: color,
          fields: fields,
          footer: { text: Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm') + ' 送信' }
        }]
      })
    });
    return res.getResponseCode() < 300;
  } catch (err) {
    return false;
  }
}

// ═══ Discord接続テスト ═══
// GASエディタで関数「checkDiscord」を選んで実行し、実行ログを見る。
// 日報は記録されず、Discordにテストメッセージが1件だけ投稿される。
function checkDiscord() {
  var props = PropertiesService.getScriptProperties();
  var keys = props.getKeys();
  Logger.log('登録済みのプロパティ名: ' + (keys.length ? keys.join(', ') : '(ひとつも無し)'));

  var url = props.getProperty('DISCORD_WEBHOOK');
  if (!url) {
    Logger.log('❌ DISCORD_WEBHOOK が未設定です。ここが原因です。');
    Logger.log('   → プロジェクトの設定 → スクリプト プロパティ → プロパティを追加');
    Logger.log('     名前: DISCORD_WEBHOOK / 値: Discordの日報チャンネルのWebhook URL');
    return;
  }
  // Webhook ID は秘密ではない（トークンは末尾の別セグメント）。
  // どのチャンネル宛か取り違えたときに気付けるよう、IDだけ出す。
  var idMatch = String(url).match(/\/webhooks\/(\d+)\//);
  Logger.log('✅ DISCORD_WEBHOOK は設定済み / Webhook ID: '
    + (idMatch ? idMatch[1] : '(URLの形式が不正です)'));

  try {
    var res = UrlFetchApp.fetch(url, {
      method: 'post',
      contentType: 'application/json',
      muteHttpExceptions: true,
      payload: JSON.stringify({ content: '✅ 日報システムからの接続テストです' })
    });
    var code = res.getResponseCode();
    Logger.log('HTTPステータス: ' + code);
    Logger.log('レスポンス本文: ' + (res.getContentText() || '(空 ＝ 正常)'));
    if (code < 300) {
      Logger.log('🎉 送信成功。Discordの日報チャンネルを確認してください。');
      Logger.log('   ここで届くのにフォームからは届かない場合は、デプロイが古いのが原因です。');
    } else if (code === 401 || code === 404) {
      Logger.log('❌ Webhook URL が無効です（削除済み、または貼り間違い）。作り直してください。');
    } else {
      Logger.log('❌ 送信失敗。上のレスポンス本文を確認してください。');
    }
  } catch (err) {
    Logger.log('❌ 例外が発生: ' + err);
    Logger.log('   「承認が必要」と出た場合は、承認してからもう一度実行してください。');
  }
}

// ═══ 二重登録の防止 ═══
// 通信が途中で切れると、記録は済んでいるのにフォームには失敗と出る。
// そこで送り直すと、シートに2行・Discordに2件入り、今月の合計も水増しされていた。
// フォームは同じ日報には同じ「送信ID」を付けて送るので、6時間以内に同じIDが
// 来たら、記録もDiscord投稿もせず、前回の結果をそのまま返す。
var DEDUPE_SEC = 21600;
function dedupeKey_(id) { return 'nippou:' + id; }

// ═══ 読み込み ═══
// /exec をブラウザで直接開くと、デプロイ確認の応答が返る。日報は記録しない。
//   ?action=gains&name=名前&date=yyyy-mm-dd … その人のその日の獲得報告
//   ?action=gainStats&from=yyyy-mm          … その月以降の獲得報告の月ごとの合計（戦闘力用）
function doGet(e) {
  var q = (e && e.parameter) || {};
  try {
    if (q.action === 'gains') return json_(dayGains_(String(q.name || ''), String(q.date || '')));
    if (q.action === 'gainStats') return json_(gainStats_(String(q.from || '')));
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  }
  return json_({ ok: true, status: 'ready', version: CODE_VERSION, dedupe: true, gain: true });
}

function doPost(e) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(25000);
    var res = handlePost_(e);
    return json_(res);
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  } finally {
    try { lock.releaseLock(); } catch (ignore) {}
  }
}

function handlePost_(e) {
  var p = JSON.parse(e.postData.contents);
  var requestId = String(p.requestId || '').slice(0, 100);
  var cache = CacheService.getScriptCache();
  if (requestId) {
    var done = cache.get(dedupeKey_(requestId));
    if (done) {
      var prevRes = JSON.parse(done);
      prevRes.duplicate = true;   // 2回目以降。記録・投稿はしていない
      return prevRes;
    }
  }
  // 獲得報告の取り消しは、何度来ても結果が同じなので送信IDでの確認はしない
  if (p.action === 'gainUndo') return undoGain_(p);
  if (p.action === 'fighter') return setFighter_(p);
  var out = p.action === 'gain' ? buildGain_(p, requestId) : buildReport_(p, requestId);
  if (out.ok && requestId) {
    try { cache.put(dedupeKey_(requestId), JSON.stringify(out), DEDUPE_SEC); } catch (ignore) {}
  }
  return out;
}

function buildReport_(p, requestId) {
  if (!p.name) return { ok: false, error: '名前がありません' };

  var approachTotal = sumBreakdown_(p.approaches);
  var gainTotal = sumBreakdown_(p.gains);
  var mnpCount = mnpCount_(p.gains);
  var target = Number(p.target) || 0;
  var prev = Number(p.prev) || 0;
  var cumulative = prev + mnpCount;
  var pct = target > 0 ? Math.round(cumulative / target * 100) : '';

  getSheet_().appendRow([
    new Date(),
    p.date ? String(p.date) : '',
    String(p.name),
    String(p.carrier || ''),
    String(p.place || ''),
    approachTotal,
    fmtBreakdown_(p.approaches),
    gainTotal,
    fmtBreakdown_(p.gains),
    target || '',
    cumulative,
    pct !== '' ? pct + '%' : '',
    String(p.note || '').trim(),
    requestId || ''
  ]);

  var stats = monthlyStats_(String(p.name), ymd_(p.date, 'Asia/Tokyo'));
  var monthly = stats.gain;

  // ─── 日付表記 ───
  var dateStr = '';
  if (p.date) {
    var d = new Date(p.date);
    if (!isNaN(d.getTime())) {
      var youbi = ['日', '月', '火', '水', '木', '金', '土'][d.getDay()];
      dateStr = (d.getMonth() + 1) + '/' + d.getDate() + '（' + youbi + '）';
    }
  }

  // ─── コピー用テキスト（Teamsにそのまま貼れる形）───
  var lines = [];
  lines.push('日報');
  lines.push('入店日：' + (dateStr || '-'));
  lines.push('キャリア：' + (p.carrier || '-'));
  lines.push('稼働場所：' + (p.place || '-'));
  lines.push('');
  lines.push('【アプローチ】' + approachTotal + '件');
  if (fmtBreakdown_(p.approaches)) lines.push(fmtBreakdown_(p.approaches));
  lines.push('');
  lines.push('【獲得】' + gainTotal + '件');
  if (fmtBreakdown_(p.gains)) lines.push(fmtBreakdown_(p.gains));
  if (target > 0) {
    lines.push('');
    lines.push('🎯 MNP目標 ' + target + '件中 通算 ' + cumulative + '件（' + pct + '%）');
    lines.push(bar_(pct));
    if (prev > 0) lines.push('（本日MNP ' + mnpCount + '件 ＋ 前日まで ' + prev + '件）');
  }
  lines.push('');
  lines.push('📈 今月通算：' + monthly + '件');
  lines.push('📱 今月MNP：' + stats.mnp + '件（稼働' + stats.days + '日・平均' + stats.mnpAvg + '件/日）');
  if (p.note && String(p.note).trim()) {
    lines.push('');
    lines.push('【所感】');
    lines.push(String(p.note).trim());
  }
  var copyText = lines.join('\n');

  // ─── Teams自動投稿（URLが設定されている場合のみ試行）───
  var teamsOk = null;
  if (TEAMS_WEBHOOK_URL) {
    var facts = [
      { title: '入店日', value: dateStr || '-' },
      { title: 'キャリア', value: String(p.carrier || '-') },
      { title: '稼働場所', value: String(p.place || '-') },
      { title: 'アプローチ', value: approachTotal + '件' + (fmtBreakdown_(p.approaches) ? '（' + fmtBreakdown_(p.approaches) + '）' : '') },
      { title: '獲得', value: gainTotal + '件' + (fmtBreakdown_(p.gains) ? '（' + fmtBreakdown_(p.gains) + '）' : '') }
    ];
    var body = [
      { type: 'TextBlock', text: '📋 日報｜' + p.name, weight: 'Bolder', size: 'Medium', wrap: true },
      { type: 'FactSet', facts: facts }
    ];
    if (target > 0) {
      body.push({
        type: 'TextBlock',
        text: '🎯 MNP目標 ' + target + '件中 通算 ' + cumulative + '件（' + pct + '%）\n' + bar_(pct) +
          (prev > 0 ? '\n（本日MNP ' + mnpCount + '件 ＋ 前日まで ' + prev + '件）' : ''),
        wrap: true
      });
    }
    body.push({
      type: 'TextBlock',
      text: '📈 今月通算：' + monthly + '件\n📱 今月MNP：' + stats.mnp + '件（稼働' + stats.days + '日・平均' + stats.mnpAvg + '件/日）',
      wrap: true
    });
    if (p.note && String(p.note).trim()) {
      body.push({ type: 'TextBlock', text: '💬 所感\n' + String(p.note).trim(), wrap: true, separator: true });
    }
    var card = {
      type: 'message',
      attachments: [{
        contentType: 'application/vnd.microsoft.card.adaptive',
        content: { '$schema': 'http://adaptivecards.io/schemas/adaptive-card.json', type: 'AdaptiveCard', version: '1.4', body: body }
      }]
    };
    try {
      var res = UrlFetchApp.fetch(TEAMS_WEBHOOK_URL, {
        method: 'post', contentType: 'application/json', payload: JSON.stringify(card), muteHttpExceptions: true
      });
      teamsOk = res.getResponseCode() < 300;
    } catch (err2) {
      teamsOk = false;
    }
  }

  // ─── Discord自動投稿（スクリプトプロパティにURLが設定されている場合のみ試行）───
  var discordOk = sendToDiscord_(copyText, p, {
    dateStr: dateStr,
    pct: pct,
    cumulative: cumulative,
    prev: prev,
    mnpCount: mnpCount,
    target: target,
    monthly: monthly,
    monthlyMnp: stats.mnp,
    workDays: stats.days,
    mnpAvg: stats.mnpAvg,
    approachTotal: approachTotal,
    gainTotal: gainTotal
  });

  return {
    ok: true,
    dedupe: true,   // このデプロイは送信IDでの二重登録防止に対応している（フォームが自動再試行の判断に使う）
    teams: teamsOk,
    discord: discordOk,
    monthly: monthly,
    monthlyMnp: stats.mnp,
    workDays: stats.days,
    mnpAvg: stats.mnpAvg,
    text: copyText
  };
}

// ═══════════════════════════════════════════
//  獲得報告（ポータルの獲得報告フォームから）
//  1回の報告を、商品ごとに1行ずつ「獲得」シートに書く。シートは無ければ作る。
//  点数や区分はここでは持たない（ポータルの sieg-products.js で計算する）。
// ═══════════════════════════════════════════
var GAIN_SHEET = '獲得';
var GAIN_HEAD = ['送信日時', '獲得日', '名前', 'キャリア', '商品', '件数', 'メモ', '送信ID', 'DiscordメッセージID'];
var CARRIER_LABEL = { docomo: 'ドコモ', Softbank: 'SB', au: 'au' };

function gainSheet_() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName(GAIN_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(GAIN_SHEET);
    sheet.getRange(1, 1, 1, GAIN_HEAD.length).setValues([GAIN_HEAD]);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function todayYmd_() { return Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd'); }

// 'yyyy-mm-dd' → '10/30（木）'
function dayLabel_(ymd) {
  var m = String(ymd).match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (!m) return '';
  var d = new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]));
  return Number(m[2]) + '/' + Number(m[3]) + '（' + ['日', '月', '火', '水', '木', '金', '土'][d.getDay()] + '）';
}

function cleanGainItems_(arr) {
  var out = [];
  (arr || []).forEach(function(x) {
    var label = String((x && x.label) || '').trim().slice(0, 30);
    var count = Math.floor(Number(x && x.count) || 0);
    if (label && count > 0) out.push({ label: label, count: Math.min(count, 99) });
  });
  return out;
}

function buildGain_(p, requestId) {
  var name = String(p.name || '').trim().slice(0, 20);
  if (!name) return { ok: false, error: '名前がありません' };
  var carrier = String(p.carrier || '');
  if (!CARRIER_LABEL[carrier]) return { ok: false, error: 'キャリアを選んでください' };
  var items = cleanGainItems_(p.items);
  if (!items.length) return { ok: false, error: '件数が入っていません' };
  var date = ymd_(p.date, 'Asia/Tokyo') || todayYmd_();
  var memo = String(p.memo || '').trim().slice(0, 500);

  var sheet = gainSheet_();
  var now = new Date();
  var first = sheet.getLastRow() + 1;
  sheet.getRange(first, 1, items.length, GAIN_HEAD.length).setValues(items.map(function(x) {
    return [now, date, name, carrier, x.label, x.count, memo, requestId || '', ''];
  }));

  var sent = sendGainToDiscord_(name, carrier, date, items, memo);
  if (sent && sent.id) {
    var ids = items.map(function() { return [sent.id]; });
    sheet.getRange(first, GAIN_HEAD.length, items.length, 1).setValues(ids);
  }
  var day = dayGains_(name, date);
  return {
    ok: true,
    dedupe: true,
    id: requestId || '',
    discord: sent ? sent.ok : null,
    items: day.items,
    reports: day.reports
  };
}

// Discord の獲得報告チャンネルに、いつもの書き方（「ハイエンド1件」を1行ずつ）で投稿する。
// 取り消しで消せるように ?wait=true でメッセージIDを受け取る。
// 戻り値：URL未設定なら null、それ以外は { ok, id }
function sendGainToDiscord_(name, carrier, date, items, memo) {
  var url = PropertiesService.getScriptProperties().getProperty('DISCORD_GAIN_WEBHOOK');
  if (!url) return null;
  var lines = ['**' + CARRIER_LABEL[carrier] + '**　' + dayLabel_(date)];
  items.forEach(function(x) { lines.push(x.label + x.count + '件'); });
  if (memo) lines.push('', memo);
  try {
    var res = UrlFetchApp.fetch(url + (url.indexOf('?') >= 0 ? '&' : '?') + 'wait=true', {
      method: 'post',
      contentType: 'application/json',
      muteHttpExceptions: true,
      payload: JSON.stringify({
        username: name,
        content: lines.join('\n').slice(0, 1900),
        allowed_mentions: { parse: [] }   // メモに @everyone などがあっても通知を飛ばさない
      })
    });
    var ok = res.getResponseCode() < 300, id = '';
    if (ok) { try { id = String(JSON.parse(res.getContentText()).id || ''); } catch (ignore) {} }
    return { ok: ok, id: id };
  } catch (err) {
    return { ok: false, id: '' };
  }
}

// 送った報告を取り消す（本人の名前と送信IDが両方合う行だけ消す）。Discord の投稿も消す
function undoGain_(p) {
  var name = String(p.name || ''), id = String(p.id || '');
  if (!name || !id) return { ok: false, error: '取り消す報告が分かりません' };
  var sheet = gainSheet_();
  var data = sheet.getDataRange().getValues();
  var rows = [], msgId = '';
  for (var i = data.length - 1; i >= 1; i--) {
    if (String(data[i][7]) === id && String(data[i][2]) === name) {
      rows.push(i + 1);
      if (data[i][8]) msgId = String(data[i][8]);
    }
  }
  rows.forEach(function(r) { sheet.deleteRow(r); });   // 下の行から消すので番号はずれない

  var discord = null;
  var url = PropertiesService.getScriptProperties().getProperty('DISCORD_GAIN_WEBHOOK');
  if (url && msgId) {
    try {
      var res = UrlFetchApp.fetch(url.split('?')[0] + '/messages/' + msgId, { method: 'delete', muteHttpExceptions: true });
      discord = res.getResponseCode() < 300 || res.getResponseCode() === 404;
    } catch (err) {
      discord = false;
    }
  }
  return { ok: true, removed: rows.length, discord: discord };
}

// その人のその日の獲得報告。items は商品ごとの合計、reports は送った報告ごと（新しい順）
function dayGains_(name, date) {
  if (!name) return { ok: false, error: '名前がありません' };
  date = ymd_(date, 'Asia/Tokyo') || todayYmd_();
  var sheet = gainSheet_();
  var tz = sheet.getParent().getSpreadsheetTimeZone();
  var data = sheet.getDataRange().getValues();
  var sum = {}, order = [], reports = {}, rorder = [];
  for (var i = 1; i < data.length; i++) {
    var r = data[i];
    if (String(r[2]) !== name || ymd_(r[1], tz) !== date) continue;
    var carrier = String(r[3]), label = String(r[4]), count = Number(r[5]) || 0;
    var key = carrier + '\t' + label;
    if (!(key in sum)) { sum[key] = 0; order.push(key); }
    sum[key] += count;
    var rid = String(r[7] || ('row' + i));
    if (!reports[rid]) {
      var at = Object.prototype.toString.call(r[0]) === '[object Date]' ? Utilities.formatDate(r[0], 'Asia/Tokyo', 'HH:mm') : '';
      reports[rid] = { id: String(r[7] || ''), at: at, carrier: carrier, memo: String(r[6] || ''), items: [] };
      rorder.push(rid);
    }
    reports[rid].items.push({ label: label, count: count });
  }
  return {
    ok: true,
    date: date,
    items: order.map(function(k) { var p = k.split('\t'); return { carrier: p[0], label: p[1], count: sum[k] }; }),
    reports: rorder.reverse().map(function(k) { return reports[k]; })
  };
}

// from の月（yyyy-mm）以降の、月・名前・キャリア・商品ごとの合計。from が無ければ5か月前から
//   rows: [[月, 名前, キャリア, 商品, 件数], …]
function gainStats_(from) {
  if (!/^\d{4}-\d{2}$/.test(from)) {
    var d = new Date();
    d.setDate(1);
    d.setMonth(d.getMonth() - 5);
    from = Utilities.formatDate(d, 'Asia/Tokyo', 'yyyy-MM');
  }
  var sheet = gainSheet_();
  var tz = sheet.getParent().getSpreadsheetTimeZone();
  var data = sheet.getDataRange().getValues();
  var sum = {}, order = [];
  for (var i = 1; i < data.length; i++) {
    var r = data[i], month = ymd_(r[1], tz).slice(0, 7);
    if (!month || month < from) continue;
    var key = [month, String(r[2]), String(r[3]), String(r[4])].join('\t');
    if (!(key in sum)) { sum[key] = 0; order.push(key); }
    sum[key] += Number(r[5]) || 0;
  }
  return {
    ok: true,
    from: from,
    month: todayYmd_().slice(0, 7),
    rows: order.map(function(k) { return k.split('\t').concat([sum[k]]); }),
    fighters: fighters_()
  };
}

// ═══ 戦闘力で使うキャラ ═══
// 「ファイター」シートに 名前・キャラの番号（'01'〜）を1人1行で持つ。シートは無ければ作る
var FIGHTER_SHEET = 'ファイター';
function fighterSheet_() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName(FIGHTER_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(FIGHTER_SHEET);
    sheet.getRange(1, 1, 1, 3).setValues([['名前', 'キャラ', '更新日時']]);
    sheet.setFrozenRows(1);
  }
  return sheet;
}
function fighters_() {
  var data = fighterSheet_().getDataRange().getValues(), out = {};
  for (var i = 1; i < data.length; i++) {
    // 数字として読まれても '01' の形にそろえる
    if (data[i][0]) out[String(data[i][0])] = ('0' + String(data[i][1]).replace(/^'/, '')).slice(-2);
  }
  return out;
}
function setFighter_(p) {
  var name = String(p.name || '').trim().slice(0, 20), id = String(p.id || '');
  if (!name || !/^\d{2}$/.test(id)) return { ok: false, error: '名前かキャラが正しくありません' };
  var sheet = fighterSheet_(), data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]) === name) {
      sheet.getRange(i + 1, 2, 1, 2).setValues([["'" + id, new Date()]]);
      return { ok: true };
    }
  }
  sheet.getRange(data.length + 1, 1, 1, 3).setValues([[name, "'" + id, new Date()]]);
  return { ok: true };
}
