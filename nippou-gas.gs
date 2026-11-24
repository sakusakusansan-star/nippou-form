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
var CODE_VERSION = '2026-11-25';
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
  var url = webhook_('DISCORD_WEBHOOK');
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

  return postDiscord_(url, {
    // 投稿者の名前を、Webhook の名前（Bot）ではなく日報を送った人の名前にする
    username: String(p.name || '日報').slice(0, 80),
    allowed_mentions: { parse: [] },
    embeds: [{
      title: '📋 日報｜' + String(p.name || '') + '　' + (m.dateStr || '') + (achieved ? ' 🎉' : ''),
      color: color,
      fields: fields,
      footer: { text: Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm') + ' 送信' }
    }]
  }).ok;
}

// ═══ Discord への投稿（送り直し付き） ═══
// Google のサーバーから Discord へ送ると、混み合いで一時的に断られる（429）ことや、
// Discord 側の一時的なエラー（5xx）・通信の失敗がときどきある。
// そのときは少し待って最大3回まで送る。429 のときは Discord が指定した秒数だけ待つ（最大5秒）。
// 4xx（429以外）は送り直しても同じなので、すぐ失敗にする。
// 戻り値：{ ok, code, body }（code は最後の HTTP ステータス、通信の失敗は 0）
function postDiscord_(url, payload) {
  var last = { ok: false, code: 0, body: '' };
  for (var attempt = 0; attempt < 3; attempt++) {
    var wait = 0;
    try {
      var res = UrlFetchApp.fetch(url, {
        method: 'post',
        contentType: 'application/json',
        muteHttpExceptions: true,
        payload: JSON.stringify(payload)
      });
      var code = res.getResponseCode(), body = res.getContentText() || '';
      last = { ok: code < 300, code: code, body: body };
      if (last.ok) return last;
      if (code === 429) {
        var after = 1;
        try { after = Number(JSON.parse(body).retry_after) || 1; } catch (ignore) {}
        wait = Math.min(5000, Math.ceil(after * 1000) + 250);
      } else if (code >= 500) {
        wait = 1000 * (attempt + 1);
      } else {
        return last;
      }
    } catch (err) {
      last = { ok: false, code: 0, body: String(err) };
      wait = 1000 * (attempt + 1);
    }
    if (attempt < 2) Utilities.sleep(wait);
  }
  return last;
}

// ═══ みんなの状況（team.html）の合言葉 ═══
// 合言葉はスクリプト プロパティ TEAM_PASS に入れる（コードや GitHub には書かない）。
// 正しければ「通行証」（合言葉から作る記号）を返し、端末はそれを覚えて次からは聞かれない。
// TEAM_PASS を変えると通行証も変わるので、前の合言葉で入っていた端末は入り直しになる。
// まちがいが10分で10回続いたら、しばらく受け付けない（総当たり対策）。
var TEAM_FAIL_MAX = 10;
var TEAM_FAIL_SEC = 600;
function teamPass_(p) {
  var pass = webhook_('TEAM_PASS');
  if (!pass) return { ok: false, notSet: true, error: '合言葉がまだ設定されていません（GAS のスクリプト プロパティ TEAM_PASS）' };
  var want = teamToken_(pass);
  if (p.token !== undefined) {
    return String(p.token) === want ? { ok: true } : { ok: false, expired: true, error: '合言葉が変わりました。入れ直してください' };
  }
  var cache = CacheService.getScriptCache();
  var fails = Number(cache.get('teamPassFails') || 0);
  if (fails >= TEAM_FAIL_MAX) return { ok: false, locked: true, error: 'まちがいが続いたので、10分ほど待ってからもう一度入れてください' };
  if (String(p.pass || '').trim() !== pass) {
    cache.put('teamPassFails', String(fails + 1), TEAM_FAIL_SEC);
    return { ok: false, error: '合言葉がちがいます' };
  }
  return { ok: true, token: want };
}
function teamToken_(pass) {
  var d = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, 'sieg-team:' + pass, Utilities.Charset.UTF_8);
  return Utilities.base64EncodeWebSafe(d).replace(/=+$/, '');
}

// みんなの状況の通行証が合っているか（合言葉 TEAM_PASS から作る記号と同じか）
function teamOk_(token) {
  var pass = webhook_('TEAM_PASS');
  return !!pass && String(token || '') === teamToken_(pass);
}

// ═══ 月末の締め：みんなの状況から、まだの人へ声をかける（日報の Discord チャンネルに出す） ═══
// 通行証が合うときだけ。続けて押しても 10分に1回まで。@メンションは鳴らさない（文に入っていても）
var REMIND_GAP_SEC = 600;
function remind_(p) {
  if (!teamOk_(p.token)) return { ok: false, expired: true, error: '合言葉を確かめられませんでした。みんなの状況に入り直してください' };
  var text = String(p.text || '').replace(/\r/g, '').trim().slice(0, 1800);
  if (!text) return { ok: false, error: '送る文がありません' };
  var url = webhook_('DISCORD_WEBHOOK');
  if (!url) return { ok: false, notSet: true, error: 'Discord の送り先（スクリプト プロパティ DISCORD_WEBHOOK）が設定されていません' };
  var cache = CacheService.getScriptCache();
  if (cache.get('remindAt')) return { ok: false, tooSoon: true, error: 'さっき送ったばかりです。10分ほどあけてから送ってください' };
  try { cache.put('remindAt', '1', REMIND_GAP_SEC); } catch (e0) {}   // 2人が同時に押しても1回だけ
  var r = postDiscord_(url, { username: 'Sieg 締めのお知らせ', content: text, allowed_mentions: { parse: [] } });
  if (!r.ok) {
    try { cache.remove('remindAt'); } catch (e1) {}
    return { ok: false, discord: false, code: r.code, error: 'Discord に送れませんでした（' + r.code + '）。少し待ってからもう一度押してください' };
  }
  return { ok: true, discord: true };
}

// ═══ 年末の「あなたの○年」（ポータルの wrapped.html 用） ═══
//   日報（入店日がその年）：日報を出した日数・はじめての日・現場（稼働場所）の数と多い順・アプローチと獲得の合計・いちばん獲得した日
//   獲得報告（獲得シート）：獲得の合計・多かった商品
//   AI のひとこと（1年をねぎらう）は「ふりかえり」シートに残し、7日たったら作り直す（12月中に数字が増えるため）
// 同じ日に日報を2回送っていたら、あとの方だけ数える
var REVIEW_SHEET = 'ふりかえり';
var REVIEW_HEAD = ['作った日時', '名前', '年', 'ひとこと'];
var REVIEW_MAX_LEN = 70;
var REVIEW_KEEP_MS = 7 * 86400000;
function yearReview_(name, year) {
  name = String(name || '').trim();
  if (!name) return { ok: false, error: '名前がありません' };
  if (!/^\d{4}$/.test(year)) year = todayYmd_().slice(0, 4);
  var sheet = getSheet_(), tz = sheet.getParent().getSpreadsheetTimeZone();
  var data = sheet.getDataRange().getValues(), byDay = {};
  for (var i = 1; i < data.length; i++) {
    var r = data[i];
    if (String(r[2] || '') !== name) continue;
    var d = ymd_(r[1], tz) || ymd_(r[0], tz);
    if (!d || d.slice(0, 4) !== year) continue;
    byDay[d] = { place: String(r[4] || '').replace(/\s+/g, ' ').trim(), ap: Number(r[5]) || 0, gain: Number(r[7]) || 0 };
  }
  var dates = Object.keys(byDay).sort(), places = {}, ap = 0, gain = 0, best = null;
  dates.forEach(function(d) {
    var x = byDay[d];
    if (x.place) places[x.place] = (places[x.place] || 0) + 1;
    ap += x.ap; gain += x.gain;
    if (x.gain > 0 && (!best || x.gain > best.gain)) best = { date: d, gain: x.gain };
  });
  var placeList = Object.keys(places).map(function(k) { return { name: k, days: places[k] }; })
    .sort(function(a, b) { return b.days - a.days || (a.name < b.name ? -1 : 1); });
  // 獲得報告
  var gs = gainSheet_(), gtz = gs.getParent().getSpreadsheetTimeZone(), gd = gs.getDataRange().getValues(), items = {}, gTotal = 0;
  for (var j = 1; j < gd.length; j++) {
    if (String(gd[j][2] || '') !== name || ymd_(gd[j][1], gtz).slice(0, 4) !== year) continue;
    var key = String(gd[j][3]) + '\t' + String(gd[j][4]), n = Number(gd[j][5]) || 0;
    items[key] = (items[key] || 0) + n; gTotal += n;
  }
  var top = Object.keys(items).map(function(k) { var a = k.split('\t'); return { carrier: a[0], label: a[1], count: items[k] }; })
    .filter(function(x) { return x.count > 0; }).sort(function(a, b) { return b.count - a.count; }).slice(0, 3);
  var out = { ok: true, name: name, year: year, days: dates.length, first: dates[0] || '', placeCount: placeList.length, places: placeList.slice(0, 5),
    approach: ap, gain: gain, bestDay: best, gains: { total: gTotal, top: top } };
  out.ai = (dates.length || gTotal) ? reviewAi_(name, year, out) : '';
  return out;
}

function reviewSheet_() {
  var ss = SpreadsheetApp.openById(SHEET_ID), sh = ss.getSheetByName(REVIEW_SHEET);
  if (!sh) { sh = ss.insertSheet(REVIEW_SHEET); sh.appendRow(REVIEW_HEAD); sh.setFrozenRows(1); }
  return sh;
}
var REVIEW_RULES =
  'あなたは営業チームのマスコット。メンバーの1年分の数字を見て、その本人に1年をねぎらう言葉を作る。\n' +
  'JSON で text を返す。60文字くらいまで、1〜2文。\n' +
  '- 下の数字から1つか2つを具体的に拾って、ほめる・ねぎらう。最後は来年への前向きなひとこと。\n' +
  '- ひらがな多めの、やさしくのんびりした話し言葉（です・ます は使わない）。\n' +
  '- 数字は下に書いてあるものだけを使う。計算し直したり、作ったりしない。ほかの人の名前や数字は出さない。\n' +
  '- 絵文字・顔文字・かぎかっこ・記号の飾りは使わない。';
function reviewAi_(name, year, s) {
  var sh = reviewSheet_(), last = sh.getLastRow();
  if (last >= 2) {
    var vals = sh.getRange(2, 1, last - 1, 4).getValues();
    for (var i = vals.length - 1; i >= 0; i--) {
      if (String(vals[i][1]) !== name || String(vals[i][2]) !== String(year)) continue;
      var at = vals[i][0];
      if (Object.prototype.toString.call(at) === '[object Date]' && Date.now() - at.getTime() < REVIEW_KEEP_MS) return String(vals[i][3] || '');
      break;
    }
  }
  if (!webhook_('GEMINI_API_KEY')) return '';
  var cache = CacheService.getScriptCache(), k = 'review:' + name + ':' + year;
  if (cache.get(k)) return '';   // 作っている途中・失敗したばかり
  try { cache.put(k, '1', 600); } catch (e0) {}
  var lines = ['名前：' + name, '【' + year + '年】',
    '日報を出した日：' + s.days + '日' + (s.first ? '（はじめては ' + s.first + '）' : ''),
    '行った現場：' + s.placeCount + 'か所' + (s.places.length ? '（多い順：' + s.places.slice(0, 3).map(function(x) { return x.name + ' ' + x.days + '日'; }).join('、') + '）' : ''),
    'アプローチ：' + s.approach + '件　日報の獲得：' + s.gain + '件',
    '獲得報告：' + s.gains.total + '件' + (s.gains.top.length ? '（多い商品：' + s.gains.top.map(function(x) { return x.label + ' ' + x.count + '件'; }).join('、') + '）' : '')];
  if (s.bestDay) lines.push('いちばん獲得した日：' + s.bestDay.date + '（' + s.bestDay.gain + '件）');
  var obj = gemini_(REVIEW_RULES + '\n\n' + lines.join('\n'), { type: 'OBJECT', properties: { text: { type: 'STRING' } }, required: ['text'] }, 300);
  var text = obj ? cleanText_(obj.text, REVIEW_MAX_LEN) : '';
  if (!text) return '';   // 失敗：10分は作り直さない（上の印がそのまま残る）
  try { cache.remove(k); } catch (e1) {}
  sh.appendRow([new Date(), name, String(year), text]);
  return text;
}
// Gemini に JSON で答えてもらう（ひとことと同じモデルの順）。だめなら null
function gemini_(prompt, schema, maxTokens) {
  var key = webhook_('GEMINI_API_KEY');
  if (!key) return null;
  for (var i = 0; i < CHEER_MODELS.length; i++) {
    var model = CHEER_MODELS[i];
    var payload = {
      contents: [{ role: 'user', parts: [{ text: prompt }] }],
      generationConfig: { temperature: 0.8, maxOutputTokens: maxTokens || 400, responseMimeType: 'application/json', responseSchema: schema,
        thinkingConfig: /gemini-2\.5/.test(model) ? { thinkingBudget: 0 } : { thinkingLevel: 'minimal' } }
    };
    try {
      var res = UrlFetchApp.fetch('https://generativelanguage.googleapis.com/v1beta/models/' + model + ':generateContent?key=' + encodeURIComponent(key), {
        method: 'post', contentType: 'application/json', payload: JSON.stringify(payload), muteHttpExceptions: true
      });
      if (res.getResponseCode() !== 200) { Logger.log('Gemini ' + model + ' HTTP ' + res.getResponseCode()); continue; }
      var json = JSON.parse(res.getContentText());
      var parts = json.candidates && json.candidates[0] && json.candidates[0].content && json.candidates[0].content.parts || [];
      var raw = parts.map(function(x) { return x.text || ''; }).join('');
      try { return JSON.parse(raw); } catch (e) { return { text: raw }; }
    } catch (e2) {
      Logger.log('Gemini ' + model + ' ' + e2.message);
    }
  }
  return null;
}
// ひとことと同じ整え方（長さだけ変えられる）
function cleanText_(s, max) {
  var t = String(s || '').split(/\r?\n/).map(function(x) { return x.trim(); }).filter(function(x) { return x; }).join('');
  t = unwrapQuotes_(t).replace(/[\uD800-\uDBFF][\uDC00-\uDFFF]/g, '').replace(/[☀-➿️]/g, '').trim();
  if (t.length > max) t = t.slice(0, max - 1) + '…';
  return t;
}
/** GASのエディタから実行して、ふりかえりがどう出るか試す（下の名前を変えて実行し、実行ログを見る） */
function testYearReview() {
  var r = yearReview_('鈴木', todayYmd_().slice(0, 4));
  Logger.log(JSON.stringify(r, null, 1));
}

// ═══ 不具合の記録 ═══
// ポータル・各フォームの gas-client.js が、Google とのやりとりに最後まで失敗したとき・画面のプログラムが止まったときに送ってくる。
// 「エラー記録」シートに1行ずつ残す（多すぎるときは10分に60件まで）。
// 10分に3件以上続いたら、管理用の Discord（スクリプト プロパティ DISCORD_ADMIN_WEBHOOK。無ければ送らない）に1時間に1回知らせる
var ERR_SHEET = 'エラー記録';
var ERR_HEAD = ['日時', '名前', '画面', 'していたこと', '種類', '案内', '詳しく', '端末'];
var ERR_MAX_PER_10MIN = 60;
function errSheet_() {
  var ss = SpreadsheetApp.openById(SHEET_ID), sh = ss.getSheetByName(ERR_SHEET);
  if (!sh) { sh = ss.insertSheet(ERR_SHEET); sh.appendRow(ERR_HEAD); sh.setFrozenRows(1); }
  return sh;
}
function logError_(p) {
  var cache = CacheService.getScriptCache(), n = Number(cache.get('errCount') || 0);
  if (n >= ERR_MAX_PER_10MIN) return { ok: true, dropped: true };
  try { cache.put('errCount', String(n + 1), 600); } catch (e0) {}
  var cut = function(v, k) {
    var t = String(v == null ? '' : v).replace(/[\r\n\t]+/g, ' ').slice(0, k);
    return /^[=+\-@]/.test(t) ? "'" + t : t;   // シートの式として読まれないように
  };
  // 名前は who で受け取る（name だと、更新前の GAS が日報として記録してしまうため）
  var row = [new Date(), cut(p.who, 20), cut(p.page, 80), cut(p.what, 60), cut(p.kind, 20), cut(p.msg, 200), cut(p.detail, 400), cut(p.ua, 120)];
  errSheet_().appendRow(row);
  var burst = Number(cache.get('errBurst') || 0) + 1;
  try { cache.put('errBurst', String(burst), 600); } catch (e1) {}
  var admin = webhook_('DISCORD_ADMIN_WEBHOOK');
  if (admin && burst >= 3 && !cache.get('errAlerted')) {
    try { cache.put('errAlerted', '1', 3600); } catch (e2) {}
    postDiscord_(admin, { username: 'Sieg 見守り', allowed_mentions: { parse: [] },
      content: 'ここ10分で ' + burst + '件、Google とのやりとりに失敗しています（最新：' + [row[1], row[2], row[5]].filter(String).join('・') + '）。「エラー記録」シートを見てください。' });
  }
  return { ok: true };
}
// 最近7日の不具合（新しい順に50件まで。みんなの状況の画面用。通行証が合うときだけ）
function errors_(p) {
  if (!teamOk_(p.token)) return { ok: false, expired: true, error: '合言葉を確かめられませんでした' };
  var sh = SpreadsheetApp.openById(SHEET_ID).getSheetByName(ERR_SHEET);
  if (!sh || sh.getLastRow() < 2) return { ok: true, items: [] };
  var last = sh.getLastRow(), from = Math.max(2, last - 299);
  var vals = sh.getRange(from, 1, last - from + 1, 8).getValues(), since = Date.now() - 7 * 86400000, items = [];
  for (var i = vals.length - 1; i >= 0 && items.length < 50; i--) {
    var at = vals[i][0];
    if (Object.prototype.toString.call(at) !== '[object Date]') continue;
    if (at.getTime() < since) break;
    items.push({ at: Utilities.formatDate(at, 'Asia/Tokyo', 'M/d HH:mm'), name: String(vals[i][1] || ''), page: String(vals[i][2] || ''),
      what: String(vals[i][3] || ''), kind: String(vals[i][4] || ''), msg: String(vals[i][5] || '') });
  }
  return { ok: true, items: items };
}

// ═══ 日報へのAIのひとこと（ポータルのぼのぼのがしゃべる） ═══
// その人のいちばん新しい日報（きょうかきのうの分）と、ここ2週間の流れを Gemini に読ませて、本人への
//   cheer … ほめる・ねぎらうひとこと
//   grow  … のびたところ（「できるようになったね」。前の週より本当に伸びたときだけ）
//   notice … 気づいたこと（同じ失敗のくり返し・書いていることと結果の食い違いの指摘）
//   undone … やれてないこと（前に「次は〜する」と書いたのに、きょうの所感に出てこないこと）
//   next   … 次はこれ（あしたやることを1つ。言い方の例つき）
//   todo   … きょうの所感に本人が書いた「次やること」（シートに残し、次の日報で できた／まだ を見る）
//   （前は hint 1つにまとめていた。古い画面のために hint にも3つをつなげて返す）
// アプローチ・着座の件数は材料にしない（所感の文章を中心に見る）
// を作る。伸びたかどうかの数字は GAS で計算してから渡す（AIに計算させない）。
// 日報を送るときには呼ばない（送信を遅くしないため）。ポータルのホームを開いたときに1回だけ作り、
// 「ひとこと」シートに残して、2回目からはそれを返す。
// 鍵はスクリプト プロパティ GEMINI_API_KEY（経費精算と同じ値）。無いとき・失敗したときは text: ''（ぼのぼのは数字のひとことだけ）。
var CHEER_SHEET = 'ひとこと';
var CHEER_HEAD = ['作った日時', '名前', '日報のキー', '日付', 'ひとこと', 'のびた', 'ヒント', '気づいたこと', 'やれてないこと', '次はこれ', '次やること', '前回の次やること', '前回の判定'];
var CHEER_COLS = CHEER_HEAD.length;
var CHEER_MAX_LEN = 60;   // ひとこと（AIには40文字以内と頼み、これは念のための上限）
var GROW_MAX_LEN = 80;    // のびた（AIには50文字以内と頼む）
var NOTICE_MAX_LEN = 110;  // 気づいたこと・やれてないこと（AIには70文字以内と頼む）
var NEXT_MAX_LEN = 140;    // 次はこれ（AIには90文字以内と頼む）
var TODO_MAX_LEN = 50;     // 次やること（AIには30文字以内と頼む）
var PREV_RESULTS = ['できた', 'まだ', 'わからない'];
var HINT_MAX_LEN = 200;   // アドバイスは過去の日報と照らし合わせて書くので、長めでよい（AIには120文字までと頼み、これは念のための上限）
var CHEER_MODELS = ['gemini-2.5-flash', 'gemini-flash-lite-latest', 'gemini-3.5-flash'];   // 経費精算の領収書読み取りと同じ順

function cheerSheet_() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sh = ss.getSheetByName(CHEER_SHEET);
  if (!sh) { sh = ss.insertSheet(CHEER_SHEET); sh.appendRow(CHEER_HEAD); sh.setFrozenRows(1); }
  else if (!String(sh.getRange(1, 8).getValue() || '').trim()) sh.getRange(1, 6, 1, CHEER_COLS - 5).setValues([CHEER_HEAD.slice(5)]);   // 後から足した列
  return sh;
}

// 内訳 "キャッチ6 / 店内2" → { キャッチ: 6, 店内: 2 }
function parseBreakdown_(s) {
  var out = {};
  String(s || '').split(' / ').forEach(function(part) {
    var m = part.trim().match(/^(.*\D)(\d+)$/);
    if (m) out[m[1]] = (out[m[1]] || 0) + Number(m[2]);
  });
  return out;
}

// その人の日報（新しい順）。[{ key, date, place, carrier, ap, apText, aps, gain, gainText, target, cum, note }]
function recentReports_(name, n) {
  var sheet = getSheet_(), tz = sheet.getParent().getSpreadsheetTimeZone();
  var data = sheet.getDataRange().getValues(), out = [];
  for (var i = data.length - 1; i >= 1 && out.length < n; i--) {
    var r = data[i];
    if (String(r[2] || '') !== name) continue;
    var at = r[0];
    out.push({
      key: name + '|' + (r[13] || (Object.prototype.toString.call(at) === '[object Date]' ? at.getTime() : String(at))),
      date: ymd_(r[1], tz) || ymd_(at, tz), carrier: String(r[3] || ''), place: String(r[4] || ''),
      ap: Number(r[5]) || 0, apText: String(r[6] || ''), aps: parseBreakdown_(r[6]),
      gain: Number(r[7]) || 0, gainText: String(r[8] || ''),
      target: Number(r[9]) || 0, cum: Number(r[10]) || 0, note: String(r[12] || '')
    });
  }
  return out;
}

// 直近7日（きょうまで）と、その前の7日の合計。同じ日に2回出していたら新しい方だけ数える
function twoWeeks_(reports, today) {
  var t = new Date(today + 'T00:00:00+09:00').getTime(), seen = {};
  var w = [{ n: 0, ap: 0, gain: 0, aps: {} }, { n: 0, ap: 0, gain: 0, aps: {} }];
  reports.forEach(function(r) {
    if (!r.date || seen[r.date]) return;
    seen[r.date] = true;
    var days = Math.round((t - new Date(r.date + 'T00:00:00+09:00').getTime()) / 86400000);
    var k = days >= 0 && days < 7 ? 0 : days >= 7 && days < 14 ? 1 : -1;
    if (k < 0) return;
    w[k].n++; w[k].ap += r.ap; w[k].gain += r.gain;
    Object.keys(r.aps).forEach(function(l) { w[k].aps[l] = (w[k].aps[l] || 0) + r.aps[l]; });
  });
  return w;
}
// アプローチ・着座の件数は AI に渡さない（材料にしない）
function weekLine_(w) {
  if (!w.n) return '日報なし';
  return '日報' + w.n + '回　獲得' + w.gain + '件';
}

function cheer_(name) {
  name = String(name || '').trim();
  if (!name) return { ok: false, error: '名前がありません' };
  var reports = recentReports_(name, 30);
  var latest = reports[0];
  if (!latest) return { ok: true, text: '' };
  var today = todayYmd_(), yest = Utilities.formatDate(new Date(Date.now() - 86400000), 'Asia/Tokyo', 'yyyy-MM-dd');
  if (latest.date !== today && latest.date !== yest) return { ok: true, text: '' };   // 古い日報にはしゃべらない
  if (!webhook_('GEMINI_API_KEY')) return { ok: true, text: '' };   // 鍵が無い（失敗ではないので、覚えておかない）

  var sh = cheerSheet_();
  var found = findCheer_(sh, latest.key);
  if (found) return cheerOut_(found, latest.date);

  // 同じ日報を何台かで同時に開いても、AIを呼ぶのは1回だけ（作っている間は印を置く）。
  // 日報の送信と同じロックは使わない（AIが遅いときに、日報の送信を待たせないため）
  var cache = CacheService.getScriptCache(), failKey = 'cheerFail:' + latest.key, busyKey = 'cheerBusy:' + latest.key;
  if (cache.get(failKey) || cache.get(busyKey)) return { ok: true, text: '' };   // 失敗したばかり・作っている途中
  try { cache.put(busyKey, '1', 60); } catch (e0) {}
  var c = null;
  try { c = makeCheer_(name, reports, today, todosOf_(sh, reports)); } catch (e) { Logger.log('ひとことを作れませんでした: ' + e.message); }
  try { cache.remove(busyKey); } catch (e1) {}
  if (!c || !c.text) { try { cache.put(failKey, '1', 600); } catch (e2) {} return { ok: true, text: '' }; }
  sh.appendRow([new Date(), name, latest.key, latest.date, c.text, c.grow, c.hint, c.notice, c.undone, c.next, c.todo, c.prev.todo, c.prev.todo ? c.prev.result : '']);
  return cheerOut_(c, latest.date);
}
function cheerOut_(c, date) {
  return { ok: true, text: c.text, grow: c.grow, hint: c.hint, notice: c.notice || '', undone: c.undone || '', next: c.next || '',
    todo: c.todo || '', prev: c.prev && c.prev.todo ? { todo: c.prev.todo, result: c.prev.result || 'わからない' } : null, date: date };
}
// これまでの日報で残した「次やること」（日報のキー → 次やること）
function todosOf_(sh, reports) {
  var last = sh.getLastRow(), want = {}, out = {};
  reports.forEach(function(r) { want[r.key] = true; });
  if (last < 2) return out;
  sh.getRange(2, 3, last - 1, CHEER_COLS - 2).getValues().forEach(function(v) {
    if (want[String(v[0])] && String(v[8] || '').trim()) out[String(v[0])] = String(v[8]).trim();
  });
  return out;
}

function findCheer_(sh, key) {
  var last = sh.getLastRow();
  if (last < 2) return null;
  var vals = sh.getRange(2, 3, last - 1, CHEER_COLS - 2).getValues();   // C:キー D:日付 E:ひとこと F:のびた G:ヒント H〜M:気づいたこと…前回の判定
  for (var i = vals.length - 1; i >= 0; i--) {
    var v = vals[i];
    if (String(v[0]) === key) return { text: String(v[2] || ''), grow: String(v[3] || ''), hint: String(v[4] || ''),
      notice: String(v[5] || ''), undone: String(v[6] || ''), next: String(v[7] || ''), todo: String(v[8] || ''),
      prev: { todo: String(v[9] || ''), result: String(v[10] || '') } };
  }
  return null;
}

// AIへのお願い（ここを書き換えると、しゃべり方が変わる）
var CHEER_RULES =
  'あなたは営業チームのマスコット。メンバーが送った日報（とくに本人が書いた所感の文章）を、その人の過去の日報と照らし合わせて読み、本人にかける言葉を作る。\n' +
  'アプローチ・着座の件数は材料にしない。所感の中身（やったこと・うまくいったこと・困ったこと・次にやると書いたこと）を中心に考える。\n' +
  'JSON で cheer・grow・notice・undone・next・todo・prevTodo・prevResult を返す。どれも1行。\n' +
  '- cheer：きょうの所感から1つ具体的に拾って、ほめる・ねぎらう（40文字以内。文の終わりまで書き切る）。\n' +
  '- grow：過去の所感とくらべて、できるようになったこと・変わったことを書く（50文字以内。文の終わりまで書き切る）。前に「次は〜する」と書いていて、きょうそれができていれば、それを拾う。なければ空文字。\n' +
  '- アドバイスは3つに分けて書く。どれも言うことがなければ空文字。ふわっとした励ましにしない。人を責める・ばかにする・怒る言い方はしない。\n' +
  '- notice（気づいたこと）：きょうの所感と過去の所感を照らし合わせて、同じ悩み・同じ失敗のくり返しや、書いていることと結果の食い違いを指摘する（70文字以内、1文）。\n' +
  '  - 根拠を書く。過去の日報を引くときは「11/22に」のように日付を書き、本人の文をそのまま「」でくくって引用する（言い換えた文を「」でくくらない）。\n' +
  '  - 食い違いや考え違いは、はっきり指摘・否定してよい。\n' +
  '- undone（やれてないこと）：これまでの日報で本人が「次は〜する」「明日は〜する」と書いたのに、きょうの所感に出てこない・できていないことを1つ（70文字以内、1文）。いつ書いたことか日付を入れる。なければ空文字。\n' +
  '- next（次はこれ）：あしたの現場でやることを1つだけ、行動で書く。お客さまへの言い方の例を「」で1つ入れる（90文字以内、1〜2文）。所感が短くて材料がないときは、獲得件数や目標までのペースから考える。\n' +
  '- todo：きょうの所感に本人が書いた「次は〜する」「明日は〜する」を、本人の言葉に近い形で短く（30文字以内）。書いていなければ空文字。\n' +
  '- prevTodo と prevResult：下に【前回の次やること】があればそれを、なければ1つ前の日報の所感から本人が次にやると書いたことを prevTodo に（なければ空文字）。' +
  'きょうの所感を読んで、それができたなら prevResult を「できた」、できていない・やっていないなら「まだ」、所感から分からないなら「わからない」にする。\n' +
  '共通のルール：\n' +
  '- やさしい話し言葉（です・ます は使わない）。cheer と grow はひらがな多め、notice・undone・next は読みやすさ優先で漢字を使ってよい。\n' +
  '- 数字は下に書いてあるものだけを使う。計算し直したり、作ったりしない。\n' +
  '- 日報に書いてないことを想像で足さない。\n' +
  '- ほかのメンバーの名前や数字は出さない。\n' +
  '- 絵文字・顔文字・記号の飾りは使わない。かぎかっこ「」は、日報の文の引用と、お客さまへの言い方の例にだけ使う。';

var CHEER_SCHEMA = { type: 'OBJECT', properties: { cheer: { type: 'STRING' }, grow: { type: 'STRING' }, notice: { type: 'STRING' }, undone: { type: 'STRING' },
  next: { type: 'STRING' }, todo: { type: 'STRING' }, prevTodo: { type: 'STRING' }, prevResult: { type: 'STRING' } },
  required: ['cheer', 'grow', 'notice', 'undone', 'next', 'todo', 'prevTodo', 'prevResult'] };

function cheerPrompt_(name, reports, today, todos) {
  todos = todos || {};
  var r = reports[0], lines = [];
  var note = function(x, n) { var t = String(x.note || '').replace(/\s+/g, ' ').trim(); return t ? t.slice(0, n) : '（なし）'; };
  lines.push('名前：' + name);
  lines.push('【きょうの日報】' + r.date + '　場所：' + (r.place || '-') + '　キャリア：' + (r.carrier || '-') + '　獲得：' + r.gain + '件' + (r.gainText ? '（' + r.gainText + '）' : ''));
  lines.push('所感：' + note(r, 500));
  if (reports.length > 1) {
    lines.push('【これまでの日報（新しい順）】');
    reports.slice(1, 11).forEach(function(x) {
      lines.push('- ' + x.date + '　' + (x.place || '-') + '　獲得' + x.gain + '件　所感：' + note(x, 250) + (todos[x.key] ? '　（次やること：' + todos[x.key] + '）' : ''));
    });
  }
  if (reports[1] && todos[reports[1].key]) lines.push('【前回の次やること】' + reports[1].date + '：' + todos[reports[1].key]);
  var w = twoWeeks_(reports, today || r.date);
  lines.push('【直近7日】' + weekLine_(w[0]));
  lines.push('【その前の7日】' + weekLine_(w[1]));
  if (r.target) {
    var d = new Date(today + 'T00:00:00+09:00'), last = new Date(d.getFullYear(), d.getMonth() + 1, 0).getDate();
    var left = last - d.getDate(), rest = Math.max(0, r.target - r.cum);
    lines.push('【今月のMNP】' + r.cum + '件 / 目標 ' + r.target + '件（月末まで あと' + left + '日）');
    // ペースは GAS で計算して渡す（AI に計算させない）
    if (rest > 0) lines.push('【目標までのペース】あと' + rest + '件。' + (left > 0 ? '残り' + left + '日で 1日あたり' + (Math.round(rest / left * 10) / 10) + '件' : 'きょうが月末'));
    else lines.push('【目標までのペース】目標たっせい');
  }
  return CHEER_RULES + '\n\n' + lines.join('\n');
}

function makeCheer_(name, reports, today, todos) {
  var key = webhook_('GEMINI_API_KEY');
  if (!key) return null;
  todos = todos || {};
  var prompt = cheerPrompt_(name, reports, today || todayYmd_(), todos);
  for (var i = 0; i < CHEER_MODELS.length; i++) {
    var model = CHEER_MODELS[i];
    var payload = {
      contents: [{ role: 'user', parts: [{ text: prompt }] }],
      generationConfig: { temperature: 0.8, maxOutputTokens: 800, responseMimeType: 'application/json', responseSchema: CHEER_SCHEMA,
        thinkingConfig: /gemini-2\.5/.test(model) ? { thinkingBudget: 0 } : { thinkingLevel: 'minimal' } }
    };
    try {
      var res = UrlFetchApp.fetch('https://generativelanguage.googleapis.com/v1beta/models/' + model + ':generateContent?key=' + encodeURIComponent(key), {
        method: 'post', contentType: 'application/json', payload: JSON.stringify(payload), muteHttpExceptions: true
      });
      if (res.getResponseCode() !== 200) { Logger.log('Gemini ' + model + ' HTTP ' + res.getResponseCode()); continue; }
      var json = JSON.parse(res.getContentText());
      var parts = json.candidates && json.candidates[0] && json.candidates[0].content && json.candidates[0].content.parts || [];
      var raw = parts.map(function(p) { return p.text || ''; }).join('');
      var obj = null;
      try { obj = JSON.parse(raw); } catch (e) { obj = { cheer: raw }; }   // JSON で返らなかったら、全体をひとことに
      var out = adviceFrom_(obj, reports, todos);
      if (out.text) return out;
    } catch (e) {
      Logger.log('Gemini ' + model + ' ' + e.message);
    }
  }
  return null;
}

// AIの答え → 画面に出す形。日報の引用は本当にその日報に書いてあるか確かめる（違えばその文を出さない）
function adviceFrom_(obj, reports, todos) {
  var notes = reports.map(function(r) { return { date: r.date, note: r.note }; });
  var out = {
    text: checkQuotes_(cleanCheer_(obj.cheer), notes),
    grow: checkQuotes_(cleanCheer_(obj.grow, GROW_MAX_LEN), notes),
    notice: checkQuotes_(cleanCheer_(obj.notice, NOTICE_MAX_LEN), notes),
    undone: checkQuotes_(cleanCheer_(obj.undone, NOTICE_MAX_LEN), notes),
    next: cleanCheer_(obj.next, NEXT_MAX_LEN),   // 「」はお客さまへの言い方の例なので確かめない
    todo: cleanCheer_(obj.todo, TODO_MAX_LEN)
  };
  if (!out.text && obj.cheer) out.text = cleanCheer_(String(obj.cheer).replace(/「[^「」]*」/g, ''));   // ほめる文は、引用を外して残す
  // 前回の次やること：残してあればそれ、なければ AI が前の所感から取り出したもの
  var prevKey = reports[1] && reports[1].key, prevTodo = (prevKey && todos[prevKey]) || cleanCheer_(obj.prevTodo, TODO_MAX_LEN);
  var res = String(obj.prevResult || '').trim();
  out.prev = { todo: prevTodo, result: prevTodo ? (PREV_RESULTS.indexOf(res) >= 0 ? res : 'わからない') : '' };
  out.hint = [out.notice, out.undone, out.next].filter(String).join(' ');   // 古い画面用
  if (out.hint.length > HINT_MAX_LEN) out.hint = cutAtSentence_(out.hint, HINT_MAX_LEN);
  return out;
}
// 引用をくらべる形（空白・句読点・記号をのぞき、全角英数を半角に）
function quoteKey_(v) {
  return String(v || '').replace(/[Ａ-Ｚａ-ｚ０-９]/g, function(c) { return String.fromCharCode(c.charCodeAt(0) - 0xFEE0); })
    .replace(/[\s　、。,.，．！？!?・…「」『』（）()〜~ー－\-]/g, '').toLowerCase();
}
// 文の中の「」が日報の文か確かめる。日付（11/22 など）があればその日の日報で、なければどれかの日報で。
// 見つからない引用をふくむ文は外す。短い引用（4文字未満）は確かめない。「…」で省いた引用は、残りの部分で確かめる
function checkQuotes_(text, notes) {
  if (!text || text.indexOf('「') < 0) return text;
  var sentences = text.match(/[^。！？!?]+[。！？!?]?/g) || [text];
  var kept = sentences.filter(function(sen) {
    var quotes = sen.match(/「[^「」]+」/g);
    if (!quotes) return true;
    var dates = (sen.match(/(\d{1,2})\/(\d{1,2})/g) || []).map(function(d) { var p = d.split('/'); return ('0' + p[0]).slice(-2) + '-' + ('0' + p[1]).slice(-2); });
    var pool = notes.filter(function(n) { return !dates.length || dates.some(function(d) { return String(n.date).slice(5) === d; }); })
      .map(function(n) { return quoteKey_(n.note); }).join('|');
    return quotes.every(function(q) {
      var pieces = q.slice(1, -1).split(/…|\.\.\.|・・・/).map(quoteKey_).filter(function(x) { return x.length >= 4; });
      return pieces.every(function(x) { return pool.indexOf(x) >= 0; });
    });
  });
  return kept.join('').trim();
}

// 1行目だけ・かぎかっこや引用符・絵文字をのぞいて、長すぎたら切る
function cleanCheer_(s, max) {
  max = max || CHEER_MAX_LEN;
  var t = String(s || '').split(/\r?\n/).map(function(x) { return x.trim(); }).filter(function(x) { return x; })[0] || '';
  t = unwrapQuotes_(t)
       .replace(/[\uD800-\uDBFF][\uDC00-\uDFFF]/g, '')      // 絵文字（サロゲートペア）
       .replace(/[☀-➿️]/g, '')               // 記号の絵文字
       .trim();
  if (t.length > max) t = cutAtSentence_(t, max);
  return t;
}
// 全体がかぎかっこ・引用符でくくられているときだけ外す（「…」と11/22に書いて…の頭の「を消さない）
function unwrapQuotes_(t) {
  var pairs = [['「', '」'], ['『', '』'], ['"', '"'], ["'", "'"], ['“', '”']];
  for (var i = 0; i < pairs.length; i++) {
    var o = pairs[i][0], c = pairs[i][1];
    if (t.length >= 2 && t.charAt(0) === o && t.charAt(t.length - 1) === c && t.slice(1, -1).indexOf(o) < 0 && t.slice(1, -1).indexOf(c) < 0) return t.slice(1, -1).trim();
  }
  return t;
}
// 長すぎるときは、上限の中の最後の「。」「！」で切る（文の途中で「…」にしない）。文の区切りが前のほうにしかなければ「…」
function cutAtSentence_(t, max) {
  var head = t.slice(0, max), m = head.match(/^[\s\S]*[。！？!?]/);
  if (m && m[0].length >= max * 0.5) return m[0];
  return t.slice(0, max - 1) + '…';
}

/**
 * GASのエディタから実行して、ひとことがどう作られるか試す（シートには残さない）。
 * 下の名前を、日報を出している人に変えてから実行し、実行ログを見る。
 */
function testCheer() {
  var name = '鈴木';
  var reports = recentReports_(name, 30);
  if (!reports.length) { Logger.log(name + 'さんの日報がありません'); return; }
  Logger.log('鍵：' + (webhook_('GEMINI_API_KEY') ? 'あり' : 'なし（スクリプト プロパティ GEMINI_API_KEY を登録してください）'));
  var c = makeCheer_(name, reports, todayYmd_());
  if (!c) { Logger.log('（作れませんでした）'); return; }
  Logger.log('ほめる：' + c.text);
  Logger.log('のびた：' + (c.grow || '（なし）'));
  Logger.log('気づいたこと：' + (c.notice || '（なし）'));
  Logger.log('やれてないこと：' + (c.undone || '（なし）'));
  Logger.log('次はこれ：' + (c.next || '（なし）'));
  Logger.log('次やること：' + (c.todo || '（なし）'));
  Logger.log('前回の次やること：' + (c.prev.todo ? c.prev.todo + '（' + c.prev.result + '）' : '（なし）'));
}

// スクリプト プロパティの Webhook URL（前後の空白・改行は取る）
function webhook_(key) {
  return String(PropertiesService.getScriptProperties().getProperty(key) || '').trim();
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

// ═══ 名簿（ポータルと各ツールに出す人） ═══
// 「名簿」シートに1人1行。みんなの状況（team.html）の「メンバー」から、通行証のある人だけが追加・直す・外すをする。
// 各画面は ?action=roster で読む（sieg-roster.js）。読めないときは、画面に入っている前の名簿で動く。
//   区分 … 社員／パートナー。日報・シフト入力・掲示板は区分で決まる
//   稼働費・経費精算・みんなの状況 … その人が使うなら ○
//   外した日 … 入っていれば一覧に出さない（記録は消さない。もう一度「戻す」もできる）
// 名前はシートやスプレッドシートのあちこちで目印になっているので、あとから変えられない（外して、新しく追加する）。
var ROSTER_SHEET = '名簿';
var ROSTER_HEAD = ['名前', '区分', 'フルネーム', 'キャリア', '稼働費', '経費精算', 'みんなの状況', '外した日', '直した日時'];
var ROSTER_TOOLS = ['commute', 'keihi', 'team'];
var ROSTER_CARRIERS = ['docomo', 'au', 'Softbank'];
var ROSTER_CACHE = 'roster:v1';
// シートがまだ無いときの名簿（いままで sieg-roster.js に書いていた内容）
var ROSTER_SEED = [
  { name: '福山', group: 'staff', full: '福山康平', carrier: 'Softbank', tools: ['commute', 'keihi', 'team'] },
  { name: '古川', group: 'staff', full: '古川泰成', carrier: 'Softbank', tools: ['commute', 'keihi'] },
  { name: '清田', group: 'staff', full: '清田和男', carrier: 'docomo', tools: ['commute', 'keihi'] },
  { name: '福岡', group: 'staff', full: '', carrier: 'docomo', tools: ['team'] },
  { name: '山口', group: 'partner', full: '山口真人', carrier: 'au', tools: ['commute'] },
  { name: '鈴木', group: 'partner', full: '', carrier: 'docomo', tools: ['commute'] },
  { name: '毛利', group: 'partner', full: '', carrier: 'Softbank', tools: ['commute'] },
  { name: '山内', group: 'partner', full: '', carrier: '', tools: ['commute'] }
];
function rosterRow_(x) {
  var t = x.tools || [];
  return [x.name, x.group === 'partner' ? 'パートナー' : '社員', x.full || '', x.carrier || '',
    t.indexOf('commute') >= 0 ? '○' : '', t.indexOf('keihi') >= 0 ? '○' : '', t.indexOf('team') >= 0 ? '○' : '',
    x.hidden || '', x.at || ''];
}
function rosterFromRow_(r, tz) {
  var hid = r[7];
  if (hid instanceof Date) hid = Utilities.formatDate(hid, tz, 'yyyy-MM-dd');
  var tools = [];
  if (String(r[4]).trim()) tools.push('commute');
  if (String(r[5]).trim()) tools.push('keihi');
  if (String(r[6]).trim()) tools.push('team');
  var c = String(r[3] || '').trim(), ck = c.toLowerCase();
  return {
    name: String(r[0] || '').trim(),
    group: String(r[1] || '').indexOf('パートナー') >= 0 ? 'partner' : 'staff',
    full: String(r[2] || '').trim(),
    carrier: ck === 'docomo' ? 'docomo' : ck === 'au' ? 'au' : ck === 'softbank' ? 'Softbank' : '',
    tools: tools,
    hidden: String(hid || '').trim()
  };
}
function rosterSheet_(create) {
  var ss = SpreadsheetApp.openById(SHEET_ID), sh = ss.getSheetByName(ROSTER_SHEET);
  if (!sh && create) {
    sh = ss.insertSheet(ROSTER_SHEET);
    sh.appendRow(ROSTER_HEAD);
    ROSTER_SEED.forEach(function(x) { sh.appendRow(rosterRow_(x)); });
    sh.setFrozenRows(1);
  }
  return sh;
}
function readRoster_() {
  var sh = rosterSheet_(false);
  if (!sh || sh.getLastRow() < 2) return { people: ROSTER_SEED.map(function(x) { var o = JSON.parse(JSON.stringify(x)); o.hidden = ''; return o; }), seed: true };
  var tz = sh.getParent().getSpreadsheetTimeZone();
  var people = sh.getRange(2, 1, sh.getLastRow() - 1, ROSTER_HEAD.length).getValues()
    .map(function(r) { return rosterFromRow_(r, tz); })
    .filter(function(x) { return x.name; });
  return { people: people };
}
function roster_() {
  var cache = CacheService.getScriptCache(), hit = cache.get(ROSTER_CACHE);
  if (hit) { try { return JSON.parse(hit); } catch (e) {} }
  var r = readRoster_(), m = readMaster_();
  var out = { ok: true, people: r.people, seed: !!r.seed, clients: m.clients, places: m.places, masterSeed: !!m.seed };
  try { cache.put(ROSTER_CACHE, JSON.stringify(out), 600); } catch (e) {}   // シートを直接直したときも10分で出る
  return out;
}
// 名前に使えない文字：空白・区切り記号・式の始まり（= + - @）
function rosterName_(v) {
  var n = String(v || '').replace(/[\s　]+/g, '');
  if (!n) return { error: '名前を入れてください' };
  if (/^[=+\-@]/.test(n) || /[,，、\/\\<>"'`()（）]/.test(n)) return { error: '名前に使えない文字があります' };
  if (n.length > 10) return { error: '名前は10文字までにしてください' };
  return { name: n };
}
function rosterSave_(p) {
  if (!teamOk_(p.token)) return { ok: false, expired: true, error: '合言葉を確かめられませんでした。みんなの状況に入り直してください' };
  var op = String(p.op || ''), x = p.person || {};
  if (['add', 'update', 'hide', 'show'].indexOf(op) < 0) return { ok: false, error: 'できない操作です' };
  var nm = rosterName_(x.name);
  if (nm.error) return { ok: false, error: nm.error };
  var sh = rosterSheet_(true), tz = sh.getParent().getSpreadsheetTimeZone();
  var n = sh.getLastRow() - 1;
  var rows = n > 0 ? sh.getRange(2, 1, n, ROSTER_HEAD.length).getValues() : [];
  var people = rows.map(function(r) { return rosterFromRow_(r, tz); });
  var idx = -1;
  people.forEach(function(q, i) { if (q.name === nm.name) idx = i; });
  var now = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd HH:mm');
  var clean = {
    name: nm.name,
    group: x.group === 'partner' ? 'partner' : 'staff',
    full: String(x.full || '').replace(/[\r\n\t]/g, '').trim().replace(/^[=+\-@]+/, '').slice(0, 20),
    carrier: ROSTER_CARRIERS.indexOf(x.carrier) >= 0 ? x.carrier : '',
    tools: (Array.isArray(x.tools) ? x.tools : []).filter(function(t, i, a) { return ROSTER_TOOLS.indexOf(t) >= 0 && a.indexOf(t) === i; }),
    hidden: '', at: now
  };
  if (op === 'add') {
    if (idx >= 0) {
      var same = people[idx];
      // 通信が切れて送り直したとき（同じ内容）は、追加できたことにする
      if (!same.hidden && same.group === clean.group && same.full === clean.full && same.carrier === clean.carrier && same.tools.join() === clean.tools.join()) return rosterOk_(true);
      return { ok: false, error: '「' + nm.name + '」はもう名簿にいます' + (same.hidden ? '（外した人の中）。外した人から「戻す」を押してください' : '') };
    }
    people.push(clean); idx = people.length - 1;
  } else {
    if (idx < 0) return { ok: false, error: '「' + nm.name + '」が名簿に見つかりません。開き直してください' };
    var cur = people[idx];
    if (op === 'update') { clean.hidden = cur.hidden; people[idx] = clean; }
    if (op === 'hide') { cur.hidden = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd'); cur.at = now; }
    if (op === 'show') { cur.hidden = ''; cur.at = now; }
  }
  // みんなの状況を見る人が1人もいなくなると、だれも名簿を直せなくなる
  if (!people.some(function(q) { return !q.hidden && q.tools.indexOf('team') >= 0; })) {
    return { ok: false, error: 'みんなの状況を見る人が1人もいなくなります。ほかの人に「みんなの状況」を付けてからにしてください' };
  }
  sh.getRange(idx + 2, 1, 1, ROSTER_HEAD.length).setValues([rosterRow_(people[idx])]);
  return rosterOk_(false);
}
// ═══ 取引先・場所の一覧（ポータルと各ツールが使う。名簿と一緒に ?action=roster で配る） ═══
// 「取引先」シート：シフトの取引先のボタン・シフト表のプルダウンと集計・みんなの状況の色
// 「場所」シート：シフトの開催場所 → 住所（稼働費の距離）・店舗稼働・担当店舗（Discord用の文のカッコ）
// お店の住所だけなので、合言葉なしで配ってよい（自宅は入れない）。直すのはみんなの状況（通行証）から。
var CLIENT_SHEET = '取引先';
var CLIENT_HEAD = ['名前', 'キャリア', 'いつもの場所', '色', '外した日', '直した日時'];
var CLIENT_SEED = [
  { name: '南国殖産', carrier: 'docomo', place: '近見', color: '#2E7D32' },
  { name: 'ハイコム', carrier: 'SoftBank', place: '', color: '#E65100' },
  { name: 'サンコミ', carrier: 'au', place: '', color: '#6A1B9A' },
  { name: 'ダル', carrier: 'docomo', place: '', color: '#00838F' },
  { name: 'NOON', carrier: '', place: '', color: '#AD1457' },
  { name: '九情', carrier: 'au', place: '', color: '#1565C0' }
];
var PLACE_SHEET = '場所';
var PLACE_HEAD = ['名前', '別名', '住所', '店舗稼働', '担当店舗', 'カッコなし', '外した日', '直した日時'];
// シートがまだ無いときの場所（住所は「稼働費の場所から取り込む」で入る）
var PLACE_SEED = [
  { name: '下通', noParen: true }, { name: 'ゆめタウン大牟田', noParen: true },
  { name: '健軍', shop: true }, { name: '松橋', shop: true }, { name: '光の森', shop: true }, { name: '大津', shop: true },
  { name: '麻生田', shop: true }, { name: '八代', shop: true, alias: ['八代中央'] }, { name: '玉名', shop: true, alias: ['玉名中央'] },
  { name: '熊本西', shop: true }, { name: '東バイパス', shop: true }, { name: '菊池', shop: true },
  { name: '上熊本', shop: true, alias: ['上熊'] }, { name: '近見', shop: true }
];
var MASTER_CARRIERS = ['docomo', 'au', 'SoftBank'];
var MASTER_PREFIX = /^(?:ＳＢ|SB|ソフトバンク|ドコモ|docomo|ａｕ|au)[\s　・･]*/i;

// 場所の名前をくらべる形（空白・全角英数・頭の会社名をのぞく）。「sb 玉名」「SB玉名」「玉名」は同じ
function placeKey_(v) {
  var t = String(v || '').replace(/[Ａ-Ｚａ-ｚ０-９]/g, function(c) { return String.fromCharCode(c.charCodeAt(0) - 0xFEE0); }).replace(/[\s　]+/g, '');
  for (var i = 0; i < 3 && MASTER_PREFIX.test(t); i++) { var t2 = t.replace(MASTER_PREFIX, ''); if (!t2 || t2 === t) break; t = t2; }
  return t.toLowerCase();
}
function hasPrefix_(v) { return MASTER_PREFIX.test(String(v || '').replace(/[\s　]+/g, '')); }
function stripPrefix_(v) {
  var t = String(v || '').trim();
  for (var i = 0; i < 3 && MASTER_PREFIX.test(t); i++) { var t2 = t.replace(MASTER_PREFIX, '').trim(); if (!t2 || t2 === t) break; t = t2; }
  return t;
}
function masterSheet_(name, create) {
  var ss = SpreadsheetApp.openById(SHEET_ID), sh = ss.getSheetByName(name);
  if (!sh && create) {
    sh = ss.insertSheet(name);
    if (name === CLIENT_SHEET) { sh.appendRow(CLIENT_HEAD); CLIENT_SEED.forEach(function(c) { sh.appendRow(clientRow_(c)); }); }
    else { sh.appendRow(PLACE_HEAD); PLACE_SEED.forEach(function(p) { sh.appendRow(placeRow_(p)); }); }
    if (sh.setFrozenRows) sh.setFrozenRows(1);
  }
  return sh;
}
function clientRow_(c) { return [c.name, c.carrier || '', c.place || '', c.color || '', c.hidden || '', c.at || '']; }
function placeRow_(p) { return [p.name, (p.alias || []).join('、'), p.address || '', p.shop ? '○' : '', p.store || '', p.noParen ? '○' : '', p.hidden || '', p.at || '']; }
function dateText__(v, tz) { return Object.prototype.toString.call(v) === '[object Date]' ? Utilities.formatDate(v, tz, 'yyyy-MM-dd') : String(v || '').trim(); }
function readMaster_() {
  var tz = 'Asia/Tokyo', seed = false;
  var cs = masterSheet_(CLIENT_SHEET, false), ps = masterSheet_(PLACE_SHEET, false);
  var clients = cs && cs.getLastRow() >= 2
    ? cs.getRange(2, 1, cs.getLastRow() - 1, CLIENT_HEAD.length).getValues().map(function(r) {
        var col = String(r[3] || '').trim();
        return { name: String(r[0] || '').trim(), carrier: MASTER_CARRIERS.indexOf(String(r[1]).trim()) >= 0 ? String(r[1]).trim() : '',
          place: String(r[2] || '').trim(), color: /^#[0-9a-fA-F]{6}$/.test(col) ? col : '', hidden: dateText__(r[4], tz) };
      }).filter(function(c) { return c.name; })
    : (seed = true, CLIENT_SEED.map(function(c) { return { name: c.name, carrier: c.carrier, place: c.place, color: c.color, hidden: '' }; }));
  var places = ps && ps.getLastRow() >= 2
    ? ps.getRange(2, 1, ps.getLastRow() - 1, PLACE_HEAD.length).getValues().map(function(r) {
        return { name: String(r[0] || '').trim(), alias: String(r[1] || '').split(/[、,，]/).map(function(a) { return a.trim(); }).filter(String),
          address: String(r[2] || '').trim(), shop: String(r[3] || '').trim() !== '', store: String(r[4] || '').trim(),
          noParen: String(r[5] || '').trim() !== '', hidden: dateText__(r[6], tz) };
      }).filter(function(p) { return p.name; })
    : (seed = true, PLACE_SEED.map(function(p) { return { name: p.name, alias: p.alias || [], address: '', shop: !!p.shop, store: '', noParen: !!p.noParen, hidden: '' }; }));
  return { clients: clients, places: places, seed: seed };
}
// みんなの状況から直す。kind: client|place、op: add|update|hide|show、key: 直す前の名前
function masterText_(v, n) { return String(v || '').replace(/[\r\n\t]/g, ' ').trim().replace(/^[=+\-@]+/, '').slice(0, n); }
function masterSave_(p) {
  if (!teamOk_(p.token)) return { ok: false, expired: true, error: '合言葉を確かめられませんでした。みんなの状況に入り直してください' };
  var kind = p.kind === 'place' ? 'place' : p.kind === 'client' ? 'client' : '';
  var op = String(p.op || ''), x = p.item || {};
  if (!kind || ['add', 'update', 'hide', 'show'].indexOf(op) < 0) return { ok: false, error: 'できない操作です' };
  var name = masterText_(x.name, 20).replace(/[,，、]/g, '');
  if (!name) return { ok: false, error: '名前を入れてください' };
  var sh = masterSheet_(kind === 'client' ? CLIENT_SHEET : PLACE_SHEET, true), head = kind === 'client' ? CLIENT_HEAD : PLACE_HEAD;
  var m = readMaster_(), list = kind === 'client' ? m.clients : m.places;
  var key = op === 'add' ? name : masterText_(p.key || x.name, 20);
  var idx = -1, dup = -1;
  list.forEach(function(it, i) { if (it.name === key) idx = i; if (it.name === name) dup = i; });
  var now = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd HH:mm');
  var clean;
  if (kind === 'client') {
    var col = String(x.color || '').trim();
    if (col && !/^#[0-9a-fA-F]{6}$/.test(col)) return { ok: false, error: '色の形がちがいます（#から始まる6けた）' };
    clean = { name: name, carrier: MASTER_CARRIERS.indexOf(x.carrier) >= 0 ? x.carrier : '', place: masterText_(x.place, 30), color: col, hidden: '', at: now };
  } else {
    var alias = (Array.isArray(x.alias) ? x.alias : String(x.alias || '').split(/[、,，]/)).map(function(a) { return masterText_(a, 20); }).filter(String).slice(0, 10);
    clean = { name: name, alias: alias, address: masterText_(x.address, 100), shop: !!x.shop, store: masterText_(x.store, 20), noParen: !!x.noParen, hidden: '', at: now };
  }
  if (op === 'add') {
    if (dup >= 0) return { ok: false, error: '「' + name + '」はもうあります' + (list[dup].hidden ? '（外したものの中）。「戻す」を押してください' : '') };
    list.push(clean); idx = list.length - 1;
  } else {
    if (idx < 0) return { ok: false, error: '「' + key + '」が見つかりません。開き直してください' };
    if (op === 'update') {
      if (dup >= 0 && dup !== idx) return { ok: false, error: '「' + name + '」はもうあります' };
      clean.hidden = list[idx].hidden; list[idx] = clean;
    }
    if (op === 'hide') { list[idx].hidden = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd'); list[idx].at = now; }
    if (op === 'show') { list[idx].hidden = ''; list[idx].at = now; }
  }
  var rows = list.map(function(it) { return kind === 'client' ? clientRow_(it) : placeRow_(it); });
  // シートの中身を一覧で書き直す（見出しの下）。行が減ることはない（外すのは印だけ）
  sh.getRange(2, 1, rows.length, head.length).setValues(rows);
  return rosterOk_(false);
}
// 稼働費の「稼働場所」（社員のブック・山口さんのブック）とパートナーの「場所」から、場所と住所を取り込む。
// 「SB光の森」→ 名前「光の森」・別名「SB光の森」・店舗稼働。もうある場所は、住所が空のときだけ住所を入れ、別名を足す
var MASTER_IMPORT = [
  { id: '1shpJvauT5rwKCaWH9RQBNZt3DzAA11yD5jNFUg2c6d4', sheet: '稼働場所', name: 0, addr: 1 },   // 稼働費計算_GAS版
  { id: '1RRLf47jjFvL1MdFvkt43GdSsw_r7NLQOqOQzRCEwmpA', sheet: '稼働場所', name: 0, addr: 1 },   // 山口_稼働費計算_GAS版
  { id: '19TTixFCWjVUR9S4XvsMjdbQt_U3v4MlCdvVpgTMsK_U', sheet: '場所', name: 2, addr: 3, del: 5 }   // パートナー_交通費（場所名・住所・削除）
];
function masterImport_(p) {
  if (!teamOk_(p.token)) return { ok: false, expired: true, error: '合言葉を確かめられませんでした。みんなの状況に入り直してください' };
  var m = readMaster_(), places = m.places, byKey = {}, added = 0, filled = 0, read = 0, skipped = [];
  places.forEach(function(pl, i) { byKey[placeKey_(pl.name)] = i; pl.alias.forEach(function(a) { if (byKey[placeKey_(a)] === undefined) byKey[placeKey_(a)] = i; }); });
  MASTER_IMPORT.forEach(function(src) {
    var sh;
    try { sh = SpreadsheetApp.openById(src.id).getSheetByName(src.sheet); } catch (e) { skipped.push(src.sheet + '（開けません）'); return; }
    if (!sh || sh.getLastRow() < 2) return;
    var w = Math.max(src.name, src.addr, src.del || 0) + 1;
    sh.getRange(2, 1, sh.getLastRow() - 1, w).getValues().forEach(function(r) {
      var raw = String(r[src.name] || '').trim(), addr = String(r[src.addr] || '').trim().replace(/^〒?\d{3}-\d{4}\s*/, '');
      if (!raw || !addr || /自宅/.test(raw) || (src.del !== undefined && String(r[src.del] || '').trim())) return;
      read++;
      var k = placeKey_(raw), i = byKey[k];
      if (i === undefined) {
        var nm = stripPrefix_(raw) || raw;
        places.push({ name: nm, alias: nm !== raw ? [raw] : [], address: addr, shop: hasPrefix_(raw), store: '', noParen: false, hidden: '' });
        byKey[k] = places.length - 1; added++;
      } else {
        var pl = places[i];
        if (!pl.address) { pl.address = addr; filled++; }
        if (raw !== pl.name && pl.alias.indexOf(raw) < 0 && pl.alias.length < 10) pl.alias.push(raw);
      }
    });
  });
  var sh = masterSheet_(PLACE_SHEET, true);
  if (places.length) sh.getRange(2, 1, places.length, PLACE_HEAD.length).setValues(places.map(placeRow_));
  var out = rosterOk_(false);
  out.imported = { read: read, added: added, filled: filled, skipped: skipped };
  return out;
}

function rosterOk_(already) {
  try { CacheService.getScriptCache().remove(ROSTER_CACHE); } catch (e) {}
  var out = roster_();
  out.already = already;
  return out;
}

// ═══ 読み込み ═══
// /exec をブラウザで直接開くと、デプロイ確認の応答が返る。日報は記録しない。
//   ?action=gains&name=名前&date=yyyy-mm-dd … その人のその日の獲得報告
//   ?action=gainStats&from=yyyy-mm          … その月以降の獲得報告の月ごとの合計（戦闘力用）
//   ?action=nippouStatus&month=yyyy-mm      … その月に日報を出した日（人ごと。みんなの状況の画面用）
//   ?action=cheer&name=名前                  … いちばん新しい日報への、AIのひとこと（ポータルのぼのぼの用）
//   ?action=yearReview&name=名前&year=yyyy   … 1年のふりかえり（ポータルの「あなたの○年」用）
//   ?action=roster                           … 名簿（ポータルと各ツールに出す人）
//   ?action=contactStatus&name=名前          … 連絡先を登録したか（電話の下4桁と日付だけ。中身は返さない）

// ═══ 連絡先（パートナー本人が「連絡先の登録」の画面で入れる。みんなの状況だけで見られる） ═══
// 「連絡先」シートに1人1行。名簿シートには入れない（名簿は合言葉なしで読まれるため）。
// 本人の画面は名前を選ぶだけで開けるので、そこには中身を返さない（登録したか・下4桁・日付だけ）。
// 全部を返すのは、合言葉の通行証があるとき（みんなの状況）だけ。
var CONTACT_SHEET = '連絡先';
var CONTACT_HEAD = ['名前', '電話', 'メール', '緊急連絡先（名前・続柄）', '緊急連絡先の電話', '更新日時'];
function contactSheet_(create) {
  var ss = SpreadsheetApp.openById(SHEET_ID), sh = ss.getSheetByName(CONTACT_SHEET);
  if (!sh && create) {
    sh = ss.insertSheet(CONTACT_SHEET);
    sh.appendRow(CONTACT_HEAD);
    if (sh.setFrozenRows) sh.setFrozenRows(1);
  }
  return sh;
}
function contactRows_() {
  var sh = contactSheet_(false);
  if (!sh || sh.getLastRow() < 2) return [];
  return sh.getRange(2, 1, sh.getLastRow() - 1, CONTACT_HEAD.length).getValues().map(function(r) {
    var at = r[5];
    return { name: String(r[0] || '').trim(), tel: String(r[1] || '').trim(), mail: String(r[2] || '').trim(),
      emName: String(r[3] || '').trim(), emTel: String(r[4] || '').trim(),
      at: Object.prototype.toString.call(at) === '[object Date]' ? Utilities.formatDate(at, 'Asia/Tokyo', 'yyyy-MM-dd') : String(at || '').slice(0, 10) };
  }).filter(function(x) { return x.name; });
}
// 電話番号：数字とハイフンだけ。数字が10〜11桁（0から始まる）
function cleanTel_(v, need) {
  var t = String(v || '').replace(/[０-９]/g, function(c) { return String.fromCharCode(c.charCodeAt(0) - 0xFEE0); })
    .replace(/[ー－‐−―]/g, '-').replace(/[\s　()（）]/g, '');
  if (!t) return need ? { error: '電話番号を入れてください' } : { tel: '' };
  if (!/^[0-9-]+$/.test(t)) return { error: '電話番号は数字で入れてください' };
  var digits = t.replace(/-/g, '');
  if (!/^0\d{9,10}$/.test(digits)) return { error: '電話番号の桁数を確かめてください（0から始まる10〜11桁）' };
  return { tel: t };
}
function contactSave_(p) {
  var name = String(p.name || '').trim();
  var people = (roster_().people || []).filter(function(x) { return !x.hidden; });
  if (!name || !people.some(function(x) { return x.name === name; })) return { ok: false, error: '名簿にいない名前です。ポータルで名前を選び直してください' };
  var tel = cleanTel_(p.tel, true); if (tel.error) return { ok: false, error: tel.error };
  var emTel = cleanTel_(p.emTel, false); if (emTel.error) return { ok: false, error: '緊急連絡先の' + emTel.error };
  var mail = String(p.mail || '').trim().slice(0, 100);
  if (mail && !/^[^\s@=+]+@[^\s@]+\.[^\s@]+$/.test(mail)) return { ok: false, error: 'メールアドレスの形を確かめてください' };
  var emName = String(p.emName || '').replace(/[\r\n\t]/g, ' ').trim().replace(/^[=+\-@]+/, '').slice(0, 30);
  var sh = contactSheet_(true), now = new Date();
  var row = [name, tel.tel, mail, emName, emTel.tel, now];
  var n = sh.getLastRow() - 1, at = -1;
  if (n > 0) sh.getRange(2, 1, n, 1).getValues().forEach(function(r, i) { if (String(r[0]).trim() === name) at = i + 2; });
  if (at < 0) at = sh.getLastRow() + 1;
  var rg = sh.getRange(at, 1, 1, row.length);
  if (rg.setNumberFormat) rg.setNumberFormat('@');   // 電話番号の先頭の0を消さない
  rg.setValues([row.map(function(v, i) { return i === 5 ? Utilities.formatDate(now, 'Asia/Tokyo', 'yyyy-MM-dd HH:mm') : v; })]);
  return { ok: true, status: contactStatus_(name) };
}
function contactStatus_(name) {
  name = String(name || '').trim();
  var me = contactRows_().filter(function(x) { return x.name === name; })[0];
  if (!me) return { ok: true, registered: false };
  var d = me.tel.replace(/-/g, '');
  return { ok: true, registered: true, tel4: d.slice(-4), at: me.at };
}
function contacts_(p) {
  if (!teamOk_(p.token)) return { ok: false, expired: true, error: '合言葉を確かめられませんでした。みんなの状況に入り直してください' };
  return { ok: true, people: contactRows_() };
}

function doGet(e) {
  var q = (e && e.parameter) || {};
  try {
    if (q.action === 'gains') return json_(dayGains_(String(q.name || ''), String(q.date || '')));
    if (q.action === 'gainStats') return json_(gainStats_(String(q.from || '')));
    if (q.action === 'nippouStatus') return json_(nippouStatus_(String(q.month || '')));
    if (q.action === 'cheer') return json_(cheer_(String(q.name || '')));
    if (q.action === 'yearReview') return json_(yearReview_(String(q.name || ''), String(q.year || '')));
    if (q.action === 'roster') return json_(roster_());
    if (q.action === 'contactStatus') return json_(contactStatus_(q.name));
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  }
  return json_({ ok: true, status: 'ready', version: CODE_VERSION, dedupe: true, gain: true, teamPass: true, cheer: true, remind: true, yearReview: true, logError: true, roster: true, contacts: true, master: true });
}

function doPost(e) {
  // ロックのいらないもの（不具合の記録・締めの声かけ・不具合の一覧）は、日報の送信を待たせないよう先に返す
  var p0 = null;
  try { p0 = JSON.parse(e.postData.contents); } catch (ignore) {}
  if (p0 && (p0.action === 'logError' || p0.action === 'remind' || p0.action === 'errors')) {
    try { return json_(p0.action === 'logError' ? logError_(p0) : p0.action === 'remind' ? remind_(p0) : errors_(p0)); }
    catch (err) { return json_({ ok: false, error: String(err) }); }
  }
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
  if (p.action === 'teamPass') return teamPass_(p);
  if (p.action === 'rosterSave') return rosterSave_(p);
  if (p.action === 'contactSave') return contactSave_(p);
  if (p.action === 'masterSave') return masterSave_(p);
  if (p.action === 'masterImport') return masterImport_(p);
  if (p.action === 'contacts') return contacts_(p);
  if (p.action === 'gainResend') return resendGain_(p);
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
var CARRIER_LABEL = { docomo: 'docomo', Softbank: 'SoftBank', au: 'au' };   // 書き方はポータル（sieg-products.js の label）と同じ

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
    discordCode: sent ? sent.code : null,
    items: day.items,
    reports: day.reports,
    discordOn: day.discordOn
  };
}

// Discord の獲得報告チャンネルに、いつもの書き方（「ハイエンド1件」を1行ずつ）で投稿する。
// 取り消しで消せるように ?wait=true でメッセージIDを受け取る。
// 戻り値：URL未設定なら null、それ以外は { ok, id, code }
function sendGainToDiscord_(name, carrier, date, items, memo) {
  var url = webhook_('DISCORD_GAIN_WEBHOOK');
  if (!url) return null;
  var lines = ['**' + (CARRIER_LABEL[carrier] || carrier) + '**　' + dayLabel_(date)];
  items.forEach(function(x) { lines.push(x.label + x.count + '件'); });
  if (memo) lines.push('', memo);
  var r = postDiscord_(url + (url.indexOf('?') >= 0 ? '&' : '?') + 'wait=true', {
    username: name,
    content: lines.join('\n').slice(0, 1900),
    allowed_mentions: { parse: [] }   // メモに @everyone などがあっても通知を飛ばさない
  });
  var id = '';
  if (r.ok) { try { id = String(JSON.parse(r.body).id || ''); } catch (ignore) {} }
  return { ok: r.ok, id: id, code: r.code };
}

// Discord に出せなかった報告を送り直す（本人の名前と送信IDが合う行）。
// すでに出ている（メッセージIDがある）ときは何もしない
function resendGain_(p) {
  var name = String(p.name || ''), id = String(p.id || '');
  if (!name || !id) return { ok: false, error: '送り直す報告が分かりません' };
  var sheet = gainSheet_();
  var tz = sheet.getParent().getSpreadsheetTimeZone();
  var data = sheet.getDataRange().getValues();
  var rows = [], items = [], carrier = '', date = '', memo = '', posted = false;
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][7]) !== id || String(data[i][2]) !== name) continue;
    rows.push(i + 1);
    carrier = String(data[i][3]); date = ymd_(data[i][1], tz); memo = String(data[i][6] || '');
    items.push({ label: String(data[i][4]), count: Number(data[i][5]) || 0 });
    if (data[i][8]) posted = true;
  }
  if (!rows.length) return { ok: false, error: 'その報告は見つかりませんでした（取り消し済みかもしれません）' };
  if (posted) return { ok: true, discord: true, already: true };
  var sent = sendGainToDiscord_(name, carrier, date, items, memo);
  if (!sent) return { ok: true, discord: null };
  if (sent.id) rows.forEach(function(r) { sheet.getRange(r, GAIN_HEAD.length).setValues([[sent.id]]); });
  return { ok: true, discord: sent.ok, code: sent.code };
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
  var url = webhook_('DISCORD_GAIN_WEBHOOK');
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
      reports[rid] = { id: String(r[7] || ''), at: at, carrier: carrier, memo: String(r[6] || ''), items: [], posted: !!r[8] };
      rorder.push(rid);
    }
    reports[rid].items.push({ label: label, count: count });
  }
  return {
    ok: true,
    date: date,
    discordOn: !!webhook_('DISCORD_GAIN_WEBHOOK'),   // false なら Discord に出ないのは設定がないため
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

// ═══ 日報の提出状況（みんなの状況の画面用） ═══
// month（yyyy-mm、無ければ今月）に、だれが何日の日報を出したか。
//   days: { 名前: ['yyyy-mm-dd', …] }（入店日。同じ日に何回送っても1日）
//   last: { 名前: 'M/d HH:mm' }（その月に最後に送った日時）
function nippouStatus_(month) {
  if (!/^\d{4}-\d{2}$/.test(month)) month = todayYmd_().slice(0, 7);
  var sheet = getSheet_();
  var tz = sheet.getParent().getSpreadsheetTimeZone();
  var data = sheet.getDataRange().getValues();
  var days = {}, last = {}, lastAt = {};
  for (var i = 1; i < data.length; i++) {
    var name = String(data[i][2] || ''), d = ymd_(data[i][1], tz);
    if (!name || d.slice(0, 7) !== month) continue;
    days[name] = days[name] || [];
    if (days[name].indexOf(d) < 0) days[name].push(d);
    var at = data[i][0];
    if (Object.prototype.toString.call(at) === '[object Date]' && (!lastAt[name] || at > lastAt[name])) {
      lastAt[name] = at;
      last[name] = Utilities.formatDate(at, 'Asia/Tokyo', 'M/d HH:mm');
    }
  }
  Object.keys(days).forEach(function(n) { days[n].sort(); });
  return { ok: true, month: month, today: todayYmd_(), days: days, last: last };
}
