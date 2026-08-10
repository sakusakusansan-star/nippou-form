// ═══════════════════════════════════════════
//  Sieg 日報 GAS
//  1. 新しいスプレッドシートを作成し、IDを下に貼る
//  2. Teams → Workflows →「Webhook 要求を受信したらチャットに投稿する」
//     テンプレートで日報グループ宛のフローを作成し、URLを下に貼る（使わない場合は空文字のままでOK）
//  3. Discordへ自動投稿する場合は、プロジェクトの設定 → スクリプト プロパティに
//     DISCORD_WEBHOOK として Discord の Webhook URL を登録する（コードには書かない）
//  4. setup() を1回実行（ヘッダー作成）
//  5. デプロイ → ウェブアプリ →「全員（匿名含む）」で公開
//  6. デプロイURLを nippou-form の NIPPOU_API に貼る
// ═══════════════════════════════════════════

var SHEET_ID = '1ffkmHDt_sz1ZQ3ZB_fXIaJeFhuW0wSI7bqwMpC9Z33g';
var SHEET_NAME = '日報';
// Teams自動投稿を使わない場合は空文字のままでOK（コピペ運用ならシート保存のみで十分）
var TEAMS_WEBHOOK_URL = '';

function setup() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) sheet = ss.insertSheet(SHEET_NAME);
  sheet.getRange(1, 1, 1, 13).setValues([[
    '送信日時', '入店日', '名前', 'キャリア', '稼働場所',
    'アプローチ合計', 'アプローチ内訳', '獲得合計', '獲得内訳',
    'MNP目標', 'MNP通算進捗', '進捗率', '所感'
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

// 内訳配列から特定ラベルの件数だけ取り出す（進捗はMNPのみでカウント）
function pickCount_(arr, label) {
  if (!arr || !arr.length) return 0;
  var hit = arr.filter(function(x) { return x.label === label; });
  return hit.length ? (Number(hit[0].count) || 0) : 0;
}

// 同じ名前の今月の獲得合計（今回分を含めるため、追記後に呼ぶ）
function monthlyTotal_(name) {
  var data = getSheet_().getDataRange().getValues();
  var now = new Date();
  var total = 0;
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][2]) !== name) continue;
    var d = new Date(data[i][1]);
    if (isNaN(d.getTime())) continue;
    if (d.getFullYear() === now.getFullYear() && d.getMonth() === now.getMonth()) {
      total += Number(data[i][7]) || 0;
    }
  }
  return total;
}

// 進捗バー "■■■□□□□□□□"
function bar_(pct) {
  var filled = Math.max(0, Math.min(10, Math.round(pct / 10)));
  return Array(filled + 1).join('■') + Array(10 - filled + 1).join('□');
}

// ═══ Discord送信 ═══
// スクリプトプロパティ「DISCORD_WEBHOOK」にURLが設定されていれば #受注速報 等へ自動投稿
//
// meta には doPost で算出済みの値を渡す：
//   { dateStr, pct, cumulative, prev, mnpCount, target, monthly, approachTotal, gainTotal }
// 渡っていない項目は p から計算し直すので、meta が dateStr だけでも動作する
// （ただし monthly はシート由来のため、渡さないと 0件 表示になる）
//
// 戻り値：URL未設定なら null、投稿を試みたら成功/失敗の true/false（doPost の応答に使う）
function sendToDiscord_(copyText, p, meta) {
  var url = PropertiesService.getScriptProperties().getProperty('DISCORD_WEBHOOK');
  if (!url) return null;

  var m = meta || {};

  var target        = Number(m.target) || Number(p.target) || 0;
  var prev          = Number(m.prev) || Number(p.prev) || 0;
  var mnpCount      = m.mnpCount != null ? (Number(m.mnpCount) || 0) : pickCount_(p.gains, 'MNP');
  var cumulative    = m.cumulative != null ? (Number(m.cumulative) || 0) : (prev + mnpCount);
  var approachTotal = m.approachTotal != null ? (Number(m.approachTotal) || 0) : sumBreakdown_(p.approaches);
  var gainTotal     = m.gainTotal != null ? (Number(m.gainTotal) || 0) : sumBreakdown_(p.gains);
  var monthly       = Number(m.monthly) || 0;
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

  fields.push({ name: '📈 今月通算', value: monthly + '件' });

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

function doPost(e) {
  try {
    var p = JSON.parse(e.postData.contents);
    if (!p.name) return json_({ ok: false, error: '名前がありません' });

    var approachTotal = sumBreakdown_(p.approaches);
    var gainTotal = sumBreakdown_(p.gains);
    var mnpCount = pickCount_(p.gains, 'MNP');
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
      String(p.note || '').trim()
    ]);

    var monthly = monthlyTotal_(String(p.name));

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
      body.push({ type: 'TextBlock', text: '📈 今月通算：' + monthly + '件', wrap: true });
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
      approachTotal: approachTotal,
      gainTotal: gainTotal
    });

    return json_({ ok: true, teams: teamsOk, discord: discordOk, monthly: monthly, text: copyText });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  }
}
