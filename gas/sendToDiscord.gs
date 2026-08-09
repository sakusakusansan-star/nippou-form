// ═══════════════════════════════════════════════════════════
//  Discord 投稿（fields 分割 ＋ MNP進捗連動カラー）
//  nippou-gas.gs の sendToDiscord_ 関数と丸ごと差し替えて使う
// ═══════════════════════════════════════════════════════════

/** MNP進捗に応じた embed カラー */
var DISCORD_COLOR_IDLE  = 0x8A93A6; // target未設定 / 50%未満
var DISCORD_COLOR_AMBER = 0xE8A33D; // 50%以上100%未満
var DISCORD_COLOR_DONE  = 0x4CD97B; // 100%以上（目標達成）

/**
 * Discord へ日報を投稿する。
 *
 * @param {Object} p       フォームから受け取ったペイロード
 *        { date, name, carrier, place, approaches[], gains[], target, prev, note }
 * @param {string} text    コピー用テキスト（互換のため残置。embed では未使用）
 * @param {Object} [stats] doPost 内で算出済みの集計値。省略時は p から再計算する。
 *        { pct, cumulative, prev, mnpCount, target, monthly, approachTotal, gainTotal }
 */
function sendToDiscord_(p, text, stats) {
  // ▼ Webhook URL の取得は既存コードのまま（変更しないこと）
  var url = DISCORD_WEBHOOK_URL;
  if (!url) return;

  // ── 小道具（グローバルを汚さないようローカル関数で持つ） ──
  var n = function (v) {
    if (typeof v === 'number' && isFinite(v)) return v;
    return parseInt(v, 10) || 0;
  };
  var totalOf = function (arr) {
    var t = 0;
    (arr || []).forEach(function (x) { t += n(x && x.count); });
    return t;
  };
  var countOf = function (arr, label) {
    var t = 0;
    (arr || []).forEach(function (x) { if (x && x.label === label) t += n(x.count); });
    return t;
  };
  var pick = function (v, fallback) { return (v === undefined || v === null) ? fallback : n(v); };

  // ── 集計値（stats が渡されていればそれを優先、無ければ p から再計算） ──
  var s             = stats || {};
  var target        = pick(s.target, n(p.target));
  var prev          = pick(s.prev, n(p.prev));
  var mnpCount      = pick(s.mnpCount, countOf(p.gains, 'MNP'));
  var cumulative    = pick(s.cumulative, prev + mnpCount);
  var pct           = pick(s.pct, target > 0 ? Math.round(cumulative / target * 100) : 0);
  var monthly       = pick(s.monthly, 0);
  var approachTotal = pick(s.approachTotal, totalOf(p.approaches));
  var gainTotal     = pick(s.gainTotal, totalOf(p.gains));

  var hasTarget = target > 0;
  var achieved  = hasTarget && pct >= 100;

  // ── カラー ──
  var color = DISCORD_COLOR_IDLE;
  if (hasTarget && pct >= 100)     color = DISCORD_COLOR_DONE;
  else if (hasTarget && pct >= 50) color = DISCORD_COLOR_AMBER;

  // ── タイトル（形式は現状維持。達成時のみ末尾に 🎉） ──
  var dateLabel = (typeof fmtDate_ === 'function') ? fmtDate_(p.date) : p.date;
  var title = '📋 日報｜' + p.name + '　' + dateLabel + (achieved ? ' 🎉' : '');

  // ── fields ──
  var withTotal = function (total, breakdown) {
    var v = '**' + total + '件**';
    if (breakdown) v += '\n' + breakdown;
    return v;
  };

  var fields = [
    { name: '📍 稼働場所', value: p.place || '—',   inline: true },
    { name: '📶 キャリア', value: p.carrier || '—', inline: true },
    { name: '🚶 アプローチ', value: withTotal(approachTotal, fmtBreakdown_(p.approaches)), inline: false },
    { name: '🎯 獲得',       value: withTotal(gainTotal, fmtBreakdown_(p.gains)),          inline: false }
  ];

  // MNP進捗：target 未設定のときはフィールドごと省略
  if (hasTarget) {
    fields.push({
      name: '📊 MNP進捗',
      value: bar_(pct) + ' ' + pct + '%\n'
           + '通算 ' + cumulative + '/' + target + '件'
           + '（前日 ' + prev + '件 ＋ 本日MNP ' + mnpCount + '件）',
      inline: false
    });
  }

  fields.push({ name: '📈 今月通算', value: monthly + '件', inline: false });

  // 所感：入力があるときだけ追加（Discord のフィールド上限 1024 文字に丸める）
  var note = (p.note || '').trim();
  if (note) {
    fields.push({
      name: '💬 所感',
      value: note.length > 1000 ? note.slice(0, 1000) + '…' : note,
      inline: false
    });
  }

  var payload = {
    embeds: [{
      title: title,
      color: color,
      fields: fields,
      footer: { text: Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm') + ' 送信' }
    }]
  };

  UrlFetchApp.fetch(url, {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  });
}

// ═══════════════════════════════════════════════════════════
//  doPost 内の呼び出し部分（任意）
//  ─ 引数を増やさず sendToDiscord_(p, text) のままでも動作する。
//    計算の二度手間を避けたい場合のみ、第3引数を渡す形に差し替える。
// ═══════════════════════════════════════════════════════════
//
//   sendToDiscord_(p, text, {
//     pct: pct,
//     cumulative: cumulative,
//     prev: prev,
//     mnpCount: mnpCount,
//     target: target,
//     monthly: monthly,
//     approachTotal: approachTotal,
//     gainTotal: gainTotal
//   });
//
