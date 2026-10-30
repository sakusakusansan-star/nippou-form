// 実行: node nippou-gas.test.js  （Node単体で動く。GASへ貼る前の確認用）
// スプレッドシート・Discord をモックにして、nippou-gas.gs の日報と獲得報告を確かめる
const fs = require('fs');
const path = require('path');
const assert = require('assert');

function Sheet(name, grid) {
  const g = grid;
  const width = () => g.reduce((w, r) => Math.max(w, r.length), 0);
  const s = {
    getName: () => name,
    getLastRow: () => g.length,
    appendRow: (vals) => g.push(vals.slice()),
    setFrozenRows() {},
    deleteRow: (r) => g.splice(r - 1, 1),
    getParent: () => ss,
    getDataRange: () => s.getRange(1, 1, g.length || 1, width() || 1),
    getRange(r, c, nr, nc) {
      nr = nr || 1; nc = nc || 1;
      return {
        getValues() {
          const out = [];
          for (let i = 0; i < nr; i++) {
            const line = [];
            for (let j = 0; j < nc; j++) line.push(g[r - 1 + i] && g[r - 1 + i][c - 1 + j] !== undefined ? g[r - 1 + i][c - 1 + j] : '');
            out.push(line);
          }
          return out;
        },
        setValues(vals) {
          for (let i = 0; i < vals.length; i++) {
            while (g.length < r + i) g.push([]);
            for (let j = 0; j < vals[i].length; j++) g[r - 1 + i][c - 1 + j] = vals[i][j];
          }
        }
      };
    },
    _grid: g
  };
  return s;
}

let ss;
const posts = [];
let props = {};
function load() {
  const sheets = [Sheet('日報', [['送信日時']])];
  ss = {
    getSheetByName: (n) => sheets.find(x => x.getName() === n) || null,
    insertSheet: (n) => { const x = Sheet(n, []); sheets.push(x); return x; },
    getSpreadsheetTimeZone: () => 'Asia/Tokyo'
  };
  const cache = {};
  global.SpreadsheetApp = { openById: () => ss };
  global.LockService = { getScriptLock: () => ({ waitLock() {}, releaseLock() {} }) };
  global.CacheService = { getScriptCache: () => ({ get: (k) => cache[k] || null, put: (k, v) => { cache[k] = v; } }) };
  global.ContentService = { MimeType: { JSON: 'json' }, createTextOutput: (t) => ({ text: t, setMimeType() { return this; } }) };
  global.PropertiesService = { getScriptProperties: () => ({ getProperty: (k) => props[k] || null, getKeys: () => Object.keys(props) }) };
  global.Logger = { log() {} };
  // Asia/Tokyo 固定の簡易版
  global.Utilities = {
    formatDate(d, tz, f) {
      const j = new Date(d.getTime() + 9 * 3600000);
      const p = (n) => String(n).padStart(2, '0');
      return f.replace('yyyy', j.getUTCFullYear()).replace('MM', p(j.getUTCMonth() + 1)).replace('dd', p(j.getUTCDate()))
        .replace('HH', p(j.getUTCHours())).replace('mm', p(j.getUTCMinutes()));
    }
  };
  let msg = 1000;
  global.UrlFetchApp = {
    fetch(url, opt) {
      posts.push({ url, method: opt.method, body: opt.payload ? JSON.parse(opt.payload) : null });
      return { getResponseCode: () => (opt.method === 'delete' ? 204 : 200), getContentText: () => JSON.stringify({ id: String(++msg) }) };
    }
  };
  const code = fs.readFileSync(path.join(__dirname, 'nippou-gas.gs'), 'utf8');
  (0, eval)(code);
}
const post = (p) => JSON.parse(doPost({ postData: { contents: JSON.stringify(p) } }).text);
const get = (q) => JSON.parse(doGet({ parameter: q }).text);

load();

// ─── MNP の数え方 ───
assert.strictEqual(mnpFromBreakdown_('MNP3 / 新規1'), 3, '古い書き方');
assert.strictEqual(mnpFromBreakdown_('SMNP2 / YMNP1 / 光新規1 / HOME5G12'), 3, 'SB の MNP 2種類');
assert.strictEqual(mnpFromBreakdown_('auMNP1 / UQMNP4 / auカード1'), 5);
assert.strictEqual(mnpCount_([{ label: 'SMNP', count: 2 }, { label: 'YMNP', count: 1 }, { label: 'Air新規', count: 5 }]), 3);

// ─── 読み込み（未設定のとき）───
let r = get({});
assert.strictEqual(r.version, '2026-10-31');
assert.strictEqual(r.gain, true);

// ─── Discord 未設定でも記録できる ───
r = post({ action: 'gain', requestId: 'a1', name: '清田', carrier: 'docomo', date: '2026-10-30',
  items: [{ label: 'ハイエンド', count: 2 }, { label: 'Dカードプラチナ', count: 1 }, { label: '新規', count: 0 }], memo: '' });
assert.strictEqual(r.ok, true);
assert.strictEqual(r.discord, null);
assert.deepStrictEqual(r.items, [{ carrier: 'docomo', label: 'ハイエンド', count: 2 }, { carrier: 'docomo', label: 'Dカードプラチナ', count: 1 }]);
assert.strictEqual(ss.getSheetByName('獲得')._grid.length, 3, '見出し＋2行（0件は書かない）');

// ─── Discord に投稿し、メッセージIDを覚える ───
props.DISCORD_GAIN_WEBHOOK = 'https://discord.com/api/webhooks/1/tok';
r = post({ action: 'gain', requestId: 'a2', name: '清田', carrier: 'docomo', date: '2026-10-30',
  items: [{ label: 'ハイエンド', count: 1 }], memo: '@everyone 達成のため11月まわし' });
assert.strictEqual(r.discord, true);
const sent = posts[posts.length - 1];
assert.ok(/wait=true$/.test(sent.url));
assert.strictEqual(sent.body.username, '清田');
assert.deepStrictEqual(sent.body.allowed_mentions, { parse: [] });
assert.strictEqual(sent.body.content, '**ドコモ**　10/30（金）\nハイエンド1件\n\n@everyone 達成のため11月まわし');
assert.deepStrictEqual(r.items[0], { carrier: 'docomo', label: 'ハイエンド', count: 3 }, '同じ日の分は合計');
assert.strictEqual(r.reports.length, 2);
assert.strictEqual(r.reports[0].id, 'a2', '新しい順');

// ─── 同じ送信IDは二重に記録しない ───
const before = ss.getSheetByName('獲得')._grid.length, nPosts = posts.length;
r = post({ action: 'gain', requestId: 'a2', name: '清田', carrier: 'docomo', date: '2026-10-30', items: [{ label: 'ハイエンド', count: 1 }] });
assert.strictEqual(r.duplicate, true);
assert.strictEqual(ss.getSheetByName('獲得')._grid.length, before);
assert.strictEqual(posts.length, nPosts);

// ─── 入力の確認 ───
assert.strictEqual(post({ action: 'gain', requestId: 'x', name: '', carrier: 'au', items: [{ label: 'auMNP', count: 1 }] }).ok, false);
assert.strictEqual(post({ action: 'gain', requestId: 'y', name: '鈴木', carrier: 'rakuten', items: [{ label: 'MNP', count: 1 }] }).ok, false);
assert.strictEqual(post({ action: 'gain', requestId: 'z', name: '鈴木', carrier: 'au', items: [{ label: 'auMNP', count: 0 }] }).ok, false);

// ─── 別の人・別の日 ───
post({ action: 'gain', requestId: 'b1', name: '鈴木', carrier: 'Softbank', date: '2026-10-30', items: [{ label: 'SMNP', count: 2 }, { label: 'YMNP', count: 1 }] });
post({ action: 'gain', requestId: 'b2', name: '鈴木', carrier: 'Softbank', date: '2026-09-15', items: [{ label: 'SMNP', count: 4 }] });
r = get({ action: 'gains', name: '鈴木', date: '2026-10-30' });
assert.deepStrictEqual(r.items.map(x => x.label + x.count), ['SMNP2', 'YMNP1']);

// ─── 戦闘力用の集計 ───
r = get({ action: 'gainStats', from: '2026-09' });
assert.strictEqual(r.ok, true);
const rows = r.rows.map(x => x.join('|')).sort();
assert.deepStrictEqual(rows, [
  '2026-09|鈴木|Softbank|SMNP|4',
  '2026-10|清田|docomo|Dカードプラチナ|1',
  '2026-10|清田|docomo|ハイエンド|3',
  '2026-10|鈴木|Softbank|SMNP|2',
  '2026-10|鈴木|Softbank|YMNP|1'
]);
assert.strictEqual(get({ action: 'gainStats', from: '2026-10' }).rows.length, 4);

// ─── キャラの選択（1人1行・上書き）───
assert.deepStrictEqual(get({ action: 'gainStats' }).fighters, {});
assert.strictEqual(post({ action: 'fighter', name: '鈴木', id: '03' }).ok, true);
assert.strictEqual(post({ action: 'fighter', name: '清田', id: '10' }).ok, true);
assert.strictEqual(post({ action: 'fighter', name: '鈴木', id: '07' }).ok, true);
assert.strictEqual(post({ action: 'fighter', name: '鈴木', id: 'x' }).ok, false);
assert.deepStrictEqual(get({ action: 'gainStats' }).fighters, { 鈴木: '07', 清田: '10' });
assert.strictEqual(ss.getSheetByName('ファイター')._grid.length, 3);

// ─── 取り消し（本人の分だけ。Discord の投稿も消す）───
r = post({ action: 'gainUndo', name: '鈴木', id: 'a2' });
assert.strictEqual(r.removed, 0, '他人の報告は消せない');
r = post({ action: 'gainUndo', name: '清田', id: 'a2' });
assert.strictEqual(r.removed, 1);
assert.strictEqual(r.discord, true);
assert.strictEqual(posts[posts.length - 1].method, 'delete');
assert.ok(/\/webhooks\/1\/tok\/messages\/\d+$/.test(posts[posts.length - 1].url));
r = get({ action: 'gains', name: '清田', date: '2026-10-30' });
assert.deepStrictEqual(r.items.map(x => x.label + x.count), ['ハイエンド2', 'Dカードプラチナ1']);
assert.strictEqual(post({ action: 'gainUndo', name: '清田', id: 'a2' }).removed, 0, '2回目は何もしない');

// ─── 日報は今まで通り（SB の MNP は SMNP・YMNP の合計）───
props.DISCORD_WEBHOOK = 'https://discord.com/api/webhooks/2/tok';
r = post({ requestId: 'n1', name: '鈴木', date: '2026-10-30', carrier: 'Softbank', place: '下通',
  approaches: [{ label: 'キャッチ', count: 10 }], gains: [{ label: 'SMNP', count: 2 }, { label: 'YMNP', count: 1 }, { label: '光新規', count: 1 }],
  target: 10, prev: 4, note: '' });
assert.strictEqual(r.ok, true);
assert.strictEqual(r.monthlyMnp, 3);
assert.ok(r.text.indexOf('通算 7件') >= 0, r.text);
const embed = posts[posts.length - 1].body.embeds[0];
assert.strictEqual(posts[posts.length - 1].body.username, '鈴木', '日報も送った人の名前で投稿');
assert.ok(embed.fields.some(f => /本日MNP 3件/.test(f.value)));

console.log('nippou-gas: all tests passed');
