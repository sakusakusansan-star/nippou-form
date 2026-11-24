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
        getValue() { return g[r - 1] && g[r - 1][c - 1] !== undefined ? g[r - 1][c - 1] : ''; },
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
const gemini = [], geminiQueue = [];
const failQueue = [];
const sleeps = [];
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
  global.LockService = { getScriptLock: () => ({ waitLock() {}, tryLock: () => true, releaseLock() {} }) };
  global.CacheService = { getScriptCache: () => ({ get: (k) => cache[k] || null, put: (k, v) => { cache[k] = v; }, remove: (k) => { delete cache[k]; } }) };
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
      if (/generativelanguage/.test(url)) {   // Gemini
        gemini.push({ url, body: JSON.parse(opt.payload) });
        const g = geminiQueue.shift() || { text: JSON.stringify({ cheer: '「下通で キャッチ30件、すごいがんばり😊」\n（説明）', grow: '成約率 先週より 上がったね', hint: '着座に つなげると もっと伸びそう' }) };
        if (g.code) return { getResponseCode: () => g.code, getContentText: () => '{"error":{"message":"x"}}' };
        return { getResponseCode: () => 200, getContentText: () => JSON.stringify({ candidates: [{ content: { parts: [{ text: g.text }] } }] }) };
      }
      posts.push({ url, method: opt.method, body: opt.payload ? JSON.parse(opt.payload) : null });
      const f = failQueue.shift();   // テストで失敗させたい応答を順に入れる
      if (f === 'throw') throw new Error('Address unavailable');
      if (f) return { getResponseCode: () => f.code, getContentText: () => JSON.stringify(f.body || {}) };
      return { getResponseCode: () => (opt.method === 'delete' ? 204 : 200), getContentText: () => JSON.stringify({ id: String(++msg) }) };
    }
  };
  global.Utilities.sleep = (ms) => sleeps.push(ms);
  global.Utilities.DigestAlgorithm = { SHA_256: 'sha256' };
  global.Utilities.Charset = { UTF_8: 'utf8' };
  global.Utilities.computeDigest = (alg, s) => [...require('crypto').createHash('sha256').update(s, 'utf8').digest()].map(b => b > 127 ? b - 256 : b);
  global.Utilities.base64EncodeWebSafe = (bytes) => Buffer.from(bytes.map(b => b & 255)).toString('base64').replace(/\+/g, '-').replace(/\//g, '_');
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
assert.strictEqual(r.version, '2026-11-24');
assert.strictEqual(r.roster, true);
assert.ok(r.remind && r.yearReview && r.logError, '新しい読み書きがある');
assert.strictEqual(r.teamPass, true);
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
assert.strictEqual(sent.body.content, '**docomo**　10/30（金）\nハイエンド1件\n\n@everyone 達成のため11月まわし');
assert.deepStrictEqual(r.items[0], { carrier: 'docomo', label: 'ハイエンド', count: 3 }, '同じ日の分は合計');
assert.strictEqual(r.reports.length, 2);
assert.strictEqual(r.reports[0].id, 'a2', '新しい順');

// ─── Discord が一時的に断ったら送り直す ───
failQueue.push({ code: 429, body: { retry_after: 0.7 } }, 'throw');
let n0 = posts.length;
r = post({ action: 'gain', requestId: 'r429', name: '福山', carrier: 'Softbank', date: '2026-10-30', items: [{ label: 'SMNP', count: 1 }] });
assert.strictEqual(r.discord, true, '3回目で届く');
assert.strictEqual(posts.length - n0, 3);
assert.deepStrictEqual(sleeps.slice(-2), [950, 2000], '429 は Discord の指定どおり待つ');
assert.strictEqual(get({ action: 'gains', name: '福山', date: '2026-10-30' }).reports[0].posted, true);

// 400 などは送り直さない
failQueue.push({ code: 400, body: { message: 'bad' } });
n0 = posts.length;
r = post({ action: 'gain', requestId: 'r400', name: '福山', carrier: 'Softbank', date: '2026-10-30', items: [{ label: 'YMNP', count: 1 }] });
assert.strictEqual(r.discord, false);
assert.strictEqual(r.discordCode, 400);
assert.strictEqual(posts.length - n0, 1);

// 3回とも失敗 → 記録は残り、あとから「再送」できる
failQueue.push({ code: 502 }, { code: 502 }, { code: 502 });
r = post({ action: 'gain', requestId: 'r502', name: '福山', carrier: 'Softbank', date: '2026-10-30', items: [{ label: 'pixel', count: 2 }] });
assert.strictEqual(r.discord, false);
let day = get({ action: 'gains', name: '福山', date: '2026-10-30' });
assert.strictEqual(day.discordOn, true);
assert.deepStrictEqual(day.reports.map(x => x.id + ':' + x.posted), ['r502:false', 'r400:false', 'r429:true']);
r = post({ action: 'gainResend', name: '福山', id: 'r502' });
assert.strictEqual(r.discord, true);
assert.strictEqual(posts[posts.length - 1].body.content, '**SoftBank**　10/30（金）\npixel2件');
assert.strictEqual(get({ action: 'gains', name: '福山', date: '2026-10-30' }).reports[0].posted, true);
n0 = posts.length;
assert.strictEqual(post({ action: 'gainResend', name: '福山', id: 'r502' }).already, true, '2回目は投稿しない');
assert.strictEqual(posts.length, n0);
assert.strictEqual(post({ action: 'gainResend', name: '鈴木', id: 'r400' }).ok, false, '他人の報告は送れない');
// 片付け（このあとの集計のテストに混ざらないように）
['r429', 'r400', 'r502'].forEach(id => post({ action: 'gainUndo', name: '福山', id }));

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

// ─── 日報の提出状況 ───
post({ requestId: 'n2', name: '鈴木', date: '2026-10-30', carrier: 'Softbank', place: '下通', approaches: [], gains: [], target: 0, prev: 0, note: '' });
post({ requestId: 'n3', name: '毛利', date: '2026-10-29', carrier: 'Softbank', place: '下通', approaches: [], gains: [], target: 0, prev: 0, note: '' });
post({ requestId: 'n4', name: '毛利', date: '2026-09-29', carrier: 'Softbank', place: '下通', approaches: [], gains: [], target: 0, prev: 0, note: '' });
r = get({ action: 'nippouStatus', month: '2026-10' });
assert.strictEqual(r.ok, true);
assert.deepStrictEqual(r.days, { 鈴木: ['2026-10-30'], 毛利: ['2026-10-29'] }, '同じ日に2回送っても1日・別の月は数えない');
assert.ok(r.last['鈴木']);

// ─── みんなの状況の合言葉 ───
const rowsBefore = ss.getSheetByName('日報')._grid.length;
r = post({ action: 'teamPass', pass: 'x' });
assert.strictEqual(r.ok, false); assert.strictEqual(r.notSet, true, '未設定なら入れない');
props.TEAM_PASS = ' sieg-2026 \n';
r = post({ action: 'teamPass', pass: 'ちがう' });
assert.deepStrictEqual(r, { ok: false, error: '合言葉がちがいます' }, '通行証は返さない');
r = post({ action: 'teamPass', pass: 'sieg-2026' });
assert.strictEqual(r.ok, true); assert.ok(/^[A-Za-z0-9_-]{43}$/.test(r.token), r.token);
assert.ok(r.token.indexOf('sieg') < 0);
const tok = r.token;
assert.deepStrictEqual(post({ action: 'teamPass', token: tok }), { ok: true }, '覚えた通行証で入れる');
// 連絡先：本人が入れる。本人の画面には下4桁と日付だけ。全部はみんなの状況（通行証）だけ
assert.deepStrictEqual(get({ action: 'contactStatus', name: '鈴木' }), { ok: true, registered: false });
let cr = post({ action: 'contactSave', name: '鈴木', tel: '０９０ー１２３４ー５６７８', mail: 'suzuki@example.com', emName: '鈴木花子（母）', emTel: '096-123-4567' });
assert.ok(cr.ok, cr.error);
assert.deepStrictEqual([cr.status.registered, cr.status.tel4], [true, '5678']);
assert.ok(!JSON.stringify(cr).includes('090') && !JSON.stringify(get({ action: 'contactStatus', name: '鈴木' })).includes('1234'), '本人の画面には番号を返さない');
assert.deepStrictEqual(ss.getSheetByName('連絡先')._grid[0], ['名前', '電話', 'メール', '緊急連絡先（名前・続柄）', '緊急連絡先の電話', '更新日時']);
// 2回目は上書き（行は増えない）
cr = post({ action: 'contactSave', name: '鈴木', tel: '080-1111-2222' });
assert.ok(cr.ok, cr.error);
assert.strictEqual(ss.getSheetByName('連絡先')._grid.length, 2);
assert.strictEqual(post({ action: 'contacts', token: 'nise' }).expired, true, '通行証が無いと読めない');
let cl = post({ action: 'contacts', token: tok });
assert.ok(cl.ok);
assert.deepStrictEqual(cl.people.map(x => [x.name, x.tel, x.mail, x.emName, x.emTel]), [['鈴木', '080-1111-2222', '', '', '']]);
// 断るもの：名簿にいない名前・桁数の違う番号・番号なし・変なメール
assert.ok(/名簿にいない/.test(post({ action: 'contactSave', name: 'だれか', tel: '090-1234-5678' }).error));
assert.ok(/桁数/.test(post({ action: 'contactSave', name: '鈴木', tel: '090-1234' }).error));
assert.ok(/入れてください/.test(post({ action: 'contactSave', name: '鈴木', tel: '' }).error));
assert.ok(/メール/.test(post({ action: 'contactSave', name: '鈴木', tel: '09012345678', mail: '=HYPERLINK(1)' }).error));
assert.ok(/緊急連絡先/.test(post({ action: 'contactSave', name: '鈴木', tel: '09012345678', emTel: 'abc' }).error));
assert.strictEqual(post({ action: 'teamPass', token: 'nise' }).expired, true);
assert.strictEqual(ss.getSheetByName('日報')._grid.length, rowsBefore, '合言葉の確認で日報は増えない');
props.TEAM_PASS = 'kaeta';
assert.strictEqual(post({ action: 'teamPass', token: tok }).expired, true, '合言葉を変えたら前の端末は入り直し');
for (let i = 0; i < 9; i++) post({ action: 'teamPass', pass: 'x' + i });   // ここまでで10回まちがい
r = post({ action: 'teamPass', pass: 'kaeta' });
assert.strictEqual(r.locked, true, '10回まちがえたら正しくてもしばらく受け付けない');

// ─── 日報へのAIのひとこと（ほめる・のびた・ヒント） ───
{
  const day = n => Utilities.formatDate(new Date(Date.now() - n * 86400000), 'Asia/Tokyo', 'yyyy-MM-dd');
  const today = day(0), old = day(4);
  // 山内：前の週（8〜13日前）は キャッチ20・獲得0 が2回、今週は きょう キャッチ30＋着座5・獲得2 と 4日前 キャッチ12
  post({ requestId: 'w1', name: '山内', date: day(10), carrier: 'au', place: '近見', approaches: [{ label: 'キャッチ', count: 20 }], gains: [], target: 5, prev: 0, note: '' });
  post({ requestId: 'w2', name: '山内', date: day(9), carrier: 'au', place: '近見', approaches: [{ label: 'キャッチ', count: 20 }], gains: [], target: 5, prev: 0, note: '' });
  post({ requestId: 'c1', name: '山内', date: old, carrier: 'au', place: '近見', approaches: [{ label: 'キャッチ', count: 12 }], gains: [], target: 5, prev: 0, note: '' });
  post({ requestId: 'c2', name: '山内', date: today, carrier: 'au', place: '下通', approaches: [{ label: 'キャッチ', count: 30 }, { label: '着座', count: 5 }], gains: [{ label: 'auMNP', count: 2 }], target: 5, prev: 1, note: '雨だったけど粘った' });
  // 鍵が無いうちは空（ぼのぼのは数字のひとことだけ）
  delete props.GEMINI_API_KEY;
  let c = get({ action: 'cheer', name: '山内' });
  assert.deepStrictEqual([c.ok, c.text], [true, '']);
  assert.strictEqual(gemini.length, 0);
  props.GEMINI_API_KEY = 'AIza-test';
  c = get({ action: 'cheer', name: '山内' });
  assert.deepStrictEqual([c.text, c.grow, c.hint], ['下通で キャッチ30件、すごいがんばり', '成約率 先週より 上がったね', '着座に つなげると もっと伸びそう'], '1行目だけ・かぎかっこと絵文字をのぞく');
  assert.strictEqual(c.date, today);
  assert.strictEqual(gemini.length, 1);
  const body = gemini[0].body, prompt = body.contents[0].parts[0].text;
  assert.strictEqual(body.generationConfig.responseMimeType, 'application/json');
  // 所感を中心に、過去の日報（所感つき）と照らし合わせる。アプローチ・着座の件数は渡さない
  const data = prompt.split('名前：')[1];
  assert.ok(data.includes('【直近7日】日報2回　獲得2件') && data.includes('【その前の7日】日報2回　獲得0件'), data);
  assert.ok(data.includes('所感：雨だったけど粘った') && data.includes('近見　獲得0件　所感：（なし）'), data);
  assert.ok(!data.includes('アプローチ') && !data.includes('着座') && !data.includes('キャッチ'), 'アプローチ・着座の件数は渡さない');
  assert.ok(prompt.includes('【今月のMNP】3件 / 目標 5件'), prompt);
  // 目標までのペースは GAS で計算して渡す。アドバイスは具体的に・食い違いは指摘してよい
  assert.ok(/【目標までのペース】あと2件。(残り\d+日で 1日あたり[\d.]+件|きょうが月末)/.test(prompt), prompt);
  assert.ok(prompt.includes('照らし合わせた、総合的なアドバイス') && prompt.includes('はっきり指摘・否定してよい') && prompt.includes('人を責める'), 'アドバイスのルール');
  assert.ok(!prompt.includes('鈴木') && !prompt.includes('毛利'), 'ほかの人の日報は渡さない');
  assert.ok(gemini[0].url.includes('gemini-2.5-flash'));
  // 2回目は残したものを返す（AIは呼ばない）
  c = get({ action: 'cheer', name: '山内' });
  assert.deepStrictEqual([c.text, c.grow, c.hint], ['下通で キャッチ30件、すごいがんばり', '成約率 先週より 上がったね', '着座に つなげると もっと伸びそう']);
  assert.strictEqual(gemini.length, 1);
  assert.deepStrictEqual(ss.getSheetByName('ひとこと')._grid[0], ['作った日時', '名前', '日報のキー', '日付', 'ひとこと', 'のびた', 'ヒント']);
  assert.strictEqual(ss.getSheetByName('ひとこと')._grid.length, 2, '見出し＋1行');
  // 伸びもヒントもない日は空のまま
  post({ requestId: 'c5', name: '鈴木', date: today, carrier: 'Softbank', place: '下通', approaches: [], gains: [], target: 0, prev: 0, note: '' });
  geminiQueue.push({ text: JSON.stringify({ cheer: 'きょうも おつかれさま', grow: '', hint: '' }) });
  c = get({ action: 'cheer', name: '鈴木' });
  assert.deepStrictEqual([c.text, c.grow, c.hint], ['きょうも おつかれさま', '', '']);
  // JSON で返らなかったときは、全体をひとことに
  post({ requestId: 'c6', name: '福山', date: today, carrier: 'au', place: '下通', approaches: [], gains: [], target: 0, prev: 0, note: '' });
  geminiQueue.push({ text: 'ゆっくり やすんでね' });
  assert.strictEqual(get({ action: 'cheer', name: '福山' }).text, 'ゆっくり やすんでね');
  // 古い日報（4日前が最新）の人にはしゃべらない
  post({ requestId: 'c3', name: '山口', date: old, carrier: 'au', place: '下通', approaches: [], gains: [], target: 0, prev: 0, note: '' });
  assert.strictEqual(get({ action: 'cheer', name: '山口' }).text, '');
  // 全部のモデルが失敗 → 空。しばらくは呼び直さない
  post({ requestId: 'c4', name: '毛利', date: today, carrier: 'Softbank', place: '下通', approaches: [], gains: [], target: 0, prev: 0, note: '' });
  geminiQueue.push({ code: 500 }, { code: 503 }, { code: 429 });
  const before = gemini.length;
  assert.strictEqual(get({ action: 'cheer', name: '毛利' }).text, '');
  assert.strictEqual(gemini.length - before, 3, '3つのモデルを順に試す');
  assert.strictEqual(get({ action: 'cheer', name: '毛利' }).text, '');
  assert.strictEqual(gemini.length - before, 3, '失敗したばかりなので呼ばない');
  assert.strictEqual(get({}).cheer, true);
  // 長すぎるひとことは切る・内訳の読み取り
  assert.strictEqual(cleanCheer_('あ'.repeat(70)).length, 60);
  // のびたも文の終わりで切る・長めまで残す
  assert.strictEqual(cleanCheer_('前は家族に確認しなきゃと帰っちゃうことがあったけど、きょうは家族の携帯のことまで聞けたね', GROW_MAX_LEN), '前は家族に確認しなきゃと帰っちゃうことがあったけど、きょうは家族の携帯のことまで聞けたね');
  // かぎかっこ：全体をくくっているときだけ外す
  assert.strictEqual(cleanCheer_('「外現場で苦労した」と11/22に書いているね。', HINT_MAX_LEN), '「外現場で苦労した」と11/22に書いているね。');
  assert.strictEqual(cleanCheer_('「すごいね」'), 'すごいね');
  assert.strictEqual(cleanCheer_('"いいね"'), 'いいね');
  // アドバイスは文の終わりで切る（途中で「…」にしない）
  var longHint = 'い'.repeat(100) + '。' + 'う'.repeat(150) + '。';
  assert.strictEqual(cleanCheer_(longHint, HINT_MAX_LEN), 'い'.repeat(100) + '。');
  assert.strictEqual(cleanCheer_('え'.repeat(150), HINT_MAX_LEN), 'え'.repeat(150));
  assert.strictEqual(cleanCheer_('お'.repeat(250), HINT_MAX_LEN), 'お'.repeat(199) + '…');
  assert.strictEqual(cleanCheer_('か。' + 'き'.repeat(250), HINT_MAX_LEN), 'か。' + 'き'.repeat(197) + '…');
  assert.deepStrictEqual(JSON.parse(JSON.stringify(parseBreakdown_('キャッチ6 / 店内2 / キャッチ1'))), { キャッチ: 7, 店内: 2 });
}
// ─── 月末の締め：まだの人への声かけ（通行証・10分に1回・メンションは鳴らさない） ───
{
  props.DISCORD_WEBHOOK = 'https://discord.com/api/webhooks/9/day';
  const t2 = teamToken_('kaeta');
  assert.strictEqual(post({ action: 'remind', token: 'nise', text: 'x' }).expired, true, '通行証がちがえば送らない');
  assert.strictEqual(post({ action: 'remind', token: t2, text: '  ' }).ok, false, '空の文は送らない');
  const n = posts.length;
  const rr = post({ action: 'remind', token: t2, text: '【11月の締め】\n・鈴木さん：稼働費 3日分 @everyone' });
  assert.deepStrictEqual(rr, { ok: true, discord: true });
  const pd = posts[posts.length - 1];
  assert.strictEqual(pd.url, props.DISCORD_WEBHOOK);
  assert.deepStrictEqual(pd.body.allowed_mentions, { parse: [] }, '@everyone は鳴らさない');
  assert.ok(pd.body.content.startsWith('【11月の締め】'));
  assert.strictEqual(post({ action: 'remind', token: t2, text: 'もう一回' }).tooSoon, true);
  assert.strictEqual(posts.length, n + 1, '10分に1回');
}

// ─── 1年のふりかえり（日報：日数・現場・アプローチ・獲得／獲得報告：合計・多い商品／AIのひとこと） ───
{
  props.GEMINI_API_KEY = 'k';
  const rep = (id, date, place, ap, mnp) => post({ requestId: id, name: '年末太郎', date, carrier: 'au', place, approaches: [{ label: 'キャッチ', count: ap }], gains: mnp ? [{ label: 'auMNP', count: mnp }] : [], target: 0, prev: 0, note: '' });
  rep('y1', '2026-03-02', '下通', 10, 1);
  rep('y2', '2026-03-02', '下通', 12, 2);            // 同じ日の2回目（こちらを数える）
  rep('y3', '2026-05-10', ' イオン  熊本 ', 5, 0);
  rep('y4', '2026-06-01', '下通', 8, 3);
  rep('y5', '2025-12-30', '下通', 99, 9);            // 前の年
  post({ action: 'gain', requestId: 'yg1', name: '年末太郎', carrier: 'au', date: '2026-04-01', items: [{ label: 'auMNP', count: 2 }, { label: 'auでんき', count: 5 }] });
  post({ action: 'gain', requestId: 'yg2', name: '年末太郎', carrier: 'au', date: '2025-11-01', items: [{ label: 'auMNP', count: 9 }] });
  geminiQueue.push({ text: JSON.stringify({ text: '「1年で 下通に よく いったね😊\nらいねんも よろしく」' }) });
  const y = get({ action: 'yearReview', name: '年末太郎', year: '2026' });
  assert.strictEqual(y.ok, true);
  assert.strictEqual(y.days, 3);
  assert.strictEqual(y.first, '2026-03-02');
  assert.strictEqual(y.placeCount, 2);
  assert.deepStrictEqual(y.places, [{ name: '下通', days: 2 }, { name: 'イオン 熊本', days: 1 }]);
  assert.strictEqual(y.approach, 12 + 5 + 8);
  assert.strictEqual(y.gain, 2 + 0 + 3);
  assert.deepStrictEqual(y.bestDay, { date: '2026-06-01', gain: 3 });
  assert.strictEqual(y.gains.total, 7);
  assert.deepStrictEqual(y.gains.top, [{ carrier: 'au', label: 'auでんき', count: 5 }, { carrier: 'au', label: 'auMNP', count: 2 }]);
  assert.strictEqual(y.ai, '1年で 下通に よく いったねらいねんも よろしく');
  assert.ok(/行った現場：2か所/.test(gemini[gemini.length - 1].body.contents[0].parts[0].text), 'AIに数字を渡す');
  const g0 = gemini.length;
  assert.strictEqual(get({ action: 'yearReview', name: '年末太郎', year: '2026' }).ai, y.ai, '2回目はシートから');
  assert.strictEqual(gemini.length, g0, 'AIを呼び直さない');
  const none = get({ action: 'yearReview', name: 'だれも', year: '2026' });
  assert.ok(none.ok && none.days === 0 && none.ai === '' && gemini.length === g0, '何もない人にはAIを呼ばない');
  assert.strictEqual(get({ action: 'yearReview', name: '' }).ok, false);
}

// ─── 不具合の記録（シート・式にしない・続いたら管理用 Discord・一覧は通行証） ───
{
  assert.deepStrictEqual(post({ action: 'logError', who: '鈴木', page: '/nippou-form/', what: '(送信)', kind: 'timeout', msg: '時間がかかりすぎ', detail: '=HYPERLINK("x")\n2行目', ua: 'iPhone' }), { ok: true });
  const eg = ss.getSheetByName('エラー記録')._grid;
  assert.deepStrictEqual(eg[0], ['日時', '名前', '画面', 'していたこと', '種類', '案内', '詳しく', '端末']);
  assert.strictEqual(eg[1][6], "'=HYPERLINK(\"x\") 2行目", '式として読まれない・改行は空白');
  props.DISCORD_ADMIN_WEBHOOK = 'https://discord.com/api/webhooks/7/admin';
  const adm = () => posts.filter(x => x.url === props.DISCORD_ADMIN_WEBHOOK).length;
  post({ action: 'logError', who: '毛利', page: '/sieg-portal/', kind: 'google', msg: 'y' });
  assert.strictEqual(adm(), 0, '2件ではまだ知らせない');
  post({ action: 'logError', who: '毛利', page: '/sieg-portal/', kind: 'google', msg: 'z' });
  assert.strictEqual(adm(), 1, '10分に3件で知らせる');
  assert.ok(posts.filter(x => x.url === props.DISCORD_ADMIN_WEBHOOK)[0].body.content.includes('3件'));
  post({ action: 'logError', who: '毛利', page: '/sieg-portal/', kind: 'google', msg: 'w' });
  assert.strictEqual(adm(), 1, '1時間に1回まで');
  assert.strictEqual(post({ action: 'errors', token: 'nise' }).expired, true);
  const list = post({ action: 'errors', token: teamToken_('kaeta') });
  assert.ok(list.ok);
  assert.deepStrictEqual(list.items.map(x => x.msg), ['w', 'z', 'y', '時間がかかりすぎ'], '新しい順');
  assert.strictEqual(list.items[3].page, '/nippou-form/');
  for (let i = 0; i < 70; i++) post({ action: 'logError', msg: 'many' + i });
  assert.ok(eg.length <= 61, '10分に60件まで ' + eg.length);
}
// ─── 名簿（みんなの状況のメンバー） ───
{
  const T = teamToken_('kaeta');
  // シートが無いうちは、いままでの名簿を返す（シートは作らない）
  let g = get({ action: 'roster' });
  assert.strictEqual(g.ok, true); assert.strictEqual(g.seed, true);
  assert.deepStrictEqual(g.people.map(x => x.name), ['福山', '古川', '清田', '福岡', '山口', '鈴木', '毛利', '山内']);
  assert.deepStrictEqual(g.people[0], { name: '福山', group: 'staff', full: '福山康平', carrier: 'Softbank', tools: ['commute', 'keihi', 'team'], hidden: '' });
  assert.strictEqual(ss.getSheetByName('名簿'), null, '読むだけではシートを作らない');
  // 通行証が無いと直せない
  assert.strictEqual(post({ action: 'rosterSave', token: 'nise', op: 'add', person: { name: '田中', group: 'partner' } }).expired, true);
  assert.strictEqual(ss.getSheetByName('名簿'), null);
  // 追加：シートを作って、いままでの8人のあとに足す
  let r = post({ action: 'rosterSave', token: T, op: 'add', person: { name: ' 田中 ', group: 'partner', full: '田中一郎', carrier: 'au', tools: ['commute', 'commute', 'xx'] } });
  assert.strictEqual(r.ok, true, JSON.stringify(r));
  const sh = ss.getSheetByName('名簿');
  assert.deepStrictEqual(sh._grid[0], ['名前', '区分', 'フルネーム', 'キャリア', '稼働費', '経費精算', 'みんなの状況', '外した日', '直した日時']);
  assert.strictEqual(sh._grid.length, 10);
  assert.deepStrictEqual(sh._grid[9].slice(0, 8), ['田中', 'パートナー', '田中一郎', 'au', '○', '', '', '']);
  assert.ok(/^\d{4}-\d\d-\d\d \d\d:\d\d$/.test(sh._grid[9][8]));
  assert.deepStrictEqual(r.people[8], { name: '田中', group: 'partner', full: '田中一郎', carrier: 'au', tools: ['commute'], hidden: '' }, '同じものは1つ・知らないツールは入れない');
  assert.strictEqual(get({ action: 'roster' }).people.length, 9, '読み直しにすぐ出る（キャッシュを消す）');
  // 同じ内容を送り直したら、追加できたことにする。違う内容なら「もういます」
  r = post({ action: 'rosterSave', token: T, op: 'add', person: { name: '田中', group: 'partner', full: '田中一郎', carrier: 'au', tools: ['commute'] } });
  assert.strictEqual(r.ok, true); assert.strictEqual(r.already, true); assert.strictEqual(sh._grid.length, 10);
  r = post({ action: 'rosterSave', token: T, op: 'add', person: { name: '田中', group: 'staff' } });
  assert.strictEqual(r.ok, false); assert.ok(r.error.includes('もう名簿にいます'));
  // 名前のチェック
  assert.ok(post({ action: 'rosterSave', token: T, op: 'add', person: { name: '　' } }).error.includes('名前を入れて'));
  assert.ok(post({ action: 'rosterSave', token: T, op: 'add', person: { name: '=HYPERLINK(1)' } }).error.includes('使えない文字'));
  assert.ok(post({ action: 'rosterSave', token: T, op: 'add', person: { name: 'あいうえおかきくけこさ' } }).error.includes('10文字'));
  assert.strictEqual(post({ action: 'rosterSave', token: T, op: 'zzz', person: { name: '田中' } }).ok, false);
  // 直す：区分・キャリア・ツール。フルネームの式は外す
  r = post({ action: 'rosterSave', token: T, op: 'update', person: { name: '田中', group: 'staff', full: '=田中', carrier: 'Softbank', tools: ['keihi'] } });
  assert.deepStrictEqual(sh._grid[9].slice(0, 8), ['田中', '社員', '田中', 'Softbank', '', '○', '', '']);
  assert.strictEqual(post({ action: 'rosterSave', token: T, op: 'update', person: { name: 'いない人' } }).error.includes('見つかりません'), true);
  // 外す → 一覧に出さない印（記録は消さない）。戻す → 印を消す
  r = post({ action: 'rosterSave', token: T, op: 'hide', person: { name: '田中' } });
  assert.ok(r.ok); assert.ok(/^\d{4}-\d\d-\d\d$/.test(r.people[8].hidden), r.people[8].hidden);
  assert.strictEqual(sh._grid.length, 10, '行は消さない');
  assert.ok(post({ action: 'rosterSave', token: T, op: 'add', person: { name: '田中', group: 'staff' } }).error.includes('戻す'), '外した人の名前で追加はできない');
  r = post({ action: 'rosterSave', token: T, op: 'show', person: { name: '田中' } });
  assert.strictEqual(r.people[8].hidden, '');
  // みんなの状況を見る人が0人になる変更はしない
  r = post({ action: 'rosterSave', token: T, op: 'hide', person: { name: '福岡' } });
  assert.ok(r.ok);
  r = post({ action: 'rosterSave', token: T, op: 'update', person: { name: '福山', group: 'staff', full: '福山康平', carrier: 'Softbank', tools: ['commute', 'keihi'] } });
  assert.strictEqual(r.ok, false); assert.ok(r.error.includes('1人もいなくなります'));
  assert.strictEqual(post({ action: 'rosterSave', token: T, op: 'hide', person: { name: '福山' } }).ok, false);
  assert.deepStrictEqual(get({ action: 'roster' }).people[0].tools, ['commute', 'keihi', 'team'], '断ったときは変えない');
  post({ action: 'rosterSave', token: T, op: 'show', person: { name: '福岡' } });
  // シートを人が直接直しても読める（キャリアの大文字小文字・○以外の印）
  sh._grid[2][3] = 'SoftBank'; sh._grid[2][5] = 'あり';
  CacheService.getScriptCache().remove('roster:v1');
  g = get({ action: 'roster' });
  assert.strictEqual(g.people[1].carrier, 'Softbank'); assert.deepStrictEqual(g.people[1].tools, ['commute', 'keihi']);
  assert.strictEqual(g.seed, false);
}
// ─── 取引先・場所の一覧 ───
{
  const T = teamToken_('kaeta');
  CacheService.getScriptCache().remove('roster:v1');
  let g = get({ action: 'roster' });
  // シートが無いうちは、初めの一覧（今の取引先6つ・店舗稼働の場所）を返す
  assert.strictEqual(g.masterSeed, true);
  assert.deepStrictEqual(g.clients.map(c => c.name), ['南国殖産', 'ハイコム', 'サンコミ', 'ダル', 'NOON', '九情']);
  assert.deepStrictEqual(g.clients[0], { name: '南国殖産', carrier: 'docomo', place: '近見', color: '#2E7D32', hidden: '' });
  assert.ok(g.places.find(p => p.name === '健軍').shop && g.places.find(p => p.name === '下通').noParen);
  assert.strictEqual(ss.getSheetByName('取引先'), null, '読むだけではシートを作らない');
  // 通行証が無いと直せない
  assert.strictEqual(post({ action: 'masterSave', token: 'nise', kind: 'client', op: 'add', item: { name: 'X' } }).expired, true);
  // 取引先を足す・直す・外す
  let r = post({ action: 'masterSave', token: T, kind: 'client', op: 'add', item: { name: '新取引先', carrier: 'au', color: '#123456' } });
  assert.ok(r.ok, r.error);
  assert.deepStrictEqual(r.clients[6], { name: '新取引先', carrier: 'au', place: '', color: '#123456', hidden: '' });
  assert.strictEqual(ss.getSheetByName('取引先')._grid.length, 8, '見出し＋7');
  assert.ok(post({ action: 'masterSave', token: T, kind: 'client', op: 'add', item: { name: 'ダル' } }).error.includes('もうあります'));
  assert.ok(post({ action: 'masterSave', token: T, kind: 'client', op: 'add', item: { name: 'Y', color: 'red' } }).error.includes('色'));
  r = post({ action: 'masterSave', token: T, kind: 'client', op: 'update', key: '新取引先', item: { name: '新取引先2', carrier: 'xx', color: '' } });
  assert.deepStrictEqual([r.clients[6].name, r.clients[6].carrier], ['新取引先2', ''], '名前を変えられる・知らないキャリアは空');
  r = post({ action: 'masterSave', token: T, kind: 'client', op: 'hide', item: { name: '新取引先2' } });
  assert.ok(/^\d{4}-\d\d-\d\d$/.test(r.clients[6].hidden));
  // 稼働費の場所から取り込む（SB光の森 → 光の森・別名・店舗稼働。自宅・削除した場所は入れない。もうある場所は住所だけ入れる）
  const k = Sheet('稼働場所', []);
  [['短縮名', '住所'], ['SB光の森', '熊本県菊池郡菊陽町光の森2丁目3-1'], ['福山自宅', '熊本県玉名郡'], ['ドコモ近見', '熊本県熊本市南区近見7丁目'],
   ['au・SB下通', '熊本県熊本市中央区手取本町4-12'], ['DCMダイキ', '〒862-0000 熊本県熊本市東区御領2丁目']].forEach(x => k.appendRow(x));
  const pp = Sheet('場所', []);
  [['ID', '名前', '場所名', '住所', '作成日時', '削除'], ['a', '毛利', 'SB松橋', '熊本県宇城市松橋町', '', ''], ['b', '鈴木', 'けした場所', '熊本県', '', '削除']].forEach(x => pp.appendRow(x));
  // 取り込み元は別のブック（日報のブックとは別）
  const src = { getSheetByName: n => ({ '稼働場所': k, '場所': pp })[n] || null };
  const realOpen = SpreadsheetApp.openById;
  SpreadsheetApp.openById = id => id === SHEET_ID ? ss : src;
  r = post({ action: 'masterImport', token: T });
  assert.ok(r.ok, r.error);
  const pl = n => r.places.find(p => p.name === n);
  assert.deepStrictEqual([pl('光の森').address, pl('光の森').alias, pl('光の森').shop], ['熊本県菊池郡菊陽町光の森2丁目3-1', ['SB光の森'], true]);
  assert.deepStrictEqual([pl('近見').address, pl('近見').shop], ['熊本県熊本市南区近見7丁目', true]);
  assert.deepStrictEqual([pl('下通').address, pl('下通').noParen, pl('下通').alias], ['熊本県熊本市中央区手取本町4-12', true, ['au・SB下通']]);
  assert.deepStrictEqual([pl('DCMダイキ').address, pl('DCMダイキ').shop], ['熊本県熊本市東区御領2丁目', false], '〒は取る・会社名の無い場所は店舗稼働にしない');
  assert.strictEqual(pl('松橋').address, '熊本県宇城市松橋町');
  assert.ok(!r.places.some(p => /自宅/.test(p.name) || p.name === 'けした場所'));
  assert.ok(r.imported.read >= 5);
  // もう一度取り込んでも増えない
  const n1 = r.places.length;
  r = post({ action: 'masterImport', token: T });
  assert.strictEqual(r.places.length, n1);
  // 場所を直す（別名は、区切り）
  r = post({ action: 'masterSave', token: T, kind: 'place', op: 'update', key: 'DCMダイキ', item: { name: 'DCMダイキ', alias: 'ダイキ、DCMダイキ東', address: '熊本県熊本市東区御領2丁目', store: '上熊本' } });
  assert.deepStrictEqual([pl('DCMダイキ') && r.places.find(p => p.name === 'DCMダイキ').alias, r.places.find(p => p.name === 'DCMダイキ').store], [['ダイキ', 'DCMダイキ東'], '上熊本']);
  assert.strictEqual(get({ action: 'roster' }).masterSeed, false);
  SpreadsheetApp.openById = realOpen;
}
console.log('nippou-gas: all tests passed');
