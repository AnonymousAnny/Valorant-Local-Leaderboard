// End-to-end test: runs Code.gs against mocked Apps Script services and a fake HenrikDev API.
// Usage: node test/run.js
const fs = require('fs');
const path = require('path');
const vm = require('vm');
const assert = require('assert');

function makeSheet(rows, cols = 12) {
  const grid = rows.map(r => { const a = r.slice(); while (a.length < cols) a.push(''); return a; });
  const bg = grid.map(() => new Array(cols).fill(null));
  const hidden = new Set();
  const sheet = {
    grid, bg, hidden, hideCalls: 0,
    getLastRow: () => grid.length,
    getLastColumn: () => cols,
    isColumnHiddenByUser: c => hidden.has(c),
    hideColumns: c => { hidden.add(c); sheet.hideCalls++; },
    getRange(r, c, nr = 1, nc = 1) {
      assert(r >= 1 && nr >= 1 && r - 1 + nr <= grid.length, `bad range r=${r} nr=${nr}`);
      const slice = src => src.slice(r - 1, r - 1 + nr).map(row => row.slice(c - 1, c - 1 + nc));
      const put = (dst, vals) => {
        assert.strictEqual(vals.length, nr);
        vals.forEach((row, i) => {
          assert.strictEqual(row.length, nc);
          row.forEach((v, j) => { dst[r - 1 + i][c - 1 + j] = v; });
        });
      };
      return {
        getValues: () => slice(grid),
        getFormulas: () => slice(grid).map(r => r.map(v => (typeof v === 'string' && v.startsWith('=') ? v : ''))),
        getBackgrounds: () => slice(bg).map(r => r.map(v => v || '#ffffff')),
        setValues: v => put(grid, v),
        setBackgrounds: v => put(bg, v),
        sort(specs) {
          const part = grid.slice(r - 1, r - 1 + nr);
          const bgp = bg.slice(r - 1, r - 1 + nr);
          const cmp = (a, b) => {
            for (const spec of specs) {
              const idx = spec.column - c;
              const x = part[a][idx], y = part[b][idx];
              const bx = x === '', by = y === '';
              if (bx !== by) return bx ? 1 : -1; // blanks last
              if (bx && by) continue;
              if (x !== y) return spec.ascending ? x - y : y - x;
            }
            return 0;
          };
          const order = part.map((_, i) => i).sort(cmp);
          order.forEach((o, i) => { grid[r - 1 + i] = part[o]; bg[r - 1 + i] = bgp[o]; });
        },
      };
    },
  };
  return sheet;
}

function run({ rows, props, responder }) {
  const logs = [], fetched = [], sleeps = [], triggers = [];
  const sheetObj = makeSheet(rows);
  const ctx = {
    console,
    Logger: { log: m => logs.push(m) },
    PropertiesService: { getScriptProperties: () => ({ getProperty: k => (k in props ? props[k] : null), getProperties: () => Object.assign({}, props) }) },
    LockService: { getScriptLock: () => ({ tryLock: () => true, releaseLock() {} }) },
    Utilities: { sleep: ms => sleeps.push(ms) },
    SpreadsheetApp: {
      getUi: () => ({
        createMenu: name => {
          const m = { addItem: (...a) => { ctx.menu.push([name, ...a]); return m; }, addToUi() {} };
          return m;
        },
        showSidebar: h => { ctx.sidebar = h; },
      }),
      getActiveSpreadsheet: () => ({
        getSheets: () => [sheetObj, { other: true }],
        getSheetByName: n => (n === 'Board' ? sheetObj : null),
      }),
    },
    ScriptApp: {
      getProjectTriggers: () => triggers,
      newTrigger: fn => {
        const t = { forSpreadsheet: () => t, onOpen: () => t, create: () => triggers.push({ getHandlerFunction: () => fn }) };
        return t;
      },
    },
    UrlFetchApp: {
      fetchAll: reqs => reqs.map(r => {
        fetched.push(r);
        const { code, body, headers } = responder(r);
        return { getResponseCode: () => code, getHeaders: () => headers || {}, getContentText: () => (typeof body === 'string' ? body : JSON.stringify(body)) };
      }),
    },
    HtmlService: { createHtmlOutput: html => ({ html, setTitle(t) { this.title = t; return this; } }) },
    menu: [],
  };
  vm.createContext(ctx);
  vm.runInContext(fs.readFileSync(path.join(__dirname, '..', 'Code.gs'), 'utf8'), ctx);
  return { ctx, sheet: sheetObj, logs, fetched, sleeps, triggers };
}

const player = (tier, elo, rr, games = 0, peak = 'Radiant', season = 'e9a1') => ({
  code: 200,
  body: {
    status: 200,
    data: {
      current_data: {
        currenttierpatched: tier, ranking_in_tier: rr, elo, games_needed_for_rating: games,
        images: { large: 'https://img/' + tier.replace(/ /g, '_') + '.png' },
      },
      highest_rank: { patched_tier: peak, season },
    },
  },
});
const db = {
  'Ann/1111': player('Gold 2', 1200, 40),
  'Bob/2222': player('Radiant', 2700, 90),
  'Cy Z/3333': player('Silver 1', 800, 10, 1), // placement -> grey; name with space
  'Dee/4444': { code: 404, body: { status: 404, errors: [{ message: 'Player not found' }] } },
  'Eve/5555': { code: 200, body: { status: 200, data: {} } },
  'Flo/6666': { code: 502, body: '<html>bad gateway</html>' },
};
const responder = r => {
  const m = r.url.match(/mmr\/([^/]+)\/([^/]+)\/([^/]+)$/);
  assert(m, 'url shape ' + r.url);
  assert(/^KEY\d*$/.test(r.headers.Authorization), 'unknown key ' + r.headers.Authorization);
  const key = decodeURIComponent(m[2]) + '/' + decodeURIComponent(m[3]);
  return db[key] || { code: 404, body: { status: 404 } };
};

let n = 0;
const ok = name => console.log('ok ' + (++n) + ' - ' + name);

// 1. Full refresh
{
  const rows = [
    ['Name', 'Tag', '', '', 'Label'],
    ['Ann', '#1111', '', '', 'A'],
    ['Cy Z', '3333', '', '', 'C \u{1F451}'], // tag without '#', stale crown
    ['Bob', '#2222', '', '', 'B'],
    ['Dee', '#4444', '', '', 'D'],
    ['Eve', '#5555', '', '', 'E'],
    ['', '', '', '', 'blank'],
    ['Gus', '', '', '', 'G'], // missing tag
  ];
  const t = run({ rows, props: { API_KEY: 'KEY' }, responder });
  t.ctx.refreshValorantData();
  const g = t.sheet.grid;
  assert.deepStrictEqual(g.slice(1, 4).map(r => r[0]), ['Bob', 'Ann', 'Cy Z'], 'sorted by ELO desc');
  ok('sorted by ELO descending');
  assert.strictEqual(g[1][4], '\u{1F451} B');
  assert.strictEqual(g[3][4], 'C');
  assert.strictEqual(g.filter(r => String(r[4]).includes('\u{1F451}')).length, 1);
  ok('exactly one crown, on top row');
  assert.strictEqual(g[1][5], 'Radiant');
  assert.strictEqual(g[1][6], '=IMAGE("https://img/Radiant.png")');
  assert.strictEqual(g[1][7], 90);
  assert.strictEqual(g[1][8], 2700);
  assert.strictEqual(g[1][9], 'Radiant');
  assert.strictEqual(g[1][11], 0);
  ok('rank/image/RR/ELO/peak written');
  assert(t.fetched.some(r => r.url.endsWith('/ap/Cy%20Z/3333')), 'encoded url');
  assert(!t.fetched.some(r => r.url.includes('Gus')));
  ok('names URL-encoded, "#"-less tags handled, invalid rows not fetched');
  assert.strictEqual(t.sheet.bg[3][5], '#D3D3D3');
  assert.strictEqual(t.sheet.bg[1][5], null);
  ok('placement rows greyed, backgrounds follow sort');
  const byName = Object.fromEntries(g.slice(1).map(r => [r[0], r[5]]));
  assert.strictEqual(byName.Dee, 'API Error: 404 (Player not found)');
  assert.strictEqual(byName.Eve, 'API Error: unexpected response');
  assert.strictEqual(byName.Gus, 'Missing name or tag');
  assert.strictEqual(byName[''], '');
  ok('error rows report clear messages; blank row cleared');
  assert(t.sheet.hidden.has(9) && t.sheet.hidden.has(12));
  ok('columns I and L hidden');
  t.ctx.refreshValorantData();
  assert.strictEqual(t.sheet.hideCalls, 2);
  assert.strictEqual(t.sheet.grid.filter(r => String(r[4]).includes('\u{1F451}')).length, 1);
  ok('second run is idempotent (no double hide / crown)');
}

// 2. Rate limit retry then success; persistent 5xx ends in error
{
  let calls = 0;
  const rows = [['N', 'T'], ['Ann', '#1111'], ['Flo', '#6666']];
  const t = run({
    rows, props: { API_KEY: 'KEY' },
    responder: r => (r.url.includes('Ann') && ++calls < 3 ? { code: 429, body: { status: 429 } } : responder(r)),
  });
  t.ctx.refreshValorantData();
  assert.strictEqual(calls, 3);
  assert.strictEqual(t.sleeps.length, 3); // Ann: 2 retry rounds, Flo (502) keeps retrying to the 4th attempt
  const by = Object.fromEntries(t.sheet.grid.slice(1).map(r => [r[0], r[5]]));
  assert.strictEqual(by.Ann, 'Gold 2');
  assert.match(by.Flo, /^JSON Parse Error \(HTTP 502\)/);
  assert.deepStrictEqual(t.sleeps, [10000, 20000, 30000]);
  ok('429 retried with linear backoff; non-JSON 502 reported');
}

// 3. Batching beyond 20 rows
{
  const rows = [['N', 'T']];
  for (let i = 0; i < 45; i++) rows.push(['P' + i, '#0001']);
  let max = 0;
  const t = run({ rows, props: { API_KEY: 'KEY' }, responder: () => player('Iron 1', 100, 1) });
  const orig = t.ctx.UrlFetchApp.fetchAll;
  t.ctx.UrlFetchApp.fetchAll = reqs => { max = Math.max(max, reqs.length); return orig(reqs); };
  t.ctx.refreshValorantData();
  assert.strictEqual(t.fetched.length, 45);
  assert(max <= 10);
  ok('45 players fetched in batches of <=10');
}

// 4. Edge cases
{
  const t = run({ rows: [['N', 'T']], props: { API_KEY: 'KEY' }, responder });
  t.ctx.refreshValorantData();
  assert.strictEqual(t.fetched.length, 0);
  ok('header-only sheet does nothing (no crash)');

  const one = run({ rows: [['N', 'T'], ['Ann', '#1111']], props: { API_KEY: 'KEY' }, responder });
  one.ctx.refreshValorantData();
  assert.strictEqual(one.sheet.grid[1][5], 'Gold 2');
  ok('single-player sheet works');

  const nokey = run({ rows: [['N', 'T'], ['Ann', '#1111']], props: {}, responder });
  assert.throws(() => nokey.ctx.refreshValorantData(), /API key not found/);
  ok('missing API key throws clear error');

  const named = run({ rows: [['N', 'T'], ['Ann', '#1111']], props: { API_KEY: 'KEY', SHEET_NAME: 'Board', REGION: 'eu' }, responder });
  named.ctx.refreshValorantData();
  assert(named.fetched[0].url.includes('/mmr/eu/'));
  ok('SHEET_NAME and REGION properties honoured');

  const bad = run({ rows: [['N', 'T']], props: { API_KEY: 'KEY', SHEET_NAME: 'Nope' }, responder });
  assert.throws(() => bad.ctx.refreshValorantData(), /not found/);
  ok('unknown SHEET_NAME throws clear error');

  const fe = run({ rows: [['N', 'T'], ['Ann', '#1111']], props: { API_KEY: 'KEY' }, responder });
  fe.ctx.UrlFetchApp.fetchAll = () => { throw new Error('boom'); };
  fe.ctx.refreshValorantData();
  assert.strictEqual(fe.sheet.grid[1][5], 'Fetch Error: boom');
  ok('fetchAll exception becomes per-row error');
}

// 6. Rate-limit header drives the wait; transient errors keep previous data, definitive ones don't
{
  const rows = [
    ['N', 'T', '', '', 'L', 'Gold 2', '=IMAGE(\"https://img/old.png\")', 40, 1200, 'Radiant', 'e9a1', 0],
    ['Ann', '#1111', '', '', 'A', 'Gold 2', '=IMAGE(\"https://img/old.png\")', 40, 1200, 'Radiant', 'e9a1', 0],
    ['Dee', '#4444', '', '', 'D', 'Iron 1', '=IMAGE(\"https://img/old.png\")', 5, 100, 'Iron 1', 'e1a1', 0],
  ];
  const t = run({
    rows, props: { API_KEY: 'KEY' },
    responder: r => (r.url.includes('Ann')
      ? { code: 429, body: { status: 429 }, headers: { 'X-RateLimit-Reset': '7' } }
      : responder(r)),
  });
  t.ctx.refreshValorantData();
  assert.deepStrictEqual(t.sleeps, [7500, 7500, 7500]);
  ok('x-ratelimit-reset header honoured for retry delay');
  const by = Object.fromEntries(t.sheet.grid.slice(1).map(r => [r[0], r]));
  assert.strictEqual(by.Ann[5], 'Gold 2');
  assert.strictEqual(by.Ann[6], '=IMAGE("https://img/old.png")');
  assert.strictEqual(by.Ann[8], 1200);
  ok('rate-limited row keeps previous rank, image and ELO');
  assert.strictEqual(by.Dee[5], 'API Error: 404 (Player not found)');
  assert.strictEqual(by.Dee[8], '');
  ok('definitive 404 overwrites stale data with the error');
}

// 7. Multiple API keys: round-robin, rotation on 429, parsing, dedupe
{
  const rows = [['N', 'T'], ['Ann', '#1111'], ['Bob', '#2222'], ['Ann', '#1111'], ['Bob', '#2222']];
  const t = run({ rows, props: { API_KEY: 'KEY, KEY2', API_KEY3: 'KEY3 KEY2' }, responder });
  t.ctx.refreshValorantData();
  assert.deepStrictEqual(t.fetched.map(r => r.headers.Authorization), ['KEY', 'KEY2', 'KEY3', 'KEY']);
  ok('keys parsed from API_KEY and API_KEY_n, deduped, assigned round-robin');

  // key "KEY" is rate limited, KEY2 is fine: retry must switch key
  const seen = [];
  const t2 = run({
    rows: [['N', 'T'], ['Ann', '#1111'], ['Bob', '#2222']],
    props: { API_KEY: 'KEY', API_KEY_2: 'KEY2' },
    responder: r => {
      seen.push(r.url.split('/').slice(-2).join('/') + '@' + r.headers.Authorization);
      return r.headers.Authorization === 'KEY' ? { code: 429, body: { status: 429 } } : responder(r);
    },
  });
  t2.ctx.refreshValorantData();
  assert.deepStrictEqual(seen, ['Ann/1111@KEY', 'Bob/2222@KEY2', 'Ann/1111@KEY2']);
  assert.strictEqual(t2.sheet.grid[1][0] === 'Bob' ? t2.sheet.grid[1][5] : t2.sheet.grid[2][5], 'Radiant');
  assert(t2.sheet.grid.every(r => !String(r[5]).startsWith('API Error: 429')));
  assert.strictEqual(t2.sleeps.length, 1);
  ok('429 on one key: retry rotates to the other key and succeeds');
}

// 8. The user's actual property names: API_KEY1 and API_KEY2 (no plain API_KEY)
{
  const t = run({ rows: [['N', 'T'], ['Ann', '#1111'], ['Bob', '#2222']], props: { API_KEY1: 'KEY1', API_KEY2: 'KEY2', SHEET_NAME: 'Board' }, responder });
  t.ctx.refreshValorantData();
  assert.deepStrictEqual(t.fetched.map(r => r.headers.Authorization), ['KEY1', 'KEY2']);
  ok('API_KEY1 / API_KEY2 property names both used');
}

// 9. Inactive/unrated accounts: greyed (after sort) and placed below all active accounts
{
  const rows = [['N', 'T', '', '', 'L'], ['Ann', '#1111', '', '', 'A'], ['Bob', '#2222', '', '', 'B'], ['Cy Z', '#3333', '', '', 'C'], ['Una', '#7777', '', '', 'U']];
  const dbx = Object.assign({}, db, {
    'Cy Z/3333': player('Radiant', 2900, 5, 1),      // highest ELO but inactive (needs 1 game)
    'Una/7777': player('Unrated', 0, 0, 3),           // unrated
  });
  const t = run({ rows, props: { API_KEY: 'KEY' }, responder: r => {
    const m = r.url.match(/mmr\/[^/]+\/([^/]+)\/([^/]+)$/);
    return dbx[decodeURIComponent(m[1]) + '/' + decodeURIComponent(m[2])];
  } });
  t.ctx.refreshValorantData();
  assert.deepStrictEqual(t.sheet.grid.slice(1).map(r => r[0]), ['Bob', 'Ann', 'Cy Z', 'Una']);
  ok('inactive account with highest ELO sorts below active accounts');
  assert.deepStrictEqual(t.sheet.bg.slice(1).map(r => r[5]), [null, null, '#D3D3D3', '#D3D3D3']);
  assert.deepStrictEqual(t.sheet.bg.slice(1).map(r => r[7]), [null, null, '#D3D3D3', '#D3D3D3']);
  ok('inactive and unrated rows greyed on F..H, active rows not');
  t.ctx.refreshValorantData();
  assert.deepStrictEqual(t.sheet.bg.slice(1).map(r => r[5]), [null, null, '#D3D3D3', '#D3D3D3']);
  ok('grey styling stable across refreshes');
}

// 10. Public layout: no credential columns, so Owner is column C and outputs are D..J
{
  const rows = [
    ['Riot ID', 'Tag', 'Owner'],
    ['Ann', '#1111', 'A'],
    ['Bob', '#2222', 'B'],
    ['Cy Z', '#3333', 'C'],
  ];
  const dbx = Object.assign({}, db, { 'Cy Z/3333': player('Radiant', 2900, 5, 1) });
  const t = run({ rows, props: { API_KEY: 'KEY' }, responder: r => {
    const m = r.url.match(/mmr\/[^/]+\/([^/]+)\/([^/]+)$/);
    return dbx[decodeURIComponent(m[1]) + '/' + decodeURIComponent(m[2])];
  } });
  t.ctx.refreshValorantData();
  const g = t.sheet.grid;
  assert.deepStrictEqual(g.slice(1).map(r => r[0]), ['Bob', 'Ann', 'Cy Z']);
  assert.strictEqual(g[1][2], '\u{1F451} B');
  assert.strictEqual(g[1][3], 'Radiant'); // D = rank
  assert.strictEqual(g[1][6], 2700); // G = ELO
  assert.strictEqual(g[3][9], 1); // J = games needed
  ok('Owner header anchors the layout (crown in C, ELO in G, games in J)');
  assert.deepStrictEqual([...t.sheet.hidden].sort((a, b) => a - b), [7, 10]);
  assert.deepStrictEqual(t.sheet.bg.slice(1).map(r => r[3]), [null, null, '#D3D3D3']);
  ok('hidden columns and grey follow the detected layout');
}

// 5. Menu and trigger
{
  const t = run({ rows: [['N', 'T']], props: {}, responder });
  t.ctx.onOpen();
  assert.strictEqual(t.ctx.menu.length, 3);
  assert(t.ctx.sidebar && t.ctx.sidebar.title === 'Discord' && t.ctx.sidebar.html.includes('discord.com/widget'));
  t.ctx.installAutoRefreshTrigger();
  t.ctx.installAutoRefreshTrigger();
  assert.strictEqual(t.triggers.length, 1);
  t.ctx.autoRefresh(); // no key -> logged, not thrown
  assert(t.logs.some(l => /autoRefresh failed/.test(l)));
  ok('menus + Discord sidebar built, trigger installed once, autoRefresh swallows errors');
}
console.log('\nAll ' + n + ' checks passed');
