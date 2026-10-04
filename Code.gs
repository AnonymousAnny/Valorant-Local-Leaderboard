// Valorant Local Leaderboard
// Sheet layout: A = name, B = tag (e.g. #1234), then an "Owner" column (the crown goes here), followed by
// 7 output columns: rank, rank image, RR, ELO (hidden), peak rank, peak season, placement games (hidden).
// The Owner column is found by its header in row 1; if there is none it is assumed to be column E.

const DEFAULT_OWNER_COLUMN = 5;
const OUT_COUNT = 7;
const CROWN = '👑';
const DEFAULT_REGION = 'ap';
const BATCH_SIZE = 10;
const MAX_ATTEMPTS = 4;
const RETRY_DELAY_MS = 10000; // fallback when the API gives no reset header
const MAX_RETRY_DELAY_MS = 60000;
const UNRANKED_GREY = '#D3D3D3';

// Simple trigger: only builds the menus and opens the sidebar (simple triggers cannot call UrlFetchApp).
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Discord')
    .addItem('Join Discord Server', 'showSidebar')
    .addToUi();
  ui.createMenu('Valorant')
    .addItem('Refresh Valorant Stats', 'refreshValorantData')
    .addItem('Enable auto-refresh on open', 'installAutoRefreshTrigger')
    .addToUi();
  showSidebar();
}

function showSidebar() {
  const html = HtmlService.createHtmlOutput(
    '<iframe src="https://discord.com/widget?id=695645680189833256&theme=dark" width="280" height="680" ' +
    'allowtransparency="true" frameborder="0" ' +
    'sandbox="allow-popups allow-popups-to-escape-sandbox allow-same-origin allow-scripts"></iframe>')
    .setTitle('Discord');
  SpreadsheetApp.getUi().showSidebar(html);
}

// Installable trigger target: refreshes when the spreadsheet is opened.
function autoRefresh() {
  try {
    refreshValorantData();
  } catch (error) {
    Logger.log(`autoRefresh failed: ${error}`);
  }
}

// Creates the installable on-open trigger once (needs to be run by the sheet owner).
function installAutoRefreshTrigger() {
  const exists = ScriptApp.getProjectTriggers().some(t => t.getHandlerFunction() === 'autoRefresh');
  if (!exists) {
    ScriptApp.newTrigger('autoRefresh')
      .forSpreadsheet(SpreadsheetApp.getActiveSpreadsheet())
      .onOpen()
      .create();
  }
}

function getLayout(sheet) {
  const header = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const owner = header.findIndex(cell => String(cell).trim().toLowerCase() === 'owner');
  const crown = owner >= 0 ? owner + 1 : DEFAULT_OWNER_COLUMN;
  const first = crown + 1;
  return { NAME: 1, TAG: 2, CROWN: crown, OUT_FIRST: first, ELO: first + 3, GAMES: first + 6, OUT_COUNT: OUT_COUNT };
}

function getLeaderboardSheet() {
  const props = PropertiesService.getScriptProperties();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const name = props.getProperty('SHEET_NAME');
  if (name) {
    const sheet = ss.getSheetByName(name);
    if (!sheet) throw new Error(`Sheet "${name}" not found.`);
    return sheet;
  }
  return ss.getSheets()[0];
}

// Keys come from any script property named API_KEY, API_KEY1, API_KEY_2, ... (each value may also hold
// several keys separated by commas/spaces/newlines).
function getApiKeys(props) {
  const all = props.getProperties();
  const names = Object.keys(all)
    .filter(name => /^API_KEY_?\d*$/.test(name))
    .sort((a, b) => Number(a.replace(/\D/g, '') || 0) - Number(b.replace(/\D/g, '') || 0));
  const keys = [];
  names.forEach(name => String(all[name]).split(/[\s,;]+/).forEach(key => {
    if (key && keys.indexOf(key) === -1) keys.push(key);
  }));
  return keys;
}

function refreshValorantData() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) {
    throw new Error('Another refresh is already running.');
  }
  try {
    const props = PropertiesService.getScriptProperties();
    const apiKeys = getApiKeys(props);
    if (!apiKeys.length) throw new Error('API key not found.');
    const region = props.getProperty('REGION') || DEFAULT_REGION;

    const sheet = getLeaderboardSheet();
    const L = getLayout(sheet);
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return; // header only, nothing to do
    const rowCount = lastRow - 1;

    const players = sheet.getRange(2, L.NAME, rowCount, 2).getValues();
    const requests = [];
    const results = players.map(() => null);

    players.forEach(([name, tag], i) => {
      const cleanName = String(name == null ? '' : name).trim();
      const cleanTag = String(tag == null ? '' : tag).trim().replace(/^#/, '');
      if (!cleanName && !cleanTag) {
        results[i] = { blank: true };
      } else if (!cleanName || !cleanTag) {
        results[i] = { error: 'Missing name or tag' };
      } else {
        requests.push({
          index: i,
          url: `https://api.henrikdev.xyz/valorant/v2/mmr/${encodeURIComponent(region)}/` +
               `${encodeURIComponent(cleanName)}/${encodeURIComponent(cleanTag)}`,
        });
      }
    });

    batchFetch(requests, apiKeys).forEach(r => { results[r.index] = r; });

    writeResults(sheet, L, results);
    sortSheetByELO(sheet, L);
    styleInactiveRows(sheet, L);
    hideColumns(sheet, [L.ELO, L.GAMES]);
    manageCrownSymbols(sheet, L);
  } catch (error) {
    Logger.log(`Error in refreshValorantData: ${error}`);
    throw error;
  } finally {
    lock.releaseLock();
  }
}

// Fetches in chunks, spreading requests across the API keys (round-robin) so each key's rate limit
// is shared. A retry moves a request to the next key. Retries rate-limited (429) / transient (5xx) responses.
function batchFetch(requests, apiKeys) {
  const done = [];
  for (let start = 0; start < requests.length; start += BATCH_SIZE) {
    let pending = requests.slice(start, start + BATCH_SIZE);
    for (let attempt = 1; pending.length && attempt <= MAX_ATTEMPTS; attempt++) {
      let responses;
      try {
        responses = UrlFetchApp.fetchAll(pending.map(r => ({
          url: r.url,
          method: 'get',
          headers: { accept: 'application/json', Authorization: apiKeys[(r.index + attempt - 1) % apiKeys.length] },
          muteHttpExceptions: true,
        })));
      } catch (error) {
        pending.forEach(r => done.push({ index: r.index, error: `Fetch Error: ${error.message}`, transient: true }));
        pending = [];
        break;
      }
      const retry = [];
      let delay = 0;
      responses.forEach((response, i) => {
        const req = pending[i];
        const code = response.getResponseCode();
        if ((code === 429 || code >= 500) && attempt < MAX_ATTEMPTS) {
          retry.push(req);
          delay = Math.max(delay, retryDelayMs(response, attempt));
        } else {
          done.push(Object.assign({ index: req.index }, parseResponse(response, code)));
        }
      });
      pending = retry;
      if (pending.length) Utilities.sleep(delay);
    }
  }
  return done;
}

// Waits for the API's own reset time when it sends one (seconds), else backs off linearly.
function retryDelayMs(response, attempt) {
  const headers = (response.getHeaders && response.getHeaders()) || {};
  for (const key in headers) {
    const name = key.toLowerCase();
    if (name === 'retry-after' || name === 'x-ratelimit-reset') {
      const seconds = Number(headers[key]);
      if (seconds > 0) return Math.min(seconds * 1000 + 500, MAX_RETRY_DELAY_MS);
    }
  }
  return Math.min(RETRY_DELAY_MS * attempt, MAX_RETRY_DELAY_MS);
}

function parseResponse(response, code) {
  let json;
  try {
    json = JSON.parse(response.getContentText());
  } catch (error) {
    return { error: `JSON Parse Error (HTTP ${code})`, transient: true };
  }
  const status = json && json.status != null ? json.status : code;
  if (status !== 200) {
    const detail = json && json.errors && json.errors[0] && json.errors[0].message;
    return { error: `API Error: ${status}${detail ? ` (${detail})` : ''}`, transient: status === 429 || status >= 500 };
  }
  const current = json.data && json.data.current_data;
  if (!current) return { error: 'API Error: unexpected response' };
  return { data: json.data };
}

// Builds the F..L block for every row and writes it with one call per kind.
// A transient failure (rate limit, 5xx, network) keeps the row's previous data instead of wiping it.
function writeResults(sheet, L, results) {
  const previous = readPrevious(sheet, L, results.length);
  const rows = [];
  const backgrounds = [];
  const blank = new Array(L.OUT_COUNT).fill('');
  results.forEach((result, index) => {
    if (result.blank) {
      rows.push(blank.slice());
      backgrounds.push([null, null, null]);
    } else if (result.error) {
      Logger.log(`Error fetching row ${index + 2}: ${result.error}`);
      if (result.transient && previous[index].hasData) {
        rows.push(previous[index].cells);
        backgrounds.push(previous[index].backgrounds);
      } else {
        rows.push([result.error].concat(blank.slice(1)));
        backgrounds.push([null, null, null]);
      }
    } else {
      const current = result.data.current_data;
      const highest = result.data.highest_rank;
      const image = current.images && current.images.large;
      rows.push([
        current.currenttierpatched || '',
        image ? `=IMAGE("${String(image).replace(/"/g, '""')}")` : '',
        current.ranking_in_tier != null ? current.ranking_in_tier : '',
        current.elo != null ? current.elo : '',
        highest ? highest.patched_tier : '',
        highest ? highest.season : '',
        current.games_needed_for_rating != null ? current.games_needed_for_rating : '',
      ]);
      backgrounds.push([null, null, null]);
    }
  });
  sheet.getRange(2, L.OUT_FIRST, rows.length, L.OUT_COUNT).setValues(rows);
  sheet.getRange(2, L.OUT_FIRST, rows.length, 3).setBackgrounds(backgrounds);
}

// Existing F..L cells (formulas preserved) and backgrounds; hasData = a numeric ELO is present.
function readPrevious(sheet, L, count) {
  const range = sheet.getRange(2, L.OUT_FIRST, count, L.OUT_COUNT);
  const values = range.getValues();
  const formulas = range.getFormulas();
  const backgrounds = sheet.getRange(2, L.OUT_FIRST, count, 3).getBackgrounds();
  return values.map((row, i) => ({
    cells: row.map((value, j) => formulas[i][j] || value),
    backgrounds: backgrounds[i].map(color => (color === '#ffffff' ? null : color)),
    hasData: typeof row[L.ELO - L.OUT_FIRST] === 'number',
  }));
}

function sortSheetByELO(sheet, L) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 3) return; // nothing to sort
  // Active accounts (0 games needed for a rating) first, inactive/unrated below them, then by ELO.
  sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn())
    .sort([{ column: L.GAMES, ascending: true }, { column: L.ELO, ascending: false }]);
}

// Greys out F..H of accounts that still need games for a rating (games needed > 0). Runs after the sort
// so it always matches the final rows.
function styleInactiveRows(sheet, L) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;
  const count = lastRow - 1;
  const games = sheet.getRange(2, L.GAMES, count, 1).getValues();
  const elo = sheet.getRange(2, L.ELO, count, 1).getValues();
  const backgrounds = games.map((row, i) => {
    const hasData = typeof elo[i][0] === 'number';
    const inactive = hasData && row[0] > 0;
    const color = inactive ? UNRANKED_GREY : null;
    return [color, color, color];
  });
  sheet.getRange(2, L.OUT_FIRST, count, 3).setBackgrounds(backgrounds);
}

function hideColumns(sheet, columnIndexes) {
  columnIndexes.forEach(index => {
    if (!sheet.isColumnHiddenByUser(index)) sheet.hideColumns(index);
  });
}

// Keeps exactly one crown, on the top row of column E.
function manageCrownSymbols(sheet, L) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;
  const range = sheet.getRange(2, L.CROWN, lastRow - 1, 1);
  const values = range.getValues().map(([value]) => {
    const text = value == null ? '' : String(value);
    return [text.split(CROWN).join('').trim()];
  });
  values[0][0] = `${CROWN} ${values[0][0]}`.trim();
  range.setValues(values);
}
