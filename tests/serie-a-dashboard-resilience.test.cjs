const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');

const html = fs.readFileSync(path.join(__dirname, '..', 'serie-a', 'index.html'), 'utf8');
const scriptStart = html.lastIndexOf('<script>') + '<script>'.length;
const scriptEnd = html.lastIndexOf('</script>');
const frontendSource = html.slice(scriptStart, scriptEnd);

function frontendHarness(fetchImplementation) {
  const storage = new Map();
  let retryWaits = 0;
  const context = vm.createContext({
    console,
    Date,
    Error,
    TypeError,
    SyntaxError,
    Math,
    Number,
    String,
    Array,
    Object,
    JSON,
    Promise,
    Intl,
    AbortController,
    fetch: fetchImplementation,
    setTimeout(callback, milliseconds) {
      // Do not allow the request timeout to fire in these immediate-response tests.
      if (milliseconds >= 10000) return setTimeout(callback, 1000000);
      retryWaits++;
      return setTimeout(callback, 0);
    },
    clearTimeout,
    localStorage: {
      getItem: key => storage.has(key) ? storage.get(key) : null,
      setItem: (key, value) => storage.set(key, String(value)),
      removeItem: key => storage.delete(key)
    },
    document: { addEventListener() {} },
    window: { scrollTo() {} },
    confirm: () => true
  });
  vm.runInContext(frontendSource, context);
  return { context, storage, retryWaits: () => retryWaits };
}

test('dashboard transport retries one interrupted request and then succeeds', async () => {
  let calls = 0;
  const harness = frontendHarness(async () => {
    calls++;
    if (calls === 1) throw new TypeError('Load failed');
    return { ok: true, status: 200, json: async () => ({ success: true }) };
  });
  let retryNotices = 0;
  const result = await harness.context.api({ action: 'serieAGetDashboard' }, {
    retries: 1,
    onRetry: () => { retryNotices++; }
  });
  assert.equal(result.success, true);
  assert.equal(calls, 2);
  assert.equal(retryNotices, 1);
  assert.equal(harness.retryWaits(), 1);
});

test('non-retryable errors and normal write calls are never repeated', async () => {
  let calls = 0;
  const harness = frontendHarness(async () => {
    calls++;
    throw new Error('Application failure');
  });
  await assert.rejects(() => harness.context.api({ action: 'serieASubmitPrediction' }, {
    retries: 1
  }), /Application failure/);
  assert.equal(calls, 1);
});

test('dashboard cache is private to the player, bounded and cleared on logout helpers', () => {
  const harness = frontendHarness(async () => {
    throw new Error('fetch not expected');
  });
  vm.runInContext(`
    session = { player_id: 'P1' };
    matches = [{ match_id: 'M1', can_edit: true }];
    leaderboard = [{ player_id: 'P1', total_points: 7 }];
    writeDashboardCache();
  `, harness.context);
  assert.ok(harness.storage.has('ss_serie_a_dashboard_v1_P1'));
  assert.equal(JSON.parse(harness.storage.get('ss_serie_a_dashboard_v1_P1')).matches[0].match_id, 'M1');
  vm.runInContext(`session = { player_id: 'P2' };`, harness.context);
  assert.equal(harness.context.readDashboardCache(), null);
  vm.runInContext(`session = { player_id: 'P1' }; clearDashboardCache();`, harness.context);
  assert.equal(harness.storage.size, 0);
});

test('combined backend endpoint authenticates once and returns both datasets', () => {
  const context = vm.createContext({});
  const source = fs.readFileSync(path.join(__dirname, '..', 'backend', 'SerieAApi.gs'), 'utf8');
  vm.runInContext(source, context);
  vm.runInContext(`
    var authChecks = 0;
    requireSerieASession_ = function(playerId, token) {
      if (playerId !== 'P1' || token !== 'token') throw new Error('bad auth');
      authChecks++;
    };
    buildSerieAMatchesResponse_ = function(playerId) {
      return { success: true, matches: [{ match_id: 'M1', player_id: playerId }] };
    };
    buildSerieALeaderboardResponse_ = function() {
      return { success: true, leaderboard: [{ player_id: 'P1', total_points: 9 }] };
    };
  `, context);
  const result = context.handleSerieARequest_({
    action: 'serieAGetDashboard', player_id: 'P1', session_token: 'token'
  });
  assert.equal(result.success, true);
  assert.equal(result.matches.length, 1);
  assert.equal(result.leaderboard[0].total_points, 9);
  assert.equal(vm.runInContext('authChecks', context), 1);
});

test('frontend uses combined reads, bounded timeout, cache safety and a legacy rollout fallback', () => {
  assert.match(frontendSource, /action: 'serieAGetDashboard'/);
  assert.match(frontendSource, /API_TIMEOUT_MS = 18000/);
  assert.match(frontendSource, /retries: 1/);
  assert.match(frontendSource, /unknown action/i);
  assert.match(frontendSource, /kickoff <= now[\s\S]*match\.can_edit = false/);
  assert.match(frontendSource, /writeDashboardCache\(\)/);
  assert.match(frontendSource, /data-retry-load/);
});
