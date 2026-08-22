// ============================================================
// Serie A fixture import and synchronisation
// Development-only module: it does not touch the World Cup Matches tab.
// ============================================================

const SERIE_A_CONFIG = Object.freeze({
  competitionCode: 'SA',
  seasonStartYear: 2026,
  seasonId: '2026-27',
  sheetName: 'Matches_SerieA',
  provider: 'football-data.org',
  syncEveryHours: 6
});

const SERIE_A_MATCH_HEADERS = Object.freeze([
  'match_id',
  'provider',
  'provider_match_id',
  'competition_id',
  'season_id',
  'round',
  'stage',
  'matchday',
  'home_team_id',
  'home_team',
  'away_team_id',
  'away_team',
  'kickoff_datetime',
  'status',
  'home_score',
  'away_score',
  'provider_last_updated',
  'last_synced_at',
  'manual_kickoff_override',
  'notes'
]);

/**
 * Creates the development tab without changing or deleting the existing
 * World Cup Matches tab. Safe to run more than once.
 */
function setupSerieAMatchesSheet() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = spreadsheet.getSheetByName(SERIE_A_CONFIG.sheetName);

  if (!sheet) {
    sheet = spreadsheet.insertSheet(SERIE_A_CONFIG.sheetName);
  }

  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, SERIE_A_MATCH_HEADERS.length)
      .setValues([SERIE_A_MATCH_HEADERS]);
    sheet.setFrozenRows(1);
    sheet.getRange('M:M').setNumberFormat('yyyy-mm-dd hh:mm');
    sheet.getRange('Q:R').setNumberFormat('yyyy-mm-dd hh:mm:ss');
    return { success: true, created: true, sheet: SERIE_A_CONFIG.sheetName };
  }

  validateSerieAMatchHeaders_(sheet);
  return { success: true, created: false, sheet: SERIE_A_CONFIG.sheetName };
}

/**
 * Read-only connectivity check. Fetches the season but writes nothing.
 */
function previewSerieAFixtures() {
  const matches = fetchSerieAFixtures_();
  const statuses = {};

  matches.forEach(match => {
    const status = mapSerieAStatus_(match.status);
    statuses[status] = (statuses[status] || 0) + 1;
  });

  return {
    success: true,
    competition: SERIE_A_CONFIG.competitionCode,
    season: SERIE_A_CONFIG.seasonId,
    matches_found: matches.length,
    statuses,
    sample: matches.slice(0, 5).map(match => ({
      provider_match_id: match.id,
      matchday: match.matchday,
      home_team: match.homeTeam && match.homeTeam.name,
      away_team: match.awayTeam && match.awayTeam.name,
      kickoff_datetime: match.utcDate,
      status: mapSerieAStatus_(match.status),
      provider_last_updated: match.lastUpdated
    }))
  };
}

/**
 * Imports new fixtures and updates existing ones by provider_match_id.
 * Existing rows that are temporarily absent from the API are preserved.
 * A row with manual_kickoff_override=TRUE keeps its manually entered kickoff.
 */
function syncSerieAFixtures() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) {
    throw new Error('Another Serie A synchronisation is already running');
  }

  try {
    setupSerieAMatchesSheet();

    const sheet = SpreadsheetApp.getActiveSpreadsheet()
      .getSheetByName(SERIE_A_CONFIG.sheetName);
    const providerMatches = fetchSerieAFixtures_();
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const rows = data.slice(1).filter(row => row.some(value => value !== ''));
    const col = headerMap_(headers);
    const now = new Date();

    const existingByProviderId = {};
    rows.forEach((row, index) => {
      const providerId = row[col.provider_match_id];
      if (providerId !== '' && providerId !== null && providerId !== undefined) {
        existingByProviderId[providerId.toString()] = index;
      }
    });

    let added = 0;
    let updated = 0;

    providerMatches.forEach(match => {
      const providerId = match.id.toString();
      const existingIndex = existingByProviderId[providerId];
      const isNew = existingIndex === undefined;
      const row = isNew
        ? new Array(headers.length).fill('')
        : rows[existingIndex];
      const beforeProviderUpdate = JSON.stringify(row);
      const manualKickoff = truthy_(row[col.manual_kickoff_override]);

      setField_(row, col, 'match_id', buildSerieAMatchId_(match.id));
      setField_(row, col, 'provider', SERIE_A_CONFIG.provider);
      setField_(row, col, 'provider_match_id', match.id);
      setField_(row, col, 'competition_id', SERIE_A_CONFIG.competitionCode);
      setField_(row, col, 'season_id', SERIE_A_CONFIG.seasonId);
      setField_(row, col, 'round', match.matchday ? 'Matchday ' + match.matchday : 'Serie A');
      setField_(row, col, 'stage', 'league');
      setField_(row, col, 'matchday', match.matchday || '');
      setField_(row, col, 'home_team_id', match.homeTeam && match.homeTeam.id || '');
      setField_(row, col, 'home_team', serieATeamName_(match.homeTeam));
      setField_(row, col, 'away_team_id', match.awayTeam && match.awayTeam.id || '');
      setField_(row, col, 'away_team', serieATeamName_(match.awayTeam));

      if (!manualKickoff) {
        setField_(row, col, 'kickoff_datetime', match.utcDate ? new Date(match.utcDate) : '');
      }

      setField_(row, col, 'status', mapSerieAStatus_(match.status));
      setField_(row, col, 'home_score', serieAScore_(match, 'home'));
      setField_(row, col, 'away_score', serieAScore_(match, 'away'));
      setField_(row, col, 'provider_last_updated',
        match.lastUpdated ? new Date(match.lastUpdated) : '');
      const providerChanged = JSON.stringify(row) !== beforeProviderUpdate;
      setField_(row, col, 'last_synced_at', now);

      if (isNew) {
        rows.push(row);
        existingByProviderId[providerId] = rows.length - 1;
        added++;
      } else if (providerChanged) {
        updated++;
      }
    });

    rows.sort((a, b) => {
      const matchdayDiff = Number(a[col.matchday] || 999) - Number(b[col.matchday] || 999);
      if (matchdayDiff !== 0) return matchdayDiff;
      return dateValue_(a[col.kickoff_datetime]) - dateValue_(b[col.kickoff_datetime]);
    });

    if (rows.length) {
      sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);
    }

    CacheService.getScriptCache().remove('vals_' + SERIE_A_CONFIG.sheetName);
    auditLog('serie_a_fixtures_synced', null, null, null,
      'added=' + added + ', updated=' + updated, 'trigger');

    return {
      success: true,
      sheet: SERIE_A_CONFIG.sheetName,
      matches_received: providerMatches.length,
      added,
      updated,
      total_rows: rows.length,
      synced_at: now.toISOString()
    };
  } finally {
    lock.releaseLock();
  }
}

/**
 * Installs one six-hour trigger. Safe to run more than once.
 */
function createSerieAFixtureSyncTrigger() {
  const handler = 'syncSerieAFixtures';
  const exists = ScriptApp.getProjectTriggers()
    .some(trigger => trigger.getHandlerFunction() === handler);

  if (exists) {
    return { success: true, created: false, message: 'Trigger already exists' };
  }

  ScriptApp.newTrigger(handler)
    .timeBased()
    .everyHours(SERIE_A_CONFIG.syncEveryHours)
    .create();

  return { success: true, created: true, every_hours: SERIE_A_CONFIG.syncEveryHours };
}

function fetchSerieAFixtures_() {
  const url = 'https://api.football-data.org/v4/competitions/' +
    SERIE_A_CONFIG.competitionCode + '/matches?season=' +
    SERIE_A_CONFIG.seasonStartYear;

  const response = UrlFetchApp.fetch(url, {
    method: 'get',
    headers: { 'X-Auth-Token': getSerieAFootballApiKey_() },
    muteHttpExceptions: true
  });

  const statusCode = response.getResponseCode();
  const body = response.getContentText();
  if (statusCode !== 200) {
    throw new Error('football-data.org returned HTTP ' + statusCode + ': ' + body.slice(0, 250));
  }

  const payload = JSON.parse(body);
  if (!Array.isArray(payload.matches)) {
    throw new Error('football-data.org response does not contain a matches array');
  }

  return payload.matches;
}

function getSerieAFootballApiKey_() {
  const key = PropertiesService.getScriptProperties()
    .getProperty('FOOTBALL_API_KEY');
  if (!key) {
    throw new Error('Missing FOOTBALL_API_KEY script property');
  }
  return key;
}

function validateSerieAMatchHeaders_(sheet) {
  const lastColumn = Math.max(sheet.getLastColumn(), SERIE_A_MATCH_HEADERS.length);
  const headers = sheet.getRange(1, 1, 1, lastColumn).getValues()[0];
  const missing = SERIE_A_MATCH_HEADERS.filter(header => headers.indexOf(header) === -1);
  if (missing.length) {
    throw new Error(
      'The ' + SERIE_A_CONFIG.sheetName + ' tab is missing columns: ' + missing.join(', ')
    );
  }
}

function headerMap_(headers) {
  const map = {};
  headers.forEach((header, index) => { map[header] = index; });
  return map;
}

function setField_(row, columnMap, field, value) {
  const index = columnMap[field];
  if (index === undefined) throw new Error('Missing required column: ' + field);
  row[index] = value;
}

function buildSerieAMatchId_(providerMatchId) {
  return 'SA-' + SERIE_A_CONFIG.seasonStartYear + '-' + providerMatchId;
}

function serieATeamName_(team) {
  if (!team) return '';
  return team.shortName || team.name || '';
}

function mapSerieAStatus_(providerStatus) {
  const statuses = {
    SCHEDULED: 'upcoming',
    TIMED: 'upcoming',
    IN_PLAY: 'live',
    PAUSED: 'live',
    FINISHED: 'finished',
    POSTPONED: 'postponed',
    SUSPENDED: 'suspended',
    // A permanently cancelled match is treated as void and never scored.
    CANCELLED: 'void',
    AWARDED: 'finished'
  };
  return statuses[providerStatus] || (providerStatus || 'unknown').toString().toLowerCase();
}

/**
 * Central Serie A editing rule, ready to be used by the new API layer.
 *
 * - upcoming: editable until the normal lock time;
 * - postponed: editable only when the provider supplies a future kickoff that
 *   is still before lock;
 * - live/suspended/finished/void: never editable.
 */
function canEditSerieAMatch_(match, lockMinutes, now) {
  const status = (match.status || '').toString().trim().toLowerCase();
  if (status !== 'upcoming' && status !== 'postponed') return false;

  const kickoff = match.kickoff_datetime instanceof Date
    ? match.kickoff_datetime
    : new Date(match.kickoff_datetime);
  if (isNaN(kickoff.getTime())) return false;

  const referenceTime = now instanceof Date ? now : new Date();
  const lockTime = kickoff.getTime() - Number(lockMinutes || 0) * 60 * 1000;
  return referenceTime.getTime() < lockTime;
}

function shouldScoreSerieAMatch_(match) {
  return (match.status || '').toString().trim().toLowerCase() === 'finished';
}

function serieAScore_(match, side) {
  if (!match || !match.score) return '';
  const score = match.score.fullTime || match.score.regularTime;
  if (!score || score[side] === null || score[side] === undefined) return '';
  return score[side];
}

function truthy_(value) {
  return value === true || value === 1 ||
    (typeof value === 'string' && value.trim().toLowerCase() === 'true');
}

function dateValue_(value) {
  if (!value) return Number.MAX_SAFE_INTEGER;
  const timestamp = value instanceof Date ? value.getTime() : new Date(value).getTime();
  return isNaN(timestamp) ? Number.MAX_SAFE_INTEGER : timestamp;
}
