// ============================================================
// Serie A players, authentication and empty game-data sheets
// Keeps the same player-facing PINs while storing only salted hashes.
// ============================================================

const SERIE_A_PLAYER_SHEETS = Object.freeze({
  sourcePlayers: 'Players',
  players: 'Players_SerieA',
  predictions: 'Predictions_SerieA',
  leaderboard: 'Leaderboard_SerieA',
  audit: 'Audit_Log_SerieA'
});

const SERIE_A_PLAYER_HEADERS = Object.freeze([
  'player_id', 'name', 'pin_salt', 'pin_hash', 'email', 'active', 'created_at'
]);

const SERIE_A_PREDICTION_HEADERS = Object.freeze([
  'prediction_id', 'player_id', 'match_id',
  'pred_home_score', 'pred_away_score',
  'submitted_at', 'updated_at', 'points', 'scoring_version'
]);

const SERIE_A_LEADERBOARD_HEADERS = Object.freeze([
  'rank', 'player_id', 'player_name', 'total_points',
  'exact_scores', 'correct_outcomes', 'golden_matches_scored',
  'golden_points', 'predictions_submitted', 'missed_predictions',
  'last_matchday_points', 'most_wrong_prediction', 'updated_at'
]);

const SERIE_A_AUDIT_HEADERS = Object.freeze([
  'timestamp', 'action', 'player_id', 'match_id',
  'old_value', 'new_value', 'source'
]);

const SERIE_A_SESSION_TTL_SECONDS = 21600; // Apps Script cache maximum: 6 hours

/**
 * Creates isolated Serie A tabs and migrates active/inactive player accounts
 * from Players. Predictions and leaderboard begin empty.
 * Safe to run more than once; existing Serie A players are not overwritten.
 */
function setupSerieAPlayerAndGameSheets() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const playerSheet = ensureSerieASheet_(
    spreadsheet,
    SERIE_A_PLAYER_SHEETS.players,
    SERIE_A_PLAYER_HEADERS
  );
  ensureSerieASheet_(
    spreadsheet,
    SERIE_A_PLAYER_SHEETS.predictions,
    SERIE_A_PREDICTION_HEADERS
  );
  ensureSerieASheet_(
    spreadsheet,
    SERIE_A_PLAYER_SHEETS.leaderboard,
    SERIE_A_LEADERBOARD_HEADERS
  );
  ensureSerieASheet_(
    spreadsheet,
    SERIE_A_PLAYER_SHEETS.audit,
    SERIE_A_AUDIT_HEADERS
  );

  const migration = migrateSerieAPlayers_(spreadsheet, playerSheet);
  return {
    success: true,
    players_migrated: migration.added,
    players_already_present: migration.existing,
    same_pin_for_players: true,
    plaintext_pin_stored_in_serie_a: false,
    sheets: [
      SERIE_A_PLAYER_SHEETS.players,
      SERIE_A_PLAYER_SHEETS.predictions,
      SERIE_A_PLAYER_SHEETS.leaderboard,
      SERIE_A_PLAYER_SHEETS.audit
    ]
  };
}

function migrateSerieAPlayers_(spreadsheet, targetSheet) {
  const sourceSheet = spreadsheet.getSheetByName(SERIE_A_PLAYER_SHEETS.sourcePlayers);
  if (!sourceSheet) {
    throw new Error('Source Players tab not found');
  }

  const source = serieASheetObjects_(sourceSheet);
  const target = serieASheetObjects_(targetSheet);
  const existingIds = new Set(
    target.filter(row => row.player_id !== '' && row.player_id !== null)
      .map(row => row.player_id.toString())
  );
  const now = new Date();
  const newRows = [];

  source.forEach(player => {
    if (!player.player_id || !player.name || player.pin === '' || player.pin === null) return;
    const playerId = player.player_id.toString();
    if (existingIds.has(playerId)) return;

    const salt = Utilities.getUuid();
    newRows.push([
      player.player_id,
      player.name,
      salt,
      hashSerieAPin_(player.pin, salt),
      player.email || '',
      player.active === true,
      player.created_at || now
    ]);
    existingIds.add(playerId);
  });

  if (newRows.length) {
    targetSheet.getRange(
      targetSheet.getLastRow() + 1,
      1,
      newRows.length,
      SERIE_A_PLAYER_HEADERS.length
    ).setValues(newRows);
  }

  auditSerieA_('players_migrated', null, null, null,
    'added=' + newRows.length, 'setup');
  return { added: newRows.length, existing: target.length };
}

/**
 * Authenticates with the same name and PIN used by the player, then issues a
 * short-lived random session token. The token, not player_id alone, must be
 * required by prediction write endpoints.
 */
function loginSerieAPlayer(name, pin) {
  if (!name || pin === '' || pin === null || pin === undefined) {
    return { success: false, error: 'Name and PIN required' };
  }

  const sheet = SpreadsheetApp.getActiveSpreadsheet()
    .getSheetByName(SERIE_A_PLAYER_SHEETS.players);
  if (!sheet) return { success: false, error: 'Serie A players are not configured' };

  const normalizedName = name.toString().trim().toLowerCase();
  const player = serieASheetObjects_(sheet).find(row =>
    row.name && row.name.toString().trim().toLowerCase() === normalizedName &&
    row.active === true
  );

  if (!player || !secureEquals_(
      hashSerieAPin_(pin, player.pin_salt),
      player.pin_hash.toString()
    )) {
    auditSerieA_('login_failed', null, null, name, null, 'frontend');
    return { success: false, error: 'Invalid name or PIN' };
  }

  const token = Utilities.getUuid().replace(/-/g, '') +
    Utilities.getUuid().replace(/-/g, '');
  const session = {
    player_id: player.player_id.toString(),
    player_name: player.name.toString(),
    created_at: new Date().toISOString()
  };
  CacheService.getScriptCache().put(
    serieASessionKey_(token),
    JSON.stringify(session),
    SERIE_A_SESSION_TTL_SECONDS
  );

  auditSerieA_('login_success', player.player_id, null, null, null, 'frontend');
  return {
    success: true,
    player_id: player.player_id,
    player_name: player.name,
    session_token: token,
    expires_in_seconds: SERIE_A_SESSION_TTL_SECONDS
  };
}

function requireSerieASession_(playerId, token) {
  if (!playerId || !token) throw new Error('Authentication required');
  const cached = CacheService.getScriptCache().get(serieASessionKey_(token));
  if (!cached) throw new Error('Session expired. Please sign in again.');

  const session = JSON.parse(cached);
  if (session.player_id.toString() !== playerId.toString()) {
    throw new Error('Session does not belong to this player');
  }
  return session;
}

function hashSerieAPin_(pin, salt) {
  const value = salt.toString() + ':' + pin.toString().trim() + ':' +
    getSerieAAuthPepper_();
  const bytes = Utilities.computeDigest(
    Utilities.DigestAlgorithm.SHA_256,
    value,
    Utilities.Charset.UTF_8
  );
  return bytes.map(byte => ('0' + ((byte + 256) % 256).toString(16)).slice(-2)).join('');
}

function getSerieAAuthPepper_() {
  const properties = PropertiesService.getScriptProperties();
  let pepper = properties.getProperty('SERIE_A_AUTH_PEPPER');
  if (!pepper) {
    pepper = Utilities.getUuid() + Utilities.getUuid();
    properties.setProperty('SERIE_A_AUTH_PEPPER', pepper);
  }
  return pepper;
}

function serieASessionKey_(token) {
  const bytes = Utilities.computeDigest(
    Utilities.DigestAlgorithm.SHA_256,
    token.toString(),
    Utilities.Charset.UTF_8
  );
  const digest = bytes
    .map(byte => ('0' + ((byte + 256) % 256).toString(16)).slice(-2))
    .join('');
  return 'sa_session_' + digest;
}

function secureEquals_(left, right) {
  if (left.length !== right.length) return false;
  let difference = 0;
  for (let i = 0; i < left.length; i++) {
    difference |= left.charCodeAt(i) ^ right.charCodeAt(i);
  }
  return difference === 0;
}

function ensureSerieASheet_(spreadsheet, name, headers) {
  let sheet = spreadsheet.getSheetByName(name);
  if (!sheet) sheet = spreadsheet.insertSheet(name);

  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.setFrozenRows(1);
    return sheet;
  }

  const actual = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const missing = headers.filter(header => actual.indexOf(header) === -1);
  if (missing.length) {
    throw new Error(name + ' is missing columns: ' + missing.join(', '));
  }
  return sheet;
}

function serieASheetObjects_(sheet) {
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return [];
  const headers = data[0];
  return data.slice(1)
    .filter(row => row.some(value => value !== ''))
    .map(row => {
      const object = {};
      headers.forEach((header, index) => { object[header] = row[index]; });
      return object;
    });
}

function auditSerieA_(action, playerId, matchId, oldValue, newValue, source) {
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ensureSerieASheet_(
      spreadsheet,
      SERIE_A_PLAYER_SHEETS.audit,
      SERIE_A_AUDIT_HEADERS
    );
    sheet.appendRow([
      new Date(), action, playerId || '', matchId || '',
      oldValue || '', newValue || '', source || 'api'
    ]);
  } catch (_) {
    // Auditing must not break login or setup.
  }
}

