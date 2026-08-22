// ============================================================
// Serie A API layer
// Uses isolated Serie A sheets and session-authenticated prediction writes.
// ============================================================

/**
 * Called from the existing Code.gs request router. Returns null when the action
 * does not belong to the Serie A API, preserving World Cup compatibility.
 */
function handleSerieARequest_(params) {
  switch (params.action) {
    case 'serieALogin':
      return loginSerieAPlayer(params.name, params.pin);
    case 'serieAGetMatches':
      return getSerieAMatches(
        params.player_id,
        params.session_token
      );
    case 'serieASubmitPrediction':
      return submitSerieAPrediction(
        params.player_id,
        params.session_token,
        params.match_id,
        params.pred_home_score,
        params.pred_away_score
      );
    case 'serieAGetLeaderboard':
      return getSerieALeaderboard(
        params.player_id,
        params.session_token
      );
    default:
      return null;
  }
}

function getSerieAMatches(playerId, sessionToken) {
  requireSerieASession_(playerId, sessionToken);

  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const matchSheet = spreadsheet.getSheetByName(SERIE_A_CONFIG.sheetName);
  const predictionSheet = spreadsheet.getSheetByName(
    SERIE_A_PLAYER_SHEETS.predictions
  );
  const playerSheet = spreadsheet.getSheetByName(SERIE_A_PLAYER_SHEETS.players);
  if (!matchSheet || !predictionSheet || !playerSheet) {
    return { success: false, error: 'Serie A sheets are not configured' };
  }

  const settings = getSerieASettings_();
  const lockMinutes = numberSetting_(settings.lock_minutes_before_kickoff, 0);
  const now = new Date();
  const matches = serieASheetObjects_(matchSheet);
  const predictions = serieASheetObjects_(predictionSheet);
  const players = serieASheetObjects_(playerSheet);
  const playerNames = {};
  players.forEach(player => {
    playerNames[player.player_id.toString()] = player.name;
  });

  const predictionsByMatch = {};
  const myPredictions = {};
  predictions.forEach(prediction => {
    if (!prediction.match_id || !prediction.player_id) return;
    const matchId = prediction.match_id.toString();
    if (!predictionsByMatch[matchId]) predictionsByMatch[matchId] = [];
    predictionsByMatch[matchId].push(prediction);
    if (prediction.player_id.toString() === playerId.toString()) {
      myPredictions[matchId] = prediction;
    }
  });

  return {
    success: true,
    competition: SERIE_A_CONFIG.competitionCode,
    season: SERIE_A_CONFIG.seasonId,
    server_time: now.toISOString(),
    matches: matches.map(match => {
      const matchId = match.match_id.toString();
      const canEdit = canEditSerieAMatch_(match, lockMinutes, now);
      const reveal = shouldRevealSerieAPredictions_(match, now);
      const mine = myPredictions[matchId] || null;
      const allPredictions = reveal
        ? (predictionsByMatch[matchId] || []).map(prediction => ({
            player_id: prediction.player_id,
            player_name: playerNames[prediction.player_id.toString()] || 'Unknown',
            is_me: prediction.player_id.toString() === playerId.toString(),
            pred_home_score: prediction.pred_home_score,
            pred_away_score: prediction.pred_away_score,
            points: prediction.points
          }))
        : null;

      return {
        match_id: match.match_id,
        provider_match_id: match.provider_match_id,
        round: match.round,
        stage: 'league',
        matchday: match.matchday,
        home_team: match.home_team,
        away_team: match.away_team,
        kickoff_datetime: serieAIsoDate_(match.kickoff_datetime),
        status: match.status,
        locked: !canEdit,
        can_edit: canEdit,
        show_all_predictions: reveal,
        is_golden: truthy_(match.is_golden),
        points_multiplier: numberSetting_(match.points_multiplier, 1),
        home_score: shouldScoreSerieAMatch_(match) ? match.home_score : null,
        away_score: shouldScoreSerieAMatch_(match) ? match.away_score : null,
        user_prediction: mine ? {
          pred_home_score: mine.pred_home_score,
          pred_away_score: mine.pred_away_score,
          points: mine.points
        } : null,
        all_predictions: allPredictions
      };
    })
  };
}

function submitSerieAPrediction(playerId, sessionToken, matchId, home, away) {
  requireSerieASession_(playerId, sessionToken);
  const homeScore = strictSerieAScore_(home);
  const awayScore = strictSerieAScore_(away);

  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) {
    return { success: false, error: 'Another save is in progress. Please retry.' };
  }

  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const matchSheet = spreadsheet.getSheetByName(SERIE_A_CONFIG.sheetName);
    const predictionSheet = spreadsheet.getSheetByName(
      SERIE_A_PLAYER_SHEETS.predictions
    );
    if (!matchSheet || !predictionSheet) {
      return { success: false, error: 'Serie A sheets are not configured' };
    }

    const match = serieASheetObjects_(matchSheet).find(row =>
      row.match_id && row.match_id.toString() === matchId.toString()
    );
    if (!match) return { success: false, error: 'Match not found' };

    const settings = getSerieASettings_();
    const lockMinutes = numberSetting_(settings.lock_minutes_before_kickoff, 0);
    if (!canEditSerieAMatch_(match, lockMinutes, new Date())) {
      auditSerieA_('prediction_rejected_locked', playerId, matchId,
        null, null, 'frontend');
      return { success: false, error: 'This match is locked' };
    }

    const data = predictionSheet.getDataRange().getValues();
    const headers = data[0];
    const rows = data.slice(1);
    const col = headerMap_(headers);
    const existingIndex = rows.findIndex(row =>
      row[col.player_id] && row[col.match_id] &&
      row[col.player_id].toString() === playerId.toString() &&
      row[col.match_id].toString() === matchId.toString()
    );
    const now = new Date();

    if (existingIndex >= 0) {
      const row = rows[existingIndex];
      const oldScore = row[col.pred_home_score] + '-' + row[col.pred_away_score];
      row[col.pred_home_score] = homeScore;
      row[col.pred_away_score] = awayScore;
      row[col.updated_at] = now;
      row[col.points] = '';
      row[col.scoring_version] = 'serie-a-v1';
      predictionSheet.getRange(existingIndex + 2, 1, 1, headers.length)
        .setValues([row]);
      auditSerieA_('prediction_updated', playerId, matchId,
        oldScore, homeScore + '-' + awayScore, 'frontend');
      return { success: true, message: 'Prediction updated' };
    }

    const newRow = new Array(headers.length).fill('');
    setField_(newRow, col, 'prediction_id', 'SAP-' + Utilities.getUuid());
    setField_(newRow, col, 'player_id', playerId);
    setField_(newRow, col, 'match_id', matchId);
    setField_(newRow, col, 'pred_home_score', homeScore);
    setField_(newRow, col, 'pred_away_score', awayScore);
    setField_(newRow, col, 'submitted_at', now);
    setField_(newRow, col, 'updated_at', now);
    setField_(newRow, col, 'points', '');
    setField_(newRow, col, 'scoring_version', 'serie-a-v1');
    predictionSheet.appendRow(newRow);
    auditSerieA_('prediction_created', playerId, matchId,
      null, homeScore + '-' + awayScore, 'frontend');
    return { success: true, message: 'Prediction saved' };
  } finally {
    lock.releaseLock();
  }
}

function getSerieALeaderboard(playerId, sessionToken) {
  requireSerieASession_(playerId, sessionToken);
  const sheet = SpreadsheetApp.getActiveSpreadsheet()
    .getSheetByName(SERIE_A_PLAYER_SHEETS.leaderboard);
  if (!sheet) return { success: false, error: 'Serie A leaderboard is not configured' };
  return { success: true, leaderboard: serieASheetObjects_(sheet) };
}

/**
 * Internal/admin function. It is intentionally not exposed as a public web
 * action. Fixture sync calls it after updated results are stored.
 */
function recalculateSerieALeaderboard_() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const matchSheet = spreadsheet.getSheetByName(SERIE_A_CONFIG.sheetName);
  const predictionSheet = spreadsheet.getSheetByName(
    SERIE_A_PLAYER_SHEETS.predictions
  );
  const playerSheet = spreadsheet.getSheetByName(SERIE_A_PLAYER_SHEETS.players);
  const leaderboardSheet = spreadsheet.getSheetByName(
    SERIE_A_PLAYER_SHEETS.leaderboard
  );
  if (!matchSheet || !predictionSheet || !playerSheet || !leaderboardSheet) {
    return { success: false, skipped: true, error: 'Serie A sheets are incomplete' };
  }

  const settings = getSerieASettings_();
  const matches = serieASheetObjects_(matchSheet);
  const players = serieASheetObjects_(playerSheet).filter(player => truthy_(player.active));
  const matchById = {};
  const finishedMatches = matches.filter(match =>
    shouldScoreSerieAMatch_(match) &&
    serieAValidScoreValue_(match.home_score) &&
    serieAValidScoreValue_(match.away_score)
  );
  finishedMatches.forEach(match => { matchById[match.match_id.toString()] = match; });

  const data = predictionSheet.getDataRange().getValues();
  const headers = data[0];
  const rows = data.slice(1);
  const col = headerMap_(headers);
  const stats = {};
  players.forEach(player => {
    stats[player.player_id.toString()] = {
      player_id: player.player_id,
      player_name: player.name,
      total_points: 0,
      exact_scores: 0,
      correct_outcomes: 0,
      golden_matches_scored: 0,
      golden_points: 0,
      predictions_submitted: 0,
      missed_predictions: 0,
      last_matchday_points: 0,
      most_wrong_prediction: '',
      _worst_wrongness: -1
    };
  });

  const finishedMatchdays = finishedMatches
    .map(match => Number(match.matchday))
    .filter(matchday => Number.isFinite(matchday));
  const latestFinishedMatchday = finishedMatchdays.length
    ? Math.max.apply(null, finishedMatchdays)
    : null;

  rows.forEach(row => {
    const playerId = row[col.player_id] && row[col.player_id].toString();
    const matchId = row[col.match_id] && row[col.match_id].toString();
    const stat = stats[playerId];
    const match = matchById[matchId];
    if (!stat || !match) return;

    const prediction = {
      pred_home_score: row[col.pred_home_score],
      pred_away_score: row[col.pred_away_score]
    };
    const points = calculateSerieAPoints_(prediction, match, settings);
    row[col.points] = points;
    stat.total_points += points;
    stat.predictions_submitted++;

    const predictedHome = Number(prediction.pred_home_score);
    const predictedAway = Number(prediction.pred_away_score);
    const actualHome = Number(match.home_score);
    const actualAway = Number(match.away_score);
    if (predictedHome === actualHome && predictedAway === actualAway) {
      stat.exact_scores++;
    } else if (serieAOutcome_(predictedHome, predictedAway) ===
               serieAOutcome_(actualHome, actualAway)) {
      stat.correct_outcomes++;
    }

    if (truthy_(match.is_golden)) {
      stat.golden_matches_scored++;
      stat.golden_points += points;
    }
    if (Number(match.matchday) === latestFinishedMatchday) {
      stat.last_matchday_points += points;
    }

    const wrongness = Math.abs(predictedHome - actualHome) +
      Math.abs(predictedAway - actualAway);
    if (Number.isFinite(wrongness) && wrongness > stat._worst_wrongness) {
      stat._worst_wrongness = wrongness;
      stat.most_wrong_prediction = match.home_team + ' ' +
        predictedHome + '-' + predictedAway + ' ' + match.away_team +
        ' → actual ' + actualHome + '-' + actualAway;
    }
  });

  if (rows.length) {
    predictionSheet.getRange(2, 1, rows.length, headers.length).setValues(rows);
  }

  Object.values(stats).forEach(stat => {
    stat.missed_predictions = Math.max(
      0,
      finishedMatches.length - stat.predictions_submitted
    );
  });
  const sorted = Object.values(stats).sort((left, right) => {
    if (right.total_points !== left.total_points) {
      return right.total_points - left.total_points;
    }
    if (right.exact_scores !== left.exact_scores) {
      return right.exact_scores - left.exact_scores;
    }
    return right.correct_outcomes - left.correct_outcomes;
  });

  leaderboardSheet.clearContents();
  leaderboardSheet.getRange(1, 1, 1, SERIE_A_LEADERBOARD_HEADERS.length)
    .setValues([SERIE_A_LEADERBOARD_HEADERS]);
  if (sorted.length) {
    const leaderboardRows = sorted.map((stat, index) => [
      index + 1,
      stat.player_id,
      stat.player_name,
      stat.total_points,
      stat.exact_scores,
      stat.correct_outcomes,
      stat.golden_matches_scored,
      stat.golden_points,
      stat.predictions_submitted,
      stat.missed_predictions,
      stat.last_matchday_points,
      stat.most_wrong_prediction,
      new Date()
    ]);
    leaderboardSheet.getRange(
      2, 1, leaderboardRows.length, SERIE_A_LEADERBOARD_HEADERS.length
    ).setValues(leaderboardRows);
  }

  auditSerieA_('leaderboard_recalculated', null, null, null,
    'finished_matches=' + finishedMatches.length, 'trigger');
  return {
    success: true,
    players: sorted.length,
    finished_matches: finishedMatches.length
  };
}

function strictSerieAScore_(value) {
  if (value === '' || value === null || value === undefined) {
    throw new Error('Enter both scores');
  }
  const text = value.toString().trim();
  if (!/^\d{1,2}$/.test(text)) {
    throw new Error('Scores must be whole numbers from 0 to 99');
  }
  return Number(text);
}

function shouldRevealSerieAPredictions_(match, now) {
  const status = (match.status || '').toString().trim().toLowerCase();
  if (status === 'finished' || status === 'live' || status === 'suspended' ||
      status === 'void') {
    return true;
  }
  if (status === 'postponed') return false;

  const kickoff = match.kickoff_datetime instanceof Date
    ? match.kickoff_datetime
    : new Date(match.kickoff_datetime);
  return !isNaN(kickoff.getTime()) && now.getTime() >= kickoff.getTime();
}

function serieAIsoDate_(value) {
  if (!value) return null;
  const date = value instanceof Date ? value : new Date(value);
  return isNaN(date.getTime()) ? null : date.toISOString();
}

/**
 * Read-only setup check for the Apps Script editor. It logs a compact report so
 * the result is visible in Execution log when run manually.
 */
function verifySerieABackendSetup() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const requiredSheets = [
    SERIE_A_CONFIG.sheetName,
    SERIE_A_SETTINGS_SHEET,
    SERIE_A_PLAYER_SHEETS.players,
    SERIE_A_PLAYER_SHEETS.predictions,
    SERIE_A_PLAYER_SHEETS.leaderboard,
    SERIE_A_PLAYER_SHEETS.audit
  ];
  const missingSheets = requiredSheets.filter(name => !spreadsheet.getSheetByName(name));
  const result = {
    success: missingSheets.length === 0,
    missing_sheets: missingSheets,
    matches: 0,
    players: 0,
    predictions: 0,
    duplicate_provider_match_ids: 0,
    matchdays_with_multiple_golden: [],
    plaintext_pin_column_present: false,
    lock_minutes_before_kickoff: null,
    golden_multiplier: null
  };

  if (!missingSheets.includes(SERIE_A_CONFIG.sheetName)) {
    const matches = serieASheetObjects_(
      spreadsheet.getSheetByName(SERIE_A_CONFIG.sheetName)
    );
    result.matches = matches.length;
    const providerIds = matches
      .map(match => match.provider_match_id && match.provider_match_id.toString())
      .filter(Boolean);
    result.duplicate_provider_match_ids =
      providerIds.length - new Set(providerIds).size;

    const goldenByMatchday = {};
    matches.filter(match => truthy_(match.is_golden)).forEach(match => {
      const key = match.matchday.toString();
      goldenByMatchday[key] = (goldenByMatchday[key] || 0) + 1;
    });
    result.matchdays_with_multiple_golden = Object.keys(goldenByMatchday)
      .filter(matchday => goldenByMatchday[matchday] > 1);
  }

  if (!missingSheets.includes(SERIE_A_PLAYER_SHEETS.players)) {
    const playerSheet = spreadsheet.getSheetByName(SERIE_A_PLAYER_SHEETS.players);
    result.players = serieASheetObjects_(playerSheet).length;
    const headers = playerSheet.getRange(1, 1, 1, playerSheet.getLastColumn())
      .getValues()[0];
    result.plaintext_pin_column_present = headers.indexOf('pin') !== -1;
  }

  if (!missingSheets.includes(SERIE_A_PLAYER_SHEETS.predictions)) {
    result.predictions = serieASheetObjects_(
      spreadsheet.getSheetByName(SERIE_A_PLAYER_SHEETS.predictions)
    ).length;
  }

  if (!missingSheets.includes(SERIE_A_SETTINGS_SHEET)) {
    const settings = getSerieASettings_();
    result.lock_minutes_before_kickoff =
      numberSetting_(settings.lock_minutes_before_kickoff, null);
    result.golden_multiplier = numberSetting_(settings.golden_match_multiplier, null);
  }

  result.success = result.success &&
    result.matches > 0 &&
    result.players > 0 &&
    result.duplicate_provider_match_ids === 0 &&
    result.matchdays_with_multiple_golden.length === 0 &&
    result.plaintext_pin_column_present === false &&
    result.lock_minutes_before_kickoff === 0 &&
    result.golden_multiplier === 3;

  Logger.log(JSON.stringify(result, null, 2));
  return result;
}
