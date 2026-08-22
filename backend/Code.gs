// ============================================================
// Sportsbook & Steak WC Predictions — Google Apps Script
// Paste this entire file into your Apps Script project
//
// v4 — First Goalscorer dropdown: single list with one labelled section
//      per team (home squad, then away squad) via scorer_options_home/away.
//      Correct-method bonus changed from 3 to 2 points.
//      (Max knockout = 5 exact + 2 qualifier + 2 method + 3 scorer = 12 pts.)
// ============================================================

// -------------------------------------------------------
// ENTRY POINTS
// -------------------------------------------------------

function doGet(e) {
  const params = e.parameter || {};
  return handleRequest(params);
}

function doPost(e) {
  let params = {};
  try {
    if (e.postData && e.postData.contents) {
      params = JSON.parse(e.postData.contents);
    }
  } catch (_) {
    params = e.parameter || {};
  }
  return handleRequest(params);
}

function handleRequest(params) {
  let result;
  try {
    const action = params.action;
    switch (action) {
      case 'login':
        result = loginPlayer(params.name, params.pin);
        break;
      case 'getMatches':
        result = getMatches(params.player_id);
        break;
      case 'submitPrediction':
        result = submitPrediction(
          params.player_id,
          params.match_id,
          params.pred_home_score_90,
          params.pred_away_score_90,
          params.predicted_qualifier,
          params.predicted_qualification_method,
          params.predicted_first_scorer
        );
        break;
      case 'getLeaderboard':
        result = getLeaderboard();
        break;
      case 'getMatchPredictions':
        result = getMatchPredictions(params.match_id, params.player_id);
        break;
      case 'recalculateAll':
        result = recalculateAll();
        break;
      case 'generateReminderMessage':
        result = { success: true, message: generateReminderMessage() };
        break;
      default:
        result = { success: false, error: 'Unknown action: ' + action };
    }
  } catch (err) {
    result = { success: false, error: err.message };
  }

  return ContentService
    .createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}


// -------------------------------------------------------
// HELPERS
// -------------------------------------------------------

function getSheet(name) {
  return SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);
}

function sheetToObjects(sheet) {
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return [];
  const headers = data[0];
  return data.slice(1).map(row => {
    const obj = {};
    headers.forEach((h, i) => { obj[h] = row[i]; });
    return obj;
  });
}

// -------------------------------------------------------
// CACHING (CacheService) — read-heavy, slow-changing tabs only.
// Matches / Players / Squads are cached ~60s. Predictions are NEVER cached
// (they change on every save). Trigger jobs (recalc/sync) keep reading fresh.
// -------------------------------------------------------
function cachedValues(name, ttl) {
  const cache = CacheService.getScriptCache();
  const key = 'vals_' + name;
  const hit = cache.get(key);
  if (hit) {
    try { return JSON.parse(hit); } catch (e) {}
  }
  const sheet = getSheet(name);
  const values = sheet ? sheet.getDataRange().getValues() : [];
  try { cache.put(key, JSON.stringify(values), ttl || 60); } catch (e) { /* >100KB — skip caching */ }
  return values;
}

// Same output shape as sheetToObjects(), but reads by sheet name via the cache.
function cachedObjects(name, ttl) {
  const data = cachedValues(name, ttl);
  if (data.length < 2) return [];
  const headers = data[0];
  return data.slice(1).map(row => {
    const obj = {};
    headers.forEach((h, i) => { obj[h] = row[i]; });
    return obj;
  });
}

function getSettings() {
  const rows = sheetToObjects(getSheet('Settings'));
  const s = {};
  rows.forEach(r => { s[r.setting] = r.value; });
  return s;
}

// Returns { "Team Name": ["Player 1", "Player 2", ...], ... }
// Reads from the 'Squads' tab (columns: team, pos, player_name)
function getSquadsByTeam() {
  const rows = cachedObjects('Squads');
  const byTeam = {};
  rows.forEach(r => {
    if (!r.team || !r.player_name) return;
    const t = r.team.toString().trim();
    if (!byTeam[t]) byTeam[t] = [];
    byTeam[t].push(r.player_name.toString().trim());
  });
  return byTeam;
}

function isLocked(kickoffDatetime, lockMinutes) {
  const kickoff = new Date(kickoffDatetime);
  const lockTime = new Date(kickoff.getTime() - lockMinutes * 60 * 1000);
  return new Date() >= lockTime;
}

function getOutcome(home, away) {
  const h = parseInt(home), a = parseInt(away);
  if (h > a) return '1';
  if (h < a) return '2';
  return 'X';
}

function auditLog(action, playerId, matchId, oldVal, newVal, source) {
  try {
    getSheet('Audit_Log').appendRow([
      new Date(), action,
      playerId || '', matchId || '',
      oldVal || '', newVal || '',
      source || 'api'
    ]);
  } catch (_) { /* non-fatal */ }
}


// -------------------------------------------------------
// 1. loginPlayer
// -------------------------------------------------------

function loginPlayer(name, pin) {
  if (!name || !pin) return { success: false, error: 'Name and PIN required' };

  const players = cachedObjects('Players');
  const player = players.find(p =>
    p.name.toString().trim().toLowerCase() === name.toString().trim().toLowerCase() &&
    p.pin.toString().trim() === pin.toString().trim() &&
    p.active === true
  );

  if (player) {
    auditLog('login_success', player.player_id, null, null, null, 'frontend');
    return { success: true, player_id: player.player_id, player_name: player.name };
  }

  auditLog('login_failed', null, null, name, null, 'frontend');
  return { success: false, error: 'Invalid name or PIN' };
}


// -------------------------------------------------------
// 2. getMatches
// -------------------------------------------------------

function getMatches(player_id) {
  const settings = getSettings();
  const lockMins = parseInt(settings.lock_minutes_before_kickoff) || 5;

  const matches = cachedObjects('Matches');
  const predictions = sheetToObjects(getSheet('Predictions'));   // never cached — changes on every save
  const players = cachedObjects('Players');
  const squads = getSquadsByTeam();                              // cached internally

  // Build player name map
  const playerNames = {};
  players.forEach(p => { playerNames[p.player_id.toString()] = p.name; });

  // Index user's own predictions
  const myPreds = {};
  predictions.forEach(p => {
    if (p.player_id.toString() === player_id.toString()) {
      myPreds[p.match_id.toString()] = p;
    }
  });

  // Index all predictions by match (for locked/finished matches)
  const allPredsByMatch = {};
  predictions.forEach(p => {
    if (!p.match_id) return;
    const mid = p.match_id.toString();
    if (!allPredsByMatch[mid]) allPredsByMatch[mid] = [];
    allPredsByMatch[mid].push(p);
  });

  return {
    success: true,
    matches: matches.map(m => {
      const locked = isLocked(m.kickoff_datetime, lockMins);
      const finished = m.status === 'finished';
      const void_ = m.status === 'void';
      const showAll = locked || finished;

      const myPred = myPreds[m.match_id.toString()] || null;

      // Build all_predictions only when revealed
      let all_predictions = null;
      if (showAll) {
        const matchPreds = allPredsByMatch[m.match_id.toString()] || [];
        all_predictions = matchPreds.map(p => ({
          player_name: playerNames[p.player_id.toString()] || 'Unknown',
          is_me: p.player_id.toString() === player_id.toString(),
          pred_home_score_90: p.pred_home_score_90,
          pred_away_score_90: p.pred_away_score_90,
          predicted_qualifier: p.predicted_qualifier,
          predicted_qualification_method: p.predicted_qualification_method,
          predicted_first_scorer: p.predicted_first_scorer,
          points: p.points
        }));
      }

      // First-scorer dropdown options (knockout + editable only).
      // Home and away squads sent separately so the frontend can show
      // one labelled section per team within a single dropdown.
      const koEditable = (m.stage === 'knockout' && !locked && !finished && !void_);
      const scorerHome = koEditable ? (squads[m.home_team] || []) : null;
      const scorerAway = koEditable ? (squads[m.away_team] || []) : null;
      const scorerOptions = koEditable
        ? [].concat(scorerHome, scorerAway)
        : null;

      return {
        match_id: m.match_id,
        round: m.round,
        stage: m.stage,
        matchday: m.matchday,
        home_team: m.home_team,
        away_team: m.away_team,
        kickoff_datetime: m.kickoff_datetime instanceof Date
          ? m.kickoff_datetime.toISOString()
          : m.kickoff_datetime,
        status: m.status,
        locked,
        can_edit: !locked && !finished && !void_,
        show_all_predictions: showAll,
        // Results only revealed when finished
        home_score_90: finished ? m.home_score_90 : null,
        away_score_90: finished ? m.away_score_90 : null,
        actual_qualifier: finished ? m.actual_qualifier : null,
        actual_qualification_method: finished ? m.actual_qualification_method : null,
        actual_first_scorer: finished ? m.actual_first_scorer : null,
        // Player lists for the First Scorer dropdown
        scorer_options: scorerOptions,        // flat home+away (fallback)
        scorer_options_home: scorerHome,      // home squad (own labelled section)
        scorer_options_away: scorerAway,      // away squad (own labelled section)
        // User's own prediction (always visible to them)
        user_prediction: myPred ? {
          pred_home_score_90: myPred.pred_home_score_90,
          pred_away_score_90: myPred.pred_away_score_90,
          predicted_qualifier: myPred.predicted_qualifier,
          predicted_qualification_method: myPred.predicted_qualification_method,
          predicted_first_scorer: myPred.predicted_first_scorer,
          points: myPred.points
        } : null,
        // All predictions (only after lock)
        all_predictions
      };
    })
  };
}


// -------------------------------------------------------
// 3. submitPrediction
// -------------------------------------------------------

function submitPrediction(player_id, match_id, pred_home, pred_away, pred_qualifier, pred_method, pred_scorer) {
  const settings = getSettings();
  const lockMins = parseInt(settings.lock_minutes_before_kickoff) || 5;

  // Validate player
  const players = sheetToObjects(getSheet('Players'));
  const player = players.find(p =>
    p.player_id && p.player_id.toString() === player_id.toString() && p.active === true
  );
  if (!player) return { success: false, error: 'Player not found or inactive' };

  // Validate match
  const matches = sheetToObjects(getSheet('Matches'));
  const match = matches.find(m => m.match_id.toString() === match_id.toString());
  if (!match) return { success: false, error: 'Match not found' };
  if (match.status === 'void') return { success: false, error: 'Match is void' };
  if (match.status === 'finished') return { success: false, error: 'Match already finished' };

  // Lock check (server-side — cannot be bypassed by frontend)
  if (isLocked(match.kickoff_datetime, lockMins)) {
    auditLog('prediction_rejected_locked', player_id, match_id, null, null, 'frontend');
    return { success: false, error: 'This match is locked. Predictions are closed.' };
  }

  // Validate scores
  const homeScore = parseInt(pred_home);
  const awayScore = parseInt(pred_away);
  if (isNaN(homeScore) || isNaN(awayScore) || homeScore < 0 || awayScore < 0) {
    return { success: false, error: 'Scores must be non-negative numbers' };
  }

  // Knockout-specific validation
  if (match.stage === 'knockout') {
    if (!pred_qualifier || pred_qualifier.trim() === '') {
      return { success: false, error: 'Please select the qualifying team' };
    }
    const validMethods = ['90_minutes', 'extra_time', 'penalties'];
    if (!pred_method || !validMethods.includes(pred_method)) {
      return { success: false, error: 'Please select a valid qualification method' };
    }
    if (!pred_scorer || pred_scorer.toString().trim() === '') {
      return { success: false, error: 'Please select the first goalscorer' };
    }
  }

  const predSheet = getSheet('Predictions');
  const predData = predSheet.getDataRange().getValues();
  const headers = predData[0];
  const predictions = predData.slice(1);

  // Find existing row index
  const pidCol = headers.indexOf('player_id');
  const midCol = headers.indexOf('match_id');
  const existingIdx = predictions.findIndex(row =>
    row[pidCol] && row[midCol] &&
    row[pidCol].toString() === player_id.toString() &&
    row[midCol].toString() === match_id.toString()
  );

  const now = new Date();
  const qualifier = pred_qualifier || '';
  const method = (match.stage === 'knockout') ? pred_method : '';
  const firstScorer = (match.stage === 'knockout') ? (pred_scorer || '') : '';

  if (existingIdx >= 0) {
    // Update existing prediction (row is existingIdx + 2 due to header + 1-indexing)
    const rowNum = existingIdx + 2;
    const colMap = {};
    headers.forEach((h, i) => { colMap[h] = i + 1; });

    const oldHome = predictions[existingIdx][headers.indexOf('pred_home_score_90')];
    const oldAway = predictions[existingIdx][headers.indexOf('pred_away_score_90')];

    predSheet.getRange(rowNum, colMap['pred_home_score_90']).setValue(homeScore);
    predSheet.getRange(rowNum, colMap['pred_away_score_90']).setValue(awayScore);
    predSheet.getRange(rowNum, colMap['predicted_qualifier']).setValue(qualifier);
    predSheet.getRange(rowNum, colMap['predicted_qualification_method']).setValue(method);
    if (colMap['predicted_first_scorer']) {
      predSheet.getRange(rowNum, colMap['predicted_first_scorer']).setValue(firstScorer);
    }
    predSheet.getRange(rowNum, colMap['updated_at']).setValue(now);

    auditLog('prediction_updated', player_id, match_id,
      `${oldHome}-${oldAway}`, `${homeScore}-${awayScore}`, 'frontend');
    return { success: true, message: 'Prediction updated successfully' };

  } else {
    // Create new prediction — build the row dynamically from the header order
    // so it stays correct regardless of column positions.
    const rowArr = new Array(headers.length).fill('');
    const setCol = (name, val) => {
      const idx = headers.indexOf(name);
      if (idx >= 0) rowArr[idx] = val;
    };
    const newId = 'P' + Date.now();
    setCol('prediction_id', newId);
    setCol('player_id', player_id);
    setCol('match_id', match_id);
    setCol('pred_home_score_90', homeScore);
    setCol('pred_away_score_90', awayScore);
    setCol('predicted_qualifier', qualifier);
    setCol('predicted_qualification_method', method);
    setCol('predicted_first_scorer', firstScorer);
    setCol('submitted_at', now);
    setCol('updated_at', now);
    setCol('points', '');

    predSheet.appendRow(rowArr);

    auditLog('prediction_created', player_id, match_id,
      null, `${homeScore}-${awayScore}`, 'frontend');
    return { success: true, message: 'Prediction saved successfully' };
  }
}


// -------------------------------------------------------
// 4. getLeaderboard
// -------------------------------------------------------

function getLeaderboard() {
  const data = sheetToObjects(getSheet('Leaderboard'));
  return { success: true, leaderboard: data };
}


// -------------------------------------------------------
// 5. getMatchPredictions  (legacy — kept for compatibility)
//    getMatches() now returns all_predictions inline for locked/finished matches
// -------------------------------------------------------

function getMatchPredictions(match_id, player_id) {
  const settings = getSettings();
  const lockMins = parseInt(settings.lock_minutes_before_kickoff) || 5;

  const matches = sheetToObjects(getSheet('Matches'));
  const match = matches.find(m => m.match_id.toString() === match_id.toString());
  if (!match) return { success: false, error: 'Match not found' };

  const locked = isLocked(match.kickoff_datetime, lockMins);
  const predictions = sheetToObjects(getSheet('Predictions'));
  const players = sheetToObjects(getSheet('Players'));
  const playerNames = {};
  players.forEach(p => { playerNames[p.player_id.toString()] = p.name; });

  const matchPreds = predictions.filter(p => p.match_id.toString() === match_id.toString());

  if (!locked) {
    // Only own prediction
    const mine = matchPreds.find(p => p.player_id.toString() === player_id.toString());
    return {
      success: true, locked: false,
      predictions: mine ? [{
        player_name: 'You', is_me: true,
        pred_home_score_90: mine.pred_home_score_90,
        pred_away_score_90: mine.pred_away_score_90,
        predicted_qualifier: mine.predicted_qualifier,
        predicted_qualification_method: mine.predicted_qualification_method,
        predicted_first_scorer: mine.predicted_first_scorer,
        points: mine.points
      }] : []
    };
  }

  // Locked: return all
  return {
    success: true, locked: true,
    predictions: matchPreds.map(p => ({
      player_name: playerNames[p.player_id.toString()] || 'Unknown',
      is_me: p.player_id.toString() === player_id.toString(),
      pred_home_score_90: p.pred_home_score_90,
      pred_away_score_90: p.pred_away_score_90,
      predicted_qualifier: p.predicted_qualifier,
      predicted_qualification_method: p.predicted_qualification_method,
      predicted_first_scorer: p.predicted_first_scorer,
      points: p.points
    }))
  };
}


// -------------------------------------------------------
// 6. recalculateAll  (also callable by trigger)
// -------------------------------------------------------

function recalculateAll() {
  const predSheet = getSheet('Predictions');
  const matchSheet = getSheet('Matches');
  const playerSheet = getSheet('Players');
  const lbSheet = getSheet('Leaderboard');

  const matches = sheetToObjects(matchSheet);
  const players = sheetToObjects(playerSheet).filter(p => p.active === true);
  const predData = predSheet.getDataRange().getValues();
  const predHeaders = predData[0];
  const predictions = predData.slice(1);

  // Map of finished matches
  const finishedMap = {};
  matches.forEach(m => {
    if (m.status === 'finished') finishedMap[m.match_id.toString()] = m;
  });

  const totalFinished = Object.keys(finishedMap).length;

  // Column indexes in Predictions sheet
  const colIdx = {};
  predHeaders.forEach((h, i) => { colIdx[h] = i; });

  // Re-score every prediction for a finished match
  predictions.forEach((row, i) => {
    const mid = row[colIdx['match_id']]?.toString();
    const match = finishedMap[mid];
    if (!match) return;

    const pred = {};
    predHeaders.forEach((h, j) => { pred[h] = row[j]; });
    const pts = calculatePoints(pred, match);

    predSheet.getRange(i + 2, colIdx['points'] + 1).setValue(pts);
  });

  // Build player stats
  const stats = {};
  players.forEach(p => {
    stats[p.player_id.toString()] = {
      player_id: p.player_id, player_name: p.name,
      total_points: 0, exact_scores: 0, correct_outcomes: 0,
      qualifier_hits: 0, method_hits: 0,
      predictions_submitted: 0, missed_predictions: 0,
      last_matchday_points: 0, most_wrong_prediction: '',
      _worst_wrongness: -1, updated_at: new Date()
    };
  });

  // Sort matchdays to find the last one
  const matchdays = matches.map(m => m.matchday).filter(Boolean);
  const lastMatchday = matchdays.length ? Math.max(...matchdays.map(Number)) : null;

  // Re-read predictions with updated points
  const freshData = predSheet.getDataRange().getValues();
  const freshHeaders = freshData[0];
  const freshPreds = freshData.slice(1);
  const fIdx = {};
  freshHeaders.forEach((h, i) => { fIdx[h] = i; });

  freshPreds.forEach(row => {
    const pid = row[fIdx['player_id']]?.toString();
    const mid = row[fIdx['match_id']]?.toString();
    const stat = stats[pid];
    if (!stat) return;

    const match = finishedMap[mid];
    if (!match) return;

    stat.predictions_submitted++;
    const pts = parseFloat(row[fIdx['points']]) || 0;
    stat.total_points += pts;

    if (match.matchday?.toString() === lastMatchday?.toString()) {
      stat.last_matchday_points += pts;
    }

    const pH = parseInt(row[fIdx['pred_home_score_90']]);
    const pA = parseInt(row[fIdx['pred_away_score_90']]);
    const aH = parseInt(match.home_score_90);
    const aA = parseInt(match.away_score_90);

    if (pH === aH && pA === aA) {
      stat.exact_scores++;
    } else if (getOutcome(pH, pA) === getOutcome(aH, aA)) {
      stat.correct_outcomes++;
    }

    if (match.stage === 'knockout') {
      const pq = row[fIdx['predicted_qualifier']];
      const pm = row[fIdx['predicted_qualification_method']];
      if (pq && pq === match.actual_qualifier) {
        stat.qualifier_hits++;
        if (pm && normMethod(pm) === normMethod(match.actual_qualification_method)) stat.method_hits++;
      }
    }

    // Wrongness (only for finished + predicted)
    if (!isNaN(pH) && !isNaN(pA) && !isNaN(aH) && !isNaN(aA)) {
      const w = Math.abs(pH - aH) + Math.abs(pA - aA);
      if (w > stat._worst_wrongness) {
        stat._worst_wrongness = w;
        stat.most_wrong_prediction =
          `${match.home_team} ${pH}-${pA} ${match.away_team} → actual ${aH}-${aA}`;
      }
    }
  });

  // Missed predictions
  players.forEach(p => {
    const s = stats[p.player_id.toString()];
    s.missed_predictions = Math.max(0, totalFinished - s.predictions_submitted);
  });

  // Sort by points desc
  const sorted = Object.values(stats).sort((a, b) => {
    if (b.total_points !== a.total_points) return b.total_points - a.total_points;
    if (b.exact_scores !== a.exact_scores) return b.exact_scores - a.exact_scores;
    return b.correct_outcomes - a.correct_outcomes;
  });

  // Rewrite Leaderboard
  lbSheet.clearContents();
  lbSheet.appendRow([
    'rank','player_id','player_name','total_points','exact_scores','correct_outcomes',
    'qualifier_hits','method_hits','predictions_submitted','missed_predictions',
    'last_matchday_points','most_wrong_prediction','updated_at'
  ]);
  sorted.forEach((s, i) => {
    lbSheet.appendRow([
      i + 1, s.player_id, s.player_name, s.total_points,
      s.exact_scores, s.correct_outcomes, s.qualifier_hits, s.method_hits,
      s.predictions_submitted, s.missed_predictions, s.last_matchday_points,
      s.most_wrong_prediction, s.updated_at
    ]);
  });

  auditLog('points_recalculated', null, null, null, null, 'trigger');
  auditLog('leaderboard_recalculated', null, null, null, null, 'trigger');

  return { success: true, message: 'Recalculation complete' };
}

// Normalise a qualification-method string so the sheet's human labels
// ("full time", "after extra time", "penalty shootout", …) match the app's
// tokens (90_minutes / extra_time / penalties). Prevents method bonuses from
// being missed due to a label mismatch.
function normMethod(s) {
  if (!s && s !== 0) return '';
  let t = s.toString().trim().toLowerCase().replace(/[^a-z0-9 ]+/g, ' ').replace(/\s+/g, ' ').trim();
  if (!t) return '';
  if (t.indexOf('pen') !== -1 || t.indexOf('shoot') !== -1) return 'penalties';
  if (t.indexOf('extra') !== -1 || t === 'aet' || t === 'et' || t.indexOf('aet') !== -1) return 'extra_time';
  if (t.indexOf('90') !== -1 || t.indexOf('full') !== -1 || t.indexOf('regular') !== -1 || t === 'ft' || t === 'reg' || t.indexOf('normal') !== -1) return '90_minutes';
  return t;
}

function calculatePoints(pred, match) {
  const pH = parseInt(pred.pred_home_score_90);
  const pA = parseInt(pred.pred_away_score_90);
  const aH = parseInt(match.home_score_90);
  const aA = parseInt(match.away_score_90);

  if ([pH, pA, aH, aA].some(isNaN)) return 0;

  let pts = 0;

  if (pH === aH && pA === aA) {
    // Exact score — 5 pts, no bonuses
    pts += 5;
  } else {
    // Correct outcome 1/X/2
    if (getOutcome(pH, pA) === getOutcome(aH, aA)) pts += 1;
    // Correct goals home team
    if (pH === aH) pts += 1;
    // Correct goals away team
    if (pA === aA) pts += 1;
    // Correct goal difference (signed)
    if ((pH - pA) === (aH - aA)) pts += 1;
  }

  if (match.stage === 'knockout') {
    // Qualifier (+2) and method (+2, only if qualifier correct)
    if (pred.predicted_qualifier && pred.predicted_qualifier === match.actual_qualifier) {
      pts += 2;
      if (pred.predicted_qualification_method &&
          normMethod(pred.predicted_qualification_method) === normMethod(match.actual_qualification_method)) {
        pts += 2;
      }
    }
    // First goalscorer (+3) — exact string match.
    // Handles "No Goalscorer / 0-0" and "Own Goal" identically to player names.
    if (pred.predicted_first_scorer && match.actual_first_scorer &&
        pred.predicted_first_scorer.toString().trim() === match.actual_first_scorer.toString().trim()) {
      pts += 3;
    }
  }

  return pts;
}


// -------------------------------------------------------
// 7. generateReminderMessage  (WhatsApp-ready plain text)
// -------------------------------------------------------

function generateReminderMessage() {
  const settings = getSettings();
  const lockMins = parseInt(settings.lock_minutes_before_kickoff) || 5;
  const siteUrl = settings.prediction_website_url || '[site URL not configured]';
  const tz = Session.getScriptTimeZone();

  const now = new Date();
  const matches = sheetToObjects(getSheet('Matches'));
  const predictions = sheetToObjects(getSheet('Predictions')).filter(p => p.player_id && p.match_id);
  const players = sheetToObjects(getSheet('Players')).filter(p => p.active === true);
  const leaderboard = sheetToObjects(getSheet('Leaderboard'));

  // Group stage only
  const groupMatches = matches.filter(m => m.stage === 'group' && m.match_id);
  const totalGroupMatches = groupMatches.length;
  const groupMatchIds = new Set(groupMatches.map(m => m.match_id.toString()));

  // Predictions per player (group stage only)
  const predCount = {};
  players.forEach(p => { predCount[p.player_id.toString()] = 0; });
  predictions.forEach(p => {
    if (p.player_id && groupMatchIds.has(p.match_id.toString())) {
      const pid = p.player_id.toString();
      if (predCount[pid] !== undefined) predCount[pid]++;
    }
  });

  // Predictions in last 24h
  const since24h = new Date(now.getTime() - 24 * 3600 * 1000);
  const recent = [];
  predictions.forEach(p => {
    if (!p.player_id || !p.match_id) return;
    const updated = new Date(p.updated_at || p.submitted_at);
    if (updated >= since24h) {
      const player = players.find(pl => pl.player_id.toString() === p.player_id.toString());
      const match = matches.find(m => m.match_id.toString() === p.match_id.toString());
      if (player && match) recent.push({ name: player.name, match: match.home_team + ' vs ' + match.away_team, score: p.pred_home_score_90 + '-' + p.pred_away_score_90, updated });
    }
  });
  recent.sort((a, b) => b.updated - a.updated);

  // Next kickoff
  const nextMatch = matches
    .filter(m => m.status === 'upcoming')
    .sort((a, b) => new Date(a.kickoff_datetime) - new Date(b.kickoff_datetime))[0];

  let msg = `🥩🏆 *S&S PREDICTIONS — DAILY DIGEST*\n`;
  msg += `━━━━━━━━━━━━━━━━━━━━━━━━\n\n`;

  // Player prediction counts
  msg += `📊 *Group stage predictions (/${totalGroupMatches}):*\n`;
  const sorted = players.slice().sort((a, b) =>
    (predCount[b.player_id.toString()] || 0) - (predCount[a.player_id.toString()] || 0));
  sorted.forEach(p => {
    const cnt = predCount[p.player_id.toString()] || 0;
    const icon = cnt === 0 ? '❌' : cnt === totalGroupMatches ? '🏆' : '✅';
    msg += `${icon} ${p.name}: *${cnt}/${totalGroupMatches}*\n`;
  });

  // Next kickoff
  msg += `\n━━━━━━━━━━━━━━━━━━━━━━━━\n`;
  if (nextMatch) {
    const kickoff = new Date(nextMatch.kickoff_datetime);
    const lockTime = new Date(kickoff.getTime() - lockMins * 60 * 1000);
    const diffMs = lockTime - now;
    const kickoffStr = Utilities.formatDate(kickoff, tz, 'EEE d MMM, HH:mm');
    if (diffMs > 0) {
      const diffH = Math.floor(diffMs / 3600000);
      const diffM = Math.floor((diffMs % 3600000) / 60000);
      msg += `⏰ *Next:* ${nextMatch.home_team} vs ${nextMatch.away_team}\n`;
      msg += `   📅 ${kickoffStr} · 🔒 Lock in ${diffH}h ${diffM}m\n`;
    }
  }

  // Last 24h
  msg += `\n━━━━━━━━━━━━━━━━━━━━━━━━\n`;
  msg += `🕐 *Last 24h activity:*\n`;
  if (recent.length === 0) {
    msg += `   😴 No predictions in the last 24 hours\n`;
  } else {
    recent.slice(0, 8).forEach(r => {
      msg += `   ⚽ ${r.name} → ${r.match} *${r.score}*\n`;
    });
  }

  msg += `\n━━━━━━━━━━━━━━━━━━━━━━━━\n`;
  msg += `🔗 ${siteUrl}\n`;
  msg += `━━━━━━━━━━━━━━━━━━━━━━━━`;

  return msg;
}

// WhatsApp-ready overnight update block
function generateOvernightBlock() {
  const tz = Session.getScriptTimeZone();
  const now = new Date();
  const since = new Date(now.getTime() - 18 * 3600 * 1000); // last 18h

  const matches = sheetToObjects(getSheet('Matches'));
  const predictions = sheetToObjects(getSheet('Predictions')).filter(p => p.player_id && p.match_id);
  const players = sheetToObjects(getSheet('Players')).filter(p => p.active === true);
  const leaderboard = sheetToObjects(getSheet('Leaderboard'));

  // Matches finished in last 18h
  const overnightMatches = matches.filter(m => {
    if (m.status !== 'finished') return false;
    const ko = new Date(m.kickoff_datetime);
    return ko >= since && ko <= now;
  }).sort((a, b) => new Date(a.kickoff_datetime) - new Date(b.kickoff_datetime));

  if (overnightMatches.length === 0) return null;

  // Points per player from overnight matches
  const overnightIds = new Set(overnightMatches.map(m => m.match_id.toString()));
  const playerPts = {};
  players.forEach(p => { playerPts[p.player_id.toString()] = { name: p.name, pts: 0, details: [] }; });

  predictions.forEach(p => {
    if (!p.player_id || !p.match_id) return;
    if (!overnightIds.has(p.match_id.toString())) return;
    const pid = p.player_id.toString();
    if (!playerPts[pid]) return;
    const pts = parseFloat(p.points) || 0;
    playerPts[pid].pts += pts;
    const match = overnightMatches.find(m => m.match_id.toString() === p.match_id.toString());
    if (match && pts > 0) {
      playerPts[pid].details.push(`${match.home_team.split(' ')[0]} +${pts}`);
    }
  });

  const medals = ['🥇','🥈','🥉'];

  let block = `🌙 *OVERNIGHT UPDATE*\n`;
  block += `━━━━━━━━━━━━━━━━━━━━━━━━\n\n`;

  // Results
  block += `⚽ *Last night\'s results:*\n`;
  overnightMatches.forEach(m => {
    block += `   ${m.home_team} *${m.home_score_90}-${m.away_score_90}* ${m.away_team}\n`;
  });

  // Points per player
  block += `\n📊 *Points earned:*\n`;
  const sortedPlayers = Object.values(playerPts).sort((a, b) => b.pts - a.pts);
  sortedPlayers.forEach(p => {
    const icon = p.pts > 0 ? '⭐' : '—';
    const detail = p.details.length > 0 ? ' (' + p.details.join(', ') + ')' : '';
    block += `${icon} ${p.name}: *+${p.pts} pts*${detail}\n`;
  });

  // Current standings
  block += `\n🏅 *Current standings:*\n`;
  leaderboard.slice(0, players.length).forEach((p, i) => {
    const medal = i < 3 ? medals[i] : `${i+1}.`;
    block += `${medal} ${p.player_name}: *${p.total_points} pts*\n`;
  });

  block += `━━━━━━━━━━━━━━━━━━━━━━━━`;
  return block;
}


// -------------------------------------------------------
// 8. sendDailyAdminReminder  (called by time trigger)
// -------------------------------------------------------

function sendDailyAdminReminder() {
  const settings = getSettings();
  const adminEmail = settings.admin_email;

  if (!adminEmail || adminEmail === '[ADMIN_EMAIL_HERE]') {
    Logger.log('No admin email configured — skipping reminder');
    return;
  }

  const msg = generateReminderMessage();
  const overnightBlock = generateOvernightBlock();
  const tz = Session.getScriptTimeZone();
  const now = new Date();
  const matches = sheetToObjects(getSheet('Matches'));
  const predictions = sheetToObjects(getSheet('Predictions')).filter(p => p.player_id && p.match_id);
  const players = sheetToObjects(getSheet('Players')).filter(p => p.active === true);
  const lockMins = parseInt(settings.lock_minutes_before_kickoff) || 5;

  const groupMatches = matches.filter(m => m.stage === 'group' && m.match_id);
  const totalGroupMatches = groupMatches.length;
  const groupMatchIds = new Set(groupMatches.map(m => m.match_id.toString()));

  const predCount = {};
  players.forEach(p => { predCount[p.player_id.toString()] = 0; });
  predictions.forEach(p => {
    if (p.player_id && groupMatchIds.has(p.match_id.toString())) {
      const pid = p.player_id.toString();
      if (predCount[pid] !== undefined) predCount[pid]++;
    }
  });

  const sortedPlayers = players.slice().sort((a, b) =>
    (predCount[b.player_id.toString()] || 0) - (predCount[a.player_id.toString()] || 0));

  const playerRows = sortedPlayers.map(p => {
    const cnt = predCount[p.player_id.toString()] || 0;
    const pct = Math.round((cnt / totalGroupMatches) * 100);
    const color = cnt === 0 ? '#f87171' : cnt === totalGroupMatches ? '#4ade80' : '#c9a227';
    return `<tr>
      <td style="padding:8px 12px;font-size:13px;color:#fff;font-weight:600">${p.name}</td>
      <td style="padding:8px 12px;font-size:14px;color:${color};font-weight:700;text-align:right">${cnt}/${totalGroupMatches}</td>
      <td style="padding:8px 12px;width:120px">
        <div style="background:#1c2d54;border-radius:4px;height:6px">
          <div style="background:${color};border-radius:4px;height:6px;width:${pct}%"></div>
        </div>
      </td>
    </tr>`;
  }).join('');

  // Build overnight HTML section
  let overnightHtml = '';
  if (overnightBlock) {
    overnightHtml = `<div style="margin-bottom:20px">`
      + `<div style="font-size:11px;color:#7a90b8;text-transform:uppercase;letter-spacing:1px;margin-bottom:8px">🌙 Overnight Update — tap & copy to WhatsApp</div>`
      + `<div style="background:#0a1628;border:2px solid rgba(201,162,39,0.4);border-radius:8px;padding:16px;font-size:13px;color:#fff;line-height:1.9;white-space:pre-wrap;font-family:'Courier New',monospace">${overnightBlock.replace(/\*/g,'')}</div>`
      + `</div>`;
  }

  const nextMatch = matches.filter(m => m.status === 'upcoming')
    .sort((a, b) => new Date(a.kickoff_datetime) - new Date(b.kickoff_datetime))[0];
  let nextMatchHtml = '';
  if (nextMatch) {
    const kickoff = new Date(nextMatch.kickoff_datetime);
    const lockTime = new Date(kickoff.getTime() - lockMins * 60 * 1000);
    const diffMs = lockTime - now;
    const kickoffStr = Utilities.formatDate(kickoff, tz, 'EEE d MMM, HH:mm');
    const countdown = diffMs > 0
      ? Math.floor(diffMs/3600000) + 'h ' + Math.floor((diffMs%3600000)/60000) + 'm to lock'
      : 'LOCKED';
    nextMatchHtml = `<div style="background:#152240;border:1px solid rgba(201,162,39,0.2);border-radius:8px;padding:12px 16px;margin-bottom:20px">
      <div style="font-size:11px;color:#7a90b8;text-transform:uppercase;letter-spacing:1px;margin-bottom:6px">Next Kickoff</div>
      <div style="font-size:15px;font-weight:700;color:#fff">${nextMatch.home_team} vs ${nextMatch.away_team}</div>
      <div style="font-size:12px;color:#c9a227;margin-top:4px">📅 ${kickoffStr} &nbsp;·&nbsp; 🔒 ${countdown}</div>
    </div>`;
  }

  const since24h = new Date(now.getTime() - 24 * 3600 * 1000);
  const recentItems = [];
  predictions.forEach(p => {
    if (!p.player_id || !p.match_id) return;
    const updated = new Date(p.updated_at || p.submitted_at);
    if (updated >= since24h) {
      const player = players.find(pl => pl.player_id.toString() === p.player_id.toString());
      const match = matches.find(m => m.match_id.toString() === p.match_id.toString());
      if (player && match) recentItems.push({ name: player.name, match: match.home_team + ' vs ' + match.away_team, score: p.pred_home_score_90 + '-' + p.pred_away_score_90 });
    }
  });

  const recentHtml = recentItems.length === 0
    ? '<p style="color:#7a90b8;font-size:13px;margin:0">No predictions in the last 24 hours.</p>'
    : recentItems.slice(0, 10).map(r => `<div style="display:flex;justify-content:space-between;padding:7px 0;border-bottom:1px solid rgba(255,255,255,0.05)">
        <span style="font-size:13px;color:#adc0dc"><strong style="color:#fff">${r.name}</strong> · ${r.match}</span>
        <span style="font-size:13px;font-weight:700;color:#c9a227;margin-left:12px">${r.score}</span>
      </div>`).join('');

  const htmlBody = `<div style="font-family:'Segoe UI',Arial,sans-serif;max-width:600px;margin:0 auto;background:#0a1628;color:#fff;border-radius:12px;overflow:hidden">
    <div style="background:#0e1a35;border-bottom:3px solid #c9a227;padding:20px 24px">
      <h2 style="color:#c9a227;margin:0 0 2px;font-size:18px">Sportsbook &amp; Steak</h2>
      <p style="color:#7a90b8;font-size:12px;margin:0;text-transform:uppercase;letter-spacing:1px">Daily Digest · ${Utilities.formatDate(now, tz, 'EEE d MMM yyyy')}</p>
    </div>
    <div style="padding:20px 24px">
      ${overnightHtml}
      ${nextMatchHtml}
      <div style="font-size:11px;color:#7a90b8;text-transform:uppercase;letter-spacing:1px;margin-bottom:10px">Group Stage Predictions (/${totalGroupMatches})</div>
      <table style="width:100%;border-collapse:collapse;background:#0e1a35;border-radius:8px;overflow:hidden;margin-bottom:20px">${playerRows}</table>
      <div style="font-size:11px;color:#7a90b8;text-transform:uppercase;letter-spacing:1px;margin-bottom:10px">Last 24 Hours</div>
      <div style="background:#0e1a35;border-radius:8px;padding:12px 16px;margin-bottom:20px">${recentHtml}</div>
      <div style="background:#152240;border-radius:8px;padding:12px 16px">
        <p style="font-size:11px;color:#7a90b8;margin:0 0 4px">Prediction site</p>
        <a href="${settings.prediction_website_url}" style="color:#c9a227;font-size:13px;font-weight:600">${settings.prediction_website_url}</a>
      </div>
      <p style="color:#4a5568;font-size:11px;margin-top:16px;text-align:center">Automated daily digest · Sportsbook &amp; Steak Predictions</p>
    </div>
  </div>`;

  MailApp.sendEmail({
    to: adminEmail,
    subject: '🥩 S&S Daily Digest — ' + Utilities.formatDate(now, tz, 'd MMM'),
    htmlBody: htmlBody,
    body: msg
  });

  Logger.log('Admin digest sent to ' + adminEmail);
}


// -------------------------------------------------------
// SETUP: Run once manually after pasting this code
// -------------------------------------------------------

function createTriggers() {
  // Remove old triggers first
  ScriptApp.getProjectTriggers().forEach(t => ScriptApp.deleteTrigger(t));

  // Daily admin email at 09:00
  ScriptApp.newTrigger('sendDailyAdminReminder')
    .timeBased().everyDays(1).atHour(9).create();

  // Recalculate every 30 min
  // NOTE: Apps Script free tier allows ~6 triggers and 6h/day execution.
  // 30-min interval is safer than 10-min for staying within limits.
  // Alternatively, just run recalculateAll() manually after entering results.
  ScriptApp.newTrigger('recalculateAll')
    .timeBased().everyMinutes(30).create();

  Logger.log('✅ Triggers created: daily reminder + 30-min recalc');
}


// -------------------------------------------------------
// SAMPLE DATA SETUP: Run once to populate the sheet
// -------------------------------------------------------

function populateSampleData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // --- Players ---
  const ps = ss.getSheetByName('Players');
  ps.clearContents();
  ps.appendRow(['player_id','name','pin','email','active','created_at']);
  const players = [
    ['P1','Marco','1234','','TRUE', new Date()],
    ['P2','Luca','2345','','TRUE', new Date()],
    ['P3','Sofia','3456','','TRUE', new Date()],
    ['P4','Elena','4567','','TRUE', new Date()],
    ['P5','Davide','5678','','TRUE', new Date()],
    ['P6','Giulia','6789','','TRUE', new Date()],
    ['P7','Matteo','7890','','TRUE', new Date()],
  ];
  players.forEach(r => ps.appendRow(r));

  // --- Matches ---
  const ms = ss.getSheetByName('Matches');
  ms.clearContents();
  ms.appendRow(['match_id','round','stage','matchday','home_team','away_team','kickoff_datetime','status','home_score_90','away_score_90','actual_qualifier','actual_qualification_method','actual_first_scorer']);

  const now = new Date();
  const d = (addDays, hour=19) => {
    const dt = new Date(now);
    dt.setDate(dt.getDate() + addDays);
    dt.setHours(hour, 0, 0, 0);
    return dt;
  };

  const matchRows = [
    ['M1','Group A','group',1,'Brazil','Germany',d(0,19),'upcoming','','','','',''],
    ['M2','Group B','group',1,'France','Argentina',d(0,21),'upcoming','','','','',''],
    ['M3','Group A','group',2,'Spain','Portugal',d(2,19),'upcoming','','','','',''],
    ['M4','Group C','group',1,'England','Italy',d(2,21),'upcoming','','','','',''],
    ['M5','Quarter-Final','knockout',5,'Brazil','France',d(7,19),'upcoming','','','','',''],
    ['M6','Semi-Final','knockout',6,'Spain','England',d(14,19),'upcoming','','','','',''],
  ];
  matchRows.forEach(r => ms.appendRow(r));

  // --- Settings ---
  const sets = ss.getSheetByName('Settings');
  sets.clearContents();
  sets.appendRow(['setting','value']);
  const settingRows = [
    ['lock_minutes_before_kickoff', 5],
    ['show_predictions_after_lock', 'TRUE'],
    ['scoring_exact_score', 5],
    ['scoring_correct_outcome', 1],
    ['scoring_qualifier', 2],
    ['scoring_qualification_method', 2],
    ['scoring_first_scorer', 3],
    ['admin_email', '[ADMIN_EMAIL_HERE]'],
    ['prediction_website_url', '[GITHUB_PAGES_URL_HERE]'],
    ['reminder_window_hours', 24],
    ['daily_reminder_time', '09:00'],
  ];
  settingRows.forEach(r => sets.appendRow(r));

  Logger.log('✅ Sample data populated. Update admin_email and prediction_website_url in Settings tab!');
}


// ============================================================
// AUTO SYNC RESULTS FROM football-data.org
// Paste this at the bottom of your Code.gs
// Then run syncResults() manually or it runs via trigger
// ============================================================

// Configure this once in Apps Script under Project Settings > Script Properties:
// FOOTBALL_API_KEY = your football-data.org token
// Never commit API tokens to the repository.
function getFootballApiKey() {
  const key = PropertiesService.getScriptProperties().getProperty('FOOTBALL_API_KEY');
  if (!key) throw new Error('Missing FOOTBALL_API_KEY script property');
  return key;
}

// Team name mapping: football-data.org name → your Sheet name
// Add more here if results don't match
const TEAM_NAME_MAP = {
  // Confirmed from API (football-data.org → Sheet names)
  'Bosnia-Herzegovina':        'Bosnia & Herzegovina',
  'Bosnia and Herzegovina':    'Bosnia & Herzegovina',
  'Turkey':                    'Türkiye',
  'Cape Verde Islands':        'Cape Verde',
  'Cabo Verde':                'Cape Verde',
  'United States':             'USA',
  'Congo DR':                  'DR Congo',
  // Pre-emptive
  "Côte d'Ivoire":             'Ivory Coast',
  'Czech Republic':            'Czechia',
  'IR Iran':                   'Iran',
  'Korea Republic':            'South Korea',
};

function syncResults() {
  try {
    const url = 'https://api.football-data.org/v4/competitions/WC/matches?status=FINISHED&season=2026';

    const response = UrlFetchApp.fetch(url, {
      headers: { 'X-Auth-Token': getFootballApiKey() },
      muteHttpExceptions: true
    });

    const code = response.getResponseCode();
    if (code !== 200) {
      Logger.log('API error: HTTP ' + code + ' — ' + response.getContentText());
      return;
    }

    const data = JSON.parse(response.getContentText());
    const apiMatches = data.matches || [];

    if (apiMatches.length === 0) {
      Logger.log('No finished matches returned from API yet.');
      return;
    }

    const matchSheet = getSheet('Matches');
    const sheetData  = matchSheet.getDataRange().getValues();
    const headers    = sheetData[0];

    const colIdx = {};
    headers.forEach((h, i) => { colIdx[h] = i; });

    let updated = 0;

    apiMatches.forEach(am => {
      if (am.status !== 'FINISHED') return;

      // Normalise team names
      let apiHome = mapTeam(am.homeTeam.name);
      let apiAway = mapTeam(am.awayTeam.name);

      // Determine 90-min score
      // regularTime is the 90-min score; fullTime includes ET
      const homeScore90 = (am.score.regularTime && am.score.regularTime.home !== null)
        ? am.score.regularTime.home
        : am.score.fullTime.home;
      const awayScore90 = (am.score.regularTime && am.score.regularTime.away !== null)
        ? am.score.regularTime.away
        : am.score.fullTime.away;

      // Determine qualification method for knockout
      let qualMethod = '';
      let qualifier  = '';
      const duration = am.score.duration; // REGULAR, EXTRA_TIME, PENALTY_SHOOTOUT
      const winner   = am.score.winner;   // HOME_TEAM, AWAY_TEAM, DRAW

      if (duration === 'REGULAR')          qualMethod = '90_minutes';
      else if (duration === 'EXTRA_TIME')  qualMethod = 'extra_time';
      else if (duration === 'PENALTY_SHOOTOUT') qualMethod = 'penalties';

      if (winner === 'HOME_TEAM') qualifier = apiHome;
      else if (winner === 'AWAY_TEAM') qualifier = apiAway;

      // Find matching row in sheet
      for (let i = 1; i < sheetData.length; i++) {
        const row = sheetData[i];
        const shHome   = row[colIdx['home_team']] ? row[colIdx['home_team']].toString().trim() : '';
        const shAway   = row[colIdx['away_team']] ? row[colIdx['away_team']].toString().trim() : '';
        const shStatus = row[colIdx['status']] ? row[colIdx['status']].toString().trim() : '';

        if (shHome === apiHome && shAway === apiAway && shStatus !== 'finished') {
          const rowNum = i + 1;

          matchSheet.getRange(rowNum, colIdx['home_score_90'] + 1).setValue(homeScore90);
          matchSheet.getRange(rowNum, colIdx['away_score_90'] + 1).setValue(awayScore90);
          matchSheet.getRange(rowNum, colIdx['status'] + 1).setValue('finished');

          // Knockout only
          const stage = row[colIdx['stage']] ? row[colIdx['stage']].toString() : '';
          if (stage === 'knockout' && qualifier) {
            matchSheet.getRange(rowNum, colIdx['actual_qualifier'] + 1).setValue(qualifier);
            matchSheet.getRange(rowNum, colIdx['actual_qualification_method'] + 1).setValue(qualMethod);
          }

          Logger.log('✅ Updated: ' + apiHome + ' ' + homeScore90 + '-' + awayScore90 + ' ' + apiAway
            + (qualifier ? ' | Qualifier: ' + qualifier + ' (' + qualMethod + ')' : ''));
          updated++;
          break;
        }
      }
    });

    if (updated > 0) {
      Logger.log('Synced ' + updated + ' results — running recalculateAll...');
      recalculateAll();
      Logger.log('✅ Done. Leaderboard updated.');
    } else {
      Logger.log('No new results to sync (all already marked finished or no matches yet).');
    }

  } catch (err) {
    Logger.log('syncResults error: ' + err.message);
  }
}

function mapTeam(name) {
  if (!name) return '';
  return TEAM_NAME_MAP[name.trim()] || name.trim();
}

// -------------------------------------------------------
// Run once to add syncResults to the trigger schedule
// (replaces existing triggers — run createTriggers again
//  if you want to keep the existing ones)
// -------------------------------------------------------
function addSyncTrigger() {
  // Check if trigger already exists
  const existing = ScriptApp.getProjectTriggers();
  const alreadyExists = existing.some(t => t.getHandlerFunction() === 'syncResults');
  if (alreadyExists) {
    Logger.log('syncResults trigger already exists');
    return;
  }
  ScriptApp.newTrigger('syncResults')
    .timeBased()
    .everyMinutes(60)
    .create();
  Logger.log('✅ syncResults trigger added — runs every 60 minutes');
}

// -------------------------------------------------------
// DEBUG: Run once to see exact team names from API
// -------------------------------------------------------

function checkTeamNames() {
  const url = 'https://api.football-data.org/v4/competitions/WC/matches?season=2026';
  const response = UrlFetchApp.fetch(url, {
    headers: { 'X-Auth-Token': getFootballApiKey() },
    muteHttpExceptions: true
  });
  const data = JSON.parse(response.getContentText());
  const matches = data.matches || [];
  Logger.log('Total matches from API: ' + matches.length);
  matches.forEach(m => {
    Logger.log(m.homeTeam.name + ' vs ' + m.awayTeam.name + ' [' + m.status + ']');
  });
}

// -------------------------------------------------------
// ONE-TIME: populate Round of 32 real team names
// (already run — kept for reference)
// -------------------------------------------------------

function populateR32Teams() {
  const R32_MAP = {
    "2026-06-28 21:00:00": ["South Africa", "Canada"],
    "2026-06-29 19:00:00": ["Brazil", "Japan"],
    "2026-06-29 22:30:00": ["Germany", "Paraguay"],
    "2026-06-30 03:00:00": ["Netherlands", "Morocco"],
    "2026-06-30 19:00:00": ["Ivory Coast", "Norway"],
    "2026-06-30 23:00:00": ["France", "Sweden"],
    "2026-07-01 03:00:00": ["Mexico", "Ecuador"],
    "2026-07-01 18:00:00": ["England", "DR Congo"],
    "2026-07-01 22:00:00": ["Belgium", "Senegal"],
    "2026-07-02 02:00:00": ["USA", "Bosnia & Herzegovina"],
    "2026-07-02 21:00:00": ["Spain", "Austria"],
    "2026-07-03 01:00:00": ["Portugal", "Croatia"],
    "2026-07-03 05:00:00": ["Switzerland", "Algeria"],
    "2026-07-03 20:00:00": ["Australia", "Egypt"],
    "2026-07-04 00:00:00": ["Argentina", "Cape Verde"],
    "2026-07-04 03:30:00": ["Colombia", "Ghana"]
  };

  const sheet = getSheet('Matches');
  const data = sheet.getDataRange().getValues();
  const headers = data[0];

  const colIdx = {};
  headers.forEach((h, i) => { colIdx[h] = i; });

  const koCol = colIdx['kickoff_datetime'];
  const homeCol = colIdx['home_team'];
  const awayCol = colIdx['away_team'];
  const tz = Session.getScriptTimeZone();

  let updated = 0;
  let notFound = [];

  for (let i = 1; i < data.length; i++) {
    const rawKo = data[i][koCol];
    if (!rawKo) continue;

    let koStr;
    if (rawKo instanceof Date) {
      koStr = Utilities.formatDate(rawKo, tz, 'yyyy-MM-dd HH:mm:ss');
    } else {
      koStr = rawKo.toString().trim();
    }

    if (R32_MAP[koStr]) {
      const realHome = R32_MAP[koStr][0];
      const realAway = R32_MAP[koStr][1];
      const rowNum = i + 1;

      const oldHome = data[i][homeCol];
      const oldAway = data[i][awayCol];

      sheet.getRange(rowNum, homeCol + 1).setValue(realHome);
      sheet.getRange(rowNum, awayCol + 1).setValue(realAway);

      Logger.log('✅ ' + koStr + ': "' + oldHome + ' vs ' + oldAway + '" → "' + realHome + ' vs ' + realAway + '"');
      updated++;
    }
  }

  const sheetKickoffs = new Set();
  for (let i = 1; i < data.length; i++) {
    const rawKo = data[i][koCol];
    if (!rawKo) continue;
    const koStr = (rawKo instanceof Date)
      ? Utilities.formatDate(rawKo, tz, 'yyyy-MM-dd HH:mm:ss')
      : rawKo.toString().trim();
    sheetKickoffs.add(koStr);
  }
  Object.keys(R32_MAP).forEach(dt => {
    if (!sheetKickoffs.has(dt)) notFound.push(dt + ' (' + R32_MAP[dt].join(' vs ') + ')');
  });

  Logger.log('━━━━━━━━━━━━━━━━━━━━━━━━');
  Logger.log('Aggiornate: ' + updated + '/16 partite');
  if (notFound.length) {
    Logger.log('⚠️ Date NON trovate nello Sheet:');
    notFound.forEach(d => Logger.log('   ' + d));
  } else {
    Logger.log('✅ Tutte le 16 date abbinate correttamente');
  }

  return { success: true, updated: updated, notFound: notFound };
}
