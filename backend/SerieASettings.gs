// ============================================================
// Serie A application settings
// Creates an isolated settings tab without touching World Cup settings.
// ============================================================

const SERIE_A_SETTINGS_SHEET = 'Settings_SerieA';

const SERIE_A_DEFAULT_SETTINGS = Object.freeze([
  ['competition_code', 'SA', 'football-data.org competition code'],
  ['season_id', '2026-27', 'Application season identifier'],
  ['matches_sheet', 'Matches_SerieA', 'Active development fixture sheet'],
  ['timezone', 'Europe/Rome', 'Competition time zone'],
  ['lock_minutes_before_kickoff', 0, 'Editable until the exact kickoff'],
  ['show_predictions_after_lock', true, 'Reveal predictions at kickoff'],
  ['scoring_exact_score', 5, 'Exact score points'],
  ['scoring_correct_outcome', 1, 'Correct 1/X/2 points'],
  ['scoring_correct_home_goals', 1, 'Correct home goals points'],
  ['scoring_correct_away_goals', 1, 'Correct away goals points'],
  ['scoring_correct_goal_difference', 1, 'Correct signed goal difference points'],
  ['golden_match_enabled', true, 'One Golden Match per matchday'],
  ['golden_match_multiplier', 3, 'Multiplier applied to all match points'],
  ['golden_selection_lookahead_days', 7, 'Select before the matchday starts']
]);

/**
 * Creates Settings_SerieA and adds missing settings. Existing values are never
 * overwritten, so later administrator changes survive reruns.
 */
function setupSerieASettingsSheet() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = spreadsheet.getSheetByName(SERIE_A_SETTINGS_SHEET);
  let created = false;

  if (!sheet) {
    sheet = spreadsheet.insertSheet(SERIE_A_SETTINGS_SHEET);
    created = true;
  }

  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, 3).setValues([['setting', 'value', 'description']]);
    sheet.setFrozenRows(1);
  }

  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const settingColumn = headers.indexOf('setting');
  const valueColumn = headers.indexOf('value');
  const descriptionColumn = headers.indexOf('description');
  if (settingColumn === -1 || valueColumn === -1 || descriptionColumn === -1) {
    throw new Error('Settings_SerieA must contain setting, value and description columns');
  }

  const existing = new Set(
    data.slice(1)
      .map(row => row[settingColumn])
      .filter(Boolean)
      .map(value => value.toString())
  );
  const missing = SERIE_A_DEFAULT_SETTINGS.filter(row => !existing.has(row[0]));

  if (missing.length) {
    sheet.getRange(sheet.getLastRow() + 1, 1, missing.length, 3).setValues(missing);
  }

  sheet.autoResizeColumns(1, 3);
  return {
    success: true,
    created,
    settings_added: missing.map(row => row[0]),
    total_settings: SERIE_A_DEFAULT_SETTINGS.length
  };
}

function getSerieASettings_() {
  const settings = {};
  SERIE_A_DEFAULT_SETTINGS.forEach(row => { settings[row[0]] = row[1]; });

  const sheet = SpreadsheetApp.getActiveSpreadsheet()
    .getSheetByName(SERIE_A_SETTINGS_SHEET);
  if (!sheet || sheet.getLastRow() < 2) return settings;

  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const settingColumn = headers.indexOf('setting');
  const valueColumn = headers.indexOf('value');
  if (settingColumn === -1 || valueColumn === -1) return settings;

  data.slice(1).forEach(row => {
    if (row[settingColumn] !== '' && row[settingColumn] !== null) {
      settings[row[settingColumn].toString()] = row[valueColumn];
    }
  });
  return settings;
}

/**
 * Serie A score-only calculation. Golden Match multiplication is applied to
 * the entire score after normal points have been calculated.
 */
function calculateSerieAPoints_(prediction, match, suppliedSettings) {
  const status = (match.status || '').toString().trim().toLowerCase();
  if (status === 'void' || status === 'cancelled') return 0;

  const settings = suppliedSettings || getSerieASettings_();
  const rawScores = [
    prediction.pred_home_score,
    prediction.pred_away_score,
    match.home_score,
    match.away_score
  ];
  if (rawScores.some(score => !serieAValidScoreValue_(score))) return 0;

  const predictedHome = Number(prediction.pred_home_score);
  const predictedAway = Number(prediction.pred_away_score);
  const actualHome = Number(match.home_score);
  const actualAway = Number(match.away_score);

  let points = 0;
  if (predictedHome === actualHome && predictedAway === actualAway) {
    points = numberSetting_(settings.scoring_exact_score, 5);
  } else {
    if (serieAOutcome_(predictedHome, predictedAway) ===
        serieAOutcome_(actualHome, actualAway)) {
      points += numberSetting_(settings.scoring_correct_outcome, 1);
    }
    if (predictedHome === actualHome) {
      points += numberSetting_(settings.scoring_correct_home_goals, 1);
    }
    if (predictedAway === actualAway) {
      points += numberSetting_(settings.scoring_correct_away_goals, 1);
    }
    if ((predictedHome - predictedAway) === (actualHome - actualAway)) {
      points += numberSetting_(settings.scoring_correct_goal_difference, 1);
    }
  }

  const isGolden = truthy_(match.is_golden);
  const multiplier = isGolden
    ? numberSetting_(match.points_multiplier,
        numberSetting_(settings.golden_match_multiplier, 3))
    : 1;
  return points * multiplier;
}

function serieAOutcome_(home, away) {
  if (home > away) return '1';
  if (home < away) return '2';
  return 'X';
}

function numberSetting_(value, fallback) {
  const number = Number(value);
  return Number.isFinite(number) ? number : fallback;
}

function serieAValidScoreValue_(value) {
  if (value === '' || value === null || value === undefined) return false;
  const number = Number(value);
  return Number.isInteger(number) && number >= 0 && number <= 99;
}
