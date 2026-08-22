# Migration plan: World Cup app to Serie A

## Objective

Convert the existing single-file prediction game to a season-based Serie A app
without interrupting the current GitHub Pages deployment or corrupting existing
World Cup data.

## Baseline recorded on 22 August 2026

- Production source: root `index.html` on `main`.
- Frontend: dependency-free HTML, CSS and JavaScript in one file.
- Backend: deployed Google Apps Script web app.
- Storage: Google Sheets.
- `Code.gs` has now been exported into `backend/Code.gs`.
- The expected sheet schema has been reconstructed in `docs/DATA_MODEL.md` and
  still needs comparison with the real spreadsheet headers.
- The frontend uses `login`, `getMatches`, `getLeaderboard` and
  `submitPrediction`.
- World Cup assumptions are embedded in rendering and scoring help, including
  `stage`, `round`, knockout qualifiers, qualification method and first scorer.

## Non-negotiable safety rules

1. Keep `main` unchanged until the Serie A version passes acceptance tests.
2. Never test migrations against the production spreadsheet.
3. Make a full copy of the spreadsheet and Apps Script project for development.
4. Give the development Apps Script a new deployment URL.
5. Store URLs in environment configuration; do not silently repoint production.
6. Preserve the current API contract during the first frontend refactor.
7. Do not commit PINs, sheet IDs, access tokens or private player data.

## Target workspace

The root `index.html` stays in place so GitHub Pages keeps working. Code is
separated incrementally rather than moved in one large change.

```text
index.html                 GitHub Pages entry point
src/
  config.js                app/season/API configuration
  api.js                   API client, timeout and retry behaviour
  state.js                 browser state
  ui/                      page renderers and interactions
backend/
  Code.gs                  exported Apps Script source
  appsscript.json          Apps Script manifest
docs/
  MIGRATION.md             this plan
  DATA_MODEL.md            sheet/API schema after export
tests/
  fixtures/                anonymised API responses
  smoke/                   critical user-flow checks
```

No build tool is required for the first migration. This keeps GitHub Pages
simple and makes regressions easier to isolate.

## Delivery phases

### Phase 0 — freeze and inventory (started)

- Clone the production repository.
- Keep `main` as the untouched production baseline.
- Create `serie-a-migration` for all work.
- Record the current API actions and embedded tournament assumptions.
- Compare the exported backend and reconstructed schema with the live project.

Exit condition: frontend, Apps Script and sheet contracts are all documented.

### Phase 1 — make the existing app testable

- Add anonymised fixtures for every API action.
- Add a development configuration with a separate Apps Script URL.
- Add smoke tests for login, match loading, saving and leaderboard loading.
- Add optional loading diagnostics behind `?debug=1`.
- Confirm that production behaviour is unchanged.

Exit condition: the existing World Cup version can be tested locally and still
produces the same UI/API calls.

### Phase 2 — extract the single-file frontend

Extract one responsibility at a time from `index.html`: configuration, API
client, state, formatting, then renderers. After every extraction, run the smoke
tests and compare the live flows.

Exit condition: the application is modular, but functionally identical.

### Phase 3 — introduce a season model

Replace tournament-specific decisions with explicit configuration:

```js
const COMPETITION = {
  id: 'serie-a-2026-27',
  name: 'Serie A',
  season: '2026/27',
  format: 'league',
  matchdays: 38,
  timezone: 'Europe/Rome'
};
```

The data model should support `competition_id` and `season_id` so old World Cup
records can remain separate and future seasons do not require another rewrite.

Exit condition: switching competition data does not require changing UI code.

### Phase 4 — migrate the backend and spreadsheet copy

- Add competition/season identifiers to matches, predictions and leaderboard
  calculations.
- Import the Serie A calendar into the development sheet.
- Use stable match IDs; do not derive identity only from team names or dates.
- Set lock times from kickoff timestamps in `Europe/Rome` and store timestamps
  unambiguously.
- Decide how postponed matches and changed kickoff times update locks.
- Recalculate scores only in the copied sheet until verified.

Exit condition: the development backend returns a full Serie A season correctly.

### Phase 5 — Serie A UI and scoring decisions

Serie A has 38 league matchdays and no knockout rounds. Remove or hide qualifier
and qualification-method fields for league matches. Before implementation,
confirm whether first-goalscorer predictions stay and confirm the scoring rules.

Exit condition: all 380 fixtures render correctly, filters remain usable on
mobile, and scoring matches agreed examples.

### Phase 6 — staged release

- Deploy the Serie A branch to a preview URL or separate Pages repository.
- Test with one administrator and at least one ordinary test player.
- Back up production immediately before cutover.
- Merge only after acceptance; retain the old deployment for rollback.

## Acceptance checklist

- Login succeeds and invalid PINs fail without exposing sensitive information.
- Only the correct player's prediction is editable before lock.
- A prediction can be created and updated before lock.
- Editing is impossible at and after the lock timestamp.
- Other players' predictions reveal only at the agreed time.
- Leaderboard totals match hand-calculated examples.
- Postponed/rescheduled fixtures behave predictably.
- All 38 matchdays and 380 fixtures can load without freezing the page.
- Slow/cold Apps Script responses show useful feedback and can retry safely.
- Existing World Cup records remain readable and unchanged.
- Mobile layout works on representative iPhone and Android widths.

## First decisions needed after backend export

1. Which Serie A season is the initial target?
2. Keep first goalscorer for every league match, or use score-only predictions?
3. Keep the current points system or define a new Serie A system?
4. Should old World Cup accounts/PINs carry over?
5. When should all predictions become visible: kickoff or full time?
