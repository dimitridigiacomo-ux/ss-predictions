# Google Apps Script backend

`Code.gs` is the exported backend for the current application. The repository
copy has one deliberate security change: the football-data.org token has been
removed from source control.

## Development setup

1. Make a copy of the production Google Sheet.
2. Make a separate Apps Script project connected to that copied sheet.
3. Add a Script Property named `FOOTBALL_API_KEY` containing a newly generated
   football-data.org token.
4. Set the Apps Script project time zone to `Europe/Rome` for Serie A.
5. Deploy the development project as a separate web app.
6. Use its URL only from the migration branch during testing.

Do not reuse the token that appeared in the original pasted source. It must be
revoked or regenerated before result synchronisation is used again.

## Serie A fixture module

`SerieAMatches.gs` is isolated from the live World Cup data. It uses a separate
`Matches_SerieA` tab and supports:

- read-only API preview;
- initial season import;
- updates by stable provider match ID;
- postponed/rescheduled fixtures;
- manual kickoff overrides;
- a six-hour synchronisation trigger.

Run functions in this order in the development Apps Script project:

1. `previewSerieAFixtures()`
2. `setupSerieAMatchesSheet()`
3. `syncSerieAFixtures()`
4. `createSerieAFixtureSyncTrigger()` after the imported fixtures have been
   checked.
