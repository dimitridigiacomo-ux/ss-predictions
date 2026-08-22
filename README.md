# S&S Predictions

Football prediction game published with GitHub Pages and backed by Google Apps
Script and Google Sheets.

## Live safety

- `main` is the production branch used by the existing live site.
- Serie A work belongs on `serie-a-migration` until it has been tested.
- The root `index.html` remains the GitHub Pages entry point.
- Do not replace the current Apps Script deployment URL during development.

## Current architecture

```text
GitHub Pages
  index.html (HTML, CSS and JavaScript in one file)
       |
       | HTTPS requests
       v
Google Apps Script web app (Code.gs; not yet exported here)
       |
       v
Google Sheets (players, matches, predictions and results)
```

The browser currently calls four API actions:

- `login`
- `getMatches`
- `getLeaderboard`
- `submitPrediction`

See [docs/MIGRATION.md](docs/MIGRATION.md) for the safe migration and Serie A
conversion plan.

## Files still needed

Before changing application behaviour, export these items into the repository:

1. The complete Google Apps Script source (`Code.gs` and any other `.gs` or
   `.html` files in that Apps Script project).
2. A copy of the Google Sheets column headers, with private player data removed.
3. The Apps Script project time zone and deployment settings.

Never commit real PINs, private player information, or Google credentials.

