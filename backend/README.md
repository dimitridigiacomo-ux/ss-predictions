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

