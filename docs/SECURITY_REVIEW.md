# Initial backend security review

This review describes the imported backend and is not a production penetration
test.

## Critical before a new deployment

1. The imported source contained a football-data.org API token. It was removed
   from the repository copy. Revoke/regenerate it and store the replacement in
   Apps Script Script Properties.
2. `submitPrediction` trusts a public `player_id`. After login, the browser holds
   no session token, so a caller can potentially submit as another player by
   changing the ID.
3. Administrative actions such as `recalculateAll` are exposed through the same
   public web endpoint without an administrator check.

## High priority

1. Login sends the PIN as a URL query parameter from the current frontend. URLs
   can appear in browser history and intermediary logs. Move authentication and
   writes to POST requests.
2. PINs are compared with plain values stored in the Players sheet. Introduce a
   migration to salted hashes or replace PIN authentication with short-lived
   server-issued sessions.
3. Concurrent submissions can race: the code checks for an existing row and
   then updates/appends without `LockService`. Two simultaneous saves could
   create duplicates.

## Correct controls already present

- Prediction locking is rechecked on the server from the kickoff timestamp.
- Scores are validated as non-negative integers.
- Inactive players and finished/void matches are rejected.
- Other players' predictions are withheld until the match is locked.

## Recommended order

1. Rotate the exposed provider token.
2. Copy the sheet and backend into a development environment.
3. Restrict administrative actions.
4. Add server-issued player sessions and POST writes.
5. Add `LockService` around prediction create/update operations.
6. Only then begin the Serie A data migration.

