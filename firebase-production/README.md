# BTCwebapp production Firebase backend

This directory contains the production Firestore cache publisher, asynchronous
measurement submission pipeline, and notification backend for project
`btcwebapp-551bd` in `asia-northeast3`.

## Safety boundaries

- Every Function refuses to use a project other than `btcwebapp-551bd`.
- Publisher and Sheets operations accept only spreadsheet
  `19rgzRnTQtOwwW7Ts5NbBuItNey94dAZsEnO7Tk0cm6s`.
- Public HTTPS endpoints accept only origin `https://somyun.github.io`.
- Browser access to submissions, rate limits, notification devices, receipts,
  and `systemConfig` is denied by Firestore Rules.
- `systemConfig/submissions.enabled` is fail-closed. A missing document or any
  value other than `true` rejects new submissions.
- `systemConfig/notificationDispatch.enabled` is fail-closed and affects only
  scheduled Happy Hugether dispatch. Self-test and heartbeat remain separate.

The submission gate can additionally contain `allowedFormKeys` and
`allowedSheetNames`. Non-empty arrays restrict writes to the listed forms. The
notification dispatch gate must remain disabled until direct Humetro collection,
legacy device migration, and the notification baseline have been verified, and the
legacy GAS notification trigger is disabled.

## Required cloud setup

- Enable Google Sheets API.
- Grant the Functions service account edit access to the production spreadsheet.
- Create `BWA_PUBLISHER_TOKEN` for administrator endpoints.
- Create `HUMETRO_ID` and `HUMETRO_PW` in Secret Manager. Scheduled collection
  logs in to Humetro directly; credentials never enter Firestore or browser code.
- Import active legacy rows from the `FCM_Tokens` sheet with the protected admin
  migration endpoint. The sheet remains an untouched rollback reference.
- Deploy Functions, Firestore Rules, and indexes while both gates remain OFF.

## Verification

From the repository root:

```powershell
node --test tests\*.test.js
```

From `firebase-production/functions`:

```powershell
npm.cmd run verify
```

## Deployment

Always pass the production project explicitly. Deployment and gate changes are
separate operations so code can be verified while writes and scheduled dispatch
remain blocked.

```powershell
firebase deploy --only "firestore:rules,firestore:indexes,functions" `
  --project btcwebapp-551bd --config firebase.json
```

Opening the submission gate for a real form and enabling scheduled notification
dispatch are operational changes that require the approvals documented in
`docs/PRODUCTION_TEST_CODE_ADOPTION_PLAN.md`.
