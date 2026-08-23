# BTCwebapp production Firebase backend

This directory contains the production Firestore cache publisher, asynchronous
measurement submission pipeline, and notification backend for project
`btcwebapp-551bd` in `asia-northeast3`.

Form cache publishing is event-driven. Controlled GAS uploads and measurement
writes enqueue `formPublishJobs`; a Firestore create trigger publishes only the
affected form. The scheduled publisher runs once per day at 03:15 Asia/Seoul as
a recovery reconciliation, not as the primary update path.

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
- Set the same `BWA_PUBLISHER_TOKEN` value in the Apps Script project's Script
  Properties. GAS uses it only to enqueue a form publish job through the
  protected `enqueueFormPublish` endpoint.
- Create `HUMETRO_ID` and `HUMETRO_PW` in Secret Manager. Scheduled collection
  logs in to Humetro directly; credentials never enter Firestore or browser code.
- Import active legacy rows from the `FCM_Tokens` sheet with the protected admin
  migration endpoint. The sheet remains an untouched rollback reference.
- Deploy Functions, Firestore Rules, and indexes while both gates remain OFF.

## Verification

From `firebase-production/functions`:

```powershell
npm.cmd run lint
npm.cmd test
```

## Deployment

Always pass the production project explicitly. Deployment and gate changes are
separate operations so code can be verified while writes and scheduled dispatch
remain blocked.

Roll out the compatible browser reader (`production-read.js` and `index.html`)
first. Then deploy Firestore Rules and Functions, and finally deploy the updated
Apps Script after setting its `BWA_PUBLISHER_TOKEN` Script Property. This order
keeps existing legacy chunks readable until clients understand revision-scoped
chunks.

```powershell
firebase deploy --only "firestore:rules,firestore:indexes,functions" `
  --project btcwebapp-551bd --config firebase.json
```

Opening the submission gate for a real form and enabling scheduled notification
dispatch are separate operational changes and require explicit production approval.
