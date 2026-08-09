# BTCwebapp production Firestore cache

This directory owns the production Firestore read cache and asynchronous measurement-save
pipeline for project `btcwebapp-551bd`. Upload and XLSX export remain in Apps Script;
measurement values are no longer posted directly from the browser to Apps Script.

## Resources

- Firestore Standard `(default)` database in `asia-northeast3`
- Public read-only collections: `publicCache`, `publicForms`
- Scheduled function: `publishAllChangedFormsScheduled`
- Public submission functions: `submitMeasurements`, `getMeasurementSubmission`
- Firestore worker: `syncMeasurementSubmission`
- Schedule: every 5 minutes in `Asia/Seoul`
- Runtime guard: the function refuses to run outside `btcwebapp-551bd`
- Source guard: GAS responses must reference spreadsheet
  `19rgzRnTQtOwwW7Ts5NbBuItNey94dAZsEnO7Tk0cm6s`

The scheduler always checks the form list. It fetches full form data only when that
form's source revision changed, which keeps Apps Script traffic bounded.

## Async measurement save

1. The browser submits a validated, idempotent request to `submitMeasurements`.
2. Firestore stores it as `queued`; clients cannot read or write the server-only queue.
3. `syncMeasurementSubmission` writes every measurement to column F and the accepted
   revision to `FormList` with one Google Sheets API `values.batchUpdate` request.
4. After Sheets succeeds, the worker advances the public form and form-list cache to
   that same revision and marks the submission `synced`.
5. The browser prepares XLSX only after observing `synced`. It never falls back to a
   direct Apps Script measurement write.

The implementation therefore includes the final Google Sheet write. It is not usable
in a deployed environment until all of these operations are completed:

- enable the Google Sheets API in `btcwebapp-551bd`;
- share spreadsheet `19rgzRnTQtOwwW7Ts5NbBuItNey94dAZsEnO7Tk0cm6s` with the
  Functions runtime service account as an editor;
- deploy the Functions and Firestore rules below;
- set secret `BWA_PRODUCTION_SUBMISSION_ADMIN_TOKEN`;
- explicitly open `systemConfig/submissions.enabled` through `setSubmissionGate` only
  after a smoke test. A missing gate document is fail-closed.

## Verification

From the repository root:

```powershell
& 'C:\Program Files\nodejs\node.exe' scripts\compare-production-published.js
```

From `firebase-production/functions`:

```powershell
$env:Path = 'C:\Program Files\nodejs;' + $env:Path
& 'C:\Program Files\nodejs\npm.cmd' run verify
```

## Deployment

Always pass the project ID explicitly:

```powershell
$env:Path = 'C:\Program Files\nodejs;' + $env:Path
& "$env:APPDATA\npm\firebase.cmd" deploy --only 'firestore:rules,firestore:indexes' `
  --project btcwebapp-551bd --config firebase.json
& "$env:APPDATA\npm\firebase.cmd" deploy `
  --only 'functions:publishAllChangedFormsScheduled,functions:submitMeasurements,functions:getMeasurementSubmission,functions:syncMeasurementSubmission,functions:retryMeasurementSubmission,functions:setSubmissionGate' `
  --project btcwebapp-551bd --config firebase.json
```

Container images older than seven days are removed by the Artifact Registry cleanup policy.
Measurement submission remains disabled until the gate is explicitly opened.
