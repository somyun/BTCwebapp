# BTCwebapp production Firestore cache

This directory owns only the production read cache for project `btcwebapp-551bd`.
It does not replace the existing Apps Script write, upload, notification, or XLSX paths.

## Resources

- Firestore Standard `(default)` database in `asia-northeast3`
- Public read-only collections: `publicCache`, `publicForms`
- Scheduled function: `publishAllChangedFormsScheduled`
- Schedule: every 5 minutes in `Asia/Seoul`
- Runtime guard: the function refuses to run outside `btcwebapp-551bd`
- Source guard: GAS responses must reference spreadsheet
  `19rgzRnTQtOwwW7Ts5NbBuItNey94dAZsEnO7Tk0cm6s`

The scheduler always checks the form list. It fetches full form data only when that
form's source revision changed, which keeps Apps Script traffic bounded.

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
  --only functions:publishAllChangedFormsScheduled `
  --project btcwebapp-551bd --config firebase.json
```

Container images older than seven days are removed by the Artifact Registry cleanup policy.
Measurement submission through Firestore remains intentionally disabled.
