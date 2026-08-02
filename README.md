# TrackTalents Outlook Add-in

This repo now contains the current TrackTalents Outlook add-in foundation. The current build focuses on:

- Outlook task pane hosting
- TrackTalents-themed launcher UI
- login gating when an action is clicked
- email and resume context capture from the open Outlook message
- app handoff into TrackTalents pages in a new browser tab

## What is included

- `manifest/tracktalents-outlook-manifest.xml`
- `public/taskpane.html`
- `public/taskpane.js`
- `public/taskpane.css`
- `scripts/dev-server.js`

## Local setup

1. Install dependencies.

```bash
npm install
```

2. Install or refresh the localhost developer certificate.

```bash
npm run certs:install
```

3. Start the local HTTPS server.

```bash
npm run dev
```

The add-in files will be hosted at:

- `https://localhost:3201/taskpane.html`

The Codex in-app browser may not trust local developer certificates. For previewing the same page there, use:

- `http://localhost:3202/taskpane.html`

These ports intentionally avoid `3001/3002` so the extension can run at the same time as `tracktalents-v2`.

## Sideload into Outlook

Use the manifest file:

- `manifest/tracktalents-outlook-manifest.xml`

For Outlook on the web or new Outlook, go to:

- `Apps`
- `Add apps`
- `My add-ins`
- `Add a custom add-in`
- `Add from file`

Then select the manifest file.

After sideloading:

1. Open any email message.
2. Open the add-in from the message action bar or ribbon.
3. The task pane should show the action launcher.
4. Open an email with a resume attachment to see real message context.
5. Click any action such as `Add Candidate`.
6. If not logged in, the add-in will ask for TrackTalents login.
7. After login, the corresponding TrackTalents page opens in a new tab and lands on the matching form flow with Outlook context in the query string.

## What to build next

The next logical steps are:

1. replace demo query-string handoff with the final resume-prefill API flow;
2. extend the same auto-open pattern to every remaining Outlook action;
3. attach real parsed resume data to `Add Candidate`;
4. add shared auth or SSO so the web app and add-in reuse one session.
