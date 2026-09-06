# TrackTalents Outlook Add-in

This repo now contains the current TrackTalents Outlook add-in foundation. The current build focuses on:

- Outlook task pane hosting
- TrackTalents-themed launcher UI
- login gating when an action is clicked
- email and resume context capture from the open Outlook message
- AI email parsing for `Add Contact` and `Add Job`
- app handoff into TrackTalents pages in a new browser tab

## What is included

- `manifest/tracktalents-outlook-development.xml`
- `manifest/tracktalents-outlook-production.xml`
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
- `https://localhost:3201/manifest.xml`

The Codex in-app browser may not trust local developer certificates. For previewing the same page there, use:

- `http://localhost:3202/taskpane.html`

These ports intentionally avoid `3001/3002` so the extension can run at the same time as `tracktalents-v2`.

## Parser API

`Add Contact` and `Add Job` call the email parser through the local add-in bridge:

```text
POST /api/email/parse
```

By default, the bridge forwards parser requests to the deployed Railway API:

```text
https://tracktalents-ai-production.up.railway.app
```

Override it when needed:

```bash
EMAIL_PARSER_API_URL=https://tracktalents-ai-production.up.railway.app npm run dev
```

## Railway Deployment

The deployed add-in exposes a manifest at:

```text
https://tracktalents-outlook-extension-production.up.railway.app/manifest.xml
```

The manifest is generated from `manifest/tracktalents-outlook-manifest.xml` and rewrites local URLs to the public Railway host. If the Railway domain changes, set `ADDIN_PUBLIC_URL` to the new public base URL.

Set `APP_HOST` in Railway to the public TrackTalents ATS web app URL when the add-in should open live ATS pages instead of local development pages.

## Outlook manifests

- `manifest/tracktalents-outlook-development.xml` is named **TrackTalents Outlook Dev** and targets `https://localhost:3201`. Use it only while the local add-in server is running for development/testing.
- `manifest/tracktalents-outlook-production.xml` is named **TrackTalents Outlook** and targets the deployed Railway add-in host. Use it for production sideloading/deployment.

The manifests use distinct add-in IDs, so Outlook can install the development and production add-ins side by side. The `/manifest.xml` endpoint serves the development manifest locally and the production manifest when hosted on Railway.

## Sideload into Outlook

Use the appropriate manifest file:

- development/testing: `manifest/tracktalents-outlook-development.xml`
- production: `manifest/tracktalents-outlook-production.xml`

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

1. map additional AI parser fields into the TrackTalents dynamic form configs;
2. extend the same auto-open pattern to every remaining Outlook action;
3. attach real parsed resume data to `Add Candidate`;
4. add shared auth or SSO so the web app and add-in reuse one session.
