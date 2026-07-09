# TrackTalents Outlook Add-in

This repo now contains the first real TrackTalents Outlook add-in app foundation. The current build starts with:

- Outlook task pane hosting
- TrackTalents login screen
- local auth proxy for development
- authenticated app shell
- basic current-email context summary

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

- `https://localhost:3001/taskpane.html`

The Codex in-app browser may not trust local developer certificates. For previewing the same page there, use:

- `http://localhost:3002/taskpane.html`

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
3. The task pane should open the TrackTalents login screen.
4. After login, the app shell should show the current message subject and basic metadata.

## What to build next

The next logical steps are:

1. add better session bootstrapping and logout handling;
2. add mail-context adapter utilities;
3. add `Add Contact`;
4. add `Add Candidate`.
