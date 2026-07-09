# TrackTalents Outlook Add-in MVP Architecture

## Purpose

This document defines the recommended architecture for the new TrackTalents Outlook add-in project. It is based on the current `tracktalents-v2` application and is designed to:

- fit Outlook add-in platform constraints cleanly;
- maximize reuse of existing TrackTalents APIs and business rules;
- avoid copying the full `tracktalents-v2` UI into Outlook;
- deliver a focused MVP that is practical to build and support.

## Executive Summary

The recommended solution is a modern Outlook web add-in with a task pane as the primary user experience. The add-in should be a thin Outlook-specific shell that:

- reads message context from Outlook using Office.js;
- authenticates the user into TrackTalents;
- transforms Outlook message data into TrackTalents payloads;
- calls TrackTalents APIs to create or link ATS records;
- keeps the UI intentionally compact and action-oriented.

We should reuse the existing TrackTalents API contracts and business rules where possible, but not attempt to port the full `tracktalents-v2` UI and dynamic form system into Outlook.

## Recommended Product Scope

### MVP actions

The first release should support these workflows:

1. Add Candidate from email
2. Add Contact from email
3. Add Job from email
4. Submit Resume to Contact
5. Source Candidate to Job
6. Link Email to existing ATS record

### Not in MVP

These items should be deferred until the core workflows are stable:

- Smart Alerts or send-time validation
- Full `tracktalents-v2` page parity
- Full dynamic form rendering inside Outlook
- Complex reporting or dashboards
- Advanced background automation
- Legacy file-download-based plugin behavior

## Why This Architecture

### Why task pane first

A task pane add-in gives the best balance of speed, UX, and maintainability.

Benefits:

- natural place for forms, search, and confirmation states;
- works well for mail-context-driven actions;
- easier to debug and evolve than event-only approaches;
- better fit for candidate/contact/job creation flows.

### Why not continue the old Office 365 prototype

The existing `tracktalents-v2/office365` prototype is useful as reference, but not as the foundation for the new project.

Issues with the old approach:

- it relies on legacy plain JavaScript files;
- it downloads payload files and attachments instead of completing ATS actions in-app;
- it mixes Outlook logic, auth logic, and export behavior in a single layer;
- it is not aligned with the cleaner service-driven architecture already present in `tracktalents-v2`.

### Why not port the full TrackTalents frontend

The main application contains large, stateful, desktop-oriented workflows. Outlook needs focused, fast, narrow interactions. Reusing the full UI would add complexity without adding much value.

## Recommended Technical Architecture

## High-level shape

```text
Outlook Client
  -> Office.js + Task Pane UI
  -> Outlook Context Adapter
  -> TrackTalents Session/Auth Layer
  -> TrackTalents API Adapter
  -> Existing TrackTalents Backend APIs
```

## Layers

### 1. Outlook shell

This is the new project in this repository.

Responsibilities:

- add-in manifest;
- task pane UI;
- Outlook item readers;
- command surface integration;
- action orchestration;
- small local state and notifications.

### 2. Outlook context adapter

This layer converts Office.js item data into normalized objects the app can use.

Example outputs:

- current message metadata;
- sender and recipients;
- subject and HTML body;
- attachments and attachment descriptors;
- compose vs read mode context;
- mailbox/account identifiers if available.

Suggested model:

```ts
type OutlookMessageContext = {
  itemId: string | null;
  mode: "read" | "compose";
  subject: string;
  bodyHtml: string;
  from?: {
    displayName?: string;
    email?: string;
  };
  to: Array<{ displayName?: string; email?: string }>;
  cc: Array<{ displayName?: string; email?: string }>;
  attachments: Array<{
    id: string;
    name: string;
    size?: number;
    contentType?: string;
    isInline?: boolean;
  }>;
};
```

### 3. TrackTalents domain layer

This layer holds reusable business transformations and validation rules.

Responsibilities:

- map Outlook context to candidate/contact/job payloads;
- perform duplicate checks before create/submit;
- select the primary attachment to parse as resume;
- normalize ATS-specific payload fields;
- keep business logic independent from the UI.

### 4. TrackTalents API adapter

This layer talks to TrackTalents backend endpoints and hides low-level request details from the task pane UI.

Responsibilities:

- candidate APIs;
- contact APIs;
- job APIs;
- pipeline APIs;
- file upload and resume parsing APIs;
- lookup/config APIs;
- email/linking APIs;
- permission checks.

### 5. Optional backend-for-frontend

For MVP, this is optional but strongly recommended if auth/session handling becomes messy.

A small BFF can:

- exchange Microsoft identity for a TrackTalents session;
- proxy TrackTalents APIs;
- simplify token refresh;
- avoid pushing too much security logic into the add-in;
- normalize inconsistent legacy payloads.

## Manifest Recommendation

Start with the Outlook add-in-only XML manifest for the MVP.

Reasoning:

- it is still the safest path for broad Outlook compatibility;
- it avoids early complexity around multi-client packaging choices;
- it keeps the first milestone focused on working Outlook behavior.

We can revisit the unified Microsoft 365 manifest later if product goals shift toward broader Microsoft 365 packaging strategy.

## Authentication Recommendation

## Current reality in `tracktalents-v2`

The main app currently depends on browser storage values such as:

- `token`
- `user-security`
- `userId`

This works inside the current web app, but should not be assumed to transfer cleanly to Outlook add-in runtimes.

## Recommended approach

### MVP auth

Use a dedicated TrackTalents login flow inside the task pane.

Benefits:

- fastest to implement;
- least ambiguity;
- matches existing backend session expectations;
- easiest to debug while the product surface is still evolving.

### Preferred long-term auth

Move to Microsoft identity plus TrackTalents session exchange.

Flow:

1. user opens add-in;
2. add-in gets Microsoft identity through MSAL/nested app auth;
3. backend validates identity and issues TrackTalents session/token;
4. add-in calls TrackTalents APIs with the issued session.

This is the cleaner long-term model, but it should not block the MVP if the backend side is not ready.

## Feature Reuse Map from `tracktalents-v2`

## High reuse

These areas should be reused heavily through service and payload logic:

- candidate create API
- contact create API
- job create API
- resume parse API
- attachment upload API
- pipeline submission API
- duplicate checks
- lookups/config endpoints
- plugin version/download concepts if still needed

## Medium reuse

These areas can be reused conceptually but likely need adaptation:

- search flows for candidates/jobs/contacts
- permission-based action visibility
- email-related workflows
- pipeline-related actions

## Low reuse

These should not be copied directly into Outlook:

- full DynamicForm UI
- large page-level layouts
- heavy modal chains
- broad navigation model
- desktop-scale multi-tab workflows

## Proposed Repo Structure

```text
Tracktalents-Extension/
  docs/
    outlook-addin-mvp-architecture.md
  apps/
    outlook-addin/
      manifest/
      src/
        app/
        components/
        features/
        services/
        adapters/
        domain/
        types/
  packages/
    tt-domain/
    tt-api/
    tt-outlook-adapters/
```

## Suggested module structure inside the add-in

```text
src/
  app/
    bootstrap/
    routes/
  components/
    common/
    forms/
  features/
    add-candidate/
    add-contact/
    add-job/
    source-to-job/
    submit-to-contact/
    link-email/
  services/
    auth/
    tracktalents/
    outlook/
  adapters/
    office-js/
  domain/
    mappers/
    validators/
    workflows/
  types/
```

## UX Recommendation

The add-in should feel like an action console, not a mini clone of the full ATS.

### Suggested layout

- top summary card for current email context;
- action list for primary workflows;
- compact form or stepper for the selected action;
- confirmation state with created record details and deep link back to TrackTalents.

### UX principles

- avoid full-width enterprise form sprawl;
- prefill aggressively from email context;
- keep decisions to the minimum required fields;
- use search-and-select instead of rendering giant config-driven forms;
- always show what data will be sent to TrackTalents.

## MVP Workflow Design

## Add Candidate

1. Read email context
2. Identify resume-like attachment
3. Upload and parse resume
4. Prefill candidate fields
5. Run duplicate email check
6. Confirm or edit minimal fields
7. Create candidate
8. Offer optional next step: source or submit

## Add Contact

1. Read sender and recipient data
2. Prefill contact fields
3. Run duplicate check
4. Choose company if needed
5. Create contact

## Add Job

1. Use email subject as draft job title
2. Use email body as draft job description
3. Select contact/company
4. Create job

## Submit Resume to Contact

1. Identify candidate
2. Identify contact
3. Pick resume/document
4. Check duplicate submission
5. Submit to pipeline

## Source Candidate to Job

1. Identify candidate
2. Search and select job
3. Check duplicate submission
4. Create source/submission record

## Link Email to ATS Record

1. Search record
2. Attach message metadata/body
3. Save activity or communication linkage

## Suggested Delivery Phases

## Phase 1: foundation

- scaffold add-in project
- choose stack and build tooling
- create manifest
- implement Outlook context reader
- implement TrackTalents auth
- build shell UI

## Phase 2: core workflows

- add contact
- add candidate
- add job
- attachment handling
- resume parse integration

## Phase 3: pipeline actions

- submit resume to contact
- source candidate to job
- link email to record

## Phase 4: refinements

- permissions and role checks
- deep links back to ATS
- analytics/logging
- better error recovery
- optional event-based features

## Risks and Mitigations

## Risk: session mismatch between Outlook and TrackTalents

Mitigation:

- use dedicated add-in auth for MVP;
- isolate auth behind a service abstraction;
- add BFF if direct token handling becomes brittle.

## Risk: trying to reuse too much frontend

Mitigation:

- reuse services and payload mappers;
- rebuild Outlook UX specifically for task pane constraints.

## Risk: attachment APIs behave differently across Outlook contexts

Mitigation:

- build a single attachment adapter abstraction;
- test read mode and compose mode separately;
- define fallback behavior when attachment access is unavailable.

## Risk: platform feature fragmentation

Mitigation:

- keep MVP focused on supported mail scenarios;
- avoid event-based automation until task pane flows are stable;
- document minimum supported Outlook clients during development.

## Decisions

## Confirmed recommendations

- Build a new Outlook web add-in in this repository.
- Use a task pane as the primary UX.
- Reuse TrackTalents backend APIs and business rules.
- Do not reuse the old `office365` prototype as the core architecture.
- Do not attempt full `tracktalents-v2` UI parity.
- Start with a focused MVP.
- Use dedicated TrackTalents login for MVP unless backend is ready for Microsoft identity exchange.
- Start with add-in-only XML manifest.

## Open decisions

The following items still need product or engineering confirmation:

1. Should MVP auth be TrackTalents-only or Microsoft identity plus TrackTalents exchange?
2. Should the initial release target Outlook Web and new Outlook first, or include classic Outlook from day one?
3. What should "Link Email to ATS Record" save exactly: an activity note, a sent-email record, a document, or a communication object?
4. Should "Reply All" exist in MVP, and if yes, should it create an Outlook draft, log communication, or both?

## Recommended Immediate Next Step

Build the project foundation in this repository with:

1. add-in scaffold
2. manifest
3. task pane shell
4. auth screen
5. Outlook message context adapter
6. one end-to-end workflow: Add Contact

That gives the fastest real proof that the architecture works before we add candidate parsing and pipeline actions.
