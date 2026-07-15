const ACTIONS = [
  {
    id: "add-candidate",
    label: "Add Candidate",
    description: "Open the candidate flow in TrackTalents with resume context from this email.",
    icon: "+",
    iconClass: "icon-candidate",
    path: "/candidates",
    intent: "create-candidate"
  },
  {
    id: "submit-resume-contact",
    label: "Submit Resume To Contact",
    description: "Open the contact workflow and prepare resume submission from the selected email.",
    icon: "C",
    iconClass: "icon-contact",
    path: "/contacts",
    intent: "submit-resume-contact"
  },
  {
    id: "add-job",
    label: "Add Job",
    description: "Open a new job flow and reuse the current email details as source context.",
    icon: "J",
    iconClass: "icon-job",
    path: "/jobs",
    intent: "create-job"
  },
  {
    id: "add-contact",
    label: "Add Contact",
    description: "Open the contact form with sender details carried over from Outlook.",
    icon: "P",
    iconClass: "icon-person",
    path: "/contacts",
    intent: "create-contact"
  },
  {
    id: "source-resume-job",
    label: "Source Resume To Job",
    description: "Open the sourcing workflow to match the resume from this email with jobs.",
    icon: "S",
    iconClass: "icon-source",
    path: "/local-search",
    intent: "source-resume-job"
  },
  {
    id: "reply-all",
    label: "Reply All",
    description: "Jump into the communication area with the Outlook message context attached.",
    icon: "R",
    iconClass: "icon-reply",
    path: "/sentmails",
    intent: "reply-all"
  }
];

const state = {
  officeReady: false,
  officeHost: null,
  officePlatform: null,
  currentItem: buildPreviewItem(),
  config: {
    appHost: "https://test.tracktalents.com",
    loginPath: "/login",
    forgotPasswordPath: "/forgotpassword"
  },
  showLoginPrompt: false,
  pendingActionId: null,
  contextBanner: "",
  launchMessage: ""
};

function getRoot() {
  return document.getElementById("app");
}

function escapeHtml(value) {
  return String(value ?? "")
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#039;");
}

function escapeAttribute(value) {
  return escapeHtml(value).replaceAll("`", "&#096;");
}

function buildPreviewItem() {
  return {
    itemId: "preview-mail-001",
    subject: "Senior Java Developer Resume - Ananya Sharma",
    from: {
      displayName: "Ananya Sharma",
      email: "ananya.sharma@example.com"
    },
    fromDisplay: "Ananya Sharma <ananya.sharma@example.com>",
    toCount: 1,
    attachmentCount: 2,
    attachmentNames: ["Ananya-Sharma-Resume.pdf", "portfolio.txt"],
    hasResumeAttachment: true,
    primaryResumeName: "Ananya-Sharma-Resume.pdf",
    bodyPreview:
      "Hello team, please find my latest resume attached for the senior Java developer role. I have 7 years of backend and cloud experience.",
    mode: "preview"
  };
}

function getInitials(value) {
  return String(value || "")
    .split(/\s+/)
    .filter(Boolean)
    .slice(0, 2)
    .map((part) => part.charAt(0).toUpperCase())
    .join("") || "TT";
}

function isResumeFile(name) {
  return /\.(pdf|doc|docx|rtf|txt)$/i.test(String(name || ""));
}

function formatContextId(item) {
  const source = item?.itemId || "preview-context";
  return String(source).replace(/[^a-zA-Z0-9]/g, "").slice(0, 18) || "previewcontext";
}

function getItemSummary() {
  if (!state.currentItem) {
    return {
      subject: "No email selected",
      fromDisplay: "Open an Outlook email with a resume attachment to use this add-in.",
      attachmentLabel: "No attachments",
      statusLabel: state.officeReady ? "Waiting for email" : "Preview mode",
      hasResume: false
    };
  }

  const names = state.currentItem.attachmentNames || [];
  const attachmentLabel = names.length ? names.join(", ") : "No attachments";

  return {
    subject: state.currentItem.subject || "No subject",
    fromDisplay: state.currentItem.fromDisplay || "Unknown sender",
    attachmentLabel,
    statusLabel: state.currentItem.hasResumeAttachment ? "Resume detected" : "No resume detected",
    hasResume: state.currentItem.hasResumeAttachment
  };
}

function actionLabelFromId(actionId) {
  return ACTIONS.find((action) => action.id === actionId)?.label || "TrackTalents action";
}

function render() {
  const root = getRoot();
  const itemSummary = getItemSummary();
  const loginChip = `<button id="open-login-button" class="ghost-button" type="button">Login</button>`;
  const hostPill = state.officeReady
    ? `Connected to Outlook${state.officePlatform ? ` on ${state.officePlatform}` : ""}`
    : "Preview mode";

  root.innerHTML = `
    <main class="shell">
      <section class="hero-card">
        <div class="hero-top">
          <div class="brand-lockup">
            <div class="brand-mark">TT</div>
            <div>
              <p class="eyebrow">TrackTalents Outlook</p>
              <h1>Import talent from your inbox</h1>
            </div>
          </div>
          ${loginChip}
        </div>

        <p class="hero-copy">
          Open an email with a resume, choose an action here, and we’ll hand you off into TrackTalents with the mail context ready.
        </p>

        <div class="pill-row">
          <span class="pill ${itemSummary.hasResume ? "pill-accent" : "pill-muted"}">${escapeHtml(itemSummary.statusLabel)}</span>
          <span class="pill pill-muted">${escapeHtml(hostPill)}</span>
        </div>
      </section>

      <section class="context-card">
        <div class="context-head">
          <div>
            <p class="section-label">Selected Email</p>
            <h2>${escapeHtml(itemSummary.subject)}</h2>
          </div>
          <div class="avatar">${escapeHtml(getInitials(state.currentItem?.from?.displayName || "TrackTalents"))}</div>
        </div>

        <div class="context-grid">
          <div class="context-row">
            <span>From</span>
            <strong>${escapeHtml(itemSummary.fromDisplay)}</strong>
          </div>
          <div class="context-row">
            <span>Context ID</span>
            <strong>${escapeHtml(formatContextId(state.currentItem))}</strong>
          </div>
          <div class="context-row context-row-wide">
            <span>Attachments</span>
            <strong>${escapeHtml(itemSummary.attachmentLabel)}</strong>
          </div>
        </div>

        <p class="context-note">
          ${escapeHtml(
            state.currentItem?.bodyPreview
              ? state.currentItem.bodyPreview
              : "When your API is ready, this same email context can power automatic form filling in the main TrackTalents app."
          )}
        </p>
      </section>

      ${
        state.contextBanner
          ? `<div class="banner banner-info">${escapeHtml(state.contextBanner)}</div>`
          : ""
      }
      ${
        state.launchMessage
          ? `<div class="banner banner-success">${escapeHtml(state.launchMessage)}</div>`
          : ""
      }

      <section class="actions-card">
        <div class="actions-head">
          <div>
            <p class="section-label">Actions</p>
            <h2>Choose what to do with this email</h2>
          </div>
        </div>

        <div class="action-list">
          ${ACTIONS.map(renderActionButton).join("")}
        </div>
      </section>

      ${
        state.showLoginPrompt
          ? renderLoginPrompt()
          : ""
      }
    </main>
  `;

  bindEvents();
}

function renderActionButton(action) {
  return `
    <button
      class="action-button"
      type="button"
      data-action-id="${escapeAttribute(action.id)}"
      aria-label="${escapeAttribute(action.label)}"
    >
      <span class="action-icon ${escapeAttribute(action.iconClass)}" aria-hidden="true">${escapeHtml(action.icon)}</span>
      <span class="action-copy">
        <strong>${escapeHtml(action.label)}</strong>
        <small>${escapeHtml(action.description)}</small>
      </span>
      <span class="action-arrow" aria-hidden="true">Open</span>
    </button>
  `;
}

function renderLoginPrompt() {
  const pendingAction = state.pendingActionId
    ? actionLabelFromId(state.pendingActionId)
    : "TrackTalents";
  const launchButtonLabel = state.pendingActionId
    ? `Login and Open ${pendingAction}`
    : "Open TrackTalents Login";

  return `
    <section class="login-overlay" aria-label="Login prompt">
      <div class="login-dialog">
        <div class="login-header">
          <div>
            <p class="section-label">Login Required</p>
            <h2>Continue to ${escapeHtml(pendingAction)}</h2>
          </div>
          <button id="close-login-button" class="icon-button" type="button" aria-label="Close login prompt">x</button>
        </div>

        <p class="login-copy">
          Sign in through the real TrackTalents website first. After login, you will land directly on the ${escapeHtml(pendingAction)} form with this Outlook email context attached.
        </p>

        <div class="login-note">
          The Outlook add-in will not do a separate login anymore. This avoids being asked twice.
        </div>

        <div class="login-actions">
          <button id="continue-login-button" class="primary-button" type="button">${escapeHtml(launchButtonLabel)}</button>
          <button id="forgot-password" class="text-button" type="button">Forgot password</button>
        </div>
      </div>
    </section>
  `;
}

function bindEvents() {
  const openLoginButton = document.getElementById("open-login-button");
  if (openLoginButton) {
    openLoginButton.addEventListener("click", () => {
      state.pendingActionId = null;
      state.showLoginPrompt = true;
      render();
    });
  }

  const closeLoginButton = document.getElementById("close-login-button");
  if (closeLoginButton) {
    closeLoginButton.addEventListener("click", () => {
      state.showLoginPrompt = false;
      render();
    });
  }

  const forgotPassword = document.getElementById("forgot-password");
  if (forgotPassword) {
    forgotPassword.addEventListener("click", () => {
      safeOpenWindow(buildAbsoluteAppUrl(state.config.forgotPasswordPath));
    });
  }

  const continueLoginButton = document.getElementById("continue-login-button");
  if (continueLoginButton) {
    continueLoginButton.addEventListener("click", () => {
      const actionId = state.pendingActionId;
      if (actionId) {
        launchAction(actionId);
        return;
      }

      safeOpenWindow(buildAbsoluteAppUrl(state.config.loginPath));
      state.showLoginPrompt = false;
      state.launchMessage = "TrackTalents login opened in a new tab.";
      render();
    });
  }

  document.querySelectorAll("[data-action-id]").forEach((button) => {
    button.addEventListener("click", () => {
      const actionId = button.getAttribute("data-action-id");
      if (!actionId) {
        return;
      }

      state.pendingActionId = actionId;
      state.showLoginPrompt = true;
      render();
    });
  });
}

function buildLaunchContext(action) {
  const item = state.currentItem || {};
  const query = new URLSearchParams({
    source: "outlook-addin",
    action: action.id,
    intent: action.intent,
    contextId: formatContextId(item),
    subject: item.subject || "",
    fromName: item.from?.displayName || "",
    fromEmail: item.from?.email || "",
    attachmentCount: String(item.attachmentCount || 0),
    hasResume: item.hasResumeAttachment ? "true" : "false",
    resumeFile: item.primaryResumeName || "",
    platform: state.officePlatform || "preview",
    host: state.officeHost || "preview"
  });

  return query.toString();
}

function buildAbsoluteAppUrl(pathname) {
  return new URL(pathname, `${state.config.appHost}/`).toString();
}

function buildTargetPath(actionId) {
  const action = ACTIONS.find((entry) => entry.id === actionId);
  if (!action) {
    throw new Error("Unknown action requested.");
  }

  return `${action.path}?${buildLaunchContext(action)}`;
}

function buildLaunchUrl(actionId) {
  const loginUrl = new URL(buildAbsoluteAppUrl(state.config.loginPath));
  loginUrl.searchParams.set("redirectTo", buildTargetPath(actionId));
  return loginUrl.toString();
}

function safeOpenWindow(url) {
  if (window.Office?.context?.ui?.openBrowserWindow) {
    Office.context.ui.openBrowserWindow(url);
    return;
  }

  window.open(url, "_blank", "noopener,noreferrer");
}

function launchAction(actionId) {
  try {
    const url = buildLaunchUrl(actionId);
    safeOpenWindow(url);
    state.showLoginPrompt = false;
    state.contextBanner = "";
    state.launchMessage = `${actionLabelFromId(actionId)} is opening through the TrackTalents login handoff, then the matching form will open with Outlook mail context attached.`;
    render();
  } catch (error) {
    state.launchMessage = "";
    state.contextBanner = error instanceof Error ? error.message : "Unable to open TrackTalents.";
    render();
  }
}

async function loadRuntimeConfig() {
  try {
    const response = await fetch("/api/config");
    if (!response.ok) {
      return;
    }

    const data = await response.json();
    state.config = {
      ...state.config,
      ...data
    };
  } catch {
    // Keep the local defaults when the config endpoint is unavailable.
  }
}

function setItemStateFromOffice(item) {
  const attachmentNames = Array.isArray(item.attachments)
    ? item.attachments.map((attachment) => attachment.name).filter(Boolean)
    : [];
  const resumeNames = attachmentNames.filter(isResumeFile);
  const fromDisplay =
    item.from?.emailAddress
      ? `${item.from.displayName || "Unknown"} <${item.from.emailAddress}>`
      : "Unavailable in this mode";

  state.currentItem = {
    itemId: item.itemId || item.internetMessageId || "",
    subject: item.subject || "",
    from: item.from
      ? {
          displayName: item.from.displayName || "",
          email: item.from.emailAddress || ""
        }
      : { displayName: "", email: "" },
    fromDisplay,
    toCount: Array.isArray(item.to) ? item.to.length : 0,
    attachmentCount: attachmentNames.length,
    attachmentNames,
    hasResumeAttachment: resumeNames.length > 0,
    primaryResumeName: resumeNames[0] || "",
    bodyPreview: "",
    mode: "outlook"
  };

  if (item.body?.getAsync) {
    item.body.getAsync(Office.CoercionType.Text, (result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded && state.currentItem) {
        state.currentItem.bodyPreview = String(result.value || "")
          .replace(/\s+/g, " ")
          .trim()
          .slice(0, 420);
        render();
      }
    });
  }
}

function readCurrentItem() {
  if (!state.officeReady || !window.Office?.context?.mailbox?.item) {
    if (!state.currentItem) {
      state.currentItem = buildPreviewItem();
    }
    return;
  }

  setItemStateFromOffice(Office.context.mailbox.item);
}

async function boot() {
  await loadRuntimeConfig();
  render();

  if (!window.Office?.onReady) {
    state.contextBanner = "Office.js is not available here, so the add-in is showing a preview email context.";
    render();
    return;
  }

  Office.onReady((info) => {
    state.officeReady = info.host === Office.HostType.Outlook;
    state.officeHost = info.host;
    state.officePlatform = info.platform;
    readCurrentItem();

    if (!state.officeReady) {
      state.contextBanner = "This page is in preview mode. Open it inside Outlook to read the real email context.";
    } else if (!state.currentItem?.hasResumeAttachment) {
      state.contextBanner = "Open an email that includes a resume so the add-in can pass resume context into TrackTalents.";
    } else {
      state.contextBanner = "Resume email detected. Choose any action to continue into TrackTalents.";
    }

    render();
  });
}

if (document.readyState === "loading") {
  document.addEventListener("DOMContentLoaded", boot);
} else {
  boot();
}
