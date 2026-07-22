const ACTIONS = [
  {
    id: "add-candidate",
    label: "Add Candidate",
    iconSrc: "/assets/action-icons/add-candidate.png",
    path: "/candidates",
    intent: "create-candidate"
  },
  {
    id: "submit-resume-contact",
    label: "Submit Resume To Contact",
    iconSrc: "/assets/action-icons/submit-resume-contact.png",
    path: "/contacts",
    intent: "submit-resume-contact"
  },
  {
    id: "add-job",
    label: "Add Job",
    iconSrc: "/assets/action-icons/add-job.png",
    path: "/jobs",
    intent: "create-job"
  },
  {
    id: "add-contact",
    label: "Add Contact",
    iconSrc: "/assets/action-icons/add-contact.png",
    path: "/contacts",
    intent: "create-contact"
  },
  {
    id: "source-resume-job",
    label: "Source Resume To Job",
    iconSrc: "/assets/action-icons/source-resume-job.png",
    path: "/local-search",
    intent: "source-resume-job"
  },
  {
    id: "reply-all",
    label: "Reply All",
    iconSrc: "/assets/action-icons/reply-all.png",
    path: "/sentmails",
    intent: "reply-all"
  }
];

const AUTH_STORAGE_KEY = "tracktalents-outlook-auth";

const state = {
  officeReady: false,
  officeHost: null,
  officePlatform: null,
  officeUser: null,
  currentItem: buildPreviewItem(),
  config: {
    appHost: "http://localhost:3000",
    authBridgePath: "/outlook-auth-bridge",
    loginPath: "/login",
    forgotPasswordPath: "/forgotpassword"
  },
  auth: loadPersistedAuth(),
  launchMessage: "",
  loginError: "",
  loginSubmitting: false,
  pendingActionId: null,
  showLoginModal: false,
  importModal: buildImportModalState(),
  loginForm: {
    email: "",
    password: ""
  }
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

function buildImportModalState() {
  return {
    open: false,
    actionId: null,
    attachments: [],
    resumeAttachmentId: "",
    error: "",
    submitting: false
  };
}

function buildPreviewAuth() {
  return {
    email: "preview.user@example.com",
    displayName: "Preview User",
    accessToken: "preview-token",
    refreshToken: "preview-refresh",
    tokenType: "Bearer",
    tokenExpiration: Date.now() + 60 * 60 * 1000,
    userId: "PREVIEW1",
    mId: "MID1",
    loginData: {
      userId: "PREVIEW1",
      MId: "MID1",
      expires_in: 3600
    },
    authenticatedAt: new Date().toISOString()
  };
}

function buildPreviewItem() {
  const attachments = [
    {
      id: "preview-att-1",
      name: "Ananya-Sharma-Resume.pdf",
      size: 348120,
      contentType: "application/pdf",
      isInline: false,
      previewText:
        "Ananya Sharma\nSenior Java Developer\nEmail: ananya.sharma@example.com\nPhone: +1 555-0102\nLocation: Dallas, TX\nSkills: Java, Spring Boot, AWS, Microservices\nExperience: 7 years\nCurrent Employer: Nimbus Systems"
    },
    {
      id: "preview-att-2",
      name: "portfolio.txt",
      size: 2480,
      contentType: "text/plain",
      isInline: false,
      previewText: "Portfolio links and certifications for Ananya Sharma."
    }
  ];
  const attachmentNames = attachments.map((attachment) => attachment.name);
  const resumeNames = attachmentNames.filter(isResumeFile);

  return {
    itemId: "preview-mail-001",
    subject: "Senior Java Developer Resume - Ananya Sharma",
    from: {
      displayName: "Ananya Sharma",
      email: "ananya.sharma@example.com"
    },
    fromDisplay: "Ananya Sharma <ananya.sharma@example.com>",
    toCount: 1,
    attachments,
    attachmentCount: attachmentNames.length,
    attachmentNames,
    hasResumeAttachment: resumeNames.length > 0,
    primaryResumeName: resumeNames[0] || "",
    bodyPreview:
      "Hello team, please find my latest resume attached for the senior Java developer role. I have 7 years of backend and cloud experience.",
    mode: "preview"
  };
}

function isResumeFile(name) {
  return /\.(pdf|doc|docx|rtf|txt)$/i.test(String(name || ""));
}

function isAddCandidateAction(actionId) {
  return actionId === "add-candidate";
}

function formatContextId(item) {
  const source = item?.itemId || "preview-context";
  return String(source).replace(/[^a-zA-Z0-9]/g, "").slice(0, 18) || "previewcontext";
}

function formatFileSize(bytes) {
  const size = Number(bytes || 0);
  if (size <= 0) {
    return "0 KB";
  }

  if (size >= 1024 * 1024) {
    return `${(size / (1024 * 1024)).toFixed(1)} MB`;
  }

  return `${Math.max(1, Math.round(size / 1024))} KB`;
}

function normalizeAttachmentContentFormat(format) {
  if (typeof format !== "string") {
    return "base64";
  }

  const normalized = format.trim().toLowerCase();
  if (normalized === "url") {
    return "url";
  }

  if (normalized === "eml") {
    return "eml";
  }

  if (normalized === "icalendar") {
    return "icalendar";
  }

  return "base64";
}

function loadPersistedAuth() {
  try {
    const raw = window.localStorage.getItem(AUTH_STORAGE_KEY);
    if (!raw) {
      return null;
    }

    const parsed = JSON.parse(raw);
    if (!parsed || typeof parsed.email !== "string" || typeof parsed.accessToken !== "string") {
      return null;
    }

    return parsed;
  } catch {
    return null;
  }
}

function persistAuth(auth) {
  try {
    if (!auth) {
      window.localStorage.removeItem(AUTH_STORAGE_KEY);
      return;
    }

    window.localStorage.setItem(AUTH_STORAGE_KEY, JSON.stringify(auth));
  } catch {
    // Ignore storage failures and keep the in-memory session alive.
  }
}

function isAuthenticated() {
  return Boolean(
    state.auth?.email &&
      state.auth?.accessToken &&
      state.auth?.loginData &&
      (state.auth?.userId || state.auth?.loginData?.userId || state.auth?.loginData?.UserId) &&
      (state.auth?.mId || state.auth?.loginData?.MId)
  );
}

function extractDisplayNameFromEmail(email) {
  const localPart = String(email || "").split("@")[0] || "";

  return localPart
    .split(/[._-]+/)
    .filter(Boolean)
    .map((part) => part.charAt(0).toUpperCase() + part.slice(1))
    .join(" ");
}

function getWelcomeName() {
  return (
    state.officeUser?.displayName ||
    state.auth?.displayName ||
    extractDisplayNameFromEmail(state.auth?.email || "") ||
    "there"
  );
}

function actionLabelFromId(actionId) {
  return ACTIONS.find((action) => action.id === actionId)?.label || "TrackTalents action";
}

function createImportAttachmentSelection(attachment, index) {
  return {
    id: attachment.id || `attachment-${index + 1}`,
    name: attachment.name || `Attachment ${index + 1}`,
    size: Number(attachment.size || 0),
    contentType: attachment.contentType || "application/octet-stream",
    isInline: Boolean(attachment.isInline),
    previewText: attachment.previewText || "",
    canBeResume: isResumeFile(attachment.name),
    contentFormat: normalizeAttachmentContentFormat(attachment.contentFormat),
    content: attachment.content || "",
    source: attachment.source || "email"
  };
}

function getImportableAttachments() {
  return Array.isArray(state.currentItem?.attachments)
    ? state.currentItem.attachments.filter((attachment) => !attachment.isInline)
    : [];
}

function getSelectedImportResume() {
  return state.importModal.attachments.find(
    (attachment) => attachment.id === state.importModal.resumeAttachmentId
  );
}

function getImportSummary() {
  const totalAttachments = state.importModal.attachments.length;
  const selectedResume = getSelectedImportResume();
  const documentCount = totalAttachments - (selectedResume ? 1 : 0);

  return {
    totalAttachments,
    documentCount,
    selectedResume
  };
}

function closeImportModal() {
  state.importModal = buildImportModalState();
}

function openImportModal(actionId) {
  state.importModal = buildImportModalForAction(actionId);
  render();
}

function buildImportModalForAction(actionId) {
  const attachments = getImportableAttachments().map(createImportAttachmentSelection);
  const defaultResume = attachments.find((attachment) => attachment.canBeResume);

  return {
    open: true,
    actionId,
    attachments,
    resumeAttachmentId: defaultResume?.id || "",
    error: "",
    submitting: false
  };
}

function handleActionRequest(actionId) {
  if (isAddCandidateAction(actionId)) {
    openImportModal(actionId);
    return;
  }

  launchAction(actionId);
}

function buildAttachmentBadgeMarkup(count) {
  return `<span class="import-badge">${escapeHtml(String(count))}</span>`;
}

function renderWelcomeNote() {
  return `
    <div class="welcome-note">
      <p class="welcome-copy">Welcome, ${escapeHtml(getWelcomeName())}</p>
      <span class="welcome-tag">${isAuthenticated() ? "Signed In" : "Sign In Required"}</span>
    </div>
  `;
}

function renderLoginModal() {
  return `
    <div class="modal-overlay" role="presentation">
      <section class="login-dialog" role="dialog" aria-modal="true" aria-labelledby="login-title">
        <div class="login-header">
          <div>
            <p class="section-label">TrackTalents Access</p>
            <h2 id="login-title">Sign in inside Outlook</h2>
          </div>
          <button id="close-login-button" class="icon-button" type="button" aria-label="Close login dialog">X</button>
        </div>

        <p class="login-copy">
          Sign in once here. After that, every action button will open the matching TrackTalents create form in a separate tab.
        </p>

        ${
          state.loginError
            ? `<div class="banner banner-error modal-banner">${escapeHtml(state.loginError)}</div>`
            : ""
        }

        <form id="login-form" class="login-form">
          <label class="field">
            <span>Email</span>
            <input
              id="login-email"
              name="email"
              type="email"
              value="${escapeAttribute(state.loginForm.email)}"
              autocomplete="email"
              required
            />
          </label>

          <label class="field">
            <span>Password</span>
            <input
              id="login-password"
              name="password"
              type="password"
              value="${escapeAttribute(state.loginForm.password)}"
              autocomplete="current-password"
              required
            />
          </label>

          <div class="login-actions">
            <button class="primary-button" type="submit" ${state.loginSubmitting ? "disabled" : ""}>
              ${state.loginSubmitting ? "Signing In..." : "Sign In"}
            </button>
            <button id="forgot-password-button" class="text-button" type="button">
              Forgot Password
            </button>
          </div>
        </form>
      </section>
    </div>
  `;
}

function renderImportAttachmentRow(attachment) {
  const isResume = attachment.id === state.importModal.resumeAttachmentId;

  return `
    <div class="import-attachment-row">
      <div class="import-attachment-copy">
        <strong>${escapeHtml(attachment.name)}</strong>
        <span>${escapeHtml(formatFileSize(attachment.size))}${
          attachment.canBeResume ? " • Resume supported" : ""
        }</span>
      </div>

      <label class="resume-checkbox ${attachment.canBeResume ? "" : "resume-checkbox-disabled"}">
        <input
          type="checkbox"
          data-resume-select-id="${escapeAttribute(attachment.id)}"
          ${isResume ? "checked" : ""}
          ${attachment.canBeResume ? "" : "disabled"}
        />
        <span>Resume</span>
      </label>

      <button
        type="button"
        class="delete-attachment-button"
        data-remove-attachment-id="${escapeAttribute(attachment.id)}"
        aria-label="Remove ${escapeAttribute(attachment.name)}"
      >
        <svg class="delete-attachment-icon" viewBox="0 0 24 24" fill="none" aria-hidden="true">
          <path d="M3 6h18"></path>
          <path d="M8 6V4h8v2"></path>
          <path d="M19 6l-1 14H6L5 6"></path>
          <path d="M10 11v6"></path>
          <path d="M14 11v6"></path>
        </svg>
      </button>
    </div>
  `;
}

function renderImportModal() {
  const summary = getImportSummary();
  const hasAttachments = summary.totalAttachments > 0;
  const selectedResume = summary.selectedResume;
  const importDisabled = !hasAttachments || !selectedResume || state.importModal.submitting;
  const submitLabel = state.importModal.submitting ? "Importing..." : "Import";
  const importSummaryText = selectedResume
    ? `1 resume and ${summary.documentCount} document${
        summary.documentCount === 1 ? "" : "s"
      } will be imported on submit`
    : "Choose one attachment as the resume before importing";

  return `
    <div class="modal-overlay" role="presentation">
      <section class="import-dialog" role="dialog" aria-modal="true" aria-labelledby="import-title">
        <div class="import-header">
          <div>
            <p class="section-label">Add Candidate</p>
            <h2 id="import-title">Import attachments from this email</h2>
          </div>
        </div>

        <div class="import-summary-row">
          <div class="import-info-banner">
            ${buildAttachmentBadgeMarkup(summary.totalAttachments)}
            <span>${escapeHtml(importSummaryText)}</span>
          </div>
        </div>

        ${
          state.importModal.error
            ? `<div class="banner banner-error">${escapeHtml(state.importModal.error)}</div>`
            : ""
        }

        ${
          hasAttachments
            ? `<div class="import-attachment-list">${state.importModal.attachments
                .map(renderImportAttachmentRow)
                .join("")}</div>`
            : `<div class="import-empty-state">No supported attachments were found on this email.</div>`
        }

        <div class="import-footer">
          <div class="import-footer-actions">
            <button
              id="cancel-import-button"
              class="ghost-button import-footer-button"
              type="button"
              ${state.importModal.submitting ? "disabled" : ""}
            >
              Cancel
            </button>
            <button
              id="import-attachments-button"
              class="primary-button import-footer-primary"
              type="button"
              ${importDisabled ? "disabled" : ""}
            >
              ${escapeHtml(submitLabel)}
            </button>
          </div>
        </div>
      </section>
    </div>
  `;
}

function render() {
  const root = getRoot();
  const authLabel = isAuthenticated() ? "Logout" : "Login";
  const authButtonId = isAuthenticated() ? "logout-button" : "open-login-button";

  root.innerHTML = `
    <main class="shell">
      <section class="main-card">
        <div class="hero-strip">
          <div class="hero-top">
            <div class="brand-lockup">
              <div class="brand-mark">
                <img src="/assets/track-talents-profile.png" alt="TrackTalents" class="brand-mark-image" />
              </div>
              <div>
                <h1>TrackTalents Outlook</h1>
                <p class="hero-copy">TrackTalents Outlook will add Candidates, Clients, Job etc directly from your email into TrackTalents ATS</p>
              </div>
            </div>
            <button id="${authButtonId}" class="ghost-button" type="button">${authLabel}</button>
          </div>

          ${renderWelcomeNote()}
        </div>

        ${
          state.launchMessage
            ? `<section class="status-stack"><div class="banner banner-success">${escapeHtml(state.launchMessage)}</div></section>`
            : ""
        }

        <div class="actions-frame">
          <div class="actions-head">
            <p class="section-label">Quick Actions</p>
          </div>

          <div class="action-list">
            ${ACTIONS.map(renderActionButton).join("")}
          </div>
        </div>
      </section>
    </main>

    ${state.showLoginModal ? renderLoginModal() : ""}
    ${state.importModal.open ? renderImportModal() : ""}
  `;

  bindEvents();
}

function renderActionButton(action) {
  return `
    <button
      class="action-button"
      type="button"
      data-action-id="${escapeAttribute(action.id)}"
      style="--action-watermark: url('${escapeAttribute(action.iconSrc)}');"
      aria-label="${escapeAttribute(action.label)}"
    >
      <span class="action-icon" aria-hidden="true">
        <img src="${escapeAttribute(action.iconSrc)}" alt="" class="action-icon-image" />
      </span>
      <span class="action-copy">
        <strong>${escapeHtml(action.label)}</strong>
      </span>
    </button>
  `;
}

function bindEvents() {
  const openLoginButton = document.getElementById("open-login-button");
  if (openLoginButton) {
    openLoginButton.addEventListener("click", () => {
      state.showLoginModal = true;
      state.loginError = "";
      render();
    });
  }

  const logoutButton = document.getElementById("logout-button");
  if (logoutButton) {
    logoutButton.addEventListener("click", () => {
      state.auth = null;
      state.pendingActionId = null;
      state.showLoginModal = false;
      state.loginError = "";
      state.loginSubmitting = false;
      state.loginForm.password = "";
      closeImportModal();
      persistAuth(null);
      state.launchMessage = "You have been logged out from the TrackTalents extension.";
      render();
    });
  }

  const closeLoginButton = document.getElementById("close-login-button");
  if (closeLoginButton) {
    closeLoginButton.addEventListener("click", () => {
      state.showLoginModal = false;
      state.loginError = "";
      render();
    });
  }

  const forgotPasswordButton = document.getElementById("forgot-password-button");
  if (forgotPasswordButton) {
    forgotPasswordButton.addEventListener("click", () => {
      safeOpenWindow(buildAbsoluteAppUrl(state.config.forgotPasswordPath));
    });
  }

  const loginForm = document.getElementById("login-form");
  if (loginForm) {
    loginForm.addEventListener("submit", handleLoginSubmit);
  }

  const loginEmail = document.getElementById("login-email");
  if (loginEmail) {
    loginEmail.addEventListener("input", (event) => {
      state.loginForm.email = event.target.value;
    });
  }

  const loginPassword = document.getElementById("login-password");
  if (loginPassword) {
    loginPassword.addEventListener("input", (event) => {
      state.loginForm.password = event.target.value;
    });
  }

  const cancelImportButton = document.getElementById("cancel-import-button");
  if (cancelImportButton) {
    cancelImportButton.addEventListener("click", () => {
      if (!state.importModal.submitting) {
        closeImportModal();
        render();
      }
    });
  }

  const importAttachmentsButton = document.getElementById("import-attachments-button");
  if (importAttachmentsButton) {
    importAttachmentsButton.addEventListener("click", handleImportSubmit);
  }

  document.querySelectorAll("[data-remove-attachment-id]").forEach((button) => {
    button.addEventListener("click", () => {
      if (state.importModal.submitting) {
        return;
      }

      const attachmentId = button.getAttribute("data-remove-attachment-id");
      if (!attachmentId) {
        return;
      }

      state.importModal.attachments = state.importModal.attachments.filter(
        (attachment) => attachment.id !== attachmentId
      );

      if (state.importModal.resumeAttachmentId === attachmentId) {
        const nextResume = state.importModal.attachments.find((attachment) => attachment.canBeResume);
        state.importModal.resumeAttachmentId = nextResume?.id || "";
      }

      state.importModal.error = "";
      render();
    });
  });

  document.querySelectorAll("[data-resume-select-id]").forEach((input) => {
    input.addEventListener("change", () => {
      if (state.importModal.submitting) {
        return;
      }

      const attachmentId = input.getAttribute("data-resume-select-id");
      if (!attachmentId) {
        return;
      }

      const isChecked = Boolean(input.checked);
      state.importModal.resumeAttachmentId = isChecked ? attachmentId : "";
      state.importModal.error = "";
      render();
    });
  });

  document.querySelectorAll("[data-action-id]").forEach((button) => {
    button.addEventListener("click", () => {
      const actionId = button.getAttribute("data-action-id");
      if (!actionId) {
        return;
      }

      if (!isAuthenticated()) {
        state.pendingActionId = actionId;
        state.showLoginModal = true;
        state.loginError = "";
        state.launchMessage = `Sign in to continue to ${actionLabelFromId(actionId)}.`;
        render();
        return;
      }

      handleActionRequest(actionId);
    });
  });
}

async function handleLoginSubmit(event) {
  event.preventDefault();

  const email = state.loginForm.email.trim();
  const password = state.loginForm.password;

  if (!email || !password) {
    state.loginError = "Enter both your email and password to continue.";
    render();
    return;
  }

  state.loginSubmitting = true;
  state.loginError = "";
  render();

  try {
    const response = await fetch("/api/auth/login", {
      method: "POST",
      headers: {
        "Content-Type": "application/json"
      },
      body: JSON.stringify({ email, password })
    });

    const data = await response.json().catch(() => ({}));

    if (!response.ok) {
      throw new Error(data?.message || "Login failed. Please verify your credentials.");
    }

    state.auth = {
      email,
      displayName: String(data?.name || extractDisplayNameFromEmail(email)),
      accessToken: String(data?.access_token || ""),
      tokenType: String(data?.token_type || "Bearer"),
      refreshToken: String(data?.refresh_token || ""),
      tokenExpiration: Date.now() + Number(data?.expires_in || 0) * 1000,
      userId: String(data?.userId || data?.UserId || ""),
      mId: String(data?.MId || ""),
      loginData: data,
      authenticatedAt: new Date().toISOString()
    };
    persistAuth(state.auth);

    state.loginSubmitting = false;
    state.loginError = "";
    state.showLoginModal = false;
    state.loginForm.password = "";
    state.launchMessage = `Welcome ${getWelcomeName()}. You are signed in and ready to open TrackTalents create forms in separate tabs.`;

    const pendingActionId = state.pendingActionId;
    state.pendingActionId = null;

    if (pendingActionId) {
      handleActionRequest(pendingActionId);
      return;
    }

    render();
  } catch (error) {
    state.loginSubmitting = false;
    state.loginError = error instanceof Error ? error.message : "Unable to sign in right now.";
    render();
  }
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

  if (state.auth?.email) {
    query.set("signedInUser", state.auth.email);
  }

  return query.toString();
}

function buildAbsoluteAppUrl(pathname) {
  return new URL(pathname, `${state.config.appHost}/`).toString();
}

function encodeBridgePayload(value) {
  const json = JSON.stringify(value);
  const utf8 = new TextEncoder().encode(json);
  let binary = "";

  utf8.forEach((byte) => {
    binary += String.fromCharCode(byte);
  });

  return btoa(binary).replaceAll("+", "-").replaceAll("/", "_").replaceAll("=", "");
}

function buildTargetPath(actionId) {
  const action = ACTIONS.find((entry) => entry.id === actionId);
  if (!action) {
    throw new Error("Unknown action requested.");
  }

  return `${action.path}?${buildLaunchContext(action)}`;
}

function buildAuthBridgeUrl(actionId, outlookCandidateImport) {
  const redirectTo = buildTargetPath(actionId);
  const loginExpirySeconds = Number(state.auth?.loginData?.expires_in || 0);
  const fallbackExpiration =
    loginExpirySeconds > 0 ? Date.now() + loginExpirySeconds * 1000 : Date.now() + 60 * 60 * 1000;
  const tokenExpiration = Number(state.auth?.tokenExpiration || 0) || fallbackExpiration;
  const payload = encodeBridgePayload({
    accessToken: state.auth?.accessToken || "",
    refreshToken: state.auth?.refreshToken || "",
    tokenType: state.auth?.tokenType || "Bearer",
    tokenExpiration,
    email: state.auth?.email || "",
    userId: state.auth?.userId || state.auth?.loginData?.userId || state.auth?.loginData?.UserId || "",
    mId: state.auth?.mId || state.auth?.loginData?.MId || "",
    loginData: state.auth?.loginData || {},
    outlookCandidateImport: outlookCandidateImport || null
  });
  const hash = new URLSearchParams({
    payload,
    redirectTo
  });

  return `${buildAbsoluteAppUrl(state.config.authBridgePath)}#${hash.toString()}`;
}

function safeOpenWindow(url) {
  if (window.Office?.context?.ui?.openBrowserWindow) {
    Office.context.ui.openBrowserWindow(url);
    return;
  }

  window.open(url, "_blank", "noopener,noreferrer");
}

function launchAction(actionId, options = {}) {
  try {
    const url = buildAuthBridgeUrl(actionId, options.outlookCandidateImport);
    safeOpenWindow(url);
    state.launchMessage =
      options.successMessage || `${actionLabelFromId(actionId)} opened in a separate tab.`;
    state.showLoginModal = false;
    closeImportModal();
    render();
  } catch (error) {
    state.launchMessage = "";
    state.loginError = error instanceof Error ? error.message : "Unable to open TrackTalents.";
    render();
  }
}

async function resolveAttachmentForImport(attachment) {
  if (attachment?.content) {
    return attachment;
  }

  if (state.currentItem?.mode === "preview") {
    return {
      ...attachment,
      contentFormat: "base64",
      content: utf8TextToBase64(
        attachment.previewText || `${attachment.name}\nTrackTalents preview attachment content`
      )
    };
  }

  const contentResult = await readOfficeAttachmentContent(attachment.id);
  return {
    ...attachment,
    contentFormat: normalizeAttachmentContentFormat(contentResult.format),
    content: String(contentResult.content || "")
  };
}

function utf8TextToBase64(value) {
  const utf8 = new TextEncoder().encode(String(value || ""));
  let binary = "";

  utf8.forEach((byte) => {
    binary += String.fromCharCode(byte);
  });

  return btoa(binary);
}

function readOfficeAttachmentContent(attachmentId) {
  return new Promise((resolve, reject) => {
    const item = window.Office?.context?.mailbox?.item;
    if (!item?.getAttachmentContentAsync) {
      reject(
        new Error("This Outlook client does not support reading attachment content in the task pane.")
      );
      return;
    }

    item.getAttachmentContentAsync(attachmentId, (result) => {
      if (result.status !== Office.AsyncResultStatus.Succeeded) {
        reject(
          new Error(
            result.error?.message ||
              "TrackTalents could not read one of the selected Outlook attachments."
          )
        );
        return;
      }

      resolve(result.value || {});
    });
  });
}

async function handleImportSubmit() {
  const summary = getImportSummary();
  const selectedResume = summary.selectedResume;

  if (!selectedResume) {
    state.importModal.error = "Select one attachment as the resume before importing.";
    render();
    return;
  }

  if (!state.auth?.accessToken) {
    state.importModal.error = "Your TrackTalents session expired. Please sign in again.";
    render();
    return;
  }

  state.importModal.submitting = true;
  state.importModal.error = "";
  render();

  try {
    const attachments = await Promise.all(
      state.importModal.attachments.map(resolveAttachmentForImport)
    );
    const response = await fetch("/api/candidate/import-from-email", {
      method: "POST",
      headers: {
        "Content-Type": "application/json"
      },
      body: JSON.stringify({
        accessToken: state.auth.accessToken,
        userId:
          state.auth.userId || state.auth.loginData?.userId || state.auth.loginData?.UserId || "",
        selectedResumeId: selectedResume.id,
        previewMode: state.currentItem?.mode === "preview",
        emailContext: {
          subject: state.currentItem?.subject || "",
          fromName: state.currentItem?.from?.displayName || "",
          fromEmail: state.currentItem?.from?.email || "",
          bodyPreview: state.currentItem?.bodyPreview || ""
        },
        attachments
      })
    });

    const payload = await response.json().catch(() => ({}));
    if (!response.ok) {
      throw new Error(payload?.message || "Unable to import the selected attachments.");
    }

    launchAction(state.importModal.actionId || "add-candidate", {
      outlookCandidateImport: payload,
      successMessage: `Add Candidate opened with ${selectedResume.name} selected as the resume.`
    });
  } catch (error) {
    state.importModal.submitting = false;
    state.importModal.error =
      error instanceof Error ? error.message : "Unable to import the selected attachments.";
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

function normalizeOfficeAttachment(attachment, index) {
  return {
    id: attachment.id || `office-att-${index + 1}`,
    name: attachment.name || `Attachment ${index + 1}`,
    size: Number(attachment.size || 0),
    contentType: attachment.contentType || "application/octet-stream",
    isInline: Boolean(attachment.isInline),
    contentFormat: "base64",
    content: ""
  };
}

function setItemStateFromOffice(item) {
  const attachments = Array.isArray(item.attachments)
    ? item.attachments.map(normalizeOfficeAttachment)
    : [];
  const attachmentNames = attachments.map((attachment) => attachment.name).filter(Boolean);
  const resumeNames = attachmentNames.filter(isResumeFile);
  const fromDisplay = item.from?.emailAddress
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
    attachments,
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
  if (!state.officeReady) {
    if (!state.currentItem) {
      state.currentItem = buildPreviewItem();
    }
    return;
  }

  const item = window.Office?.context?.mailbox?.item;
  if (!item) {
    state.currentItem = {
      itemId: "",
      subject: "",
      from: { displayName: "", email: "" },
      fromDisplay: "No message selected",
      toCount: 0,
      attachments: [],
      attachmentCount: 0,
      attachmentNames: [],
      hasResumeAttachment: false,
      primaryResumeName: "",
      bodyPreview: "",
      mode: "outlook"
    };
    render();
    return;
  }

  setItemStateFromOffice(item);
}

function applyPreviewPageStateFromUrl() {
  if (state.currentItem?.mode !== "preview") {
    return;
  }

  const params = new URLSearchParams(window.location.search);
  if (params.get("previewAuth") === "1" && !isAuthenticated()) {
    state.auth = buildPreviewAuth();
    persistAuth(state.auth);
  }

  const previewAction = params.get("previewAction");
  if (previewAction === "add-candidate" && isAuthenticated()) {
    state.importModal = buildImportModalForAction(previewAction);
  }
}

function setOfficeUserState() {
  const profile = window.Office?.context?.mailbox?.userProfile;
  if (!profile) {
    return;
  }

  state.officeUser = {
    displayName: profile.displayName || "",
    email: profile.emailAddress || ""
  };
}

function handleOutlookItemChanged() {
  readCurrentItem();

  if (state.importModal.open && state.importModal.actionId) {
    state.importModal = buildImportModalForAction(state.importModal.actionId);
  }

  render();
}

async function boot() {
  await loadRuntimeConfig();
  applyPreviewPageStateFromUrl();
  render();

  if (!window.Office?.onReady) {
    render();
    return;
  }

  Office.onReady((info) => {
    state.officeReady = info.host === Office.HostType.Outlook;
    state.officeHost = info.host;
    state.officePlatform = info.platform;
    setOfficeUserState();
    readCurrentItem();
    applyPreviewPageStateFromUrl();

    if (state.officeReady && Office.context?.mailbox?.addHandlerAsync) {
      Office.context.mailbox.addHandlerAsync(Office.EventType.ItemChanged, handleOutlookItemChanged);
    }

    render();
  });
}

if (document.readyState === "loading") {
  document.addEventListener("DOMContentLoaded", boot);
} else {
  boot();
}
