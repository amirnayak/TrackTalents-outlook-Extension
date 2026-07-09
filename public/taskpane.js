const STORAGE_KEYS = {
  token: "token",
  loginData: "loginData",
  tokenExpiration: "tokenExpiration",
  userId: "userId",
  mid: "MId",
  rememberCheck: "rememberCheck"
};

const state = {
  officeReady: false,
  officeHost: null,
  officePlatform: null,
  session: loadSession(),
  currentItem: null,
  loginError: "",
  loggingIn: false
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

function loadSession() {
  const token = localStorage.getItem(STORAGE_KEYS.token);
  const loginDataRaw = localStorage.getItem(STORAGE_KEYS.loginData);
  const tokenExpirationRaw = localStorage.getItem(STORAGE_KEYS.tokenExpiration);

  if (!token || !loginDataRaw) {
    return null;
  }

  const tokenExpiration = Number(tokenExpirationRaw || "0");

  if (tokenExpiration && Number.isFinite(tokenExpiration) && Date.now() > tokenExpiration) {
    clearSession();
    return null;
  }

  try {
    const loginData = JSON.parse(loginDataRaw);
    return { token, loginData, tokenExpiration };
  } catch {
    clearSession();
    return null;
  }
}

function clearSession() {
  localStorage.removeItem(STORAGE_KEYS.token);
  localStorage.removeItem(STORAGE_KEYS.loginData);
  localStorage.removeItem(STORAGE_KEYS.tokenExpiration);
  localStorage.removeItem(STORAGE_KEYS.userId);
  localStorage.removeItem(STORAGE_KEYS.mid);
}

function persistSession(data, remember) {
  const accessToken = data.access_token;
  const expirationTimeInMs = Number(data.expires_in || 0) * 1000;
  const tokenExpiration = Date.now() + expirationTimeInMs;

  localStorage.setItem(STORAGE_KEYS.token, accessToken);
  localStorage.setItem(STORAGE_KEYS.loginData, JSON.stringify(data));
  localStorage.setItem(STORAGE_KEYS.tokenExpiration, String(tokenExpiration));
  localStorage.setItem(STORAGE_KEYS.userId, data.userId || data.UserId || "");
  localStorage.setItem(STORAGE_KEYS.mid, data.MId || "");

  if (remember) {
    localStorage.setItem(STORAGE_KEYS.rememberCheck, "true");
  } else {
    localStorage.removeItem(STORAGE_KEYS.rememberCheck);
  }

  state.session = {
    token: accessToken,
    tokenExpiration,
    loginData: data
  };
}

function getSessionUserName() {
  const loginData = state.session?.loginData || {};
  return (
    loginData.Email ||
    loginData.email ||
    loginData.userName ||
    loginData.userId ||
    loginData.UserId ||
    "TrackTalents user"
  );
}

function getSessionRole() {
  const role = state.session?.loginData?.Role;
  return role ? String(role) : "Recruiter";
}

function renderLogin() {
  const root = getRoot();
  root.innerHTML = `
    <main class="app-shell">
      <section class="login-layout">
        <section class="panel brand-panel">
          <p class="eyebrow">TrackTalents Outlook</p>
          <h1 class="hero-title">Recruiting<br />from your inbox</h1>
          <p class="hero-copy">
            Sign in first, then we’ll start wiring candidate, contact, and job workflows directly into Outlook.
          </p>
        </section>

        <section class="panel login-panel">
          <h2 class="section-title">Login</h2>
          <p class="section-copy">
            Use your TrackTalents account to continue into the Outlook add-in.
          </p>
          ${
            state.loginError
              ? `<div class="alert alert-error">${escapeHtml(state.loginError)}</div>`
              : ""
          }
          ${
            state.officeReady
              ? `<div class="alert alert-info">Outlook connection ready on ${escapeHtml(state.officePlatform || "unknown platform")}.</div>`
              : ""
          }
          <form id="login-form" class="form-stack">
            <div class="field">
              <label for="email">Email</label>
              <input id="email" name="email" type="email" placeholder="name@company.com" autocomplete="username" required />
              <span class="field-error" id="email-error"></span>
            </div>
            <div class="field">
              <label for="password">Password</label>
              <input id="password" name="password" type="password" placeholder="Enter your password" autocomplete="current-password" required />
              <span class="field-error" id="password-error"></span>
            </div>
            <div class="checkbox-row">
              <label class="checkbox">
                <input id="remember" name="remember" type="checkbox" ${localStorage.getItem(STORAGE_KEYS.rememberCheck) === "true" ? "checked" : ""} />
                <span>Remember me on this device</span>
              </label>
              <button id="forgot-password" type="button" class="link-button">Forgot password</button>
            </div>
            <button class="primary-button" type="submit" ${state.loggingIn ? "disabled" : ""}>
              ${
                state.loggingIn
                  ? '<span class="spinner" aria-hidden="true"></span>'
                  : "Continue"
              }
            </button>
          </form>
        </section>
      </section>
    </main>
  `;

  document.getElementById("login-form").addEventListener("submit", onLoginSubmit);
  document.getElementById("forgot-password").addEventListener("click", () => {
    window.open("https://test.tracktalents.com/forgotpassword", "_blank", "noopener,noreferrer");
  });
}

function renderApp() {
  const root = getRoot();
  const item = state.currentItem;
  const itemSubject = item?.subject || "No message selected";
  const itemFrom = item?.fromDisplay || "Open an Outlook email to load sender details.";
  const itemRecipients = item ? String(item.toCount) : "-";
  const itemAttachments = item ? String(item.attachmentCount) : "-";
  const hostChip = state.officeReady ? "chip chip-ready" : "chip chip-muted";
  const hostText = state.officeReady
    ? `Connected to Outlook${state.officePlatform ? ` on ${state.officePlatform}` : ""}`
    : "Preview mode";

  root.innerHTML = `
    <main class="app-shell">
      <section class="panel header-card">
        <div class="identity-row">
          <div class="identity-meta">
            <p class="eyebrow">TrackTalents Outlook</p>
            <h1>Welcome, ${escapeHtml(getSessionUserName())}</h1>
            <p>${escapeHtml(getSessionRole())} workspace</p>
          </div>
          <button id="logout-button" class="secondary-button" type="button">Logout</button>
        </div>
        <div class="${hostChip}">${escapeHtml(hostText)}</div>
      </section>

      <section class="panel status-card">
        <div class="toolbar">
          <div>
            <h2>Current Email Context</h2>
            <p>The app is ready to start building real workflows on top of this Outlook item.</p>
          </div>
          <button id="refresh-context" class="primary-button" type="button">Refresh Mail Data</button>
        </div>
      </section>

      <section class="grid feature-grid">
        <section class="panel feature-card">
          <h3>Add Candidate</h3>
          <p>Next build target. This will parse attachments, prefill candidate data, and create the record in TrackTalents.</p>
          <button class="secondary-button" type="button" disabled>Coming Next</button>
        </section>

        <section class="panel feature-card">
          <h3>Add Contact</h3>
          <p>Use the sender and recipients from the current email to create a TrackTalents contact.</p>
          <button class="secondary-button" type="button" disabled>Coming Soon</button>
        </section>

        <section class="panel feature-card">
          <h3>Add Job</h3>
          <p>Use email subject and body to create a job draft directly from Outlook.</p>
          <button class="secondary-button" type="button" disabled>Coming Soon</button>
        </section>
      </section>

      <section class="panel mail-card">
        <h3>Selected Mail Snapshot</h3>
        <div class="details-grid">
          <div class="detail-row">
            <dt>Subject</dt>
            <dd>${escapeHtml(itemSubject)}</dd>
          </div>
          <div class="detail-row">
            <dt>From</dt>
            <dd>${escapeHtml(itemFrom)}</dd>
          </div>
          <div class="detail-row">
            <dt>Recipients</dt>
            <dd>${escapeHtml(itemRecipients)}</dd>
          </div>
          <div class="detail-row">
            <dt>Attachments</dt>
            <dd>${escapeHtml(itemAttachments)}</dd>
          </div>
        </div>
      </section>
    </main>
  `;

  document.getElementById("logout-button").addEventListener("click", () => {
    clearSession();
    state.session = null;
    state.loginError = "";
    render();
  });

  document.getElementById("refresh-context").addEventListener("click", () => {
    readCurrentItem();
    render();
  });
}

function render() {
  if (state.session) {
    renderApp();
  } else {
    renderLogin();
  }
}

function setFieldError(id, message) {
  const element = document.getElementById(id);
  if (element) {
    element.textContent = message || "";
  }
}

function validateLoginForm(email, password) {
  let valid = true;

  setFieldError("email-error", "");
  setFieldError("password-error", "");

  if (!email || !/\S+@\S+\.\S+/.test(email)) {
    setFieldError("email-error", "Enter a valid email address.");
    valid = false;
  }

  if (!password || password.length < 6) {
    setFieldError("password-error", "Password must be at least 6 characters long.");
    valid = false;
  }

  return valid;
}

async function onLoginSubmit(event) {
  event.preventDefault();

  const form = event.currentTarget;
  const formData = new FormData(form);
  const email = String(formData.get("email") || "").trim();
  const password = String(formData.get("password") || "");
  const remember = formData.get("remember") === "on";

  if (!validateLoginForm(email, password)) {
    return;
  }

  state.loggingIn = true;
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
      throw new Error(data?.message || "Login failed.");
    }

    persistSession(data, remember);
    readCurrentItem();
  } catch (error) {
    state.loginError = error instanceof Error ? error.message : "Login failed.";
  } finally {
    state.loggingIn = false;
    render();
  }
}

function readCurrentItem() {
  if (!state.officeReady || !window.Office?.context?.mailbox?.item) {
    state.currentItem = null;
    return;
  }

  const item = Office.context.mailbox.item;
  const fromDisplay =
    item.from?.emailAddress
      ? `${item.from.displayName || "Unknown"} <${item.from.emailAddress}>`
      : "Unavailable in this mode";

  state.currentItem = {
    subject: item.subject || "",
    fromDisplay,
    toCount: Array.isArray(item.to) ? item.to.length : 0,
    attachmentCount: Array.isArray(item.attachments) ? item.attachments.length : 0
  };
}

function boot() {
  render();

  Office.onReady((info) => {
    state.officeReady = info.host === Office.HostType.Outlook;
    state.officeHost = info.host;
    state.officePlatform = info.platform;
    readCurrentItem();
    render();
  });
}

if (document.readyState === "loading") {
  document.addEventListener("DOMContentLoaded", boot);
} else {
  boot();
}
