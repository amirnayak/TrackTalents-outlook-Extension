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
    id: "attach-email",
    label: "Link Emails",
    iconSrc: "/assets/action-icons/attach-email.svg",
    path: "/sentmails",
    intent: "attach-email"
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
  // {
  //   id: "reply-all",
  //   label: "Reply All",
  //   iconSrc: "/assets/action-icons/reply-all.png",
  //   path: "/sentmails",
  //   intent: "reply-all"
  // }
];

const AUTH_STORAGE_KEY = "tracktalents-outlook-auth";
const ATTACH_EMAIL_PAGE_SIZE = 10;

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
  attachEmail: buildAttachEmailState(),
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

function buildAttachEmailState() {
  return {
    open: false,
    activeTab: "contacts",
    search: "",
    loading: false,
    error: "",
    submitting: false,
    contacts: [],
    candidates: [],
    contactsTotal: 0,
    candidatesTotal: 0,
    contactsPage: 0,
    candidatesPage: 0,
    selectedContacts: [],
    selectedCandidates: []
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
    to: [
      {
        displayName: "TrackTalents Hiring",
        email: "hiring@tracktalents.com"
      }
    ],
    fromDisplay: "Ananya Sharma <ananya.sharma@example.com>",
    toCount: 1,
    attachments,
    attachmentCount: attachmentNames.length,
    attachmentNames,
    hasResumeAttachment: resumeNames.length > 0,
    primaryResumeName: resumeNames[0] || "",
    bodyPreview:
      "Hello team, please find my latest resume attached for the senior Java developer role. I have 7 years of backend and cloud experience.",
    bodyHtml:
      "<p>Hello team,</p><p>Please find my latest resume attached for the senior Java developer role. I have <strong>7 years</strong> of backend and cloud experience.</p>",
    mode: "preview"
  };
}

function isResumeFile(name) {
  return /\.(pdf|doc|docx|rtf|txt)$/i.test(String(name || ""));
}

function isAddCandidateAction(actionId) {
  return actionId === "add-candidate";
}

function isResumeImportAction(actionId) {
  return (
    actionId === "add-candidate" ||
    actionId === "submit-resume-contact" ||
    actionId === "source-resume-job"
  );
}

function isAttachEmailAction(actionId) {
  return actionId === "attach-email";
}

function isEmailParserAction(actionId) {
  return actionId === "add-contact" || actionId === "add-job";
}

function getEmailAddinRecordType(actionId) {
  switch (actionId) {
    case "add-job":
    case "source-resume-job":
      return "job";
    case "add-contact":
    case "submit-resume-contact":
      return "contact";
    case "reply-all":
      return "email";
    case "add-candidate":
    default:
      return "candidate";
  }
}

function formatContextId(item) {
  const source = item?.itemId || "preview-context";
  return String(source).replace(/[^a-zA-Z0-9]/g, "").slice(0, 18) || "previewcontext";
}

function formatFileSize(bytes) {
  const size = Number(bytes || 0);
  if (size <= 0) {
    return "File Size - 0KB";
  }

  if (size >= 1024 * 1024) {
    return `File Size - ${(size / (1024 * 1024)).toFixed(1)}MB`;
  }

  return `File Size - ${Math.max(1, Math.round(size / 1024))}KB`;
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

function normalizeRecipient(recipient) {
  if (!recipient || typeof recipient !== "object") {
    return {
      displayName: "",
      email: ""
    };
  }

  return {
    displayName:
      recipient.displayName ||
      recipient.name ||
      recipient.Name ||
      "",
    email:
      recipient.email ||
      recipient.emailAddress ||
      recipient.Email ||
      ""
  };
}

function normalizeRecipientList(recipients) {
  if (!Array.isArray(recipients)) {
    return [];
  }

  return recipients
    .map(normalizeRecipient)
    .filter((recipient) => recipient.displayName || recipient.email);
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
  const attachments = getImportableAttachments()
    .map(createImportAttachmentSelection)
    .filter((attachment) => (isResumeImportAction(actionId) ? attachment.canBeResume : true));
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

function getAttachEmailSelectedTotal() {
  return state.attachEmail.selectedContacts.length + state.attachEmail.selectedCandidates.length;
}

function getAttachEmailPageCount(type) {
  const total =
    type === "candidates"
      ? Number(state.attachEmail.candidatesTotal || 0)
      : Number(state.attachEmail.contactsTotal || 0);
  return Math.max(1, Math.ceil(total / ATTACH_EMAIL_PAGE_SIZE));
}

function getAttachEmailCurrentPage(type) {
  return type === "candidates"
    ? Number(state.attachEmail.candidatesPage || 0)
    : Number(state.attachEmail.contactsPage || 0);
}

function getAttachEmailRows(type) {
  return type === "candidates" ? state.attachEmail.candidates : state.attachEmail.contacts;
}

function isAttachEmailSelected(type, recordId) {
  const selected =
    type === "candidates" ? state.attachEmail.selectedCandidates : state.attachEmail.selectedContacts;
  return selected.some((item) => item.id === recordId);
}

function resetAttachEmailState() {
  state.attachEmail = buildAttachEmailState();
}

function closeAttachEmailPanel() {
  resetAttachEmailState();
}

function buildAttachEmailSearchFilter(type, search) {
  const value = String(search || "").trim();
  if (!value) {
    return "";
  }

  if (type === "candidates") {
    return JSON.stringify([
      ["CandidateData.FirstName", "contains", value],
      "or",
      ["CandidateData.LastName", "contains", value],
      "or",
      ["CandidateData.Contact.Email1", "contains", value],
      "or",
      ["CandidateData.JobTitle", "contains", value],
      "or",
      ["CandidateData.CurrentLocation", "contains", value]
    ]);
  }

  return JSON.stringify([
    ["ContactData.Name", "contains", value],
    "or",
    ["ContactData.Contact.Email1", "contains", value],
    "or",
    ["ContactData.CompanyName", "contains", value],
    "or",
    ["ContactData.JobTitle", "contains", value]
  ]);
}

function normalizeAttachEmailContact(record) {
  const contactData = record?.ContactData && typeof record.ContactData === "object" ? record.ContactData : {};
  const contactInfo = contactData?.Contact && typeof contactData.Contact === "object" ? contactData.Contact : {};
  const firstName = String(contactData.FirstName || "").trim();
  const lastName = String(contactData.LastName || "").trim();
  const fullName = String(contactData.Name || `${firstName} ${lastName}`).trim() || "Unnamed Contact";

  return {
    id: String(record?._id || ""),
    name: fullName,
    email: String(contactInfo.Email1 || contactInfo.Email2 || ""),
    company: String(contactData.CompanyName || ""),
    subtitle: String(contactData.JobTitle || ""),
    raw: record
  };
}

function normalizeAttachEmailCandidate(record) {
  const candidateData =
    record?.CandidateData && typeof record.CandidateData === "object" ? record.CandidateData : {};
  const contactInfo =
    candidateData?.Contact && typeof candidateData.Contact === "object" ? candidateData.Contact : {};
  const firstName = String(candidateData.FirstName || "").trim();
  const lastName = String(candidateData.LastName || "").trim();
  const fullName = `${firstName} ${lastName}`.trim() || "Unnamed Candidate";

  return {
    id: String(record?._id || ""),
    name: fullName,
    email: String(contactInfo.Email1 || contactInfo.Email2 || ""),
    title: String(candidateData.JobTitle || ""),
    location: String(candidateData.CurrentLocation || candidateData.Relocation || ""),
    raw: record
  };
}

async function fetchAttachEmailRecords(type, options = {}) {
  const page = Math.max(0, Number(options.page || 0));
  const search = String(options.search ?? state.attachEmail.search ?? "").trim();
  const skip = page * ATTACH_EMAIL_PAGE_SIZE;
  const filter = buildAttachEmailSearchFilter(type, search);
  const query = new URLSearchParams({
    skip: String(skip),
    take: String(ATTACH_EMAIL_PAGE_SIZE),
    requireTotalCount: "true"
  });

  if (filter) {
    query.set("filter", filter);
  }

  const response = await fetch(`/api/attach-email/${type}?${query.toString()}`, {
    method: "GET",
    headers: {
      Authorization: `Bearer ${state.auth?.accessToken || ""}`
    }
  });

  const payload = await response.json().catch(() => ({}));
  if (!response.ok) {
    throw new Error(payload?.message || `Unable to load TrackTalents ${type}.`);
  }

  const rows = Array.isArray(payload?.data) ? payload.data : [];
  const totalCount = Math.max(0, Number(payload?.totalCount || 0));

  return {
    rows:
      type === "candidates"
        ? rows.map(normalizeAttachEmailCandidate)
        : rows.map(normalizeAttachEmailContact),
    totalCount
  };
}

async function loadAttachEmailRecords(options = {}) {
  const loadContacts = options.loadContacts !== false;
  const loadCandidates = options.loadCandidates !== false;
  const contactsPage = Math.max(0, Number(options.contactsPage ?? state.attachEmail.contactsPage));
  const candidatesPage = Math.max(
    0,
    Number(options.candidatesPage ?? state.attachEmail.candidatesPage)
  );
  const renderLoading = options.renderLoading !== false;

  state.attachEmail.loading = true;
  state.attachEmail.error = "";
  if (renderLoading) {
    render();
  }

  try {
    const tasks = [];

    if (loadContacts) {
      tasks.push(
        fetchAttachEmailRecords("contacts", {
          page: contactsPage,
          search: state.attachEmail.search
        }).then((result) => ({
          type: "contacts",
          result
        }))
      );
    }

    if (loadCandidates) {
      tasks.push(
        fetchAttachEmailRecords("candidates", {
          page: candidatesPage,
          search: state.attachEmail.search
        }).then((result) => ({
          type: "candidates",
          result
        }))
      );
    }

    const results = await Promise.allSettled(tasks);

    const errors = [];

    results.forEach((entry) => {
      if (entry.status === "fulfilled") {
        if (entry.value.type === "contacts") {
          state.attachEmail.contacts = entry.value.result.rows;
          state.attachEmail.contactsTotal = entry.value.result.totalCount;
          state.attachEmail.contactsPage = contactsPage;
          return;
        }

        if (entry.value.type === "candidates") {
          state.attachEmail.candidates = entry.value.result.rows;
          state.attachEmail.candidatesTotal = entry.value.result.totalCount;
          state.attachEmail.candidatesPage = candidatesPage;
        }

        return;
      }

      if (entry.reason instanceof Error) {
        errors.push(entry.reason.message);
      } else if (entry.reason) {
        errors.push(String(entry.reason));
      }
    });

    state.attachEmail.loading = false;
    state.attachEmail.error = errors[0] || "";
    render();
  } catch (error) {
    state.attachEmail.loading = false;
    state.attachEmail.error =
      error instanceof Error ? error.message : "Unable to load TrackTalents records.";
    render();
  }
}

function openAttachEmailPanel() {
  state.attachEmail = buildAttachEmailState();
  state.attachEmail.open = true;
  render();
  void loadAttachEmailRecords({
    loadContacts: true,
    loadCandidates: true
  });
}

function toggleAttachEmailSelection(type, recordId) {
  if (!recordId) {
    return;
  }

  const rows = getAttachEmailRows(type);
  const selectedKey = type === "candidates" ? "selectedCandidates" : "selectedContacts";
  const selectedRows = state.attachEmail[selectedKey];
  const existingIndex = selectedRows.findIndex((item) => item.id === recordId);

  if (existingIndex >= 0) {
    state.attachEmail[selectedKey] = selectedRows.filter((item) => item.id !== recordId);
  } else {
    const nextRecord = rows.find((item) => item.id === recordId);
    if (!nextRecord) {
      return;
    }

    state.attachEmail[selectedKey] = [...selectedRows, nextRecord];
  }

  render();
}

function buildAttachEmailSummaryCopy() {
  const contactCount = state.attachEmail.selectedContacts.length;
  const candidateCount = state.attachEmail.selectedCandidates.length;

  if (!contactCount && !candidateCount) {
    return "Choose one or more contacts or candidates to prepare this email for attachment.";
  }

  return `${contactCount} contact${contactCount === 1 ? "" : "s"} and ${candidateCount} candidate${
    candidateCount === 1 ? "" : "s"
  } selected`;
}

function buildAttachEmailActivityNote(recordId) {
  const item = state.currentItem || {};
  const fromName = String(item.from?.displayName || "").trim();
  const fromEmail = String(item.from?.email || "").trim();
  const subject = String(item.subject || "").trim();
  const body = String(item.bodyPreview || "")
    .replace(/\s+/g, " ")
    .trim()
    .slice(0, 1400);
  const attachmentCount = Number(item.attachmentCount || 0);
  const attachmentSummary =
    attachmentCount > 0 ? `Attachments: ${attachmentCount}` : "Attachments: 0";
  const emailSource = fromName || fromEmail
    ? `From: ${fromName || "Unknown"}${fromEmail ? ` <${fromEmail}>` : ""}`
    : "From: Unknown";

  return [
    "Outlook email attached from TrackTalents Outlook add-in.",
    recordId ? `Email Add-in Record ID: ${recordId}` : "",
    emailSource,
    subject ? `Subject: ${subject}` : "Subject: (empty)",
    attachmentSummary,
    body ? `Body: ${body}` : "Body: (empty)"
  ]
    .filter(Boolean)
    .join("\n");
}

function buildAttachEmailSuccessMessage(contactCount, candidateCount) {
  return `Email linked to ${contactCount} contact${contactCount === 1 ? "" : "s"} and ${candidateCount} candidate${
    candidateCount === 1 ? "" : "s"
  }.`;
}

async function attachEmailToTrackTalentsRecord(
  entityType,
  entityId,
  noteDescription,
  userId,
  documents = []
) {
  const response = await fetch("/api/attach-email/link", {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
      Authorization: `Bearer ${state.auth?.accessToken || ""}`
    },
    body: JSON.stringify({
      entityType,
      entityId,
      noteDescription,
      userId: String(userId || "").trim(),
      documents: Array.isArray(documents) ? documents : []
    })
  });

  const payload = await response.json().catch(() => ({}));
  if (!response.ok) {
    throw new Error(payload?.message || `Unable to link the email to the selected ${entityType}.`);
  }

  return payload;
}

function buildAttachEmailSummaryRows() {
  const rows = [];

  if (state.attachEmail.selectedContacts.length) {
    rows.push({
      label: `Contacts (${state.attachEmail.selectedContacts.length})`,
      values: state.attachEmail.selectedContacts.map((contact) => contact.name)
    });
  }

  if (state.attachEmail.selectedCandidates.length) {
    rows.push({
      label: `Candidates (${state.attachEmail.selectedCandidates.length})`,
      values: state.attachEmail.selectedCandidates.map((candidate) => candidate.name)
    });
  }

  return rows;
}

function renderAttachEmailSummaryRows() {
  const rows = buildAttachEmailSummaryRows();
  if (!rows.length) {
    return "";
  }

  return `
    <div class="attach-email-summary-list">
      ${rows
        .map(
          (row) => `
            <div class="attach-email-summary-item">
              <strong>${escapeHtml(row.label)}</strong>
              <div class="attach-email-chip-list">
                ${row.values
                  .map((value) => `<span class="attach-email-chip">${escapeHtml(value)}</span>`)
                  .join("")}
              </div>
            </div>
          `
        )
        .join("")}
    </div>
  `;
}

function renderAttachEmailTable() {
  const activeTab = state.attachEmail.activeTab;
  const rows = getAttachEmailRows(activeTab);
  const columnTemplate =
    activeTab === "candidates"
      ? "minmax(180px, 1.15fr) minmax(220px, 1fr) minmax(160px, 0.85fr)"
      : "minmax(180px, 1.05fr) minmax(220px, 1fr) minmax(180px, 0.9fr)";
  const tableMinWidth = activeTab === "candidates" ? "620px" : "640px";

  if (!rows.length) {
    return `
      <div class="attach-email-empty-state">
        ${escapeHtml(
          state.attachEmail.loading
            ? "Loading TrackTalents records..."
            : `No ${activeTab} matched this search.`
        )}
      </div>
    `;
  }

  if (activeTab === "candidates") {
    return `
      <div class="attach-email-table-shell">
        <div
          class="attach-email-table"
          style="--attach-email-columns: ${escapeAttribute(columnTemplate)}; --attach-email-table-min-width: ${escapeAttribute(tableMinWidth)};"
        >
          <div class="attach-email-table-head">
            <span>Name</span>
            <span>Email</span>
            <span>Location</span>
          </div>
          <div class="attach-email-table-body">
            ${rows
              .map((candidate) => {
                const selected = isAttachEmailSelected("candidates", candidate.id);
                return `
                  <button
                    type="button"
                    class="attach-email-row ${selected ? "attach-email-row-selected" : ""}"
                    data-attach-email-row-type="candidates"
                    data-attach-email-row-id="${escapeAttribute(candidate.id)}"
                  >
                    <span>${escapeHtml(candidate.name)}</span>
                    <span>${escapeHtml(candidate.email || "—")}</span>
                    <span>${escapeHtml(candidate.location || "—")}</span>
                  </button>
                `;
              })
              .join("")}
          </div>
        </div>
      </div>
    `;
  }

  return `
    <div class="attach-email-table-shell">
      <div
        class="attach-email-table"
        style="--attach-email-columns: ${escapeAttribute(columnTemplate)}; --attach-email-table-min-width: ${escapeAttribute(tableMinWidth)};"
      >
        <div class="attach-email-table-head">
          <span>Name</span>
          <span>Email</span>
          <span>Company</span>
        </div>
        <div class="attach-email-table-body">
          ${rows
            .map((contact) => {
              const selected = isAttachEmailSelected("contacts", contact.id);
              return `
                <button
                  type="button"
                  class="attach-email-row ${selected ? "attach-email-row-selected" : ""}"
                  data-attach-email-row-type="contacts"
                  data-attach-email-row-id="${escapeAttribute(contact.id)}"
                >
                  <span>${escapeHtml(contact.name)}</span>
                  <span>${escapeHtml(contact.email || "—")}</span>
                  <span>${escapeHtml(contact.company || "—")}</span>
                </button>
              `;
            })
            .join("")}
        </div>
      </div>
    </div>
  `;
}

function renderAttachEmailPanel() {
  const activeTab = state.attachEmail.activeTab;
  const currentPage = getAttachEmailCurrentPage(activeTab);
  const pageCount = getAttachEmailPageCount(activeTab);
  const hasSelection = getAttachEmailSelectedTotal() > 0;
  const submitLabel = state.attachEmail.submitting ? "Linking..." : "Link Email";
  const searchPlaceholder = "Search any column...";

  return `
    <div class="modal-overlay" role="presentation">
      <section
        class="attach-email-dialog"
        role="dialog"
        aria-modal="true"
        aria-labelledby="attach-email-title"
      >
        <div class="attach-email-header">
          <div>
            <p class="section-label">Link Emails</p>
            <h2 id="attach-email-title">Link this email to TrackTalents records</h2>
          </div>
        </div>

        <div class="attach-email-summary ${hasSelection ? "attach-email-summary-active" : ""}">
          <div class="attach-email-summary-copy">
            <strong>${hasSelection ? "Ready to link" : "Select records"}</strong>
            <span>${escapeHtml(buildAttachEmailSummaryCopy())}</span>
          </div>
          ${renderAttachEmailSummaryRows()}
        </div>

        ${state.attachEmail.error ? `<div class="banner banner-error">${escapeHtml(state.attachEmail.error)}</div>` : ""}

        <div class="attach-email-tabs" role="tablist" aria-label="TrackTalents lists">
          <button
            type="button"
            class="attach-email-tab ${activeTab === "contacts" ? "attach-email-tab-active" : ""}"
            data-attach-email-tab="contacts"
          >
            Contacts
          </button>
          <button
            type="button"
            class="attach-email-tab ${activeTab === "candidates" ? "attach-email-tab-active" : ""}"
            data-attach-email-tab="candidates"
          >
            Candidates
          </button>
        </div>

        <div class="attach-email-toolbar">
          <input
            id="attach-email-search"
            class="attach-email-search"
            type="search"
            placeholder="${escapeAttribute(searchPlaceholder)}"
            value="${escapeAttribute(state.attachEmail.search)}"
          />
        </div>

        ${renderAttachEmailTable()}

        <div class="attach-email-footer">
          <div class="attach-email-pager">
            <button
              type="button"
              class="ghost-button attach-email-pager-button"
              id="attach-email-prev-page"
              ${state.attachEmail.loading || currentPage <= 0 ? "disabled" : ""}
            >
              Prev
            </button>
            <span class="attach-email-pager-copy">
              ${escapeHtml(`${currentPage + 1} / ${pageCount}`)}
            </span>
            <button
              type="button"
              class="ghost-button attach-email-pager-button"
              id="attach-email-next-page"
              ${state.attachEmail.loading || currentPage + 1 >= pageCount ? "disabled" : ""}
            >
              Next
            </button>
          </div>

          <div class="attach-email-footer-actions">
            <button
              id="cancel-attach-email-button"
              class="ghost-button attach-email-footer-button"
              type="button"
              ${state.attachEmail.submitting ? "disabled" : ""}
            >
              Cancel
            </button>
            <button
              id="submit-attach-email-button"
              class="primary-button attach-email-footer-primary"
              type="button"
              ${!hasSelection || state.attachEmail.loading || state.attachEmail.submitting ? "disabled" : ""}
            >
              ${escapeHtml(submitLabel)}
            </button>
          </div>
        </div>
      </section>
    </div>
  `;
}

function handleActionRequest(actionId) {
  if (isAttachEmailAction(actionId)) {
    openAttachEmailPanel();
    return;
  }

  if (isResumeImportAction(actionId)) {
    openImportModal(actionId);
    return;
  }

  void handleDirectActionLaunch(actionId);
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
            <h2 id="login-title">Sign in to TrackTalents</h2>
          </div>
          <button id="close-login-button" class="icon-button" type="button" aria-label="Close login dialog">X</button>
        </div>

        <p class="login-copy">
          Use your TrackTalents Credentials to signin to Link your Outlook emails to TrackTalents Records
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
        <span>${escapeHtml(formatFileSize(attachment.size))}</span>
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
  const actionId = state.importModal.actionId;
  const actionLabel = actionLabelFromId(actionId);
  const submitLabel = state.importModal.submitting ? "Importing..." : "Import";
  const additionalDocumentsLabel = `${summary.documentCount} additional document${
    summary.documentCount === 1 ? "" : "s"
  }`;

  let modalTitle = "Select the resume from this email";
  let importSummaryText = selectedResume
    ? `1 resume and ${additionalDocumentsLabel} will be imported`
    : "Choose one attachment as the resume before importing";
  let emptyStateText = "No resume-supported attachments were found on this email.";

  switch (actionId) {
    case "add-job":
      modalTitle = "Select the document to create this job";
      importSummaryText = selectedResume
        ? `1 document and ${additionalDocumentsLabel} will be imported into the job form`
        : "Choose one attachment to parse into the job form before importing";
      emptyStateText = "No resume-supported job documents were found on this email.";
      break;
    case "add-contact":
      modalTitle = "Select the document to create this contact";
      importSummaryText = selectedResume
        ? `1 document and ${additionalDocumentsLabel} will be imported into the contact form`
        : "Choose one attachment to parse into the contact form before importing";
      emptyStateText = "No resume-supported contact documents were found on this email.";
      break;
    case "submit-resume-contact":
      modalTitle = "Select the resume to submit to this contact";
      importSummaryText = selectedResume
        ? `1 resume and ${additionalDocumentsLabel} will be imported for contact submission`
        : "Choose one attachment as the resume before submitting to the contact";
      emptyStateText = "No resume-supported attachments were found for contact submission on this email.";
      break;
    case "source-resume-job":
      modalTitle = "Select the resume to source to this job";
      importSummaryText = selectedResume
        ? `1 resume and ${additionalDocumentsLabel} will be imported for job sourcing`
        : "Choose one attachment as the resume before sourcing it to the job";
      emptyStateText = "No resume-supported attachments were found for job sourcing on this email.";
      break;
    case "add-candidate":
    default:
      break;
  }

  return `
    <div class="modal-overlay" role="presentation">
      <section class="import-dialog" role="dialog" aria-modal="true" aria-labelledby="import-title">
        <div class="import-header">
          <div>
            <p class="section-label">${escapeHtml(actionLabel)}</p>
            <h2 id="import-title">${escapeHtml(modalTitle)}</h2>
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
            : `<div class="import-empty-state">${escapeHtml(emptyStateText)}</div>`
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
          <div class="hero-card">
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
        </div>

        <div class="actions-frame">
          <div class="actions-surface">
            <div class="actions-head">
              <p class="section-label">Quick Actions</p>
            </div>

            <div class="action-list">
              ${ACTIONS.map(renderActionButton).join("")}
            </div>
          </div>
        </div>
      </section>
    </main>

    ${state.showLoginModal ? renderLoginModal() : ""}
    ${state.attachEmail.open ? renderAttachEmailPanel() : ""}
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
      closeAttachEmailPanel();
      closeImportModal();
      persistAuth(null);
      state.launchMessage = "";
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

  const cancelAttachEmailButton = document.getElementById("cancel-attach-email-button");
  if (cancelAttachEmailButton) {
    cancelAttachEmailButton.addEventListener("click", () => {
      if (!state.attachEmail.submitting) {
        closeAttachEmailPanel();
        render();
      }
    });
  }

  const submitAttachEmailButton = document.getElementById("submit-attach-email-button");
  if (submitAttachEmailButton) {
    submitAttachEmailButton.addEventListener("click", handleAttachEmailSubmit);
  }

  const attachEmailSearch = document.getElementById("attach-email-search");
  if (attachEmailSearch) {
    attachEmailSearch.addEventListener("input", (event) => {
      state.attachEmail.search = event.target.value;
      state.attachEmail.contactsPage = 0;
      state.attachEmail.candidatesPage = 0;
      void loadAttachEmailRecords({
        loadContacts: true,
        loadCandidates: true,
        contactsPage: 0,
        candidatesPage: 0,
        renderLoading: false
      });
    });
  }

  document.querySelectorAll("[data-attach-email-tab]").forEach((button) => {
    button.addEventListener("click", () => {
      const tab = button.getAttribute("data-attach-email-tab");
      if (!tab || tab === state.attachEmail.activeTab) {
        return;
      }

      state.attachEmail.activeTab = tab === "candidates" ? "candidates" : "contacts";
      render();
    });
  });

  document.querySelectorAll("[data-attach-email-row-id]").forEach((button) => {
    button.addEventListener("click", () => {
      const recordId = button.getAttribute("data-attach-email-row-id");
      const type = button.getAttribute("data-attach-email-row-type");

      if (!recordId || !type || state.attachEmail.submitting) {
        return;
      }

      toggleAttachEmailSelection(type === "candidates" ? "candidates" : "contacts", recordId);
    });
  });

  const attachEmailPrevPage = document.getElementById("attach-email-prev-page");
  if (attachEmailPrevPage) {
    attachEmailPrevPage.addEventListener("click", () => {
      const activeTab = state.attachEmail.activeTab;
      const nextPage = Math.max(0, getAttachEmailCurrentPage(activeTab) - 1);

      if (activeTab === "candidates") {
        void loadAttachEmailRecords({
          loadContacts: false,
          loadCandidates: true,
          candidatesPage: nextPage
        });
        return;
      }

      void loadAttachEmailRecords({
        loadContacts: true,
        loadCandidates: false,
        contactsPage: nextPage
      });
    });
  }

  const attachEmailNextPage = document.getElementById("attach-email-next-page");
  if (attachEmailNextPage) {
    attachEmailNextPage.addEventListener("click", () => {
      const activeTab = state.attachEmail.activeTab;
      const nextPage = getAttachEmailCurrentPage(activeTab) + 1;

      if (activeTab === "candidates") {
        void loadAttachEmailRecords({
          loadContacts: false,
          loadCandidates: true,
          candidatesPage: nextPage
        });
        return;
      }

      void loadAttachEmailRecords({
        loadContacts: true,
        loadCandidates: false,
        contactsPage: nextPage
      });
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
        state.launchMessage = "";
        render();
        return;
      }

      handleActionRequest(actionId);
    });
  });
}

async function handleAttachEmailSubmit() {
  const contactCount = state.attachEmail.selectedContacts.length;
  const candidateCount = state.attachEmail.selectedCandidates.length;

  if (!contactCount && !candidateCount) {
    state.attachEmail.error = "Select at least one contact or candidate before attaching this email.";
    render();
    return;
  }

  state.attachEmail.submitting = true;
  state.attachEmail.error = "";
  render();

  try {
    const importedDocuments = await prepareEmailAddinDocuments(getImportableAttachments());
    const emailAddinRecord = await createEmailAddinRecord("email", {
      body: state.currentItem?.bodyPreview || "",
      bodyHtml: state.currentItem?.bodyHtml || "",
      subject: state.currentItem?.subject || "",
      fromName: state.currentItem?.from?.displayName || "",
      fromEmail: state.currentItem?.from?.email || "",
      toRecipients: normalizeRecipientList(state.currentItem?.to),
      resumes: null,
      documents: importedDocuments
    });
    const recordId = String(emailAddinRecord?._id || "");
    const noteDescription = buildAttachEmailActivityNote(recordId);
    const activityUserId = String(
      state.auth?.userId ||
        state.auth?.loginData?.userId ||
        state.auth?.loginData?.UserId ||
        ""
    ).trim();

    const attachTasks = [
      ...state.attachEmail.selectedContacts.map((contact) =>
        attachEmailToTrackTalentsRecord(
          "contacts",
          contact.id,
          noteDescription,
          activityUserId,
          importedDocuments
        )
      ),
      ...state.attachEmail.selectedCandidates.map((candidate) =>
        attachEmailToTrackTalentsRecord(
          "candidates",
          candidate.id,
          noteDescription,
          activityUserId,
          importedDocuments
        )
      )
    ];

    await Promise.all(attachTasks);

    state.launchMessage = "";
    closeAttachEmailPanel();
    render();
  } catch (error) {
    state.attachEmail.submitting = false;
    state.attachEmail.error =
      error instanceof Error ? error.message : "Unable to prepare this email for attachment.";
    render();
  }
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
    state.launchMessage = "";

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

function buildResumeRecordEntry(attachment) {
  if (!attachment) {
    return null;
  }

  return {
    FileName: String(attachment.name || ""),
    ContentType: String(attachment.contentType || "application/octet-stream"),
    ContentFormat: String(attachment.contentFormat || "base64"),
    Content: String(attachment.content || ""),
    Size: Number(attachment.size || 0),
    Source: String(attachment.source || "email")
  };
}

function buildResumeBridgePreview(attachment) {
  if (!attachment) {
    return null;
  }

  return {
    FileName: String(attachment.name || ""),
    ContentType: String(attachment.contentType || "application/octet-stream"),
    Size: Number(attachment.size || 0),
    Source: String(attachment.source || "email")
  };
}

async function createOutlookImportSession(payload) {
  const response = await fetch("/api/outlook-import-session", {
    method: "POST",
    headers: {
      "Content-Type": "application/json"
    },
    body: JSON.stringify({
      payload
    })
  });

  const data = await response.json().catch(() => ({}));
  if (!response.ok) {
    throw new Error(data?.message || "Unable to prepare the Outlook import session.");
  }

  return {
    sessionId: String(data?.sessionId || ""),
    apiHost: window.location.origin
  };
}

async function updateOutlookImportSession(sessionId, payload) {
  const response = await fetch(
    `/api/outlook-import-session/${encodeURIComponent(String(sessionId || ""))}`,
    {
      method: "POST",
      headers: {
        "Content-Type": "application/json"
      },
      body: JSON.stringify({
        payload
      })
    }
  );

  const data = await response.json().catch(() => ({}));
  if (!response.ok) {
    throw new Error(data?.message || "Unable to update the Outlook import session.");
  }

  return {
    sessionId: String(data?.sessionId || sessionId || ""),
    apiHost: window.location.origin
  };
}

async function parseResumeAttachment(attachment) {
  const response = await fetch("/api/resume/parse", {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
      Authorization: `Bearer ${state.auth?.accessToken || ""}`
    },
    body: JSON.stringify({
      attachment
    })
  });

  const payload = await response.json().catch(() => ({}));
  if (!response.ok) {
    throw new Error(payload?.message || "Unable to parse the selected Outlook resume.");
  }

  return payload;
}

function buildEmailParserRequest(actionId) {
  const item = state.currentItem || {};
  const body = firstNonEmpty(toPlainString(item.bodyHtml), toPlainString(item.bodyPreview), "");

  return {
    parse_type: getEmailAddinRecordType(actionId),
    source: "outlook_extension",
    subject: toPlainString(item.subject) || null,
    from_name: toPlainString(item.from?.displayName) || null,
    from_email: toPlainString(item.from?.email) || null,
    to: normalizeRecipientList(item.to)
      .map((recipient) => recipient.email)
      .filter(Boolean),
    cc: [],
    sent_at: null,
    message_id: toPlainString(item.itemId) || null,
    thread_id: toPlainString(item.conversationId) || null,
    body: body || null
  };
}

async function parseCurrentEmailForAction(actionId) {
  const response = await fetch("/api/email/parse", {
    method: "POST",
    headers: {
      "Content-Type": "application/json"
    },
    body: JSON.stringify(buildEmailParserRequest(actionId))
  });

  const payload = await response.json().catch(() => ({}));
  if (!response.ok) {
    throw new Error(payload?.message || "Unable to parse the current Outlook email.");
  }

  return payload;
}

function splitFullName(fullName) {
  const [firstName, ...lastNameParts] = String(fullName || "").trim().split(/\s+/).filter(Boolean);
  return {
    firstName: firstName || "",
    lastName: lastNameParts.join(" ")
  };
}

function toPlainString(value) {
  if (typeof value === "string") {
    return value.trim();
  }

  if (typeof value === "number" || typeof value === "boolean") {
    return String(value);
  }

  return "";
}

function firstNonEmpty(...values) {
  return values.find((value) => {
    if (Array.isArray(value)) {
      return value.length > 0;
    }

    return value !== undefined && value !== null && String(value).trim() !== "";
  });
}

function normalizePayType(value) {
  const text = String(value || "").toLowerCase();
  if (!text) {
    return "";
  }

  if (/\b(hour|hourly|\/hr|\/hour|w2|c2c)\b/.test(text)) {
    return "Hourly";
  }

  if (/\b(year|yearly|annual|annum|salary)\b/.test(text)) {
    return "Yearly";
  }

  if (/\b(month|monthly)\b/.test(text)) {
    return "Monthly";
  }

  return "";
}

function buildParsedResumeDataFromEmailParser(actionId, parserResult) {
  const data = parserResult?.structured_data || {};

  if (actionId === "add-contact") {
    const nameParts = splitFullName(data.full_name);

    return {
      FirstName: data.first_name || nameParts.firstName,
      LastName: data.last_name || nameParts.lastName,
      JobTitle: data.job_title || "",
      CompanyName: data.company_name || "",
      Notes: data.notes || "",
      Contact: {
        Email1: data.email || "",
        CellNumber: data.cell_number || "",
        WorkNumber: data.work_number || "",
        DirectNumber: "",
        StreetAddress: data.street_address || "",
        City: data.city || "",
        State: data.state || "",
        PostalCode: data.postal_code || "",
        Country: data.country || "USA"
      },
      StreetAddress: data.street_address || "",
      City: data.city || "",
      State: data.state || "",
      PostalCode: data.postal_code || "",
      Country: data.country || "USA",
      EmailParserData: data,
      EmailParserConfidence: parserResult?.confidence_summary || ""
    };
  }

  const requiredSkills = Array.isArray(data.required_skills) ? data.required_skills : [];
  const preferredSkills = Array.isArray(data.preferred_skills) ? data.preferred_skills : [];
  const salaryOrBudget = data.salary_or_budget || "";
  const yearsOfExperience = firstNonEmpty(data.years_of_experience, data.experience_required, "");
  const payType = firstNonEmpty(data.pay_type, normalizePayType(salaryOrBudget), "");

  return {
    JobTitle: data.job_title || "",
    CompanyName: data.company_name || "",
    ClientName: data.company_name || "",
    Location: data.location || "",
    WorkMode: data.work_mode || "",
    EmploymentType: data.employment_type || "",
    ExperienceRequired: data.experience_required || "",
    ExperienceLevel: data.experience_level || "",
    YearsOfExperience: yearsOfExperience || "",
    NoOfOpenings: data.no_of_openings || "",
    PayRate: data.pay_rate || salaryOrBudget,
    ClientRate: data.client_rate || salaryOrBudget,
    PayType: payType || "",
    JobType: data.job_type || "",
    SalaryOrBudget: salaryOrBudget,
    NoticePeriod: data.notice_period || "",
    JobSummary: data.job_summary || "",
    JobDescription: data.job_description_raw || data.job_summary || "",
    RequiredSkills: requiredSkills,
    PreferredSkills: preferredSkills,
    skills: [...requiredSkills, ...preferredSkills]
      .map((skill) => String(skill || "").trim())
      .filter(Boolean)
      .map((skill) => ({ skill })),
    EmailParserData: data,
    EmailParserConfidence: parserResult?.confidence_summary || ""
  };
}

function buildEmailParserImportOptions(actionId, parserResult, parserWarning = "") {
  const data = parserResult?.structured_data || {};
  const warnings = [
    ...(Array.isArray(parserResult?.warnings) ? parserResult.warnings : []),
    parserWarning
  ].filter(Boolean);
  const options = {
    parsedResumeData: parserResult
      ? buildParsedResumeDataFromEmailParser(actionId, parserResult)
      : null,
    emailParserResult: parserResult || null,
    warnings
  };

  if (actionId === "add-job" && data.job_title) {
    options.subject = data.job_title;
  }

  if (actionId === "add-job" && (data.job_description_raw || data.job_summary)) {
    options.bodyPreview = data.job_description_raw || data.job_summary;
  }

  return options;
}

async function prepareImportDocuments(attachments) {
  const response = await fetch("/api/outlook-import-documents", {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
      Authorization: `Bearer ${state.auth?.accessToken || ""}`
    },
    body: JSON.stringify({
      userId:
        String(
          state.auth?.userId ||
            state.auth?.loginData?.userId ||
            state.auth?.loginData?.UserId ||
            ""
        ),
      attachments
    })
  });

  const payload = await response.json().catch(() => ({}));
  if (!response.ok) {
    throw new Error(payload?.message || "Unable to prepare the selected Outlook documents.");
  }

  return Array.isArray(payload?.documents) ? payload.documents : [];
}

function buildDocumentPayload(documents) {
  if (!Array.isArray(documents) || documents.length === 0) {
    return null;
  }

  const documentValues = documents
    .map((document) => {
      if (typeof document === "string") {
        const value = String(document).trim();
        return value || null;
      }

      if (document && typeof document === "object") {
        return document;
      }

      return null;
    })
    .filter(Boolean);

  if (documentValues.length === 0) {
    return null;
  }

  return documentValues;
}

function buildEmailAddinRecordDebugPayload(payload) {
  if (!payload || typeof payload !== "object") {
    return payload;
  }

  return {
    ...payload,
    emailData:
      payload.emailData && typeof payload.emailData === "object"
        ? {
            ...payload.emailData,
            BodyHtml: payload.emailData.BodyHtml
              ? `[redacted html length=${String(payload.emailData.BodyHtml).length}]`
              : ""
          }
        : payload.emailData,
    accessToken: payload.accessToken ? "[redacted]" : "",
    resumes: Array.isArray(payload.resumes)
      ? payload.resumes.map((resume) => ({
          ...resume,
          Content: resume?.Content ? `[redacted base64 length=${String(resume.Content).length}]` : ""
        }))
      : payload.resumes
  };
}

function logEmailAddinRecordDebug(stage, details) {
  const timestamp = new Date().toISOString();
  console.groupCollapsed(`[TrackTalents][EmailAddinRecord] ${stage} ${timestamp}`);
  console.log(details);
  console.groupEnd();
}

function buildEmailAddinRecordRequest(type, options = {}) {
  const item = state.currentItem || {};

  return {
    type,
    resumes: Array.isArray(options.resumes) ? options.resumes : null,
    documents: buildDocumentPayload(options.documents),
    emailData: {
      Body: String(options.body ?? item.bodyPreview ?? ""),
      BodyHtml: String(options.bodyHtml ?? item.bodyHtml ?? ""),
      Subject: String(options.subject ?? item.subject ?? ""),
      From: {
        Name: String(options.fromName ?? item.from?.displayName ?? ""),
        Email: String(options.fromEmail ?? item.from?.email ?? "")
      },
      To: normalizeRecipientList(options.toRecipients ?? item.to).map((recipient) => ({
        Name: recipient.displayName,
        Email: recipient.email
      }))
    }
  };
}

async function createEmailAddinRecord(type, options = {}) {
  const requestPayload = buildEmailAddinRecordRequest(type, options);
  logEmailAddinRecordDebug("POST request", buildEmailAddinRecordDebugPayload(requestPayload));

  const response = await fetch("/api/EmailAddinRecord", {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
      Authorization: `Bearer ${state.auth?.accessToken || ""}`
    },
    body: JSON.stringify(requestPayload)
  });

  const payload = await response.json().catch(() => ({}));
  logEmailAddinRecordDebug("POST response", {
    ok: response.ok,
    status: response.status,
    payload
  });

  if (!response.ok) {
    throw new Error(payload?.message || "Unable to create the email add-in record.");
  }

  return payload;
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

function buildBridgeLoginData() {
  const loginData = state.auth?.loginData || {};

  return {
    expires_in: Number(loginData.expires_in || 0),
    userId: String(loginData.userId || state.auth?.userId || ""),
    UserId: String(loginData.UserId || loginData.userId || state.auth?.userId || ""),
    MId: String(loginData.MId || state.auth?.mId || ""),
    Role: String(loginData.Role || ""),
    Permissions: loginData.Permissions || "[]",
    Databases: String(loginData.Databases || ""),
    Admin: loginData.Admin ?? false,
    OnboardingLicense: loginData.OnboardingLicense ?? false,
    Email: String(loginData.Email || state.auth?.email || ""),
    FirstName: String(loginData.FirstName || ""),
    LastName: String(loginData.LastName || "")
  };
}

function buildTargetPath(actionId) {
  const action = ACTIONS.find((entry) => entry.id === actionId);
  if (!action) {
    throw new Error("Unknown action requested.");
  }

  return `${action.path}?${buildLaunchContext(action)}`;
}

function buildOutlookActionImportPayload(actionId, options = {}) {
  return {
    actionId,
    source: "outlook-addin",
    importedAt: new Date().toISOString(),
    selectedResumeName: options.selectedResume?.FileName || "",
    emailContext: {
      subject: options.subject || state.currentItem?.subject || "",
      fromName: options.fromName || state.currentItem?.from?.displayName || "",
      fromEmail: options.fromEmail || state.currentItem?.from?.email || "",
      bodyPreview: options.bodyPreview || state.currentItem?.bodyPreview || ""
    },
    parsedResumeData: options.parsedResumeData || null,
    resumes: Array.isArray(options.resumes) ? options.resumes : [],
    documents: Array.isArray(options.documents) ? options.documents : [],
    emailAddinRecord: options.emailAddinRecord || null,
    selectedResume: options.selectedResume || null,
    emailParserResult: options.emailParserResult || null,
    warnings: Array.isArray(options.warnings) ? options.warnings : []
  };
}

function buildAuthBridgeUrl(actionId, bridgeData = {}) {
  const redirectTo = buildTargetPath(actionId);
  const bridgeLoginData = buildBridgeLoginData();
  const loginExpirySeconds = Number(bridgeLoginData.expires_in || 0);
  const fallbackExpiration =
    loginExpirySeconds > 0 ? Date.now() + loginExpirySeconds * 1000 : Date.now() + 60 * 60 * 1000;
  const tokenExpiration = Number(state.auth?.tokenExpiration || 0) || fallbackExpiration;
  const payload = encodeBridgePayload({
    accessToken: state.auth?.accessToken || "",
    tokenType: state.auth?.tokenType || "Bearer",
    tokenExpiration,
    email: state.auth?.email || "",
    userId: state.auth?.userId || bridgeLoginData.userId || bridgeLoginData.UserId || "",
    mId: state.auth?.mId || bridgeLoginData.MId || "",
    loginData: bridgeLoginData,
    outlookCandidateImport: bridgeData.outlookCandidateImport || null,
    outlookImportSessionId: bridgeData.outlookImportSessionId || "",
    outlookImportApiHost: bridgeData.outlookImportApiHost || "",
    emailAddinRecord: bridgeData.emailAddinRecord || null,
    selectedResume: bridgeData.selectedResume || null
  });
  const hash = new URLSearchParams({
    payload,
    redirectTo
  });

  return `${buildAbsoluteAppUrl(state.config.authBridgePath)}#${hash.toString()}`;
}

function createPendingImportLaunch(actionId) {
  const sessionId =
    typeof window.crypto?.randomUUID === "function"
      ? window.crypto.randomUUID()
      : `outlook-import-${Date.now()}-${Math.random().toString(36).slice(2, 8)}`;

  launchAction(actionId, {
    outlookImportSessionId: sessionId,
    outlookImportApiHost: window.location.origin
  });

  return {
    sessionId,
    apiHost: window.location.origin
  };
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
    const url = buildAuthBridgeUrl(actionId, {
      outlookCandidateImport: options.outlookCandidateImport,
      outlookImportSessionId: options.outlookImportSessionId,
      outlookImportApiHost: options.outlookImportApiHost,
      emailAddinRecord: options.emailAddinRecord,
      selectedResume: options.selectedResume
    });
    safeOpenWindow(url);
    state.launchMessage = "";
    state.showLoginModal = false;
    closeAttachEmailPanel();
    closeImportModal();
    render();
  } catch (error) {
    state.launchMessage = "";
    state.loginError = error instanceof Error ? error.message : "Unable to open TrackTalents.";
    render();
  }
}

async function handleDirectActionLaunch(actionId) {
  if (!state.auth?.accessToken) {
    state.loginError = "Your TrackTalents session expired. Please sign in again.";
    render();
    return;
  }

  const pendingLaunch = createPendingImportLaunch(actionId);

  try {
    state.launchMessage = "";
    render();

    const emailParserTask = isEmailParserAction(actionId)
      ? parseCurrentEmailForAction(actionId)
          .then((parserResult) => {
            logEmailAddinRecordDebug("Email parsed", {
              actionId,
              status: parserResult?.status || "",
              confidence: parserResult?.confidence_summary || "",
              missingRequiredFields: parserResult?.missing_required_fields || [],
              step: "parse-email"
            });

            return buildEmailParserImportOptions(actionId, parserResult);
          })
          .catch((error) => {
            const parserWarning =
              error instanceof Error ? error.message : "Unable to parse the current Outlook email.";
            logEmailAddinRecordDebug("Email parse warning", {
              actionId,
              message: parserWarning,
              step: "parse-email"
            });

            return buildEmailParserImportOptions(actionId, null, parserWarning);
          })
      : Promise.resolve({});
    const importedDocumentsTask = prepareEmailAddinDocuments(getImportableAttachments());
    const [emailParserImportOptions, importedDocuments] = await Promise.all([
      emailParserTask,
      importedDocumentsTask
    ]);

    logEmailAddinRecordDebug("Submit start", {
      actionId,
      itemId: state.currentItem?.itemId || "",
      subject: state.currentItem?.subject || "",
      step: "create-email-addin-record"
    });

    logEmailAddinRecordDebug("Documents prepared", {
      actionId,
      count: importedDocuments.length,
      names: importedDocuments.map((document) => document?.DocumentName || "").filter(Boolean),
      step: "create-email-addin-record"
    });

    const payload = await createEmailAddinRecord(getEmailAddinRecordType(actionId), {
      body: state.currentItem?.bodyPreview || "",
      bodyHtml: state.currentItem?.bodyHtml || "",
      subject: state.currentItem?.subject || "",
      fromName: state.currentItem?.from?.displayName || "",
      fromEmail: state.currentItem?.from?.email || "",
      toRecipients: normalizeRecipientList(state.currentItem?.to),
      resumes: null,
      documents: importedDocuments
    });

    logEmailAddinRecordDebug("Email add-in record created", {
      actionId,
      recordId: String(payload?._id || ""),
      step: "load-email-addin-record"
    });

    const recordId = String(payload?._id || "");
    const outlookActionImport = buildOutlookActionImportPayload(actionId, {
      ...emailParserImportOptions,
      emailAddinRecord: {
        ...payload,
        _id: recordId || payload?._id || "",
        Type: String(payload?.Type || getEmailAddinRecordType(actionId))
      }
    });

    const importSession = await updateOutlookImportSession(
      pendingLaunch.sessionId,
      outlookActionImport
    );

    logEmailAddinRecordDebug("Outlook import session created", {
      actionId,
      sessionId: importSession.sessionId,
      step: "launch-action"
    });

    state.launchMessage = "";
    state.showLoginModal = false;
    closeImportModal();
    render();
  } catch (error) {
    logEmailAddinRecordDebug("Submit error", {
      actionId,
      message: error instanceof Error ? error.message : String(error)
    });
    state.launchMessage = "";
    state.loginError =
      error instanceof Error ? error.message : "Unable to open the TrackTalents action.";
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

async function prepareEmailAddinDocuments(attachments) {
  const documentAttachments = Array.isArray(attachments) ? attachments.filter(Boolean) : [];
  if (documentAttachments.length === 0) {
    return [];
  }

  const resolvedAttachments = await Promise.all(
    documentAttachments.map((attachment) => resolveAttachmentForImport(attachment))
  );

  return prepareImportDocuments(resolvedAttachments);
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
  const actionId = state.importModal.actionId || "add-candidate";
  const summary = getImportSummary();
  const selectedResume = summary.selectedResume;
  const documentAttachments = state.importModal.attachments.filter(
    (attachment) => attachment.id !== state.importModal.resumeAttachmentId
  );

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

  const pendingLaunch = createPendingImportLaunch(actionId);

  try {
    logEmailAddinRecordDebug("Submit start", {
      actionId,
      itemId: state.currentItem?.itemId || "",
      subject: state.currentItem?.subject || "",
      selectedResume: buildResumeBridgePreview(selectedResume),
      step: "resolve-attachment"
    });

    const selectedResumeTask = resolveAttachmentForImport(selectedResume);
    const resolvedDocumentAttachmentsTask = Promise.all(
      documentAttachments.map((attachment) => resolveAttachmentForImport(attachment))
    );
    const resolvedSelectedResume = await selectedResumeTask;
    logEmailAddinRecordDebug("Attachment resolved", {
      fileName: resolvedSelectedResume?.name || "",
      contentType: resolvedSelectedResume?.contentType || "",
      size: Number(resolvedSelectedResume?.size || 0),
      step: "parse-resume"
    });

    const [parsedResumeData, resolvedDocumentAttachments] = await Promise.all([
      parseResumeAttachment(resolvedSelectedResume),
      resolvedDocumentAttachmentsTask
    ]);
    logEmailAddinRecordDebug("Resume parsed", {
      fileName: resolvedSelectedResume?.name || "",
      hasParsedData: Boolean(parsedResumeData),
      parsedFirstName: String(parsedResumeData?.FirstName || ""),
      parsedLastName: String(parsedResumeData?.LastName || ""),
      step: "create-email-addin-record"
    });

    const importedDocuments = await prepareImportDocuments(resolvedDocumentAttachments);
    logEmailAddinRecordDebug("Documents prepared", {
      count: importedDocuments.length,
      names: importedDocuments.map((document) => document?.DocumentName || "").filter(Boolean),
      step: "create-email-addin-record"
    });

    const payload = await createEmailAddinRecord(getEmailAddinRecordType(actionId), {
      body: state.currentItem?.bodyPreview || "",
      bodyHtml: state.currentItem?.bodyHtml || "",
      subject: state.currentItem?.subject || "",
      fromName: state.currentItem?.from?.displayName || "",
      fromEmail: state.currentItem?.from?.email || "",
      toRecipients: normalizeRecipientList(state.currentItem?.to),
      resumes: null,
      documents: importedDocuments
    });
    logEmailAddinRecordDebug("Email add-in record created", {
      recordId: String(payload?._id || ""),
      step: "load-email-addin-record"
    });

    const recordId = String(payload?._id || "");
    logEmailAddinRecordDebug("Email add-in record loaded", {
      recordId,
      step: "create-outlook-import-session"
    });
    const resumePreview = buildResumeBridgePreview(resolvedSelectedResume);
    const outlookCandidateImport = buildOutlookActionImportPayload(actionId, {
      parsedResumeData: parsedResumeData || null,
      resumes: Array.isArray(parsedResumeData?.Resumes) ? parsedResumeData.Resumes : [],
      documents: importedDocuments,
      emailAddinRecord: {
        ...payload,
        _id: recordId || payload?._id || "",
        Type: String(payload?.Type || getEmailAddinRecordType(actionId))
      },
      selectedResume: resumePreview
    });
    const importSession = await updateOutlookImportSession(
      pendingLaunch.sessionId,
      outlookCandidateImport
    );
    logEmailAddinRecordDebug("Outlook import session created", {
      sessionId: importSession.sessionId,
      step: "launch-action"
    });

    state.launchMessage = "";
    state.showLoginModal = false;
    closeImportModal();
    render();
  } catch (error) {
    logEmailAddinRecordDebug("Submit error", {
      message: error instanceof Error ? error.message : String(error)
    });
    state.importModal.submitting = false;
    state.importModal.error =
      error instanceof Error ? error.message : "Unable to save the email add-in record.";
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
  const toRecipients = normalizeRecipientList(item.to);

  state.currentItem = {
    itemId: toPlainString(item.itemId) || toPlainString(item.internetMessageId),
    subject: toPlainString(item.subject),
    from: item.from
      ? {
          displayName: toPlainString(item.from.displayName),
          email: toPlainString(item.from.emailAddress)
        }
      : { displayName: "", email: "" },
    to: toRecipients,
    fromDisplay,
    toCount: toRecipients.length,
    attachments,
    attachmentCount: attachmentNames.length,
    attachmentNames,
    hasResumeAttachment: resumeNames.length > 0,
    primaryResumeName: resumeNames[0] || "",
    bodyPreview: "",
    bodyHtml: "",
    mode: "outlook"
  };

  if (item.body?.getAsync) {
    item.body.getAsync(Office.CoercionType.Text, (result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded && state.currentItem) {
        state.currentItem.bodyPreview = String(result.value || "")
          .replace(/\r\n/g, "\n")
          .replace(/\r/g, "\n")
          .trim();
        render();
      }
    });

    item.body.getAsync(Office.CoercionType.Html, (result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded && state.currentItem) {
        state.currentItem.bodyHtml = String(result.value || "").trim();
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
      to: [],
      fromDisplay: "No message selected",
      toCount: 0,
      attachments: [],
      attachmentCount: 0,
      attachmentNames: [],
      hasResumeAttachment: false,
      primaryResumeName: "",
      bodyPreview: "",
      bodyHtml: "",
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
