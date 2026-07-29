const http = require("http");
const https = require("https");
const path = require("path");
const { randomUUID } = require("crypto");

const express = require("express");
const devCerts = require("office-addin-dev-certs");

const HTTPS_PORT = Number(process.env.PORT || 3201);
const HTTP_PREVIEW_PORT = Number(process.env.PREVIEW_PORT || 3202);
const HOST = "localhost";
const API_HOST = process.env.API_HOST || "https://testapi.tracktalents.com/api/";
const APP_HOST = process.env.APP_HOST || "http://localhost:3000";
const REQUEST_BODY_LIMIT = process.env.REQUEST_BODY_LIMIT || "100mb";
const OUTLOOK_IMPORT_SESSION_TTL_MS = 15 * 60 * 1000;
const API_REQUEST_TIMEOUT_MS = Number(process.env.API_REQUEST_TIMEOUT_MS || 30000);
const app = express();
const outlookImportSessions = new Map();

app.use((req, res, next) => {
  res.setHeader("Access-Control-Allow-Origin", "*");
  res.setHeader("Access-Control-Allow-Headers", "Content-Type, Authorization");
  res.setHeader("Access-Control-Allow-Methods", "GET,POST,OPTIONS");
  res.setHeader("Cache-Control", "no-store");

  if (req.method === "OPTIONS") {
    res.status(204).end();
    return;
  }

  next();
});

app.use(express.json({ limit: REQUEST_BODY_LIMIT }));
app.use(express.static(path.join(__dirname, "..", "public")));

app.get("/health", (_req, res) => {
  res.json({
    ok: true,
    host: HOST,
    httpsPort: HTTPS_PORT,
    httpPreviewPort: HTTP_PREVIEW_PORT,
    apiHost: API_HOST,
    appHost: APP_HOST,
    requestBodyLimit: REQUEST_BODY_LIMIT
  });
});

app.get("/api/config", (_req, res) => {
  res.json({
    appHost: APP_HOST,
    authBridgePath: "/outlook-auth-bridge",
    loginPath: "/login",
    forgotPasswordPath: "/forgotpassword"
  });
});

app.use((error, _req, res, next) => {
  if (error?.type === "entity.too.large") {
    res.status(413).json({
      message: `The selected Outlook attachments are too large for the local import bridge (${REQUEST_BODY_LIMIT} limit).`
    });
    return;
  }

  next(error);
});

app.post("/api/auth/login", async (req, res) => {
  const email = typeof req.body?.email === "string" ? req.body.email.trim() : "";
  const password = typeof req.body?.password === "string" ? req.body.password : "";

  if (!email || !password) {
    res.status(400).json({ message: "Email and password are required." });
    return;
  }

  try {
    const body = new URLSearchParams({
      username: email,
      password,
      grant_type: "password",
      id: ""
    });

    const response = await fetch(new URL("Token", API_HOST), {
      method: "POST",
      headers: {
        "Content-Type": "application/x-www-form-urlencoded"
      },
      body
    });

    const data = await readApiResponse(response);

    if (!response.ok) {
      res.status(response.status).json({
        message:
          data?.error_description ||
          data?.message ||
          "Login failed. Please verify your credentials.",
        details: data
      });
      return;
    }

    res.json(data);
  } catch (error) {
    res.status(502).json({
      message: "Unable to reach the TrackTalents API.",
      details: error instanceof Error ? error.message : String(error)
    });
  }
});

app.post("/api/EmailAddinRecord", async (req, res) => {
  const accessToken = readAccessToken(req);
  const type = typeof req.body?.type === "string" ? req.body.type.trim() : "";
  const emailData = req.body?.emailData && typeof req.body.emailData === "object"
    ? req.body.emailData
    : {};
  const resumes = Array.isArray(req.body?.resumes) ? req.body.resumes : null;
  const documents = Array.isArray(req.body?.documents) ? req.body.documents : null;

  if (!accessToken) {
    res.status(401).json({ message: "TrackTalents access token is required." });
    return;
  }

  if (!type) {
    res.status(400).json({ message: "Email add-in record type is required." });
    return;
  }

  try {
    const response = await fetch(new URL("EmailAddinRecord", API_HOST), {
      method: "POST",
      headers: {
        Authorization: `Bearer ${accessToken}`,
        "Content-Type": "application/json"
      },
      body: JSON.stringify(buildEmailAddinRecordPayload(emailData, type, resumes, documents))
    });

    const data = await readApiResponse(response);
    if (!response.ok) {
      res.status(response.status).json({
        message: data?.message || "Unable to create the email add-in record.",
        details: data
      });
      return;
    }

    res.json(data);
  } catch (error) {
    res.status(502).json({
      message: "Unable to reach the TrackTalents API.",
      details: error instanceof Error ? error.message : String(error)
    });
  }
});

app.get("/api/EmailAddinRecord/:recordId", async (req, res) => {
  const accessToken = readAccessToken(req);
  const recordId = typeof req.params?.recordId === "string" ? req.params.recordId.trim() : "";

  if (!accessToken) {
    res.status(401).json({ message: "TrackTalents access token is required." });
    return;
  }

  if (!recordId) {
    res.status(400).json({ message: "Email add-in record id is required." });
    return;
  }

  try {
    const response = await fetch(new URL(`EmailAddinRecord/${encodeURIComponent(recordId)}`, API_HOST), {
      method: "GET",
      headers: {
        Authorization: `Bearer ${accessToken}`
      }
    });

    const data = await readApiResponse(response);
    if (!response.ok) {
      res.status(response.status).json({
        message: data?.message || "Unable to load the email add-in record.",
        details: data
      });
      return;
    }

    res.json(data);
  } catch (error) {
    res.status(502).json({
      message: "Unable to reach the TrackTalents API.",
      details: error instanceof Error ? error.message : String(error)
    });
  }
});

app.get("/api/attach-email/contacts", async (req, res) => {
  const accessToken = readAccessToken(req);

  if (!accessToken) {
    res.status(401).json({ message: "TrackTalents access token is required." });
    return;
  }

  try {
    const payload = await fetchAttachEmailGrid("Contacts/Perf", accessToken, req.query, {
      sort: '[{"selector":"ContactData.ModifiedDate","desc":true}]'
    });
    res.json(payload);
  } catch (error) {
    res.status(502).json({
      message:
        error instanceof Error ? error.message : "Unable to load TrackTalents contacts.",
      details: error instanceof Error ? error.stack : String(error)
    });
  }
});

app.get("/api/attach-email/candidates", async (req, res) => {
  const accessToken = readAccessToken(req);

  if (!accessToken) {
    res.status(401).json({ message: "TrackTalents access token is required." });
    return;
  }

  try {
    const payload = await fetchAttachEmailCandidates(accessToken, req.query);
    res.json(payload);
  } catch (error) {
    res.status(502).json({
      message: error instanceof Error ? error.message : "Unable to load TrackTalents candidates.",
      details: error instanceof Error ? error.stack : String(error)
    });
  }
});

app.post("/api/attach-email/link", async (req, res) => {
  const accessToken = readAccessToken(req);
  const entityType = typeof req.body?.entityType === "string" ? req.body.entityType.trim() : "";
  const entityId = typeof req.body?.entityId === "string" ? req.body.entityId.trim() : "";
  const userId = typeof req.body?.userId === "string" ? req.body.userId.trim() : "";
  const noteDescription =
    typeof req.body?.noteDescription === "string" ? req.body.noteDescription.trim() : "";

  if (!accessToken) {
    res.status(401).json({ message: "TrackTalents access token is required." });
    return;
  }

  if (!entityType || !entityId) {
    res.status(400).json({ message: "A valid TrackTalents entity type and id are required." });
    return;
  }

  if (!noteDescription) {
    res.status(400).json({ message: "A note description is required for the email attachment." });
    return;
  }

  const normalizedEntityName = normalizeAttachEmailEntityName(entityType);
  if (!normalizedEntityName) {
    res.status(400).json({
      message: "Only Contacts and Candidates can be linked from Attach Email."
    });
    return;
  }

  try {
    const response = await fetch(new URL("Activities", API_HOST), {
      method: "POST",
      headers: {
        Authorization: `Bearer ${accessToken}`,
        "Content-Type": "application/json"
      },
      body: JSON.stringify({
        NoteType: "Email",
        NoteStatus: "Open",
        NoteDescription: noteDescription,
        EntityName: normalizedEntityName,
        EntityID: entityId,
        ...(userId ? { UserId: userId } : {})
      }),
      signal: AbortSignal.timeout(API_REQUEST_TIMEOUT_MS)
    });

    const data = await readApiResponse(response);
    if (!response.ok) {
      res.status(response.status).json({
        message: extractApiMessage(data) || "Unable to attach the email to the selected record.",
        details: data
      });
      return;
    }

    res.json({
      success: true,
      entityType: normalizedEntityName,
      entityId,
      data
    });
  } catch (error) {
    res.status(502).json({
      message:
        error instanceof Error ? error.message : "Unable to reach the TrackTalents API.",
      details: error instanceof Error ? error.stack : String(error)
    });
  }
});

app.post("/api/resume/parse", async (req, res) => {
  const accessToken = readAccessToken(req);
  const attachment =
    req.body?.attachment && typeof req.body.attachment === "object" ? req.body.attachment : null;

  if (!accessToken) {
    res.status(401).json({ message: "TrackTalents access token is required." });
    return;
  }

  if (!attachment) {
    res.status(400).json({ message: "A selected resume attachment is required." });
    return;
  }

  try {
    const parsedResumeData = await parseResumeWithTrackTalents(attachment, accessToken);
    res.json(parsedResumeData);
  } catch (error) {
    res.status(502).json({
      message:
        error instanceof Error ? error.message : "Unable to parse the selected Outlook resume.",
      details: error instanceof Error ? error.stack : String(error)
    });
  }
});

app.post("/api/outlook-import-documents", async (req, res) => {
  const accessToken = readAccessToken(req);
  const userId = typeof req.body?.userId === "string" ? req.body.userId.trim() : "";
  const attachments = Array.isArray(req.body?.attachments) ? req.body.attachments : [];

  if (!accessToken) {
    res.status(401).json({ message: "TrackTalents access token is required." });
    return;
  }

  if (attachments.length === 0) {
    res.json({ documents: [] });
    return;
  }

  try {
    const importedDocuments = [];

    for (const attachment of attachments) {
      const documentGuid = await uploadAttachmentToTrackTalents(attachment, accessToken);
      importedDocuments.push(buildImportedDocument(attachment, documentGuid, userId));
    }

    res.json({
      documents: importedDocuments
    });
  } catch (error) {
    res.status(502).json({
      message:
        error instanceof Error
          ? error.message
          : "Unable to prepare the selected Outlook documents.",
      details: error instanceof Error ? error.stack : String(error)
    });
  }
});

app.post("/api/outlook-import-session", (req, res) => {
  const payload = req.body?.payload;

  if (!payload || typeof payload !== "object") {
    res.status(400).json({ message: "A valid Outlook import payload is required." });
    return;
  }

  const sessionId = randomUUID();
  const expiresAt = Date.now() + OUTLOOK_IMPORT_SESSION_TTL_MS;

  outlookImportSessions.set(sessionId, {
    payload,
    expiresAt
  });

  res.json({
    sessionId,
    expiresAt
  });
});

app.post("/api/outlook-import-session/:sessionId", (req, res) => {
  const sessionId = typeof req.params?.sessionId === "string" ? req.params.sessionId.trim() : "";
  const payload = req.body?.payload;

  if (!sessionId) {
    res.status(400).json({ message: "A valid Outlook import session id is required." });
    return;
  }

  if (!payload || typeof payload !== "object") {
    res.status(400).json({ message: "A valid Outlook import payload is required." });
    return;
  }

  const expiresAt = Date.now() + OUTLOOK_IMPORT_SESSION_TTL_MS;
  outlookImportSessions.set(sessionId, {
    payload,
    expiresAt
  });

  res.json({
    sessionId,
    expiresAt
  });
});

app.get("/api/outlook-import-session/:sessionId", (req, res) => {
  const sessionId = typeof req.params?.sessionId === "string" ? req.params.sessionId.trim() : "";
  const session = outlookImportSessions.get(sessionId);

  if (!sessionId || !session) {
    res.status(404).json({ message: "The Outlook import session could not be found." });
    return;
  }

  if (Number(session.expiresAt || 0) <= Date.now()) {
    outlookImportSessions.delete(sessionId);
    res.status(410).json({ message: "The Outlook import session expired. Please import again." });
    return;
  }

  res.json(session.payload);
});

app.post("/api/candidate/import-from-email", async (req, res) => {
  const accessToken = readAccessToken(req);
  const userId = typeof req.body?.userId === "string" ? req.body.userId.trim() : "";
  const selectedResumeId =
    typeof req.body?.selectedResumeId === "string" ? req.body.selectedResumeId.trim() : "";
  const previewMode = Boolean(req.body?.previewMode);
  const attachments = Array.isArray(req.body?.attachments) ? req.body.attachments : [];
  const emailContext = req.body?.emailContext && typeof req.body.emailContext === "object"
    ? req.body.emailContext
    : {};

  if (!accessToken) {
    res.status(401).json({ message: "TrackTalents access token is required." });
    return;
  }

  if (!selectedResumeId) {
    res.status(400).json({ message: "Please choose one email attachment as the resume." });
    return;
  }

  if (attachments.length === 0) {
    res.status(400).json({ message: "No attachments were provided for import." });
    return;
  }

  const selectedResume = attachments.find((attachment) => attachment?.id === selectedResumeId);
  if (!selectedResume) {
    res.status(400).json({ message: "The selected resume attachment could not be found." });
    return;
  }

  try {
    if (previewMode) {
      res.json(buildPreviewCandidateImportPayload(selectedResume, attachments, emailContext, userId));
      return;
    }

    const uploadedResumeId = await uploadAttachmentToTrackTalents(selectedResume, accessToken);
    const uploadedDocuments = [];
    let parsedResumeData = null;
    let parseWarning = "";

    for (const attachment of attachments) {
      if (attachment?.id === selectedResumeId) {
        continue;
      }

      const documentGuid = await uploadAttachmentToTrackTalents(attachment, accessToken);
      uploadedDocuments.push(buildImportedDocument(attachment, documentGuid, userId));
    }

    try {
      parsedResumeData = await parseResumeWithTrackTalents(selectedResume, accessToken);
    } catch (error) {
      parseWarning =
        error instanceof Error
          ? error.message
          : "TrackTalents could not parse the selected resume.";
    }

    const resumeText = Array.isArray(parsedResumeData?.Resumes) && parsedResumeData.Resumes[0]
      ? String(parsedResumeData.Resumes[0].ResumeText || "")
      : "";
    const importedResume = {
      ResumeId: uploadedResumeId,
      ResumeTitle: selectedResume.name || "Imported Resume",
      ResumeName: selectedResume.name || "Imported Resume",
      ResumeText: resumeText,
      IsPrimary: true
    };

    res.json({
      source: "outlook-addin",
      importedAt: new Date().toISOString(),
      emailContext: sanitizeEmailContext(emailContext),
      selectedResumeName: selectedResume.name || "",
      parsedResumeData: parsedResumeData
        ? {
            ...parsedResumeData,
            Resumes: [importedResume]
          }
        : buildFallbackCandidateImportData(importedResume, emailContext),
      resumes: [importedResume],
      documents: uploadedDocuments,
      warnings: parseWarning ? [parseWarning] : []
    });
  } catch (error) {
    console.error("Candidate import from email failed:", error);
    res.status(502).json({
      message: error instanceof Error ? error.message : "Unable to import the selected attachments.",
      details: error instanceof Error ? error.stack : String(error)
    });
  }
});

function sanitizeEmailContext(emailContext) {
  return {
    subject: typeof emailContext?.subject === "string" ? emailContext.subject : "",
    fromName: typeof emailContext?.fromName === "string" ? emailContext.fromName : "",
    fromEmail: typeof emailContext?.fromEmail === "string" ? emailContext.fromEmail : "",
    toRecipients: sanitizeRecipientList(emailContext?.toRecipients),
    bodyPreview: typeof emailContext?.bodyPreview === "string" ? emailContext.bodyPreview : ""
  };
}

function readAccessToken(req) {
  const authHeader =
    typeof req.headers?.authorization === "string" ? req.headers.authorization.trim() : "";
  if (/^Bearer\s+/i.test(authHeader)) {
    return authHeader.replace(/^Bearer\s+/i, "").trim();
  }

  if (typeof req.body?.accessToken === "string") {
    return req.body.accessToken.trim();
  }

  if (typeof req.query?.accessToken === "string") {
    return req.query.accessToken.trim();
  }

  return "";
}

function normalizeAttachEmailEntityName(value) {
  const normalized = String(value || "").trim().toLowerCase();

  if (normalized === "contact" || normalized === "contacts") {
    return "Contacts";
  }

  if (normalized === "candidate" || normalized === "candidates") {
    return "Candidates";
  }

  if (normalized === "job" || normalized === "jobs") {
    return "Jobs";
  }

  return "";
}

function normalizeGridNumber(value, fallback, minimum = 0, maximum = 100) {
  const parsed = Number(value);
  if (!Number.isFinite(parsed)) {
    return fallback;
  }

  return Math.min(maximum, Math.max(minimum, Math.floor(parsed)));
}

async function fetchAttachEmailGrid(endpoint, accessToken, query, options = {}) {
  const params = new URLSearchParams({
    skip: String(normalizeGridNumber(query?.skip, 0, 0, 100000)),
    take: String(normalizeGridNumber(query?.take, 10, 1, 50)),
    requireTotalCount:
      String(query?.requireTotalCount || "true").toLowerCase() === "false" ? "false" : "true",
    sort: typeof options.sort === "string" && options.sort ? options.sort : "[]",
    _: String(Date.now())
  });

  if (typeof query?.filter === "string" && query.filter.trim()) {
    params.set("filter", query.filter.trim());
  }

  const response = await fetch(new URL(`${endpoint}?${params.toString()}`, API_HOST), {
    method: "GET",
    headers: {
      Authorization: `Bearer ${accessToken}`
    },
    signal: AbortSignal.timeout(API_REQUEST_TIMEOUT_MS)
  });

  const data = await readApiResponse(response);
  if (!response.ok) {
    throw new Error(
      extractApiMessage(data) || `TrackTalents could not load records from ${endpoint}.`
    );
  }

  return {
    data: Array.isArray(data?.data) ? data.data : [],
    totalCount: Number(data?.totalCount || 0)
  };
}

async function fetchAttachEmailGridWithFallback(endpoints, accessToken, query, options = {}) {
  const candidates = Array.isArray(endpoints) ? endpoints.filter(Boolean) : [];
  let lastError = null;

  for (const endpoint of candidates) {
    try {
      return await fetchAttachEmailGrid(endpoint, accessToken, query, options);
    } catch (error) {
      lastError = error;
    }
  }

  throw (
    lastError ||
    new Error("TrackTalents could not load the requested records from any supported endpoint.")
  );
}

async function fetchAttachEmailCandidates(accessToken, query) {
  const skip = normalizeGridNumber(query?.skip, 0, 0, 100000);
  const take = normalizeGridNumber(query?.take, 10, 1, 50);
  const page = Math.floor(skip / take) + 1;
  const search = extractAttachEmailSearchValue(query?.filter);

  const params = new URLSearchParams({
    pg: String(page),
    ps: String(take),
    SearchType: search ? "SearchAll" : "Default",
    SortColumn: "CandidateData.ModifiedDate",
    _: String(Date.now())
  });

  if (search) {
    params.set("CandidateData.FirstName", search);
    params.set("CandidateData.LastName", search);
    params.set("CandidateData.Contact.Email1", search);
    params.set("CandidateData.Contact.Email2", search);
    params.set("CandidateData.JobTitle", search);
    params.set("CandidateData.CurrentLocation", search);
    params.set("CandidateData.CandidateStatus", "");
  }

  const response = await fetch(new URL("v1/Candidates/Elastic", API_HOST), {
    method: "POST",
    headers: {
      Authorization: `Bearer ${accessToken}`,
      "Content-Type": "application/x-www-form-urlencoded;charset=UTF-8"
    },
    body: params.toString(),
    signal: AbortSignal.timeout(API_REQUEST_TIMEOUT_MS)
  });

  const data = await readApiResponse(response);
  if (!response.ok) {
    throw new Error(
      extractApiMessage(data) || "TrackTalents could not load records from v1/Candidates/Elastic."
    );
  }

  return {
    data: Array.isArray(data?.data) ? data.data : [],
    totalCount: Number(data?.totalCount || 0)
  };
}

function extractAttachEmailSearchValue(filterValue) {
  if (typeof filterValue !== "string" || !filterValue.trim()) {
    return "";
  }

  try {
    const parsed = JSON.parse(filterValue);
    const queue = [parsed];

    while (queue.length) {
      const current = queue.shift();
      if (!Array.isArray(current)) {
        continue;
      }

      if (
        current.length >= 3 &&
        typeof current[1] === "string" &&
        current[1].toLowerCase() === "contains" &&
        typeof current[2] === "string"
      ) {
        return current[2].trim();
      }

      current.forEach((entry) => {
        if (Array.isArray(entry)) {
          queue.push(entry);
        }
      });
    }
  } catch {
    return "";
  }

  return "";
}

function sanitizeEmailAddinData(emailData) {
  return {
    Body: typeof emailData?.Body === "string"
      ? emailData.Body
      : typeof emailData?.body === "string"
        ? emailData.body
        : "",
    Subject: typeof emailData?.Subject === "string"
      ? emailData.Subject
      : typeof emailData?.subject === "string"
        ? emailData.subject
        : "",
    From: sanitizeSingleRecipient(emailData?.From ?? emailData?.from),
    To: sanitizeRecipientList(emailData?.To ?? emailData?.to)
  };
}

function buildEmailAddinRecordPayload(emailData, type, resumes, documents) {
  return {
    EmailData: sanitizeEmailAddinData(emailData),
    Resumes: resumes,
    Documents: documents,
    Type: type
  };
}

function sanitizeSingleRecipient(recipient) {
  if (!recipient || typeof recipient !== "object") {
    return {
      Name: "",
      Email: ""
    };
  }

  return {
    Name: typeof recipient?.Name === "string"
      ? recipient.Name
      : typeof recipient?.name === "string"
        ? recipient.name
        : typeof recipient?.displayName === "string"
          ? recipient.displayName
          : "",
    Email: typeof recipient?.Email === "string"
      ? recipient.Email
      : typeof recipient?.email === "string"
        ? recipient.email
        : typeof recipient?.emailAddress === "string"
          ? recipient.emailAddress
          : ""
  };
}

function sanitizeRecipientList(recipients) {
  if (!Array.isArray(recipients)) {
    return [];
  }

  return recipients
    .map(sanitizeSingleRecipient)
    .filter((recipient) => recipient.Name || recipient.Email);
}

function buildImportedDocument(attachment, documentGuid, userId) {
  const safeUserId = String(userId || "");
  const owners = safeUserId
    ? [
        {
          value: safeUserId,
          label: safeUserId
        }
      ]
    : [];

  return {
    DocumentGuid: documentGuid,
    DocumentId: buildRandomId(),
    DocumentName: String(attachment?.name || "Imported Document"),
    DocumentDesc: String(attachment?.name || "Imported Document"),
    DocumentType: "Email Attachment",
    CreateDate: new Date().toISOString(),
    Owners: owners,
    Private: false,
    UserId: safeUserId
  };
}

function buildRandomId() {
  return Math.random().toString(36).slice(2, 8).toUpperCase();
}

function decodeAttachmentBuffer(attachment) {
  const format = String(attachment?.contentFormat || "base64").trim().toLowerCase();
  const content = typeof attachment?.content === "string" ? attachment.content : "";

  if (!content) {
    throw new Error(`"${attachment?.name || "Selected attachment"}" is empty and could not be imported.`);
  }

  if (format === "url") {
    throw new Error(
      `"${attachment?.name || "Selected attachment"}" is a cloud attachment URL and cannot be imported from Outlook yet.`
    );
  }

  if (format === "base64") {
    return Buffer.from(content, "base64");
  }

  return Buffer.from(content, "utf8");
}

function buildAttachmentFormData(attachment) {
  const buffer = decodeAttachmentBuffer(attachment);
  const mimeType =
    typeof attachment?.contentType === "string" && attachment.contentType
      ? attachment.contentType
      : "application/octet-stream";
  const blob = new Blob([buffer], { type: mimeType });
  const formData = new FormData();
  formData.append("file", blob, String(attachment?.name || "attachment.bin"));
  return formData;
}

async function uploadAttachmentToTrackTalents(attachment, accessToken) {
  const response = await fetch(new URL("Attachment/Save", API_HOST), {
    method: "POST",
    headers: {
      Authorization: `Bearer ${accessToken}`
    },
    body: buildAttachmentFormData(attachment),
    signal: AbortSignal.timeout(API_REQUEST_TIMEOUT_MS)
  });

  const data = await readApiResponse(response);
  if (!response.ok) {
    throw new Error(
      extractApiMessage(data) ||
        `TrackTalents could not upload "${attachment?.name || "the selected attachment"}".`
    );
  }

  const value = unwrapApiValue(data);
  if (!value) {
    throw new Error(
      `TrackTalents returned an empty attachment id for "${attachment?.name || "the selected attachment"}".`
    );
  }

  return String(value);
}

async function parseResumeWithTrackTalents(attachment, accessToken) {
  const startTime = Date.now();
  let response;

  try {
    response = await fetch(new URL("resume/parse", API_HOST), {
      method: "POST",
      headers: {
        Authorization: `Bearer ${accessToken}`
      },
      body: buildAttachmentFormData(attachment),
      signal: AbortSignal.timeout(API_REQUEST_TIMEOUT_MS)
    });
  } catch (error) {
    if (error?.name === "TimeoutError") {
      throw new Error(
        `TrackTalents did not finish parsing "${attachment?.name || "the selected resume"}" within ${
          Math.round(API_REQUEST_TIMEOUT_MS / 1000)
        } seconds.`
      );
    }

    throw error;
  }

  const data = await readApiResponse(response);
  console.log("Resume parse completed", {
    fileName: attachment?.name || "",
    durationMs: Date.now() - startTime,
    ok: response.ok,
    status: response.status
  });

  if (!response.ok) {
    throw new Error(
      extractApiMessage(data) ||
        `TrackTalents could not parse "${attachment?.name || "the selected resume"}".`
    );
  }

  return data;
}

async function readApiResponse(response) {
  const text = await response.text();
  if (!text) {
    return {};
  }

  try {
    return JSON.parse(text);
  } catch {
    return text;
  }
}

function unwrapApiValue(value) {
  if (!value) {
    return "";
  }

  if (typeof value === "string") {
    return value.replace(/^"+|"+$/g, "").trim();
  }

  if (typeof value === "object") {
    if (typeof value.value === "string") {
      return value.value;
    }

    if (typeof value.id === "string") {
      return value.id;
    }

    if (typeof value.guid === "string") {
      return value.guid;
    }
  }

  return "";
}

function extractApiMessage(data) {
  if (!data) {
    return "";
  }

  if (typeof data === "string") {
    return data.trim();
  }

  if (typeof data === "object") {
    if (typeof data.message === "string" && data.message.trim()) {
      return data.message.trim();
    }

    if (typeof data.error_description === "string" && data.error_description.trim()) {
      return data.error_description.trim();
    }

    if (typeof data.error === "string" && data.error.trim()) {
      return data.error.trim();
    }
  }

  return "";
}

function buildFallbackCandidateImportData(importedResume, emailContext) {
  const safeContext = sanitizeEmailContext(emailContext);
  const fullName = safeContext.fromName || "Candidate Imported";
  const nameParts = fullName.split(/\s+/).filter(Boolean);

  return {
    FirstName: nameParts[0] || "Candidate",
    LastName: nameParts.slice(1).join(" ") || "Imported",
    Contact: {
      CellNumber: "",
      WorkNumber: null,
      DirectNumber: null,
      Email1: safeContext.fromEmail || "",
      Email2: null
    },
    Resumes: [importedResume]
  };
}

function buildPreviewCandidateImportPayload(selectedResume, attachments, emailContext, userId) {
  const safeContext = sanitizeEmailContext(emailContext);
  const firstLine = String(selectedResume?.previewText || "")
    .split(/\r?\n/)
    .map((line) => line.trim())
    .filter(Boolean)[0] || "";
  const fullName = firstLine || safeContext.fromName || "Candidate";
  const nameParts = fullName.split(/\s+/).filter(Boolean);
  const firstName = nameParts[0] || "Candidate";
  const lastName = nameParts.slice(1).join(" ") || "Imported";
  const otherAttachments = attachments.filter((attachment) => attachment?.id !== selectedResume?.id);

  const importedResume = {
    ResumeId: "PREVIEW-RESUME",
    ResumeTitle: selectedResume?.name || "Preview Resume",
    ResumeName: selectedResume?.name || "Preview Resume",
    ResumeText: String(selectedResume?.previewText || safeContext.bodyPreview || ""),
    IsPrimary: true
  };

  return {
    source: "outlook-addin",
    importedAt: new Date().toISOString(),
    emailContext: safeContext,
    selectedResumeName: selectedResume?.name || "",
    parsedResumeData: {
      FirstName: firstName,
      LastName: lastName,
      Contact: {
        CellNumber: "+1 555-0102",
        WorkNumber: null,
        DirectNumber: null,
        Email1: safeContext.fromEmail || "candidate@example.com",
        Email2: null
      },
      CurrentLocation: "Dallas, TX",
      Relocation: "Open",
      JobTitle: "Senior Java Developer",
      TotalExperience: "7",
      EducationLevel: "Bachelor's",
      WorkAuthorization: "Authorized to work",
      skills: [
        { skill: "Java", Years: 7 },
        { skill: "Spring Boot", Years: 5 },
        { skill: "AWS", Years: 4 }
      ],
      WorkExperiences: [
        {
          Employer: "Nimbus Systems",
          JobTitle: "Senior Java Developer",
          JobLocation: "Dallas, TX",
          StartDate: "2020-01-01",
          EndDate: ""
        }
      ],
      EducationDetails: [
        {
          University: "University of Texas",
          Degree: "Computer Science",
          DegreeType: "Bachelor's",
          YearPassed: "2017"
        }
      ],
      Addresses: [
        {
          StreetAddress: "",
          City: "Dallas",
          State: "TX",
          PostalCode: "",
          Country: "USA"
        }
      ],
      Resumes: [importedResume]
    },
    resumes: [importedResume],
    documents: otherAttachments.map((attachment, index) =>
      buildImportedDocument(
        attachment,
        `PREVIEW-DOC-${index + 1}`,
        userId
      )
    )
  };
}

async function start() {
  const httpsOptions = await devCerts.getHttpsServerOptions();

  if (!httpsOptions || !httpsOptions.key || !httpsOptions.cert) {
    throw new Error("Unable to load localhost HTTPS certificates.");
  }

  const httpsServer = https.createServer(
    {
      key: httpsOptions.key,
      cert: httpsOptions.cert,
      ca: httpsOptions.ca
    },
    app
  );

  const httpServer = http.createServer(app);

  httpsServer.listen(HTTPS_PORT, HOST, () => {
    console.log(`TrackTalents Outlook app running at https://${HOST}:${HTTPS_PORT}`);
    console.log(`Health check: https://${HOST}:${HTTPS_PORT}/health`);
    console.log(`Preview in app browser: http://${HOST}:${HTTP_PREVIEW_PORT}/taskpane.html`);
    console.log("Sideload manifest: manifest/tracktalents-outlook-manifest.xml");
  });

  httpServer.listen(HTTP_PREVIEW_PORT, HOST, () => {
    console.log(`HTTP preview running at http://${HOST}:${HTTP_PREVIEW_PORT}`);
  });
}

start().catch((error) => {
  console.error("Failed to start the Outlook server.");
  console.error(error);
  process.exit(1);
});
