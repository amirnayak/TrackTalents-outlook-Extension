const http = require("http");
const https = require("https");
const path = require("path");

const express = require("express");
const devCerts = require("office-addin-dev-certs");

const HTTPS_PORT = Number(process.env.PORT || 3201);
const HTTP_PREVIEW_PORT = Number(process.env.PREVIEW_PORT || 3202);
const HOST = "localhost";
const API_HOST = process.env.API_HOST || "https://testapi.tracktalents.com/api/";
const APP_HOST = process.env.APP_HOST || "http://localhost:3000";
const app = express();

app.use((req, res, next) => {
  res.setHeader("Access-Control-Allow-Origin", "*");
  res.setHeader("Cache-Control", "no-store");
  next();
});

app.use(express.json({ limit: "25mb" }));
app.use(express.static(path.join(__dirname, "..", "public")));

app.get("/health", (_req, res) => {
  res.json({
    ok: true,
    host: HOST,
    httpsPort: HTTPS_PORT,
    httpPreviewPort: HTTP_PREVIEW_PORT,
    apiHost: API_HOST,
    appHost: APP_HOST
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

app.post("/api/candidate/import-from-email", async (req, res) => {
  const accessToken = typeof req.body?.accessToken === "string" ? req.body.accessToken.trim() : "";
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
    const parsedResumeData = await parseResumeWithTrackTalents(selectedResume, accessToken);
    const uploadedDocuments = [];

    for (const attachment of attachments) {
      if (attachment?.id === selectedResumeId) {
        continue;
      }

      const documentGuid = await uploadAttachmentToTrackTalents(attachment, accessToken);
      uploadedDocuments.push(buildImportedDocument(attachment, documentGuid, userId));
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
      parsedResumeData: {
        ...parsedResumeData,
        Resumes: [importedResume]
      },
      resumes: [importedResume],
      documents: uploadedDocuments
    });
  } catch (error) {
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
    bodyPreview: typeof emailContext?.bodyPreview === "string" ? emailContext.bodyPreview : ""
  };
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
    body: buildAttachmentFormData(attachment)
  });

  const data = await readApiResponse(response);
  if (!response.ok) {
    throw new Error(
      data?.message ||
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
  const response = await fetch(new URL("resume/parse", API_HOST), {
    method: "POST",
    headers: {
      Authorization: `Bearer ${accessToken}`
    },
    body: buildAttachmentFormData(attachment)
  });

  const data = await readApiResponse(response);
  if (!response.ok) {
    throw new Error(
      data?.message ||
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
