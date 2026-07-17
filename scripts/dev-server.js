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

app.use(express.json());
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

    const text = await response.text();
    let data;

    try {
      data = text ? JSON.parse(text) : {};
    } catch {
      data = { raw: text };
    }

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
