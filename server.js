require("dotenv").config();

const express = require("express");
const session = require("express-session");
const { ConfidentialClientApplication } = require("@azure/msal-node");
const path = require("path");
const crypto = require("crypto");

const app = express();
const PORT = process.env.PORT || 3000;

app.set("trust proxy", 1);

app.use((req, res, next) => {
  res.setHeader(
    "Strict-Transport-Security",
    "max-age=31536000; includeSubDomains"
  );
  res.setHeader("X-Content-Type-Options", "nosniff");
  res.setHeader("Referrer-Policy", "strict-origin-when-cross-origin");
  res.setHeader(
    "Content-Security-Policy",
    "default-src 'self'; " +
      "connect-src 'self' https://graph.microsoft.com https://login.microsoftonline.com https://provider-vivacious-greyhound.ngrok-free.dev https://zoom.us https://*.zoom.us https://zoom.com https://*.zoom.com; " +
      "img-src 'self' data: https:; " +
      "style-src 'self' 'unsafe-inline'; " +
      "script-src 'self' 'unsafe-inline'; " +
      "frame-ancestors https://*.zoom.us https://*.zoom.com;"
  );
  next();
});

app.use(
  session({
    secret: process.env.SESSION_SECRET,
    resave: false,
    saveUninitialized: false,
    cookie: {
      httpOnly: true,
      secure: false,
      sameSite: "lax"
    }
  })
);

app.use(express.json());
app.use(express.static(path.join(__dirname, "public")));

const msalConfig = {
  auth: {
    clientId: process.env.MS_CLIENT_ID,
    authority: `https://login.microsoftonline.com/${process.env.MS_TENANT_ID}`,
    clientSecret: process.env.MS_CLIENT_SECRET
  }
};

const cca = new ConfidentialClientApplication(msalConfig);

const SCOPES = ["openid", "profile", "email", "offline_access", "Mail.Read"];

// In-memory store for local testing
const outlookConnections = new Map();

function getOrCreateConnection(linkId) {
  if (!outlookConnections.has(linkId)) {
    outlookConnections.set(linkId, {
      connected: false,
      accessToken: null,
      account: null,
      createdAt: new Date().toISOString()
    });
  }
  return outlookConnections.get(linkId);
}

app.get("/", (req, res) => {
  res.sendFile(path.join(__dirname, "public", "index.html"));
});

app.get("/auth/outlook/start", async (req, res) => {
  try {
    const linkId = req.query.linkId || crypto.randomUUID();

    getOrCreateConnection(linkId);

    const authCodeUrlParameters = {
      scopes: SCOPES,
      redirectUri: process.env.MS_REDIRECT_URI,
      state: linkId
    };

    const authUrl = await cca.getAuthCodeUrl(authCodeUrlParameters);
    res.redirect(authUrl);
  } catch (err) {
    console.error("Error generating auth URL:", err);
    res.status(500).send("Failed to start Outlook auth");
  }
});

app.get("/auth/outlook/callback", async (req, res) => {
  try {
    if (!req.query.code) {
      return res.status(400).send("Missing authorization code");
    }

    const linkId = req.query.state;
    if (!linkId) {
      return res.status(400).send("Missing state/linkId");
    }

    const tokenRequest = {
      code: req.query.code,
      scopes: SCOPES,
      redirectUri: process.env.MS_REDIRECT_URI
    };

    const response = await cca.acquireTokenByCode(tokenRequest);

    const record = getOrCreateConnection(linkId);
    record.connected = true;
    record.accessToken = response.accessToken;
    record.account = response.account;

    req.session.accessToken = response.accessToken;
    req.session.account = response.account;

    res.send(`
      <!doctype html>
      <html>
      <head>
        <meta charset="UTF-8" />
        <title>Outlook Connected</title>
        <style>
          body {
            font-family: Arial, sans-serif;
            max-width: 700px;
            margin: 40px auto;
            padding: 0 16px;
          }
          code {
            background: #f3f3f3;
            padding: 2px 6px;
          }
        </style>
      </head>
      <body>
        <h2>Outlook connected successfully</h2>
        <p>Signed in as: ${response.account?.username || "Unknown user"}</p>
        <p>Connection ID: <code>${linkId}</code></p>
        <p>You can now return to the Zoom app and click <strong>Fetch last 3 emails</strong>.</p>
      </body>
      </html>
    `);
  } catch (err) {
    console.error("Callback error:", err);
    res.status(500).send(`
      <h2>Outlook callback failed</h2>
      <pre>${err.message}</pre>
    `);
  }
});

app.get("/auth/outlook/status", (req, res) => {
  const linkId = req.query.linkId;

  if (!linkId) {
    return res.status(400).json({ error: "Missing linkId" });
  }

  const record = outlookConnections.get(linkId);

  if (!record) {
    return res.json({
      connected: false,
      account: null
    });
  }

  res.json({
    connected: !!record.connected,
    account: record.account?.username || null
  });
});

app.get("/emails/latest", async (req, res) => {
  try {
    const linkId = req.query.linkId;
    let accessToken = null;

    if (linkId && outlookConnections.has(linkId)) {
      accessToken = outlookConnections.get(linkId).accessToken;
    }

    if (!accessToken && req.session.accessToken) {
      accessToken = req.session.accessToken;
    }

    if (!accessToken) {
      return res.status(401).json({
        error: "Not connected",
        message: "Connect Outlook first"
      });
    }

    const graphResponse = await fetch(
      "https://graph.microsoft.com/v1.0/me/messages?$top=3&$orderby=receivedDateTime desc&$select=subject,from,receivedDateTime,bodyPreview",
      {
        headers: {
          Authorization: `Bearer ${accessToken}`
        }
      }
    );

    const data = await graphResponse.json();

    if (!graphResponse.ok) {
      console.error("Graph error:", data);
      return res.status(graphResponse.status).json(data);
    }

    const emails = (data.value || []).map((mail) => ({
      subject: mail.subject || "(no subject)",
      fromName: mail.from?.emailAddress?.name || "",
      fromAddress: mail.from?.emailAddress?.address || "",
      receivedDateTime: mail.receivedDateTime || "",
      bodyPreview: mail.bodyPreview || ""
    }));

    res.json({ emails });
  } catch (err) {
    console.error("Email fetch error:", err);
    res.status(500).json({
      error: "Failed to fetch emails",
      message: err.message
    });
  }
});

app.get("/auth/zoom/callback", (req, res) => {
  res.send(`
    <h2>Zoom authorization successful</h2>
    <p>Your Zoom app is installed for testing.</p>
    <p><a href="/">Open Zoom X Outlook app</a></p>
    <pre>${JSON.stringify(req.query, null, 2)}</pre>
  `);
});

app.get("/debug/sdk", (req, res) => {
  res.json({
    ok: true,
    message: "Backend reachable from app",
    timestamp: new Date().toISOString()
  });
});

app.listen(PORT, () => {
  console.log(`Server running on http://localhost:${PORT}`);
});