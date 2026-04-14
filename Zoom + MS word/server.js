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

const SCOPES = ["openid", "profile", "email", "offline_access", "Files.Read"];

const wordConnections = new Map();

function getOrCreateConnection(linkId) {
  if (!wordConnections.has(linkId)) {
    wordConnections.set(linkId, {
      connected: false,
      accessToken: null,
      account: null,
      createdAt: new Date().toISOString()
    });
  }
  return wordConnections.get(linkId);
}

function isWordDocument(item) {
  const name = (item.name || "").toLowerCase();
  const mime = item.file?.mimeType || "";

  return (
    name.endsWith(".docx") ||
    name.endsWith(".doc") ||
    mime === "application/vnd.openxmlformats-officedocument.wordprocessingml.document" ||
    mime === "application/msword"
  );
}

function sortByLastModifiedDesc(items) {
  return [...items].sort((a, b) => {
    const aTime = new Date(a.lastModifiedDateTime || 0).getTime();
    const bTime = new Date(b.lastModifiedDateTime || 0).getTime();
    return bTime - aTime;
  });
}

function getAccessToken(req, linkId) {
  if (linkId && wordConnections.has(linkId)) {
    const token = wordConnections.get(linkId).accessToken;
    if (token) return token;
  }

  if (req.session.accessToken) {
    return req.session.accessToken;
  }

  return null;
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
    res.status(500).send("Failed to start Microsoft auth");
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
        <title>Microsoft Connected</title>
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
        <h2>Microsoft account connected successfully</h2>
        <p>Signed in as: ${response.account?.username || "Unknown user"}</p>
        <p>Connection ID: <code>${linkId}</code></p>
        <p>You can now return to the Zoom app and use the Word buttons.</p>
      </body>
      </html>
    `);
  } catch (err) {
    console.error("Callback error:", err);
    res.status(500).send(`
      <h2>Microsoft callback failed</h2>
      <pre>${err.message}</pre>
    `);
  }
});

app.get("/auth/outlook/status", (req, res) => {
  const linkId = req.query.linkId;

  if (!linkId) {
    return res.status(400).json({ error: "Missing linkId" });
  }

  const record = wordConnections.get(linkId);

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

app.get("/word/docs", async (req, res) => {
  try {
    const linkId = req.query.linkId;
    const accessToken = getAccessToken(req, linkId);

    if (!accessToken) {
      return res.status(401).json({
        error: "Not connected",
        message: "Connect Microsoft first"
      });
    }

    const graphResponse = await fetch(
      "https://graph.microsoft.com/v1.0/me/drive/root/children?$top=50&$select=id,name,webUrl,lastModifiedDateTime,size,file,parentReference,createdDateTime",
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

    const items = data.value || [];
    const wordDocs = sortByLastModifiedDesc(items.filter(isWordDocument))
      .slice(0, 3)
      .map((item) => ({
        id: item.id,
        name: item.name || "",
        webUrl: item.webUrl || "",
        lastModifiedDateTime: item.lastModifiedDateTime || "",
        createdDateTime: item.createdDateTime || "",
        size: item.size || 0,
        parentPath: item.parentReference?.path || "",
        mimeType: item.file?.mimeType || ""
      }));

    res.json({ documents: wordDocs });
  } catch (err) {
    console.error("Word docs fetch error:", err);
    res.status(500).json({
      error: "Failed to fetch Word documents",
      message: err.message
    });
  }
});

app.get("/word/latest-metadata", async (req, res) => {
  try {
    const linkId = req.query.linkId;
    const accessToken = getAccessToken(req, linkId);

    if (!accessToken) {
      return res.status(401).json({
        error: "Not connected",
        message: "Connect Microsoft first"
      });
    }

    const graphResponse = await fetch(
      "https://graph.microsoft.com/v1.0/me/drive/root/children?$top=50&$select=id,name,webUrl,lastModifiedDateTime,size,file,parentReference,createdDateTime",
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

    const items = data.value || [];
    const wordDocs = sortByLastModifiedDesc(items.filter(isWordDocument));

    if (!wordDocs.length) {
      return res.json({ document: null });
    }

    const latest = wordDocs[0];

    res.json({
      document: {
        id: latest.id,
        name: latest.name || "",
        webUrl: latest.webUrl || "",
        lastModifiedDateTime: latest.lastModifiedDateTime || "",
        createdDateTime: latest.createdDateTime || "",
        size: latest.size || 0,
        parentPath: latest.parentReference?.path || "",
        mimeType: latest.file?.mimeType || ""
      }
    });
  } catch (err) {
    console.error("Latest metadata fetch error:", err);
    res.status(500).json({
      error: "Failed to fetch latest document metadata",
      message: err.message
    });
  }
});

app.get("/auth/zoom/callback", (req, res) => {
  res.send(`
    <h2>Zoom authorization successful</h2>
    <p>Your Zoom app is installed for testing.</p>
    <p><a href="/">Open Zoom X Word app</a></p>
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