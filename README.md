# 🚀 Zoom X Microsoft Integration

## 🌐 Architecture Diagram

![Architecture Diagram](architecture_diagram.png)

---

## 🔐 Authentication Flow

Zoom App → Backend → Microsoft OAuth → Access Token → Microsoft Graph → Zoom UI

---

## 📧 Outlook Integration

- Connect MS Account
- Fetch last 3 emails

---

## 📄 Word Integration

- Connect MS Account
- Fetch last 3 Word documents
- Show metadata of latest document

---

## 🧪 Demo Walkthrough

1. Start backend:
   node server.js

2. Start ngrok:
   ngrok http 3000

3. Open Zoom App

4. Click **Connect MS Account**

5. Login via Microsoft

6. Click:
   - Show Emails
   - Show Word Documents

---

## ⚠️ Notes

- Word documents cannot open inside Zoom
- Tokens stored in-memory
- Restart server → reconnect required
