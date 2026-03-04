# 📅 Timetable Generator — Streamlit + Google Drive

Full-featured school timetable scheduler with Google Sign-In and Drive sync.

---

## 🚀 Deploy to Streamlit Cloud (3 steps)

**1.** Push all files to a GitHub repo (maintain the `.streamlit/` folder)

**2.** Go to **share.streamlit.io** → New app → select repo → main file: `app.py` → Deploy

**3.** Add Google OAuth secrets (see below) — then share the URL

---

## 🔐 Google OAuth + Drive Setup (10 mins, one-time)

### A. Create OAuth credentials

1. Go to **https://console.cloud.google.com**
2. Create a new project (e.g. "Timetable App")
3. **APIs & Services → OAuth consent screen** → External → fill app name, add your email
4. **APIs & Services → Credentials → Create Credentials → OAuth 2.0 Client ID**
   - Application type: **Web application**
   - Authorized redirect URIs: `https://YOUR-APP.streamlit.app` *(exact URL from step 2, no trailing slash)*
5. Copy the **Client ID** and **Client Secret**
6. **APIs & Services → Library → search "Google Drive API" → Enable**

### B. Add secrets to Streamlit Cloud

In your app → **Settings → Secrets**, paste:

```toml
[google_oauth]
client_id     = "123456-abc.apps.googleusercontent.com"
client_secret = "GOCSPX-xxxxxxxxxxxxxxx"
redirect_uri  = "https://your-app-name.streamlit.app"
```

That's it. Users now see "Sign in with Google" and get Drive sync automatically.

> **Without secrets** the app still works in Demo Mode — full functionality, just no persistent login or Drive sync.

---

## 📂 Files

```
app.py              ← Streamlit UI (5 steps + viewer + downloads)
engine.py           ← Timetable generation engine (pure Python)
auth.py             ← Google OAuth 2.0 (login, token refresh, logout)
drive_sync.py       ← Google Drive API (save, load, list, delete configs)
demo_config.json    ← 8-class demo school for instant testing
requirements.txt    ← streamlit, openpyxl, requests
.streamlit/
  config.toml       ← Theme
  secrets.toml      ← OAuth credentials template (DO NOT commit to public repo!)
```

---

## ✅ Feature Status

| Feature | Phase |
|---|---|
| Full 5-step timetable wizard | ✅ 1 |
| Stage 1 → Task Analysis → Stage 2 → Force Fill | ✅ 1 |
| Classwise + teacherwise viewer | ✅ 1 |
| 5 Excel download formats | ✅ 1 |
| Manual JSON config backup/restore | ✅ 1 |
| Google Sign-In (OAuth 2.0) | ✅ 2 |
| Demo Mode (no login needed) | ✅ 2 |
| Token refresh (stays logged in) | ✅ 2 |
| **Drive auto-save on every step navigation** | ✅ 3 |
| **Drive config browser (load/delete saved configs)** | ✅ 3 |
| **Per-user Drive folder isolation** | ✅ 3 |
| **drive.file scope (can only see own app files)** | ✅ 3 |

---

## 🧪 Testing Locally

```bash
pip install streamlit openpyxl requests
streamlit run app.py
```

Click **"Continue as Demo User"** → **"Load Demo Config"** → Step 4 → generate.

Drive sync won't work locally without OAuth credentials, but all other features will.
