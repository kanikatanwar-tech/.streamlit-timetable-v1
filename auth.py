"""
auth.py — Google OAuth 2.0 for Streamlit Cloud
─────────────────────────────────────────────────
Scopes requested:
  openid email profile
  https://www.googleapis.com/auth/drive.file   ← Phase 3 Drive sync

The access_token is stored in st.session_state['_access_token'] for use
by drive_sync.py.  The refresh_token (if issued) is stored in
st.session_state['_refresh_token'] so we can silently refresh mid-session.

Setup (.streamlit/secrets.toml):
  [google_oauth]
  client_id     = "xxx.apps.googleusercontent.com"
  client_secret = "xxx"
  redirect_uri  = "https://your-app.streamlit.app"
"""

import streamlit as st
import urllib.parse, secrets, time

try:
    import requests as _req
    _requests_ok = True
except ImportError:
    _requests_ok = False

GOOGLE_AUTH_URL   = "https://accounts.google.com/o/oauth2/v2/auth"
GOOGLE_TOKEN_URL  = "https://oauth2.googleapis.com/token"
GOOGLE_REVOKE_URL = "https://oauth2.googleapis.com/revoke"
GOOGLE_USERINFO   = "https://www.googleapis.com/oauth2/v2/userinfo"

# Drive.file = only files THIS app creates/opens (safest scope)
SCOPES = " ".join([
    "openid",
    "email",
    "profile",
    "https://www.googleapis.com/auth/drive.file",
])


# ── Helpers ───────────────────────────────────────────────────────────────────
def _secrets_ok():
    try:
        st.secrets["google_oauth"]["client_id"]
        st.secrets["google_oauth"]["client_secret"]
        st.secrets["google_oauth"]["redirect_uri"]
        return True
    except (KeyError, FileNotFoundError):
        return False

def _get(key):
    return st.secrets["google_oauth"][key]

def _post(url, **kwargs):
    if not _requests_ok:
        raise RuntimeError("'requests' package missing — add to requirements.txt")
    return _req.post(url, **kwargs)

def _get_req(url, **kwargs):
    if not _requests_ok:
        raise RuntimeError("'requests' package missing")
    return _req.get(url, **kwargs)


# ── Token refresh ─────────────────────────────────────────────────────────────
def _maybe_refresh():
    """
    Silently refresh access_token if it is close to expiry.
    Called automatically by get_access_token().
    """
    exp  = st.session_state.get("_token_expiry", 0)
    now  = time.time()
    rt   = st.session_state.get("_refresh_token")
    if exp - now > 120 or not rt:
        return   # still valid or no refresh token

    if not _secrets_ok():
        return
    try:
        resp = _post(GOOGLE_TOKEN_URL, data={
            "client_id":     _get("client_id"),
            "client_secret": _get("client_secret"),
            "refresh_token": rt,
            "grant_type":    "refresh_token",
        }, timeout=10)
        resp.raise_for_status()
        data = resp.json()
        st.session_state["_access_token"] = data["access_token"]
        st.session_state["_token_expiry"]  = now + int(data.get("expires_in", 3600))
    except Exception:
        pass   # will fail gracefully on next Drive call


def get_access_token():
    """Return current valid access token, or None (demo / not logged in)."""
    _maybe_refresh()
    return st.session_state.get("_access_token")


# ── Build login URL ───────────────────────────────────────────────────────────
def build_login_url():
    state = secrets.token_urlsafe(32)
    st.session_state["_oauth_state"] = state
    params = {
        "client_id":     _get("client_id"),
        "redirect_uri":  _get("redirect_uri"),
        "response_type": "code",
        "scope":         SCOPES,
        "state":         state,
        "access_type":   "offline",   # get refresh_token
        "prompt":        "consent",   # force consent so refresh_token is always issued
    }
    return GOOGLE_AUTH_URL + "?" + urllib.parse.urlencode(params), state


# ── OAuth callback ────────────────────────────────────────────────────────────
def handle_callback():
    """
    Must be called at the very top of app.py before any rendering.
    Returns True when a fresh login just completed.
    """
    if not _secrets_ok() or not _requests_ok:
        return False

    params = st.query_params
    code   = params.get("code")
    state  = params.get("state")
    if not code:
        return False
    if st.session_state.get("_oauth_code_handled") == code:
        return False
    if state != st.session_state.get("_oauth_state"):
        st.query_params.clear()
        return False

    # Exchange code → tokens
    try:
        resp = _post(GOOGLE_TOKEN_URL, data={
            "code":          code,
            "client_id":     _get("client_id"),
            "client_secret": _get("client_secret"),
            "redirect_uri":  _get("redirect_uri"),
            "grant_type":    "authorization_code",
        }, timeout=10)
        resp.raise_for_status()
        token_data = resp.json()
    except Exception as e:
        st.error("OAuth token exchange failed: {}".format(e))
        st.query_params.clear()
        return False

    access_token  = token_data.get("access_token")
    refresh_token = token_data.get("refresh_token")
    expires_in    = int(token_data.get("expires_in", 3600))

    if not access_token:
        st.error("No access token returned: {}".format(token_data))
        st.query_params.clear()
        return False

    # Fetch user profile
    try:
        ui_resp = _get_req(GOOGLE_USERINFO,
            headers={"Authorization": "Bearer {}".format(access_token)}, timeout=10)
        ui_resp.raise_for_status()
        ui = ui_resp.json()
    except Exception as e:
        st.error("Failed to fetch user info: {}".format(e))
        st.query_params.clear()
        return False

    # Persist everything in session
    st.session_state["user"] = {
        "id":      ui.get("id", ""),
        "email":   ui.get("email", ""),
        "name":    ui.get("name", ui.get("email", "User")),
        "picture": ui.get("picture", ""),
    }
    st.session_state["_access_token"]  = access_token
    st.session_state["_refresh_token"] = refresh_token or ""
    st.session_state["_token_expiry"]  = time.time() + expires_in
    st.session_state["_oauth_code_handled"] = code
    st.query_params.clear()
    return True


# ── User helpers ──────────────────────────────────────────────────────────────
def get_user():
    return st.session_state.get("user")

def logout():
    token = st.session_state.get("_access_token")
    if token and _requests_ok:
        try:
            _post(GOOGLE_REVOKE_URL, params={"token": token}, timeout=5)
        except Exception:
            pass
    for k in ["user","_access_token","_refresh_token","_token_expiry",
              "_oauth_state","_oauth_code_handled"]:
        st.session_state.pop(k, None)


# ── UI components ─────────────────────────────────────────────────────────────
def render_login_page():
    st.markdown("""
    <style>
    .lw  { max-width:480px; margin:60px auto; text-align:center; }
    .lc  { background:white; border-radius:16px; padding:2.5rem 2rem;
           box-shadow:0 8px 32px rgba(0,0,0,.12); }
    .lt  { font-size:2.2rem; font-weight:700; color:#2c3e50; margin-bottom:.3rem; }
    .ls  { color:#7f8c8d; margin-bottom:2rem; font-size:1rem; }
    .gb  { display:inline-flex; align-items:center; gap:12px; background:white;
           border:2px solid #dadce0; border-radius:8px; padding:12px 28px;
           font-size:1rem; font-weight:600; color:#3c4043; text-decoration:none;
           box-shadow:0 2px 6px rgba(0,0,0,.08); transition:box-shadow .2s; }
    .gb:hover { box-shadow:0 4px 16px rgba(0,0,0,.16); border-color:#aaa; }
    </style>
    """, unsafe_allow_html=True)

    GSVG = """<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 48 48" width="24" height="24">
      <path fill="#EA4335" d="M24 9.5c3.54 0 6.71 1.22 9.21 3.6l6.85-6.85C35.9 2.38 30.47 0 24 0
        14.62 0 6.51 5.38 2.56 13.22l7.98 6.19C12.43 13.72 17.74 9.5 24 9.5z"/>
      <path fill="#4285F4" d="M46.98 24.55c0-1.57-.15-3.09-.38-4.55H24v9.02h12.94
        c-.58 2.96-2.26 5.48-4.78 7.18l7.73 6c4.51-4.18 7.09-10.36 7.09-17.65z"/>
      <path fill="#FBBC05" d="M10.53 28.59c-.48-1.45-.76-2.99-.76-4.59s.27-3.14.76-4.59
        l-7.98-6.19C.92 16.46 0 20.12 0 24c0 3.88.92 7.54 2.56 10.78l7.97-6.19z"/>
      <path fill="#34A853" d="M24 48c6.48 0 11.93-2.13 15.89-5.81l-7.73-6
        c-2.18 1.48-4.97 2.36-8.16 2.36-6.26 0-11.57-4.22-13.47-9.91l-7.98 6.19
        C6.51 42.62 14.62 48 24 48z"/>
    </svg>"""

    if not _secrets_ok():
        st.markdown('<div class="lw"><div class="lc">', unsafe_allow_html=True)
        st.markdown('<div class="lt">📅 Timetable Generator</div>', unsafe_allow_html=True)
        st.markdown('<div class="ls">School timetable scheduling</div>', unsafe_allow_html=True)
        st.warning("Google OAuth not configured — running in Demo Mode.")
        st.markdown("""
Add to `.streamlit/secrets.toml` to enable Google login:
```toml
[google_oauth]
client_id     = "xxx.apps.googleusercontent.com"
client_secret = "xxx"
redirect_uri  = "https://your-app.streamlit.app"
```
""")
        if st.button("▶ Continue as Demo User", type="primary", use_container_width=True):
            st.session_state["user"] = {
                "id": "demo", "email": "demo@school.local",
                "name": "Demo User", "picture": ""
            }
            st.rerun()
        st.markdown('</div></div>', unsafe_allow_html=True)
        return

    login_url, _ = build_login_url()
    st.markdown("""
    <div class="lw"><div class="lc">
      <div class="lt">📅 Timetable Generator</div>
      <div class="ls">Sign in to access your saved timetable configs from Google Drive</div>
      <a href="{url}" class="gb">{svg}&nbsp; Sign in with Google</a>
      <br><br>
      <p style="color:#bbb;font-size:.8rem">
        We only access files <em>this app creates</em> in your Drive.<br>
        Your school data is never shared.
      </p>
    </div></div>
    """.format(url=login_url, svg=GSVG), unsafe_allow_html=True)


def render_user_badge():
    """Call inside st.sidebar block."""
    user = get_user()
    if not user:
        return
    pic = user.get("picture", "")
    if pic and user.get("id") != "demo":
        st.markdown(
            '<img src="{}" style="border-radius:50%;width:48px;height:48px;'
            'display:block;margin:0 auto 6px;">'.format(pic),
            unsafe_allow_html=True)
    else:
        st.markdown('<div style="font-size:2rem;text-align:center">👤</div>',
            unsafe_allow_html=True)
    st.markdown("<div style='text-align:center'><b>{}</b><br>"
        "<span style='font-size:.75rem;color:#aaa'>{}</span></div>".format(
        user["name"], user["email"]), unsafe_allow_html=True)
