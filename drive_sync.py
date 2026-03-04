"""
drive_sync.py — Google Drive sync for Timetable Generator
──────────────────────────────────────────────────────────
All configs are stored in the user's Drive inside a folder:
  "Timetable Generator"  (created automatically on first save)

Each school config is saved as:
  timetable_{slug}.json

Uses only the drive.file scope — the app can only see files IT created,
so users' other Drive files are completely private.

All API calls use the access_token from auth.get_access_token().
No additional Google libraries needed — pure requests.
"""

import json
import time
import streamlit as st

try:
    import requests as _req
    _requests_ok = True
except ImportError:
    _requests_ok = False

import auth

DRIVE_FILES_URL   = "https://www.googleapis.com/drive/v3/files"
DRIVE_UPLOAD_URL  = "https://www.googleapis.com/upload/drive/v3/files"
APP_FOLDER_NAME   = "Timetable Generator"
MIME_JSON         = "application/json"
MIME_FOLDER       = "application/vnd.google-apps.folder"

# Session-state keys
_FOLDER_ID_KEY    = "_drive_folder_id"
_LAST_SAVED_KEY   = "_drive_last_saved"
_SAVE_STATUS_KEY  = "_drive_save_status"   # 'ok' | 'error' | 'saving' | None
_CONFIGS_CACHE    = "_drive_configs_cache"  # list of {id, name, school, modified}
_CACHE_TS_KEY     = "_drive_cache_ts"


# ── Internal helpers ──────────────────────────────────────────────────────────

def _token():
    return auth.get_access_token()

def _headers():
    return {"Authorization": "Bearer {}".format(_token())}

def _is_demo():
    u = auth.get_user()
    return not u or u.get("id") == "demo"

def _slug(school_name: str) -> str:
    """Convert school name to a safe filename slug."""
    s = school_name.strip().lower() if school_name.strip() else "untitled"
    safe = "".join(c if c.isalnum() or c in "-_ " else "_" for c in s)
    return safe.replace(" ", "_")[:32]


def drive_available() -> bool:
    """True if user is logged in with a real Google account and has a token."""
    if not _requests_ok:
        return False
    if _is_demo():
        return False
    return bool(_token())


# ── Folder management ─────────────────────────────────────────────────────────

def _get_or_create_folder() -> str | None:
    """
    Return the Drive folder ID for "Timetable Generator".
    Creates the folder if it doesn't exist.
    Caches in session_state to avoid repeated API calls.
    """
    cached = st.session_state.get(_FOLDER_ID_KEY)
    if cached:
        return cached

    tok = _token()
    if not tok:
        return None

    # Search for existing folder
    try:
        resp = _req.get(DRIVE_FILES_URL, headers=_headers(), params={
            "q":      "name='{}' and mimeType='{}' and trashed=false".format(
                      APP_FOLDER_NAME, MIME_FOLDER),
            "fields": "files(id,name)",
            "spaces": "drive",
        }, timeout=10)
        resp.raise_for_status()
        files = resp.json().get("files", [])
    except Exception as e:
        st.session_state[_SAVE_STATUS_KEY] = "error"
        return None

    if files:
        folder_id = files[0]["id"]
        st.session_state[_FOLDER_ID_KEY] = folder_id
        return folder_id

    # Create the folder
    try:
        resp = _req.post(DRIVE_FILES_URL, headers=_headers(), json={
            "name":     APP_FOLDER_NAME,
            "mimeType": MIME_FOLDER,
        }, timeout=10)
        resp.raise_for_status()
        folder_id = resp.json()["id"]
        st.session_state[_FOLDER_ID_KEY] = folder_id
        return folder_id
    except Exception:
        return None


# ── Core API: save ────────────────────────────────────────────────────────────

def save_config(config: dict, class_config: dict, step3: dict,
                school_name: str = "") -> bool:
    """
    Save / update config to Google Drive.
    Returns True on success, False on failure.
    Stores save status in session_state for UI display.
    """
    if not drive_available():
        return False

    st.session_state[_SAVE_STATUS_KEY] = "saving"
    folder_id = _get_or_create_folder()
    if not folder_id:
        st.session_state[_SAVE_STATUS_KEY] = "error"
        return False

    slug      = _slug(school_name or config.get("school_name", ""))
    filename  = "timetable_{}.json".format(slug)
    payload   = json.dumps({
        "config":       config,
        "class_config": class_config,
        "step3":        step3,
        "_meta": {
            "school":    school_name or config.get("school_name", ""),
            "saved_at":  time.strftime("%Y-%m-%dT%H:%M:%SZ", time.gmtime()),
            "saved_by":  (auth.get_user() or {}).get("email", ""),
        },
    }, indent=2).encode("utf-8")

    # Check if a file with this name already exists in our folder
    existing_id = _find_file_id(filename, folder_id)

    try:
        if existing_id:
            # Update existing file (PATCH)
            resp = _req.patch(
                "{}/{}".format(DRIVE_UPLOAD_URL, existing_id),
                headers={**_headers(), "Content-Type": MIME_JSON},
                params={"uploadType": "media"},
                data=payload,
                timeout=15,
            )
        else:
            # Create new file (multipart)
            import io
            metadata = json.dumps({
                "name":    filename,
                "parents": [folder_id],
            }).encode("utf-8")
            boundary = "boundary_timetable_app_xyz"
            body = (
                b"--" + boundary.encode() + b"\r\n"
                b"Content-Type: application/json; charset=UTF-8\r\n\r\n"
                + metadata + b"\r\n"
                b"--" + boundary.encode() + b"\r\n"
                b"Content-Type: " + MIME_JSON.encode() + b"\r\n\r\n"
                + payload + b"\r\n"
                b"--" + boundary.encode() + b"--"
            )
            resp = _req.post(
                DRIVE_UPLOAD_URL,
                headers={
                    **_headers(),
                    "Content-Type": "multipart/related; boundary={}".format(boundary),
                },
                params={"uploadType": "multipart"},
                data=body,
                timeout=15,
            )

        resp.raise_for_status()
        st.session_state[_LAST_SAVED_KEY]  = time.time()
        st.session_state[_SAVE_STATUS_KEY] = "ok"
        # Bust the config list cache
        st.session_state.pop(_CONFIGS_CACHE, None)
        return True

    except Exception as e:
        st.session_state[_SAVE_STATUS_KEY] = "error"
        return False


# ── Core API: list ────────────────────────────────────────────────────────────

def list_configs(force_refresh: bool = False) -> list:
    """
    Return list of saved configs from Drive.
    Each item: {id, name, school, modified, filename}
    Cached for 30 seconds to avoid hammering the API.
    """
    if not drive_available():
        return []

    # Use cache unless stale or forced
    cache_ts = st.session_state.get(_CACHE_TS_KEY, 0)
    if not force_refresh and (time.time() - cache_ts) < 30:
        return st.session_state.get(_CONFIGS_CACHE, [])

    folder_id = _get_or_create_folder()
    if not folder_id:
        return []

    try:
        resp = _req.get(DRIVE_FILES_URL, headers=_headers(), params={
            "q":      "'{}' in parents and trashed=false and mimeType='{}'".format(
                      folder_id, MIME_JSON),
            "fields": "files(id,name,modifiedTime,size)",
            "orderBy": "modifiedTime desc",
        }, timeout=10)
        resp.raise_for_status()
        files = resp.json().get("files", [])
    except Exception:
        return st.session_state.get(_CONFIGS_CACHE, [])

    results = []
    for f in files:
        fname = f.get("name", "")
        # Strip prefix/suffix to get school name
        school = fname
        if school.startswith("timetable_"):
            school = school[len("timetable_"):]
        if school.endswith(".json"):
            school = school[:-5]
        school = school.replace("_", " ").title()

        results.append({
            "id":       f["id"],
            "filename": fname,
            "school":   school,
            "modified": f.get("modifiedTime", ""),
        })

    st.session_state[_CONFIGS_CACHE] = results
    st.session_state[_CACHE_TS_KEY]  = time.time()
    return results


# ── Core API: load ────────────────────────────────────────────────────────────

def load_config(file_id: str) -> dict | None:
    """
    Download and parse a config JSON from Drive by file_id.
    Returns the parsed dict or None on failure.
    """
    if not drive_available():
        return None
    try:
        resp = _req.get(
            "{}/{}".format(DRIVE_FILES_URL, file_id),
            headers=_headers(),
            params={"alt": "media"},
            timeout=15,
        )
        resp.raise_for_status()
        return resp.json()
    except Exception:
        return None


def delete_config(file_id: str) -> bool:
    """Move a Drive config to trash."""
    if not drive_available():
        return False
    try:
        resp = _req.patch(
            "{}/{}".format(DRIVE_FILES_URL, file_id),
            headers=_headers(),
            json={"trashed": True},
            timeout=10,
        )
        resp.raise_for_status()
        st.session_state.pop(_CONFIGS_CACHE, None)
        return True
    except Exception:
        return False


# ── Internal ──────────────────────────────────────────────────────────────────

def _find_file_id(filename: str, folder_id: str) -> str | None:
    """Return Drive file ID if filename exists in folder, else None."""
    try:
        resp = _req.get(DRIVE_FILES_URL, headers=_headers(), params={
            "q":      "name='{}' and '{}' in parents and trashed=false".format(
                      filename, folder_id),
            "fields": "files(id)",
        }, timeout=10)
        resp.raise_for_status()
        files = resp.json().get("files", [])
        return files[0]["id"] if files else None
    except Exception:
        return None


# ── UI helpers ────────────────────────────────────────────────────────────────

def render_save_status():
    """
    Small inline status indicator — call in sidebar after Save button.
    Shows: ✅ Saved X mins ago  |  ⏳ Saving…  |  ❌ Save failed
    """
    if not drive_available():
        return
    status = st.session_state.get(_SAVE_STATUS_KEY)
    last   = st.session_state.get(_LAST_SAVED_KEY)

    if status == "saving":
        st.sidebar.caption("⏳ Saving to Drive…")
    elif status == "ok" and last:
        elapsed = int((time.time() - last) / 60)
        if elapsed < 1:
            st.sidebar.caption("✅ Saved just now")
        elif elapsed == 1:
            st.sidebar.caption("✅ Saved 1 min ago")
        else:
            st.sidebar.caption("✅ Saved {} mins ago".format(elapsed))
    elif status == "error":
        st.sidebar.caption("❌ Drive save failed — check connection")


def render_drive_configs_panel():
    """
    Render the "Your saved configs" panel shown on the dashboard / step 1.
    Returns the loaded payload dict if user picks a config, else None.
    """
    if not drive_available():
        return None

    configs = list_configs()
    if not configs:
        st.info("☁️ No saved configs in your Drive yet. "
                "Your configurations will auto-save to Google Drive as you work.")
        return None

    st.markdown("### ☁️ Your Saved Configs")
    st.caption("Stored in Google Drive → Timetable Generator folder")

    loaded = None
    for cfg in configs:
        mod   = cfg["modified"][:10] if cfg.get("modified") else "—"
        col1, col2, col3 = st.columns([4, 2, 1])
        with col1:
            st.markdown("**{}**".format(cfg["school"]))
            st.caption("Last saved: {}".format(mod))
        with col2:
            if st.button("📂 Load", key="load_drive_{}".format(cfg["id"]),
                         use_container_width=True):
                with st.spinner("Loading from Drive…"):
                    payload = load_config(cfg["id"])
                if payload:
                    loaded = payload
                    st.session_state[_SAVE_STATUS_KEY] = None
                    st.session_state[_LAST_SAVED_KEY]  = None
                else:
                    st.error("Could not load config — try again.")
        with col3:
            if st.button("🗑️", key="del_drive_{}".format(cfg["id"]),
                         help="Delete from Drive"):
                if delete_config(cfg["id"]):
                    st.success("Deleted.")
                    st.rerun()

        st.divider()

    if st.button("🔄 Refresh list", use_container_width=True):
        list_configs(force_refresh=True)
        st.rerun()

    return loaded


def auto_save(config: dict, class_config: dict, step3: dict):
    """
    Trigger a background-style save.
    Call this at the end of any step that modifies config.
    Only saves if school_name is set and Drive is available.
    """
    school = config.get("school_name", "").strip()
    if not school or not drive_available():
        return
    # Rate-limit: don't save more than once every 20 seconds
    last = st.session_state.get(_LAST_SAVED_KEY, 0)
    if time.time() - last < 20:
        return
    save_config(config, class_config, step3, school)
