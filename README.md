# Timetable Generator V4.0 — Streamlit Cloud App

Converted from the original tkinter desktop app (`timetable_generator_v16.py`) to a web app that can be used by anyone, anywhere via Streamlit Cloud.

## 🚀 Deploy to Streamlit Cloud (Free)

### Step 1: Create a GitHub repository
1. Go to [github.com](https://github.com) → New repository
2. Name it (e.g. `timetable-generator`)
3. Upload these two files:
   - `timetable_app.py`
   - `requirements.txt`

### Step 2: Deploy on Streamlit Cloud
1. Go to [share.streamlit.io](https://share.streamlit.io)
2. Sign in with your GitHub account
3. Click **"New app"**
4. Select your repository, branch (`main`), and file (`timetable_app.py`)
5. Click **"Deploy!"**

That's it! In ~2 minutes your app will be live at a URL like:
`https://your-app-name.streamlit.app`

---

## 📖 How to Use

### Saving & Loading Configurations
- **Save**: Click the **"💾 Download ... JSON"** button → a JSON file downloads to your computer
- **Load**: Click the upload area → select your previously downloaded JSON → form fields auto-populate
- Each step (Step 1, Step 2, Step 3) has its own save/load

### Step 1 — School Configuration
1. Set periods per day, working days, period halves
2. Upload teacher Excel file (Column A only, one name per row)
3. Set number of sections for each class (6–12)
4. Click **"Continue to Step 2"**

### Step 2 — Class Configuration
- Each class has its own tab
- Set the Class Teacher and CT Period
- Add subjects with: teacher, periods/week, period preferences, day preferences, consecutive flag, parallel options
- Click **"Validate & Complete"** — shows conflict report
- If all clear, proceed to Step 3

### Step 3 — Teacher Manager
- View all teachers with their workload
- **Skip**: Mark overloaded teachers to skip workload enforcement
- **Edit/Combine**: Select entries to combine (share same slot across classes)
- **Unavailability**: Block specific days/periods for any teacher
- Click **"Proceed to Step 4"**

### Step 4 — Generate Timetable
1. Click **"Stage 1"** → Places Class Teacher slots and preference-pinned periods
2. Click **"Task Analysis"** → Review parallel/consecutive groups → **"Allocate Periods"**
3. Click **"Proceed to Stage 2"** → Fills remaining periods
4. If unplaced periods remain, click **"Force Fill"** (CSP solver)
5. Download Excel in 5 formats: Classwise, Teacherwise, CT List, Workload, One-Sheet

---

## 📁 Files
| File | Purpose |
|------|---------|
| `timetable_app.py` | Main Streamlit app (all-in-one) |
| `requirements.txt` | Python dependencies for Streamlit Cloud |

---

## 💡 Notes
- All computation logic is **100% identical** to the original desktop app
- Config files are saved to **your local computer** (browser downloads)
- Upload saved JSON files to restore your work in any session
- No server-side storage — your data stays with you
