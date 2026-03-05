"""
extract_my_configs.py
─────────────────────────────────────────────────────────────────────────────
Run this script on YOUR computer (the same machine where you used the original
timetable app).  It will:

  1. Find all your saved Step 1, Step 2, and Step 3 config files in
     ~/TimetableConfigs/
  2. Merge them together into a single Streamlit-compatible JSON file
  3. Save one file per combination to your Desktop

Usage:
  python extract_my_configs.py

Requirements: Python 3.8+  (no extra packages needed)
─────────────────────────────────────────────────────────────────────────────
"""

import json
import sys
from pathlib import Path
from datetime import datetime


# ── Locate config folder ──────────────────────────────────────────────────────
CONFIG_DIR = Path.home() / "TimetableConfigs"
DESKTOP    = Path.home() / "Desktop"

if not CONFIG_DIR.exists():
    print("❌  ~/TimetableConfigs not found.")
    print("    Make sure you're running this on the same computer as the original app.")
    sys.exit(1)

files = list(CONFIG_DIR.glob("*.json"))
if not files:
    print("❌  No .json files found in ~/TimetableConfigs/")
    sys.exit(1)

print(f"✓  Found {len(files)} file(s) in {CONFIG_DIR}\n")


# ── Categorise files ──────────────────────────────────────────────────────────
step1_files = []
step2_files = []
step3_files = []

for f in sorted(files, reverse=True):
    try:
        d = json.loads(f.read_text(encoding="utf-8"))
    except Exception:
        print(f"  ⚠  Could not parse {f.name} — skipping")
        continue

    if d.get("step") == 3 or "step3_data" in d:
        step3_files.append((f, d))
        print(f"  [Step 3] {f.stem}  — saved {d.get('saved_at','?')}")
    elif "assignments" in d:
        step2_files.append((f, d))
        classes = list(d["assignments"].keys())
        print(f"  [Step 2] {f.stem}  — {len(classes)} classes  — saved {d.get('saved_at','?')}")
    elif "teacher_names" in d or "periods_per_day" in d:
        step1_files.append((f, d))
        print(f"  [Step 1] {f.stem}  — saved {d.get('saved_at','?')}")
    else:
        print(f"  [?????] {f.stem} — unrecognised format, skipping")

print()


# ── Converter helpers ─────────────────────────────────────────────────────────

def convert_step1(d1: dict) -> dict:
    """Convert Step 1 data → config dict for Streamlit app."""
    classes_raw = d1.get("classes", {})
    return {
        "school_name":        "",          # user can fill in the app
        "working_days":       int(d1.get("working_days", 5)),
        "periods_per_day":    int(d1.get("periods_per_day", 8)),
        "periods_first_half": int(d1.get("periods_first_half", 4)),
        "classes": {int(k): int(v) for k, v in classes_raw.items()},
        "_teacher_names": d1.get("teacher_names", []),  # informational
    }


def convert_step2(d2: dict) -> dict:
    """Convert Step 2 assignments → class_config dict for Streamlit app."""
    raw = d2.get("assignments", {})
    class_config = {}
    for cn, cd in raw.items():
        teacher = cd.get("teacher", "")
        # teacher may be a plain string or an old StringVar-serialised value
        if isinstance(teacher, dict):
            teacher = teacher.get("value", "") or ""

        tp = cd.get("teacher_period", 1)
        try:
            tp = int(tp)
        except (ValueError, TypeError):
            tp = 1

        subjects = []
        for s in cd.get("subjects", []):
            # Normalise consecutive field (bool or 'Yes'/'No')
            consec = s.get("consecutive", False)
            if isinstance(consec, bool):
                consec_str = "Yes" if consec else "No"
            else:
                consec_str = "Yes" if str(consec).lower() in ("yes", "true", "1") else "No"

            # Normalise parallel flag
            par = s.get("parallel", False)
            if isinstance(par, str):
                par = par.lower() in ("true", "yes", "1")

            subjects.append({
                "name":             s.get("name", "").strip(),
                "teacher":          s.get("teacher", "").strip(),
                "periods":          int(s.get("periods", 1)),
                "consecutive":      consec_str,
                "parallel":         par,
                "parallel_subject": (s.get("parallel_subject") or "").strip(),
                "parallel_teacher": (s.get("parallel_teacher") or "").strip(),
                "periods_pref":     list(s.get("periods_pref", [])),
                "days_pref":        list(s.get("days_pref", [])),
            })

        class_config[cn] = {
            "teacher":        teacher.strip(),
            "teacher_period": tp,
            "subjects":       subjects,
        }
    return class_config


def convert_step3(d3: dict) -> dict:
    """Convert Step 3 data → step3 dict for Streamlit app."""
    raw_s3   = d3.get("step3_data", {})
    raw_unav = d3.get("step3_unavailability", {})
    step3 = {}

    for teacher, td in raw_s3.items():
        combines = []
        for cb in td.get("combines", []):
            combines.append({
                "classes":  list(cb.get("classes", [])),
                "subjects": list(cb.get("subjects", [])),
            })
        step3[teacher] = {
            "combines":       combines,
            "unavailability": {},
        }

    # Merge unavailability
    for teacher, u in raw_unav.items():
        days    = list(u.get("days", []))
        periods = list(u.get("periods", []))
        if teacher not in step3:
            step3[teacher] = {"combines": [], "unavailability": {}}
        step3[teacher]["unavailability"] = {"days": days, "periods": periods}

    return step3


# ── Build and export combinations ─────────────────────────────────────────────

DESKTOP.mkdir(parents=True, exist_ok=True)
exported = 0

def export(config, class_config, step3, name):
    global exported
    payload = {
        "config":       config,
        "class_config": class_config,
        "step3":        step3,
        "_meta": {
            "exported_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "source":      "extract_my_configs.py",
        },
    }
    out = DESKTOP / f"{name}.json"
    with open(out, "w", encoding="utf-8") as fh:
        json.dump(payload, fh, indent=2, ensure_ascii=False)
    print(f"  ✅  Exported → {out}")
    exported += 1


# Strategy:
#   • If we have Step 2 (assignments) — that's the most complete data.
#     Pair it with each Step 3 we have, or just export it alone.
#   • If we only have Step 1 — export that alone.
#   • Always include a Step 1 config dict if available.

if step2_files:
    # Use the most recent Step 1 as base config (for periods/days/classes)
    base_config = convert_step1(step1_files[0][1]) if step1_files else {
        "school_name": "", "working_days": 5, "periods_per_day": 8,
        "periods_first_half": 4, "classes": {},
    }

    for s2f, s2d in step2_files:
        class_config = convert_step2(s2d)

        # Try to find a matching Step 3
        if step3_files:
            for s3f, s3d in step3_files:
                step3 = convert_step3(s3d)
                name  = f"streamlit_{s2f.stem}__{s3f.stem}"
                print(f"\nExporting: Step2={s2f.stem} + Step3={s3f.stem}")
                export(base_config, class_config, step3, name)
        else:
            name = f"streamlit_{s2f.stem}"
            print(f"\nExporting: Step2={s2f.stem} (no Step 3 found)")
            export(base_config, class_config, {}, name)

elif step1_files:
    # Only Step 1 data available
    for s1f, s1d in step1_files:
        config = convert_step1(s1d)
        name   = f"streamlit_{s1f.stem}"
        print(f"\nExporting: Step1={s1f.stem} only (no Step 2 found)")
        export(config, {}, {}, name)
else:
    print("⚠  No usable config files found.")


# ── Summary ───────────────────────────────────────────────────────────────────
print()
print("─" * 60)
print(f"Done! {exported} file(s) saved to: {DESKTOP}")
print()
print("Next steps:")
print("  1. Open the Streamlit app in your browser")
print("  2. Sidebar → 📂 Load Config (JSON)")
print("  3. Select any of the exported files from your Desktop")
print("  4. Your configuration will be restored immediately.")
print("─" * 60)
