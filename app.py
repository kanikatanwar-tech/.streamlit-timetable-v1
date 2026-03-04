import streamlit as st
import json, copy, time
from engine import TimetableEngine
import auth
import drive_sync

st.set_page_config(page_title="Timetable Generator", page_icon="📅", layout="wide")

# ── OAuth callback must run before any rendering ──────────────────────────────
just_logged_in = auth.handle_callback()

# ── Session defaults ──────────────────────────────────────────────────────────
DEFAULTS = {
    'step': 1,
    'config': {
        'school_name': '',
        'working_days': 5,
        'periods_per_day': 8,
        'periods_first_half': 4,
        'classes': {6:0, 7:0, 8:0, 9:0, 10:0, 11:0, 12:0},
    },
    'class_config': {},
    'step3': {},
    'timetable': None,
    'engine': None,
    'stage': 0,
    'stage1_result': None,
    'task_analysis': None,
    'force_fill_notes': None,
    'unplaced': 0,
}
for k, v in DEFAULTS.items():
    if k not in st.session_state:
        st.session_state[k] = copy.deepcopy(v) if isinstance(v, dict) else v

def s():  return st.session_state
def go(n): s()['step'] = n; st.rerun()
def cfg(): return s()['config']
DAYS_ALL = ['Mon','Tue','Wed','Thu','Fri','Sat','Sun']


# ── CSS ───────────────────────────────────────────────────────────────────────
st.markdown("""
<style>
.stApp { background:#f0f2f6; }
.step-header { background:linear-gradient(135deg,#1a252f,#2980b9); color:white;
    padding:1.1rem 1.5rem; border-radius:12px; margin-bottom:1.5rem; }
.step-header h2 { margin:0; font-size:1.4rem; }
.step-header p  { margin:0; opacity:.8; font-size:.88rem; margin-top:.2rem; }
.card { background:white; border-radius:12px; padding:1.2rem 1.5rem;
    box-shadow:0 2px 8px rgba(0,0,0,.07); margin-bottom:1rem; }
.user-card { background:linear-gradient(135deg,#1a252f,#2c3e50); color:white;
    border-radius:12px; padding:1rem; text-align:center; margin-bottom:.8rem; }
.drive-badge { display:inline-flex; align-items:center; gap:6px;
    background:#e8f5e9; border:1px solid #4caf50; border-radius:20px;
    padding:3px 12px; font-size:.78rem; color:#2e7d32; margin-top:4px; }
.demo-badge  { display:inline-flex; align-items:center; gap:6px;
    background:#fff8e1; border:1px solid #ffc107; border-radius:20px;
    padding:3px 12px; font-size:.78rem; color:#e65100; margin-top:4px; }
.cfg-card { background:white; border-radius:10px; padding:.9rem 1.1rem;
    box-shadow:0 2px 6px rgba(0,0,0,.07); margin-bottom:.6rem;
    border-left:4px solid #2980b9; }
.cfg-school { font-weight:700; font-size:1rem; color:#1a252f; }
.cfg-meta   { font-size:.78rem; color:#7f8c8d; margin-top:2px; }
.tt-cell { border:1px solid #e0e0e0; border-radius:7px; padding:6px 8px;
    font-size:.73rem; text-align:center; min-height:54px;
    display:flex; flex-direction:column; justify-content:center;
    background:white; }
.tt-free { background:#fafafa; color:#bbb; }
.tt-normal   { background:#d5e8d4; border-color:#82b366; }
.tt-ct       { background:#a8d5a2; border-color:#6aaf65; font-weight:600; }
.tt-combined { background:#dae8fc; border-color:#6c8ebf; }
.tt-parallel { background:#ffe6cc; border-color:#d6b656; }
.tt-cp       { background:#f8cecc; border-color:#b85450; }
.badge-ok  { background:#27ae60; color:white; padding:2px 9px; border-radius:12px; font-size:.73rem; }
.badge-err { background:#e74c3c; color:white; padding:2px 9px; border-radius:12px; font-size:.73rem; }
.save-ok   { color:#27ae60; font-size:.8rem; }
.save-err  { color:#e74c3c; font-size:.8rem; }
</style>
""", unsafe_allow_html=True)


# ── Gate: show login page if not authenticated ────────────────────────────────
if not auth.get_user():
    auth.render_login_page()
    st.stop()

user    = auth.get_user()
is_demo = (user.get("id") == "demo")


# ── Sidebar ───────────────────────────────────────────────────────────────────
with st.sidebar:

    # User card
    pic = user.get("picture","")
    if pic and not is_demo:
        st.markdown('<div class="user-card"><img src="{}" style="border-radius:50%;'
            'width:52px;height:52px;margin-bottom:6px;"><br>'
            '<b>{}</b><br><span style="font-size:.75rem;opacity:.75">{}</span>'
            '</div>'.format(pic, user["name"], user["email"]), unsafe_allow_html=True)
    else:
        icon = "🧪" if is_demo else "👤"
        st.markdown('<div class="user-card"><div style="font-size:2.2rem">{}</div>'
            '<br><b>{}</b><br><span style="font-size:.75rem;opacity:.75">{}</span>'
            '</div>'.format(icon, user["name"], user["email"]), unsafe_allow_html=True)

    # Drive / Demo badge
    if drive_sync.drive_available():
        st.markdown('<div class="drive-badge">☁️ Drive Sync Active</div>',
            unsafe_allow_html=True)
    elif is_demo:
        st.markdown('<div class="demo-badge">🧪 Demo Mode</div>',
            unsafe_allow_html=True)

    st.markdown("")

    if cfg().get('school_name'):
        st.markdown("**🏫 {}**".format(cfg()['school_name']))

    st.divider()

    # Nav
    steps = [
        (1, "⚙️ School Setup"),
        (2, "📋 Class Config"),
        (3, "🔗 Combines & Rules"),
        (4, "🚀 Generate"),
        (5, "📊 Final Timetable"),
    ]
    for n, label in steps:
        if n == s()['step']:
            st.markdown("&nbsp;&nbsp;**→ {}**".format(label))
        else:
            if st.button(label, key="nav_{}".format(n), use_container_width=True):
                go(n)

    st.divider()

    # Drive Save
    st.markdown("**☁️ Drive Sync**")
    if drive_sync.drive_available():
        school = cfg().get('school_name','').strip()
        if school:
            if st.button("💾 Save to Drive", use_container_width=True, type="primary"):
                with st.spinner("Saving…"):
                    ok = drive_sync.save_config(cfg(), s()['class_config'], s()['step3'], school)
                if ok: st.success("Saved!")
                else:  st.error("Save failed — check connection.")
            drive_sync.render_save_status()
        else:
            st.caption("Set a school name in Step 1 to enable Drive save.")
    else:
        st.caption("Sign in with Google to enable Drive sync." if is_demo
                   else "Drive sync unavailable.")

    st.divider()

    # Manual JSON import/export (always available)
    st.markdown("**📦 Manual Backup**")
    st.download_button(
        "⬇ Export JSON",
        data=json.dumps({
            'config': cfg(),
            'class_config': s()['class_config'],
            'step3': s()['step3'],
        }, indent=2),
        file_name="timetable_{}.json".format(
            cfg().get('school_name','config').replace(' ','_').lower()[:20]),
        mime="application/json",
        use_container_width=True,
    )
    uploaded = st.file_uploader("📂 Import JSON", type="json", label_visibility="collapsed")
    if uploaded:
        try:
            payload = json.load(uploaded)
            s()['config']       = payload.get('config', cfg())
            s()['config']['classes'] = {int(k): v for k,v in s()['config']['classes'].items()}
            s()['class_config'] = payload.get('class_config', {})
            s()['step3']        = payload.get('step3', {})
            s()['stage']        = 0; s()['timetable'] = None; s()['task_analysis'] = None
            st.success("✅ Imported!")
            st.rerun()
        except Exception as e:
            st.error("Import failed: {}".format(e))

    st.divider()

    # Demo config loader
    st.markdown("**🧪 Quick Test**")
    if st.button("Load Demo Config", use_container_width=True):
        import os
        demo_path = os.path.join(os.path.dirname(__file__), 'demo_config.json')
        try:
            with open(demo_path) as f:
                payload = json.load(f)
            s()['config']       = payload['config']
            s()['config']['classes'] = {int(k): v for k,v in s()['config']['classes'].items()}
            s()['class_config'] = payload['class_config']
            s()['step3']        = payload['step3']
            s()['stage']        = 0; s()['timetable'] = None; s()['task_analysis'] = None
            st.success("Demo loaded — 8 classes, 10 teachers!")
            st.rerun()
        except Exception as e:
            st.error(str(e))

    st.divider()

    col1, col2 = st.columns(2)
    with col1:
        if st.button("🔄 Reset", use_container_width=True):
            for k, v in DEFAULTS.items():
                st.session_state[k] = copy.deepcopy(v) if isinstance(v, dict) else v
            st.rerun()
    with col2:
        if st.button("🚪 Sign Out", use_container_width=True):
            auth.logout()
            st.rerun()


# ─────────────────────────────────────────────────────────────────────────────
# RENDER HELPERS
# ─────────────────────────────────────────────────────────────────────────────
def _cell_html(cn, e, eng):
    if e is None:
        st.markdown('<div class="tt-cell tt-free">FREE</div>', unsafe_allow_html=True); return
    et = e.get('type','normal')
    if et == 'combined_parallel':
        l1, l2 = eng.get_combined_par_display(cn, e)
        st.markdown('<div class="tt-cell tt-cp"><b>{}</b><br><span style="color:#666">{}</span></div>'.format(l1,l2), unsafe_allow_html=True)
    elif et == 'parallel':
        st.markdown('<div class="tt-cell tt-parallel"><b>{}</b>/<b>{}</b><br>'
            '<span style="color:#666">{} / {}</span></div>'.format(
            e['subject'],e.get('par_subj','?'),e['teacher'],e.get('par_teach','?')), unsafe_allow_html=True)
    elif et == 'combined':
        cc = e.get('combined_classes',[])
        st.markdown('<div class="tt-cell tt-combined"><b>{}</b><br>'
            '<span style="color:#555;font-size:.68rem">[{}]</span><br>'
            '<span style="color:#666">{}</span></div>'.format(
            e['subject'],'+'.join(cc),e['teacher']), unsafe_allow_html=True)
    else:
        mark  = " ★" if e.get('is_ct') else ""
        color = "tt-ct" if e.get('is_ct') else "tt-normal"
        st.markdown('<div class="tt-cell {}"><b>{}{}</b><br>'
            '<span style="color:#555">{}</span></div>'.format(
            color, e['subject'], mark, e['teacher']), unsafe_allow_html=True)


def _render_class_tt(cn, grid, days, ppd, half1, eng):
    rows = grid.get(cn, [])
    if not rows: st.info("No data for {}".format(cn)); return
    hcols = st.columns([2]+[3]*ppd)
    hcols[0].markdown("**Day**")
    for p in range(ppd):
        hcols[p+1].markdown("**P{} {}**".format(p+1,"①" if p<half1 else "②"))
    for d, dn in enumerate(days):
        dc = st.columns([2]+[3]*ppd)
        dc[0].markdown("**{}**".format(dn))
        for p in range(ppd):
            with dc[p+1]:
                _cell_html(cn, rows[d][p] if d<len(rows) else None, eng)


def _render_teacher_tt(tdata, days, ppd, half1):
    hcols = st.columns([2]+[3]*ppd)
    hcols[0].markdown("**Day**")
    for p in range(ppd):
        hcols[p+1].markdown("**P{} {}**".format(p+1,"①" if p<half1 else "②"))
    for d, dn in enumerate(days):
        dc = st.columns([2]+[3]*ppd)
        dc[0].markdown("**{}**".format(dn))
        for p in range(ppd):
            with dc[p+1]:
                e = tdata[d][p] if d<len(tdata) else None
                if not e:
                    st.markdown('<div class="tt-cell tt-free">FREE</div>', unsafe_allow_html=True)
                else:
                    color = "tt-ct" if e.get('is_ct') else "tt-normal"
                    st.markdown('<div class="tt-cell {}"><b>{}</b><br>'
                        '<span style="color:#555">{}</span></div>'.format(
                        color, e['class'], e['subject']), unsafe_allow_html=True)


def _build_teacher_grid(grid, all_classes, days, ppd):
    tg = {}
    for cn in all_classes:
        for d in range(len(days)):
            for p in range(ppd):
                e = grid.get(cn,[[]])[d][p] if d<len(grid.get(cn,[])) else None
                if not e: continue
                et  = e.get('type','normal')
                cc2 = e.get('combined_classes',[])
                def _add(tn, tc, ts, tct):
                    if not tn: return
                    tg.setdefault(tn, [[None]*ppd for _ in range(len(days))])
                    tg[tn][d][p] = {'class':tc,'subject':ts,'is_ct':tct}
                if et=='combined_parallel':
                    if not cc2 or cn==cc2[0]: _add(e.get('teacher'),'+'.join(cc2),e.get('subject',''),False)
                    pt=e.get('par_teach','')
                    if pt and pt not in ('—','?',''): _add(pt,cn,e.get('par_subj',''),e.get('is_ct',False))
                elif et=='combined':
                    if not cc2 or cn==cc2[0]: _add(e.get('teacher'),'+'.join(cc2),e.get('subject',''),e.get('is_ct',False))
                else:
                    _add(e.get('teacher'),cn,e.get('subject',''),e.get('is_ct',False))
                    pt=e.get('par_teach','')
                    if pt and pt not in ('—','?',''): _add(pt,cn,e.get('par_subj',''),False)
    return tg


# ─────────────────────────────────────────────────────────────────────────────
# STEP 1 — School Setup  +  Drive config picker
# ─────────────────────────────────────────────────────────────────────────────
if s()['step'] == 1:

    # ── Drive saved configs panel (only if Drive is available) ────────────────
    if drive_sync.drive_available():
        loaded = drive_sync.render_drive_configs_panel()
        if loaded:
            s()['config']       = loaded.get('config', cfg())
            s()['config']['classes'] = {int(k): v for k,v in s()['config']['classes'].items()}
            s()['class_config'] = loaded.get('class_config', {})
            s()['step3']        = loaded.get('step3', {})
            s()['stage']        = 0; s()['timetable'] = None; s()['task_analysis'] = None
            meta = loaded.get('_meta', {})
            school = meta.get('school') or s()['config'].get('school_name','')
            st.success("✅ Loaded **{}** from Drive!".format(school))
            st.rerun()
        st.divider()

    st.markdown("""<div class="step-header">
        <h2>⚙️ Step 1 — School Setup</h2>
        <p>Configure basic parameters and number of class sections per grade</p>
    </div>""", unsafe_allow_html=True)

    c = cfg()
    st.markdown('<div class="card">', unsafe_allow_html=True)
    c['school_name'] = st.text_input("🏫 School Name", value=c.get('school_name',''),
        placeholder="e.g. Delhi Public School")
    col1, col2, col3 = st.columns(3)
    with col1: c['working_days']      = st.number_input("Working Days/Week", 1, 7, int(c.get('working_days',5)))
    with col2: c['periods_per_day']   = st.number_input("Periods/Day", 1, 15, int(c.get('periods_per_day',8)))
    with col3: c['periods_first_half']= st.number_input("Periods in 1st Half", 1, int(c['periods_per_day']),
        min(int(c.get('periods_first_half',4)), int(c['periods_per_day'])))
    st.markdown('</div>', unsafe_allow_html=True)

    st.markdown("### 📚 Sections per Grade")
    st.markdown('<div class="card">', unsafe_allow_html=True)
    grade_cols = st.columns(7)
    for i, cls in enumerate(range(6, 13)):
        with grade_cols[i]:
            c['classes'][cls] = st.number_input("Class {}".format(cls), 0, 10,
                int(c['classes'].get(cls,0)), key="cls_{}".format(cls))
    preview = ["{}{}".format(cls, chr(65+si))
               for cls in range(6,13) for si in range(c['classes'].get(cls,0))]
    if preview:
        st.info("**{}** sections: {}".format(len(preview), ', '.join(preview)))
    st.markdown('</div>', unsafe_allow_html=True)

    col1, col2 = st.columns(2)
    with col2:
        if st.button("Next: Class Config →", type="primary", use_container_width=True):
            if not preview:
                st.error("Add at least one class section.")
            else:
                cc = s()['class_config']
                for cn in preview:
                    if cn not in cc:
                        cc[cn] = {'teacher':'', 'teacher_period':1, 'subjects':[]}
                drive_sync.auto_save(cfg(), s()['class_config'], s()['step3'])
                go(2)


# ─────────────────────────────────────────────────────────────────────────────
# STEP 2 — Class Configuration
# ─────────────────────────────────────────────────────────────────────────────
elif s()['step'] == 2:
    st.markdown("""<div class="step-header">
        <h2>📋 Step 2 — Class Configuration</h2>
        <p>Set class teacher and subjects for each section</p>
    </div>""", unsafe_allow_html=True)

    c = cfg()
    all_classes = ["{}{}".format(cls, chr(65+si))
                   for cls in range(6,13) for si in range(c['classes'].get(cls,0))]
    cc = s()['class_config']
    for cn in all_classes:
        if cn not in cc:
            cc[cn] = {'teacher':'', 'teacher_period':1, 'subjects':[]}

    tabs = st.tabs(all_classes)
    for tab, cn in zip(tabs, all_classes):
        with tab:
            data = cc[cn]
            st.markdown('<div class="card">', unsafe_allow_html=True)
            col1, col2 = st.columns([4,1])
            with col1:
                data['teacher'] = st.text_input("Class Teacher", value=data.get('teacher',''),
                    key="ct_{}".format(cn), placeholder="e.g. Mrs. Sharma")
            with col2:
                data['teacher_period'] = st.number_input("CT Period #", 1, c['periods_per_day'],
                    int(data.get('teacher_period',1)), key="ctp_{}".format(cn))
            st.markdown('</div>', unsafe_allow_html=True)

            subjects = data.get('subjects', [])
            if subjects:
                st.markdown("**Subjects ({})**".format(len(subjects)))
                for i, sub in enumerate(subjects):
                    flags = []
                    if sub.get('consecutive')=='Yes': flags.append("↔")
                    if sub.get('parallel'):           flags.append("∥ {}".format(sub.get('parallel_subject','')))
                    if sub.get('periods_pref'):       flags.append("P{}".format(sub['periods_pref']))
                    if sub.get('days_pref'):          flags.append("D:{}".format(sub['days_pref']))
                    rc = st.columns([3,2,1,3,1])
                    rc[0].write("**{}**".format(sub['name']))
                    rc[1].write("👤 {}".format(sub['teacher']))
                    rc[2].write("{}/wk".format(sub['periods']))
                    rc[3].caption(' | '.join(flags) or '—')
                    with rc[4]:
                        if st.button("🗑️", key="del_{}_{}".format(cn,i)):
                            subjects.pop(i); data['subjects']=subjects; st.rerun()

            with st.expander("➕ Add Subject to {}".format(cn), expanded=(len(subjects)==0)):
                a1,a2,a3 = st.columns([3,3,1])
                with a1: ns_name    = st.text_input("Subject Name", key="ns_n_{}".format(cn), placeholder="Mathematics")
                with a2: ns_teacher = st.text_input("Teacher",      key="ns_t_{}".format(cn), placeholder="Mr. Gupta")
                with a3: ns_per     = st.number_input("Pd/Wk", 1, 30, 4, key="ns_p_{}".format(cn))
                b1,b2 = st.columns(2)
                with b1: ns_consec = st.checkbox("Consecutive", key="ns_c_{}".format(cn))
                with b2: ns_par    = st.checkbox("Has Parallel", key="ns_par_{}".format(cn))
                ns_ps = ns_pt = ''
                if ns_par:
                    p1,p2 = st.columns(2)
                    with p1: ns_ps = st.text_input("Parallel Subject", key="ns_ps_{}".format(cn))
                    with p2: ns_pt = st.text_input("Parallel Teacher", key="ns_pt_{}".format(cn))
                with st.expander("⚙️ Period / Day preferences"):
                    pp = st.multiselect("Preferred Periods", list(range(1,c['periods_per_day']+1)), key="pp_{}".format(cn))
                    pd = st.multiselect("Preferred Days",    DAYS_ALL[:c['working_days']],           key="pd_{}".format(cn))
                if st.button("✅ Add Subject", key="add_{}".format(cn), type="primary"):
                    if ns_name.strip() and ns_teacher.strip():
                        subjects.append({'name':ns_name.strip(),'teacher':ns_teacher.strip(),
                            'periods':int(ns_per),'consecutive':'Yes' if ns_consec else 'No',
                            'parallel':ns_par,'parallel_subject':ns_ps.strip(),
                            'parallel_teacher':ns_pt.strip(),'periods_pref':pp,'days_pref':pd})
                        data['subjects']=subjects; st.rerun()
                    else: st.error("Name and teacher required.")

    st.divider()
    col1, col2 = st.columns(2)
    with col1:
        if st.button("← Back", use_container_width=True): go(1)
    with col2:
        if st.button("Next: Combines & Rules →", type="primary", use_container_width=True):
            drive_sync.auto_save(cfg(), s()['class_config'], s()['step3'])
            go(3)


# ─────────────────────────────────────────────────────────────────────────────
# STEP 3 — Combines & Unavailability
# ─────────────────────────────────────────────────────────────────────────────
elif s()['step'] == 3:
    st.markdown("""<div class="step-header">
        <h2>🔗 Step 3 — Combines & Unavailability</h2>
        <p>Set combined class groups and teacher unavailability windows</p>
    </div>""", unsafe_allow_html=True)

    c   = cfg()
    s3  = s()['step3']
    all_classes = ["{}{}".format(cls, chr(65+si))
                   for cls in range(6,13) for si in range(c['classes'].get(cls,0))]

    # Collect teachers
    teachers = set()
    for cn, data in s()['class_config'].items():
        if data.get('teacher','').strip(): teachers.add(data['teacher'].strip())
        for sub in data.get('subjects',[]):
            if sub.get('teacher','').strip():           teachers.add(sub['teacher'].strip())
            if sub.get('parallel_teacher','').strip():  teachers.add(sub['parallel_teacher'].strip())
    teachers = sorted(teachers)

    t1, t2 = st.tabs(["🔗 Combined Classes", "🚫 Teacher Unavailability"])

    with t1:
        if not teachers:
            st.info("Complete Step 2 first to see teachers here.")
        else:
            for teacher in teachers:
                td = s3.setdefault(teacher, {'combines':[], 'unavailability':{}})
                with st.expander("👤 {} — {} combine(s)".format(teacher, len(td.get('combines',[])))):
                    for i, cb in enumerate(td.get('combines',[])):
                        st.markdown("**Combine {}**".format(i+1))
                        cc1,cc2,cc3 = st.columns([3,3,1])
                        with cc1: cb['classes']  = st.multiselect("Classes", all_classes,
                            default=cb.get('classes',[]), key="cb_cls_{}_{}".format(teacher,i))
                        with cc2:
                            raw = st.text_input("Subject(s)", value=', '.join(cb.get('subjects',[])),
                                key="cb_sub_{}_{}".format(teacher,i), help="Comma-separated")
                            cb['subjects'] = [x.strip() for x in raw.split(',') if x.strip()]
                        with cc3:
                            st.markdown("<br>", unsafe_allow_html=True)
                            if st.button("🗑️", key="del_cb_{}_{}".format(teacher,i)):
                                td['combines'].pop(i); st.rerun()
                    if st.button("➕ Add combine for {}".format(teacher), key="addc_{}".format(teacher)):
                        td['combines'].append({'classes':[],'subjects':[]}); st.rerun()

    with t2:
        if not teachers:
            st.info("Complete Step 2 first.")
        else:
            for teacher in teachers:
                td = s3.setdefault(teacher, {'combines':[], 'unavailability':{}})
                u  = td.setdefault('unavailability', {})
                with st.expander("👤 {}".format(teacher)):
                    uc1, uc2 = st.columns(2)
                    with uc1: u['days']    = st.multiselect("Unavailable Days",
                        DAYS_ALL[:c['working_days']], default=u.get('days',[]),
                        key="ud_{}".format(teacher))
                    with uc2: u['periods'] = st.multiselect("Unavailable Periods",
                        list(range(1, c['periods_per_day']+1)), default=u.get('periods',[]),
                        key="up_{}".format(teacher))
                    if u.get('days') and u.get('periods'):
                        st.caption("Will not be scheduled on {} in period(s) {}".format(
                            ', '.join(u['days']), u['periods']))

    st.divider()
    col1, col2 = st.columns(2)
    with col1:
        if st.button("← Back", use_container_width=True): go(2)
    with col2:
        if st.button("Next: Generate →", type="primary", use_container_width=True):
            drive_sync.auto_save(cfg(), s()['class_config'], s()['step3'])
            go(4)


# ─────────────────────────────────────────────────────────────────────────────
# STEP 4 — Generate
# ─────────────────────────────────────────────────────────────────────────────
elif s()['step'] == 4:
    st.markdown("""<div class="step-header">
        <h2>🚀 Step 4 — Generate Timetable</h2>
        <p>Run stages progressively to build a conflict-free timetable</p>
    </div>""", unsafe_allow_html=True)

    def _build_eng():
        unavail = {t: d['unavailability'] for t,d in s()['step3'].items()
                   if d.get('unavailability',{}).get('days') or d.get('unavailability',{}).get('periods')}
        eng = TimetableEngine(
            configuration    = cfg(),
            class_config_data= s()['class_config'],
            step3_data       = {t:{'combines':d.get('combines',[])} for t,d in s()['step3'].items()},
            step3_unavailability = unavail,
        )
        eng.init_gen_state()
        return eng

    stage = s()['stage']
    labels = ["Not started","Stage 1 ✓","Task Analysis ✓","Stage 2 ✓"]
    st.progress([0,33,55,100][min(stage,3)] / 100,
                text="Progress: {}".format(labels[min(stage,3)]))

    # Stage 1
    with st.expander("**Stage 1 — Fixed & Preference Periods**", expanded=(stage==0)):
        if stage == 0:
            st.info("Places CT periods and preference-constrained subjects.")
            if st.button("▶ Run Stage 1", type="primary"):
                with st.spinner("Running Stage 1…"):
                    try:
                        eng = _build_eng()
                        res = eng.run_stage1()
                        s()['engine']=eng; s()['stage']=1
                        s()['stage1_result']=res; s()['timetable']=eng.get_timetable()
                        st.rerun()
                    except Exception as e:
                        import traceback; traceback.print_exc()
                        st.error("Stage 1 error: {}".format(e))
        else:
            res    = s()['stage1_result'] or {}
            issues = res.get('issues',[])
            rem    = res.get('remaining',0)
            if issues: st.warning("{} issue(s). {} periods to Stage 2.".format(len(issues),rem))
            else:       st.success("✅ Stage 1 done — {} periods remaining.".format(rem))
            for iss in issues: st.error(iss)
            if st.button("🔄 Re-run Stage 1"): s()['stage']=0; s()['task_analysis']=None; st.rerun()

    # Task Analysis
    if stage >= 1:
        with st.expander("**Task Analysis — Combined & Parallel Groups**", expanded=(stage==1)):
            if s()['task_analysis'] is None:
                st.info("Allocates slots for combined/parallel/consecutive groups.")
                if st.button("📋 Run Task Analysis", type="primary"):
                    with st.spinner("Analysing groups…"):
                        gslots, alloc, rows = s()['engine'].run_task_analysis_allocation()
                        s()['task_analysis']=(gslots,alloc,rows)
                        s()['timetable']=s()['engine'].get_timetable()
                        s()['stage']=2; st.rerun()
            else:
                gslots, alloc, rows = s()['task_analysis']
                if not rows:
                    st.info("No combined/parallel/consecutive groups.")
                else:
                    hc = st.columns([1,1,2,2,2,2,2])
                    for h,t in zip(hc,["Grp","Sec","Class","Subject","Teacher","Par.","Status"]):
                        h.markdown("**{}**".format(t))
                    for row in rows:
                        gn=row['group']; al=alloc.get(gn,{})
                        ok=al.get('ok',False)
                        placed=al.get('s1_placed',0)+al.get('new_placed',0)
                        tot=al.get('total',0)
                        badge='<span class="badge-ok">✓ {}/{}</span>'.format(placed,tot) if ok \
                              else '<span class="badge-err">✗ {}/{}</span>'.format(placed,tot)
                        rc=st.columns([1,1,2,2,2,2,2])
                        rc[0].write(gn); rc[1].write(row['section'])
                        rc[2].write(row['class']); rc[3].write(row['subject'])
                        rc[4].write(row['teacher']); rc[5].write(row.get('par_subj','—'))
                        rc[6].markdown(badge, unsafe_allow_html=True)
                        if not ok and al.get('reason'): st.caption("⚠ {}".format(al['reason']))
                if st.button("🔄 Re-run Task Analysis"): s()['task_analysis']=None; s()['stage']=1; st.rerun()

    # Stage 2
    if stage >= 2:
        with st.expander("**Stage 2 — Fill Remaining Periods**", expanded=(stage==2)):
            if stage == 2:
                st.info("Fills all remaining periods with SC1/SC2/filler + repair loop.")
                if st.button("▶ Run Stage 2", type="primary"):
                    with st.spinner("Running Stage 2 (may take ~15s for large schools)…"):
                        unp = s()['engine'].run_stage2()
                        s()['unplaced']=unp; s()['timetable']=s()['engine'].get_timetable()
                        s()['stage']=3; st.rerun()
            else:
                unp = s()['unplaced']
                if unp==0: st.success("✅ Stage 2 complete — fully filled!")
                else:       st.warning("{} period(s) unplaced — try Force Fill.".format(unp))
                if st.button("🔄 Re-run Stage 2"): s()['stage']=2; st.rerun()

    # Force Fill
    if stage >= 3 and s()['unplaced'] > 0:
        with st.expander("**⚡ Force Fill — Min-Conflicts Solver**", expanded=True):
            st.warning("{} period(s) unplaced. Force Fill relaxes soft constraints progressively.".format(s()['unplaced']))
            if st.button("⚡ Run Force Fill", type="primary"):
                ph = st.empty()
                def _cb(msg):
                    if msg: ph.info("⏳ {}".format(msg))
                with st.spinner("Force filling…"):
                    notes = s()['engine'].force_fill(progress_cb=_cb)
                    unp   = s()['engine'].get_timetable()['unplaced']
                    s()['unplaced']=unp; s()['timetable']=s()['engine'].get_timetable()
                    s()['force_fill_notes']=notes
                ph.empty(); st.rerun()
            if s().get('force_fill_notes'):
                with st.expander("Solver notes"): st.text(s()['force_fill_notes'])

    st.divider()
    col1, col2 = st.columns(2)
    with col1:
        if st.button("← Back to Step 3", use_container_width=True): go(3)
    with col2:
        if stage >= 3:
            if st.button("📊 View Final Timetable →", type="primary", use_container_width=True): go(5)
        else:
            st.button("📊 Timetable (complete stages first)", disabled=True, use_container_width=True)


# ─────────────────────────────────────────────────────────────────────────────
# STEP 5 — Final Timetable & Downloads
# ─────────────────────────────────────────────────────────────────────────────
elif s()['step'] == 5:
    st.markdown("""<div class="step-header">
        <h2>📊 Final Timetable</h2>
        <p>View, download and share your generated timetable</p>
    </div>""", unsafe_allow_html=True)

    tt  = s()['timetable']
    eng = s()['engine']
    if tt is None or eng is None:
        st.warning("No timetable generated — complete Step 4 first.")
        if st.button("← Go to Generate"): go(4)
        st.stop()

    days=tt['days']; ppd=tt['ppd']; half1=tt['half1']
    grid=tt['grid']; all_classes=tt['all_classes']
    total_p = sum(t['periods'] for t in tt['tasks'])
    unp     = tt.get('unplaced', 0)

    # Metrics
    m1,m2,m3,m4 = st.columns(4)
    m1.metric("Classes",    len(all_classes))
    m2.metric("Total Periods", total_p)
    m3.metric("Placed",     total_p - unp)
    m4.metric("Unplaced",   unp, delta=("⚠ needs Force Fill" if unp else "✓ complete"))

    if unp > 0:
        st.warning("⚠ {} period(s) unplaced. Return to Step 4 → Force Fill.".format(unp))
    else:
        st.success("✅ Timetable fully generated — all {} periods placed.".format(total_p))

    # Auto-save final timetable config to Drive
    drive_sync.auto_save(cfg(), s()['class_config'], s()['step3'])

    # ── Downloads ─────────────────────────────────────────────────────────────
    st.markdown("### 📥 Download Excel")
    st.markdown('<div class="card">', unsafe_allow_html=True)
    school = cfg().get('school_name','school').replace(' ','_').lower()[:16]
    modes  = [("class","📚 Classwise"),("teacher","👩‍🏫 Teacherwise"),
              ("ct_list","📋 CT List"),("workload","📊 Workload"),("one_sheet","📄 One Sheet")]
    dl_cols = st.columns(5)
    for col, (mode, label) in zip(dl_cols, modes):
        with col:
            try:
                buf = eng.export_excel(mode, tt)
                st.download_button(label=label,
                    data=buf.getvalue(),
                    file_name="{}_{}.xlsx".format(school, mode),
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True, key="dl_{}".format(mode))
            except Exception as e:
                st.button("{} ❌".format(label), disabled=True,
                          use_container_width=True, key="dl_err_{}".format(mode))
                st.caption(str(e)[:50])
    st.markdown('</div>', unsafe_allow_html=True)

    # ── Viewer ────────────────────────────────────────────────────────────────
    st.markdown("### 👀 View Timetable")
    view_mode = st.radio("View by:", ["📚 Class","👩‍🏫 Teacher"], horizontal=True)

    if view_mode == "📚 Class":
        sel = st.selectbox("Class:", all_classes)
        st.markdown('<div class="card">', unsafe_allow_html=True)
        _render_class_tt(sel, grid, days, ppd, half1, eng)
        st.markdown('</div>', unsafe_allow_html=True)
    else:
        tg      = _build_teacher_grid(grid, all_classes, days, ppd)
        teachers= sorted(tg.keys())
        if teachers:
            sel_t = st.selectbox("Teacher:", teachers)
            st.markdown('<div class="card">', unsafe_allow_html=True)
            _render_teacher_tt(tg[sel_t], days, ppd, half1)
            st.markdown('</div>', unsafe_allow_html=True)
        else:
            st.info("No teacher data available.")

    if st.button("← Back to Generate", use_container_width=True): go(4)
