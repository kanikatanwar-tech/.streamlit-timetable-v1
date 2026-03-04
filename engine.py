"""Timetable Engine — pure Python, no tkinter."""
import random, io
from collections import defaultdict

class TimetableEngine:
    def __init__(self, configuration, class_config_data, step3_data=None, step3_unavailability=None):
        self.configuration        = configuration
        self.class_config_data    = class_config_data
        self.step3_data           = step3_data or {}
        self.step3_unavailability = step3_unavailability or {}
        self._relaxed_consec_keys = set()
        self._relaxed_main_keys   = set()
        self._gen                 = None

    @staticmethod
    def _sv(val):
        if hasattr(val, 'get'): return val.get()
        return val or ''

    def _all_classes(self):
        cfg = self.configuration
        result = []
        for cls in range(6, 13):
            for si in range(cfg['classes'].get(cls, 0)):
                result.append("{}{}".format(cls, chr(65+si)))
        return result

    def init_gen_state(self):
        cfg   = self.configuration
        ppd   = cfg['periods_per_day']
        wdays = cfg['working_days']
        half1 = cfg['periods_first_half']
        DAYS  = ['Mon','Tue','Wed','Thu','Fri','Sat','Sun'][:wdays]
        all_classes = self._all_classes()

        grid    = {cn: [[None]*ppd for _ in range(wdays)] for cn in all_classes}
        task_at = {cn: [[None]*ppd for _ in range(wdays)] for cn in all_classes}
        t_busy  = {}

        def t_free(t,d,p): return not t or (d,p) not in t_busy.get(t,set())
        def t_mark(t,d,p):
            if t: t_busy.setdefault(t,set()).add((d,p))
        def t_unmark(t,d,p):
            if t: t_busy.get(t,set()).discard((d,p))
        unavail = self.step3_unavailability
        def t_unavail(t,d,p):
            u = unavail.get(t,{})
            return bool(u) and DAYS[d] in u.get('days',[]) and (p+1) in u.get('periods',[])

        s3 = self.step3_data
        cn_subj_combined = {}
        for _teacher, s3d in s3.items():
            for cb in s3d.get('combines',[]):
                classes = sorted(cb.get('classes',[]))
                subjects = cb.get('subjects',[])
                if len(classes)>=2 and subjects:
                    for cn in classes:
                        cn_subj_combined[(cn, subjects[0])] = classes
        for cn in all_classes:
            for s in self.class_config_data.get(cn,{}).get('subjects',[]):
                pri = s.get('name','').strip()
                par = (s.get('parallel_subject') or '').strip()
                if par and (cn,par) in cn_subj_combined and (cn,pri) not in cn_subj_combined:
                    cn_subj_combined[(cn,pri)] = cn_subj_combined[(cn,par)]
                if (cn,pri) in cn_subj_combined and par and (cn,par) not in cn_subj_combined:
                    cn_subj_combined[(cn,par)] = cn_subj_combined[(cn,pri)]

        tasks = []
        seen_combined = set()
        for cn in all_classes:
            cd = self.class_config_data.get(cn)
            if not cd: continue
            ct = self._sv(cd.get('teacher','')).strip()
            try: ct_per = int(self._sv(cd.get('teacher_period',1)))
            except: ct_per = 1
            ct_assigned = False

            for s in cd.get('subjects',[]):
                subj = s.get('name','')
                t    = s.get('teacher','').strip()
                try: n = int(s.get('periods',0))
                except: n = 0
                if n <= 0: continue

                cn_list = cn_subj_combined.get((cn,subj),[cn])
                if len(cn_list)>1:
                    key = (frozenset(cn_list), subj)
                    if key in seen_combined: continue
                    seen_combined.add(key)

                is_ct = (t==ct and not ct_assigned)
                if is_ct: ct_assigned = True

                par    = bool(s.get('parallel',False))
                pt     = s.get('parallel_teacher','').strip() if par else ''
                ps     = s.get('parallel_subject','').strip() if par else ''
                consec = (s.get('consecutive','No')=='Yes')
                if consec and (cn,subj) in self._relaxed_consec_keys: consec=False
                p_pref = list(s.get('periods_pref',[]))
                d_pref = list(s.get('days_pref',[]))

                if len(cn_list)>1 and par:   ttype='combined_parallel'
                elif len(cn_list)>1:          ttype='combined'
                elif par:                     ttype='parallel'
                else:                         ttype='normal'

                if is_ct:          priority='HC1'
                elif p_pref or d_pref: priority='HC2'
                elif consec:       priority='SC1'
                elif n>=wdays:     priority='SC2'
                else:              priority='filler'

                tasks.append({'idx':len(tasks),'cn_list':cn_list,'subject':subj,'teacher':t,
                    'par_subj':ps,'par_teach':pt,'periods':n,'remaining':n,'is_ct':is_ct,
                    'ct_period':ct_per if is_ct else None,'p_pref':p_pref,'d_pref':d_pref,
                    'consec':consec,'daily':(n>=wdays),'priority':priority,'type':ttype,
                    'rx_sc1':False,'rx_sc3':False,'rx_sc2':False})

        self._gen = {'cfg':cfg,'ppd':ppd,'wdays':wdays,'half1':half1,'DAYS':DAYS,
            'all_classes':all_classes,'grid':grid,'task_at':task_at,'t_busy':t_busy,
            'tasks':tasks,'total_atoms':sum(t['periods'] for t in tasks),
            't_free':t_free,'t_mark':t_mark,'t_unmark':t_unmark,'t_unavail':t_unavail}

    def _can(self, task, d, p, ign_sc1=False, ign_sc3=False):
        g=self._gen; DAYS=g['DAYS']; ppd=g['ppd']; grid=g['grid']
        t=task['teacher']; pt=task['par_teach']; p1=p+1
        if task['is_ct'] and p1!=task['ct_period']: return False
        if task['p_pref'] and not task['is_ct'] and p1 not in task['p_pref']: return False
        if task['d_pref'] and DAYS[d] not in task['d_pref']: return False
        for cn in task['cn_list']:
            if grid[cn][d][p] is not None: return False
        if not g['t_free'](t,d,p): return False
        if pt and not g['t_free'](pt,d,p): return False
        if not (ign_sc3 or task['rx_sc3']):
            if g['t_unavail'](t,d,p): return False
            if pt and g['t_unavail'](pt,d,p): return False
        if task['consec'] and not (ign_sc1 or task['rx_sc1']):
            if p>=ppd-1: return False
            for cn in task['cn_list']:
                if grid[cn][d][p+1] is not None: return False
            if not g['t_free'](t,d,p+1): return False
            if pt and not g['t_free'](pt,d,p+1): return False
        if not task['consec']:
            for cn in task['cn_list']:
                for pp in range(ppd):
                    e=grid[cn][d][pp]
                    if e and e.get('subject')==task['subject']: return False
        return True

    def _make_cell(self, task):
        return {'type':task['type'],'subject':task['subject'],'teacher':task['teacher'],
            'par_subj':task['par_subj'],'par_teach':task['par_teach'],
            'combined_classes':task['cn_list'] if len(task['cn_list'])>1 else [],
            'is_ct':task['is_ct']}

    def _place(self, task, d, p):
        g=self._gen; cell=self._make_cell(task)
        for cn in task['cn_list']:
            g['grid'][cn][d][p]=cell; g['task_at'][cn][d][p]=task['idx']
        g['t_mark'](task['teacher'],d,p)
        if task['par_teach']: g['t_mark'](task['par_teach'],d,p)
        task['remaining']-=1

    def _unplace(self, task, d, p):
        g=self._gen
        for cn in task['cn_list']:
            g['grid'][cn][d][p]=None; g['task_at'][cn][d][p]=None
        g['t_unmark'](task['teacher'],d,p)
        if task['par_teach']: g['t_unmark'](task['par_teach'],d,p)
        task['remaining']+=1

    def get_timetable(self):
        g=self._gen
        return {'grid':g['grid'],'days':g['DAYS'],'ppd':g['ppd'],'half1':g['half1'],
            'all_classes':g['all_classes'],'tasks':g['tasks'],
            'unplaced':sum(t['remaining'] for t in g['tasks'])}

    # ── Stage 1 ──────────────────────────────────────────────────────────────
    def run_stage1(self):
        g=self._gen; tasks=g['tasks']; grid=g['grid']
        wdays=g['wdays']; ppd=g['ppd']; DAYS=g['DAYS']
        issues=[]
        for task in tasks:
            if task['priority']!='HC1': continue
            p_idx=task['ct_period']-1
            for d in range(wdays):
                if task['remaining']<=0: break
                blocked=any(grid[cn][d][p_idx] is not None for cn in task['cn_list'])
                if not blocked: self._place(task,d,p_idx)
                else: issues.append("HC1 blocked: {} {} {} {}".format(task['subject'],task['teacher'],DAYS[d],p_idx+1))
        hc2=sorted([t for t in tasks if t['priority']=='HC2'],
            key=lambda t: (len(t['p_pref']) or ppd)*(len(t['d_pref']) or wdays))
        for task in hc2:
            if task['remaining']<=0: continue
            pref_p=[x-1 for x in task['p_pref']] if task['p_pref'] else list(range(ppd))
            pref_d=([DAYS.index(x) for x in task['d_pref'] if x in DAYS] if task['d_pref'] else list(range(wdays)))
            for d,p in [(d,p) for d in pref_d for p in pref_p]:
                if task['remaining']<=0: break
                if not any(grid[cn][d][p] is not None for cn in task['cn_list']):
                    self._place(task,d,p)
        other_rem=sum(t['remaining'] for t in tasks if t['priority'] not in ('HC1','HC2'))
        return {'issues':issues,'remaining':other_rem}

    # ── Stage 2 ──────────────────────────────────────────────────────────────
    def run_stage2(self):
        g=self._gen; tasks=g['tasks']; wdays=g['wdays']; ppd=g['ppd']; grid=g['grid']
        PRIO={'HC1':0,'HC2':1,'SC1':2,'SC2':3,'filler':4}
        # SC1
        for task in sorted([t for t in tasks if t['priority']=='SC1'],key=lambda t:-t['periods']):
            if task['remaining']<=0: continue
            days=list(range(wdays)); random.shuffle(days)
            for d in days:
                if task['remaining']<=0: break
                for p in range(ppd-1):
                    if task['remaining']<=0: break
                    if self._can(task,d,p):
                        self._place(task,d,p)
                        if task['remaining']>0 and self._can(task,d,p+1): self._place(task,d,p+1)
                        break
        # SC2
        for task in sorted([t for t in tasks if t['priority']=='SC2' and t['remaining']>0],key=lambda t:-t['periods']):
            placed=False
            for p in range(ppd):
                avail=[d for d in range(wdays) if self._can(task,d,p)]
                if len(avail)>=task['remaining']:
                    for d in avail[:task['remaining']]: self._place(task,d,p)
                    placed=True; break
            if not placed:
                for d in range(wdays):
                    if task['remaining']<=0: break
                    for p in range(ppd):
                        if task['remaining']<=0: break
                        if self._can(task,d,p): self._place(task,d,p); break
        # Fillers
        for task in sorted([t for t in tasks if t['remaining']>0],key=lambda t:-t['periods']):
            for d in range(wdays):
                if task['remaining']<=0: break
                for p in range(ppd):
                    if task['remaining']<=0: break
                    if self._can(task,d,p): self._place(task,d,p)
        # Repair
        relax=0
        for _ in range(80):
            rem=[t for t in tasks if t['remaining']>0]
            if not rem: break
            if relax>=1:
                for t in tasks: t['rx_sc1']=True
            if relax>=2:
                for t in tasks: t['rx_sc3']=True
            progress=False
            for task in sorted(rem,key=lambda t:-t['remaining']):
                if task['remaining']<=0: continue
                pt=task['par_teach']
                for d in range(wdays):
                    if task['remaining']<=0: break
                    for p in range(ppd):
                        if task['remaining']<=0: break
                        if self._can(task,d,p,relax>=1,relax>=2): self._place(task,d,p); progress=True
                if task['remaining']>0:
                    tprio=PRIO.get(task['priority'],4)
                    for d in range(wdays):
                        if task['remaining']<=0: break
                        for p in range(ppd):
                            if task['remaining']<=0: break
                            t_ok=g['t_free'](task['teacher'],d,p)
                            if pt and pt not in ('','—','?'): t_ok=t_ok and g['t_free'](pt,d,p)
                            if not t_ok: continue
                            bidx=None
                            for cn in task['cn_list']:
                                if grid[cn][d][p] is not None: bidx=g['task_at'][cn][d][p]; break
                            if bidx is None: continue
                            blocker=tasks[bidx]; bprio=PRIO.get(blocker['priority'],4)
                            if bprio<=tprio: continue
                            for d2 in range(wdays):
                                moved=False
                                for p2 in range(ppd):
                                    if (d2,p2)==(d,p): continue
                                    if not self._can(blocker,d2,p2,relax>=1,relax>=2): continue
                                    self._unplace(blocker,d,p)
                                    clr=all(grid[cn][d][p] is None for cn in task['cn_list'])
                                    if clr and g['t_free'](task['teacher'],d,p):
                                        self._place(blocker,d2,p2)
                                        if self._can(task,d,p,relax>=1,relax>=2):
                                            self._place(task,d,p); moved=True; progress=True; break
                                        else: self._unplace(blocker,d2,p2); self._place(blocker,d,p)
                                    else: self._place(blocker,d,p)
                                if moved: break
            if not progress: relax=min(relax+1,4)
        return sum(t['remaining'] for t in tasks)

    # ── Force Fill ───────────────────────────────────────────────────────────
    def force_fill(self, progress_cb=None):
        def _prog(msg):
            if progress_cb: progress_cb(msg)
        g=self._gen; tasks=g['tasks']; grid=g['grid']; wdays=g['wdays']; ppd=g['ppd']
        PRIO={'HC1':0,'HC2':1,'SC1':2,'SC2':3,'filler':4}
        def _prio(t): return PRIO.get(t['priority'],4)
        def _unplaced(): return sum(t['remaining'] for t in tasks)

        def _greedy(is1=False,is3=False):
            rem=sorted([t for t in tasks if t['remaining']>0],
                key=lambda t: sum(1 for d in range(wdays) for p in range(ppd) if self._can(t,d,p,is1,is3)))
            for task in rem:
                for d in range(wdays):
                    if task['remaining']<=0: break
                    for p in range(ppd):
                        if task['remaining']<=0: break
                        if self._can(task,d,p,is1,is3): self._place(task,d,p)

        def _swap(is1=False,is3=False):
            for task in sorted(tasks,key=lambda t:-t['remaining']):
                if task['remaining']<=0 or _prio(task)==0: continue
                for d in range(wdays):
                    if task['remaining']<=0: break
                    for p in range(ppd):
                        if task['remaining']<=0: break
                        tname=task['teacher']; pt=task.get('par_teach','')
                        tok=g['t_free'](tname,d,p)
                        if not is3: tok=tok and not g['t_unavail'](tname,d,p)
                        if pt and pt not in ('','—','?'): tok=tok and g['t_free'](pt,d,p)
                        if not tok: continue
                        bidx=None
                        for cn in task['cn_list']:
                            if grid[cn][d][p] is not None: bidx=g['task_at'][cn][d][p]; break
                        if bidx is None:
                            if self._can(task,d,p,is1,is3): self._place(task,d,p)
                            continue
                        blocker=tasks[bidx]
                        if _prio(blocker)<=_prio(task): continue
                        for d2 in range(wdays):
                            moved=False
                            for p2 in range(ppd):
                                if (d2,p2)==(d,p): continue
                                if not self._can(blocker,d2,p2,is1,is3): continue
                                self._unplace(blocker,d,p)
                                clr=all(grid[cn][d][p] is None for cn in task['cn_list'])
                                if clr and g['t_free'](tname,d,p):
                                    self._place(blocker,d2,p2)
                                    if self._can(task,d,p,is1,is3): self._place(task,d,p); moved=True; break
                                    else: self._unplace(blocker,d2,p2); self._place(blocker,d,p)
                                else: self._place(blocker,d,p)
                            if moved: break

        def _run_a(is1=False,is3=False):
            for _ in range(4):
                if _unplaced()==0: return
                _greedy(is1,is3)
            for _ in range(4):
                if _unplaced()==0: return
                _swap(is1,is3); _greedy(is1,is3)

        notes=[]
        _prog("Stage A — greedy…"); _run_a()
        if _unplaced()==0: return None
        _prog("Stage A — relax consecutive…")
        for t in tasks:
            if t['consec'] and t['remaining']>0:
                t['rx_sc1']=True
                for cn in t['cn_list']: self._relaxed_consec_keys.add((cn,t['subject']))
        _run_a(is1=True)
        if _unplaced()==0: return '\n'.join(notes)
        _prog("Stage A — relax unavailability…")
        for t in tasks:
            if t['remaining']>0: t['rx_sc3']=True
        _run_a(is1=True,is3=True)
        if _unplaced()==0: return '\n'.join(notes)
        _prog("Stage A — relax preferences…")
        for t in tasks:
            if t['remaining']==0 or t.get('is_ct'): continue
            if t['p_pref'] or t['d_pref'] or t.get('daily') or t['priority']=='SC2':
                t['p_pref']=[]; t['d_pref']=[]; t['daily']=False; t['priority']='filler'
        _run_a(is1=True,is3=True)
        if _unplaced()==0: return '\n'.join(notes)

        # Stage B: Min-Conflicts
        notes.append("Min-Conflicts solver applied.")
        _prog("Stage B — force-complete…")
        for t in tasks: t['rx_sc1']=True; t['rx_sc3']=True
        for task in sorted(tasks,key=lambda t:-t['remaining']):
            if task['remaining']<=0: continue
            for d in range(wdays):
                if task['remaining']<=0: break
                for p in range(ppd):
                    if task['remaining']<=0: break
                    if not all(grid[cn][d][p] is None for cn in task['cn_list']): continue
                    if any(g['task_at'][cn][d][p] is not None and tasks[g['task_at'][cn][d][p]]['priority']=='HC1' for cn in task['cn_list']): continue
                    for cn in task['cn_list']:
                        grid[cn][d][p]=self._make_cell(task); g['task_at'][cn][d][p]=task['idx']
                    g['t_mark'](task['teacher'],d,p)
                    pt=task.get('par_teach','')
                    if pt and pt not in ('','—','?'): g['t_mark'](pt,d,p)
                    task['remaining']-=1
        if _unplaced()>0: return '\n'.join(notes)

        def _conflicts(tname,pt,d,p,own):
            s=0
            for cn2 in g['all_classes']:
                idx2=g['task_at'][cn2][d][p]
                if idx2 is None or idx2==own: continue
                other=tasks[idx2]
                if other['teacher']==tname: s+=1
                if pt and pt not in ('','—','?') and (other['teacher']==pt or other.get('par_teach','')==pt): s+=1
            return s

        def _task_slots():
            ts={t['idx']:[] for t in tasks}
            for cn in g['all_classes']:
                for d in range(wdays):
                    for p in range(ppd):
                        idx=g['task_at'][cn][d][p]
                        if idx is not None and (d,p) not in ts[idx]: ts[idx].append((d,p))
            return ts

        best=None; no_imp=0
        for _iter in range(1500):
            ts=_task_slots()
            total=sum(_conflicts(t['teacher'],t.get('par_teach',''),d,p,t['idx'])
                for t in tasks if t['priority']!='HC1' for d,p in ts[t['idx']])
            if _iter%20==0: _prog("Stage B — conflicts:{} iter:{}/1500".format(total,_iter))
            if total==0: break
            if best is None or total<best: best=total; no_imp=0
            else: no_imp+=1
            if no_imp>=150:
                non_hc1=[t for t in tasks if t['priority']!='HC1']; random.shuffle(non_hc1)
                for t in non_hc1:
                    for d,p in ts[t['idx']]:
                        for cn in t['cn_list']: grid[cn][d][p]=None; g['task_at'][cn][d][p]=None
                        g['t_unmark'](t['teacher'],d,p)
                        pt2=t.get('par_teach','')
                        if pt2 and pt2 not in ('','—','?'): g['t_unmark'](pt2,d,p)
                        t['remaining']+=1
                    free=[(d,p) for d in range(wdays) for p in range(ppd) if all(grid[cn][d][p] is None for cn in t['cn_list'])]
                    random.shuffle(free)
                    for d,p in free:
                        if t['remaining']<=0: break
                        for cn in t['cn_list']: grid[cn][d][p]=self._make_cell(t); g['task_at'][cn][d][p]=t['idx']
                        g['t_mark'](t['teacher'],d,p)
                        pt2=t.get('par_teach','')
                        if pt2 and pt2 not in ('','—','?'): g['t_mark'](pt2,d,p)
                        t['remaining']-=1
                no_imp=0; best=None; continue
            conflicted=[(t,sum(_conflicts(t['teacher'],t.get('par_teach',''),d,p,t['idx']) for d,p in ts[t['idx']]))
                for t in tasks if t['priority']!='HC1' and
                sum(_conflicts(t['teacher'],t.get('par_teach',''),d,p,t['idx']) for d,p in ts[t['idx']])>0]
            if not conflicted: break
            target,_=max(conflicted,key=lambda x:x[1])
            t_sl=ts[target['idx']]
            if not t_sl: continue
            wd,wp=max(t_sl,key=lambda dp:_conflicts(target['teacher'],target.get('par_teach',''),dp[0],dp[1],target['idx']))
            for cn in target['cn_list']: grid[cn][wd][wp]=None; g['task_at'][cn][wd][wp]=None
            g['t_unmark'](target['teacher'],wd,wp)
            pt=target.get('par_teach','')
            if pt and pt not in ('','—','?'): g['t_unmark'](pt,wd,wp)
            target['remaining']+=1
            best_s=None; bd,bp=wd,wp
            for d in range(wdays):
                for p in range(ppd):
                    if not all(grid[cn][d][p] is None for cn in target['cn_list']): continue
                    sc=_conflicts(target['teacher'],pt,d,p,target['idx'])
                    if best_s is None or sc<best_s: best_s=sc; bd,bp=d,p
                    if best_s==0: break
                if best_s==0: break
            for cn in target['cn_list']: grid[cn][bd][bp]=self._make_cell(target); g['task_at'][cn][bd][bp]=target['idx']
            g['t_mark'](target['teacher'],bd,bp)
            if pt and pt not in ('','—','?'): g['t_mark'](pt,bd,bp)
            target['remaining']-=1
        _prog("")
        return '\n'.join(notes) if notes else None

    # ── Task Analysis ─────────────────────────────────────────────────────────
    def _find_parallel(self, cn, subject_name):
        for s in self.class_config_data.get(cn,{}).get('subjects',[]):
            if s['name']==subject_name and s.get('parallel'):
                return ((s.get('parallel_subject') or '?').strip(), (s.get('parallel_teacher') or '—').strip())
            if (s.get('parallel') and (s.get('parallel_subject') or '').strip()==subject_name and s['name']!=subject_name):
                return (s['name'], (s.get('teacher') or '—').strip())
        return ('—','—')

    def run_task_analysis_allocation(self):
        all_classes=self._all_classes(); s3=self.step3_data
        all_rows=[]; group_no=0; covered=set()
        for teacher,s3d in sorted(s3.items()):
            for cb in s3d.get('combines',[]):
                classes=cb.get('classes',[]); subjects=cb.get('subjects',[])
                if not classes: continue
                group_no+=1
                for j,cn in enumerate(classes):
                    tsub=subjects[j] if j<len(subjects) else (subjects[0] if subjects else '?')
                    par_subj,par_teacher=self._find_parallel(cn,tsub)
                    all_rows.append({'group':group_no,'section':'A','class':cn,'subject':tsub,
                        'teacher':teacher,'par_subj':par_subj,'par_teacher':par_teacher})
                    covered.add((cn,tsub))
                    if par_subj not in ('—','?',''): covered.add((cn,par_subj))
        for cn in all_classes:
            for s in self.class_config_data.get(cn,{}).get('subjects',[]):
                sname=s.get('name','')
                if not s.get('parallel') or (cn,sname) in covered: continue
                ps=(s.get('parallel_subject') or '').strip(); pt=(s.get('parallel_teacher') or '').strip()
                if not ps or ps in ('—','?'): continue
                group_no+=1
                all_rows.append({'group':group_no,'section':'B','class':cn,'subject':sname,
                    'teacher':s.get('teacher',''),'par_subj':ps,'par_teacher':pt})
                covered.add((cn,sname)); covered.add((cn,ps))
        for cn in all_classes:
            for s in self.class_config_data.get(cn,{}).get('subjects',[]):
                sname=s.get('name','')
                if s.get('consecutive','No')!='Yes' or (cn,sname) in covered: continue
                group_no+=1
                all_rows.append({'group':group_no,'section':'C','class':cn,'subject':sname,
                    'teacher':s.get('teacher',''),'par_subj':'—','par_teacher':'—'})
                covered.add((cn,sname))
        group_slots=self._calc_slots(all_rows)
        allocation=self._alloc_slots(all_rows,group_slots)
        return group_slots,allocation,all_rows

    def _calc_slots(self, all_rows):
        cfg=self.configuration; wdays=cfg['working_days']; ppd=cfg['periods_per_day']
        result={}
        group_rows={}
        for row in all_rows: group_rows.setdefault(row['group'],[]).append(row)
        for gn,rows in group_rows.items():
            fr=rows[0]; cn=fr['class']; subj=fr['subject']; sec=fr['section']
            periods=0
            for s in self.class_config_data.get(cn,{}).get('subjects',[]):
                if s['name']==subj: periods=int(s.get('periods',0)); break
            if periods==0:
                result[gn]={'ok':False,'slots':0,'reason':"Subject '{}' not found in '{}'".format(subj,cn)}; continue
            if sec=='C' and periods%2!=0:
                result[gn]={'ok':False,'slots':periods,'reason':"Consecutive needs even count, got {}".format(periods)}; continue
            if periods>wdays*ppd:
                result[gn]={'ok':False,'slots':periods,'reason':"Periods > grid capacity"}; continue
            result[gn]={'ok':True,'slots':periods}
        return result

    def _alloc_slots(self, all_rows, group_slots):
        if not self._gen:
            return {r['group']:{'ok':False,'total':0,'s1_placed':0,'new_placed':0,'slots':[],'reason':'No gen state'} for r in all_rows}
        g=self._gen; grid=g['grid']; t_busy=g['t_busy']; ppd=g['ppd']; wdays=g['wdays']; DAYS=g['DAYS']
        tasks=g['tasks']

        def slot_free(cn_list,d,p): return all(grid.get(cn,[[]])[d][p] is None for cn in cn_list if cn in grid)
        def tfree(t,d,p): return not t or t in ('—','?','') or ((d,p) not in t_busy.get(t,set()) and not g['t_unavail'](t,d,p))
        def all_tfree(teachers,d,p): return all(tfree(t,d,p) for t in teachers)

        def place(task,extra,d,p,cim=None):
            self._place(task,d,p)
            if cim:
                for cn,info in cim.items():
                    if cn in grid and grid[cn][d][p] is not None:
                        patch={'type':info['type'],'par_subj':info['par_subj'],'par_teach':info['par_teach']}
                        pt=info.get('primary_teacher','').strip()
                        if pt and pt not in ('—','?'): patch['teacher']=pt
                        grid[cn][d][p]=dict(grid[cn][d][p],**patch)
            all_extra=set(extra)
            if cim:
                for info in cim.values():
                    pt=info.get('par_teach','') or ''
                    if pt and pt not in ('—','?',''): all_extra.add(pt)
            ep=(task.get('par_teach') or '').strip()
            for t in all_extra:
                if t and t not in ('—','?','') and t!=ep: t_busy.setdefault(t,set()).add((d,p))

        tbp={}; tbpar={}
        for _t in tasks:
            tbp[(frozenset(_t['cn_list']),_t['subject'],_t['teacher'])]=_t
            ps=(_t.get('par_subj') or '').strip(); pt=(_t.get('par_teach') or '').strip()
            if ps and pt and ps not in ('—','?'): tbpar[(frozenset(_t['cn_list']),ps,pt)]=_t

        group_rows={}; group_sec={}
        for row in all_rows:
            gn=row['group']
            if gn not in group_rows: group_rows[gn]=[]; group_sec[gn]=row['section']
            group_rows[gn].append(row)

        result={}
        for sec in ('C','B','A'):
            for gn,rows in sorted(group_rows.items()):
                if group_sec[gn]!=sec: continue
                gs=group_slots.get(gn)
                if gs is None or not gs['ok']:
                    result[gn]={'ok':False,'total':0,'s1_placed':0,'new_placed':0,'slots':[],
                        'reason':(gs['reason'] if gs else 'Unknown')}; continue
                total=gs['slots']; fr=rows[0]; psub=fr['subject']; pteach=fr['teacher']
                all_cn=list(dict.fromkeys(r['class'] for r in rows)); cn_fs=frozenset(all_cn)
                task=tbp.get((cn_fs,psub,pteach)) or tbpar.get((cn_fs,psub,pteach))
                if task is None:
                    for to in tasks:
                        if not (frozenset(to['cn_list'])&cn_fs): continue
                        if to['subject']==psub and to['teacher']==pteach: task=to; break
                        if to.get('par_subj','')==psub and to.get('par_teach','')==pteach: task=to; break
                if task is None:
                    result[gn]={'ok':False,'total':total,'s1_placed':0,'new_placed':0,'slots':[],
                        'reason':"Task '{}'/{} not found".format(psub,pteach)}; continue
                fvp=(task.get('par_subj','').strip()==psub and task.get('par_teach','').strip()==pteach)
                s1p=task['periods']-task['remaining']; rem=task['remaining']
                if rem<=0:
                    result[gn]={'ok':True,'total':total,'s1_placed':s1p,'new_placed':0,'slots':[]}; continue
                tt_teachers=[t for t in [task['teacher'],task.get('par_teach','')] if t and t not in ('—','?','')]
                row_teachers=[r.get(f,'') for r in rows for f in ('teacher','par_teacher') if r.get(f,'') and r.get(f,'') not in ('—','?','')]
                all_t=list(dict.fromkeys(tt_teachers+row_teachers))
                em=set(filter(None,[task['teacher'],task.get('par_teach','')]))
                extra=[t for t in all_t if t not in em]
                cim={}
                for row in rows:
                    ps=(row.get('par_subj') or '').strip(); pt=(row.get('par_teacher') or '').strip()
                    hp=bool(ps and pt and ps not in ('—','?') and pt not in ('—','?'))
                    if sec=='A': ct='combined_parallel' if hp else 'combined'
                    elif sec=='B': ct='parallel'
                    else: ct='normal'
                    prim=((row.get('par_teacher') or '') if fvp else (row.get('teacher') or '')).strip()
                    cim[row['class']]={'type':ct,'par_subj':ps if hp else '','par_teach':pt if hp else '','primary_teacher':prim}
                placed=[]; fail=''
                _rel=self._relaxed_consec_keys
                _grel=(sec=='C' and rows and (rows[0]['class'],rows[0]['subject']) in _rel)
                if sec=='C' and not _grel:
                    for ps in range(ppd-2,-1,-1):
                        if len(placed)>=rem: break
                        p1,p2=ps,ps+1
                        for d in range(wdays):
                            if len(placed)>=rem: break
                            if slot_free(all_cn,d,p1) and slot_free(all_cn,d,p2) and all_tfree(all_t,d,p1) and all_tfree(all_t,d,p2):
                                if rem-len(placed)>=2:
                                    place(task,extra,d,p1,cim); place(task,extra,d,p2,cim); placed+=[(d,p1),(d,p2)]
                                else:
                                    place(task,extra,d,p1,cim); placed.append((d,p1))
                else:
                    for p in range(ppd-1,-1,-1):
                        if len(placed)>=rem: break
                        for d in range(wdays):
                            if len(placed)>=rem: break
                            if slot_free(all_cn,d,p) and all_tfree(all_t,d,p):
                                place(task,extra,d,p,cim); placed.append((d,p))
                            else:
                                busy=[cn for cn in all_cn if cn in grid and grid[cn][d][p] is not None]
                                occ=grid[busy[0]][d][p] if busy else {}
                                fail='{} P{}: {} busy'.format(DAYS[d],p+1,occ.get('subject','?') if busy else ', '.join([t for t in all_t if not tfree(t,d,p)]))
                ok=(len(placed)+s1p>=total)
                result[gn]={'ok':ok,'total':total,'s1_placed':s1p,'new_placed':len(placed),'slots':placed,
                    'reason':'' if ok else 'Only {}/{} placed. Last: {}'.format(len(placed)+s1p,total,fail)}
        return result

    def get_combined_par_display(self, cn, e):
        cc=e.get('combined_classes',[])
        ct=''; cs=''
        for _t,s3d in self.step3_data.items():
            for cb in s3d.get('combines',[]):
                if set(cb.get('classes',[]))==set(cc):
                    ct=_t; cs=cb.get('subjects',[''])[0] if cb.get('subjects') else ''; break
            if ct: break
        cls=''; clt=''
        if cs and cn in self.class_config_data:
            for _s in self.class_config_data[cn].get('subjects',[]):
                sn=_s.get('name','').strip(); pn=(_s.get('parallel_subject') or '').strip()
                if sn==cs: cls=pn; clt=(_s.get('parallel_teacher') or '').strip(); break
                elif pn==cs: cls=sn; clt=_s.get('teacher','').strip(); break
        if not cs: cs=e.get('subject',''); ct=e.get('teacher',''); cls=e.get('par_subj',''); clt=e.get('par_teach','')
        l1="{} / {}".format(cs,cls) if cls else cs
        l2="{} / {}".format(ct,clt) if clt else ct
        return l1,l2

    # ── Excel Export ─────────────────────────────────────────────────────────
    def export_excel(self, mode, timetable=None):
        import openpyxl
        from openpyxl.styles import PatternFill,Font,Alignment,Border,Side
        from openpyxl.utils import get_column_letter
        tt=timetable or self.get_timetable()
        days=tt['days']; ppd=tt['ppd']; half1=tt['half1']; grid=tt['grid']; all_classes=tt['all_classes']
        sv=self._sv
        def _f(h): return PatternFill("solid",fgColor=h.lstrip("#"))
        def _fn(bold=False,sz=9,col="000000"): return Font(bold=bold,size=sz,color=col.lstrip("#"),name="Arial")
        def _b():
            s=Side(style="thin",color="AAAAAA"); return Border(left=s,right=s,top=s,bottom=s)
        def _a(h="center",w=True): return Alignment(horizontal=h,vertical="center",wrap_text=w)
        HF=_f("#2c3e50"); HN=_fn(True,10,"FFFFFF"); DF=_f("#34495e"); DN=_fn(True,9,"FFFFFF")
        SF=_f("#d5e8d4"); CF2=_f("#dae8fc"); PF=_f("#ffe6cc"); CPF=_f("#f8cecc")
        FF=_f("#f5f5f5"); WF=_f("#FFFFFF"); SMF=_f("#eaf2ff"); CTF=_f("#1a5276"); WRF=_f("#fdebd0")
        wb=openpyxl.Workbook(); wb.remove(wb.active)
        def _ct_map():
            ct={}
            for cn in all_classes:
                cfg=self.class_config_data.get(cn,{})
                t=sv(cfg.get('teacher','')).strip()
                if t: ct.setdefault(t,[]).append(cn)
            return ct
        def _tg():
            tg={}
            for cn in all_classes:
                for d in range(len(days)):
                    for p in range(ppd):
                        e=grid.get(cn,[[]])[d][p] if d<len(grid.get(cn,[])) else None
                        if not e: continue
                        etype=e.get('type','normal'); cc=e.get('combined_classes',[])
                        is_cp=bool(cc) and etype=='combined_parallel'; is_c=bool(cc) and etype=='combined'
                        def _add(tn,tc,ts,tct):
                            if not tn: return
                            tg.setdefault(tn,[[None]*ppd for _ in range(len(days))])
                            tg[tn][d][p]={'class':tc,'subject':ts,'is_ct':tct}
                        if is_cp:
                            if not cc or cn==cc[0]: _add(e.get('teacher'),'+'.join(cc),e.get('subject',''),False)
                            pt=e.get('par_teach','')
                            if pt and pt not in ('—','?',''): _add(pt,cn,e.get('par_subj',''),e.get('is_ct',False))
                        elif is_c:
                            if not cc or cn==cc[0]: _add(e.get('teacher'),'+'.join(cc),e.get('subject',''),e.get('is_ct',False))
                        else:
                            _add(e.get('teacher'),cn,e.get('subject',''),e.get('is_ct',False))
                            pt=e.get('par_teach','')
                            if pt and pt not in ('—','?',''): _add(pt,cn,e.get('par_subj',''),False)
            return tg
        if mode=="class":
            for cn in all_classes:
                ws=wb.create_sheet(cn); cfg=self.class_config_data.get(cn,{})
                ctn=sv(cfg.get('teacher','')).strip(); ctp=sv(cfg.get('teacher_period',''))
                hdr="Class: {}   |   CT: {}{}".format(cn,ctn or '—',"   |   P:{}".format(ctp) if ctp else '')
                ws.merge_cells(start_row=1,start_column=1,end_row=1,end_column=ppd+1)
                c=ws.cell(1,1,hdr); c.fill=CTF; c.font=_fn(True,11,"FFFFFF"); c.alignment=_a(); c.border=_b()
                ws.cell(2,1,"Day"); ws.cell(2,1).fill=HF; ws.cell(2,1).font=HN; ws.cell(2,1).alignment=_a(); ws.cell(2,1).border=_b()
                for p in range(ppd):
                    h=ws.cell(2,p+2,"P{} {}".format(p+1,"①" if p<half1 else "②"))
                    h.fill=HF; h.font=HN; h.alignment=_a(); h.border=_b()
                for d,dn in enumerate(days):
                    r=3+d; ws.row_dimensions[r].height=48
                    dc=ws.cell(r,1,dn); dc.fill=DF; dc.font=DN; dc.alignment=_a(); dc.border=_b()
                    for p in range(ppd):
                        e=grid.get(cn,[[]])[d][p] if d<len(grid.get(cn,[])) else None
                        if e is None: txt="FREE"; fill=FF
                        else:
                            et=e.get('type','normal')
                            if et=='combined_parallel': l1,l2=self.get_combined_par_display(cn,e); txt="{}\n{}".format(l1,l2); fill=CPF
                            elif et=='parallel': txt="{}/{}\n{}/{}".format(e['subject'],e.get('par_subj',''),e['teacher'],e.get('par_teach','')); fill=PF
                            elif et=='combined': txt="{}[{}]\n{}".format(e['subject'],'+'.join(e.get('combined_classes',[])),e['teacher']); fill=CF2
                            else: txt="{}{}\n{}".format(e['subject']," ★" if e.get('is_ct') else "",e['teacher']); fill=SF if e.get('is_ct') else WF
                        c=ws.cell(r,p+2,txt); c.fill=fill; c.alignment=_a(); c.border=_b(); c.font=_fn(sz=8)
                sr=3+len(days)+1
                ws.merge_cells(start_row=sr,start_column=1,end_row=sr,end_column=ppd+1)
                c=ws.cell(sr,1,"Summary — {}".format(cn)); c.fill=HF; c.font=HN; c.alignment=_a("left"); c.border=_b()
                smry=defaultdict(int)
                for d in range(len(days)):
                    for p in range(ppd):
                        e=grid.get(cn,[[]])[d][p] if d<len(grid.get(cn,[])) else None
                        if e: smry[(e['subject'],e['teacher'])]+=1
                for col,txt in enumerate(["Subject","Teacher","Periods/Week"],1):
                    c=ws.cell(sr+1,col,txt); c.fill=HF; c.font=HN; c.alignment=_a(); c.border=_b()
                for i,((s,t),cnt) in enumerate(sorted(smry.items())):
                    row=sr+2+i
                    for col,val in enumerate([s,t,cnt],1):
                        c=ws.cell(row,col,val); c.fill=SMF if i%2==0 else WF; c.alignment=_a(); c.border=_b(); c.font=_fn(sz=9)
                ws.column_dimensions["A"].width=12
                for p in range(ppd): ws.column_dimensions[get_column_letter(p+2)].width=20
        elif mode=="teacher":
            tg=_tg(); ctm=_ct_map()
            for teacher in sorted(tg.keys()):
                ws=wb.create_sheet(teacher[:31]); td=tg[teacher]; ctc=ctm.get(teacher,[])
                hdr="Teacher: {}   |   CT of: {}".format(teacher,', '.join(ctc) if ctc else '—')
                ws.merge_cells(start_row=1,start_column=1,end_row=1,end_column=ppd+1)
                c=ws.cell(1,1,hdr); c.fill=CTF; c.font=_fn(True,11,"FFFFFF"); c.alignment=_a(); c.border=_b()
                ws.cell(2,1,"Day"); ws.cell(2,1).fill=HF; ws.cell(2,1).font=HN; ws.cell(2,1).alignment=_a(); ws.cell(2,1).border=_b()
                for p in range(ppd):
                    h=ws.cell(2,p+2,"P{} {}".format(p+1,"①" if p<half1 else "②"))
                    h.fill=HF; h.font=HN; h.alignment=_a(); h.border=_b()
                for d,dn in enumerate(days):
                    r=3+d; ws.row_dimensions[r].height=48
                    dc=ws.cell(r,1,dn); dc.fill=DF; dc.font=DN; dc.alignment=_a(); dc.border=_b()
                    for p in range(ppd):
                        e=td[d][p] if d<len(td) else None
                        txt="{}\n{}".format(e['class'],e['subject']) if e else "FREE"
                        fill=SF if (e and e.get('is_ct')) else (FF if not e else WF)
                        c=ws.cell(r,p+2,txt); c.fill=fill; c.alignment=_a(); c.border=_b(); c.font=_fn(sz=8)
                sr=3+len(days)+1
                ws.merge_cells(start_row=sr,start_column=1,end_row=sr,end_column=ppd+1)
                c=ws.cell(sr,1,"Summary — {}".format(teacher)); c.fill=HF; c.font=HN; c.alignment=_a("left"); c.border=_b()
                smry=defaultdict(lambda: defaultdict(int)); total=0
                for d in range(len(days)):
                    for p in range(ppd):
                        e=td[d][p] if d<len(td) else None
                        if e: smry[e['class']][e['subject']]+=1; total+=1
                for col,txt in enumerate(["Class","Subject","Periods/Week"],1):
                    c=ws.cell(sr+1,col,txt); c.fill=HF; c.font=HN; c.alignment=_a(); c.border=_b()
                row=sr+2
                for cls in sorted(smry.keys()):
                    for subj,cnt in sorted(smry[cls].items()):
                        for col,val in enumerate([cls,subj,cnt],1):
                            c=ws.cell(row,col,val); c.fill=SMF if row%2==0 else WF; c.alignment=_a(); c.border=_b(); c.font=_fn(sz=9)
                        row+=1
                for col,val in enumerate(["","TOTAL",total],1):
                    c=ws.cell(row,col,val); c.fill=_f("#d4e6f1"); c.font=_fn(True,9); c.alignment=_a(); c.border=_b()
                ws.column_dimensions["A"].width=12
                for p in range(ppd): ws.column_dimensions[get_column_letter(p+2)].width=20
        elif mode=="ct_list":
            ws=wb.create_sheet("Class Teacher List")
            ws.merge_cells("A1:C1"); c=ws["A1"]; c.value="Class Teacher List"; c.fill=HF; c.font=_fn(True,13,"FFFFFF"); c.alignment=_a(); c.border=_b()
            for col,txt in enumerate(["Class","Class Teacher","CT Period"],1):
                c=ws.cell(2,col,txt); c.fill=DF; c.font=DN; c.alignment=_a(); c.border=_b()
            for i,cn in enumerate(all_classes):
                cfg=self.class_config_data.get(cn,{}); ctn=sv(cfg.get('teacher','')).strip() or '—'; ctp=sv(cfg.get('teacher_period','')) or '—'
                for col,val in enumerate([cn,ctn,ctp],1):
                    c=ws.cell(3+i,col,val); c.fill=SMF if i%2==0 else WF; c.alignment=_a(); c.border=_b(); c.font=_fn(sz=10)
            ws.column_dimensions["A"].width=14; ws.column_dimensions["B"].width=28; ws.column_dimensions["C"].width=12
        elif mode=="workload":
            tg=_tg(); ctm=_ct_map(); ws=wb.create_sheet("Teacher Workload")
            ws.merge_cells("A1:E1"); c=ws["A1"]; c.value="Teacher Workload List"; c.fill=HF; c.font=_fn(True,13,"FFFFFF"); c.alignment=_a(); c.border=_b()
            for col,txt in enumerate(["Teacher","Subject","Class","Periods/Week","Total"],1):
                c=ws.cell(2,col,txt); c.fill=DF; c.font=DN; c.alignment=_a(); c.border=_b()
            row=3; grand=0
            for teacher in sorted(tg.keys()):
                td=tg[teacher]; smry=defaultdict(lambda: defaultdict(int))
                for d in range(len(days)):
                    for p in range(ppd):
                        e=td[d][p] if d<len(td) else None
                        if e: smry[e['subject']][e['class']]+=1
                total=sum(c2 for cd in smry.values() for c2 in cd.values()); grand+=total; ctc=ctm.get(teacher,[]); sr2=row
                for subj in sorted(smry.keys()):
                    for cls,cnt in sorted(smry[subj].items()):
                        fill=SMF if row%2==0 else WF
                        c=ws.cell(row,1,teacher if row==sr2 else ""); c.fill=WRF if ctc else fill; c.font=_fn(True if row==sr2 else False,9); c.alignment=_a(); c.border=_b()
                        for col,val in enumerate([subj,cls,cnt],2):
                            c2=ws.cell(row,col,val); c2.fill=fill; c2.alignment=_a(); c2.border=_b(); c2.font=_fn(sz=9)
                        c5=ws.cell(row,5,total if row==sr2 else ""); c5.fill=_f("#d4e6f1") if row==sr2 else fill; c5.font=_fn(True if row==sr2 else False,9); c5.alignment=_a(); c5.border=_b()
                        row+=1
                if row-sr2>1: ws.merge_cells(start_row=sr2,start_column=1,end_row=row-1,end_column=1)
            for col,val in enumerate(["","","","GRAND TOTAL",grand],1):
                c=ws.cell(row,col,val); c.fill=HF; c.font=_fn(True,10,"FFFFFF"); c.alignment=_a(); c.border=_b()
            for col,w in zip("ABCDE",[22,22,16,16,16]): ws.column_dimensions[col].width=w
        elif mode=="one_sheet":
            tg=_tg(); ws=wb.create_sheet("Teacherwise Timetable")
            for col,txt in enumerate(["Teacher","Day"],1):
                c=ws.cell(1,col,txt); c.fill=HF; c.font=HN; c.alignment=_a(); c.border=_b()
            for p in range(ppd):
                c=ws.cell(1,p+3,str(p+1)); c.fill=HF; c.font=HN; c.alignment=_a(); c.border=_b()
            row=2
            for teacher in sorted(tg.keys()):
                td=tg[teacher]; ts=row
                for d,dn in enumerate(days):
                    c=ws.cell(row,1,teacher if d==0 else ""); c.fill=WRF; c.alignment=_a(); c.font=_fn(True if d==0 else False,9); c.border=_b()
                    c2=ws.cell(row,2,dn); c2.fill=DF; c2.font=DN; c2.alignment=_a(); c2.border=_b()
                    for p in range(ppd):
                        e=td[d][p] if d<len(td) else None
                        txt="{}/{}".format(e['class'],e['subject']) if e else ""; fill=SF if (e and e.get('is_ct')) else (FF if not e else WF)
                        c3=ws.cell(row,p+3,txt); c3.fill=fill; c3.alignment=_a(); c3.border=_b(); c3.font=_fn(sz=8)
                    row+=1
                if len(days)>1: ws.merge_cells(start_row=ts,start_column=1,end_row=row-1,end_column=1)
            ws.column_dimensions["A"].width=22; ws.column_dimensions["B"].width=10
            for p in range(ppd): ws.column_dimensions[get_column_letter(p+3)].width=18
        buf=io.BytesIO(); wb.save(buf); buf.seek(0)
        return buf
