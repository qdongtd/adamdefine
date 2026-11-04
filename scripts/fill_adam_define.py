# -*- coding: utf-8 -*-
"""Fill a P21-generated ADaM Define-XML 2.1 spec workbook from the compound ADaM
spec (+ optional ADaM VLM), following SKILL.md.

Study-agnostic: the study ID, versions, datasets, origins and text all come from the
inputs. Nothing is invented - an unsourced cell is left blank, highlighted orange and
explained in a `Programmer Review` column. Red font = a value that overrides the
P21 skeleton (spec wins); orange fill = needs human review.

    python fill_adam_define.py --skeleton p21_adam_define.xlsx --spec adam_spec.xlsx \
        --out define_filled.xlsx [--vlm adam_vlm.xlsx] [--study STUDYID] \
        [--issues issues.xlsx] [--ct YYYY-MM-DD] [--keep-p21-dataset-meta]

--skeleton is never modified; --out is rebuilt each run. Then run self_qc_adam.py.
"""
import re, sys, copy, shutil, argparse, warnings
from collections import defaultdict, Counter
import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
warnings.simplefilter('ignore')

ap = argparse.ArgumentParser()
ap.add_argument('--skeleton', required=True); ap.add_argument('--spec', required=True)
ap.add_argument('--out', required=True); ap.add_argument('--vlm')
ap.add_argument('--study', help='default: Define tab StudyName')
ap.add_argument('--issues', help='write an issue workbook (Summary + Issue Details) here')
ap.add_argument('--ct', help='CT version when the user overrides the spec Study tab CTVersion')
ap.add_argument('--keep-p21-dataset-meta', action='store_true',
                help='keep P21 Class/Structure/Key Variables and only report Content-tab differences')
ap.add_argument('--nodata-text', default='Planned for collection but no data were collected for this study.')
args = ap.parse_args()

ORG = PatternFill(start_color='FFC000', end_color='FFC000', fill_type='solid'); NOFILL = PatternFill(fill_type=None)
# CDISC ADaM Controlled Terminology codelist C-codes (NCI EVS ADaM Terminology, 2026-09-25 release).
# A codelist with one of these codes belongs to ADaM CT; any other NCI-coded codelist to SDTM CT.
# Using the wrong standard makes P21 report DD0129 (it cannot find the code). Update with new CT releases.
ADAM_CT_CODES = {'C208382', 'C208381', 'C204411', 'C187458', 'C187457', 'C81223', 'C81224', 'C172334', 'C172335',
                 'C158114', 'C158115', 'C187456', 'C187455', 'C204410', 'C193278', 'C193277', 'C81225', 'C165644',
                 'C124296', 'C225073', 'C225072', 'C81226', 'C204414', 'C204413', 'C204412', 'C228401', 'C228402'}
ADAM_ORIGINS = {'Derived', 'Assigned', 'Predecessor'}
ADAM_CLASSES = {'SUBJECT LEVEL ANALYSIS DATASET', 'BASIC DATA STRUCTURE', 'OCCURRENCE DATA STRUCTURE', 'ADAM OTHER'}
DATE = {'date', 'datetime', 'time', 'partialDate', 'partialTime', 'partialDatetime', 'incompleteDatetime',
        'durationDatetime', 'intervalDatetime'}

def norm(s): return re.sub(r'[^0-9a-z]', '', str(s or '').lower())
def txt(x):
    s = '' if x is None else str(x).strip(); return '' if s == 'None' else s
def hdr(row): return {norm(x): i for i, x in enumerate(row) if x is not None}
def red(c): f = copy.copy(c.font); f.color = 'FFFF0000'; c.font = f
def wkey(w):   # where-clause key for matching skeleton vs VLM rows
    w = re.sub(r'["\'()]', '', txt(w)).upper()
    return re.sub(r'\s+', ' ', re.sub(r'\s*,\s*', ',', w)).strip()

shutil.copy(args.skeleton, args.out)
wb = openpyxl.load_workbook(args.out)
spec = openpyxl.load_workbook(args.spec, read_only=True, data_only=True)
STUDY = args.study or next(txt(r[1]) for r in wb['Define'].iter_rows(values_only=True) if txt(r[0]) == 'StudyName')
SK = norm(STUDY)
log, issues = Counter(), []
def issue(*a): issues.append(a)

# ---------------- spec: Study / Content / dataset tabs ----------------
def study_col(h): return next((i for i, x in enumerate(h) if norm(x) == SK), None)
def sheet(name): return next((spec[s] for s in spec.sheetnames if norm(s) == norm(name)), None)
study_attrs = {}
st = sheet('Study')
if st is not None:
    rows = list(st.iter_rows(values_only=True))
    k = study_col(rows[0])
    if k is None:
        sn = next((r for r in rows if txt(r[0]) == 'StudyName'), None)
        k = next((i for i, x in enumerate(sn or []) if norm(x) == SK), None)
    if k is not None: study_attrs = {txt(r[0]): txt(r[k]) for r in rows[1:] if r[0]}
IGVER = txt(study_attrs.get('StandardVersion'))
CTVER = args.ct or txt(study_attrs.get('CTVersion'))

cont = list(sheet('Content').iter_rows(values_only=True)); ch = hdr(cont[0]); cst = study_col(cont[0])
if cst is None: sys.exit('Content tab has no column for study %s' % STUDY)
content = {txt(r[0]): r for r in cont[1:] if r[0]}
in_scope = {d for d, r in content.items() if txt(r[cst])}

R_ORIG = re.compile(r'\[r_orig\]\s*([A-Za-z]+)')
def tag(s, t):   # text of one [r_t] tag, up to the next [r_...] tag
    m = re.search(r'\[r_%s\]\s*(.*?)(?=\[r_|$)' % t, s, re.S); return m.group(1).strip() if m else ''
specvars, core_adsl = {}, []
for tab in spec.sheetnames:
    if tab not in in_scope: continue
    rows = list(spec[tab].iter_rows(values_only=True))
    if not rows or norm(rows[0][0]) != 'variablename': continue
    h = hdr(rows[0]); sc = study_col(rows[0])
    if sc is None: issue('Spec', tab, '', '', 'Spec tab has no column for %s' % STUDY, '', '', ''); continue
    for r in rows[1:]:
        v = txt(r[0])
        if not v: continue
        s = txt(r[sc])
        d = dict(var=v, study=s, origin=txt(r[h['origin']]), conv=txt(r[h.get('conversiondefinition')]) if 'conversiondefinition' in h else '',
                 notes=txt(r[h['programmingnotes']]) if 'programmingnotes' in h else '',
                 core=txt(r[h['corevariable']]) if 'corevariable' in h else '',
                 type=txt(r[h['type']]) if 'type' in h else '', length=r[h['length']] if 'length' in h else None)
        m = R_ORIG.search(s); d['r_orig'] = m.group(1) if m else ''
        d['r_def'], d['r_prog'] = tag(s, 'def'), tag(s, 'prog')
        if tab == 'ADSL' and d['core'].upper() in ('Y', 'YES') and s: core_adsl.append(v)
        key = (tab, v)
        if key in specvars and specvars[key]['study'] and not s: continue
        specvars[key] = d
corevar_in = {ds for (ds, v), d in specvars.items() if v == '<COREVAR>' and d['study']}

def text_of(d):
    for lab in ('r_def', 'r_prog', 'conv', 'notes'):
        if d.get(lab): return d[lab], lab
    return '', ''
PRED_RE = re.compile(r'^\s*(?:set\s+(?:to|equal\s+to)\s+)?([A-Z][A-Z0-9]{1,7})\.([A-Z][A-Z0-9]{1,7})\s*\.?\s*$', re.I)
PRED_LEAD = re.compile(r'^\s*(?:set\s+(?:to|equal\s+to)\s+)?([A-Z][A-Z0-9]{1,7})\.([A-Z][A-Z0-9]{1,7})(.*)$', re.I | re.S)
DSVAR = re.compile(r'(?<![A-Za-z0-9_])([A-Z][A-Z0-9]{1,7})\.([A-Z][A-Z0-9]{1,7})(?![A-Za-z0-9_])')
LIBREFS = {'SDTM', 'ADAM', 'RAW', 'LOOKUP', 'WORK', 'LIBRARY'}   # SDTM.CO is a dataset reference, not DATASET.VARIABLE
def pred_of(t):
    m = PRED_RE.match(t or ''); return ('%s.%s' % (m.group(1).upper(), m.group(2).upper())) if m else ''
def pred_parse(t):
    """(predecessor, clean) from spec text. clean=False -> text carries more than a copy."""
    t = t or ''
    if pred_of(t): return pred_of(t), True
    m = PRED_LEAD.match(t)
    if m:
        dv, rest = '%s.%s' % (m.group(1).upper(), m.group(2).upper()), m.group(3).strip()
        if re.match(r'^where\s+\S', rest, re.I) and chr(10) not in rest:   # SUPPxx.QVAL where QNAM="X"
            return '%s %s' % (dv, rest), True
        return dv, False
    found = sorted({'%s.%s' % (a.upper(), b.upper()) for a, b in DSVAR.findall(t) if a.upper() not in LIBREFS})
    return (found[0], False) if len(found) == 1 else ('', False)
def map_origin(o):
    return {'derived': 'Derived', 'assigned': 'Assigned', 'predecessor': 'Predecessor'}.get(norm(o), txt(o))

# ---------------- helpers ----------------
methods, comments = {}, {}
COMMON = {'COM.NODATA': args.nodata_text}
def add_pr_col(ws):
    h = [c.value for c in ws[1]]
    if 'Programmer Review' not in h: ws.cell(1, len(h) + 1, 'Programmer Review'); h.append('Programmer Review')
    return {n: i for i, n in enumerate(h)}
def put_pr(ws, ci, row, msgs):
    if not msgs: return
    c = ws.cell(row, ci['Programmer Review'] + 1); c.value = '; '.join(([txt(c.value)] if txt(c.value) else []) + msgs)
def setv(ws, ci, row, col, val, p21=None):
    c = ws.cell(row, ci[col] + 1); c.value = val if val else None
    if p21 is not None and txt(p21) != txt(val): red(c)

def apply(ws, ci, rr, ds, var, origin, text, src, idb, p21o, p21pred, label, alts=()):
    """Write Origin/Source/Predecessor/Method/Comment for one row from (origin, text)."""
    pr = []
    if origin and origin not in ADAM_ORIGINS:
        pr.append('Origin: spec says %s - not a valid ADaM origin (Derived/Assigned/Predecessor)' % origin)
        ws.cell(rr, ci['Origin'] + 1).fill = ORG; issue(ws.title, ds, var, label, 'Spec origin not valid for ADaM', origin, '', text[:200])
        origin = ''
    if not origin:
        if p21o:
            origin = p21o; ws.cell(rr, ci['Origin'] + 1).fill = ORG
            pr.append('Origin: no spec origin - kept P21 skeleton %s' % p21o); log['origin kept from P21'] += 1
            issue(ws.title, ds, var, label, 'No spec origin - P21 skeleton origin kept', p21o, p21o, '')
        else:
            ws.cell(rr, ci['Origin'] + 1).fill = ORG; pr.append('Origin: no source in spec or skeleton'); log['origin unsourced'] += 1
            issue(ws.title, ds, var, label, 'No origin in spec or P21 skeleton', '', '', ''); return pr
    elif p21o and p21o != origin:
        pr.append('Origin: P21 skeleton=%s -> spec %s' % (p21o, origin)); log['origin P21->spec'] += 1
        issue(ws.title, ds, var, label, 'Origin: P21 skeleton differs from spec - spec used', p21o, origin, '')
    setv(ws, ci, rr, 'Origin', origin, p21o)
    for col in ('Method', 'Comment', 'Predecessor', 'Pages'):
        if col in ci: ws.cell(rr, ci[col] + 1).value = None
    if origin == 'Predecessor':
        if 'Source' in ci: ws.cell(rr, ci['Source'] + 1).value = None
        pv, clean = pred_parse(text)
        if not clean:   # another spec text field may hold the plain copy (e.g. Programming Notes 'EX.EXDOSE')
            for alt in alts:
                apv, aclean = pred_parse(alt)
                if aclean:
                    pr.append('Predecessor %s taken from another spec column (main text: "%s")' % (apv, text[:100])); pv, clean = apv, True; break
                if apv and not pv: pv = apv
        if text.strip().upper() in ('ADSL', 'ADSL.<CORE VARIABLE>'): pv, clean = 'ADSL.%s' % var, True
        if pv and not clean:
            ws.cell(rr, ci['Predecessor'] + 1).fill = ORG
            pr.append('Predecessor: spec text has more than a copy ("%s") - %s used, consider Derived' % (text[:120], pv))
            issue(ws.title, ds, var, label, 'Predecessor origin but spec text is a derivation, not a plain copy - consider Derived', origin, pv, text[:200])
        if not pv and p21pred:
            pv = p21pred; pr.append('Predecessor: spec text "%s" is not a DATASET.VARIABLE - kept P21 %s, confirm' % (text[:80], p21pred))
            ws.cell(rr, ci['Predecessor'] + 1).fill = ORG
        elif not pv:
            ws.cell(rr, ci['Predecessor'] + 1).fill = ORG; pr.append('Predecessor: no DATASET.VARIABLE in spec text "%s"' % text[:120])
            issue(ws.title, ds, var, label, 'Predecessor origin but spec text is not a DATASET.VARIABLE copy', origin, '', text[:200])
        elif p21pred and p21pred != pv:
            pr.append('Predecessor: P21=%s -> spec %s' % (p21pred, pv))
        setv(ws, ci, rr, 'Predecessor', pv, p21pred); log['predecessor'] += 1
    else:
        if 'Source' in ci: ws.cell(rr, ci['Source'] + 1).value = 'Sponsor'
        col, store, pre = ('Method', methods, 'MT.') if origin == 'Derived' else ('Comment', comments, 'COM.')
        if text:
            store[pre + idb] = text; ws.cell(rr, ci[col] + 1).value = pre + idb; log[col.lower()] += 1
            if pred_of(text) and origin != 'Predecessor':
                issue(ws.title, ds, var, label, '%s text is a pure copy (%s) - Predecessor candidate (CT C170550)' % (origin, pred_of(text)), origin, 'Predecessor', text[:120])
        else:
            ws.cell(rr, ci[col] + 1).fill = ORG; pr.append('%s: %s but no text in spec' % (col, origin))
            log['%s missing' % col.lower()] += 1
            issue(ws.title, ds, var, label, '%s origin but no derivation/comment text in spec' % origin, origin, '', '')
    if src in ('r_def', 'r_prog'): pr.append('text from [%s]' % src)
    return pr

# ---------------- Variables ----------------
V = wb['Variables']; vci = add_pr_col(V); vrow = {}
for rr in range(2, V.max_row + 1):
    ds, var = txt(V.cell(rr, vci['Dataset'] + 1).value), txt(V.cell(rr, vci['Variable'] + 1).value)
    if not ds or not var: continue
    vrow[(ds, var)] = rr
    p21o, p21pred = txt(V.cell(rr, vci['Origin'] + 1).value), txt(V.cell(rr, vci['Predecessor'] + 1).value)
    d = specvars.get((ds, var)); pr = []
    if ds not in in_scope: pr.append('Dataset %s not marked for %s on Content tab' % (ds, STUDY))
    if d is None and ds != 'ADSL' and ds in corevar_in and var in core_adsl:
        a_ = specvars.get(('ADSL', var), {})
        d = dict(origin='Predecessor', r_orig='', r_def='', r_prog='', conv='ADSL.%s' % var, notes='', study='x',
                 type=a_.get('type', ''), length=a_.get('length'))
        pr.append('ADSL core variable (<COREVAR> row)')
    if d is None:
        pr.append('Variable not found in spec %s tab' % ds)
        issue('Variables', ds, var, '', 'Variable in data but not in spec tab', p21o, '', '')
        origin, text, src = '', '', ''
    else:
        if not d['study']: pr.append('Variable not marked for %s in spec (present in data)' % STUDY)
        origin = map_origin(d['r_orig'] or d['origin']); text, src = text_of(d)
    alts = [d[k] for k in ('r_def', 'r_prog', 'conv', 'notes') if d and d.get(k)] if d else []
    pr += apply(V, vci, rr, ds, var, origin, text, src, '%s.%s' % (ds, var), p21o, p21pred, '', alts)
    if txt(V.cell(rr, vci['Has No Data'] + 1).value) == 'Yes' and 'Comment' in vci and not txt(V.cell(rr, vci['Comment'] + 1).value):
        comments['COM.NODATA'] = COMMON['COM.NODATA']; V.cell(rr, vci['Comment'] + 1).value = 'COM.NODATA'
    put_pr(V, vci, rr, pr)
    dt = V.cell(rr, vci['Data Type'] + 1)
    if not txt(dt.value) and d is not None and d.get('type'):   # MD0034: skeleton blank -> spec Type/Length
        T = {'text': 'text', 'char': 'text', 'integer': 'integer', 'float': 'float', 'date': 'date',
             'datetime': 'datetime', 'time': 'time'}.get(d['type'].strip().lower())
        if T:
            dt.value = T
            L = d.get('length'); L = str(int(L)) if (T not in DATE and str(L).strip().replace('.0', '').isdigit()) else None
            V.cell(rr, vci['Length'] + 1).value = L
            put_pr(V, vci, rr, ['DataType/Length blank in skeleton: %s/%s from spec Type/Length' % (T, L)]); log['dtype from spec'] += 1
        else:
            dt.fill = ORG; put_pr(V, vci, rr, ['DataType blank; spec Type "%s" not mappable' % d['type']])
    elif not txt(dt.value):
        dt.fill = ORG; put_pr(V, vci, rr, ['DataType blank in skeleton and spec'])
        issue('Variables', ds, var, '', 'DataType blank in skeleton and spec', '', '', '')
    if dt.value in DATE and txt(V.cell(rr, vci['Length'] + 1).value):
        V.cell(rr, vci['Length'] + 1).value = None; put_pr(V, vci, rr, ['Length cleared for date type (DD0068)'])

# ---------------- ValueLevel ----------------
VL = wb['ValueLevel']; lci = add_pr_col(VL)
vlm_rows = {}
if args.vlm:
    vb = openpyxl.load_workbook(args.vlm, read_only=True, data_only=True)
    for tab in vb.sheetnames:
        rows = list(vb[tab].iter_rows(values_only=True))
        if not rows or norm(rows[0][0]) != 'variablename' or tab not in in_scope: continue
        h = hdr(rows[0]); sc = study_col(rows[0])
        if sc is None: issue('VLM', tab, '', '', 'VLM tab has no column for %s - parent spec rows used' % STUDY, '', '', ''); continue
        for r in rows[1:]:
            if not r[0] or not txt(r[sc]): continue
            wv, cm, wval = (txt(r[h[k]]) for k in ('wherevarlist', 'comparator', 'wherevalue'))
            parts = []
            for v_, c_, x_ in zip(*[[p.strip() for p in s.split(';')] for s in (wv, cm, wval)]):
                parts.append('%s %s %s' % (v_, c_, x_))
            vlm_rows[(tab, txt(r[0]), wkey(' AND '.join(parts)))] = dict(
                origin=map_origin(r[h['origin']]), conv=txt(r[h.get('conversiondefinition')]), notes=txt(r[h.get('programmingnotes')]),
                r_orig='', r_def=tag(txt(r[sc]), 'def'), r_prog=tag(txt(r[sc]), 'prog'), study=txt(r[sc]))
used_vlm = set(); inherited = defaultdict(list)
for rr in range(2, VL.max_row + 1):
    ds, var, wc = (txt(VL.cell(rr, lci[n] + 1).value) for n in ('Dataset', 'Variable', 'Where Clause'))
    if not ds: continue
    p21o, p21pred = txt(VL.cell(rr, lci['Origin'] + 1).value), txt(VL.cell(rr, lci['Predecessor'] + 1).value)
    k = (ds, var, wkey(wc)); m = vlm_rows.get(k); pr = []
    parent = specvars.get((ds, var))
    pmo = map_origin((parent['r_orig'] or parent['origin'])) if parent else ''
    wid = re.sub(r'[^A-Za-z0-9]+', '_', re.sub(r'^\S+\s+EQ\s+', '', wc))[:40].strip('_')
    if m:
        used_vlm.add(k); origin = m['origin'] or pmo; text, src = text_of(m)
        if not m['origin']: pr.append('VLM origin blank - parent spec origin %s used' % pmo)
        log['VL from VLM'] += 1
    else:
        origin = pmo; text, src = text_of(parent) if parent else ('', '')
        log['VL from parent (no VLM row)'] += 1
        pr.append('No VLM row for this value - parent %s.%s spec row used' % (ds, var))
        if not text: VL.cell(rr, lci['Origin'] + 1).fill = ORG
        inherited[(ds, var)].append(re.sub(r'^\S+\s+EQ\s+', '', wc))
    srcd = m or parent or {}
    alts = [srcd[k] for k in ('r_def', 'r_prog', 'conv', 'notes') if srcd.get(k)]
    pr += apply(VL, lci, rr, ds, var, origin, text, src, '%s.%s.%s' % (ds, var, wid), p21o, p21pred, wc, alts)
    put_pr(VL, lci, rr, pr)
for (ds, var), vals in sorted(inherited.items()):
    issue('ValueLevel', ds, var, '%d values' % len(vals), 'No VLM row for these values - parent variable spec row used (review that the parent text fits each value)',
          '', '', ', '.join(vals)[:300])
for k in sorted(set(vlm_rows) - used_vlm):
    issue('VLM', k[0], k[1], k[2], 'VLM row marked for study but no matching value-level row in the data/skeleton', '', '', '')

# HT convention: variable with complete value level -> origin info only at value level
vln, vlf = Counter(), Counter()
for r in VL.iter_rows(min_row=2, values_only=True):
    if r[lci['Dataset']]: vln[(r[lci['Dataset']], r[lci['Variable']])] += 1; vlf[(r[lci['Dataset']], r[lci['Variable']])] += bool(txt(r[lci['Origin']]))
for k, n in vln.items():
    if k in vrow and vlf[k] == n:
        r = vrow[k]; cl = []
        for f in ('Origin', 'Source', 'Pages', 'Method', 'Predecessor', 'Comment'):
            c = V.cell(r, vci[f] + 1)
            if txt(c.value): cl.append('%s=%s' % (f, c.value)); c.value = None
            c.fill = NOFILL
        put_pr(V, vci, r, ['Defined at value level only (%d rows)%s' % (n, '; cleared ' + ', '.join(cl) if cl else '')])
        log['var-level cleared (VLM only)'] += 1

# ---------------- Datasets ----------------
D = wb['Datasets']; dci = add_pr_col(D); present = set()
for rr in range(2, D.max_row + 1):
    ds = txt(D.cell(rr, dci['Dataset'] + 1).value)
    if not ds: continue
    present.add(ds); pr = []
    if IGVER and 'Standard' in dci and not txt(D.cell(rr, dci['Standard'] + 1).value):
        D.cell(rr, dci['Standard'] + 1).value = 'ADaMIG %s' % IGVER
    if ds not in in_scope: pr.append('not marked for %s on Content tab' % STUDY); issue('Datasets', ds, '', '', 'Dataset in data but not marked for study on Content tab', '', '', '')
    cr = content.get(ds)
    if cr:
        m = re.search(r'\[r_sort\]\s*(.+)$', txt(cr[cst]), re.S)
        want = {'Class': txt(cr[ch['class']]).upper(), 'Structure': txt(cr[ch['structure']]),
                'Key Variables': (m.group(1) if m else txt(cr[ch['sortorder']])).strip()}
        for col, v in want.items():
            c = D.cell(rr, dci[col] + 1); cur = txt(c.value)
            if not v or re.sub(r'\s+', ' ', v).upper() == re.sub(r'\s+', ' ', cur).upper(): continue
            if col == 'Class' and v not in ADAM_CLASSES:
                c.fill = ORG; pr.append('Class: Content tab %s is not a valid ADaM class - kept %s' % (v, cur))
                issue('Datasets', ds, '', '', 'Content-tab Class is not a valid ADaM class', cur, v, ''); continue
            if args.keep_p21_dataset_meta and cur and not (col == 'Class' and cur not in ADAM_CLASSES):
                pr.append('%s: Content tab says %s - P21 kept' % (col, v)); issue('Datasets', ds, '', '', '%s: Content tab differs from P21 (P21 kept)' % col, cur, v, ''); continue
            c.value = v; red(c); pr.append('%s: P21 %s -> Content tab %s' % (col, cur or 'blank', v)); log['Datasets %s from Content' % col] += 1
            issue('Datasets', ds, '', '', '%s: P21 differs from Content tab - spec used' % col, cur, v, '')
    put_pr(D, dci, rr, pr)
for ds in sorted(in_scope - present): issue('Datasets', ds, '', '', 'Content tab marks dataset for study but it is not in the skeleton (no data?)', '', '', '')

# ---------------- Standards / Codelists ----------------
if 'Standards' not in wb.sheetnames and IGVER:
    S = wb.create_sheet('Standards'); S.append(['Name', 'Type', 'Publishing Set', 'Version', 'Status', 'Comment'])
    S.append(['ADaMIG', 'IG', None, IGVER, 'Final', None])
    if CTVER:
        S.append(['CDISC/NCI', 'CT', 'ADaM', CTVER, 'Final', 'Version from spec Study tab CTVersion - confirm ADaM CT release'])
        S.append(['CDISC/NCI', 'CT', 'SDTM', CTVER, 'Final', 'Version from spec Study tab CTVersion - confirm SDTM CT release'])
CL = wb['Codelists']; cli = {c.value: i for i, c in enumerate(CL[1])}
names, users = {}, defaultdict(set)
for ws, ci in ((V, vci), (VL, lci)):
    for rr in range(2, ws.max_row + 1):
        c = txt(ws.cell(rr, ci['Codelist'] + 1).value)
        if c: users[c].add(txt(ws.cell(rr, ci['Dataset'] + 1).value))
for rr in range(2, CL.max_row + 1):
    i = txt(CL.cell(rr, 1).value)
    if i: names.setdefault(i, txt(CL.cell(rr, cli['Name'] + 1).value))
dup = Counter(names.values())
for rr in range(2, CL.max_row + 1):
    i = txt(CL.cell(rr, 1).value)
    if not i: continue
    if CTVER and txt(CL.cell(rr, cli['NCI Codelist Code'] + 1).value) and not txt(CL.cell(rr, cli['Terminology'] + 1).value):
        code = txt(CL.cell(rr, cli['NCI Codelist Code'] + 1).value)
        CL.cell(rr, cli['Terminology'] + 1).value = ('ADaM ' if code in ADAM_CT_CODES else 'SDTM ') + CTVER
    if dup[names[i]] > 1:
        m2 = re.match(r'([A-Z][A-Z0-9]{1,7})\.([A-Z][A-Z0-9]{1,7})\.', i)
        tagd = ('%s.%s' % (m2.group(2), m2.group(1))) if m2 else (next(iter(users[i])) if len(users[i]) == 1 else i)
        CL.cell(rr, cli['Name'] + 1).value = '%s (%s)' % (names[i], tagd)

# ---------------- spec-content checks reported to the issue file ----------------
BAD = re.compile(r'\b(sdtm|raw|lookup)\.[a-z]|\bcatx\(|\bput\(|\bbest\.|\bstrip\(|\bupcase\(|tab "|"XX"|\[if ', re.I)
for store, kind in ((methods, 'Method'), (comments, 'Comment')):
    for i, t in store.items():
        if BAD.search(t) or t.count('"') % 2:
            issue(kind + 's', i.split('.')[1] if i.count('.') >= 2 else '', i.split('.')[2] if i.count('.') >= 2 else '', i,
                  '%s text from spec not reviewer-readable (libref/SAS code/placeholder/quotes) - copied verbatim; spec owner to reword' % kind, '', '', t[:250])
bds_vars = defaultdict(set)
for (ds, var) in vrow: bds_vars[ds].add(var)
for rr in range(2, D.max_row + 1):
    ds = txt(D.cell(rr, dci['Dataset'] + 1).value)
    if ds and txt(D.cell(rr, dci['Class'] + 1).value).upper() == 'BASIC DATA STRUCTURE' and not {'PARAMCD', 'PARAM'} <= bds_vars[ds]:
        D.cell(rr, dci['Class'] + 1).fill = ORG
        issue('Datasets', ds, '', '', 'Class BASIC DATA STRUCTURE but dataset has no PARAMCD/PARAM - check Content-tab class', txt(D.cell(rr, dci['Class'] + 1).value), '', '')

# ---------------- Methods / Comments (only referenced IDs: DD0079 / DD0080) ----------------
_ref = set()
for ws_, ci_ in ((V, vci), (VL, lci)):
    for r_ in ws_.iter_rows(min_row=2, values_only=True):
        _ref |= {txt(r_[ci_[c]]) for c in ('Method', 'Comment') if c in ci_ and txt(r_[ci_[c]])}
_ref |= {txt(r_[dci['Comment']]) for r_ in D.iter_rows(min_row=2, values_only=True) if 'Comment' in dci and txt(r_[dci['Comment']])}
for store in (methods, comments):
    for i in [i for i in store if i not in _ref]: del store[i]; log['unreferenced method/comment dropped'] += 1
M = wb['Methods']; mci = {c.value: i for i, c in enumerate(M[1])}
for mid in sorted(methods):
    r = [None] * len(mci); r[mci['ID']] = mid; r[mci['Name']] = mid; r[mci['Type']] = 'Computation'; r[mci['Description']] = methods[mid]; M.append(r)
C = wb['Comments']; cci = {c.value: i for i, c in enumerate(C[1])}
for cid in sorted(comments):
    r = [None] * len(cci); r[cci['ID']] = cid; r[cci['Description']] = comments[cid]; C.append(r)

out = args.out
try: wb.save(out)
except PermissionError:
    out = out.replace('.xlsx', '_updated.xlsx'); wb.save(out)
openpyxl.load_workbook(out)
print('saved', out, '| study', STUDY, '| ADaMIG', IGVER, '| CT', CTVER)
for k, v in sorted(log.items()): print('  %-34s %d' % (k, v))
print('  methods %d, comments %d, issues %d' % (len(methods), len(comments), len(issues)))

if args.issues:
    iw = openpyxl.Workbook(); H = Font(bold=True, color='FFFFFF'); HF = PatternFill('solid', start_color='305496')
    thin = Side(style='thin', color='BFBFBF'); B = Border(left=thin, right=thin, top=thin, bottom=thin); WR = Alignment(wrap_text=True, vertical='top')
    seen, uniq, cnt = set(), [], Counter()
    for i in issues:   # one row per variable for repeated value-level / pure-copy findings
        k = (i[0], i[1], i[2], i[4]) if ('pure copy' in i[4] or i[0] == 'ValueLevel') else tuple(i)
        cnt[k] += 1
        if k not in seen: seen.add(k); uniq.append(list(i) + [k])
    for i in uniq:
        k = i.pop()
        if cnt[k] > 1: i[3] = '%s (+%d more value-level rows)' % (i[3], cnt[k] - 1) if i[3] else '%d rows' % cnt[k]
    S = iw.active; S.title = 'Summary'
    S.append(['%s ADaM define draft - issues' % STUDY]); S['A1'].font = Font(bold=True, size=12)
    S.append(['Spec wins over the P21 skeleton (red font in the define); orange = needs review.'])
    S.append([]); S.append(['Issue', 'Count'])
    for c in S[4]: c.font = H; c.fill = HF
    for k, v in Counter(i[4] for i in uniq).most_common(): S.append([k, v])
    S.column_dimensions['A'].width = 95
    Dt = iw.create_sheet('Issue Details')
    Dt.append(['#', 'Tab', 'Dataset', 'Variable', 'Where Clause', 'Issue', 'P21 / current', 'Value used', 'Spec text', 'Programmer Resolution'])
    for c in Dt[1]: c.font = H; c.fill = HF
    for n, i in enumerate(sorted(uniq, key=lambda x: (x[4], x[0], x[1], x[2])), 1): Dt.append([n] + list(i) + [''])
    for row in Dt.iter_rows(min_row=2):
        for c in row: c.alignment = WR; c.border = B
    for col, w in zip('ABCDEFGHIJ', (5, 11, 10, 12, 28, 55, 18, 18, 60, 28)): Dt.column_dimensions[col].width = w
    Dt.freeze_panes = 'D2'; Dt.auto_filter.ref = Dt.dimensions
    iw.save(args.issues); print('issues ->', args.issues)
