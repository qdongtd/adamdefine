# -*- coding: utf-8 -*-
"""Pre-check an ADaM Define-XML 2.1 spec workbook (Pinnacle 21 Excel format) for the
spec-checkable ADaM define rules. A PRE-FILTER, not authoritative: still generate
define.xml and run the P21 Define validation (many rules need the XML, CT or data).

    python self_qc_adam.py <adam_define.xlsx>
"""
import re, sys, warnings
from collections import Counter, defaultdict
import openpyxl
warnings.simplefilter('ignore')

F = sys.argv[1]
wb = openpyxl.load_workbook(F)
def tab(name):
    ws = wb[name]; h = [c.value for c in ws[1]]
    return {n: i for i, n in enumerate(h)}, [r for r in ws.iter_rows(min_row=2, values_only=True) if any(r)]
def txt(x):
    s = '' if x is None else str(x).strip(); return '' if s == 'None' else s
res, ex = {}, {}
def chk(rule, viols): res[rule] = len(viols); ex[rule] = list(viols)[:4]

di, dr = tab('Datasets'); vi, vr = tab('Variables'); li, lr = tab('ValueLevel')
mi, mr = tab('Methods'); ci_, cr = tab('Comments'); cli, clr = tab('Codelists')
M = {r[mi['ID']]: r[mi['Description']] for r in mr if r[mi['ID']]}
C = {r[ci_['ID']]: r[ci_['Description']] for r in cr if r[ci_['ID']]}
ORIG = {'Derived', 'Assigned', 'Predecessor'}
CLASSES = {'SUBJECT LEVEL ANALYSIS DATASET', 'BASIC DATA STRUCTURE', 'OCCURRENCE DATA STRUCTURE', 'ADAM OTHER'}
DATE = {'date', 'datetime', 'time', 'partialDate', 'partialTime', 'partialDatetime', 'incompleteDatetime', 'durationDatetime', 'intervalDatetime'}
NUM = {'integer', 'float', 'text'}
PRED = re.compile(r'^[A-Z][A-Z0-9]{0,7}\.[A-Z][A-Z0-9]{0,7}(\s+where\s+.+)?$', re.I)
BAD = re.compile(r'\b(sdtm|raw|lookup)\.[a-z]|\bcatx\(|\bput\(|\bbest\.|\bstrip\(|\bupcase\(|tab "|"XX"|\[if |\[r_\w+\]', re.I)
def lab(r, ix, vl=False): return '%s.%s%s' % (r[ix['Dataset']], r[ix['Variable']], (' [%s]' % r[ix['Where Clause']]) if vl else '')

# ---- Datasets
chk('ADaM Class valid', ['%s:%s' % (r[di['Dataset']], r[di['Class']]) for r in dr if txt(r[di['Class']]).upper() not in CLASSES])
chk('Structure required', [r[di['Dataset']] for r in dr if not txt(r[di['Structure']])])
chk('Key Variables required', [r[di['Dataset']] for r in dr if not txt(r[di['Key Variables']])])
if 'Standard' in di: chk('Standard column filled', [r[di['Dataset']] for r in dr if not txt(r[di['Standard']])])
chk('ADSL present', [] if any(r[di['Dataset']] == 'ADSL' for r in dr) else ['ADSL missing'])

# ---- Variables + ValueLevel origin rules
vlo, vln = defaultdict(set), Counter()
for r in lr:
    k = (r[li['Dataset']], r[li['Variable']]); vln[k] += 1
    if txt(r[li['Origin']]): vlo[k].add(txt(r[li['Origin']]))
for name, ix, rows, vl in (('var', vi, vr, False), ('VL', li, lr, True)):
    o = lambda r: txt(r[ix['Origin']])
    chk('Origin valid for ADaM (%s)' % name, [lab(r, ix, vl) + ':' + o(r) for r in rows if o(r) and o(r) not in ORIG])
    chk('Derived needs Method (%s)' % name, [lab(r, ix, vl) for r in rows if o(r) == 'Derived' and not txt(r[ix['Method']])])
    chk('Method only on Derived (%s)' % name, [lab(r, ix, vl) for r in rows if o(r) and o(r) != 'Derived' and txt(r[ix['Method']])])
    chk('Predecessor needs value (%s)' % name, [lab(r, ix, vl) for r in rows if o(r) == 'Predecessor' and not txt(r[ix['Predecessor']])])
    chk('Predecessor syntax DATASET.VARIABLE (%s)' % name, [lab(r, ix, vl) + ':' + txt(r[ix['Predecessor']]) for r in rows
        if o(r) == 'Predecessor' and txt(r[ix['Predecessor']]) and not PRED.match(txt(r[ix['Predecessor']]))])
    chk('Predecessor value only with Predecessor origin (%s)' % name, [lab(r, ix, vl) for r in rows if txt(r[ix['Predecessor']]) and o(r) != 'Predecessor'])
    if 'Source' in ix:
        chk('Source=Sponsor for Derived/Assigned (%s)' % name, [lab(r, ix, vl) for r in rows if o(r) in ('Derived', 'Assigned') and txt(r[ix['Source']]) != 'Sponsor'])
        chk('No Source on Predecessor (%s)' % name, [lab(r, ix, vl) for r in rows if o(r) == 'Predecessor' and txt(r[ix['Source']])])
    chk('No CRF Pages in ADaM (%s)' % name, [lab(r, ix, vl) for r in rows if txt(r[ix['Pages']])])
    chk('Dangling Method/Comment (%s)' % name, [lab(r, ix, vl) for r in rows if (txt(r[ix['Method']]) and r[ix['Method']] not in M) or (txt(r[ix['Comment']]) and r[ix['Comment']] not in C)])
chk('Origin required (var, unless value level)', [lab(r, vi) for r in vr if not txt(r[vi['Origin']]) and (r[vi['Dataset']], r[vi['Variable']]) not in vlo])
chk('Origin required (VL)', [lab(r, li, True) for r in lr if not txt(r[li['Origin']])])
chk('Value level complete -> Variables row blank (HT)', ['%s.%s' % k for k in vlo if vln[k] == sum(1 for r in lr if (r[li['Dataset']], r[li['Variable']]) == k and txt(r[li['Origin']]))
    and any(txt(r[vi[f]]) for r in vr if (r[vi['Dataset']], r[vi['Variable']]) == k for f in ('Origin', 'Method', 'Comment', 'Predecessor'))])
chk('DD0074 var/VL origin mismatch', ['%s.%s' % (r[vi['Dataset']], r[vi['Variable']]) for r in vr
    if txt(r[vi['Origin']]) and (r[vi['Dataset']], r[vi['Variable']]) in vlo and {txt(r[vi['Origin']])} != vlo[(r[vi['Dataset']], r[vi['Variable']])]])
chk('Copy text not Predecessor', [lab(r, vi) + ': ' + txt(M.get(r[vi['Method']]) or C.get(r[vi['Comment']])) for r in vr
    if txt(r[vi['Origin']]) in ('Derived', 'Assigned') and re.fullmatch(r'\s*(set\s+to\s+)?[A-Z][A-Z0-9]{1,7}\.[A-Z][A-Z0-9]{1,7}\s*\.?\s*',
    txt(M.get(r[vi['Method']]) or C.get(r[vi['Comment']])), re.I)])

# ---- attributes
chk('DataType required', [lab(r, vi) for r in vr if not txt(r[vi['Data Type']])])
chk('Length integer', [lab(r, vi) + ':' + txt(r[vi['Length']]) for r in vr if txt(r[vi['Length']]) and not txt(r[vi['Length']]).isdigit()])
chk('Length empty on date types', [lab(r, vi) for r in vr if txt(r[vi['Data Type']]) in DATE and txt(r[vi['Length']])])
chk('Length required (text/integer/float)', [lab(r, vi) for r in vr if txt(r[vi['Data Type']]) in NUM and not txt(r[vi['Length']])])
chk('PARAMCD/PARAM present in BDS', [r[di['Dataset']] for r in dr if txt(r[di['Class']]).upper() == 'BASIC DATA STRUCTURE'
    and not {'PARAMCD', 'PARAM'} <= {r2[vi['Variable']] for r2 in vr if r2[vi['Dataset']] == r[di['Dataset']]}])

# ---- wording / codelists
chk('Method wording not reviewer-readable', [k for k, v in M.items() if BAD.search(txt(v)) or txt(v).count('"') % 2])
chk('Comment wording not reviewer-readable', [k for k, v in C.items() if BAD.search(txt(v)) or txt(v).count('"') % 2])
names = {}
for r in clr:
    if r[cli['ID']]: names.setdefault(r[cli['ID']], r[cli['Name']])
chk('Duplicate codelist Name', ['%s(x%d)' % (n, c) for n, c in Counter(names.values()).items() if n and c > 1])
used = {r[vi['Codelist']] for r in vr if r[vi['Codelist']]} | {r[li['Codelist']] for r in lr if r[li['Codelist']]}
dicts = {r[0] for r in wb['Dictionaries'].iter_rows(min_row=2, values_only=True) if r and r[0]} if 'Dictionaries' in wb.sheetnames else set()
chk('Referenced codelist missing', sorted(c for c in used if c not in names and c not in dicts))

viol = sorted(k for k in res if res[k]); clean = sorted(k for k in res if not res[k])
print('=== ADaM self-QC pre-check: %s ===' % F)
print('rules run: %d | clean: %d | with findings: %d   (PRE-FILTER ONLY - still run P21)' % (len(res), len(clean), len(viol)))
for k in viol: print('  X %-48s %4d  e.g. %s' % (k, res[k], ex[k]))
print('CLEAN:', ', '.join(clean))
