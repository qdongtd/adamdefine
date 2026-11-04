# adam-define-xml

A Claude Code skill for drafting and QC-ing a CDISC **ADaM Define-XML 2.1**
specification workbook (Pinnacle 21 Excel format) for a clinical study submission.
Companion to `sdtm-define-xml`.

Inputs: a P21-generated ADaM define (the skeleton), the compound ADaM spec
(TechData template, one column per study) and the ADaM VLM.

```
python scripts/fill_adam_define.py --skeleton <p21_adam_define.xlsx> --spec <adam_spec.xlsx> \
    --out <filled.xlsx> [--vlm <adam_vlm.xlsx>] [--issues <issues.xlsx>]
python scripts/self_qc_adam.py <filled.xlsx>
```

- Origins Derived / Assigned / Predecessor from the spec (`[r_orig]` overrides);
  text priority `[r_def]` > `[r_prog]` > Conversion Definition > Programming Notes.
- Predecessor `DATASET.VARIABLE` (ADSL core variables via `<COREVAR>`), Method for
  Derived, Comment for Assigned; no CRF pages.
- Value level per PARAMCD from the VLM, else the parent spec row; complete value
  level leaves the Variables row blank (reviewer convention).
- Dataset Class/Structure/Key Variables from the Content tab; ADaMIG standard.
- Spec wins over the P21 skeleton (red font); open items orange; issue workbook
  with evidence per finding.
- `self_qc_adam.py` is a pre-filter — still generate define.xml and run P21.

Requirements: Python 3.12, openpyxl (pyreadstat for codelists from data).

## Layout
```
adam-define-xml/
├── SKILL.md
├── references/adam-rules.md
└── scripts/
    ├── fill_adam_define.py
    └── self_qc_adam.py
```
