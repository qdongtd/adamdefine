# ADaM define rules and rationale

## Origin types (CDISC Define-XML CT, codelist C170449)

ADaM uses only three of the Origin Type terms:

| Code | Term | CDISC definition (verbatim) |
|---|---|---|
| C170549 | Derived | A value that is calculated by an algorithm or reproducible rule, and which is dependent upon other data values. |
| C170547 | Assigned | A value that is derived through designation, such as values from a look up table or a label on a CRF. |
| C170550 | Predecessor | A value that is copied from another variable. |

Collected and Protocol are not valid in an ADaM define. Define-XML 2.1: Derived and
Assigned take Source = Sponsor; Predecessor takes no Source, and its def:Origin
Description names the copied variable as `DATASET.VARIABLE` (P21 checks the syntax).

## Text sources and markers (TechData ADaM template)

- Study column: `x` in scope; `[r_def]` define text; `[r_prog]` study derivation;
  `[r_orig]` origin override. A cell can concatenate tags (`[r_prog] Set to TRTP.
  [r_def] Set to TRTP.`) — parse each tag to the next `[r_`.
- Priority `[r_def]` › `[r_prog]` › Conversion Definition › Programming Notes.
- The template's Conversion Definition column is "for reference/example only" per
  its Instructions tab, so study-specific `[r_*]` text always wins.
- `<COREVAR>` rows mean ADSL core variables are added to that dataset.

## Value level

- Skeleton value-level rows come from the data (one per PARAMCD per AVAL/AVALC, and
  other per-parameter variables). The VLM may describe only some of them.
- Use the VLM row when it exists; otherwise the parent variable's spec row. Both are
  sourced; the second is reported per variable so the programmer confirms the
  parent text applies to every parameter.
- When the VLM describes a value that has no row in the data, report it (do not
  add value-level rows that the data does not support).
- Reviewer convention: variables whose value level is complete carry no origin
  information on the Variables row.

## Codelists

Lessons from ADaM define reviews: derive codelist terms and coded/decoded pairs from
the real ADaM data, not the spec — numeric codes are not always 1..n (e.g. RACEN 1
and 5), TRTPN/TRTAN codings can differ between datasets, and a spec codelist can
match none of the data values. Read `.sas7bdat` with
`pyreadstat.read_sas7bdat(path, disable_datetime_conversion=True)`; format floats
with `'%.10g'` (107.2, not 107.19999999999999).

## Dataset metadata

The P21 skeleton derives Class/Key Variables from the data (often
`STUDYID, USUBJID, PARAM, PARAMCD, AVAL` for every BDS); the spec's Content tab
carries the intended class, structure and sort keys (`[r_sort]` per study). Apply
the Content tab, record the P21 value, and report every difference.

## Working conventions shared with the SDTM skill

Back up before editing; fill from sources only; orange for review, never for routine
fills; Programmer Review column with original value + reason; PermissionError →
save `_updated.xlsx`; every Method/Comment/Codelist ID must resolve; Method/Comment
wording in plain English (no librefs, SAS code, placeholders, spec-layout
references); report issues in a workbook with the evidence per row, not only in chat.
