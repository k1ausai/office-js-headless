---
title: Exact id format (paraId/textId/rsid)
labels: [wayfinder:research]
status: closed
assignee: wayfinder-charting-session
blocked_by: []
---

## Question

Real Word uses 8-hex-char uppercase ids for `w14:paraId`/`w14:textId` (already confirmed). What's
still open: the exact `w:rsid*` format, and whether a single shared rsid value gets stamped across
all paragraphs touched by one mutation (as informally observed) or varies per paragraph. Needs
real Word XML samples — before/after captures of a single edit — to confirm the rsid-stamping
pattern precisely enough for `word/paraId.ts` to generate ids the same way.

## Answer

**`w:rsid*` format:** same shape as `w14:paraId`/`w14:textId` — an 8-hex-character string
(`ST_LongHexNumber` per ECMA-376 §17.15.1.70), rendered uppercase in observed samples (e.g.
`00707319`, `00FB5DF4`), matching the uppercase-hex examples in the general OOXML rsid
documentation. Conceptually a "revision save ID": a per-editing-session marker, not a
per-paragraph content id — later sessions get strictly larger values than earlier ones in the same
file, and the value is random/time-derived rather than sequential-by-paragraph.

**Shared-vs-per-paragraph stamping:** confirmed for the web-platform case — back-to-back
`getOoxml()` calls (no edits between them) each stamp **one single fresh rsid value across
`rsidR`/`rsidRDefault`/`rsidP` for every paragraph in the document** for that call, then a
_different_ single shared value for the next call. It is not varied per paragraph within a call.
This lines up with the general rsid semantics above (one id per session/save event, applied
document-wide) and with this platform's already-documented unconditional per-`getOoxml()`-call id
regeneration (see "ParaId stability model" in the design spec) — rsid churns alongside
paraId/textId as part of that same whole-document regeneration, sharing one value per call rather
than minting one per paragraph.

**Still unconfirmed:** the available evidence for the desktop (`PC`/`Mac`) platform documents the
insert-after paraId mark-shift precisely (old id moves onto the newly inserted paragraph, the
original paragraph gets a fresh id) but does not include the corresponding rsid values for that
same mutation — so whether desktop stamps a single shared rsid across just the paragraphs touched
by one mutation (vs. regenerating more broadly, vs. leaving rsid untouched on plain structural
edits) is **not settled by primary evidence** and remains an open item. The general rsid
session-semantics above make "one shared value per mutation" the more plausible model, but
`word/paraId.ts` should not hard-code that assumption for the desktop platform without either
treating it as best-effort or gathering a direct desktop before/after sample.
