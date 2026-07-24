---
title: Confirm insert-before/mid-split id-reassignment behavior
labels: [wayfinder:task]
status: closed
assignee: user
blocked_by: []
---

## Question

Only insert-_after_'s paraId reassignment (existing paragraph gets a new id, old id moves onto
the new paragraph) is empirically confirmed. Someone with real Word desktop access needs to run
the equivalent tests for:

1. `insertParagraph(..., Word.InsertLocation.before)` — does the same reassignment happen, in
   which direction?
2. Splitting a paragraph mid-text (e.g. pressing Enter partway through) — which half keeps the
   original id?

Capture before/after `getOoxml()` output for each case (a real Word desktop session, not a
decision to make — this is a Task ticket, not Grilling). Report findings as raw before/after
paraId observations; the design-spec update to `word/paraId.ts`'s described behavior is a
follow-up ticket once this lands.

## Checklist (run in real Word desktop)

Only insert-_after_ is confirmed (see design spec's "ParaId stability model"): the existing
paragraph P gets a fresh id, and P's old id moves onto the newly inserted paragraph. Need the
same precision for two more cases.

**Setup (once):** create a doc with 3+ plain paragraphs, distinct text, so paragraphs are easy to
tell apart by content. Get a `getOoxml()` dump (Insert tab → any add-in with `body.getOoxml()`,
or read the raw `.docx` XML via `Open XML SDK Productivity Tool` / unzip `word/document.xml`).
Record each paragraph's `w14:paraId` from this baseline dump.

**Case 1 — insert-before:**

1. Pick a paragraph P (not the first).
2. Insert a new paragraph immediately _before_ P (equivalent of
   `insertParagraph(..., Word.InsertLocation.before)` — in the UI: click at the very start of P,
   press Enter, then Up-arrow and type new text into the now-empty paragraph above).
3. Take another `getOoxml()` dump. Compare every paraId to the baseline.
4. Record: does P keep its original id, or does it change? Does the new paragraph get a brand
   new id, or does it inherit P's old id (mirroring the insert-after mark-shift, just in the
   other direction)?

**Case 2 — mid-paragraph split:**

1. Pick a paragraph Q with enough text to split meaningfully (e.g. "The quick brown fox jumps").
2. Click partway through the text (e.g. right before "fox") and press Enter — this splits Q into
   two paragraphs at the cursor.
3. Take another `getOoxml()` dump. Compare paraIds to baseline.
4. Record: does the _first_ half (before the split point) keep Q's original id, or does the
   _second_ half (after the split point) keep it? Does the other half get a brand new id?

**Report back (paste or describe, doesn't need to be formatted):**

- Baseline paraIds (paragraph text → id, for the paragraphs involved in each case)
- After-insert-before paraIds (same mapping)
- After-mid-split paraIds (same mapping)
- Word version/build used, and whether this was Word desktop (PC) or Mac — spec's confirmed
  insert-after case was desktop; worth noting if this differs by desktop platform too.

## Answer

Confirmed via real Word desktop (PC), using an actual add-in-driven capture (production
`insert_with_text`/`before` call for Case 1; a real manual Enter-keypress split for Case 2 — not
scriptable via Office.js, has to be a live keystroke). Both cases converge on **one coherent
model**, not two separate quirks:

**Case 1 — insert-before:** the anchor paragraph P **keeps its original id unchanged**. The newly
inserted paragraph (placed before P) gets a **brand-new, unrelated id** — it does **not** inherit
P's old id. This is the mirror image of insert-after's behavior in outcome (the "old"/surviving
content keeps its id either way) but **not a symmetric mark-shift** — insert-after moves P's old
id onto the new paragraph; insert-before does not move anything, it just mints a fresh id for the
new paragraph and leaves P untouched.

**Case 2 — mid-paragraph split:** splitting paragraph Q at a text offset behaves exactly like
"shrink Q down to the tail-side text (keeps Q's original id) + insert a new paragraph _before_ it
containing the head-side text (fresh id)" — i.e. **identical mechanics to Case 1**, just triggered
by a split instead of an explicit insert call. The **second half** (text after the split point)
keeps the original id; the **first half** (text before the split point, effectively a new
paragraph) gets a fresh id.

**Combined model for `word/paraId.ts`:** "Insert new content before an existing paragraph" (via
`InsertLocation.before` or via a mid-paragraph split) never touches the reference paragraph's id —
it only mints a fresh id for the new content. Only `InsertLocation.after` performs a mark-shift
(existing paragraph gets a new id, its old id moves onto the new paragraph after it). This is a
directional asymmetry in real Word, not a shim bug to normalize away — the two `InsertLocation`
directions genuinely behave differently and both must be modeled as written.

Trailing-paragraph-mark churn (already confirmed) was reconfirmed as a side effect of every
capture step in this test — its id changed on every single `getOoxml()` read regardless of
mutation, while every real-content id stayed stable across reads with no mutation between them.

Platform: Word desktop (PC). Mac not separately tested — per the design spec, Mac is assumed to
match PC unless evidence says otherwise; no evidence here contradicts that assumption.

**Consistency check (follow-up):** re-ran insert-before against two more anchors in a single
automated pass — a different middle paragraph, and the very first paragraph in the document (an
edge case with no preceding sibling) — plus a fresh insert-after as a sanity baseline. All three
self-reported as expected: both insert-before cases left the anchor's id unchanged and gave the
new paragraph a fresh id (`anchorIdUnchanged: true` both times, including the first-paragraph edge
case — no special-casing needed for "no preceding sibling"), and the insert-after sanity check
reproduced the known mark-shift exactly (`markShiftReproduced: true`: anchor got a new id, the new
paragraph inherited the anchor's old one). Four independent insert-before data points and two
independent insert-after data points now agree, on the same Word build — the model above is not a
one-off observation.
