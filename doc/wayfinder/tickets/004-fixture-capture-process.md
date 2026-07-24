---
title: Fixture-capture process
labels: [wayfinder:grilling]
status: closed
assignee: wayfinder-session
blocked_by: []
---

## Question

Decide the process for capturing the v1 golden-fixture corpus (real
`{seedOoxml, operation+args, resultOoxml}` triples from actual Word, per the design spec's
Fidelity validation section):

- Who has Word desktop + Word Online access to drive the captures?
- What counts as "v1-complete" coverage — every `InsertLocation` x receiver-type combination plus
  the platform-diverging cases (paraId churn, `insertOoxml`/`insertFileFromBase64` on
  `OfficeOnline`), or a smaller starting set?
- Is there an existing human-driven Word test harness (the spec mentions one may exist) that can
  be reused to record these, or does capture tooling need to be built first?
- Where do captured fixtures land before landing in `fixtures/` — reviewed how, by whom?

## Answer

**Who captures:** the maintainer/team driving this project — real Word desktop and Word Online
access needed, no way to automate this from the shim side.

**v1-complete coverage:** the full matrix — every `InsertLocation` × receiver-type combination,
plus every platform-diverging case already identified in the design spec (`getOoxml()` paraId/
rsid churn, `insertOoxml`/`insertFileFromBase64` behavior on `OfficeOnline`). Matches the design
spec's stated fidelity-validation bar; captured upfront rather than grown incrementally per
operation.

**Capture tooling: adapt existing harness, don't build new.** The driving consumer already has an
in-app manual-testing harness with a `seed → run → snapshot` pattern and a library of realistic
seed-document + operation scenarios — reusing that (swapping its truncated/derived document
snapshot for a raw, untruncated `getOoxml()` dump, and adding a step to persist each
`{seedOoxml, operation+args, resultOoxml}` triple to disk) is cheaper than building separate
capture tooling from scratch in `office-js-headless`, and keeps capture running where a human can
already drive it interactively against real Word (desktop and web). This is downstream/consumer
work, not something implemented inside this public repo.

**Review process:** normal PR review. Captured fixtures land in `fixtures/` via a regular PR
against `office-js-headless`, reviewed the same as any other change — no separate approval step.
