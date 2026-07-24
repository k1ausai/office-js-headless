---
title: office-js-headless — implementation readiness
labels: [wayfinder:map]
status: open
---

## Destination

`doc/design-spec/2026-07-23-office-js-headless-design.md` with zero remaining items in its
"Open items for the implementation plan" section — every open decision resolved and recorded,
ready to hand to the `writing-plans` skill to produce an actual implementation plan. This map
does not write any implementation code itself.

## Notes

- Domain: Office.js Word add-in shim, OOXML (Flat-OPC), `@xmldom/xmldom`. See the design spec for
  full context — read it before resolving any ticket here.
- Public repo (`k1ausai/office-js-headless`) — never reference private repos, issues, or PRs by
  name/number in ticket bodies or resolutions, even generically. Same constraint that applies to
  the design spec and `AGENTS.md` applies here.
- Use `/grilling`/`/domain-modeling` for decision tickets, `/research` subagent for research
  tickets.
- No issue tracker wired up yet (`/setup-matt-pocock-skills` not run) — this map lives as local
  markdown under `doc/wayfinder/`, not as GitHub issues. Blocking is a body convention
  (`Blocked by:` in each ticket), not native.

## Decisions so far

- [search() semantics scope](tickets/002-search-semantics-scope.md) — v1 only needs plain-text
  search plus an optional `matchCase` flag; wildcard syntax and other `SearchOptions` are unused
  by the driving consumer and stay out of scope.
- [Exact id format (paraId/textId/rsid)](tickets/005-id-format.md) — `w:rsid*` is the same
  8-hex-char format as `paraId`/`textId`; web platform stamps one shared rsid across all
  paragraphs per `getOoxml()` call. Desktop per-mutation rsid-sharing is still unconfirmed by
  evidence — flagged, not assumed.
- [Table row/cell OOXML shape](tickets/003-table-ooxml-shape.md) — minimal element set is
  `w:tbl`/`w:tblGrid`/`w:gridCol` + `w:tr` + `w:tc`/`w:tcPr`/`w:tcW`, reusing existing
  paragraph/run/text handling for cell content; `gridSpan`/`vMerge` needed for read/round-trip
  (not write) so merged-cell tables don't silently miscount. `w:tblPr`/`w:trPr` omitted for v1.
  Whether real Word's default `tblStyle`/`tblLook` authoring boilerplate needs reproducing is
  blocked on the fixture-capture-process ticket.
- [Unloaded/unsynced-property error message shape](tickets/001-error-message-shape.md) — real
  Office.js does not distinguish "never loaded" from "loaded but unsynced"; both throw the same
  `PropertyNotLoaded`-coded error with one message template. Surfaced a new decision (error
  `.name` value) as its own ticket.
- [Error `.name` value](tickets/007-error-name-value.md) — shim throws `"RichApi.Error"`
  (matches real Word's observed runtime), not the documented-but-unobserved
  `"OfficeExtension.Error"`. No configurable option.
- [Fixture-capture process](tickets/004-fixture-capture-process.md) — full-matrix coverage,
  captured upfront by the maintainer/team with real Word access, adapting the driving consumer's
  existing in-app manual-testing harness rather than building new tooling. Fixtures land via
  normal PR review, no separate approval step.
- [Confirm insert-before/mid-split id-reassignment behavior](tickets/006-insert-before-midsplit-confirm.md)
  — empirically confirmed on Word desktop (PC), reproduced across 4 independent insert-before
  cases and 2 insert-after sanity checks. Insert-before never touches the anchor's id, only mints
  a fresh id for the new paragraph (no mark-shift, unlike insert-after); mid-split behaves
  identically — the tail half keeps the original id, the head half gets a fresh one. One coherent
  model, not two separate quirks.

## Not yet specified

- Exact structural-comparison implementation for fixture assertions (spec says "paragraph
  text/paraId/style-name lists", but the precise comparator shape depends on how the table-OOXML
  and id-format tickets land) — not sharp enough to ticket yet.

## Out of scope

- Downstream QA-runner adapter and Vitest-environment adapter — explicitly out of scope per the
  design spec's Context section (future work, separate packages that npm-install this one).
- Excel/PowerPoint/Outlook object models, real-time collaboration semantics, byte-exact OOXML
  serialization, true `.docx` zip I/O beyond the `insertFileFromBase64` stub — non-goals per the
  design spec.
