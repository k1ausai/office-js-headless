# office-js-headless v1 — PRD

Date: 2026-07-24
Source design: [`doc/design-spec/2026-07-23-office-js-headless-design.md`](../design-spec/2026-07-23-office-js-headless-design.md)
Resolution history: [`doc/wayfinder/map.md`](../wayfinder/map.md)

## Problem Statement

Add-in engineers who want to test how their Word add-in's document-mutation logic actually
behaves have had exactly one real-fidelity option: drive an actual Word process (desktop or web)
via browser/UI automation. That means a dedicated machine with a logged-in GUI session, a
sideloaded build, and some control channel to drive it — slow, GUI-session-bound, a single point
of failure, and fundamentally not something a normal CI container can run. Office.js itself has no
headless runtime; it only exists as the bridge a real Word host injects into a loaded add-in's
webview or iframe. Without a headless option, teams either skip real-fidelity testing of their
document-mutation code entirely, or pay the cost of a real-Word daemon tier just to get it.

## Solution

`office-js-headless` reimplements the subset of the `Word`/`Office` API surface that real add-ins
actually call, backed by a real in-memory OOXML (Flat-OPC) document model instead of a live Word
process. An add-in's Office.js-facing service code runs completely unmodified against it — same
`Word.run`, same `context.document.body`, same global `Word`/`Office` objects — in any Node
process, including a normal CI container. It emulates real Word's platform-specific behavior
(`PC`/`Mac`/`OfficeOnline`) faithfully enough — down to specific, empirically-confirmed quirks
like paraId mark-shift on insert-after — that a consumer's existing platform-branching code takes
the same branch it would against the real thing, and a bug that would only surface against real
Word (a missing `.load()`/`.sync()`, a stale id assumption) gets caught here too.

## User Stories

1. As an add-in engineer, I want to run my Office.js-facing service code against a headless Word
   document, so that I can write fast unit/integration tests without a real Word process.
2. As a CI pipeline, I want tests exercising `Word.run` to execute in a normal Node container, so
   that I don't need a dedicated GUI machine running Word.
3. As an add-in engineer, I want `installHeadlessOffice(seedOoxml)` to install `Word`/`Office` as
   globals, so that my add-in's unmodified production code can call bare `Word.run` without any
   test-specific shims.
4. As an add-in engineer, I want to specify which platform (`PC`/`Mac`/`OfficeOnline`) the shim
   emulates, so that I can test platform-specific behavior and fallback logic.
5. As an add-in engineer, I want `Office.context.platform` and
   `Office.context.requirements.isSetSupported` to report accurately per the chosen platform, so
   that my code's capability-branching logic (`isSetSupported(...) ? desktopPath : webFallbackPath`)
   takes the real branch.
6. As an add-in engineer, I want `Office.context.platform`/`.requirements` to read as
   `null`/throw before `Office.onReady` fires, so that a bug where I read them too early is caught
   by my tests just like it would be against real Word.
7. As an add-in engineer, I want `body.insertOoxml`/`insertFileFromBase64` to throw/reject on
   `OfficeOnline` exactly as real Word Online does, so that my platform-branching code around
   those calls is exercised correctly.
8. As an add-in engineer, I want `insertHtml`, tables, comments, and `search` to work identically
   regardless of platform, so that I don't need platform-specific test setup for those calls.
9. As an add-in engineer, I want `Word.run` to defer all mutating calls until `context.sync()`, so
   that errors from queued operations surface at `sync()` exactly like in real Office.js.
10. As an add-in engineer, I want reading an unloaded or not-yet-synced property to throw a
    `PropertyNotLoaded` error matching real Word's shape, so that my tests catch missing
    `.load()`/`.sync()` bugs.
11. As an add-in engineer, I want the thrown error's `.name` to be `"RichApi.Error"` (matching the
    observed real Word runtime), so that error-handling code tested against the shim matches what
    happens against real Word desktop.
12. As an add-in engineer, I want `Word.InsertLocation` semantics to differ correctly by receiver
    type (`Body` vs `Range` vs `Paragraph`), so that insert calls behave exactly as they would in
    real Word.
13. As an add-in engineer, I want `w14:paraId` values to be stable across reads and most mutations
    on `PC`/`Mac`, so that my paraId-addressed logic behaves predictably in tests.
14. As an add-in engineer, I want insert-after to reassign the existing paragraph's id onto the new
    paragraph (mark-shift) exactly as real Word does, so that a paraId-lookup bug after an
    insert-after operation is caught by my tests, not hidden.
15. As an add-in engineer, I want insert-before and mid-paragraph splits to leave the reference
    paragraph's id untouched (minting a fresh id for new content only), so that my tests reflect
    real Word's actual, directionally-asymmetric id-reassignment behavior.
16. As an add-in engineer, I want the document's trailing paragraph mark's id to churn on every
    `getOoxml()` call, so that I don't accidentally rely on it as a stable identifier in my tests.
17. As an add-in engineer testing `OfficeOnline` behavior, I want every `getOoxml()` call to
    regenerate all paraId/rsid/textId values unconditionally, so that I'm forced to write
    web-safe code that never treats ids as stable cross-call anchors.
18. As an add-in engineer, I want `Range`/`Body.search()` to support plain-text matching with an
    optional `matchCase` flag, so that I can test my grep/search-based tooling.
19. As an add-in engineer, I want `Table`/`TableRow`/`TableCell` proxy objects to support row/cell
    insert, cell-text read/write, and row/column count against realistic OOXML table structures,
    so that I can test table-manipulating add-in code.
20. As an add-in engineer, I want tables with pre-existing merged cells (`gridSpan`/`vMerge`) in
    seed OOXML to report correct row/column counts, so that my code doesn't silently miscount
    cells.
21. As an add-in engineer, I want `insertFileFromBase64` to apply an approximate stub import of
    `styles.xml`/`numbering.xml`, so that I can exercise the template-import code path even though
    a full-fidelity merge isn't implemented.
22. As an OSS maintainer, I want a golden-fixture regression corpus of
    `{seedOoxml, operation, resultOoxml}` triples captured from real Word, so that fidelity is an
    ongoing, checkable claim rather than a one-time assumption.
23. As an OSS maintainer, I want the shim's CI suite to replay each fixture and assert structural
    (not byte-exact) equivalence, so that regressions are caught automatically.
24. As an OSS maintainer, I want a production-found fidelity gap to become a new golden fixture
    before the shim is fixed, so that every real-world divergence becomes a permanent regression
    test.
25. As a package consumer, I want `office.getOoxml()`/`office.dispose()` escape hatches, so that I
    can assert on document state and clean up between tests.
26. As a package consumer, I want to know that global installation supports only one active
    document per process at a time, so that I structure concurrent test runs across separate
    workers/processes.
27. As a future downstream integrator (a QA-runner adapter, a Vitest-environment adapter), I want
    `office-js-headless` to be a plain, general-purpose library with no consumer-specific
    special-casing, so that I can build adapters on top of it without forking or patching the core
    package.

## Implementation Decisions

- **Backing model**: a live Flat-OPC XML tree via `@xmldom/xmldom`. `jszip` is used only for the
  one isolated `insertFileFromBase64` stub, which needs to peek inside a real `.docx` zip.
- **Batching**: `Word.run`/`context.sync()` use real deferred execution — mutating calls queue,
  nothing touches the document until `sync()`, and a queued operation's failure surfaces as that
  `sync()` call's rejection, not at the call site.
- **Load/sync gating**: reading an unloaded or loaded-but-unsynced property throws. Real Office.js
  does not distinguish these two cases either — both throw the identical, confirmed error shape
  (traced from the shipped runtime, not just its published types):

  ```ts
  {
    name: "RichApi.Error", // matches observed real Word runtime, not the documented-but-unobserved "OfficeExtension.Error"
    code: "PropertyNotLoaded",
    message: `The property '${propertyName}' is not available. Before reading the property's value, call the load method on the containing object and call "context.sync()" on the associated request context.`,
  }
  ```

- **InsertLocation is receiver-dependent**: on `Range`/`Paragraph`, `Before`/`After` splice sibling
  paragraphs, `Replace` swaps the target's content, `Start`/`End` insert within it. On `Body`,
  `Start`/`End` mean first/last child, `Replace` clears and replaces the whole body,
  `Before`/`After` don't apply.
- **Platform option** (`"PC" | "Mac" | "OfficeOnline"`, default `"PC"`) drives every
  host-observable signal a consumer might branch on: `Office.context.platform`,
  `requirements.isSetSupported("WordApiDesktop", "1.1")`, `insertOoxml`/`insertFileFromBase64`
  throw/reject behavior on `OfficeOnline`, and paraId/rsid/textId churn behavior. Both
  `Office.context.platform` and `.requirements` read as `null`/throw before `Office.onReady` has
  fired.
- **ParaId stability model**:
  - `OfficeOnline`: every `getOoxml()` call regenerates all paraId/rsid/textId values for the
    whole document unconditionally, even with zero mutations between calls. No id returned by
    `getOoxml()` is a valid cross-call anchor on this platform — only text, order, and style are.
  - `PC`/`Mac`: ids are stable across reads and most mutations, except (1) insert-after reassigns
    the _existing_ paragraph's id onto the newly inserted paragraph (a permanent mark-shift, not
    simple invalidation), while insert-before and mid-paragraph splits do **not** mirror this —
    both leave the reference paragraph's id untouched and mint a fresh id only for the new
    content, confirmed empirically across multiple independent positions including the
    first-paragraph edge case; and (2) the document's trailing paragraph mark's id churns on every
    `getOoxml()` call regardless of edits.
  - Id format: 8-hex-char uppercase strings for `paraId`/`textId`/`rsid*`
    (`ST_LongHexNumber`, ECMA-376 §17.15.1.70). Desktop's exact per-mutation rsid-sharing pattern
    is not fully confirmed by evidence — implement as best-effort, not a hard-coded assumption.
- **`search()` scope**: plain-text matching plus an optional `matchCase` flag. Wildcard syntax and
  other `SearchOptions` are unused by the driving consumer and out of scope for v1.
- **Table OOXML shape**: minimal element set — `w:tbl`/`w:tblGrid`/`w:gridCol` (table level),
  `w:tr` (row level, no `w:trPr` needed), `w:tc`/`w:tcPr`/`w:tcW`/`w:gridSpan`?/`w:vMerge`? (cell
  level), reusing existing paragraph/run/text handling for cell content (`w:tc > w:p > w:r > w:t`).
  `w:tblPr`/`w:trPr` are omitted entirely for v1. `gridSpan`/`vMerge` are read/round-tripped (not
  written) so tables with pre-existing merges don't silently miscount cells; a
  `mergeCells()`-equivalent write API is deferred until a consumer call site needs it. Cell
  property element order is schema-mandated (an XSD sequence), not stylistic.
- **`insertFileFromBase64`**: a deliberate fidelity gap. Unzips the `.docx` via `jszip`, copies
  `styles.xml`/`numbering.xml` entries in without real conflict resolution, documented explicitly
  as an approximate stub rather than a full merge-algorithm reimplementation.
- **Fixture-capture process**: full-matrix coverage (every `InsertLocation` × receiver-type
  combination, plus every platform-diverging case) captured upfront by the maintainer/team, using
  real Word desktop and Word Online access, by adapting an existing in-app manual-testing harness
  the driving consumer already has (swapping its truncated/derived document snapshot for a raw,
  untruncated `getOoxml()` dump and adding a persist-to-disk step) rather than building separate
  capture tooling in this repo. Fixtures land via normal PR review, no separate approval process.
- **Fidelity bar**: structural/semantic equivalence, not byte-exact serialization — real Word's
  precise attribute ordering/whitespace isn't chased, except where element _order_ is
  schema-mandated (e.g. table cell properties).
- **Scope discipline**: Word only for v1 (no Excel/PowerPoint/Outlook); no real-time collaboration;
  only the Office.js surface the driving consumer's code actually calls today — extended on
  demand, not speculatively.
- Downstream integration (a QA-runner adapter, a Vitest-environment adapter) is explicitly out of
  scope for this package — future work, separate packages that each npm-install this one.

## Testing Decisions

A good test here asserts observable behavior — the error actually thrown, the resulting OOXML or
paragraph state a consumer would read back — never internal implementation details like a private
queue's in-memory shape.

- **Golden-fixture regression suite** is the primary fidelity-proof mechanism, and the main test
  surface for `FlatOpcDocument` and the paraId/rsid churn logic: each fixture is a real
  `{seedOoxml, operation+args, resultOoxml}` triple captured from actual Word (per platform where
  behavior diverges), replayed in CI and compared structurally (paragraph text/paraId/style-name
  lists), not byte-for-byte. A divergence found downstream becomes a new fixture _before_ the shim
  is fixed — production-found gaps become permanent regression tests.
- **Unit tests** (not fixture-backed) for the remaining deep modules:
  - The load/sync gating logic: assert the exact `PropertyNotLoaded` error shape (name, code,
    message) fires identically for both "never loaded" and "loaded but unsynced" reads, and that a
    correctly loaded-and-synced read succeeds.
  - The InsertLocation → XML splice mapping: assert each receiver type's mapping (`Body`
    `Start`/`End`/`Replace`; `Range`/`Paragraph` `Before`/`After`/`Start`/`End`/`Replace`)
    independently of full `Word.run` wiring.
  - The deferred-batching logic: assert queued operations don't touch the document until
    `sync()`, and that a queued operation's failure surfaces as `sync()`'s rejection rather than
    at the call site.
- **Prior art**: none yet in this repo (pre-implementation) — the fixture-replay pattern described
  above is the primary regression-testing convention going forward, per the design spec's
  "Fidelity validation" section.

## Out of Scope

- Excel, PowerPoint, Outlook object models.
- Real-time collaboration / multi-user editing semantics.
- Byte-exact OOXML serialization (attribute ordering, whitespace) outside schema-mandated element
  order.
- True `.docx` zip file I/O, except the one isolated `insertFileFromBase64` stub case.
- Any Office.js surface not called by the driving consumer's code today (tracked changes, content
  controls, custom XML parts beyond what's needed, mail merge) — extend on demand only.
- The downstream QA-runner adapter and Vitest-environment adapter — future, separate packages.
- `iOS`/`Universal` platform types — not modeled, nothing in the driving consumer's code branches
  on them.
- A full-fidelity `insertFileFromBase64` merge algorithm (id-collision handling, etc.) — a
  deliberate approximate stub instead.
- A `mergeCells()`-equivalent write API for tables — read/round-trip support only for v1.
- Exact confirmation of desktop per-mutation rsid-sharing and real Word's default
  `tblStyle`/`tblLook` authoring boilerplate — both flagged as best-effort/pending-fixture-capture,
  not blocking.

## Further Notes

- Full design rationale, rejected alternatives (`jsdom`/`linkedom`/`slimdom`+`fontoxpath` for the
  XML library), and the empirical evidence behind every ParaId/platform claim live in
  `doc/design-spec/2026-07-23-office-js-headless-design.md`.
- The resolution history for every decision above — research findings, grilling transcripts,
  wayfinder tickets — lives in `doc/wayfinder/` (`doc/wayfinder/map.md` is the index).
- Package is public (`k1ausai/office-js-headless`), general-purpose by design — not special-cased
  for any one consumer, though initial API scope is driven by one real-world consumer's needs.
- No issue tracker/triage label vocabulary was configured for this session
  (`/setup-matt-pocock-skills` not run), so this PRD is saved as a local file rather than published
  to a tracker issue. Publish manually once a tracker is set up, if desired.
