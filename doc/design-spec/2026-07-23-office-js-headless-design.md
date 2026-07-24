# office-js-headless — design spec

Date: 2026-07-23
Status: approved (design), implementation-ready — all open items resolved, see
[`doc/wayfinder/map.md`](../wayfinder/map.md) for the resolution history

## Context

Real-Word integration testing for Office.js add-ins has typically meant driving an actual Word
process (desktop or web) via browser/UI automation — a daemon that launches Word desktop with a
GUI session, sideloads a build, and drives it through some control channel. That gets real
fidelity, but it's not CI-container friendly: it needs a dedicated machine with a logged-in GUI
session and is a single point of failure/bottleneck.

Office.js has no headless runtime of its own — it only exists as the bridge Word injects into a
loaded add-in's webview (desktop) or iframe (web). `office-js-headless` closes that gap by
reimplementing the subset of the `Word`/`Office` API surface real add-ins call, backed by a real
in-memory OOXML document model instead of a live Word process. This is meant to **replace** that
kind of real-Word daemon as the integration tier: an add-in's Office.js-facing service code runs
unmodified against it, in a normal CI container, no dedicated machine or GUI Word required.

Downstream integration (a QA-runner adapter speaking whatever job-dispatch protocol a consumer's
CI uses, and a Vitest-environment adapter for a consumer's own test suite) is explicitly **out of
scope for this spec** — this repo's job is the shim itself, callable as a plain library. Both
adapters are future work, downstream, each npm-installing this package.

## Goals

- Reimplement the Office.js `Word`/`Office` surface actually used by real add-in code (scoped
  initially to what a driving real-world consumer's Office.js-facing service layer calls),
  faithfully enough that a real add-in's document-mutation code runs unmodified and produces
  results equivalent to whichever real Word platform (`PC`/`Mac`/`OfficeOnline`) it's told to
  emulate — see "Platform selection".
- Publish as a general-purpose, public, OSS npm package — designed around Office.js's real API
  shape, not any one consumer's shortcuts. The first real-world consumer's needs drive initial
  scope, but isn't special-cased in the design.
- Make fidelity a checkable, ongoing claim (golden-fixture regression suite), not a one-time
  assumption.

## Non-goals (v1)

- Excel, PowerPoint, Outlook object models — Word only.
- _Why_ Word Online's `insertOoxml`/`insertFileFromBase64` fail (Microsoft's internal
  implementation — no browser-side OOXML merge engine) is out of scope. The shim only reproduces
  the _observable contract_: the same platform/capability signals, the same thrown/rejected
  errors, the same paraId churn behavior (see "Platform selection" and "ParaId stability model"
  below) — so a consumer's existing platform-branching logic takes the same branch it would
  against real Word, without the shim needing to model Microsoft's internals.
- Real-time collaboration / multi-user editing semantics.
- Any Office.js surface not called by the driving consumer's code today (tracked changes, content
  controls, custom XML parts beyond what's needed, mail merge). Extend on demand, not
  speculatively.
- Byte-exact OOXML serialization — structural/semantic equivalence is the fidelity bar; real
  Word's precise attribute ordering/whitespace isn't a contract worth chasing.
- True `.docx` zip file I/O, except the one isolated case below.

## Architecture

```
src/
  document/
    FlatOpcDocument.ts   - loads/mutates/serializes the Flat-OPC XML tree
  word/
    run.ts               - Word.run(callback), RequestContext
    proxy.ts             - tracked-object base: .load(), .context.sync(),
                            strict unloaded/unsynced-read errors
    body.ts, range.ts, paragraph.ts, table.ts, comment.ts, style.ts
    insertLocation.ts    - Word.InsertLocation semantics as XML splicing
    paraId.ts            - platform-conditional paraId/rsid/textId churn
                            (see "ParaId stability model")
  office/
    context.ts           - Office.context.platform/requirements/onReady,
                            driven by the install()-time platform option
                            (see "Platform selection")
  operations/             - one file per (receiver type x InsertLocation)
                            combination — not a single generic dispatcher,
                            since Word's semantics genuinely differ per
                            receiver type
  install.ts              - installHeadlessOffice(seedOoxml) -> installs
                            Word/Office as globals, returns a handle with
                            .getOoxml() / .reset() / .dispose()
fixtures/                 - golden {seedOoxml, operation, resultOoxml}
                            triples captured from real Word, per platform
                            where behavior diverges (see "Platform
                            selection")
test/                     - the package's own tests, replaying fixtures
```

## Core document model

`FlatOpcDocument` wraps the `<pkg:package>` XML tree — the same structure real
`body.getOoxml()`/`insertOoxml()` already exchange. No zip/`.docx` binary handling is needed for
the OOXML-fragment paths; everything stays in XML-string land, matching what Office.js's own
bridge does for these calls.

**XML library: `@xmldom/xmldom`** (`DOMParser`/`XMLSerializer`) — pure-XML, lightweight, mature.
Not `jsdom` (HTML/browser environment emulator, unnecessary weight for a test-package dependency).
Not `linkedom` (nicer `querySelector` ergonomics, but optimized for HTML-like trees rather than
strict namespaced XML). `slimdom` + `fontoxpath` (DOM4-spec XML, real XPath 3.1 queries) was
considered as a more declarative alternative for tree navigation but rejected for v1 in favor of
the more familiar xmldom-style API.

**Responsibilities:**

- Parse the seed Flat-OPC XML into a live DOM tree; keep a handle to `<w:body>` (main
  `document.xml` part) and lazily locate `comments.xml`/`styles.xml`/`numbering.xml` parts if
  present.
- Low-level mutation primitives used by the Word shim: `insertAt(targetNode, fragment, location)`,
  `deleteNode`, `search(pattern)`.
- `getOoxml(scopeNode?)`: serializes either the whole package (`body.getOoxml()`) or a
  range-scoped fragment re-wrapped as its own Flat-OPC package (`Range.getOoxml()`'s real shape).

## Word.run mechanics — batching, proxies, InsertLocation

**Deferred execution, not eager.** Real Office.js queues every mutating call (`insertText`,
`delete`, `insertOoxml`, …) into a batch; nothing touches the document until `context.sync()`.
Errors from a queued operation surface as the rejection of `sync()`, not at the call site — real
add-in code relies on this (wraps `sync()`, not individual calls, in try/catch). The shim
replicates this genuinely: each mutating call appends a pending op to the `RequestContext`'s
queue; `context.sync()` replays the queue against `FlatOpcDocument` in order, and any failure
becomes that `sync()` call's rejection.

**Strict load/track gating on reads.** Every shim object (`Body`, `Range`, `Paragraph`, `Table`,
…) tracks which properties were named in `.load(...)`. Reading a property that wasn't loaded, or
was loaded but not yet followed by a `sync()`, throws — real Office.js does not distinguish these
two cases either; both throw the identical `PropertyNotLoaded`-coded error, traced from the actual
shipped runtime (not just its published types, which document a different, unobserved error name):

```ts
{
  name: "RichApi.Error", // matches real Word's observed runtime, not the documented-but-unobserved "OfficeExtension.Error"
  code: "PropertyNotLoaded",
  message: `The property '${propertyName}' is not available. Before reading the property's value, call the load method on the containing object and call "context.sync()" on the associated request context.`,
}
```

This is what makes the shim catch real add-in bugs (missing `.load()`/`.sync()`) rather than
silently tolerating them.

**InsertLocation is receiver-dependent**, same as real Office.js: on a `Range`/`Paragraph`,
`Before`/`After` splice sibling paragraphs, `Replace` swaps the target's content, `Start`/`End`
insert within it. On `Body`, `Start`/`End` mean first/last child, `Replace` clears and replaces
the whole body, `Before`/`After` don't apply. Each receiver's insert methods get their own small
splice-mapping in `operations/`.

## Platform selection

Office.js's actual behavior — not just its types — differs by host platform, and a meaningful
share of real add-in code branches on it. `installHeadlessOffice` takes a `platform` option:
`"PC" | "Mac" | "OfficeOnline"` (the `Office.PlatformType` values that matter for Word; other
values like `iOS`/`Universal` aren't modeled — out of scope, nothing in the driving consumer's
code branches on them). Default: `"PC"`.

The platform choice drives every host-observable signal a consumer might branch on, not just
`Office.context.platform` itself:

| Signal                                                                | PC / Mac                                                  | OfficeOnline                                                                                                                                                                                             |
| --------------------------------------------------------------------- | --------------------------------------------------------- | -------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| `Office.context.platform`                                             | `Office.PlatformType.PC` / `.Mac`                         | `Office.PlatformType.OfficeOnline`                                                                                                                                                                       |
| `Office.context.requirements.isSetSupported("WordApiDesktop", "1.1")` | `true`                                                    | `false`                                                                                                                                                                                                  |
| `Range/Body.insertOoxml(...)`                                         | applies the fragment                                      | throws an `OfficeExtension.Error` (real Word Online surfaces `ooxmlIsMalformed` / `documentNotSupported` / `GeneralException` here — a genuine, unfixed Microsoft product bug, not a missing capability) |
| `Body.insertFileFromBase64(...)`                                      | applies the template import (approximate stub, see above) | rejects — Word Online has no client-side implementation of this call at all                                                                                                                              |
| `Range/Body.insertHtml(...)`                                          | applies                                                   | applies — this path is fully supported on the web and is real add-in code's actual fallback route                                                                                                        |
| Native tables, comments, `search`                                     | apply                                                     | apply — also fully web-supported, no divergence to model                                                                                                                                                 |
| `getOoxml()` paraId/rsid/textId stability                             | see "ParaId stability model"                              | see "ParaId stability model" — web churns unconditionally                                                                                                                                                |

Both `Office.context.platform` and `.requirements` must read as `null`/throw before
`Office.onReady` has fired, matching the real host (they're populated only once the runtime
library has loaded) — a consumer reading them at module-eval time instead of inside/after
`onReady` is a real bug class the strict-gating design (above) already exists to catch, and this
extends the same principle to platform signals.

This is what makes "replace the desktop daemon" actually true for consumers whose code branches on
capability rather than hardcoding a platform assumption: their existing `isSetSupported(...)
? desktopPath : webFallbackPath` branches run unmodified and take the real branch for whichever
platform the test asks for — the shim doesn't need to know that convention exists, it just needs
to make the underlying signals and call outcomes match reality.

## ParaId stability model

`w14:paraId` (and `w:rsid*`/`w14:textId`) stability is not a simple invariant — real Word violates
it in specific, platform-dependent ways, and a shim that treats ids as permanently stable would
pass tests that fail against real Word. This needs deliberate modeling, not just "don't churn
ids unless told to."

**Id format.** `w14:paraId`/`w14:textId` are 8-hex-char uppercase strings (`ST_LongHexNumber`,
ECMA-376 §17.15.1.70). `w:rsid*` uses the same 8-hex-char format — conceptually a per-editing-
session "revision save id," not a per-paragraph content id. On `OfficeOnline`, one shared rsid
value is stamped across every paragraph's `rsidR`/`rsidRDefault`/`rsidP` per `getOoxml()` call,
churning alongside paraId/textId as part of that platform's whole-document regeneration (below).
On `PC`/`Mac`, whether a single shared rsid is stamped per mutation (vs. varying per paragraph,
vs. left untouched on plain structural edits) is not settled by available evidence —
`word/paraId.ts` should treat desktop rsid stamping as best-effort rather than hard-coding an
assumption, pending a direct desktop before/after capture.

**On `OfficeOnline`:** every `getOoxml()` call — on `Body` or any `Range` — regenerates fresh
`w14:paraId`, `w:rsid*`, and `w14:textId` values for the whole document before serializing, even
with zero mutations between two calls. Two consecutive reads of an unmodified document return the
same paragraphs at the same indices with completely different ids. `w14:textId` is frequently a
non-unique placeholder value on this platform rather than a real per-paragraph id. **No id
returned by `getOoxml()` is a valid cross-call anchor on this platform** — the only stable
cross-call handles are text content, paragraph order/index, and style. This must be modeled as an
unconditional side effect of calling `getOoxml()` on this platform, not tied to any mutation.

**On `PC`/`Mac`:** ids are stable across reads and across most mutations (plain text edits,
restyling, paragraph deletion all leave surviving paragraphs' ids untouched) — but not
unconditionally, in two specific ways real Word exhibits:

1. **Insert-after and insert-before are directionally asymmetric — only `after` performs a
   mark-shift.** Inserting a new paragraph immediately after an existing paragraph P
   (`insertParagraph(..., Word.InsertLocation.after)`, or any operation that splits/appends a
   following paragraph) gives **P** a freshly generated id and moves **P's old id onto the newly
   inserted paragraph**. This is materially different from simple invalidation: a lookup by P's
   old id after this operation resolves to the _new_ paragraph, not P — silently targeting the
   wrong content rather than failing to resolve. Deleting the newly inserted paragraph afterwards
   does not undo this — P keeps its new id permanently.

   **Insert-_before_ does not mirror this.** Inserting a new paragraph immediately before P leaves
   P's id completely untouched and simply mints a fresh, unrelated id for the new paragraph —
   confirmed identically whether P is a middle paragraph or the very first paragraph in the
   document (no edge-case special behavior for "no preceding sibling").

   **Mid-paragraph split** (splitting a paragraph Q at a text offset, e.g. pressing Enter partway
   through) follows insert-before's mechanics, not insert-after's: the tail half (text after the
   split point) keeps Q's original id; the head half (text before the split point, effectively a
   newly-created paragraph) gets a fresh id.

   One coherent model covers all three cases: inserting new content _before_ a reference paragraph
   never touches that paragraph's id, only `after` performs the mark-shift. Confirmed on Word
   desktop (PC) across multiple independent insertions (middle paragraph, first-paragraph edge
   case, and a mid-split), plus a sanity re-check that the insert-after mark-shift still reproduces
   on the same build. Mac not separately tested — assumed to match PC, no contrary evidence.

2. **The document's trailing paragraph mark has no stable id, ever.** `getOoxml()`'s output
   contains one more id-bearing paragraph mark than the `body.paragraphs` collection reports — a
   final empty paragraph — and that mark's id regenerates on _every_ `getOoxml()` call regardless
   of whether anything was edited. Nothing should ever resolve or address content by this trailing
   mark's id.

Both platform behaviors are implemented in `word/paraId.ts`, invoked from `FlatOpcDocument`'s
`getOoxml()` and from the `InsertLocation.after`-splicing operation, gated on the `platform` option
from `installHeadlessOffice`.

## v1 API coverage

Driven by what a real-world consumer's Office.js-facing service layer actually calls:

- `Word.run`, `RequestContext`, `.load()`/`context.sync()`
- `Office.context.platform`, `Office.PlatformType`, `Office.context.requirements.isSetSupported`,
  `Office.onReady` — reporting capabilities per the `installHeadlessOffice` `platform` option
  (default `"PC"`), see "Platform selection"
- `Body`: `getOoxml`, `insertOoxml`, `insertHtml`, `insertText`, `insertParagraph`, `.paragraphs`,
  `.tables`, `getComments`, `getStyles`
- `Range`: `getOoxml`, `insertOoxml`, `insertText`, `delete`, `select`, `search` (plain-text only,
  with an optional `matchCase` flag — wildcard syntax and other `SearchOptions` are unused by the
  driving consumer and out of scope for v1), `getRange`
- `Paragraph`/`ParagraphCollection`: `getFirst`, `getLast`, `getRange`
- Tables, `getComments`, `getStyles`/`Word.BuiltInStyleName`
- `Word.InsertLocation` — all 5 values
- `insertFileFromBase64` — approximate stub, see below

### Table OOXML shape

Minimal element set for `word/table.ts`/`operations/`, scoped to row/cell insert, cell-text
read/write, and row/column count (per ECMA-376 §17.4 "Tables"):

| Level        | Elements                                              | Attributes needed                                                     |
| ------------ | ----------------------------------------------------- | --------------------------------------------------------------------- |
| Table        | `w:tbl`, `w:tblGrid`, `w:gridCol`                     | `w:gridCol/@w:w`                                                      |
| Row          | `w:tr`                                                | none — no `w:trPr` needed for v1 scope                                |
| Cell         | `w:tc`, `w:tcPr`, `w:tcW`, `w:gridSpan`?, `w:vMerge`? | `w:tcW/@w:w`, `w:tcW/@w:type`, `w:gridSpan/@w:val`, `w:vMerge/@w:val` |
| Cell content | `w:p`, `w:r`, `w:t`                                   | identical to existing paragraph/run/text handling                     |

`w:tblGrid` is schema-optional but treated as always-emitted here — it's the only place true
column count lives independent of per-row cell counts under `gridSpan`/merges. `w:tblPr`/`w:trPr`
are omitted entirely for v1; nothing in the covered API surface needs table-wide or row-wide
formatting. `gridSpan`/`vMerge` are read/round-tripped (not written), so a table with pre-existing
merges in seed OOXML doesn't silently miscount cells — a `mergeCells()`-equivalent write API is
deferred until a consumer call site needs it, per the "extend on demand" non-goal. Cell text nests
identically to body paragraphs (`w:tc > w:p > w:r > w:t`) — reuses `word/paragraph.ts`'s existing
construction rather than re-deriving it. Element order within `w:tcPr` is schema-mandated (an XSD
sequence), unlike attribute order/whitespace — `word/table.ts` must build cell properties in the
documented sequence. Whether real Word's default `tblStyle`/`tblLook` authoring boilerplate needs
reproducing for fixture-equivalence is unresolved without a captured fixture (see "Fidelity
validation").

### `insertFileFromBase64` — deliberate fidelity gap

Real Word's `insertFileFromBase64` imports a whole `.docx` (zip) and merges its
`styles.xml`/`numbering.xml` into the current document using an internal, undocumented merge
algorithm. Faithfully replicating that merge (id-collision handling, etc.) is disproportionate
effort for one call site. v1 stubs it: unzip the `.docx` (via `jszip` — the one place this package
touches real zip/binary `.docx` I/O), copy `styles.xml`/`numbering.xml` entries in without real
conflict resolution, and document the gap explicitly (the same pattern used for other host-specific
parity gaps). Revisit if it causes false positives/negatives in practice.

## Fidelity validation

The shim's entire value is "behaves like real Word," so that claim needs ongoing proof:

- **Golden corpus**: for each supported operation (each `InsertLocation` x receiver-type
  combination, `insertText`, `delete`, `search`, table ops, comments, styles, the
  `insertFileFromBase64` stub), a real `{seedOoxml, operation+args, resultOoxml}` triple captured
  from actual Word. Desktop-only operations need one capture; anything touching paraIds or the
  OOXML-write paths (`insertOoxml`, `insertFileFromBase64`) needs a capture **per platform**
  (`PC`/`Mac` can likely share one — `Mac` isn't known to diverge from `PC` on any covered
  operation — and `OfficeOnline` separately), since "Platform selection" and "ParaId stability
  model" above are real behavioral divergences, not implementation details. Full-matrix coverage,
  captured upfront by the maintainer/team (real Word desktop + Word Online access required), by
  adapting an existing in-app manual-testing harness the driving consumer already has — swapping
  its truncated/derived document snapshot for a raw, untruncated `getOoxml()` dump and adding a
  step to persist each captured triple to disk — rather than building separate capture tooling in
  this repo. Fixtures land via normal PR review against `office-js-headless`, same as any other
  change; no separate approval step.
- **Contract tests**: this package's CI suite replays each fixture's seed+operation through the
  shim and asserts the result matches the captured real-Word output, compared structurally
  (paragraph text/table cell text/style-name lists, paraId well-formedness), not byte-for-byte.
  paraId is checked for format validity (8-hex-uppercase, present on every real paragraph), not
  exact-value list equality against the fixture — this shim generates fresh random ids for every
  newly-created paragraph, so a synthetic (hand-written, pre-#21) fixture has no "real" captured
  id to diff a value against. Exact-value comparison for paragraphs a real capture confirms should
  stay STABLE (not newly created by the operation under test) becomes meaningful once #21's
  real-Word captures land, and can be added to the comparator then (see `test/structuralCompare.ts`
  for the harness built in #20).
- **Closing the loop on future gaps**: a divergence found downstream (in a consumer's pipeline)
  becomes a new golden fixture — capture real Word's actual output once, add it to the corpus, fix
  the shim to match. Production-found gaps become permanent regression tests.

## Consumer-facing API

Real Office.js code references bare global `Word`/`Office` — add-in code can't be run unmodified
any other way — so installation mutates globals, same as the real host injects them:

```ts
import { installHeadlessOffice } from "office-js-headless";

const office = installHeadlessOffice({
  seedOoxml: fixtureXml,
  platform: "OfficeOnline", // "PC" | "Mac" | "OfficeOnline", default "PC"
});
// Word/Office now exist as globals in this process, with platform.context,
// requirements, and getOoxml()'s paraId-churn behavior all matching that
// platform (see "Platform selection" / "ParaId stability model")

await Word.run(async (context) => {
  context.document.body.insertText("hello", Word.InsertLocation.end);
  await context.sync();
});

office.getOoxml(); // escape hatch for assertions — real Word has no equivalent
office.dispose(); // tears down globals
```

**Caveat:** global installation means one active document per process at a time — concurrent
documents need separate workers/processes.

## Remaining implementation-time caveats

Everything that blocked implementation from starting has been resolved (see
[`doc/wayfinder/map.md`](../wayfinder/map.md) for the full resolution history). Two narrow,
non-blocking items are called out at their respective sections above for implementation-time
attention rather than repeated here: `PC`/`Mac` per-mutation `w:rsid*` stamping is best-effort, not
confirmed (see "ParaId stability model"), and real Word's default `tblStyle`/`tblLook` authoring
boilerplate for tables is unconfirmed pending a captured fixture (see "Table OOXML shape").
