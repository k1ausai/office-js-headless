# Contributing

## Setup

```bash
pnpm install
pnpm test
```

## Before opening a PR

```bash
pnpm lint
pnpm format:check
pnpm type:check
pnpm test:coverage
pnpm build
pnpm pack:check   # validates the actual publishable package (publint + attw)
```

All of the above run in CI; a red check blocks merge.

## Commit messages

Enforced by commitlint ([Conventional Commits](https://www.conventionalcommits.org/)):
`type(scope?): subject`, e.g. `fix(word-run): queue sync errors instead of throwing eagerly`.
Allowed types: `feat`, `fix`, `docs`, `style`, `refactor`, `perf`, `test`, `build`, `ci`, `chore`,
`revert`. This isn't just style — release automation derives version bumps and the changelog from
these.

## Design changes

Read [`doc/design-spec/`](./doc/design-spec) before changing anything in `src/word/`,
`src/office/`, or `src/document/` — several non-obvious decisions are recorded there (deferred
batching, strict `.load()`/`.sync()` gating, structural-not-byte-exact fidelity, desktop-only
scope). If a change contradicts one of those decisions, update the spec in the same PR and explain
why.

## Fidelity fixtures

If you find a case where the shim's output diverges from real Word, don't just patch the shim —
add a new golden fixture under `fixtures/` capturing real Word's actual output for that case first
(see the design spec's "Fidelity validation" section), then fix the shim to match it. This keeps
every found gap as a permanent regression test.

### Adding a fixture

Each `fixtures/*.fixture.ts` file exports a named `fixture: Fixture` (`test/types.ts`):

```ts
import { InsertLocation } from "../src/word/insertLocation";
import type { Fixture } from "../test/types";

export const fixture: Fixture = {
  description: "Body.insertText(End) appends a new paragraph as the last child",
  // platform defaults to "PC" — set explicitly for a Mac/OfficeOnline-specific
  // fixture, e.g. to demonstrate a platform divergence (see
  // body-insertOoxml-end.fixture.ts / body-insertOoxml-end-officeonline.fixture.ts
  // for a paired example).
  seedOoxml: `<?xml version="1.0" ...`, // a full Flat-OPC document string
  apply: (context) => {
    context.document.body.insertText("New last.", InsertLocation.end);
  },
  resultOoxml: `<?xml version="1.0" ...`, // the expected document after apply()+sync()
};
```

`apply` is checked against the shim's real `RequestContext` type, so referencing an unimplemented
method is a compile error, not a silently-wrong fixture. For an operation expected to make
`sync()` reject (e.g. `insertOoxml` on `OfficeOnline`), set `expectRejection: true` (or a `RegExp`
to also match the rejection message) instead of `resultOoxml` — `Fixture` is a discriminated union,
so providing neither or both is a compile error, not a runtime surprise.

`test/fixtures.test.ts` discovers every `fixtures/*.fixture.ts` file automatically (via
`import.meta.glob`) — adding a fixture doesn't require touching the runner. It replays
`seedOoxml` → `apply` → `context.sync()` and compares the resulting document against
`resultOoxml` **structurally** (`test/structuralCompare.ts`: paragraph text, table cell text,
style names, and paraId well-formedness — not byte-for-byte, and not exact paraId _values_,
since freshly-created paragraphs get fresh random ids on every run). Real Word's own `getOoxml()`
output always carries one extra trailing, text-invisible paragraph mark (design spec's "ParaId
stability model") — the comparator drops it automatically, so don't account for it in
`resultOoxml`.

Read-only queries (`search`, `getComments`, `getStyles`) don't produce a `resultOoxml` to compare
against — those are covered by the package's regular `src/**/*.test.ts` unit tests instead, not
this harness.

The current fixture set is deliberately a small, representative sample — one per receiver
type/operation category (`Body`, `Range`, `Paragraph`, `Table`, incl. one PC/OfficeOnline
platform-divergence pair) — not full coverage of every `InsertLocation` × receiver-type
combination `src/operations/` implements. That full-matrix coverage is the real-Word-captured
golden corpus's job (issue #21), which drops into this exact same harness unchanged when it lands.
Add more synthetic fixtures here if you find a harness gap in the meantime, but exhaustive
combinatorial coverage isn't this harness's goal.

## Code style

- No comments explaining _what_ code does — only _why_, for non-obvious constraints.
- One operation per file under `src/operations/` — don't collapse receiver-type-specific
  `InsertLocation` handling into a single generic dispatcher.
