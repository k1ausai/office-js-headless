# Agent guide for office-js-headless

Headless Office.js runtime for testing Word add-ins. Public, standalone package, designed for
general reuse — not tied to or special-cased for any one consumer.

## Design spec

Full design: [`doc/design-spec/`](./doc/design-spec). Read it before making architectural changes.
Key decisions already made there — don't relitigate without a strong reason:

- Backing model is a live Flat-OPC XML tree (`@xmldom/xmldom`), not a real `.docx` zip. The one
  exception is `insertFileFromBase64`, which needs `jszip` to peek at a real `.docx`'s
  `styles.xml`/`numbering.xml` (and is a deliberately approximate stub, not full-fidelity).
- `Word.run`/`context.sync()` use real deferred batching, not eager execution — errors surface at
  `sync()`, matching real Office.js.
- `.load()`/`.sync()` gating is strict: reading an unloaded or unsynced property throws, same as
  real Office.js. This is intentional — it's what catches real add-in bugs.
- Structural/semantic OOXML fidelity is the goal, not byte-exact serialization.
- The shim targets whichever platform (`PC`/`Mac`/`OfficeOnline`) `installHeadlessOffice`'s
  `platform` option is told to emulate — not desktop-only. Platform-specific behavior (capability
  signals, `insertOoxml`/`insertFileFromBase64` failing on `OfficeOnline`, paraId churn) is
  in-scope and must match the real host per platform; only the _why_ behind Word Online's
  internals is out of scope, not its observable behavior.

## Fidelity validation

Every supported operation should have a golden fixture in `fixtures/` — a real
`{seedOoxml, operation+args, resultOoxml}` triple captured from actual Word, per platform where
behavior diverges (`PC`/`Mac` can usually share one; `OfficeOnline` separately). The shim's
own tests replay these and assert structural equivalence (not byte-diff). If a real-world case is
found where the shim diverges from real Word, capture it as a new golden fixture before fixing the
shim — production-found gaps become permanent regression tests.

## Commands

```bash
pnpm install
pnpm test            # vitest run
pnpm test:watch
pnpm lint / lint:fix
pnpm format / format:check
pnpm type:check
pnpm build            # vite build -> dist/
pnpm pack:check       # publint + attw — validates the actual publishable package
```

## Conventions

- pnpm, TypeScript strict mode, ESM-first (CJS build output also provided).
- No comments explaining _what_ code does — only _why_, for non-obvious constraints.
- One operation per file under `src/operations/` (per receiver type + InsertLocation combination)
  — don't collapse into a single generic dispatcher, since Word's semantics genuinely differ per
  receiver type.
- Named exports only. No barrel exports except the library's own `src/index.ts` entry point.
- Never `git commit` or `git push` without explicit permission — ask first. Never stage with
  `git add -A`/`git add .`; stage only the specific files changed, so unrelated working-tree
  changes aren't swept in.
