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

## Code style

- No comments explaining _what_ code does — only _why_, for non-obvious constraints.
- One operation per file under `src/operations/` — don't collapse receiver-type-specific
  `InsertLocation` handling into a single generic dispatcher.
