# office-js-headless

Headless Office.js runtime for testing Word add-ins — no real Word (desktop or web) needed.

Implements the subset of the `Word`/`Office` global API surface that add-ins actually call
(`Word.run`, `context.document.body`, `getOoxml`/`insertOoxml`, ranges, paragraphs, tables,
comments, styles, `InsertLocation`), backed by a real in-memory OOXML (Flat-OPC) document model
instead of a live Word process. Add-in code that talks to Office.js runs unmodified against it.

## Status

Early scaffold — design spec in [`doc/design-spec/`](./doc/design-spec), no implementation yet.

## Why

Office.js has no official headless runtime — it only exists as the bridge Word injects into a
loaded add-in's webview. Testing document-mutation logic for real has meant driving an actual
Word process (desktop or web) via browser/UI automation: slow, GUI-session-bound, hard to run in
CI. `office-js-headless` reimplements the used API surface directly over OOXML XML so this can
run in any Node process, including a normal CI container.

## Usage (planned)

```ts
import { installHeadlessOffice } from "office-js-headless";

const office = installHeadlessOffice({ seedOoxml: fixtureXml });

await Word.run(async (context) => {
  context.document.body.insertText("hello", Word.InsertLocation.end);
  await context.sync();
});

office.getOoxml();
office.dispose();
```

## Scope

Word only, for now. See [`doc/design-spec`](./doc/design-spec) for the full design, coverage list,
and explicit non-goals (Word Online fallback behavior, Excel/PowerPoint/Outlook, byte-exact
serialization).

## Development

```bash
pnpm install
pnpm test
pnpm lint
pnpm type:check
pnpm build
```

## Publishing

Normal path is fully automated, driven by [Conventional Commits](https://www.conventionalcommits.org/)
(enforced by commitlint):

1. Merge PRs to `main` with conventional commit messages (`feat:`, `fix:`, …).
2. `release-please` opens/updates a standing "Release PR" that bumps `package.json` version and
   writes `CHANGELOG.md` from those commits.
3. Merging that PR creates a `vX.Y.Z` git tag, which triggers `cd.yml`: lint, type-check, test,
   build, `pack:check`, then `pnpm publish --provenance`. Auth comes from the `NPM_TOKEN` repo
   secret; nothing local needed.

Manual/local publish (rare — CI is the source of truth):

1. Copy `.env.example` to `.env` and fill in `NPM_TOKEN` (a "Granular Access Token", Automation
   type, publish access scoped to `office-js-headless`, from
   https://www.npmjs.com/settings/&lt;you&gt;/tokens). `.env` is gitignored — never commit it.
2. `export $(grep -v '^#' .env | xargs)` to load it into the shell. `.npmrc` reads
   `NPM_TOKEN` via `//registry.npmjs.org/:_authToken=${NPM_TOKEN}`, so once it's exported, both
   `pnpm install` and `pnpm publish` are authenticated.
3. Bump `version` in `package.json` (or `npm version patch|minor|major`), then:
   ```bash
   pnpm lint && pnpm type:check && pnpm test && pnpm build && pnpm pack:check
   pnpm publish --access public --provenance
   ```
4. Push the matching `vX.Y.Z` tag so the GitHub release history and CD stay in sync with what you
   just published by hand.

The same `NPM_TOKEN` value must also be set as a GitHub Actions repo secret (Settings → Secrets
and variables → Actions) for the tag-triggered CD path to work.

## License

MIT
