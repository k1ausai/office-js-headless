import { describe, expect, it } from "vitest";
import { FlatOpcDocument } from "../src/document/FlatOpcDocument";
import { wordRun } from "../src/word/run";
import { expectStructuralMatch } from "./structuralCompare";
import type { Fixture } from "./types";

// Design spec's "Fidelity validation" / "Contract tests": "this package's
// CI suite replays each fixture's seed+operation through the shim and
// asserts the result matches... compared structurally... not byte-for-byte."
// Discovers every fixtures/*.fixture.ts file automatically — adding a new
// fixture doesn't require touching this runner.
const fixtureModules = import.meta.glob<{ fixture: Fixture }>("../fixtures/*.fixture.ts", {
  eager: true,
});
const fixtureEntries = Object.entries(fixtureModules);

describe("fixture replay", () => {
  it("discovers at least one fixture — proves the glob pattern itself hasn't silently broken", () => {
    expect(fixtureEntries.length).toBeGreaterThan(0);
  });

  for (const [path, mod] of fixtureEntries) {
    const fixture = mod.fixture;
    const name = path.split("/").pop();

    it(`${name}: ${fixture.description}`, async () => {
      const platform = fixture.platform ?? "PC";
      const doc = new FlatOpcDocument(fixture.seedOoxml, platform);

      if (fixture.resultOoxml === undefined) {
        const { expectRejection } = fixture;
        await wordRun(doc, platform, async (context) => {
          await fixture.apply(context);
          const expectedMessage = expectRejection instanceof RegExp ? expectRejection : undefined;
          await expect(context.sync()).rejects.toThrow(expectedMessage);
        });
        return;
      }

      // Narrowed to the resultOoxml-required union member here, synchronously
      // right after the discriminant check — TS's narrowing of `fixture`
      // itself doesn't reliably persist across the `await` below.
      const { resultOoxml } = fixture;
      await wordRun(doc, platform, async (context) => {
        await fixture.apply(context);
        await context.sync();
      });
      // Passes the live doc, not doc.getOoxml() — structuralSummaryOfDocument()
      // (used internally for a FlatOpcDocument actual side) reads directly
      // off the doc's own tracked paragraphs, no string round-trip needed
      // and no trailing-mark ambiguity to resolve for this side.
      expectStructuralMatch(doc, resultOoxml);
    });
  }
});
