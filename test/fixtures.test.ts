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
const fixtureModules = import.meta.glob<{ default: Fixture }>("../fixtures/*.fixture.ts", {
  eager: true,
});
const fixtureEntries = Object.entries(fixtureModules);

describe("fixture replay", () => {
  it("discovers at least one fixture — proves the glob pattern itself hasn't silently broken", () => {
    expect(fixtureEntries.length).toBeGreaterThan(0);
  });

  for (const [path, mod] of fixtureEntries) {
    const fixture = mod.default;
    const name = path.split("/").pop();

    it(`${name}: ${fixture.description}`, async () => {
      const platform = fixture.platform ?? "PC";
      const doc = new FlatOpcDocument(fixture.seedOoxml, platform);

      const expectRejection = fixture.expectRejection;
      if (expectRejection) {
        await wordRun(doc, platform, async (context) => {
          await fixture.apply(context);
          const expectedMessage = expectRejection instanceof RegExp ? expectRejection : undefined;
          await expect(context.sync()).rejects.toThrow(expectedMessage);
        });
        return;
      }

      if (!fixture.resultOoxml) {
        throw new Error(`${name}: resultOoxml is required unless expectRejection is set`);
      }
      await wordRun(doc, platform, async (context) => {
        await fixture.apply(context);
        await context.sync();
      });
      expectStructuralMatch(doc.getOoxml(), fixture.resultOoxml);
    });
  }
});
