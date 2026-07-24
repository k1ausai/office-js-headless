import { describe, expect, it } from "vitest";
import { FlatOpcDocument } from "../document/FlatOpcDocument";
import { MINIMAL_SEED_OOXML } from "../document/__fixtures__/minimalSeed";
import { InsertLocation } from "./insertLocation";
import { wordRun } from "./run";

describe("wordRun / RequestContext deferred batching", () => {
  it("does not apply a queued mutation until context.sync() is called", async () => {
    const doc = new FlatOpcDocument(MINIMAL_SEED_OOXML);

    await wordRun(doc, "PC", async (context) => {
      context.document.body.insertText("Queued text.", InsertLocation.end);
      expect(doc.getOoxml()).not.toContain("Queued text.");
    });

    expect(doc.getOoxml()).not.toContain("Queued text.");
  });

  it("applies queued mutations in order once context.sync() is called", async () => {
    const doc = new FlatOpcDocument(MINIMAL_SEED_OOXML);

    await wordRun(doc, "PC", async (context) => {
      context.document.body.insertText("First.", InsertLocation.end);
      context.document.body.insertText("Second.", InsertLocation.end);
      await context.sync();
    });

    const ooxml = doc.getOoxml();
    expect(ooxml).toContain("First.");
    expect(ooxml).toContain("Second.");
    expect(ooxml.indexOf("First.")).toBeLessThan(ooxml.indexOf("Second."));
  });

  it("surfaces a queued operation's failure as sync()'s rejection, not at the call site", async () => {
    const doc = new FlatOpcDocument(MINIMAL_SEED_OOXML);

    await wordRun(doc, "PC", async (context) => {
      // Calling insertText with an unimplemented location must NOT throw
      // here at the call site — real Office.js defers all validation to the
      // sync() round-trip too.
      expect(() => context.document.body.insertText("Bad.", InsertLocation.before)).not.toThrow();

      await expect(context.sync()).rejects.toThrow(/InsertLocation/);
    });
  });

  it("body.getOoxml() (inside Word.run) reads current document state, same as the top-level escape hatch", async () => {
    const doc = new FlatOpcDocument(MINIMAL_SEED_OOXML);

    await wordRun(doc, "PC", async (context) => {
      context.document.body.insertText("Inline read.", InsertLocation.end);
      await context.sync();
      expect(context.document.body.getOoxml()).toContain("Inline read.");
      expect(context.document.body.getOoxml()).toEqual(doc.getOoxml());
    });
  });

  it("clears the queue after sync() so a second sync() call is a no-op", async () => {
    const doc = new FlatOpcDocument(MINIMAL_SEED_OOXML);

    await wordRun(doc, "PC", async (context) => {
      context.document.body.insertText("Once.", InsertLocation.end);
      await context.sync();
      await context.sync();
    });

    const ooxml = doc.getOoxml();
    const firstIndex = ooxml.indexOf("Once.");
    const lastIndex = ooxml.lastIndexOf("Once.");
    expect(firstIndex).toEqual(lastIndex);
  });
});
