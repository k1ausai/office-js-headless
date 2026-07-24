import { describe, expect, it } from "vitest";
import { FlatOpcDocument } from "../document/FlatOpcDocument";
import { MINIMAL_SEED_OOXML } from "../document/__fixtures__/minimalSeed";
import { STYLE_SEED_OOXML } from "../document/__fixtures__/styleSeed";
import { BuiltInStyleName } from "./builtInStyleName";
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
      // Not a raw string comparison — the trailing paragraph mark's id
      // churns on every single getOoxml() call (see ParaId stability model),
      // so two separate calls are never byte-identical even with zero edits
      // between them. Assert both reflect the same real content instead.
      const viaBody = context.document.body.getOoxml();
      const viaEscapeHatch = doc.getOoxml();
      expect(viaBody).toContain("Inline read.");
      expect(viaEscapeHatch).toContain("Inline read.");
      expect(viaBody).toContain("Seed paragraph.");
      expect(viaEscapeHatch).toContain("Seed paragraph.");
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

describe("RequestContext.document.getStyles", () => {
  it("returns every distinct style, with BuiltInStyleName resolving against real ids — real Word.Document.getStyles(), not Body.getStyles()", async () => {
    const doc = new FlatOpcDocument(STYLE_SEED_OOXML);
    await wordRun(doc, "PC", async (context) => {
      const styles = context.document.getStyles();
      expect(styles).toHaveLength(4);

      styles.forEach((s) => s.load(["id", "nameLocal", "type", "builtIn"]));
      await context.sync();

      const heading1 = styles.find((s) => s.id === BuiltInStyleName.heading1);
      expect(heading1?.nameLocal).toBe("heading 1");
      expect(heading1?.builtIn).toBe(true);

      const custom = styles.find((s) => s.id === "MyCustomStyle");
      expect(custom?.builtIn).toBe(false);
      expect(custom?.type).toBe("Paragraph");
    });
  });

  it("returns an empty array for a document with no styles.xml part", async () => {
    const doc = new FlatOpcDocument(MINIMAL_SEED_OOXML);
    await wordRun(doc, "PC", async (context) => {
      expect(context.document.getStyles()).toEqual([]);
    });
  });
});
