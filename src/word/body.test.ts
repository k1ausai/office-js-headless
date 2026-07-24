import { describe, expect, it } from "vitest";
import { FlatOpcDocument } from "../document/FlatOpcDocument";
import { MINIMAL_SEED_OOXML } from "../document/__fixtures__/minimalSeed";
import { TABLE_SEED_OOXML } from "../document/__fixtures__/tableSeed";
import { InsertLocation } from "./insertLocation";
import { wordRun } from "./run";

describe("Body InsertLocation dispatch", () => {
  it("Start inserts a new paragraph as the first child", async () => {
    const doc = new FlatOpcDocument(MINIMAL_SEED_OOXML);
    await wordRun(doc, "PC", async (context) => {
      context.document.body.insertText("New first.", InsertLocation.start);
      await context.sync();
    });
    const ooxml = doc.getOoxml();
    expect(ooxml.indexOf("New first.")).toBeLessThan(ooxml.indexOf("Seed paragraph."));
  });

  it("End inserts a new paragraph as the last child", async () => {
    const doc = new FlatOpcDocument(MINIMAL_SEED_OOXML);
    await wordRun(doc, "PC", async (context) => {
      context.document.body.insertText("New last.", InsertLocation.end);
      await context.sync();
    });
    const ooxml = doc.getOoxml();
    expect(ooxml.indexOf("Seed paragraph.")).toBeLessThan(ooxml.indexOf("New last."));
  });

  it("Replace clears the whole body and inserts one paragraph", async () => {
    const doc = new FlatOpcDocument(MINIMAL_SEED_OOXML);
    await wordRun(doc, "PC", async (context) => {
      context.document.body.insertText("Only this.", InsertLocation.replace);
      await context.sync();
    });
    const ooxml = doc.getOoxml();
    expect(ooxml).toContain("Only this.");
    expect(ooxml).not.toContain("Seed paragraph.");
  });

  it("Before/After are not applicable to Body and reject at sync()", async () => {
    const doc = new FlatOpcDocument(MINIMAL_SEED_OOXML);
    await wordRun(doc, "PC", async (context) => {
      context.document.body.insertText("x", InsertLocation.before);
      await expect(context.sync()).rejects.toThrow(/InsertLocation/);
    });
    await wordRun(doc, "PC", async (context) => {
      context.document.body.insertText("x", InsertLocation.after);
      await expect(context.sync()).rejects.toThrow(/InsertLocation/);
    });
  });

  it("insertParagraph behaves the same as insertText for Body (both always create a new paragraph)", async () => {
    const doc = new FlatOpcDocument(MINIMAL_SEED_OOXML);
    await wordRun(doc, "PC", async (context) => {
      context.document.body.insertParagraph("Via insertParagraph.", InsertLocation.end);
      await context.sync();
    });
    expect(doc.getOoxml()).toContain("Via insertParagraph.");
  });
});

describe("Body.search", () => {
  it("returns a Range for each paragraph whose text contains the search string", async () => {
    const doc = new FlatOpcDocument(MINIMAL_SEED_OOXML);
    await wordRun(doc, "PC", async (context) => {
      context.document.body.insertParagraph("Another paragraph.", InsertLocation.end);
      await context.sync();

      const results = context.document.body.search("paragraph");
      expect(results).toHaveLength(2);
    });
  });

  it("matchCase defaults to false, and respects matchCase: true", async () => {
    const doc = new FlatOpcDocument(MINIMAL_SEED_OOXML);
    await wordRun(doc, "PC", async (context) => {
      expect(context.document.body.search("SEED")).toHaveLength(1);
      expect(context.document.body.search("SEED", { matchCase: true })).toHaveLength(0);
      expect(context.document.body.search("Seed", { matchCase: true })).toHaveLength(1);
    });
  });

  it("returns an empty array when nothing matches", async () => {
    const doc = new FlatOpcDocument(MINIMAL_SEED_OOXML);
    await wordRun(doc, "PC", async (context) => {
      expect(context.document.body.search("nonexistent")).toEqual([]);
    });
  });

  it("treats wildcard-special characters in the search text literally, through the public entry point", async () => {
    const doc = new FlatOpcDocument(MINIMAL_SEED_OOXML);
    await wordRun(doc, "PC", async (context) => {
      context.document.body.insertParagraph("Contains * and ? literally.", InsertLocation.end);
      await context.sync();

      expect(context.document.body.search("* and ?")).toHaveLength(1);
      // A regex-as-pattern interpretation of ".*" would match every
      // paragraph in the document; plain substring matching must not.
      expect(context.document.body.search(".*")).toHaveLength(0);
    });
  });
});

describe("Body.paragraphs", () => {
  it("getFirst/getLast are reachable end-to-end through Word.run and reflect the current document", async () => {
    const doc = new FlatOpcDocument(MINIMAL_SEED_OOXML);
    await wordRun(doc, "PC", async (context) => {
      context.document.body.insertParagraph("Second.", InsertLocation.end);
      await context.sync();

      const first = context.document.body.paragraphs.getFirst();
      const last = context.document.body.paragraphs.getLast();
      first.load("text");
      last.load("text");
      await context.sync();

      expect(first.text).toBe("Seed paragraph.");
      expect(last.text).toBe("Second.");
    });
  });
});

describe("Body.tables", () => {
  it("getFirst is reachable end-to-end through Word.run and reflects the current document", async () => {
    const doc = new FlatOpcDocument(TABLE_SEED_OOXML);
    await wordRun(doc, "PC", async (context) => {
      const table = context.document.body.tables.getFirst();
      table.load(["rowCount", "columnCount"]);
      await context.sync();

      expect(table.rowCount).toBe(2);
      expect(table.columnCount).toBe(3);

      const cell = table.getCell(0, 0);
      cell.load("value");
      await context.sync();
      expect(cell.value).toBe("R1C1");
    });
  });

  it("getFirst throws ItemNotFound when the document has no tables", async () => {
    const doc = new FlatOpcDocument(MINIMAL_SEED_OOXML);
    await wordRun(doc, "PC", async (context) => {
      expect(() => context.document.body.tables.getFirst()).toThrow(/not found/);
    });
  });
});

describe("Body load/sync gating", () => {
  it("reading .text without calling .load() first throws PropertyNotLoaded", async () => {
    const doc = new FlatOpcDocument(MINIMAL_SEED_OOXML);

    await wordRun(doc, "PC", async (context) => {
      expect(() => context.document.body.text).toThrow(/PropertyNotLoaded|not available/);
    });
  });

  it("reading .text after .load() but before context.sync() still throws — identical to never-loaded", async () => {
    const doc = new FlatOpcDocument(MINIMAL_SEED_OOXML);

    await wordRun(doc, "PC", async (context) => {
      context.document.body.load("text");
      expect(() => context.document.body.text).toThrow(/PropertyNotLoaded|not available/);
    });
  });

  it("reading .text after .load() + context.sync() returns the current body text", async () => {
    const doc = new FlatOpcDocument(MINIMAL_SEED_OOXML);

    await wordRun(doc, "PC", async (context) => {
      context.document.body.load("text");
      await context.sync();
      expect(context.document.body.text).toBe("Seed paragraph.");
    });
  });

  it("reads the post-mutation value when insertText and .load('text') are queued in the same batch", async () => {
    const doc = new FlatOpcDocument(MINIMAL_SEED_OOXML);

    await wordRun(doc, "PC", async (context) => {
      context.document.body.insertText("Appended.", InsertLocation.end);
      context.document.body.load("text");
      await context.sync();
      expect(context.document.body.text).toBe("Seed paragraph.\nAppended.");
    });
  });

  it("the error name/code exactly match the confirmed real Office.js shape", async () => {
    const doc = new FlatOpcDocument(MINIMAL_SEED_OOXML);

    await wordRun(doc, "PC", async (context) => {
      try {
        void context.document.body.text;
        expect.unreachable("should have thrown");
      } catch (err) {
        expect((err as Error).name).toBe("RichApi.Error");
        expect((err as Error & { code: string }).code).toBe("PropertyNotLoaded");
      }
    });
  });
});
