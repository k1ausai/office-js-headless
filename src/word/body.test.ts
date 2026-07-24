import JSZip from "jszip";
import { describe, expect, it } from "vitest";
import { COMMENT_SEED_OOXML } from "../document/__fixtures__/commentSeed";
import { FlatOpcDocument } from "../document/FlatOpcDocument";
import { MINIMAL_SEED_OOXML } from "../document/__fixtures__/minimalSeed";
import { TABLE_SEED_OOXML } from "../document/__fixtures__/tableSeed";
import { InsertLocation } from "./insertLocation";
import { wordRun } from "./run";

const W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

// Builds a minimal in-memory .docx-shaped zip (just the one part these
// tests need — real .docx files have many more) and returns it base64
// encoded, matching what Body.insertFileFromBase64 actually receives.
async function buildBase64Docx(stylesXml?: string): Promise<string> {
  const zip = new JSZip();
  if (stylesXml) zip.file("word/styles.xml", stylesXml);
  return zip.generateAsync({ type: "base64" });
}

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

describe("Body.insertHtml", () => {
  it("applies identically on PC, Mac, and OfficeOnline — the always-works insertOoxml fallback", async () => {
    for (const platform of ["PC", "Mac", "OfficeOnline"] as const) {
      const doc = new FlatOpcDocument(MINIMAL_SEED_OOXML, platform);
      await wordRun(doc, platform, async (context) => {
        context.document.body.insertHtml(
          "<p><b>Bold</b> and <i>italic</i>.</p>",
          InsertLocation.end
        );
        await context.sync();
      });
      const ooxml = doc.getOoxml();
      expect(ooxml).toContain("<w:b/>");
      expect(ooxml).toContain("<w:i/>");
      expect(doc.getBodyText()).toBe("Seed paragraph.\nBold and italic.");
    }
  });

  it("Before/After are not applicable to Body and reject at sync(), same as insertOoxml/insertText", async () => {
    const doc = new FlatOpcDocument(MINIMAL_SEED_OOXML);
    await wordRun(doc, "PC", async (context) => {
      context.document.body.insertHtml("<p>x</p>", InsertLocation.before);
      await expect(context.sync()).rejects.toThrow(/InsertLocation/);
    });
  });
});

describe("Body.insertFileFromBase64", () => {
  it("on PC/Mac: unzips the base64 .docx and merges its styles.xml into the current document", async () => {
    const base64Docx = await buildBase64Docx(
      `<w:styles xmlns:w="${W_NS}"><w:style w:type="paragraph" w:styleId="Imported"><w:name w:val="Imported"/></w:style></w:styles>`
    );

    for (const platform of ["PC", "Mac"] as const) {
      const doc = new FlatOpcDocument(MINIMAL_SEED_OOXML, platform);
      await wordRun(doc, platform, async (context) => {
        context.document.body.insertFileFromBase64(base64Docx, InsertLocation.end);
        await context.sync();

        const styles = context.document.getStyles();
        styles.forEach((s) => s.load("id"));
        await context.sync();

        expect(styles.map((s) => s.id)).toEqual(["Imported"]);
      });
    }
  });

  it("on OfficeOnline: rejects at sync() — no client-side implementation, matching real Word Online", async () => {
    const base64Docx = await buildBase64Docx();
    const doc = new FlatOpcDocument(MINIMAL_SEED_OOXML, "OfficeOnline");
    await wordRun(doc, "OfficeOnline", async (context) => {
      context.document.body.insertFileFromBase64(base64Docx, InsertLocation.end);
      await expect(context.sync()).rejects.toThrow(/insertFileFromBase64/);
    });
  });

  it("Before/After are not applicable to Body and reject at sync()", async () => {
    const base64Docx = await buildBase64Docx();
    const doc = new FlatOpcDocument(MINIMAL_SEED_OOXML);
    await wordRun(doc, "PC", async (context) => {
      context.document.body.insertFileFromBase64(base64Docx, InsertLocation.before);
      await expect(context.sync()).rejects.toThrow(/InsertLocation/);
    });
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

describe("Body.getComments", () => {
  it("returns top-level comments (with threaded replies nested via .replies), reachable end-to-end through Word.run", async () => {
    const doc = new FlatOpcDocument(COMMENT_SEED_OOXML);
    await wordRun(doc, "PC", async (context) => {
      const comments = context.document.body.getComments();
      expect(comments).toHaveLength(2);

      comments.forEach((c) => c.load(["authorName", "content", "resolved", "replies"]));
      await context.sync();

      const [resolvedComment, unresolvedComment] = comments;
      expect(resolvedComment!.authorName).toBe("Jane Doe");
      expect(resolvedComment!.resolved).toBe(true);
      expect(resolvedComment!.replies).toHaveLength(1);
      expect(resolvedComment!.replies[0]).toMatchObject({ authorName: "John Smith" });

      expect(unresolvedComment!.resolved).toBe(false);
      expect(unresolvedComment!.replies).toEqual([]);
    });
  });

  it("returns an empty array for a document with no comments", async () => {
    const doc = new FlatOpcDocument(MINIMAL_SEED_OOXML);
    await wordRun(doc, "PC", async (context) => {
      expect(context.document.body.getComments()).toEqual([]);
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
