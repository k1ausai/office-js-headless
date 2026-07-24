import { describe, expect, it } from "vitest";
import { FlatOpcDocument } from "../document/FlatOpcDocument";
import { MINIMAL_SEED_OOXML } from "../document/__fixtures__/minimalSeed";
import { InsertLocation } from "./insertLocation";
import { wordRun } from "./run";

describe("Body load/sync gating", () => {
  it("reading .text without calling .load() first throws PropertyNotLoaded", async () => {
    const doc = new FlatOpcDocument(MINIMAL_SEED_OOXML);

    await wordRun(doc, async (context) => {
      expect(() => context.document.body.text).toThrow(/PropertyNotLoaded|not available/);
    });
  });

  it("reading .text after .load() but before context.sync() still throws — identical to never-loaded", async () => {
    const doc = new FlatOpcDocument(MINIMAL_SEED_OOXML);

    await wordRun(doc, async (context) => {
      context.document.body.load("text");
      expect(() => context.document.body.text).toThrow(/PropertyNotLoaded|not available/);
    });
  });

  it("reading .text after .load() + context.sync() returns the current body text", async () => {
    const doc = new FlatOpcDocument(MINIMAL_SEED_OOXML);

    await wordRun(doc, async (context) => {
      context.document.body.load("text");
      await context.sync();
      expect(context.document.body.text).toBe("Seed paragraph.");
    });
  });

  it("reads the post-mutation value when insertText and .load('text') are queued in the same batch", async () => {
    const doc = new FlatOpcDocument(MINIMAL_SEED_OOXML);

    await wordRun(doc, async (context) => {
      context.document.body.insertText("Appended.", InsertLocation.end);
      context.document.body.load("text");
      await context.sync();
      expect(context.document.body.text).toBe("Seed paragraph.\nAppended.");
    });
  });

  it("the error name/code exactly match the confirmed real Office.js shape", async () => {
    const doc = new FlatOpcDocument(MINIMAL_SEED_OOXML);

    await wordRun(doc, async (context) => {
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
