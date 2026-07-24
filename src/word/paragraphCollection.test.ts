import { describe, expect, it } from "vitest";
import { FlatOpcDocument } from "../document/FlatOpcDocument";
import { MINIMAL_SEED_OOXML } from "../document/__fixtures__/minimalSeed";
import { ParagraphCollection } from "./paragraphCollection";

function immediateEnqueue(op: () => void): void {
  op();
}

function noopRegisterSyncable(): void {}

describe("ParagraphCollection", () => {
  it("getFirst returns a Paragraph proxy wrapping the first real paragraph", () => {
    const doc = new FlatOpcDocument(MINIMAL_SEED_OOXML);
    doc.appendParagraph("Second.");
    const collection = new ParagraphCollection(doc, "PC", immediateEnqueue, noopRegisterSyncable);

    const first = collection.getFirst();
    first.load("text");
    first.sync();
    expect(first.text).toBe("Seed paragraph.");
  });

  it("getLast returns a Paragraph proxy wrapping the last real paragraph", () => {
    const doc = new FlatOpcDocument(MINIMAL_SEED_OOXML);
    doc.appendParagraph("Second.");
    const collection = new ParagraphCollection(doc, "PC", immediateEnqueue, noopRegisterSyncable);

    const last = collection.getLast();
    last.load("text");
    last.sync();
    expect(last.text).toBe("Second.");
  });

  it("getFirst and getLast are the same paragraph when there's only one", () => {
    const doc = new FlatOpcDocument(MINIMAL_SEED_OOXML);
    const collection = new ParagraphCollection(doc, "PC", immediateEnqueue, noopRegisterSyncable);

    const first = collection.getFirst();
    const last = collection.getLast();
    first.load("text");
    last.load("text");
    first.sync();
    last.sync();
    expect(first.text).toBe(last.text);
  });

  it("getFirst/getLast throw ItemNotFound when the collection is empty", () => {
    const doc = new FlatOpcDocument(MINIMAL_SEED_OOXML);
    for (const p of doc.getParagraphElements()) {
      doc.deleteNode(p);
    }
    const collection = new ParagraphCollection(doc, "PC", immediateEnqueue, noopRegisterSyncable);

    expect(() => collection.getFirst()).toThrow(/not found/);
    expect(() => collection.getLast()).toThrow(/not found/);

    try {
      collection.getFirst();
      expect.unreachable("should have thrown");
    } catch (err) {
      expect((err as Error).name).toBe("RichApi.Error");
      expect((err as Error & { code: string }).code).toBe("ItemNotFound");
    }
  });

  it("registers each Paragraph it creates as syncable, so a later context.sync() snapshots its loaded properties", () => {
    const doc = new FlatOpcDocument(MINIMAL_SEED_OOXML);
    const registered: Array<{ sync(): void }> = [];
    const collection = new ParagraphCollection(doc, "PC", immediateEnqueue, (obj) =>
      registered.push(obj)
    );

    const first = collection.getFirst();
    first.load("text");
    expect(registered).toHaveLength(1);

    registered[0]!.sync();
    expect(first.text).toBe("Seed paragraph.");
  });
});
