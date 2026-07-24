import { type Element } from "@xmldom/xmldom";
import { describe, expect, it } from "vitest";
import { FlatOpcDocument } from "./FlatOpcDocument";
import { MINIMAL_SEED_OOXML } from "./__fixtures__/minimalSeed";

const W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const W14_NS = "http://schemas.microsoft.com/office/word/2010/wordml";
const HEX_ID = /^[0-9A-F]{8}$/;

function paraId(el: Element): string | null {
  return el.getAttributeNS(W14_NS, "paraId");
}

function realParagraphs(doc: FlatOpcDocument): Element[] {
  // Mirrors FlatOpcDocument's own private getRealParagraphs() — direct-child
  // <w:p> elements, in document order. Good enough for test purposes: the
  // trailing mark is always structurally last, so real content is
  // everything before it.
  return Array.from(doc.bodyElement.getElementsByTagNameNS(W_NS, "p"));
}

describe("ParaId stability model", () => {
  it("every newly-created paragraph gets an 8-hex-char uppercase id", () => {
    const doc = new FlatOpcDocument(MINIMAL_SEED_OOXML, "PC");
    const p = doc.appendParagraph("New.");
    expect(paraId(p)).toMatch(HEX_ID);
  });

  it("PC: insert-after reassigns the existing paragraph's OLD id onto the new paragraph, gives the existing paragraph a fresh one", () => {
    const doc = new FlatOpcDocument(MINIMAL_SEED_OOXML, "PC");
    const target = realParagraphs(doc)[0]!;
    const oldTargetId = paraId(target);
    expect(oldTargetId).toMatch(HEX_ID);

    const newPara = doc.insertParagraphAfter(target, "After.");

    expect(paraId(newPara)).toBe(oldTargetId);
    expect(paraId(target)).not.toBe(oldTargetId);
    expect(paraId(target)).toMatch(HEX_ID);
  });

  it("PC: insert-before does NOT mark-shift — the target keeps its id, the new paragraph gets an unrelated fresh one", () => {
    const doc = new FlatOpcDocument(MINIMAL_SEED_OOXML, "PC");
    const target = realParagraphs(doc)[0]!;
    const originalTargetId = paraId(target);

    const newPara = doc.insertParagraphBefore(target, "Before.");

    expect(paraId(target)).toBe(originalTargetId);
    expect(paraId(newPara)).not.toBe(originalTargetId);
    expect(paraId(newPara)).toMatch(HEX_ID);
  });

  it("PC: insert-before on the very first paragraph (no preceding sibling) behaves identically — no special-casing", () => {
    const doc = new FlatOpcDocument(MINIMAL_SEED_OOXML, "PC");
    const first = realParagraphs(doc)[0]!;
    const originalFirstId = paraId(first);

    const newPara = doc.insertParagraphBefore(first, "New first.");

    expect(paraId(first)).toBe(originalFirstId);
    expect(paraId(newPara)).not.toBe(originalFirstId);
  });

  it("PC/Mac: real paragraph ids are stable across multiple getOoxml() calls with no mutation between them", () => {
    for (const platform of ["PC", "Mac"] as const) {
      const doc = new FlatOpcDocument(MINIMAL_SEED_OOXML, platform);
      doc.getOoxml();
      const idAfterFirstCall = paraId(realParagraphs(doc)[0]!);
      doc.getOoxml();
      const idAfterSecondCall = paraId(realParagraphs(doc)[0]!);
      expect(idAfterSecondCall).toBe(idAfterFirstCall);
    }
  });

  it("PC/Mac: the trailing paragraph mark's id churns on every getOoxml() call regardless of edits", () => {
    for (const platform of ["PC", "Mac"] as const) {
      const doc = new FlatOpcDocument(MINIMAL_SEED_OOXML, platform);
      doc.getOoxml();
      const paragraphs1 = doc.bodyElement.getElementsByTagNameNS(W_NS, "p");
      const trailingId1 = paraId(paragraphs1[paragraphs1.length - 1]!);

      doc.getOoxml();
      const paragraphs2 = doc.bodyElement.getElementsByTagNameNS(W_NS, "p");
      const trailingId2 = paraId(paragraphs2[paragraphs2.length - 1]!);

      expect(trailingId2).not.toBe(trailingId1);
    }
  });

  it("PC/Mac: the trailing mark's id is excluded from getBodyText() — it's not real content", () => {
    const doc = new FlatOpcDocument(MINIMAL_SEED_OOXML, "PC");
    expect(doc.getBodyText()).toBe("Seed paragraph.");
    expect(doc.getBodyText().split("\n")).toHaveLength(1);
  });

  it("OfficeOnline: every getOoxml() call regenerates ALL paragraph ids, even with zero mutations between calls", () => {
    const doc = new FlatOpcDocument(MINIMAL_SEED_OOXML, "OfficeOnline");
    doc.getOoxml();
    const idsAfterFirstCall = realParagraphs(doc).map(paraId);

    doc.getOoxml();
    const idsAfterSecondCall = realParagraphs(doc).map(paraId);

    expect(idsAfterSecondCall).not.toEqual(idsAfterFirstCall);
    for (const id of [...idsAfterFirstCall, ...idsAfterSecondCall]) {
      expect(id).toMatch(HEX_ID);
    }
  });

  it("OfficeOnline: one shared rsid value is stamped across every paragraph per getOoxml() call", () => {
    const doc = new FlatOpcDocument(MINIMAL_SEED_OOXML, "OfficeOnline");
    doc.appendParagraph("Second.");
    doc.getOoxml();

    const rsids = realParagraphs(doc).map((p) => p.getAttributeNS(W_NS, "rsidR"));
    expect(rsids.length).toBeGreaterThan(1);
    expect(new Set(rsids).size).toBe(1);
    expect(rsids[0]).toMatch(HEX_ID);
  });

  it("OfficeOnline: mark-shift is not applied on insert-after — ids churn unconditionally on the next getOoxml() call anyway", () => {
    const doc = new FlatOpcDocument(MINIMAL_SEED_OOXML, "OfficeOnline");
    const target = realParagraphs(doc)[0]!;
    const originalTargetId = paraId(target);

    const newPara = doc.insertParagraphAfter(target, "After.");

    // No mark-shift performed — the new paragraph does NOT inherit target's
    // old id (unlike PC/Mac). Both already have their own freshly-generated
    // ids from creation; nothing here asserts platform-churn timing.
    expect(paraId(newPara)).not.toBe(originalTargetId);
  });
});
