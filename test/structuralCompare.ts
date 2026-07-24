import { expect } from "vitest";
import { FlatOpcDocument } from "../src/document/FlatOpcDocument";

const W14_NS = "http://schemas.microsoft.com/office/word/2010/wordml";
const PARA_ID_PATTERN = /^[0-9A-F]{8}$/;

export interface StructuralSummary {
  paragraphTexts: string[];
  tableCellTexts: string[];
  styleNames: string[];
  paraIdsValid: boolean;
}

// Design spec's "Fidelity validation": "compared structurally (paragraph
// text/paraId/style-name lists), not byte-for-byte." paraId is checked for
// well-formedness (8-hex-uppercase, present on every real paragraph), not
// exact-value equality — this shim generates fresh random ids for every new
// paragraph, so a synthetic (hand-written, not real-Word-captured) fixture
// has no "real" id to compare a value against. Once #21 lands real captures,
// exact-value comparison for paragraphs whose id is expected to stay STABLE
// (not newly created) would become meaningful and could be added then.
export function structuralSummary(ooxml: string): StructuralSummary {
  const doc = new FlatOpcDocument(ooxml);
  // `ooxml` here is itself a getOoxml()-shaped dump — this shim's own
  // output, a hand-written fixture, or (eventually, #21) a real Word
  // capture — which always carries a trailing paragraph mark as its actual
  // last <w:p> (design spec's "ParaId stability model"). Parsing it into a
  // FRESH FlatOpcDocument adds a second, distinct trailing mark on top,
  // but has no way to recognize the one already baked into the input
  // string as anything other than an ordinary empty paragraph. Drop it
  // here so two getOoxml() dumps compare correctly instead of spuriously
  // disagreeing on this text-invisible artifact — no real add-in ever
  // observes or cares about it either.
  const paragraphs = doc.getParagraphElements();
  const last = paragraphs[paragraphs.length - 1];
  const realParagraphs =
    last && doc.getParagraphText(last) === "" ? paragraphs.slice(0, -1) : paragraphs;
  // getParagraphElements() only reaches direct-child body paragraphs, same
  // as body.paragraphs — table-cell content needs walking tables/rows/cells
  // separately (design spec's "Golden corpus" names "table ops" as needing
  // fixture coverage too).
  const tableCellTexts = doc
    .getTableElements()
    .flatMap((table) => doc.getTableRows(table))
    .flatMap((row) => doc.getRowCells(row))
    .map((cell) => doc.getCellText(cell));
  return {
    paragraphTexts: realParagraphs.map((p) => doc.getParagraphText(p)),
    tableCellTexts,
    styleNames: doc.getStyleElements().map((s) => doc.getStyleNameLocal(s)),
    paraIdsValid: realParagraphs.every((p) => {
      const paraId = p.getAttributeNS(W14_NS, "paraId");
      return paraId !== null && PARA_ID_PATTERN.test(paraId);
    }),
  };
}

export function expectStructuralMatch(actualOoxml: string, expectedOoxml: string): void {
  const actual = structuralSummary(actualOoxml);
  const expected = structuralSummary(expectedOoxml);
  expect(actual.paragraphTexts).toEqual(expected.paragraphTexts);
  expect(actual.tableCellTexts).toEqual(expected.tableCellTexts);
  expect(actual.styleNames).toEqual(expected.styleNames);
  expect(actual.paraIdsValid).toBe(true);
  expect(expected.paraIdsValid).toBe(true);
}
