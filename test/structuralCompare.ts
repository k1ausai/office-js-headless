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
//
// Takes a live FlatOpcDocument, not a string — getParagraphElements()
// already excludes the doc's own tracked trailing mark by direct object
// reference, with no ambiguity. Prefer this over structuralSummary() below
// whenever a live doc is on hand (the fixture runner's "actual" side always
// is), since it avoids the string round-trip that motivates that
// function's less-precise trailing-paragraph heuristic entirely.
export function structuralSummaryOfDocument(doc: FlatOpcDocument): StructuralSummary {
  const paragraphs = doc.getParagraphElements();
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
    paragraphTexts: paragraphs.map((p) => doc.getParagraphText(p)),
    tableCellTexts,
    styleNames: doc.getStyleElements().map((s) => doc.getStyleNameLocal(s)),
    paraIdsValid: paragraphs.every((p) => {
      const paraId = p.getAttributeNS(W14_NS, "paraId");
      return paraId !== null && PARA_ID_PATTERN.test(paraId);
    }),
  };
}

// For an OOXML STRING rather than a live doc — needed for resultOoxml
// (always a string: hand-written today, a real Word capture from #21
// later) and for direct string-vs-string comparisons in this module's own
// tests. Unlike structuralSummaryOfDocument(), the input string here is
// itself potentially a getOoxml()-shaped dump — this shim's own output or
// a real Word capture — which always carries a trailing, text-invisible
// paragraph mark as its actual last <w:p> (design spec's "ParaId stability
// model"). Parsing it into a fresh FlatOpcDocument adds a SECOND, distinct
// trailing mark on top (correctly excluded by getParagraphElements()), but
// has no way to recognize the one already baked into the input string as
// anything other than an ordinary empty paragraph. Best-effort fix: drop a
// trailing empty paragraph. This can't distinguish that artifact from a
// document a fixture author genuinely intended to end on an empty
// paragraph — an accepted, documented limitation of comparing OOXML as
// strings, not a concern when a live doc is available instead (see above).
export function structuralSummary(ooxml: string): StructuralSummary {
  const summary = structuralSummaryOfDocument(new FlatOpcDocument(ooxml));
  const { paragraphTexts } = summary;
  const isTrailingMarkArtifact = paragraphTexts[paragraphTexts.length - 1] === "";
  return isTrailingMarkArtifact
    ? { ...summary, paragraphTexts: paragraphTexts.slice(0, -1) }
    : summary;
}

export function expectStructuralMatch(
  actual: FlatOpcDocument | string,
  expectedOoxml: string
): void {
  const actualSummary =
    typeof actual === "string" ? structuralSummary(actual) : structuralSummaryOfDocument(actual);
  const expectedSummary = structuralSummary(expectedOoxml);
  expect(actualSummary.paragraphTexts).toEqual(expectedSummary.paragraphTexts);
  expect(actualSummary.tableCellTexts).toEqual(expectedSummary.tableCellTexts);
  expect(actualSummary.styleNames).toEqual(expectedSummary.styleNames);
  expect(actualSummary.paraIdsValid).toBe(true);
  expect(expectedSummary.paraIdsValid).toBe(true);
}
