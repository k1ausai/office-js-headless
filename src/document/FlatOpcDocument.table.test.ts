import { type Element } from "@xmldom/xmldom";
import { describe, expect, it } from "vitest";
import { MERGED_TABLE_SEED_OOXML, TABLE_SEED_OOXML } from "./__fixtures__/tableSeed";
import { FlatOpcDocument } from "./FlatOpcDocument";

const W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const W14_NS = "http://schemas.microsoft.com/office/word/2010/wordml";

function firstTable(doc: FlatOpcDocument): Element {
  const table = doc.bodyElement.getElementsByTagNameNS(W_NS, "tbl")[0];
  if (!table) throw new Error("test fixture has no <w:tbl>");
  return table;
}

describe("FlatOpcDocument table primitives", () => {
  it("getTableColumnCount reads w:tblGrid/w:gridCol, not the per-row w:tc count", () => {
    const doc = new FlatOpcDocument(TABLE_SEED_OOXML);
    expect(doc.getTableColumnCount(firstTable(doc))).toBe(3);
  });

  it("getTableColumnCount is unaffected by a horizontally-merged row with fewer w:tc than columns", () => {
    const doc = new FlatOpcDocument(MERGED_TABLE_SEED_OOXML);
    expect(doc.getTableColumnCount(firstTable(doc))).toBe(2);
  });

  it("getTableRows returns each direct-child w:tr in order", () => {
    const doc = new FlatOpcDocument(TABLE_SEED_OOXML);
    const rows = doc.getTableRows(firstTable(doc));
    expect(rows).toHaveLength(2);
    expect(doc.getRowCells(rows[0]!).map((c) => doc.getCellText(c))).toEqual([
      "R1C1",
      "R1C2",
      "R1C3",
    ]);
    expect(doc.getRowCells(rows[1]!).map((c) => doc.getCellText(c))).toEqual([
      "R2C1",
      "R2C2",
      "R2C3",
    ]);
  });

  it("getRowCells on a merged row returns fewer elements than the column count", () => {
    const doc = new FlatOpcDocument(MERGED_TABLE_SEED_OOXML);
    const rows = doc.getTableRows(firstTable(doc));
    expect(doc.getRowCells(rows[0]!)).toHaveLength(1);
    expect(doc.getRowCells(rows[1]!)).toHaveLength(2);
  });

  it("getCellText/setCellText reuse paragraph text construction", () => {
    const doc = new FlatOpcDocument(TABLE_SEED_OOXML);
    const cell = doc.getRowCells(doc.getTableRows(firstTable(doc))[0]!)[0]!;
    expect(doc.getCellText(cell)).toBe("R1C1");

    doc.setCellText(cell, "Updated.");
    expect(doc.getCellText(cell)).toBe("Updated.");
  });

  it("getCellGridSpan defaults to 1 when w:gridSpan is absent, and reads it when present", () => {
    const plain = new FlatOpcDocument(TABLE_SEED_OOXML);
    const plainCell = plain.getRowCells(plain.getTableRows(firstTable(plain))[0]!)[0]!;
    expect(plain.getCellGridSpan(plainCell)).toBe(1);

    const merged = new FlatOpcDocument(MERGED_TABLE_SEED_OOXML);
    const mergedCell = merged.getRowCells(merged.getTableRows(firstTable(merged))[0]!)[0]!;
    expect(merged.getCellGridSpan(mergedCell)).toBe(2);
  });

  it("getCellVMerge returns undefined when absent, 'Restart' when explicit, 'Continue' for a bare element (schema default)", () => {
    const plain = new FlatOpcDocument(TABLE_SEED_OOXML);
    const plainCell = plain.getRowCells(plain.getTableRows(firstTable(plain))[0]!)[0]!;
    expect(plain.getCellVMerge(plainCell)).toBeUndefined();

    const merged = new FlatOpcDocument(MERGED_TABLE_SEED_OOXML);
    const rows = merged.getTableRows(firstTable(merged));
    const topCell = merged.getRowCells(rows[0]!)[0]!;
    const bottomCell = merged.getRowCells(rows[1]!)[0]!;
    expect(merged.getCellVMerge(topCell)).toBe("Restart");
    expect(merged.getCellVMerge(bottomCell)).toBe("Continue");
  });

  it("addTableRows(End) appends new rows with columnCount cells each, filled from values", () => {
    const doc = new FlatOpcDocument(TABLE_SEED_OOXML);
    const table = firstTable(doc);

    const newRows = doc.addTableRows(table, "End", 1, [["A", "B", "C"]]);

    expect(newRows).toHaveLength(1);
    const rows = doc.getTableRows(table);
    expect(rows).toHaveLength(3);
    expect(doc.getRowCells(rows[2]!).map((c) => doc.getCellText(c))).toEqual(["A", "B", "C"]);
  });

  it("addTableRows(Start) prepends new rows, preserving their given order", () => {
    const doc = new FlatOpcDocument(TABLE_SEED_OOXML);
    const table = firstTable(doc);

    doc.addTableRows(table, "Start", 2, [
      ["X1", "X2", "X3"],
      ["Y1", "Y2", "Y3"],
    ]);

    const rows = doc.getTableRows(table);
    expect(rows).toHaveLength(4);
    expect(doc.getRowCells(rows[0]!).map((c) => doc.getCellText(c))).toEqual(["X1", "X2", "X3"]);
    expect(doc.getRowCells(rows[1]!).map((c) => doc.getCellText(c))).toEqual(["Y1", "Y2", "Y3"]);
    expect(doc.getRowCells(rows[2]!).map((c) => doc.getCellText(c))).toEqual([
      "R1C1",
      "R1C2",
      "R1C3",
    ]);
  });

  it("addTableRows without values creates cells with empty text", () => {
    const doc = new FlatOpcDocument(TABLE_SEED_OOXML);
    const table = firstTable(doc);

    doc.addTableRows(table, "End", 1);

    const rows = doc.getTableRows(table);
    expect(doc.getRowCells(rows[2]!).map((c) => doc.getCellText(c))).toEqual(["", "", ""]);
  });

  it("cell property elements are emitted in schema-mandated order: w:tcW, w:gridSpan, w:vMerge", () => {
    const doc = new FlatOpcDocument(MERGED_TABLE_SEED_OOXML);
    const ooxml = doc.getOoxml();
    const tcWIndex = ooxml.indexOf("w:tcW");
    const gridSpanIndex = ooxml.indexOf("w:gridSpan");
    const vMergeIndex = ooxml.indexOf("w:vMerge");
    expect(tcWIndex).toBeLessThan(gridSpanIndex);
    expect(gridSpanIndex).toBeLessThan(vMergeIndex);
  });

  it("w:tblPr/w:trPr are never emitted by table construction (omitted entirely for v1)", () => {
    const doc = new FlatOpcDocument(TABLE_SEED_OOXML);
    doc.addTableRows(firstTable(doc), "End", 1, [["A", "B", "C"]]);
    const ooxml = doc.getOoxml();
    expect(ooxml).not.toContain("w:tblPr");
    expect(ooxml).not.toContain("w:trPr");
  });

  it("OfficeOnline's paraId churn reaches paragraphs nested inside table cells, not just body-level ones", () => {
    const doc = new FlatOpcDocument(TABLE_SEED_OOXML, "OfficeOnline");
    const cell = doc.getRowCells(doc.getTableRows(firstTable(doc))[0]!)[0]!;
    const cellParagraph = cell.getElementsByTagNameNS(W_NS, "p")[0]!;

    doc.getOoxml();
    const idAfterFirstCall = cellParagraph.getAttributeNS(W14_NS, "paraId");
    expect(idAfterFirstCall).toMatch(/^[0-9A-F]{8}$/);

    doc.getOoxml();
    const idAfterSecondCall = cellParagraph.getAttributeNS(W14_NS, "paraId");
    expect(idAfterSecondCall).not.toBe(idAfterFirstCall);
  });
});
