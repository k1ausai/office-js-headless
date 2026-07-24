import { type Element } from "@xmldom/xmldom";
import { describe, expect, it } from "vitest";
import { FlatOpcDocument } from "../document/FlatOpcDocument";
import { MERGED_TABLE_SEED_OOXML, TABLE_SEED_OOXML } from "../document/__fixtures__/tableSeed";
import { Table } from "./table";

const W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

function firstTable(doc: FlatOpcDocument): Element {
  const table = doc.bodyElement.getElementsByTagNameNS(W_NS, "tbl")[0];
  if (!table) throw new Error("test fixture has no <w:tbl>");
  return table;
}

function immediateEnqueue(op: () => void): void {
  op();
}

function noopRegisterSyncable(): void {}

describe("Table", () => {
  it("rowCount/columnCount respect load/sync gating, like Body.text", () => {
    const doc = new FlatOpcDocument(TABLE_SEED_OOXML);
    const table = new Table(doc, firstTable(doc), immediateEnqueue, noopRegisterSyncable);

    expect(() => table.rowCount).toThrow(/PropertyNotLoaded|not available/);

    table.load(["rowCount", "columnCount"]);
    expect(() => table.rowCount).toThrow(/PropertyNotLoaded|not available/);

    table.sync();
    expect(table.rowCount).toBe(2);
    expect(table.columnCount).toBe(3);
  });

  it("columnCount reflects w:tblGrid even for a table with a horizontally-merged row", () => {
    const doc = new FlatOpcDocument(MERGED_TABLE_SEED_OOXML);
    const table = new Table(doc, firstTable(doc), immediateEnqueue, noopRegisterSyncable);
    table.load(["rowCount", "columnCount"]);
    table.sync();
    expect(table.rowCount).toBe(2);
    expect(table.columnCount).toBe(2);
  });

  it("getCell reads/writes the correct cell's text", () => {
    const doc = new FlatOpcDocument(TABLE_SEED_OOXML);
    const table = new Table(doc, firstTable(doc), immediateEnqueue, noopRegisterSyncable);

    const cell = table.getCell(1, 2);
    cell.load("value");
    cell.sync();
    expect(cell.value).toBe("R2C3");

    cell.value = "Updated.";
    cell.load("value");
    cell.sync();
    expect(cell.value).toBe("Updated.");
  });

  it("getCell throws ItemNotFound for an out-of-range row or cell index", () => {
    const doc = new FlatOpcDocument(TABLE_SEED_OOXML);
    const table = new Table(doc, firstTable(doc), immediateEnqueue, noopRegisterSyncable);

    expect(() => table.getCell(99, 0)).toThrow(/not found/);
    expect(() => table.getCell(0, 99)).toThrow(/not found/);
  });

  it("getRow returns a TableRow that can resolve its own cells", () => {
    const doc = new FlatOpcDocument(TABLE_SEED_OOXML);
    const table = new Table(doc, firstTable(doc), immediateEnqueue, noopRegisterSyncable);

    const row = table.getRow(0);
    const cell = row.getCell(1);
    cell.load("value");
    cell.sync();
    expect(cell.value).toBe("R1C2");
  });

  it("a cell's gridSpan/vMerge are readable without load/sync — plain structural reads", () => {
    const doc = new FlatOpcDocument(MERGED_TABLE_SEED_OOXML);
    const table = new Table(doc, firstTable(doc), immediateEnqueue, noopRegisterSyncable);

    const topCell = table.getCell(0, 0);
    expect(topCell.gridSpan).toBe(2);
    expect(topCell.vMerge).toBe("Restart");

    const bottomLeftCell = table.getCell(1, 0);
    expect(bottomLeftCell.gridSpan).toBe(1);
    expect(bottomLeftCell.vMerge).toBe("Continue");

    const bottomRightCell = table.getCell(1, 1);
    expect(bottomRightCell.gridSpan).toBe(1);
    expect(bottomRightCell.vMerge).toBeUndefined();
  });

  it("addRows(End) grows the table, deferred until the queued op runs", () => {
    const doc = new FlatOpcDocument(TABLE_SEED_OOXML);
    const target = firstTable(doc);
    const queue: Array<() => void> = [];
    const table = new Table(doc, target, (op) => queue.push(op), noopRegisterSyncable);

    table.addRows("End", 1, [["A", "B", "C"]]);
    expect(doc.getTableRows(target)).toHaveLength(2);

    queue.forEach((op) => op());
    expect(doc.getTableRows(target)).toHaveLength(3);
    expect(doc.getCellText(doc.getRowCells(doc.getTableRows(target)[2]!)[0]!)).toBe("A");
  });
});
