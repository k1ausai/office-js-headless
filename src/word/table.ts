import { type Element } from "@xmldom/xmldom";
import { FlatOpcDocument } from "../document/FlatOpcDocument";
import { SupportedPlatform } from "../office/context";
import { richApiError } from "./errors";
import { TrackedProperties } from "./proxy";
import type { QueuedOperation, Syncable } from "./run";

type TableProperty = "rowCount" | "columnCount";
type TableCellProperty = "value";

function itemNotFoundError(): Error {
  return richApiError("ItemNotFound", "The requested item was not found in the collection.");
}

// Whole-table granularity: wraps exactly one <w:tbl>. addRows() is deferred
// like every other mutating call (Body/Range/Paragraph's insertText) and
// returns void, not the newly created rows — real Office.js's addRows()
// returns a TableRowCollection synchronously, but this shim doesn't model
// insert-method return values anywhere else either (established in #9-#12),
// so staying void here keeps Table consistent with that precedent rather
// than introducing a one-off exception.
export class Table implements Syncable {
  private readonly tracked = new TrackedProperties();

  constructor(
    private readonly doc: FlatOpcDocument,
    private readonly target: Element,
    private readonly platform: SupportedPlatform,
    private readonly enqueue: (op: QueuedOperation) => void
  ) {}

  load(propertyNames: TableProperty | TableProperty[]): void {
    this.tracked.load(propertyNames);
  }

  get rowCount(): number {
    return this.tracked.read<number>("rowCount");
  }

  get columnCount(): number {
    return this.tracked.read<number>("columnCount");
  }

  /** Called by RequestContext.sync() after the mutation queue has run. */
  sync(): void {
    this.tracked.sync((name) => this.computeProperty(name));
  }

  private computeProperty(name: string): unknown {
    switch (name) {
      case "rowCount":
        return this.doc.getTableRows(this.target).length;
      case "columnCount":
        return this.doc.getTableColumnCount(this.target);
      default:
        throw new Error(`Table: unknown property "${name}"`);
    }
  }

  getRow(rowIndex: number): TableRow {
    const rows = this.doc.getTableRows(this.target);
    const row = rows[rowIndex];
    if (!row) throw itemNotFoundError();
    return new TableRow(this.doc, row, this.platform, this.enqueue);
  }

  getCell(rowIndex: number, cellIndex: number): TableCell {
    return this.getRow(rowIndex).getCell(cellIndex);
  }

  addRows(insertLocation: "Start" | "End", rowCount: number, values?: string[][]): void {
    this.enqueue(() => {
      this.doc.addTableRows(this.target, insertLocation, rowCount, values);
    });
  }
}

export class TableRow {
  constructor(
    private readonly doc: FlatOpcDocument,
    private readonly target: Element,
    private readonly platform: SupportedPlatform,
    private readonly enqueue: (op: QueuedOperation) => void
  ) {}

  getCell(cellIndex: number): TableCell {
    const cells = this.doc.getRowCells(this.target);
    const cell = cells[cellIndex];
    if (!cell) throw itemNotFoundError();
    return new TableCell(this.doc, cell, this.platform, this.enqueue);
  }
}

export class TableCell implements Syncable {
  private readonly tracked = new TrackedProperties();

  constructor(
    private readonly doc: FlatOpcDocument,
    private readonly target: Element,
    private readonly platform: SupportedPlatform,
    private readonly enqueue: (op: QueuedOperation) => void
  ) {}

  load(propertyNames: TableCellProperty | TableCellProperty[]): void {
    this.tracked.load(propertyNames);
  }

  get value(): string {
    return this.tracked.read<string>("value");
  }

  set value(text: string) {
    this.enqueue(() => {
      this.doc.setCellText(this.target, text);
    });
  }

  /** Called by RequestContext.sync() after the mutation queue has run. */
  sync(): void {
    this.tracked.sync((name) => this.computeProperty(name));
  }

  private computeProperty(name: string): unknown {
    switch (name) {
      case "value":
        return this.doc.getCellText(this.target);
      default:
        throw new Error(`TableCell: unknown property "${name}"`);
    }
  }

  // Not gated via TrackedProperties, unlike .value — these are plain
  // structural reads over OOXML attributes issue #15 explicitly scopes as
  // read-only (no write API exists, so there's no mutation for load/sync
  // timing to protect a stale read against). Real Word.TableCell doesn't
  // expose gridSpan/vMerge at all; this is a shim-specific extension for
  // round-trip fidelity, not a mirrored real property.
  get gridSpan(): number {
    return this.doc.getCellGridSpan(this.target);
  }

  get vMerge(): "Continue" | "Restart" | undefined {
    return this.doc.getCellVMerge(this.target);
  }
}
