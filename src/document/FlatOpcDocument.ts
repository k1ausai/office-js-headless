import { DOMParser, XMLSerializer, type Document, type Element, type Node } from "@xmldom/xmldom";
import { type SupportedPlatform } from "../office/context";
import {
  applyInsertAfterMarkShift,
  assignFreshIds,
  regenerateAllIdsForOfficeOnline,
} from "../word/paraId";

const W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const W14_NS = "http://schemas.microsoft.com/office/word/2010/wordml";
const W15_NS = "http://schemas.microsoft.com/office/word/2012/wordml";
const PKG_NS = "http://schemas.microsoft.com/office/2006/xmlPackage";
const ELEMENT_NODE = 1;

// Subtree search — fine for locating the single <w:body> nested under
// <pkg:package>/<pkg:part>/<pkg:xmlData>/<w:document>, since a document.xml
// part has exactly one <w:body>.
function findFirstDescendantElementNS(root: Element, namespaceURI: string, localName: string) {
  const matches = root.getElementsByTagNameNS(namespaceURI, localName);
  return matches.length > 0 ? matches[0] : null;
}

// Direct-children-only search — required for locating <w:body>'s own
// trailing <w:sectPr> (the document's section properties). A subtree search
// here would also match a mid-document section break's <w:sectPr>, which
// lives nested inside a paragraph's <w:pPr> — a different element that isn't
// a direct child of <w:body>, so splicing against it would throw.
function findDirectChildElementNS(parent: Element, namespaceURI: string, localName: string) {
  return findDirectChildElementsNS(parent, namespaceURI, localName)[0] ?? null;
}

function findDirectChildElementsNS(
  parent: Element,
  namespaceURI: string,
  localName: string
): Element[] {
  const result: Element[] = [];
  for (const child of parent.childNodes) {
    if (
      child.nodeType === ELEMENT_NODE &&
      child.namespaceURI === namespaceURI &&
      child.localName === localName
    ) {
      result.push(child as Element);
    }
  }
  return result;
}

// Locates a named part (e.g. "/word/comments.xml") among the package's
// <pkg:part> siblings and returns its <pkg:xmlData>'s single root element —
// design spec's "Core document model": "lazily locate comments.xml/
// styles.xml/numbering.xml parts if present". Returns null if the part
// doesn't exist (a document with no comments has no comments.xml part).
function findPartRoot(pkg: Document, partName: string): Element | null {
  for (const part of pkg.getElementsByTagNameNS(PKG_NS, "part")) {
    if (part.getAttributeNS(PKG_NS, "name") !== partName) continue;
    const xmlData = findDirectChildElementNS(part, PKG_NS, "xmlData");
    if (!xmlData) return null;
    for (const child of xmlData.childNodes) {
      if (child.nodeType === ELEMENT_NODE) return child as Element;
    }
    return null;
  }
  return null;
}

// Subtree search within a single paragraph is safe — <w:t> only ever nests
// inside that paragraph's own <w:r> children, never inside a further-nested
// paragraph (paragraphs don't nest).
function paragraphText(paragraph: Element): string {
  let text = "";
  for (const t of paragraph.getElementsByTagNameNS(W_NS, "t")) {
    text += t.textContent ?? "";
  }
  return text;
}

export class FlatOpcDocument {
  private readonly seedOoxml: string;
  private readonly platform: SupportedPlatform;
  private xmlDoc: Document;
  // Real Word's getOoxml() output always carries one more id-bearing
  // paragraph mark than body.paragraphs reports — a trailing mark whose id
  // churns on every getOoxml() call regardless of platform or edits. Modeled
  // as a real, always-present, always-last <w:p>, tracked by direct object
  // reference (not positionally) so it survives being pushed around by
  // inserts. See design spec's "ParaId stability model".
  private trailingMark!: Element;

  constructor(seedOoxml: string, platform: SupportedPlatform = "PC") {
    this.seedOoxml = seedOoxml;
    this.platform = platform;
    this.xmlDoc = new DOMParser().parseFromString(seedOoxml, "text/xml");
    this.ensureTrailingMark();
  }

  get bodyElement(): Element {
    const root = this.xmlDoc.documentElement;
    const body = root && findFirstDescendantElementNS(root, W_NS, "body");
    if (!body) {
      throw new Error("FlatOpcDocument: seed OOXML has no <w:body> element");
    }
    return body;
  }

  private ensureTrailingMark(): void {
    const body = this.bodyElement;
    const phantom = this.createParagraph("");
    const sectPr = findDirectChildElementNS(body, W_NS, "sectPr");
    if (sectPr) {
      body.insertBefore(phantom, sectPr);
    } else {
      body.appendChild(phantom);
    }
    this.trailingMark = phantom;
  }

  // Direct-child <w:p> elements that represent real, user-visible content —
  // excludes the trailing mark, matching what body.paragraphs reports. Does
  // NOT include paragraphs nested inside table cells — see getAllParagraphs().
  private getRealParagraphs(): Element[] {
    return findDirectChildElementsNS(this.bodyElement, W_NS, "p").filter(
      (p) => p !== this.trailingMark
    );
  }

  // Every <w:p> anywhere under <w:body>, including ones nested inside table
  // cells — excludes the trailing mark. Unlike getRealParagraphs() (direct
  // children only, matching what body.paragraphs reports), this is for
  // whole-document operations that must reach every paragraph regardless of
  // nesting: OfficeOnline's paraId churn is one such case (see
  // churnIdsForOfficeOnlineRead()) — the spec's "every getOoxml() call...
  // regenerates fresh ids... for the whole document" doesn't carve out an
  // exception for paragraphs that happen to live inside a table.
  private getAllParagraphs(): Element[] {
    return Array.from(this.bodyElement.getElementsByTagNameNS(W_NS, "p")).filter(
      (p) => p !== this.trailingMark
    );
  }

  private createRun(text: string): Element {
    const run = this.xmlDoc.createElementNS(W_NS, "w:r");
    const textNode = this.xmlDoc.createElementNS(W_NS, "w:t");
    textNode.appendChild(this.xmlDoc.createTextNode(text));
    run.appendChild(textNode);
    return run;
  }

  private createParagraph(text: string): Element {
    const paragraph = this.xmlDoc.createElementNS(W_NS, "w:p");
    paragraph.appendChild(this.createRun(text));
    assignFreshIds(paragraph);
    return paragraph;
  }

  // Splices `node` immediately after `anchor` within `parent` — shared by
  // insertParagraphAfter and insertOoxmlAfter, which both do "insert after"
  // splicing, once for a single paragraph and once looped for a fragment's
  // paragraphs.
  private insertNodeAfterAnchor(parent: Node, anchor: Node, node: Element): void {
    if (anchor.nextSibling) {
      parent.insertBefore(node, anchor.nextSibling);
    } else {
      parent.appendChild(node);
    }
  }

  appendParagraph(text: string): Element {
    const paragraph = this.createParagraph(text);
    this.bodyElement.insertBefore(paragraph, this.trailingMark);
    return paragraph;
  }

  insertParagraphBefore(target: Element, text: string): Element {
    const paragraph = this.createParagraph(text);
    if (!target.parentNode) {
      throw new Error("FlatOpcDocument.insertParagraphBefore: target has no parent");
    }
    target.parentNode.insertBefore(paragraph, target);
    return paragraph;
  }

  insertParagraphAfter(target: Element, text: string): Element {
    const paragraph = this.createParagraph(text);
    applyInsertAfterMarkShift(this.platform, target, paragraph);
    const parent = target.parentNode;
    if (!parent) {
      throw new Error("FlatOpcDocument.insertParagraphAfter: target has no parent");
    }
    this.insertNodeAfterAnchor(parent, target, paragraph);
    return paragraph;
  }

  replaceParagraphContent(target: Element, text: string): void {
    while (target.firstChild) {
      target.removeChild(target.firstChild);
    }
    target.appendChild(this.createRun(text));
  }

  insertParagraphAsFirstChild(text: string): Element {
    const body = this.bodyElement;
    const paragraph = this.createParagraph(text);
    // insertBefore(node, null) appends — correctly handles an empty body too.
    body.insertBefore(paragraph, body.firstChild);
    return paragraph;
  }

  replaceBodyContent(text: string): Element {
    const body = this.bodyElement;
    while (body.firstChild) {
      body.removeChild(body.firstChild);
    }
    const paragraph = this.createParagraph(text);
    body.appendChild(paragraph);
    this.ensureTrailingMark();
    return paragraph;
  }

  // A <w:pPr> (paragraph properties), if present, must stay the first child
  // per the OOXML schema — new content is inserted after it, not literally
  // as the first child.
  prependTextInParagraph(target: Element, text: string): void {
    const pPr = findDirectChildElementNS(target, W_NS, "pPr");
    const insertBeforeNode = pPr ? pPr.nextSibling : target.firstChild;
    target.insertBefore(this.createRun(text), insertBeforeNode);
  }

  appendTextInParagraph(target: Element, text: string): void {
    target.appendChild(this.createRun(text));
  }

  // Real Office.js's insertOoxml exchanges the same Flat-OPC package shape
  // as getOoxml — not a bare fragment — so the incoming string is parsed the
  // same way a seed document is, then its paragraphs are imported into this
  // document's own DOM (importNode, since nodes can't move between distinct
  // xmldom Document instances directly). Imported paragraphs always get a
  // fresh id, same as any other newly-inserted content — never keep whatever
  // id happened to be in the source fragment, which could collide with an
  // existing id elsewhere in this document.
  private importOoxmlParagraphs(ooxml: string): Element[] {
    const fragmentDoc = new FlatOpcDocument(ooxml);
    const sourceParagraphs = fragmentDoc.getRealParagraphs();
    return sourceParagraphs.map((p) => {
      const imported = this.xmlDoc.importNode(p, true);
      assignFreshIds(imported);
      return imported;
    });
  }

  insertOoxmlBefore(target: Element, ooxml: string): Element[] {
    const nodes = this.importOoxmlParagraphs(ooxml);
    if (!target.parentNode) {
      throw new Error("FlatOpcDocument.insertOoxmlBefore: target has no parent");
    }
    for (const node of nodes) {
      target.parentNode.insertBefore(node, target);
    }
    return nodes;
  }

  insertOoxmlAfter(target: Element, ooxml: string): Element[] {
    const nodes = this.importOoxmlParagraphs(ooxml);
    const parent = target.parentNode;
    if (!parent) {
      throw new Error("FlatOpcDocument.insertOoxmlAfter: target has no parent");
    }
    // Only the paragraph immediately adjacent to the anchor mark-shifts —
    // the confirmed evidence (ticket 006) only covers a single inserted
    // paragraph; extending that to paragraphs 2+ of a multi-paragraph
    // fragment would be an unconfirmed guess, not a modeled behavior.
    if (nodes[0]) {
      applyInsertAfterMarkShift(this.platform, target, nodes[0]);
    }
    let anchor: Element = target;
    for (const node of nodes) {
      this.insertNodeAfterAnchor(parent, anchor, node);
      anchor = node;
    }
    return nodes;
  }

  insertOoxmlAsFirstChild(ooxml: string): Element[] {
    const body = this.bodyElement;
    const nodes = this.importOoxmlParagraphs(ooxml);
    const anchor = body.firstChild;
    for (const node of nodes) {
      // insertBefore(node, null) appends — correctly handles an empty body.
      body.insertBefore(node, anchor);
    }
    return nodes;
  }

  insertOoxmlAsLastChild(ooxml: string): Element[] {
    const nodes = this.importOoxmlParagraphs(ooxml);
    for (const node of nodes) {
      this.bodyElement.insertBefore(node, this.trailingMark);
    }
    return nodes;
  }

  replaceOoxmlBodyContent(ooxml: string): Element[] {
    const body = this.bodyElement;
    const nodes = this.importOoxmlParagraphs(ooxml);
    while (body.firstChild) {
      body.removeChild(body.firstChild);
    }
    for (const node of nodes) {
      body.appendChild(node);
    }
    this.ensureTrailingMark();
    return nodes;
  }

  replaceOoxmlAtTarget(target: Element, ooxml: string): Element[] {
    const nodes = this.importOoxmlParagraphs(ooxml);
    const parent = target.parentNode;
    if (!parent) {
      throw new Error("FlatOpcDocument.replaceOoxmlAtTarget: target has no parent");
    }
    for (const node of nodes) {
      parent.insertBefore(node, target);
    }
    parent.removeChild(target);
    return nodes;
  }

  getParagraphText(paragraph: Element): string {
    return paragraphText(paragraph);
  }

  // Shared by getOoxml() and getRangeOoxml() — the design spec's "ParaId
  // stability model" is explicit that OfficeOnline regenerates ids "for the
  // whole document" on "Body or any Range" reads alike.
  private churnIdsForOfficeOnlineRead(): void {
    regenerateAllIdsForOfficeOnline([...this.getAllParagraphs(), this.trailingMark]);
  }

  getOoxml(): string {
    if (this.platform === "OfficeOnline") {
      this.churnIdsForOfficeOnlineRead();
    } else {
      // PC/Mac: real content stays stable across whole-document reads —
      // only the trailing mark churns, on every single call, regardless of
      // edits.
      assignFreshIds(this.trailingMark);
    }
    return new XMLSerializer().serializeToString(this.xmlDoc);
  }

  // Range.getOoxml()'s real shape: a range-scoped fragment re-wrapped as its
  // own Flat-OPC package, not the whole document (design spec's "Core
  // document model"). The target's own serialization already carries its
  // required namespace declarations inline (xmldom adds them automatically
  // when a subtree is serialized outside its full ancestor chain), so the
  // wrapper only needs to declare the main `w:` namespace for its own
  // <w:document>/<w:body> elements.
  //
  // No PC/Mac trailing-mark churn here, unlike getOoxml(): the mark is never
  // part of a range-scoped fragment's output, so churning it would be a
  // side effect with no observable justification — the spec's "churns every
  // call" language describes Body's whole-document reads specifically, and
  // extending it to scoped Range reads would be an unconfirmed guess, not a
  // modeled behavior (same caution as paraId.ts's rsid omission).
  getRangeOoxml(target: Element): string {
    if (this.platform === "OfficeOnline") {
      this.churnIdsForOfficeOnlineRead();
    }
    const targetXml = new XMLSerializer().serializeToString(target);
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<pkg:package xmlns:pkg="http://schemas.microsoft.com/office/2006/xmlPackage">
  <pkg:part pkg:name="/word/document.xml" pkg:contentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml">
    <pkg:xmlData>
      <w:document xmlns:w="${W_NS}">
        <w:body>
          ${targetXml}
        </w:body>
      </w:document>
    </pkg:xmlData>
  </pkg:part>
</pkg:package>
`;
  }

  // Design spec's "Core document model" lists `deleteNode` as one of
  // FlatOpcDocument's low-level mutation primitives, alongside insertAt/search.
  deleteNode(target: Element): void {
    if (!target.parentNode) {
      throw new Error("FlatOpcDocument.deleteNode: target has no parent");
    }
    target.parentNode.removeChild(target);
  }

  getBodyText(): string {
    return this.getRealParagraphs().map(paragraphText).join("\n");
  }

  // Public counterpart of getRealParagraphs() — for callers outside this
  // class that need the actual elements, not just their joined text.
  getParagraphElements(): Element[] {
    return this.getRealParagraphs();
  }

  // Direct-child <w:tbl> elements — for Body.tables (TableCollection) to
  // enumerate, mirroring getParagraphElements()'s role for body.paragraphs.
  getTableElements(): Element[] {
    return findDirectChildElementsNS(this.bodyElement, W_NS, "tbl");
  }

  private get commentsRoot(): Element | null {
    return findPartRoot(this.xmlDoc, "/word/comments.xml");
  }

  private get commentsExtendedRoot(): Element | null {
    return findPartRoot(this.xmlDoc, "/word/commentsExtended.xml");
  }

  // Every <w:comment> in comments.xml, document order — top-level comments
  // and threaded replies alike (a reply is just another <w:comment> element;
  // only its commentsExtended.xml entry distinguishes it — see
  // getCommentReplyElements()).
  getCommentElements(): Element[] {
    const root = this.commentsRoot;
    return root ? findDirectChildElementsNS(root, W_NS, "comment") : [];
  }

  getTopLevelCommentElements(): Element[] {
    return this.getCommentElements().filter(
      (c) => !this.findCommentExtendedEntry(c)?.getAttributeNS(W15_NS, "parentParaId")
    );
  }

  // A reply's commentsExtended.xml entry has w15:parentParaId set to its
  // parent's own <w:p>/@w14:paraId — threading is expressed via paragraph
  // ids, not comment ids (ECMA-376's Word 2013+ commentsExtended extension).
  getCommentReplyElements(comment: Element): Element[] {
    const ownParaId = this.getCommentOwnParaId(comment);
    if (!ownParaId) return [];
    return this.getCommentElements().filter(
      (c) => this.findCommentExtendedEntry(c)?.getAttributeNS(W15_NS, "parentParaId") === ownParaId
    );
  }

  getCommentId(comment: Element): string {
    return comment.getAttributeNS(W_NS, "id") ?? "";
  }

  getCommentAuthor(comment: Element): string {
    return comment.getAttributeNS(W_NS, "author") ?? "";
  }

  // A comment's content nests identically to body paragraphs — reuses the
  // same paragraph text read used everywhere else. Comments may contain
  // multiple paragraphs; joined with "\n", same as getBodyText().
  getCommentContent(comment: Element): string {
    return findDirectChildElementsNS(comment, W_NS, "p").map(paragraphText).join("\n");
  }

  getCommentDate(comment: Element): Date {
    return new Date(comment.getAttributeNS(W_NS, "date") ?? "");
  }

  // Absent commentsExtended.xml, or no matching entry, means "not resolved"
  // — a comment with no extended metadata was never marked done.
  getCommentResolved(comment: Element): boolean {
    return this.findCommentExtendedEntry(comment)?.getAttributeNS(W15_NS, "done") === "1";
  }

  // Uses the comment's FIRST paragraph. Every fixture and real-world use
  // case this shim has been tested against has single-paragraph comments,
  // so this is unconfirmed for multi-paragraph ones — real Word may key
  // w15:paraId off the LAST paragraph instead. Best-effort, not a verified
  // assumption (same caution as paraId.ts's rsid omission).
  private getCommentOwnParaId(comment: Element): string | null {
    const p = findDirectChildElementNS(comment, W_NS, "p");
    return p && p.getAttributeNS(W14_NS, "paraId");
  }

  private findCommentExtendedEntry(comment: Element): Element | null {
    const root = this.commentsExtendedRoot;
    const ownParaId = root && this.getCommentOwnParaId(comment);
    if (!root || !ownParaId) return null;
    return (
      findDirectChildElementsNS(root, W15_NS, "commentEx").find(
        (entry) => entry.getAttributeNS(W15_NS, "paraId") === ownParaId
      ) ?? null
    );
  }

  // Design spec's "Core document model" lists `search(pattern)` as one of
  // FlatOpcDocument's primitives, alongside insertAt/deleteNode. Plain
  // substring matching only (issue #13: wildcard syntax and other
  // SearchOptions fields are explicitly out of scope for v1) — never
  // regex/pattern interpretation, so wildcard-special characters in
  // `searchText` are always matched literally, never silently
  // misinterpreted. Only searches direct-child body paragraphs, same scope
  // as getRealParagraphs() — table-cell paragraphs are out of scope (#13
  // predates Tables (#15); extending search() into table cells isn't asked
  // for by either issue).
  search(searchText: string, matchCase: boolean): Element[] {
    return this.getRealParagraphs().filter((p) => {
      const text = paragraphText(p);
      return matchCase
        ? text.includes(searchText)
        : text.toLowerCase().includes(searchText.toLowerCase());
    });
  }

  // True column count lives in w:tblGrid independent of per-row w:tc counts
  // under gridSpan/merges — never derive it from max(tr children) (issue
  // #15's AC).
  getTableColumnCount(table: Element): number {
    const tblGrid = findDirectChildElementNS(table, W_NS, "tblGrid");
    return tblGrid ? findDirectChildElementsNS(tblGrid, W_NS, "gridCol").length : 0;
  }

  getTableRows(table: Element): Element[] {
    return findDirectChildElementsNS(table, W_NS, "tr");
  }

  getRowCells(row: Element): Element[] {
    return findDirectChildElementsNS(row, W_NS, "tc");
  }

  // Cell content nests identically to body paragraphs (w:tc > w:p > w:r >
  // w:t) — reuses the same paragraph text read used everywhere else, per
  // design spec's "Table OOXML shape".
  getCellText(cell: Element): string {
    return paragraphText(this.requireCellParagraph(cell));
  }

  setCellText(cell: Element, text: string): void {
    this.replaceParagraphContent(this.requireCellParagraph(cell), text);
  }

  private requireCellParagraph(cell: Element): Element {
    const paragraph = findDirectChildElementNS(cell, W_NS, "p");
    if (!paragraph) {
      throw new Error("FlatOpcDocument: table cell has no <w:p> content");
    }
    return paragraph;
  }

  private findCellPropertyChild(cell: Element, localName: string): Element | null {
    const tcPr = findDirectChildElementNS(cell, W_NS, "tcPr");
    return tcPr && findDirectChildElementNS(tcPr, W_NS, localName);
  }

  // w:gridSpan absent means "spans exactly 1 grid column" (ECMA-376
  // §17.4.17) — not a merge.
  getCellGridSpan(cell: Element): number {
    const gridSpan = this.findCellPropertyChild(cell, "gridSpan");
    const val = gridSpan?.getAttributeNS(W_NS, "val");
    return val ? Number.parseInt(val, 10) : 1;
  }

  // w:vMerge present with no @w:val is the schema default "continue"
  // (ECMA-376 §17.4.86) — distinct from the element being entirely absent
  // (no merge at all).
  getCellVMerge(cell: Element): "Continue" | "Restart" | undefined {
    const vMerge = this.findCellPropertyChild(cell, "vMerge");
    if (!vMerge) return undefined;
    const val = vMerge.getAttributeNS(W_NS, "val");
    return val === "restart" ? "Restart" : "Continue";
  }

  // Reasonable schema-valid defaults for w:tcW (auto width) absent a
  // captured real-Word fixture to confirm Word's own authoring values —
  // same "unresolved without a captured fixture" caveat as tblStyle/tblLook
  // boilerplate (design spec's "Table OOXML shape").
  private createTableCell(text: string): Element {
    const cell = this.xmlDoc.createElementNS(W_NS, "w:tc");
    const tcPr = this.xmlDoc.createElementNS(W_NS, "w:tcPr");
    const tcW = this.xmlDoc.createElementNS(W_NS, "w:tcW");
    tcW.setAttributeNS(W_NS, "w:w", "0");
    tcW.setAttributeNS(W_NS, "w:type", "auto");
    tcPr.appendChild(tcW);
    cell.appendChild(tcPr);
    cell.appendChild(this.createParagraph(text));
    return cell;
  }

  // w:tblPr/w:trPr are omitted entirely for v1 (design spec's "Table OOXML
  // shape") — nothing in the covered API surface needs table-wide or
  // row-wide formatting, so table/row construction never emits them.
  addTableRows(
    table: Element,
    insertLocation: "Start" | "End",
    rowCount: number,
    values?: string[][]
  ): Element[] {
    const columnCount = this.getTableColumnCount(table);
    const newRows: Element[] = [];
    for (let i = 0; i < rowCount; i++) {
      const row = this.xmlDoc.createElementNS(W_NS, "w:tr");
      for (let col = 0; col < columnCount; col++) {
        row.appendChild(this.createTableCell(values?.[i]?.[col] ?? ""));
      }
      newRows.push(row);
    }

    if (insertLocation === "Start") {
      const anchor = this.getTableRows(table)[0] ?? null;
      for (const row of newRows) {
        table.insertBefore(row, anchor);
      }
    } else {
      for (const row of newRows) {
        table.appendChild(row);
      }
    }
    return newRows;
  }

  reset(): void {
    this.xmlDoc = new DOMParser().parseFromString(this.seedOoxml, "text/xml");
    this.ensureTrailingMark();
  }
}
