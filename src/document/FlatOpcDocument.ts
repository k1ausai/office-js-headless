import { DOMParser, XMLSerializer, type Document, type Element, type Node } from "@xmldom/xmldom";
import { type SupportedPlatform } from "../office/context";
import {
  applyInsertAfterMarkShift,
  assignFreshIds,
  regenerateAllIdsForOfficeOnline,
} from "../word/paraId";

const W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
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
  // excludes the trailing mark, matching what body.paragraphs reports.
  private getRealParagraphs(): Element[] {
    return findDirectChildElementsNS(this.bodyElement, W_NS, "p").filter(
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

  // Shared by getOoxml() and getRangeOoxml() — every read-path serialization,
  // whole-document or range-scoped, applies the same platform-driven id churn
  // (design spec's "ParaId stability model": "every getOoxml() call — on Body
  // or any Range — regenerates fresh ids ... for the whole document").
  private applyReadTimeIdChurn(): void {
    if (this.platform === "OfficeOnline") {
      regenerateAllIdsForOfficeOnline([...this.getRealParagraphs(), this.trailingMark]);
    } else {
      // PC/Mac: real content stays stable across reads — only the trailing
      // mark churns, on every single call, regardless of edits.
      assignFreshIds(this.trailingMark);
    }
  }

  getOoxml(): string {
    this.applyReadTimeIdChurn();
    return new XMLSerializer().serializeToString(this.xmlDoc);
  }

  // Range.getOoxml()'s real shape: a range-scoped fragment re-wrapped as its
  // own Flat-OPC package, not the whole document (design spec's "Core
  // document model"). The target's own serialization already carries its
  // required namespace declarations inline (xmldom adds them automatically
  // when a subtree is serialized outside its full ancestor chain), so the
  // wrapper only needs to declare the main `w:` namespace for its own
  // <w:document>/<w:body> elements.
  getRangeOoxml(target: Element): string {
    this.applyReadTimeIdChurn();
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

  reset(): void {
    this.xmlDoc = new DOMParser().parseFromString(this.seedOoxml, "text/xml");
    this.ensureTrailingMark();
  }
}
