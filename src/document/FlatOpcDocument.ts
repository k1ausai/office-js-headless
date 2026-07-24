import { DOMParser, XMLSerializer, type Document, type Element } from "@xmldom/xmldom";

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
  for (const child of parent.childNodes) {
    if (
      child.nodeType === ELEMENT_NODE &&
      child.namespaceURI === namespaceURI &&
      child.localName === localName
    ) {
      return child as Element;
    }
  }
  return null;
}

export class FlatOpcDocument {
  private readonly seedOoxml: string;
  private xmlDoc: Document;

  constructor(seedOoxml: string) {
    this.seedOoxml = seedOoxml;
    this.xmlDoc = new DOMParser().parseFromString(seedOoxml, "text/xml");
  }

  get bodyElement(): Element {
    const root = this.xmlDoc.documentElement;
    const body = root && findFirstDescendantElementNS(root, W_NS, "body");
    if (!body) {
      throw new Error("FlatOpcDocument: seed OOXML has no <w:body> element");
    }
    return body;
  }

  appendParagraph(text: string): Element {
    const body = this.bodyElement;
    const paragraph = this.xmlDoc.createElementNS(W_NS, "w:p");
    const run = this.xmlDoc.createElementNS(W_NS, "w:r");
    const textNode = this.xmlDoc.createElementNS(W_NS, "w:t");
    textNode.appendChild(this.xmlDoc.createTextNode(text));
    run.appendChild(textNode);
    paragraph.appendChild(run);

    const sectPr = findDirectChildElementNS(body, W_NS, "sectPr");
    if (sectPr) {
      body.insertBefore(paragraph, sectPr);
    } else {
      body.appendChild(paragraph);
    }
    return paragraph;
  }

  getOoxml(): string {
    return new XMLSerializer().serializeToString(this.xmlDoc);
  }

  reset(): void {
    this.xmlDoc = new DOMParser().parseFromString(this.seedOoxml, "text/xml");
  }
}
