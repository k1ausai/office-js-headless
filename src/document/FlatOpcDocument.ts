import { DOMParser, XMLSerializer, type Document, type Element } from "@xmldom/xmldom";

const W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

function findFirstElementByTagNameNS(root: Element, namespaceURI: string, localName: string) {
  const matches = root.getElementsByTagNameNS(namespaceURI, localName);
  return matches.length > 0 ? matches[0] : null;
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
    const body = root && findFirstElementByTagNameNS(root, W_NS, "body");
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

    const sectPr = findFirstElementByTagNameNS(body, W_NS, "sectPr");
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
