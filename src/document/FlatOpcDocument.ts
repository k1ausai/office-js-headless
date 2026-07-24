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
    return paragraph;
  }

  appendParagraph(text: string): Element {
    const body = this.bodyElement;
    const paragraph = this.createParagraph(text);
    const sectPr = findDirectChildElementNS(body, W_NS, "sectPr");
    if (sectPr) {
      body.insertBefore(paragraph, sectPr);
    } else {
      body.appendChild(paragraph);
    }
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
    const parent = target.parentNode;
    if (!parent) {
      throw new Error("FlatOpcDocument.insertParagraphAfter: target has no parent");
    }
    if (target.nextSibling) {
      parent.insertBefore(paragraph, target.nextSibling);
    } else {
      parent.appendChild(paragraph);
    }
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

  getParagraphText(paragraph: Element): string {
    return paragraphText(paragraph);
  }

  getOoxml(): string {
    return new XMLSerializer().serializeToString(this.xmlDoc);
  }

  getBodyText(): string {
    const paragraphs = findDirectChildElementsNS(this.bodyElement, W_NS, "p");
    return paragraphs.map(paragraphText).join("\n");
  }

  reset(): void {
    this.xmlDoc = new DOMParser().parseFromString(this.seedOoxml, "text/xml");
  }
}
