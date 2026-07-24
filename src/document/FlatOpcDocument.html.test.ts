import { type Element } from "@xmldom/xmldom";
import { describe, expect, it } from "vitest";
import { MINIMAL_SEED_OOXML } from "./__fixtures__/minimalSeed";
import { FlatOpcDocument } from "./FlatOpcDocument";

const W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

function firstParagraph(doc: FlatOpcDocument): Element {
  const paragraph = doc.bodyElement.getElementsByTagNameNS(W_NS, "p")[0];
  if (!paragraph) throw new Error("test fixture has no <w:p>");
  return paragraph;
}

describe("FlatOpcDocument HTML-to-OOXML conversion", () => {
  it("converts bare inline HTML (no <p> wrapper) into a single paragraph", () => {
    const doc = new FlatOpcDocument(MINIMAL_SEED_OOXML);
    doc.insertHtmlAsLastChild("Plain text.");
    expect(doc.getBodyText()).toBe("Seed paragraph.\nPlain text.");
  });

  it("converts multiple <p> elements into multiple paragraphs, preserving order", () => {
    const doc = new FlatOpcDocument(MINIMAL_SEED_OOXML);
    doc.insertHtmlAsLastChild("<p>One.</p><p>Two.</p>");
    expect(doc.getBodyText()).toBe("Seed paragraph.\nOne.\nTwo.");
  });

  it("<b> and <strong> both apply bold formatting (w:b in w:rPr)", () => {
    const doc = new FlatOpcDocument(MINIMAL_SEED_OOXML);
    doc.insertHtmlAsLastChild("<p><b>Bold one.</b></p><p><strong>Bold two.</strong></p>");
    const ooxml = doc.getOoxml();
    expect(ooxml).toContain("<w:b/>");
    expect(doc.getBodyText()).toBe("Seed paragraph.\nBold one.\nBold two.");
  });

  it("<i> and <em> both apply italic formatting (w:i in w:rPr)", () => {
    const doc = new FlatOpcDocument(MINIMAL_SEED_OOXML);
    doc.insertHtmlAsLastChild("<p><i>Italic one.</i></p><p><em>Italic two.</em></p>");
    const ooxml = doc.getOoxml();
    expect(ooxml).toContain("<w:i/>");
    expect(doc.getBodyText()).toBe("Seed paragraph.\nItalic one.\nItalic two.");
  });

  it("nested formatting applies both w:b and w:i to the same run", () => {
    const doc = new FlatOpcDocument(MINIMAL_SEED_OOXML);
    doc.insertHtmlAsLastChild("<p><b><i>Bold italic.</i></b></p>");
    const ooxml = doc.getOoxml();
    expect(ooxml).toContain("<w:b/>");
    expect(ooxml).toContain("<w:i/>");
    expect(doc.getBodyText()).toBe("Seed paragraph.\nBold italic.");
  });

  it("plain text alongside formatted text in the same paragraph produces separate runs", () => {
    const doc = new FlatOpcDocument(MINIMAL_SEED_OOXML);
    doc.insertHtmlAsLastChild("<p>Plain <b>bold</b> plain again.</p>");
    expect(doc.getBodyText()).toBe("Seed paragraph.\nPlain bold plain again.");
  });
});

describe("FlatOpcDocument insertHtml* location semantics (mirrors insertOoxml*)", () => {
  it("insertHtmlBefore splices before the target, preserving order", () => {
    const doc = new FlatOpcDocument(MINIMAL_SEED_OOXML);
    const target = firstParagraph(doc);
    doc.insertHtmlBefore(target, "<p>One.</p><p>Two.</p>");
    const ooxml = doc.getOoxml();
    expect(ooxml.indexOf("One.")).toBeLessThan(ooxml.indexOf("Two."));
    expect(ooxml.indexOf("Two.")).toBeLessThan(ooxml.indexOf("Seed paragraph."));
  });

  it("insertHtmlAfter splices after the target, preserving order", () => {
    const doc = new FlatOpcDocument(MINIMAL_SEED_OOXML);
    const target = firstParagraph(doc);
    doc.appendParagraph("Gamma.");
    doc.insertHtmlAfter(target, "<p>One.</p><p>Two.</p>");
    const ooxml = doc.getOoxml();
    expect(ooxml.indexOf("Seed paragraph.")).toBeLessThan(ooxml.indexOf("One."));
    expect(ooxml.indexOf("One.")).toBeLessThan(ooxml.indexOf("Two."));
    expect(ooxml.indexOf("Two.")).toBeLessThan(ooxml.indexOf("Gamma."));
  });

  it("insertHtmlAsFirstChild inserts before all existing content", () => {
    const doc = new FlatOpcDocument(MINIMAL_SEED_OOXML);
    doc.insertHtmlAsFirstChild("<p>New first.</p>");
    const ooxml = doc.getOoxml();
    expect(ooxml.indexOf("New first.")).toBeLessThan(ooxml.indexOf("Seed paragraph."));
  });

  it("insertHtmlAsLastChild inserts after all existing content, before sectPr", () => {
    const doc = new FlatOpcDocument(MINIMAL_SEED_OOXML);
    doc.insertHtmlAsLastChild("<p>New last.</p>");
    const ooxml = doc.getOoxml();
    expect(ooxml.indexOf("Seed paragraph.")).toBeLessThan(ooxml.indexOf("New last."));
  });

  it("replaceHtmlBodyContent clears the whole body and inserts the converted paragraphs", () => {
    const doc = new FlatOpcDocument(MINIMAL_SEED_OOXML);
    doc.replaceHtmlBodyContent("<p>Only this.</p>");
    expect(doc.getBodyText()).toBe("Only this.");
  });

  it("replaceHtmlAtTarget removes the target and inserts the converted paragraphs in its place", () => {
    const doc = new FlatOpcDocument(MINIMAL_SEED_OOXML);
    const target = firstParagraph(doc);
    doc.appendParagraph("Survivor.");
    doc.replaceHtmlAtTarget(target, "<p>Replacement.</p>");
    expect(doc.getBodyText()).toBe("Replacement.\nSurvivor.");
  });
});
