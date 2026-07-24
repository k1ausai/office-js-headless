import { describe, expect, it } from "vitest";
import { MINIMAL_SEED_OOXML } from "./__fixtures__/minimalSeed";
import { STYLE_SEED_OOXML } from "./__fixtures__/styleSeed";
import { FlatOpcDocument } from "./FlatOpcDocument";

const W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

describe("FlatOpcDocument.mergeStylesXml", () => {
  it("creates a styles.xml part when the document has none", () => {
    const doc = new FlatOpcDocument(MINIMAL_SEED_OOXML);
    expect(doc.getStyleElements()).toEqual([]);

    doc.mergeStylesXml(
      `<w:styles xmlns:w="${W_NS}"><w:style w:type="paragraph" w:styleId="Imported"><w:name w:val="Imported"/></w:style></w:styles>`
    );

    const styles = doc.getStyleElements();
    expect(styles.map((s) => doc.getStyleId(s))).toEqual(["Imported"]);
  });

  it("appends into an existing styles.xml without deduplicating collisions ('without real conflict resolution')", () => {
    const doc = new FlatOpcDocument(STYLE_SEED_OOXML);
    const before = doc.getStyleElements().length;

    doc.mergeStylesXml(
      `<w:styles xmlns:w="${W_NS}"><w:style w:type="paragraph" w:styleId="Normal"><w:name w:val="Normal"/></w:style></w:styles>`
    );

    const after = doc.getStyleElements();
    expect(after).toHaveLength(before + 1);
    expect(after.filter((s) => doc.getStyleId(s) === "Normal")).toHaveLength(2);
  });
});

describe("FlatOpcDocument.mergeNumberingXml", () => {
  it("creates a numbering.xml part when the document has none, and its content round-trips through getOoxml()", () => {
    const doc = new FlatOpcDocument(MINIMAL_SEED_OOXML);

    doc.mergeNumberingXml(
      `<w:numbering xmlns:w="${W_NS}"><w:abstractNum w:abstractNumId="0"/></w:numbering>`
    );

    expect(doc.getOoxml()).toContain('w:abstractNumId="0"');
  });

  it("appends into an existing numbering.xml", () => {
    const doc = new FlatOpcDocument(MINIMAL_SEED_OOXML);
    doc.mergeNumberingXml(
      `<w:numbering xmlns:w="${W_NS}"><w:abstractNum w:abstractNumId="0"/></w:numbering>`
    );
    doc.mergeNumberingXml(
      `<w:numbering xmlns:w="${W_NS}"><w:abstractNum w:abstractNumId="1"/></w:numbering>`
    );

    const ooxml = doc.getOoxml();
    expect(ooxml).toContain('w:abstractNumId="0"');
    expect(ooxml).toContain('w:abstractNumId="1"');
  });
});
