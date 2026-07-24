import { describe, expect, it } from "vitest";
import { MINIMAL_SEED_OOXML } from "./__fixtures__/minimalSeed";
import { FlatOpcDocument } from "./FlatOpcDocument";
import { STYLE_SEED_OOXML } from "./__fixtures__/styleSeed";

describe("FlatOpcDocument style primitives", () => {
  it("getStyleElements returns every w:style, in document order", () => {
    const doc = new FlatOpcDocument(STYLE_SEED_OOXML);
    const styles = doc.getStyleElements();
    expect(styles.map((s) => doc.getStyleId(s))).toEqual([
      "Normal",
      "Heading1",
      "Title",
      "MyCustomStyle",
    ]);
  });

  it("getStyleNameLocal reads w:name/@w:val", () => {
    const doc = new FlatOpcDocument(STYLE_SEED_OOXML);
    const [, heading1] = doc.getStyleElements();
    expect(doc.getStyleNameLocal(heading1!)).toBe("heading 1");
  });

  it("getStyleType reads w:type", () => {
    const doc = new FlatOpcDocument(STYLE_SEED_OOXML);
    const [normal] = doc.getStyleElements();
    expect(doc.getStyleType(normal!)).toBe("Paragraph");
  });

  it("getStyleBuiltIn is true for a style without w:customStyle, false for one with it", () => {
    const doc = new FlatOpcDocument(STYLE_SEED_OOXML);
    const [normal, , , custom] = doc.getStyleElements();
    expect(doc.getStyleBuiltIn(normal!)).toBe(true);
    expect(doc.getStyleBuiltIn(custom!)).toBe(false);
  });

  it("returns no styles when the document has no styles.xml part", () => {
    const doc = new FlatOpcDocument(MINIMAL_SEED_OOXML);
    expect(doc.getStyleElements()).toEqual([]);
  });
});
