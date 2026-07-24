import { describe, expect, it } from "vitest";
import { FlatOpcDocument } from "../document/FlatOpcDocument";
import { STYLE_SEED_OOXML } from "../document/__fixtures__/styleSeed";
import { Style } from "./style";

function styles(doc: FlatOpcDocument): Style[] {
  return doc.getStyleElements().map((s) => new Style(doc, s));
}

describe("Style load/sync gating", () => {
  it("reading a property without calling .load() first throws PropertyNotLoaded", () => {
    const doc = new FlatOpcDocument(STYLE_SEED_OOXML);
    const [style] = styles(doc);
    expect(() => style!.nameLocal).toThrow(/PropertyNotLoaded|not available/);
  });

  it("reading after .load() but before sync() still throws — identical to never-loaded", () => {
    const doc = new FlatOpcDocument(STYLE_SEED_OOXML);
    const [style] = styles(doc);
    style!.load("nameLocal");
    expect(() => style!.nameLocal).toThrow(/PropertyNotLoaded|not available/);
  });

  it("reading after .load() + sync() returns the current values", () => {
    const doc = new FlatOpcDocument(STYLE_SEED_OOXML);
    const [, heading1] = styles(doc);
    heading1!.load(["id", "nameLocal", "type", "builtIn"]);
    heading1!.sync();

    expect(heading1!.id).toBe("Heading1");
    expect(heading1!.nameLocal).toBe("heading 1");
    expect(heading1!.type).toBe("Paragraph");
    expect(heading1!.builtIn).toBe(true);
  });

  it("a custom style reports builtIn: false", () => {
    const doc = new FlatOpcDocument(STYLE_SEED_OOXML);
    const [, , , custom] = styles(doc);
    custom!.load(["id", "builtIn"]);
    custom!.sync();

    expect(custom!.id).toBe("MyCustomStyle");
    expect(custom!.builtIn).toBe(false);
  });
});
