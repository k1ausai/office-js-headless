import { afterEach, describe, expect, it } from "vitest";
import { installHeadlessOffice } from "./install";
import { MINIMAL_SEED_OOXML } from "./document/__fixtures__/minimalSeed";

// The real global `Word`/`Office` types (from @types/office-js) describe the
// full production API surface, far more than this shim implements so far.
// Tests exercise the shim through this minimal, precise slice of that shape
// instead of casting through `any`.
interface MinimalWordContext {
  document: {
    body: {
      insertText(text: string, insertLocation: string): void;
      insertOoxml(ooxml: string, insertLocation: string): void;
    };
  };
  sync(): Promise<void>;
}
interface MinimalWordGlobal {
  run<T>(callback: (context: MinimalWordContext) => Promise<T>): Promise<T>;
  InsertLocation: { end: string; start: string };
  BuiltInStyleName: { heading1: string; normal: string };
}
interface MinimalOfficeGlobal {
  context: { platform: number; requirements: { isSetSupported(name: string): boolean } };
  PlatformType: Record<string, number>;
  onReady(): Promise<unknown>;
}

function getInstalledWord(): MinimalWordGlobal {
  return (globalThis as unknown as { Word: MinimalWordGlobal }).Word;
}

function getInstalledOffice(): MinimalOfficeGlobal {
  return (globalThis as unknown as { Office: MinimalOfficeGlobal }).Office;
}

describe("installHeadlessOffice", () => {
  afterEach(() => {
    // Best-effort cleanup in case a test fails before calling dispose().
    delete (globalThis as Record<string, unknown>).Word;
    delete (globalThis as Record<string, unknown>).Office;
  });

  it("installs Word as a global with run() and InsertLocation", () => {
    const office = installHeadlessOffice({ seedOoxml: MINIMAL_SEED_OOXML });
    expect(typeof getInstalledWord().run).toBe("function");
    expect(getInstalledWord().InsertLocation.end).toBe("End");
    office.dispose();
  });

  it("exposes Word.BuiltInStyleName, matching real Word's built-in style ids", () => {
    const office = installHeadlessOffice({ seedOoxml: MINIMAL_SEED_OOXML });
    expect(getInstalledWord().BuiltInStyleName.heading1).toBe("Heading1");
    expect(getInstalledWord().BuiltInStyleName.normal).toBe("Normal");
    office.dispose();
  });

  it("runs add-in-shaped code end-to-end: Word.run -> insertText -> sync -> getOoxml", async () => {
    const office = installHeadlessOffice({ seedOoxml: MINIMAL_SEED_OOXML });

    await getInstalledWord().run(async (context) => {
      context.document.body.insertText("hello", getInstalledWord().InsertLocation.end);
      await context.sync();
    });

    expect(office.getOoxml()).toContain("hello");
    office.dispose();
  });

  it("dispose() tears down the installed globals", () => {
    const office = installHeadlessOffice({ seedOoxml: MINIMAL_SEED_OOXML });
    office.dispose();
    expect((globalThis as Record<string, unknown>).Word).toBeUndefined();
    expect((globalThis as Record<string, unknown>).Office).toBeUndefined();
  });

  it("reset() restores the document to its originally-seeded state", async () => {
    const office = installHeadlessOffice({ seedOoxml: MINIMAL_SEED_OOXML });

    await getInstalledWord().run(async (context) => {
      context.document.body.insertText("temporary", getInstalledWord().InsertLocation.end);
      await context.sync();
    });
    expect(office.getOoxml()).toContain("temporary");

    office.reset();
    expect(office.getOoxml()).not.toContain("temporary");
    office.dispose();
  });

  it("office.getOoxml() (the escape hatch) reads document state without needing a Word.run", () => {
    const office = installHeadlessOffice({ seedOoxml: MINIMAL_SEED_OOXML });
    expect(office.getOoxml()).toContain("Seed paragraph.");
    office.dispose();
  });

  it("defaults to platform PC when no platform option is given", async () => {
    const office = installHeadlessOffice({ seedOoxml: MINIMAL_SEED_OOXML });
    await getInstalledOffice().onReady();
    expect(getInstalledOffice().context.platform).toBe(getInstalledOffice().PlatformType.PC);
    office.dispose();
  });

  it("drives Office.context.platform per the platform option", async () => {
    for (const platform of ["PC", "Mac", "OfficeOnline"] as const) {
      const office = installHeadlessOffice({ seedOoxml: MINIMAL_SEED_OOXML, platform });
      await getInstalledOffice().onReady();
      expect(getInstalledOffice().context.platform).toBe(
        getInstalledOffice().PlatformType[platform]
      );
      office.dispose();
    }
  });

  it("Office.context throws before onReady() has resolved", () => {
    const office = installHeadlessOffice({ seedOoxml: MINIMAL_SEED_OOXML });
    expect(() => getInstalledOffice().context).toThrow();
    office.dispose();
  });

  it("insertOoxml applies the fragment on PC and Mac", async () => {
    const fragment = `<?xml version="1.0" encoding="UTF-8"?>
<pkg:package xmlns:pkg="http://schemas.microsoft.com/office/2006/xmlPackage">
  <pkg:part pkg:name="/word/document.xml" pkg:contentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml">
    <pkg:xmlData>
      <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
        <w:body>
          <w:p><w:r><w:t>Via insertOoxml.</w:t></w:r></w:p>
        </w:body>
      </w:document>
    </pkg:xmlData>
  </pkg:part>
</pkg:package>`;

    for (const platform of ["PC", "Mac"] as const) {
      const office = installHeadlessOffice({ seedOoxml: MINIMAL_SEED_OOXML, platform });

      await getInstalledWord().run(async (context) => {
        context.document.body.insertOoxml(fragment, getInstalledWord().InsertLocation.end);
        await context.sync();
      });

      expect(office.getOoxml()).toContain("Via insertOoxml.");
      office.dispose();
    }
  });

  it("insertOoxml rejects at sync() on OfficeOnline (no client-side merge engine there)", async () => {
    const office = installHeadlessOffice({
      seedOoxml: MINIMAL_SEED_OOXML,
      platform: "OfficeOnline",
    });
    const fragment = `<?xml version="1.0" encoding="UTF-8"?>
<pkg:package xmlns:pkg="http://schemas.microsoft.com/office/2006/xmlPackage">
  <pkg:part pkg:name="/word/document.xml" pkg:contentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml">
    <pkg:xmlData>
      <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
        <w:body>
          <w:p><w:r><w:t>Via insertOoxml.</w:t></w:r></w:p>
        </w:body>
      </w:document>
    </pkg:xmlData>
  </pkg:part>
</pkg:package>`;

    await getInstalledWord().run(async (context) => {
      context.document.body.insertOoxml(fragment, getInstalledWord().InsertLocation.end);
      await expect(context.sync()).rejects.toThrow();
    });

    expect(office.getOoxml()).not.toContain("Via insertOoxml.");
    office.dispose();
  });

  it("the platform option reaches paraId churn behavior — OfficeOnline regenerates ids on every getOoxml() call", () => {
    const office = installHeadlessOffice({
      seedOoxml: MINIMAL_SEED_OOXML,
      platform: "OfficeOnline",
    });
    const firstRead = office.getOoxml();
    const secondRead = office.getOoxml();
    const paraIdPattern = /w14:paraId="([0-9A-F]{8})"/;
    expect(firstRead.match(paraIdPattern)?.[1]).not.toBe(secondRead.match(paraIdPattern)?.[1]);
    office.dispose();
  });
});
