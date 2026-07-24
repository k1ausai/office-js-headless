import { afterEach, describe, expect, it } from "vitest";
import { installHeadlessOffice } from "./install";
import { MINIMAL_SEED_OOXML } from "./document/__fixtures__/minimalSeed";

// The real global `Word`/`Office` types (from @types/office-js) describe the
// full production API surface, far more than this shim implements so far.
// Tests exercise the shim through this minimal, precise slice of that shape
// instead of casting through `any`.
interface MinimalWordContext {
  document: { body: { insertText(text: string, insertLocation: string): void } };
  sync(): Promise<void>;
}
interface MinimalWordGlobal {
  run<T>(callback: (context: MinimalWordContext) => Promise<T>): Promise<T>;
  InsertLocation: { end: string };
}

function getInstalledWord(): MinimalWordGlobal {
  return (globalThis as unknown as { Word: MinimalWordGlobal }).Word;
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
});
