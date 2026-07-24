import { FlatOpcDocument } from "./document/FlatOpcDocument";
import { InsertLocation } from "./word/insertLocation";
import { wordRun } from "./word/run";

export interface InstallHeadlessOfficeOptions {
  seedOoxml: string;
}

export interface HeadlessOfficeHandle {
  getOoxml(): string;
  reset(): void;
  dispose(): void;
}

export function installHeadlessOffice(options: InstallHeadlessOfficeOptions): HeadlessOfficeHandle {
  const doc = new FlatOpcDocument(options.seedOoxml);

  const wordGlobal = {
    run: <T>(callback: (context: unknown) => Promise<T>) => wordRun(doc, callback as never),
    InsertLocation,
  };

  // Real add-in code references bare global `Word`/`Office` — this is the
  // only way to run it unmodified, so installation mutates globals, same as
  // the real host injects them. `Office` is an empty placeholder for now;
  // `Office.context`/`onReady` arrive with platform selection.
  (globalThis as Record<string, unknown>).Word = wordGlobal;
  (globalThis as Record<string, unknown>).Office = {};

  return {
    getOoxml: () => doc.getOoxml(),
    reset: () => doc.reset(),
    dispose: () => {
      delete (globalThis as Record<string, unknown>).Word;
      delete (globalThis as Record<string, unknown>).Office;
    },
  };
}
