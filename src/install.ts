import { FlatOpcDocument } from "./document/FlatOpcDocument";
import { InsertLocation } from "./word/insertLocation";
import { RequestContext, wordRun } from "./word/run";

export interface InstallHeadlessOfficeOptions {
  seedOoxml: string;
}

export interface HeadlessOfficeHandle {
  getOoxml(): string;
  reset(): void;
  dispose(): void;
}

function globalRecord(): Record<string, unknown> {
  return globalThis as Record<string, unknown>;
}

export function installHeadlessOffice(options: InstallHeadlessOfficeOptions): HeadlessOfficeHandle {
  const doc = new FlatOpcDocument(options.seedOoxml);

  const wordGlobal = {
    run: <T>(callback: (context: RequestContext) => Promise<T>) => wordRun(doc, callback),
    InsertLocation,
  };

  // Real add-in code references bare global `Word`/`Office` — this is the
  // only way to run it unmodified, so installation mutates globals, same as
  // the real host injects them. `Office` is an empty placeholder for now;
  // `Office.context`/`onReady` arrive with platform selection.
  globalRecord().Word = wordGlobal;
  globalRecord().Office = {};

  return {
    getOoxml: () => doc.getOoxml(),
    reset: () => doc.reset(),
    dispose: () => {
      delete globalRecord().Word;
      delete globalRecord().Office;
    },
  };
}
