import { FlatOpcDocument } from "./document/FlatOpcDocument";
import { createOfficeGlobal, SupportedPlatform } from "./office/context";
import { BuiltInStyleName } from "./word/builtInStyleName";
import { InsertLocation } from "./word/insertLocation";
import { RequestContext, wordRun } from "./word/run";

export interface InstallHeadlessOfficeOptions {
  seedOoxml: string;
  /** "PC" | "Mac" | "OfficeOnline" — defaults to "PC". */
  platform?: SupportedPlatform;
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
  const platform = options.platform ?? "PC";
  const doc = new FlatOpcDocument(options.seedOoxml, platform);

  const wordGlobal = {
    run: <T>(callback: (context: RequestContext) => Promise<T>) => wordRun(doc, platform, callback),
    InsertLocation,
    BuiltInStyleName,
  };

  // Real add-in code references bare global `Word`/`Office` — this is the
  // only way to run it unmodified, so installation mutates globals, same as
  // the real host injects them.
  globalRecord().Word = wordGlobal;
  globalRecord().Office = createOfficeGlobal(platform);

  return {
    getOoxml: () => doc.getOoxml(),
    reset: () => doc.reset(),
    dispose: () => {
      delete globalRecord().Word;
      delete globalRecord().Office;
    },
  };
}
