import { type Element } from "@xmldom/xmldom";
import { FlatOpcDocument, type StyleType } from "../document/FlatOpcDocument";
import { TrackedProperties } from "./proxy";
import type { Syncable } from "./run";

type StyleProperty = "id" | "nameLocal" | "type" | "builtIn";

// Read-only for v1 (issue #17) — no write API. No platform/enqueue:
// nothing here mutates the document, so there's nothing to defer, same
// reasoning as Comment (#16).
export class Style implements Syncable {
  private readonly tracked = new TrackedProperties();

  constructor(
    private readonly doc: FlatOpcDocument,
    private readonly target: Element
  ) {}

  load(propertyNames: StyleProperty | StyleProperty[]): void {
    this.tracked.load(propertyNames);
  }

  // A shim-only convenience, NOT a mirrored real property — real
  // Word.Style has no readable id/styleId at all (confirmed against the
  // real API surface; its scalar properties are nameLocal/type/builtIn/
  // priority/... and similar, nothing identifier-shaped). Real add-in code
  // resolves a Word.BuiltInStyleName value via
  // StyleCollection.getByName(...), a lookup, never by reading an id back
  // off an already-obtained Style. Exposed here anyway since it's the raw
  // w:styleId this shim already has on hand, and real Word's built-in
  // w:styleId values happen to equal the BuiltInStyleName enum's strings
  // (e.g. "Heading1") — useful for tests, but don't mistake it for API
  // parity.
  get id(): string {
    return this.tracked.read<string>("id");
  }

  get nameLocal(): string {
    return this.tracked.read<string>("nameLocal");
  }

  get type(): StyleType {
    return this.tracked.read<StyleType>("type");
  }

  get builtIn(): boolean {
    return this.tracked.read<boolean>("builtIn");
  }

  /** Called by RequestContext.sync() after the mutation queue has run. */
  sync(): void {
    this.tracked.sync((name) => this.computeProperty(name));
  }

  private computeProperty(name: string): unknown {
    switch (name) {
      case "id":
        return this.doc.getStyleId(this.target);
      case "nameLocal":
        return this.doc.getStyleNameLocal(this.target);
      case "type":
        return this.doc.getStyleType(this.target);
      case "builtIn":
        return this.doc.getStyleBuiltIn(this.target);
      default:
        throw new Error(`Style: unknown property "${name}"`);
    }
  }
}
