import { type Element } from "@xmldom/xmldom";
import { FlatOpcDocument } from "../document/FlatOpcDocument";
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

  // The stable OOXML identifier (w:styleId). Real Word's built-in styles
  // use exactly the Word.BuiltInStyleName enum's string values as their
  // styleId (e.g. "Heading1") — comparing `style.id` against
  // Word.BuiltInStyleName.heading1 works without any separate lookup.
  get id(): string {
    return this.tracked.read<string>("id");
  }

  get nameLocal(): string {
    return this.tracked.read<string>("nameLocal");
  }

  get type(): "Paragraph" | "Character" | "Table" | "List" {
    return this.tracked.read<"Paragraph" | "Character" | "Table" | "List">("type");
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
