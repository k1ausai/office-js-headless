import { type Element } from "@xmldom/xmldom";
import { FlatOpcDocument } from "../document/FlatOpcDocument";
import { TrackedProperties } from "./proxy";
import type { Syncable } from "./run";

type CommentProperty = "id" | "authorName" | "content" | "creationDate" | "resolved" | "replies";

// A reply is only ever obtainable through an already load/sync-gated
// Comment.replies read — by the time a caller has one, its data was already
// computed as part of the parent Comment's sync() pass. Gating each reply's
// own properties again would add TrackedProperties/Syncable machinery with
// no real protective purpose (unlike Comment itself, whose properties are
// readable — if ungated — the instant it's constructed, before any
// load()+sync() has happened). Plain data, not a further proxy.
export interface CommentReply {
  readonly id: string;
  readonly authorName: string;
  readonly content: string;
  readonly creationDate: Date;
}

// Read-only for v1 (issue #16) — no write API for content/resolved, unlike
// real Word.Comment. No `platform`/`enqueue`: nothing here mutates the
// document, so there's nothing to defer.
export class Comment implements Syncable {
  private readonly tracked = new TrackedProperties();

  constructor(
    private readonly doc: FlatOpcDocument,
    private readonly target: Element
  ) {}

  load(propertyNames: CommentProperty | CommentProperty[]): void {
    this.tracked.load(propertyNames);
  }

  get id(): string {
    return this.tracked.read<string>("id");
  }

  get authorName(): string {
    return this.tracked.read<string>("authorName");
  }

  get content(): string {
    return this.tracked.read<string>("content");
  }

  get creationDate(): Date {
    return this.tracked.read<Date>("creationDate");
  }

  get resolved(): boolean {
    return this.tracked.read<boolean>("resolved");
  }

  get replies(): CommentReply[] {
    return this.tracked.read<CommentReply[]>("replies");
  }

  /** Called by RequestContext.sync() after the mutation queue has run. */
  sync(): void {
    this.tracked.sync((name) => this.computeProperty(name));
  }

  private computeProperty(name: string): unknown {
    switch (name) {
      case "id":
        return this.doc.getCommentId(this.target);
      case "authorName":
        return this.doc.getCommentAuthor(this.target);
      case "content":
        return this.doc.getCommentContent(this.target);
      case "creationDate":
        return this.doc.getCommentDate(this.target);
      case "resolved":
        return this.doc.getCommentResolved(this.target);
      case "replies":
        return this.doc.getCommentReplyElements(this.target).map((reply): CommentReply => ({
          id: this.doc.getCommentId(reply),
          authorName: this.doc.getCommentAuthor(reply),
          content: this.doc.getCommentContent(reply),
          creationDate: this.doc.getCommentDate(reply),
        }));
      default:
        throw new Error(`Comment: unknown property "${name}"`);
    }
  }
}
