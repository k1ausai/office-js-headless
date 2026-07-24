import { describe, expect, it } from "vitest";
import { COMMENT_SEED_OOXML } from "../document/__fixtures__/commentSeed";
import { FlatOpcDocument } from "../document/FlatOpcDocument";
import { Comment } from "./comment";

function topLevelComments(doc: FlatOpcDocument): Comment[] {
  return doc.getTopLevelCommentElements().map((c) => new Comment(doc, c));
}

describe("Comment load/sync gating", () => {
  it("reading a property without calling .load() first throws PropertyNotLoaded", () => {
    const doc = new FlatOpcDocument(COMMENT_SEED_OOXML);
    const [comment] = topLevelComments(doc);
    expect(() => comment!.authorName).toThrow(/PropertyNotLoaded|not available/);
  });

  it("reading after .load() but before sync() still throws — identical to never-loaded", () => {
    const doc = new FlatOpcDocument(COMMENT_SEED_OOXML);
    const [comment] = topLevelComments(doc);
    comment!.load("authorName");
    expect(() => comment!.authorName).toThrow(/PropertyNotLoaded|not available/);
  });

  it("reading after .load() + sync() returns the current values", () => {
    const doc = new FlatOpcDocument(COMMENT_SEED_OOXML);
    const [comment] = topLevelComments(doc);
    comment!.load(["id", "authorName", "content", "creationDate", "resolved"]);
    comment!.sync();

    expect(comment!.id).toBe("0");
    expect(comment!.authorName).toBe("Jane Doe");
    expect(comment!.content).toBe("Top-level resolved comment.");
    expect(comment!.creationDate).toEqual(new Date("2024-01-15T10:00:00Z"));
    expect(comment!.resolved).toBe(true);
  });
});

describe("Comment.replies", () => {
  it("a comment with a threaded reply exposes it via .replies", () => {
    const doc = new FlatOpcDocument(COMMENT_SEED_OOXML);
    const [comment] = topLevelComments(doc);
    comment!.load("replies");
    comment!.sync();

    expect(comment!.replies).toHaveLength(1);
    expect(comment!.replies[0]).toMatchObject({
      id: "1",
      authorName: "John Smith",
      content: "Reply to the first comment.",
    });
    expect(comment!.replies[0]!.creationDate).toEqual(new Date("2024-01-15T11:00:00Z"));
  });

  it("a comment with no replies exposes an empty array", () => {
    const doc = new FlatOpcDocument(COMMENT_SEED_OOXML);
    const comments = topLevelComments(doc);
    const unresolvedComment = comments[1]!;
    unresolvedComment.load(["replies", "resolved"]);
    unresolvedComment.sync();

    expect(unresolvedComment.replies).toEqual([]);
    expect(unresolvedComment.resolved).toBe(false);
  });
});
