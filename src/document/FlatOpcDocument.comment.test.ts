import { describe, expect, it } from "vitest";
import { COMMENT_SEED_OOXML } from "./__fixtures__/commentSeed";
import { FlatOpcDocument } from "./FlatOpcDocument";
import { MINIMAL_SEED_OOXML } from "./__fixtures__/minimalSeed";

describe("FlatOpcDocument comment primitives", () => {
  it("getCommentElements returns every w:comment, in document order", () => {
    const doc = new FlatOpcDocument(COMMENT_SEED_OOXML);
    const comments = doc.getCommentElements();
    expect(comments.map((c) => doc.getCommentId(c))).toEqual(["0", "1", "2"]);
  });

  it("getTopLevelCommentElements excludes threaded replies", () => {
    const doc = new FlatOpcDocument(COMMENT_SEED_OOXML);
    const topLevel = doc.getTopLevelCommentElements();
    expect(topLevel.map((c) => doc.getCommentId(c))).toEqual(["0", "2"]);
  });

  it("getCommentAuthor/getCommentContent/getCommentDate read the comment's fields", () => {
    const doc = new FlatOpcDocument(COMMENT_SEED_OOXML);
    const [comment0] = doc.getCommentElements();
    expect(doc.getCommentAuthor(comment0!)).toBe("Jane Doe");
    expect(doc.getCommentContent(comment0!)).toBe("Top-level resolved comment.");
    expect(doc.getCommentDate(comment0!)).toEqual(new Date("2024-01-15T10:00:00Z"));
  });

  it("getCommentResolved reads the w15:done flag from commentsExtended.xml", () => {
    const doc = new FlatOpcDocument(COMMENT_SEED_OOXML);
    const [comment0, , comment2] = doc.getCommentElements();
    expect(doc.getCommentResolved(comment0!)).toBe(true);
    expect(doc.getCommentResolved(comment2!)).toBe(false);
  });

  it("getCommentReplyElements returns the replies belonging to a specific comment", () => {
    const doc = new FlatOpcDocument(COMMENT_SEED_OOXML);
    const [comment0, , comment2] = doc.getCommentElements();
    const repliesTo0 = doc.getCommentReplyElements(comment0!);
    expect(repliesTo0.map((c) => doc.getCommentId(c))).toEqual(["1"]);
    expect(doc.getCommentReplyElements(comment2!)).toEqual([]);
  });

  it("a reply is itself readable through the same primitives as a top-level comment", () => {
    const doc = new FlatOpcDocument(COMMENT_SEED_OOXML);
    const [comment0] = doc.getCommentElements();
    const [reply] = doc.getCommentReplyElements(comment0!);
    expect(doc.getCommentAuthor(reply!)).toBe("John Smith");
    expect(doc.getCommentContent(reply!)).toBe("Reply to the first comment.");
  });

  it("returns no comments when the document has no comments.xml part", () => {
    const doc = new FlatOpcDocument(MINIMAL_SEED_OOXML);
    expect(doc.getCommentElements()).toEqual([]);
    expect(doc.getTopLevelCommentElements()).toEqual([]);
  });
});
