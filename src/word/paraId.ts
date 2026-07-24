import { type Element } from "@xmldom/xmldom";
import { SupportedPlatform } from "../office/context";

const W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const W14_NS = "http://schemas.microsoft.com/office/word/2010/wordml";

// Real Word uses 8-hex-char uppercase ids for w14:paraId/w14:textId/w:rsid*
// (ST_LongHexNumber, ECMA-376 §17.15.1.70) — see
// doc/wayfinder/tickets/005-id-format.md.
export function generateHexId(): string {
  let hex = "";
  for (let i = 0; i < 8; i++) {
    hex += Math.floor(Math.random() * 16).toString(16);
  }
  return hex.toUpperCase();
}

// Every newly-created paragraph gets both a fresh paraId and a fresh textId —
// real Word always stamps both together on new content. w:rsid* is
// deliberately NOT stamped here: ticket 005 confirms rsid-sharing per
// getOoxml() call for OfficeOnline specifically, but explicitly leaves PC/Mac
// per-mutation rsid-sharing unconfirmed ("should be treated as best-effort,
// not a hard-coded assumption") — omitting it entirely is the more honest
// choice than guessing a pattern with no primary evidence behind it.
export function assignFreshIds(paragraph: Element): void {
  paragraph.setAttributeNS(W14_NS, "w14:paraId", generateHexId());
  paragraph.setAttributeNS(W14_NS, "w14:textId", generateHexId());
}

// Insert-after mark-shift: the existing paragraph's OLD paraId moves onto the
// newly inserted paragraph, and the existing paragraph gets a freshly
// generated one — a permanent reassignment, not simple invalidation. Confirmed
// on PC only (Mac assumed to match, no contrary evidence); not applied on
// OfficeOnline, where every id churns unconditionally on the next getOoxml()
// call regardless, making mark-shift unobservable there. See design spec's
// "ParaId stability model" and doc/wayfinder/tickets/006-insert-before-midsplit-confirm.md.
// Only paraId shifts — the new paragraph's own freshly-generated textId (from
// assignFreshIds, already called before this runs) is left as-is; the
// confirmed evidence only describes paraId moving, not textId, and multi-
// paragraph OOXML inserts (paragraphs 2+ relative to the anchor) are an
// unconfirmed extrapolation beyond the single-paragraph case ticket 006
// actually tested — callers only apply this to the paragraph immediately
// adjacent to the anchor, never to the rest of a multi-paragraph fragment.
export function applyInsertAfterMarkShift(
  platform: SupportedPlatform,
  existingParagraph: Element,
  newParagraph: Element
): void {
  if (platform === "OfficeOnline") return;
  const oldId = existingParagraph.getAttributeNS(W14_NS, "paraId");
  if (oldId) {
    newParagraph.setAttributeNS(W14_NS, "w14:paraId", oldId);
  }
  assignFreshIds(existingParagraph);
}

// OfficeOnline: every getOoxml() call regenerates paraId/textId independently
// per paragraph, plus one rsid value SHARED across every paragraph in the
// document for that call (not per-paragraph) — confirmed empirically, see
// doc/wayfinder/tickets/005-id-format.md.
export function regenerateAllIdsForOfficeOnline(paragraphs: Element[]): void {
  const sharedRsid = generateHexId();
  for (const p of paragraphs) {
    p.setAttributeNS(W14_NS, "w14:paraId", generateHexId());
    p.setAttributeNS(W14_NS, "w14:textId", generateHexId());
    p.setAttributeNS(W_NS, "w:rsidR", sharedRsid);
    p.setAttributeNS(W_NS, "w:rsidRDefault", sharedRsid);
    p.setAttributeNS(W_NS, "w:rsidP", sharedRsid);
  }
}
