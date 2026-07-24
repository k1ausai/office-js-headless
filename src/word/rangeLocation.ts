// Matches the real Word.RangeLocation enum's string values exactly. Not yet
// exposed on the installed Word global — Range itself isn't reachable from
// context.document until a factory method exists (Paragraph.getRange in
// #14, Range.search in #13), so exposing this now would be dead surface.
export const RangeLocation = {
  whole: "Whole",
  start: "Start",
  end: "End",
  before: "Before",
  after: "After",
  content: "Content",
} as const;

export type RangeLocationValue = (typeof RangeLocation)[keyof typeof RangeLocation];
