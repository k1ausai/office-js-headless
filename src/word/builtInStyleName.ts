// Matches the real Word.BuiltInStyleName enum's string values — real
// Word's built-in styles use exactly these strings as their w:styleId, so
// `style.id === BuiltInStyleName.heading1` works directly (see style.ts).
//
// A representative subset, not the full ~120-value real enum: the ones a
// driving consumer is actually likely to check against. Style.id/nameLocal
// read correctly for ANY style regardless of whether its name appears
// here — this list is a convenience export, not a validation gate, so
// there's no correctness cost to it being incomplete. Extend on demand,
// per the design spec's non-goal of front-loading unused surface.
export const BuiltInStyleName = {
  other: "Other",
  normal: "Normal",
  heading1: "Heading1",
  heading2: "Heading2",
  heading3: "Heading3",
  heading4: "Heading4",
  heading5: "Heading5",
  heading6: "Heading6",
  heading7: "Heading7",
  heading8: "Heading8",
  heading9: "Heading9",
  title: "Title",
  subtitle: "Subtitle",
  quote: "Quote",
  intenseQuote: "IntenseQuote",
  caption: "Caption",
  listParagraph: "ListParagraph",
  noSpacing: "NoSpacing",
  tableGrid: "TableGrid",
  strong: "Strong",
  emphasis: "Emphasis",
  hyperlink: "Hyperlink",
} as const;

export type BuiltInStyleNameValue = (typeof BuiltInStyleName)[keyof typeof BuiltInStyleName];
