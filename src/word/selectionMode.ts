// Matches the real Word.SelectionMode enum's string values exactly.
export const SelectionMode = {
  select: "Select",
  start: "Start",
  end: "End",
} as const;

export type SelectionModeValue = (typeof SelectionMode)[keyof typeof SelectionMode];
