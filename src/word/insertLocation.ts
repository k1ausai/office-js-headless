// Matches the real Word.InsertLocation enum's string values
// (node_modules/@types/office-js/index.d.ts) exactly.
export const InsertLocation = {
  before: "Before",
  after: "After",
  start: "Start",
  end: "End",
  replace: "Replace",
} as const;

export type InsertLocationValue = (typeof InsertLocation)[keyof typeof InsertLocation];
