import { describe, expect, it } from "vitest";
import { generateHexId } from "./paraId";

describe("generateHexId", () => {
  it("produces an 8-hex-char uppercase string", () => {
    for (let i = 0; i < 20; i++) {
      const id = generateHexId();
      expect(id).toMatch(/^[0-9A-F]{8}$/);
    }
  });

  it("produces different values across calls (not a fixed constant)", () => {
    const ids = new Set(Array.from({ length: 20 }, () => generateHexId()));
    expect(ids.size).toBeGreaterThan(1);
  });
});
