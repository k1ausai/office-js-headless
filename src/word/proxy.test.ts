import { describe, expect, it } from "vitest";
import { propertyNotLoadedError, TrackedProperties } from "./proxy";

describe("propertyNotLoadedError", () => {
  it("matches real Office.js's confirmed PropertyNotLoaded error shape", () => {
    const error = propertyNotLoadedError("text");
    expect(error.name).toBe("RichApi.Error");
    expect((error as Error & { code: string }).code).toBe("PropertyNotLoaded");
    expect(error.message).toBe(
      "The property 'text' is not available. Before reading the property's value, call the load method on the containing object and call \"context.sync()\" on the associated request context."
    );
  });
});

describe("TrackedProperties", () => {
  it("throws reading a property that was never loaded", () => {
    const tracked = new TrackedProperties();
    expect(() => tracked.read("text")).toThrow(/PropertyNotLoaded|not available/);
  });

  it("throws reading a property that was loaded but not yet synced — identical error to never-loaded", () => {
    const tracked = new TrackedProperties();
    tracked.load("text");

    let neverLoadedError: unknown;
    let loadedNotSyncedError: unknown;
    try {
      new TrackedProperties().read("text");
    } catch (err) {
      neverLoadedError = err;
    }
    try {
      tracked.read("text");
    } catch (err) {
      loadedNotSyncedError = err;
    }

    expect(loadedNotSyncedError).toBeInstanceOf(Error);
    expect((loadedNotSyncedError as Error).name).toBe((neverLoadedError as Error).name);
    expect((loadedNotSyncedError as Error).message).toBe((neverLoadedError as Error).message);
    expect((loadedNotSyncedError as Error & { code: string }).code).toBe(
      (neverLoadedError as Error & { code: string }).code
    );
  });

  it("returns the value once loaded and synced", () => {
    const tracked = new TrackedProperties();
    tracked.load("text");
    tracked.sync((name) => (name === "text" ? "hello" : undefined));
    expect(tracked.read("text")).toBe("hello");
  });

  it("accepts an array of property names to load() at once", () => {
    const tracked = new TrackedProperties();
    tracked.load(["text", "style"]);
    tracked.sync((name) => `value-of-${name}`);
    expect(tracked.read("text")).toBe("value-of-text");
    expect(tracked.read("style")).toBe("value-of-style");
  });

  it("sync() only snapshots loaded properties, not arbitrary ones", () => {
    const tracked = new TrackedProperties();
    tracked.load("text");
    tracked.sync((name) => `value-of-${name}`);
    expect(() => tracked.read("style")).toThrow();
  });

  it("a second sync() re-snapshots values (reflecting a mutation between syncs)", () => {
    const tracked = new TrackedProperties();
    tracked.load("text");
    tracked.sync(() => "first");
    expect(tracked.read("text")).toBe("first");

    tracked.sync(() => "second");
    expect(tracked.read("text")).toBe("second");
  });
});
