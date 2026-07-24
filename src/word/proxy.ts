// Real Office.js does not distinguish "property never named in .load()" from
// "loaded but not yet followed by context.sync()" — both throw the identical
// error, traced from the shipped runtime (@microsoft/office-js), not just its
// published types (which document a different, unobserved error name — see
// doc/wayfinder/tickets/001-error-message-shape.md and 007-error-name-value.md).
export function propertyNotLoadedError(propertyName: string): Error {
  const error = new Error(
    `The property '${propertyName}' is not available. Before reading the property's value, call the load method on the containing object and call "context.sync()" on the associated request context.`
  );
  error.name = "RichApi.Error";
  (error as Error & { code: string }).code = "PropertyNotLoaded";
  return error;
}

// Shared by every shim proxy object (Body, Range, Paragraph, Table, ...):
// tracks which properties were named in .load(...), and only lets reads
// through once RequestContext.sync() has snapshotted their values — matching
// real Office.js, where a loaded property's value is a snapshot taken at
// sync() time, not a live view (a mutation after sync() without a fresh
// load()+sync() does not change what a later read returns).
export class TrackedProperties {
  private readonly loaded = new Set<string>();
  private readonly values = new Map<string, unknown>();

  load(propertyNames: string | string[]): void {
    for (const name of Array.isArray(propertyNames) ? propertyNames : [propertyNames]) {
      this.loaded.add(name);
    }
  }

  sync(computeValue: (propertyName: string) => unknown): void {
    for (const name of this.loaded) {
      this.values.set(name, computeValue(name));
    }
  }

  read<T>(propertyName: string): T {
    if (!this.values.has(propertyName)) {
      throw propertyNotLoadedError(propertyName);
    }
    return this.values.get(propertyName) as T;
  }
}
