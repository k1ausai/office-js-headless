import type { SupportedPlatform } from "../src/office/context";
import type { RequestContext } from "../src/word/run";

// Design spec's "Fidelity validation": a golden fixture is a
// {seedOoxml, operation+args, resultOoxml} triple. `apply` IS the
// "operation+args" — expressed as code against the shim's real RequestContext
// type (not a generic descriptor needing a resolver this repo would have to
// build and maintain), so a fixture referencing an unimplemented method is a
// compile error, not a silent runtime no-op. `resultOoxml` is a full OOXML
// document (hand-written for these synthetic fixtures; a real capture from
// #21 would have the exact same shape and drop into this same harness
// unchanged).
export interface Fixture {
  description: string;
  /** Defaults to "PC". */
  platform?: SupportedPlatform;
  seedOoxml: string;
  apply: (context: RequestContext) => void | Promise<void>;
  /**
   * The expected document state after apply()+sync(). Omit when
   * expectRejection is set — a rejected sync() never produces a result to
   * compare.
   */
  resultOoxml?: string;
  /**
   * Set when this fixture's apply() is expected to make sync() reject
   * (e.g. insertOoxml on OfficeOnline). A RegExp narrows the expected
   * rejection message; `true` accepts any rejection.
   */
  expectRejection?: boolean | RegExp;
}
