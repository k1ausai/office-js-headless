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
//
// A discriminated union on expectRejection, rather than two independently-
// optional fields, makes "resultOoxml is required unless this fixture
// expects sync() to reject" a compile-time guarantee — a fixture author
// can't accidentally omit both.
interface FixtureBase {
  description: string;
  /** Defaults to "PC". */
  platform?: SupportedPlatform;
  seedOoxml: string;
  apply: (context: RequestContext) => void | Promise<void>;
}

export type Fixture =
  | (FixtureBase & {
      /** The expected document state after apply()+sync(). */
      resultOoxml: string;
      expectRejection?: undefined;
    })
  | (FixtureBase & {
      resultOoxml?: undefined;
      /**
       * This fixture's apply() is expected to make sync() reject (e.g.
       * insertOoxml on OfficeOnline). A RegExp narrows the expected
       * rejection message; `true` accepts any rejection.
       */
      expectRejection: boolean | RegExp;
    });
