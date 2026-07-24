---
title: Unloaded/unsynced-property error message shape
labels: [wayfinder:research]
status: closed
assignee: wayfinder-charting-session
blocked_by: []
---

## Question

Real Office.js throws a specific error (`OfficeExtension.Error`, roughly
`"...call load...call context.sync()..."` wording) when code reads a proxy property that was
never named in `.load()`, or was loaded but not yet followed by a `context.sync()`. What is the
exact error name, `code`, and message text real Office.js produces for each of those two cases
(unloaded-read vs. loaded-but-unsynced-read)? Source from `@types/office-js`/the real Office.js
runtime source, not guesswork — the design spec says to match real wording "where practical."

## Answer

**Bottom line: real Office.js does NOT distinguish "never loaded" from "loaded but not yet
synced." Both cases throw the exact same error — same `code`, same message template, same
class — because the client-side check only asks "is the backing field populated?", not "why
isn't it?"**

### The one relevant error: `PropertyNotLoaded`

- `code`: `"PropertyNotLoaded"`
- `message` (template, `{0}` = the property name being read): `"The property '{0}' is not
available. Before reading the property's value, call the load method on the containing object
and call \"context.sync()\" on the associated request context."`
- Thrown class: `OfficeExtension.Error` (aliased at runtime to an internal `RuntimeError` class —
  see the naming caveat below).

Evidence, from the actual shipped runtime, not just typings:

- Package `@microsoft/office-js@1.1.110` (npm; fetched via
  `https://registry.npmjs.org/@microsoft/office-js/-/office-js-1.1.110.tgz`),
  `dist/word-win32-16.01.debug.js` (Word desktop RichApi runtime — matches this repo's stated
  "desktop Word fidelity" target).
  - Every generated property getter funnels through one gate function:
    `Utility.throwIfNotLoaded(propertyName, fieldValue, entityName, isNull)` — throws iff
    `fieldValue === undefined` (and the property isn't a private `_`-prefixed field and isn't
    suppressed). Concretely, e.g. `Body.prototype.text`'s getter is:
    ```js
    get: function () {
        _throwIfNotLoaded("text", this._Te, _typeBody, this._isNull);
        return this._Te;
    }
    ```
    `this._Te` is `undefined` in _both_ of the ticket's cases: (a) `.load("text")` was never
    queued, and (b) `.load("text")` was queued but `context.sync()` hasn't resolved yet to
    populate the field from the server response. There is no separate code path, flag, or
    branch that tells these two situations apart — the same `throwIfNotLoaded` call, same
    message, same code, fires either way.
  - The message template and code are registered together:
    ```js
    ErrorCodes.propertyNotLoaded = 'PropertyNotLoaded';
    // ...
    PropertyNotLoaded: "The property '{0}' is not available. Before reading the property's
    value, call the load method on the containing object and call \"context.sync()\" on the
    associated request context."
    ```
    (`Utility.createPropertyNotLoadedException` builds the thrown object from these two pieces.)
  - Confirmed identical in `dist/excel-win32-16.01.debug.js` and
    `dist/powerpoint-win32-16.01.debug.js` from the same package — this is shared RichApi
    runtime plumbing, not Word-specific.

- Secondary corroboration (can't inspect exact runtime source, but consistent with the above):
  Microsoft Q&A thread ["PropertyNotLoaded Error Persisting after calling load and
  context.sync"](https://learn.microsoft.com/en-us/answers/questions/946867/propertynotloaded-error-persisting-after-calling-l)
  and GitHub issue
  [OfficeDev/office-js#777](https://github.com/OfficeDev/office-js/issues/777) — both describe
  `PropertyNotLoaded` firing for both mis-ordered `load`/`sync` calls, matching the "single
  undefined-field check" behavior found in the runtime source above.

### Two other error codes exist, but they are NOT this scenario

The design spec's loose wording ("...call load...call context.sync()...") maps onto exactly one
error. Two other codes are easy to confuse with it but are semantically different and should not
be reused for a bare unloaded/unsynced property read:

- **`InvalidObjectPath`** — fires when a client object reference goes stale because it was used
  across multiple `context.sync()` calls _outside_ a single `.run()` batch without being
  registered via `context.trackedObjects.add()`. Message (same source file, line ~11494):
  `"The object path '{0}' isn't working for what you're trying to do. If you're using the object
across multiple \"context.sync\" calls and outside the sequential execution of a \".run\"
batch, please use the \"context.trackedObjects.add()\" and \"context.trackedObjects.remove()\"
methods to manage the object's lifetime."` This is about object lifetime/tracking, not about
  reading an unloaded property.
- **`ValueNotLoaded`** — used specifically for `FunctionResult.value`-style "result objects"
  (e.g. custom-function results), not regular proxy properties. Message: `"The value of the
result object has not been loaded yet. Before reading the value property, call
\"context.sync()\" on the associated request context."` Doesn't apply to Body/Range/Paragraph/
  Table proxy properties, which use `PropertyNotLoaded` as described above.

Both are declared alongside `propertyNotLoaded` on `OfficeExtension.CoreErrorCodes` /
`OfficeExtension.ErrorCodes` in `@types/office-js@1.0.600`
(`node_modules/@types/office-js/index.d.ts`, class `OfficeExtension.ErrorCodes` around line
25854, and the `OfficeExtension.Error` class doc comment at line ~25826) — but the type
declarations only give the _names_ of the static members typed as `string`; they don't carry the
literal string values or message text. Those had to come from the runtime bundle above.

There's also a numeric "Office Common API error codes" table (Microsoft Learn:
https://learn.microsoft.com/en-us/office/dev/add-ins/reference/javascript-api-for-office-error-codes)
with codes like 1000–13nnn — that page explicitly states it does **not** apply to
application-specific APIs like Word/Excel JS (it's the older Common API / binding model). It has
no bearing on this question and should not be cited as a source for `Word.*` proxy errors.

### Naming caveat worth flagging (docs vs. runtime mismatch)

`@types/office-js`'s `OfficeExtension.Error` class carries the comment `/** Error name:
"OfficeExtension.Error".*/` (`node_modules/@types/office-js/index.d.ts:25828`), and Microsoft
Learn's `OfficeExtension.Error` reference page repeats this. However, in the actual runtime
source (`@microsoft/office-js@1.1.110`, `dist/word-win32-16.01.debug.js`,
`dist/excel-win32-16.01.debug.js`, `dist/powerpoint-win32-16.01.debug.js`), `OfficeExtension.Error`
is a straight alias for an internal class (`OfficeExtension_1.Error = _Internal.RuntimeError;`)
whose constructor sets:

```js
_this.name = "RichApi.Error";
```

So the `.name` property an add-in actually observes at runtime is the string `"RichApi.Error"`,
not `"OfficeExtension.Error"`, across every host application checked (Word, Excel, PowerPoint).
This is a genuine, confirmed discrepancy between the public docs/types and the shipped runtime,
not a guess — it's worth deciding deliberately whether office-js-headless's thrown error's
`.name` should match the documented `"OfficeExtension.Error"` (matches the public contract
add-in authors read) or the observed `"RichApi.Error"` (matches what real Word desktop actually
sends). Given the design spec's stated goal is desktop Word fidelity, `"RichApi.Error"` is the
more faithful choice, with `"OfficeExtension.Error"` as the documented-but-unobserved
alternative. **This decision is broken out into its own ticket — see
[Error `.name` value: "RichApi.Error" vs "OfficeExtension.Error"](007-error-name-value.md).**

### Recommended shape for office-js-headless

For a proxy property that's unloaded or loaded-but-unsynced (single unified case, per above):

```ts
{
  name: "RichApi.Error", // or "OfficeExtension.Error" — see ticket 007 for the decision
  code: "PropertyNotLoaded",
  message: `The property '${propertyName}' is not available. Before reading the property's value, call the load method on the containing object and call "context.sync()" on the associated request context.`,
}
```

No second, distinct error shape is needed for "never loaded" vs. "loaded but unsynced" — real
Office.js collapses both into this one.

### Sources

- `@types/office-js@1.0.600` (installed in this repo's `node_modules`) —
  `node_modules/@types/office-js/index.d.ts`, `OfficeExtension.Error` class (~line 25826),
  `OfficeExtension.ErrorCodes` class (~line 25854), `ClientRequestContext.load`/`loadRecursive`
  docs (~line 25705).
- `@microsoft/office-js@1.1.110` (npm registry:
  https://registry.npmjs.org/@microsoft/office-js/-/office-js-1.1.110.tgz) —
  `dist/word-win32-16.01.debug.js`, `dist/excel-win32-16.01.debug.js`,
  `dist/powerpoint-win32-16.01.debug.js`: `ErrorCodes`/`CoreErrorCodes` definitions,
  `ResourceStrings`/`CoreResourceStrings` message templates, `Utility.throwIfNotLoaded` /
  `Utility.createPropertyNotLoadedException`, `_Internal.RuntimeError` constructor (the
  `OfficeExtension.Error` alias and `this.name = 'RichApi.Error'` line), and a concrete
  `Body.prototype.text` getter showing the gate in use.
- Microsoft Learn, `OfficeExtension.ErrorCodes` class reference:
  https://learn.microsoft.com/en-us/javascript/api/office/officeextension.errorcodes?view=word-js-preview
  (lists the same member names; no literal values or message text given there either).
- Microsoft Learn, "Office Common API error codes":
  https://learn.microsoft.com/en-us/office/dev/add-ins/reference/javascript-api-for-office-error-codes
  (checked and ruled out — explicitly a different, older API model, not applicable here).
- Microsoft Q&A: https://learn.microsoft.com/en-us/answers/questions/946867/propertynotloaded-error-persisting-after-calling-l
- GitHub, OfficeDev/office-js#777: https://github.com/OfficeDev/office-js/issues/777
  (both used only as secondary corroboration of observed behavior, not as primary source for
  exact code/message text — that came from the runtime bundle).
