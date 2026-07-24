---
title: search() semantics scope
labels: [wayfinder:research]
status: closed
assignee: wayfinder-charting-session
blocked_by: []
---

## Question

`Range.search()` in real Office.js supports both plain-text matching and a wildcard syntax
(`SearchOptions`: matchCase, matchWholeWord, matchWildcards, etc). The design spec scopes v1 to
"what the driving consumer actually uses" rather than the full surface. What does the driving
consumer's Office.js-facing service layer actually call — plain-text search only, or does it pass
any `SearchOptions` (wildcards, match-case, match-whole-word)? Grep the consumer's codebase for
`.search(` call sites and their options.

Keep the answer generic in this ticket (call shape/options used, not the consumer's file paths or
repo name) — this map is public.

## Answer

The driving consumer's Office.js-facing service layer has exactly one `Range.search()` call site
(no other `Body.search()`/`Range.search()` usage exists anywhere in that codebase). It:

- Passes a **plain string**, not a wildcard pattern. The call site is guarded by a
  try/catch specifically because `search()` throws on wildcard-special characters in the input —
  that's treated as an expected failure mode to fall back from, not a feature the consumer relies
  on. `matchWildcards` is never set to `true` anywhere.
- Passes a `SearchOptions` object with a single option set: `{ matchCase: true }`.
  `matchWholeWord`, `matchWildcards`, `ignorePunct`, and `ignoreSpace` are never referenced at all
  (so they're left at their Office.js defaults, all `false`/off).

**Conclusion:** v1 only needs to support plain-text search with an optional `matchCase` flag. The
wildcard syntax, `matchWholeWord`, `ignorePunct`, and `ignoreSpace` are unused by the driving
consumer and can stay out of scope for v1's `search()` implementation.
