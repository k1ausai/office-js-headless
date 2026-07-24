---
title: Error `.name` value — "RichApi.Error" vs "OfficeExtension.Error"
labels: [wayfinder:grilling]
status: closed
assignee: wayfinder-session
blocked_by: []
---

## Question

Surfaced by [Unloaded/unsynced-property error message shape](001-error-message-shape.md): real
Office.js's public types/docs (`@types/office-js`, Microsoft Learn) document the thrown error's
class/name as `"OfficeExtension.Error"`, but the actual shipped runtime
(`@microsoft/office-js@1.1.110`, checked across Word/Excel/PowerPoint desktop bundles) sets
`error.name = "RichApi.Error"` at construction — a confirmed docs-vs-runtime mismatch, not a
guess.

Decide: should `office-js-headless`'s thrown `PropertyNotLoaded` error (and any other thrown
errors modeled the same way) have `.name === "RichApi.Error"` (matches what real Word desktop
actually sends — consistent with the design spec's stated desktop-fidelity target) or
`.name === "OfficeExtension.Error"` (matches the documented public contract add-in authors read,
even though it's not what they'd observe against real Word)?

## Answer

**Decision: `"RichApi.Error"`** — matches what real Word desktop actually sends, not the
documented-but-unobserved `"OfficeExtension.Error"`.

Rationale: the design spec's core principle is that the shim exists to catch real add-in bugs by
matching real Word's actual observable behavior (this is the same reasoning behind strict
`.load()`/`.sync()` gating and the platform-conditional capability signals). Consumer code that
branches on `error.name === "OfficeExtension.Error"` would already be silently broken against
real Word desktop today, since real Word never sends that string — reproducing the documented
value instead of the observed one would hide that bug rather than catch it. Ticket 001's finding
is treated as authoritative (confirmed from shipped runtime source, not docs).

No configurable option — kept as a fixed behavior, consistent with "extend on demand, not
speculatively" (no consumer need identified for the documented-but-fictional alternative).
