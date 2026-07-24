---
title: Table row/cell OOXML shape
labels: [wayfinder:research]
status: closed
assignee: wayfinder-charting-session
blocked_by: []
---

## Question

What OOXML structure (`w:tbl`/`w:tblPr`/`w:tblGrid`/`w:tr`/`w:tc`/`w:tcPr`) does the shim's
`Table`/`TableRow`/`TableCell` proxy objects need to read and write, per ECMA-376, scoped to the
properties the driving consumer's code actually touches (row/cell insert, cell text,
row/column count — not full table styling unless used)? Produce the minimal element/attribute set
needed for `word/table.ts` and the corresponding `operations/` entries.

## Answer

**Sources.** ECMA-376/ISO-IEC 29500-1 §17.4 "Tables" (via Microsoft's Open XML SDK reference docs,
which reproduce the ISO/IEC 29500-1 1st Edition normative text verbatim under each class's
"Remarks" section, subclause-cited below) — no repo fixtures exist yet to cross-check against
(see ticket 004; v1 golden corpus not yet captured), so this is spec/docs-only, as the ticket
allows.

- `learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.wordprocessing.table` (`w:tbl`, §17.4.38)
- `.../tableproperties` (`w:tblPr`, §17.4.60)
- `.../tablegrid` (`w:tblGrid`, §17.4.49)
- `.../gridcolumn` (`w:gridCol`, §17.4.16)
- `.../tablerow` (`w:tr`, §17.4.79)
- `.../tablerowproperties` (`w:trPr`, §17.4.82)
- `.../tablecell` (`w:tc`, §17.4.66)
- `.../tablecellproperties` (`w:tcPr`, §17.4.70)
- `.../tablecellwidth` (`w:tcW`, §17.4.72; shares content model `w:tblW` §17.4.64 via `ST_TblWidth`)
- `.../gridspan` (`w:gridSpan`, §17.4.17)
- `.../verticalmerge` (`w:vMerge`, §17.4.85)

### Structure, top to bottom

```
w:tbl
├─ w:tblPr?                    (table-wide properties — optional per schema)
├─ w:tblGrid                   (shared vertical edges / column widths)
│  └─ w:gridCol*                 one per column, @w:w = width in twips (dxa)
└─ w:tr*                       (rows, document order = row order)
   ├─ w:trPr?                    (row-wide properties — optional)
   └─ w:tc*                      (cells, document order = column order within the row)
      ├─ w:tcPr?                   (cell properties — optional per schema, but see divergence below)
      │  ├─ w:tcW?                   preferred cell width: @w:w (numeric), @w:type ∈ {auto,dxa,nil,pct}
      │  ├─ w:gridSpan?              @w:val = int, grid columns this cell spans (omitted ⇒ 1)
      │  └─ w:vMerge?                @w:val ∈ {restart,continue}; element present with no @val ⇒ continue
      └─ (w:p | w:tbl)+             ≥1 block-level child, mandatory (§17.4.66: a `tc` with none is
                                     "corrupt"). Text nests exactly like body paragraphs:
                                     w:p > w:r > w:t (§17.3.1.22/§17.3.2.25/§17.3.3.31 — same
                                     elements `word/paragraph.ts` already emits).
```

### Table-level: `w:tblPr` / `w:tblGrid` (§17.4.60 / §17.4.49)

- `w:tblPr` is schema-optional and has ~18 possible children (`tblStyle`, `tblW`, `tblBorders`,
  `tblLook`, `jc`, `bidiVisual`, etc. — full list at §17.4.60). None are load-bearing for
  row/cell insert, cell-text read/write, or row/column count. **Scope: emit `w:tblPr` only if/when
  a consumer touches table-wide formatting (not in v1 API coverage per the design spec's
  "Tables" line item — `getComments`/`getStyles`/`Word.BuiltInStyleName` are the only
  table-adjacent surface listed) — otherwise omit the element entirely, which is schema-valid.**
- `w:tblGrid` is nominally optional too — §17.4.49: "If the table grid is omitted, then a new grid
  shall be constructed from the actual contents of the table assuming that all grid columns have a
  width of 0." **In practice, treat as required to emit**: it's the only place column widths and
  logical column _count_ live independent of any one row's cell count (a row can have fewer `tc`
  than grid columns via `gridSpan`/`gridBefore`/`gridAfter`), so `Table.columnCount`/column-width
  reads need it. One `w:gridCol` per column, `@w:w` in twips (dxa).
- **Row count** = number of `w:tr` children of `w:tbl` (document order = row index).
- **Column count** = number of `w:gridCol` children of `w:tblGrid` — _not_ `max(tr children)`,
  since spanned/merged cells make per-row `tc` counts diverge from the true column count.

### Row-level: `w:trPr` (§17.4.82)

- Optional, ~14 possible children (`trHeight`, `tblHeader`, `cantSplit`, `gridBefore`/`gridAfter`,
  `jc`, etc. — full list at §17.4.82). None are load-bearing for the scoped use case (row insert,
  cell text, row/column count). **Scope: omit unless/until a consumer sets row height or a
  repeating-header row** — no v1 API surface references either.
- A bare `w:tr` containing only `w:tc` children (no `w:trPr`) is schema-valid — confirmed by the
  spec's own minimal example at §17.4.79.

### Cell-level: `w:tcPr` + content nesting (§17.4.70 / §17.4.66)

- `w:tcPr` is schema-optional (omission ⇒ cell width defaults to `auto`, no span, no merge). The
  minimal-but-real-world-typical set, in the fixed schema order (`CT_TcPr` is a sequence, not a
  choice — see "Divergences" below):
  1. `w:tcW` — preferred width. Omittable (defaults to `auto`), but every spec example that shows
     a bare cell includes it (`<w:tcW w:w="0" w:type="auto"/>`), so worth emitting for cell-width
     read/write parity. `@w:type` values per `ST_TblWidth`: `auto` | `dxa` (twips) | `nil` (zero) |
     `pct` (percent of table width, e.g. `"33.3%"` in the spec's own example — note this is a
     string-typed percentage token, not a bare number).
  2. `w:gridSpan` — **only present on cells that span >1 grid column**; omitted ⇒ span of 1.
     `@w:val` is a plain integer.
  3. `w:vMerge` — **only present on cells that are part of a vertical merge.** `@w:val="restart"`
     starts a new merged group (this cell is the visible, content-bearing top of the merge);
     `@w:val="continue"` extends the group started above it; **the element present with no `@val`
     attribute at all also means `continue`** (§17.4.85 documents `val` as optional on this
     element specifically — don't require it when parsing). Cells in the same vertical-merge group
     must span the same grid columns (non-conformant otherwise, per spec) — worth an internal
     invariant if the shim ever validates seed OOXML.
- Content: every `w:tc` needs ≥1 block-level child — `w:p` for the common case, `w:tbl` for a
  nested table (out of scope here). Text nests **identically to body paragraphs**:
  `w:tc > w:p > w:r > w:t` — reuse whatever `word/paragraph.ts` already does for
  paragraph/run/text-node construction rather than re-deriving it in `word/table.ts`.

### Merged-cell scope call

The ticket asked to include `gridSpan`/`vMerge` "if in scope." The design spec's v1 API coverage
list for Tables doesn't mention a merge-cells call, but a real add-in's `context.sync()`-then-read
flow can encounter **pre-existing** merges in any seed OOXML the driving consumer's document
contains — so **read support for `gridSpan`/`vMerge` (correct row/column-count and cell-address
math around spanned/merged cells) is load-bearing even without a `mergeCells()` write API**;
without it, `TableRow.cellCount` / cell-index addressing would silently miscount on any table that
has ever been merged in real Word. Recommend: parse and preserve both on read/round-trip; defer
_writing_ new merges (a `Table.mergeCells`-equivalent) until a consumer call site actually needs
it, consistent with the design spec's "extend on demand, not speculatively" non-goal.

### Divergences / things to flag, not fully resolved by spec text alone

1. **Element order within `w:tblPr`/`w:trPr`/`w:tcPr` is schema-mandated, not stylistic.** Unlike
   attribute order or whitespace (explicitly a non-goal per the design spec — "real Word's precise
   attribute ordering/whitespace isn't a contract worth chasing"), each of `CT_TblPr`, `CT_TrPr`,
   `CT_TcPr` is an XSD `sequence`, so child _element_ order is semantically fixed by the schema
   (the order given in each subclause's "Child Elements" table above). Emitting `w:vMerge` before
   `w:tcW`, for example, would be schema-invalid even though it wouldn't show up in a naive
   byte-diff-avoidance argument. `word/table.ts` should build cell/row/table properties in the
   documented sequence.
2. **Real Word's actual authored output vs. the bare-minimum-valid schema.** The spec confirms
   `w:tblPr`/`w:tblGrid`/`w:trPr`/`w:tcPr` are individually optional, but Word's own UI-driven
   table insertion is known (informally, not re-confirmed here) to always emit a non-empty
   `w:tblPr` (`w:tblStyle` referencing a built-in style like `"TableGrid"`, `w:tblW`, `w:tblLook`
   with its default conditional-formatting bitmask) even for a bare inserted table — this is an
   authoring convention Word applies, not a schema requirement. Whether the shim needs to
   reproduce that default `tblPr`/`tblLook` boilerplate (for structural-equivalence fixture
   comparison against real captures) or can stay at the schema-minimal shape can't be settled from
   spec text alone — **needs a real captured fixture** (blocked on ticket 004, no corpus yet) to
   confirm what `context.document.body.tables` round-trips actually look like from real Word/
   Office.js's `insertTable`-equivalent path.
3. **`w:tblGrid` omission** is schema-legal but was not observed rendered-and-reconstructed in any
   primary source example here (every spec example that shows a `w:tbl` includes `w:tblGrid`) —
   treating it as always-present (per the "Table-level" section above) is a deliberate shim
   simplification, not a strict spec requirement, and should hold up fine for the scoped use case.

### Recommended minimal element/attribute set for `word/table.ts` / `operations/`

| Level        | Elements                                              | Attributes actually needed                                            |
| ------------ | ----------------------------------------------------- | --------------------------------------------------------------------- |
| Table        | `w:tbl`, `w:tblGrid`, `w:gridCol`                     | `w:gridCol/@w:w`                                                      |
| Row          | `w:tr`                                                | none (no `w:trPr` needed for v1 scope)                                |
| Cell         | `w:tc`, `w:tcPr`, `w:tcW`, `w:gridSpan`?, `w:vMerge`? | `w:tcW/@w:w`, `w:tcW/@w:type`, `w:gridSpan/@w:val`, `w:vMerge/@w:val` |
| Cell content | `w:p`, `w:r`, `w:t`                                   | (identical to existing paragraph/run/text handling)                   |

`w:tblPr` and `w:trPr` are omittable for the v1-scoped operation set (row/cell insert, cell-text
read/write, row/column count) and should be added only when a consumer call site actually needs
table-wide or row-wide formatting — consistent with the design spec's "extend on demand" principle.
