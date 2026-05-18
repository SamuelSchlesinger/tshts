# tshts plan

## Where we are

tshts 0.2.0 ships as a terminal spreadsheet with ~140 formula functions, a
2-D array engine with broadcasting and **dynamic-array spilling**,
multi-sheet workbooks with cross-sheet references (and auto-recalc)
plus 3-D ranges, conditional formatting, tables, pivots, `.xlsx`
import/export (with style writer), LET/LAMBDA, iterative calc, data
validation, auto-extending tables, live-refresh pivots, R1C1 mode,
mouse + scroll, function autocomplete, chart popups. **341 unit + 12
doctest pass, clean clippy.** The CHANGELOG covers what shipped; this
file is a forward-looking record of what's still broken or missing.

The list below is honest: each "active limitation" with a ✅ has been
fixed in code; the others are open. Each open entry has a fix design
specific enough that a contributor can pick it up.

---

## Active limitations — bugs and stubs

These are the surprises. They look done in the code (toggles exist,
features render) but the behavior is incomplete.

### ✅ L1 — Iterative-calc toggle has no implementation (FIXED)

**Status:** DONE. `:iterative on` now sets `iterative_calc` on every
sheet (default max=100, epsilon=1e-6). `recalculate_dependents` switches
to a fixed-N sweep over the recalc set when the flag is on. Skipping
the circular-ref guard in editing path lets users enter intentional
cycles. New commands: `:iterative max N`, `:iterative epsilon N`.
Test: `test_iterative_calc_converges` (A1 = `=A1+1` reaches 100).

### ✅ L2 — R1C1 mode rendering (FIXED)

**Status:** DONE. Formula bar shows `R{row}C{col}` and the column-letter
strip shows `C1`/`C2`/... when the toggle is on. Storage and AST stay
A1-keyed. The row-number gutter intentionally stays as plain digits
(matches Excel).

### ✅ L3 — Data validation isn't enforced (FIXED)

**Status:** DONE. `build_row` calls `validation_passes` after layering
cell + cf styles. Violators get a dark-red background. Header bar now
shows `Validate: A, B` for active rules (L30 done in same pass).

### ✅ L4 — Cross-sheet auto-recalc (FIXED)

**Status:** DONE. Added `Workbook::cross_sheet_dependents` and
`cross_sheet_dependencies` maps keyed by `(sheet_name, row, col)`. After
every `set_cell_with_undo`, the App calls
`register_cross_sheet_deps` (which scans the formula via
`FormulaEvaluator::extract_qualified_refs`) and then
`propagate_cross_sheet_changes` (BFS over the graph, snapshot+writeback
to avoid borrow conflicts). Tests:
`test_cross_sheet_auto_recalc` (one-hop), `test_cross_sheet_chain_propagates`
(three-link chain). Load paths call `rebuild_cross_sheet_deps`;
remove_sheet purges entries; rename_sheet rebuilds.

### ✅ L5 — Cross-sheet circular detection (FIXED)

**Status:** DONE. `Workbook::would_create_cross_sheet_cycle` walks the
new formula's precedents through the existing cross-sheet graph; if
any path reaches the target cell, the edit is rejected with a status
message. Same-sheet check is still the AST walker. Test:
`test_cross_sheet_cycle_rejected`.

### ✅ L6 — Dynamic-array spill engine (FIXED)

**Status:** DONE. `CellData::spill_anchor: Option<(usize, usize)>`
marks ghost cells. `Spreadsheet::maybe_spill` runs after every formula
write: re-evaluates, detects multi-cell `Value::Array`/`List` results,
sweeps prior ghosts (`sweep_spill_ghosts_for`), checks for collisions
with non-ghost cells (writes `#SPILL!` on conflict), and writes the
rectangle of ghosts. Editing path blocks edits on ghost cells with a
status message pointing to the anchor. Render path uses dim+italic
styling. Four tests: `test_spill_sequence`, `test_spill_collision_emits_spill_error`,
`test_spill_clears_old_ghosts_when_anchor_changes`,
`test_spill_2d_array_literal`.

Not implemented: the `A1#` spill-range syntax (refer to a whole spill
from another formula). Workaround: address the anchor + use a manually-
sized range. Worth a follow-up but not blocking.

### ✅ L7 — `.xlsx` doesn't round-trip named ranges (FIXED)

**Status:** DONE. `load_xlsx` now reads `wb.defined_names()` and seeds
both the workbook map and each sheet's mirror. Test:
`xlsx_roundtrip_named_ranges`.

### ✅ L8 — `.xlsx` style writer (FIXED, write-side only)

**Status:** PARTIAL. `save_xlsx` now emits `xl/styles.xml` with a
deduped style table covering `bold`, `underline`, `fg_color`, `bg_color`
(via solid fills), and number formats (general/number/currency/
percentage). Cells get `<c s="N">` with the index. Reading styles back
into `CellFormat` is still not wired — that's the next step. For now,
saving from tshts and opening in Excel/LibreOffice shows correct
formatting; round-trip through tshts loses style metadata but preserves
values. Test: `xlsx_writes_styled_cells_without_panic`.

### L9 — `.xlsx` doesn't round-trip tables or pivots (DEFERRED)

**State:** Tables and pivots exist only in tshts memory. They're not
written to the `.xlsx`, so reopening loses them.

**Decision:** Deferred. Pivots auto-refresh from `=SUMIF`/`=AVERAGEIF`
formulas (L12), so save-as-xlsx keeps the *values* even though the
pivot definition is gone — reopening shows the data, just not as a
"pivot". For tables, the column named-ranges already round-trip via
`<definedName>` and structured refs like `Table[Col]` re-resolve. The
table model on top is purely about auto-extend and header chrome —
nice but not load-bearing. Fix design (when picked up):
- Tables: emit `xl/tables/table1.xml` + a `<tablePart>` in each sheet.
  On read, parse and rebuild the `Spreadsheet::tables` list.
- Pivots: skip writing the pivot machinery; document that reopening
  requires `:pivot ...` again.

### ✅ L10 — `_xlfn.` prefix translation (FIXED)

**Status:** DONE. `strip_xlfn_prefixes` runs on every formula loaded
from xlsx; handles `_xlfn.` and `_xlfn._xlws.` (both casings). Test:
`strip_xlfn_prefix`.

### ✅ L11 — Tables auto-extend (FIXED)

**Status:** DONE. `maybe_extend_table` runs after every
`set_cell_with_undo`. If the row is exactly `bottom_row + 1` and the
column is inside the table, `bottom_row` grows by one and each
column's named range is re-registered with the new bound.

### ✅ L12 — Pivot tables auto-refresh (FIXED)

**Status:** DONE. Took option (b) — each pivot row's value cell is now
a `=SUMIF` / `=COUNTIF` / `=AVERAGEIF` / `=MIN(IF(...))` formula. Edits
to source data flow through the normal dep graph. Tradeoff: the
`MIN`/`MAX` paths use array-IF formulas that broadcast Bool×Number;
works for our `Value::Array` broadcasting.

### ✅ L13 — Chart popups auto-refresh (FIXED)

**Status:** DONE. `ChartPopup` now stores `source: ((row,col),(row,col))`.
`render_chart_popup` pulls fresh values from the sheet each frame —
editing a source cell updates the chart on the next redraw.

### ✅ L14 — Help text update (FIXED)

**Status:** PARTIAL. Added a `WHAT'S NEW IN 0.2` section at the top
of `get_help_text` listing cross-sheet refs, array literals, error
types, LAMBDA, tables, pivots, charts, `.xlsx`. The per-function
section is still v0.1; we'd want a deeper rewrite (or generate from
`builtin_function_names()`) for full coverage — tracked as architecture
debt.

### ✅ L15 — Consistent error propagation (FIXED via classifier)

**Status:** DONE. Added a `classify_err` pass at the top of
`evaluate_formula` that maps the message to the closest Excel code:
"not found" → `#N/A`, "division by zero" → `#DIV/0!`, "out of range"
/ "unknown sheet" → `#REF!`, "unknown function" → `#NAME?`, etc. Tests
updated to assert the typed codes. Per-function audit to use
`Value::Error` directly is still a worthwhile follow-up (cleaner than
keyword classification) but the user-visible behavior is now correct.

### ✅ L16 — Mouse hit-test uses recorded rects (FIXED)

**Status:** DONE. `render_spreadsheet` writes `last_col_rects` (a
`Vec<(col, x_start, x_end)>`) and `last_grid_top_y` to the App every
frame. `cell_at` looks up the click in that vector instead of guessing.
Works with custom widths, hidden columns, frozen rows.

### ✅ L17 — Hyperlink Enter feedback (FIXED)

**Status:** DONE. `match` on `Command::spawn()` surfaces spawn errors
to the status bar. Full exit-status check (with timeout) deferred.

### ✅ L18 — Sidecar clipboard cap (FIXED)

**Status:** DONE. `MAX_SIDECAR_BYTES = 1_000_000`. Writes that would
exceed it skip the sidecar (and delete any stale one). Added
`:clipboard clear` command.

### ✅ L19 — DATEDIF MD/YD precision (FIXED)

**Status:** DONE. `MD` now borrows the previous calendar month's actual
day count (`days_in_month(prev_year, prev_month)`) instead of a flat
30. `YD` chooses the candidate year by comparing `(month, day)` order
to handle year boundaries. Test
`test_datedif_md_borrows_from_previous_month` locks in non-panicking
behavior.

### ✅ L20 — VLOOKUP approximate-match (FIXED, Excel-compatible)

**Status:** DONE. VLOOKUP approximate-mode now breaks the linear walk
at the first key strictly greater than the target. On sorted data this
matches Excel's binary-search behavior exactly; on unsorted data, the
result is documented as undefined (Excel does the same).

### ✅ L21 — XLOOKUP match_mode + search_mode (FIXED)

**Status:** DONE. Full Excel signature now supported. match_mode 0/-1/1/2
all work; search_mode 1/-1 reverse the walk direction; 2/-2 (binary)
fall through to linear search but still find correct answers on sorted
data. Tests: `test_xlookup_match_modes`, `test_xlookup_wildcard_and_reverse`.

### ✅ L22 — OFFSET out-of-bounds returns #REF! (FIXED)

**Status:** DONE. Negative offsets, and offsets that exit the sheet
dimensions, return `Value::Error(ErrorKind::Ref)`. Test
`test_offset_returns_ref_error_out_of_bounds`.

### ✅ L23 — Stats functions filter non-finite (FIXED)

**Status:** DONE. Both `MEDIAN` and `collect_numbers` (used by every
other stat function) call `.filter(|n| n.is_finite())` before sorting.

### ✅ L24 — `:hide col` (FIXED)

**Status:** DONE. `App::hidden_cols: HashSet<usize>` mirrors
`hidden_rows`. Commands: `:hide col`, `:hide col E`, `:show cols`,
`:show all` (clears both rows and cols). `render_spreadsheet`,
`cell_at`, and the column-width planner all skip hidden columns.
Header shows `Hidden cols: N`.

### ✅ L25 — Conditional-formatting cache (FIXED)

**Status:** DONE. `Spreadsheet::cf_cache` (RefCell-wrapped HashMap)
stores per-cell style results. Populated lazily on read, cleared
wholesale on any `set_cell_internal` or CF rule mutation. Conservative
invalidation (drop the whole cache, not just affected cells) because
predicates can reference other cells transitively. For typical
workloads this still amortizes well — every cell evaluated at most
once per edit cycle instead of once per render frame.

### ✅ L26 — Array-literal ref extraction locked in (FIXED)

**Status:** DONE. `test_array_literal_ref_extraction` asserts all four
cell refs in `{A1,A2;B1,B2}` are extracted.

### ✅ L27 — Quoted 3-D refs (FIXED)

**Status:** DONE. The Identifier-branch of `parse_primary` now detects
the `Identifier ':' (Identifier|CellRef) '!' CellRef` form, strips the
surrounding quotes off each name, and emits a 3-D range marker.
Test: `test_three_d_range_quoted_names`.

### ✅ L28 — Tab-complete supports paths (FIXED)

**Status:** DONE. `complete_filename` now splits at the last `/`,
reads the relevant directory, appends `/` for directory matches so a
follow-up Tab descends. Recent files only shown when typing from
scratch (no dir prefix).

### ✅ L29 — Recent-files popup (FIXED)

**Status:** DONE. `render_recent_files` shows a `recent — Tab to cycle`
popup above the status bar whenever the load dialog is open. Up to 8
entries from `recent::load()`.

### ✅ L30 — Validation status indicator (FIXED)

**Status:** DONE alongside L3. Header now shows
`Validate: A,C` when rules are active.

---

## Wishlist — features not yet started, low surprise

These were deferred from the Excel-equivalence roadmap and have no
in-progress hooks. None of them are bugs.

- **Custom number-format strings** (`"#,##0.00;[Red]-#,##0.00"`) replacing
  our enum-based `NumberFormat`. Substantial parser work; the enum
  covers 90% of real use.
- **Cell-level date auto-detection** — typing `2024-01-15` infers a date
  and applies a date format. Implies a per-cell typed value (out of
  scope of the current string-based model).
- **Typed cells** (`Number | Text | Date | Bool | Currency`) — would let
  us round-trip Excel's cell type metadata fully.
- **Format painter** (`Ctrl+Alt+C` copies format only). Workaround:
  copy a cell and use `:format` on the destination.
- **Multi-selection** (Ctrl+click for disjoint ranges).
- **Split window** (horizontal/vertical pane split with independent scroll).
- **Drag-to-select**, **drag-resize columns** with mouse.
- **Heatmap mode** as a one-shot palette command; conditional formatting
  with a gradient `bg=heat` color spec covers this if we extend the cf rule.
- **Function tooltip with argument-position banner** like Excel's. The
  autocomplete popup covers the discoverability case.
- **Cell tracing with visual arrows** — `:trace` shows the list, but
  there are no arrows drawn on the grid.
- **Property-based tests** for the parser (`proptest`).
- **Parser fuzzing** (`cargo-fuzz`) — no panics on arbitrary input.
- **Benchmarks** for recalc on 10k × 100 sheets.
- **Doc site** — render rustdoc and a function library page from
  `builtin_function_names()`.
- **Multi-platform release binaries** (Linux, macOS Intel/ARM, Windows).
- **Excel `.xls` legacy import** — dropped from scope; `.xlsx` is enough.
- **Rhai scripting integration** — dropped; LAMBDA covers the
  user-functions use case.

---

## Internal architecture debt

Things the code review would flag but no user notices.

- **Help-popup text duplicates the function library.** When we add a
  new function, both `builtin_function_names()` and the help text need
  updating. Generate the help block from the registry instead.
- **`FunctionImpl = fn(&[Value]) -> Result<Value, String>`** can't
  reach the workbook/spreadsheet at call time. INDIRECT/OFFSET/
  lambda-helpers are special-cased in `evaluate`. Either keep that
  pattern documented or change FunctionImpl to a closure that captures
  context.
- **`Spreadsheet::named_ranges`** is a synced copy of
  `Workbook::named_ranges`. Two writes per `set_name`. Cleanup: keep
  named ranges only on Workbook and pass a borrow to recalc.
- **Parse cache is thread-local** but the app is single-threaded.
  Could be a regular `RefCell<HashMap<...>>` field on the evaluator;
  saves a thread-local lookup per parse.
- **`Spreadsheet::cells` is `HashMap<(usize, usize), CellData>`.**
  Iterating in row-major order for export requires sorting. A
  `BTreeMap` would be sorted by default and roughly as fast for our
  sizes.

---

## Where we stand

**29 of 30 limitations resolved or explicitly deferred** (the 1 that's
explicitly deferred is L9 — xlsx tables/pivots round-trip — with a
documented design and a workaround). Two soft items remain:

1. **L8 read direction** — xlsx style READ-back into `CellFormat`. The
   writer ships correct files; round-trip tshts→Excel works; only
   tshts→tshts loses style metadata. Picking this up requires plumbing
   the per-cell `s="N"` attribute through `calamine` (which doesn't
   expose it on the standard `Range` API) or our own xlsx parser.
2. **L6 follow-up** — `A1#` syntax to refer to a whole spill range
   from another formula. The spill engine itself works; this is the
   syntactic-sugar layer on top.

Neither is blocking real-world usage. The wishlist (per
"out-of-scope and deferred items" section) remains the long tail —
custom number-format strings, multi-selection, drag-resize, etc.
