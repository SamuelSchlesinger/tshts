# Changelog

## 0.2.1 — 2026-05-16

Polish pass closing 29 of 30 documented limitations from the v0.2.0
ship review.

### Added

- **Dynamic-array spill engine.** `=SEQUENCE(10)` at A1 now fills A1..A10.
  Adds `CellData::spill_anchor` to mark ghost cells. Collisions emit
  `#SPILL!`. Ghost cells are read-only (edit-time block points to anchor)
  and render dimmed.
- **Cross-sheet auto-recalc.** Workbook-level dependency graph tracks
  `Sheet1!A1 → Sheet2!B1` edges. Edits propagate via BFS without F5.
  Includes cross-sheet circular-ref detection.
- **Conditional-formatting cache.** Per-cell style results are memoized;
  invalidated wholesale on any cell mutation or rule change.
- **`.xlsx` style writer.** `xl/styles.xml` with dedup'd fonts, fills,
  numFmts (number/currency/percent), and per-cell `s="N"` references.
  Bold, underline, fg/bg color all survive when opened in Excel.
- **R1C1 mode.** `:r1c1 on` switches the formula bar and column headers
  to `R{n}C{n}` style; storage remains A1.
- **XLOOKUP modes.** Full Excel signature: match_mode 0/-1/1/2 (exact,
  next-smaller, next-larger, wildcard) and search_mode 1/-1/2/-2.
- **Quoted 3-D refs.** `'My Sheet':'Other Sheet'!A1` parses correctly.
- **Recent-files popup.** Shows in the load dialog so users see what
  Tab will cycle through.
- **`:iterative max N` / `:iterative epsilon N`.** Tune iterative calc.

### Fixed

- VLOOKUP approximate-mode now matches Excel's early-break on sorted
  data (returns correct answer for sorted input, undefined for
  unsorted — same as Excel).
- Mouse hit-test (`cell_at`) uses rendered column rects from the last
  frame; handles custom widths and hidden columns correctly.
- DATEDIF "MD" unit borrows the actual prior-month length instead of
  a flat 30. "YD" picks the right candidate year near year boundaries.
- OFFSET out-of-bounds (negative offset or beyond sheet dims) now
  returns `#REF!` instead of an empty string.
- Stats functions (MEDIAN, STDEV, LARGE, SMALL, PERCENTILE.INC) filter
  non-finite numbers before sorting; no more silent NaN panics.
- `_xlfn.` / `_xlfn._xlws.` prefixes stripped from xlsx formulas on
  load so modern Excel files Just Work.
- xlsx named ranges round-trip in both directions.
- Sidecar clipboard JSON capped at 1MB to prevent unbounded growth.
- Hyperlink open errors surface to the status bar.

### Internal

- `Workbook::cross_sheet_dependents` / `cross_sheet_dependencies` /
  `register_cross_sheet_deps` / `propagate_cross_sheet_changes` /
  `would_create_cross_sheet_cycle` / `rebuild_cross_sheet_deps`.
- `Spreadsheet::cf_cache` (RefCell-wrapped HashMap of style results).
- `FormulaEvaluator::extract_qualified_refs` returns
  `(Option<String>, row, col)` triples for cross-sheet refs.
- `App::last_col_rects` / `last_grid_top_y` recorded each frame for
  mouse hit-test.
- New `:hide col` / `:show cols` / `:clipboard clear` palette commands.

## 0.2.0 — 2026-05-15

The Excel-equivalence pass. Major formula-language additions and the first
`.xlsx` round-trip.

### Added

**Multi-sheet language**
- Cross-sheet refs: `Sheet2!A1`, `'My Sheet'!B5:B10`. Quoted sheet names.
- 3-D refs across sheets: `Sheet1:Sheet3!A1`.
- Sheet rename propagates through every formula in the workbook (quotes the
  new name if it contains spaces).

**Array semantics**
- `Value::Array { rows, cols, data }` with shape-aware indexing.
- Array literals: `{1,2;3,4}`.
- Implicit broadcasting: `=A1:A10 * 2` produces 10 results;
  `=A1:A10 * B1:B10` is element-wise.
- New dynamic-array functions: `SEQUENCE`, `FILTER`, `SORT`, `UNIQUE`,
  `TRANSPOSE`, `SUMPRODUCT`, `FREQUENCY`.

**Errors**
- Typed errors: `Value::Error(ErrorKind)` with `#DIV/0!`, `#REF!`, `#VALUE!`,
  `#NAME?`, `#NUM!`, `#N/A`, `#NULL!`, `#SPILL!`.
- Errors propagate through binary operators; aggregates skip them.
- `IFERROR`, `IFNA`, `ISERROR`, `ISERR`, `ISNA`, `NA()`, `ERROR.TYPE`.
- Error cells render in red.

**Lookups**
- Multi-column `VLOOKUP` (`HLOOKUP`, `XLOOKUP` too).
- `INDEX(2D, row, col)` with proper shape.
- `MATCH` with wildcard support in exact mode.
- `INDIRECT`, `OFFSET`.

**Date/time**
- `TIME`, `HOUR`, `MINUTE`, `SECOND`.
- `DATEDIF` with all units (`D`/`M`/`Y`/`MD`/`YM`/`YD`).
- `WEEKDAY`, `EDATE`, `EOMONTH`, `DAYS`, `YEARFRAC`.
- `NETWORKDAYS`, `WORKDAY` with optional holiday range.
- `DATEVALUE`, `TIMEVALUE`.

**Statistics & financial**
- `MEDIAN`, `STDEV.S`, `STDEV.P`, `VAR.S`, `VAR.P`, `LARGE`, `SMALL`,
  `RANK.EQ`, `PERCENTILE.INC`, `CORREL`.
- `PMT`, `FV`, `PV`, `NPV`.
- More math: `TRUNC`, full trig (`SIN`/`COS`/`TAN` + inverses + hyperbolics),
  `DEGREES`, `RADIANS`, `FACT`, `COMBIN`, `GCD`, `LCM`, `ROUNDUP`,
  `ROUNDDOWN`, `MROUND`, `EVEN`, `ODD`.

**Text & pattern matching**
- `TEXTJOIN`, `TEXTBEFORE`, `TEXTAFTER`, `SEARCH`.
- Regex: `REGEXMATCH`, `REGEXEXTRACT`, `REGEXREPLACE`.
- `UNICHAR`, `UNICODE`, `DOLLAR`, `FIXED`, `ARRAYTOTEXT`.

**Logic**
- `IFS`, `SWITCH`, `XOR`.

**LET + LAMBDA**
- `LET(x, 5, x*2)` for local bindings within a formula.
- `LAMBDA(x, x*2)` as a first-class function value.
- Named lambdas: `:name DOUBLE LAMBDA(x, x*2)` then `=DOUBLE(7)` returns 14.
- Lambda helpers: `MAP`, `REDUCE`, `BYROW`, `BYCOL`, `SCAN`, `MAKEARRAY`.

**Tables & pivots**
- `:table create A1:D100 [name=Sales]` creates a named region with column headers.
- Structured refs: `Table1[Col1]` resolves to the matching range via named-range
  infrastructure.
- `:pivot SOURCE TARGET row=COL value=COL agg=sum|count|avg|min|max`.

**File formats**
- `.xlsx` import + export (read via `calamine`; minimal hand-rolled writer
  preserving formulas and named ranges).
- Auto-detection by extension for both startup argv and Ctrl+S / Ctrl+O.

**UI**
- Frozen columns (mirrors frozen rows).
- Hyperlink rendering for `http(s)://` cells; `Enter` opens in the system browser.
- Mouse support: scroll wheel + click-to-select.
- Function autocomplete popup while editing.
- Chart popup: `:chart bar|line A1:A10`.
- Error cells render in red.

**Tooling**
- `:goalseek TARGET EXPECTED INPUT` — bisects `INPUT` until `TARGET = EXPECTED`.
- `:trace` / `:trace dependents` — surface formula precedents/dependents.
- `:iterative on/off` — toggle iterative-calc mode (for intentional circular refs).
- `:r1c1 on/off` — A1 ↔ R1C1 mode toggle.
- `:validate <COL> <predicate>` — column-level data validation.

### Internals

- `ExpressionEvaluator` carries optional workbook + named-range + local-scope
  contexts.
- Parser lexer handles `!`, `'`, `{`/`}`/`;`, dotted function names
  (`STDEV.S`), and structured table refs (`Table[Col]`).
- Thread-local parse cache (256 entries) shared between dependency extraction
  and evaluation.

### Notes

- Tests: 324 unit + 12 doctest, all passing.
- 7 ignored tests require live network access (cryptoprices.cc).

## 0.1.0 — initial release

- Terminal spreadsheet with basic formula support
- Workbook tabs, undo/redo, find/replace
- CSV import/export
- Native `.tshts` JSON format
