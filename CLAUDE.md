# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Build & Development Commands

```bash
cargo run --release      # Build and run
cargo build --release    # Build for production
cargo test               # Run all tests
cargo test <name>        # Run tests matching name
cargo doc --open         # Generate documentation
```

## Architecture

TSHTS is a terminal spreadsheet following clean architecture with four layers:

```
src/
├── domain/           # Core business logic (no external dependencies)
│   ├── models.rs     # Spreadsheet, CellData
│   ├── services.rs   # FormulaEvaluator, CsvExporter
│   └── parser.rs     # Recursive descent formula parser
├── application/      # Application orchestration
│   └── state.rs      # App state, AppMode enum, undo/redo
├── infrastructure/   # External integrations
│   └── persistence.rs # File I/O, .tshts JSON serialization
├── presentation/     # User interface
│   ├── ui.rs         # Terminal rendering (ratatui)
│   └── input.rs      # Keyboard handling (InputHandler)
└── lib.rs, main.rs   # Entry points
```

## Architectural Invariants

**Layer Dependencies**: Domain has no external dependencies. Application depends on domain. Infrastructure and presentation depend on both.

**Formula Functions**: Registered in `FunctionRegistry` in `parser.rs`. All take `&[Value]` and return `Result<Value, String>`. The `Value` enum supports dual number/string types with `.to_number()`, `.to_string()`, and `.is_truthy()` conversions.

**Cell Dependencies**: A single workbook-wide `WorkbookGraph` on `Workbook` tracks all dependencies (same-sheet AND cross-sheet) keyed by `NodeKey = (SheetId, row, col)`. Not serialized — lazily built by `Workbook::build_dep_graph_from_scratch` on first recalc. Per-sheet `dependents`/`dependencies` maps no longer exist (deleted with the legacy cascade). `Spreadsheet::set_cell` is a pure write — propagation goes through `Workbook::recalc_via_graph_result`, which the public `set_cell_on_active` / `write_cells_on_active` / `clear_*_on_active` mutators call internally. See "Calc engine architecture" below.

**App Modes**: `AppMode` enum drives UI state. Each mode has a corresponding handler in `InputHandler` and rendering logic in `ui.rs`. State transitions go through methods on `App`.

**Undo/Redo**: Cell modifications should use `App::set_cell_with_undo()` and `App::clear_cell_with_undo()` to enable undo. Direct `spreadsheet.set_cell()` bypasses undo tracking.

**Formula Parser**: Recursive descent with operator precedence (low to high): equality, comparison, addition, concatenation, multiplication, power, unary, primary. Logical ops (AND, OR, NOT) are functions, not operators.

**Error Literals**: Excel-style error values (`#REF!`, `#N/A`, `#DIV/0!`, `#VALUE!`, `#NAME?`, `#NUM!`, `#NULL!`, `#SPILL!`) are first-class AST nodes (`Expr::ErrorLit(ErrorKind)`), lexed greedily by the longest-match-first rule. They evaluate to `Value::Error(kind)` and propagate through arithmetic, function calls, and unary ops via the standard `first_error()` cascade. Formula adjustment (paste, row/col insert/delete) emits `Expr::ErrorLit(Ref)` when a relative shift takes a reference past the origin, instead of clamping or collapsing the formula. The serialized form round-trips cleanly: `=#REF!+B5` parses back to the same AST.

**Cross-sheet structural edits**: `Workbook::insert_row_on_active`, `delete_row_on_active`, `insert_col_on_active`, `delete_col_on_active` perform the same-sheet structural mutation and also walk every OTHER sheet's formulas to shift any sheet-qualified refs to the mutated sheet (e.g. inserting a row at Sheet1!A5 shifts `=Sheet1!A5` to `=Sheet1!A6` on every other sheet). Refs to a deleted sheet's removed row/col become `#REF!`. Routes through `App::insert_row`/`delete_row`/etc.

**Sheet rename/delete**: `Workbook::rename_sheet` rewrites all formula refs (and named-range values) from old to new name, marks formula cells dirty, then triggers `recalc_via_graph_result()`. `Workbook::remove_sheet` rewrites every dangling `=GoneSheet!A1` on surviving sheets to `=#REF!` (Excel-equivalent), then calls `rebuild_cross_sheet_deps` (which wipes + rebuilds the unified graph from scratch).

**Mutation API**: Single source of truth is `Workbook::set_cell_on_active` / `clear_cell_on_active` / `write_cells_on_active` / `clear_cells_on_active`. Each writes the cell(s) via `Spreadsheet::set_cell`/etc (pure writes — no cascade), updates the unified graph, marks dirty, and runs `recalc_via_graph_result()` to flush dependents. Callers outside the workbook should never call `Spreadsheet::set_cell` directly — they'd write the value but skip propagation. The application-level helper `App::propagate_cell_change(row, col)` is kept for bespoke paths that need to manually flush but is rarely needed (the standard mutators auto-flush).

## Testing

Two layers, both run by plain `cargo test`:

**Unit tests** (in-tree under `#[cfg(test)] mod tests`): drive `InputHandler::handle_key_event` against synthetic `KeyCode` events and assert on `App` state. Fast, ~500 tests covering formula correctness, mode transitions, vim operator grammar, persistence, and UI state. See helpers `typestr` / `key` / `ctrl` at the bottom of `src/presentation/input.rs`.

**PTY end-to-end tests** (`tests/pty_*.rs` + `tests/common/mod.rs`): spawn the real release binary in a pseudo-tty via `portable-pty`, parse the rendered escape sequences back into a virtual screen via `vt100`, and assert on what a human would actually see. Catches crossterm raw-mode bugs, terminal byte conflicts (Ctrl+I ≡ Tab, Ctrl+H ≡ Backspace), the alternate-screen lifecycle, and startup regressions. Organized by area: `pty_smoke` (basic launch/quit), `pty_ux` (mode chip / dirty flag / status messages), `pty_workflows` (multi-step user flows), `pty_advanced` (cross-sheet refs / filters / freezes / charts), `pty_power` (tables / pivots / 3-D refs / validation), `pty_polish` (edge cases / error paths), `pty_edges` (case preservation / boundary conditions). Total ~120 PTY tests.

```rust
// tests/pty_smoke.rs pattern — see tests/common/mod.rs for the Harness API.
let mut h = Harness::new();
h.assert_contains("-- NORMAL --");
h.send_text("i");
h.assert_contains("-- INSERT --");
h.send_text("howdy");
h.send_enter();
h.assert_contains("howdy");
h.send_text(":q!");
h.send_enter();
assert_eq!(h.wait_for_exit(Duration::from_secs(3)), Some(0));
```

Useful Harness methods: `send_text` / `send_enter` / `send_esc` / `send_tab` / `send_ctrl(c)` / `send_arrow(Arrow::Up)`, `assert_contains` / `assert_absent` / `wait_for_text(needle, timeout)` (polls), `screen_contents()` / `row(n)` / `cursor()`, `wait_for_exit(timeout)` / `has_exited()`. Drop SIGKILLs the child.

Gotchas when writing new PTY tests:
- The default column width is narrow. Long values get clipped in the rendered grid even though the underlying cell holds the full string — keep assertion needles ≤ ~8 chars or check the formula-bar/status line, which is full-width.
- `Enter` commits an edit (cursor moves down to A2); `Esc` cancels it.
- The harness uses a 30x120 screen; assertions on layout should respect that.

## Calc engine architecture

tshts has ONE recalc engine: the unified graph-driven level executor. (The legacy per-sheet `recalculate_dependents` cascade and the `propagate_cross_sheet_changes*` family were deleted — see "Mutation discipline" below.)

**The graph**: `Workbook::graph` is a workbook-wide `WorkbookGraph` keyed by `NodeKey = (SheetId, row, col)`. Bidirectional: `dependencies[N]` is the set of cells `N` reads, `dependents[N]` is the set of cells that read `N`. Same-sheet AND cross-sheet edges live in this one structure. Built lazily via `Workbook::build_dep_graph_from_scratch` (called from `register_cross_sheet_deps` on the first write, on load, and after structural edits) and maintained incrementally by `register_cross_sheet_deps` per write.

**`Workbook::cell_purities`** is the parallel-keyed cache of per-cell `FunctionPurity` classifications. Pure cells are stored implicitly (absent from the map). Volatile / side-effecting cells are explicit.

**`Workbook::structural_targets`** is the per-`VolatileStructural`-cell cache of the dynamic targets (INDIRECT/OFFSET resolved cells) recorded after the last eval. Smart auto-seed compares against the user-dirty closure to decide whether to re-seed.

**Mutation discipline**: `Workbook::set_cell_on_active` / `clear_cell_on_active` / `write_cells_on_active` / `clear_cells_on_active` are the only public single-source-of-truth mutation APIs. Each one writes the cell(s) via `Spreadsheet::set_cell`/`clear_cell`/`set_many`/`clear_many` (which are pure writes — no cascade), updates the unified graph via `register_cross_sheet_deps`, marks dirty, and triggers a single `recalc_via_graph_result()` that propagates to all dependents. `Spreadsheet::set_cell` directly is reserved for low-level test fixtures and intra-engine machinery (the iterative-cyclic loop).

**`Workbook::recalc_via_graph_result`** is the entry point — drains `dirty`, builds `seeds`, computes `transitive_dependents` + topological levels via Kahn's algorithm, dispatches to an executor. Self-loops are retained in the graph (they're cyclic — see "Iterative calc" below).

The executor is pluggable via the `RecalcExecutor` trait (`src/domain/services/executor.rs`):
- `SequentialExecutor` — single-threaded reference impl. One workbook snapshot for the whole recalc, mutated between levels so the next level reads fresh values from prior levels.
- `ParallelExecutor` — rayon-based; partitions each level by function purity, dispatches pure cells via `par_iter().with_min_len(K)`, runs structural-volatile (`INDIRECT`, `OFFSET`) and side-effecting (`GET`) cells serially within the level barrier. Uses `Arc<Workbook>` snapshot reused across levels via `Arc::make_mut` at the level boundary (refcount = 1 after `par_iter().collect()`).

`recalc_via_graph` auto-selects: Parallel when any level has ≥ `TSHTS_PAR_THRESHOLD` cells (default 512), otherwise Sequential.

**Tuning**: `TSHTS_PAR_THRESHOLD=N` for parallel-dispatch cutoff. `RAYON_NUM_THREADS=N` for worker count. `cargo bench --bench calc_engine` runs the archetype benchmarks.

**Volatile semantics** (matched to Excel/OpenFormula):
- `NOW`/`TODAY` read a clock snapshot captured at recalc start (via the `RECALC_CLOCK` thread-local published by each executor's `run`). Two clock-volatile cells in the same pass return identical values; calls outside a recalc fall back to `SystemTime::now()`.
- `RAND`/`RANDBETWEEN` use a thread-local PRNG.
- `OFFSET`/`INDIRECT` are tagged `VolatileStructural` and consulted via the smart auto-seed: a structural cell is seeded into the dirty set only when its recorded `structural_targets` intersect the user-dirty closure.
- `GET` is `SideEffecting`; serialized via the HTTP-fetcher worker (cache-seeded for tests via `fetcher::test_hooks`).

**Iterative calc**: `Workbook::iterative_calc_cyclic` runs over the cyclic remainder from `topo_levels_from_seeds`. Gauss-Seidel-style: each pass evaluates against a fresh snapshot, mutates the live workbook between passes. Uses the highest `iter_max` and tightest `iter_epsilon` across participating sheets. Two-pass string stability for non-numeric flip-flop detection.

**Non-convergence is strict**: cycles that exhaust `iter_max` without converging get `#NUM!` written to every cyclic cell (rather than the iter_max'th iteration value, which would be a misleading artifact of the setting). `CalcError::DidNotConverge` is returned and bubbles to `App::status_message` via `recalc_via_graph_result`. Convergent cycles (legitimate fixed-point iteration like `=A1/2+1 → 2`) get the converged value, no error. Self-loops (`=A1+1` written to A1) are normal cycles: routed to `iterative_calc_cyclic` like any other cyclic remainder.

## Dependencies

- **ratatui/crossterm** - Terminal UI
- **serde/serde_json** - .tshts file serialization
- **csv** - CSV import/export
- **reqwest** (blocking) - GET() formula function
- **rayon** - Work-stealing parallel iterator (used by `ParallelExecutor`)
- **portable-pty/vt100** (dev) - PTY-based end-to-end test harness
