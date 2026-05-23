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

**Cell Dependencies**: `Spreadsheet` tracks cell dependencies bidirectionally. When a cell changes via `set_cell()`, dependent cells automatically recalculate. Dependencies are not serialized - call `rebuild_dependencies()` after loading.

**App Modes**: `AppMode` enum drives UI state. Each mode has a corresponding handler in `InputHandler` and rendering logic in `ui.rs`. State transitions go through methods on `App`.

**Undo/Redo**: Cell modifications should use `App::set_cell_with_undo()` and `App::clear_cell_with_undo()` to enable undo. Direct `spreadsheet.set_cell()` bypasses undo tracking.

**Formula Parser**: Recursive descent with operator precedence (low to high): equality, comparison, addition, concatenation, multiplication, power, unary, primary. Logical ops (AND, OR, NOT) are functions, not operators.

**Error Literals**: Excel-style error values (`#REF!`, `#N/A`, `#DIV/0!`, `#VALUE!`, `#NAME?`, `#NUM!`, `#NULL!`, `#SPILL!`) are first-class AST nodes (`Expr::ErrorLit(ErrorKind)`), lexed greedily by the longest-match-first rule. They evaluate to `Value::Error(kind)` and propagate through arithmetic, function calls, and unary ops via the standard `first_error()` cascade. Formula adjustment (paste, row/col insert/delete) emits `Expr::ErrorLit(Ref)` when a relative shift takes a reference past the origin, instead of clamping or collapsing the formula. The serialized form round-trips cleanly: `=#REF!+B5` parses back to the same AST.

**Cross-sheet structural edits**: `Workbook::insert_row_on_active`, `delete_row_on_active`, `insert_col_on_active`, `delete_col_on_active` perform the same-sheet structural mutation and also walk every OTHER sheet's formulas to shift any sheet-qualified refs to the mutated sheet (e.g. inserting a row at Sheet1!A5 shifts `=Sheet1!A5` to `=Sheet1!A6` on every other sheet). Refs to a deleted sheet's removed row/col become `#REF!`. Routes through `App::insert_row`/`delete_row`/etc.

**Sheet rename/delete**: `Workbook::rename_sheet` rewrites all formula refs (and named-range values) from old to new name, then triggers cross-sheet propagation so dependents recompute. `Workbook::remove_sheet` rewrites every dangling `=GoneSheet!A1` on surviving sheets to `=#REF!` (Excel-equivalent), then purges the dep graph.

**Cross-sheet propagation helper**: `App::propagate_cell_change(row, col)` runs `register_cross_sheet_deps` + `propagate_cross_sheet_changes` on the workbook for the current sheet. Use it from any mutation path (cut, paste, replace_all, vim delete, undo/redo) that writes/clears cells outside of `set_cell_with_undo` / `clear_cell_with_undo`, which already call it internally.

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

tshts ships two recalc engines that produce identical results:

**Legacy per-sheet cascade**: `Spreadsheet::set_cell` triggers `recalculate_dependents` which BFSes the per-sheet `dependents` HashMap. Cross-sheet propagation runs as a separate `Workbook::propagate_cross_sheet_changes` pass. This is the path most user edits flow through (every `set_cell_on_active` / `write_cells_on_active` call).

**Graph-driven level executor**: `Workbook::recalc_via_graph` builds the unified workbook-level dep graph (`WorkbookGraph`, keyed by stable `SheetId(u32)`), drains the dirty set, computes topological levels via Kahn's algorithm, and walks levels in order. Each level evaluates against an immutable workbook snapshot; results merge back at the level boundary. Used by `:recalc` (`App::recalc_all`).

The executor is pluggable via the `RecalcExecutor` trait (`src/domain/services/executor.rs`):
- `SequentialExecutor` — single-threaded reference impl.
- `ParallelExecutor` — rayon-based; partitions each level by function purity, dispatches pure cells via `par_iter().with_min_len(64)`, runs structural-volatile (`INDIRECT`, `OFFSET`) and side-effecting (`GET`) cells serially within the level barrier. Falls back to sequential below `parallel_threshold`.

`recalc_via_graph` auto-selects between the two: Parallel when any level has ≥ `TSHTS_PAR_THRESHOLD` cells (default 512), otherwise Sequential. Below that the per-level workbook clone dominates the parallel savings.

**Tuning**: set `TSHTS_PAR_THRESHOLD=N` to change the parallel-dispatch cutoff. `RAYON_NUM_THREADS=N` controls worker count. `cargo bench --bench calc_engine` runs the archetype benchmarks (wide/deep/fanout × small/medium/large) so you can pick a threshold that matches your workload.

**Volatile semantics** (matched to Excel/OpenFormula):
- `NOW`/`TODAY` read a clock snapshot captured at recalc start (via the `RECALC_CLOCK` thread-local published by each executor's `run`). Two clock-volatile cells in the same pass return identical values; calls outside a recalc fall back to `SystemTime::now()`.
- `RAND`/`RANDBETWEEN` use a thread-local PRNG; cross-worker non-determinism is the intended trade — within a pass each worker's outputs are independent.
- `OFFSET`/`INDIRECT` are tagged `VolatileStructural` and auto-seeded into the dirty set on every recalc, so changes to their value-derived targets propagate through their static dependents within one pass (matches Excel's "always recompute volatile").
- `GET` is `SideEffecting`; serialized via the existing HTTP-fetcher worker (no executor changes needed).

**Cross-sheet cycles**: handled by `Workbook::iterative_calc_cyclic` — a workbook-level Gauss-Seidel loop that walks every cyclic cell across all sheets per pass. Uses the highest `iter_max` and tightest `iter_epsilon` across participating sheets. Detects non-convergence (returns `Err(iter_max)`) and handles non-numeric flip-flop via two-pass string stability. Per-pass post-write maintenance (CF cache, spill ghosts, `maybe_spill`) is performed inside `with_recalc_context` to mirror the acyclic path.

## Dependencies

- **ratatui/crossterm** - Terminal UI
- **serde/serde_json** - .tshts file serialization
- **csv** - CSV import/export
- **reqwest** (blocking) - GET() formula function
- **rayon** - Work-stealing parallel iterator (used by `ParallelExecutor`)
- **portable-pty/vt100** (dev) - PTY-based end-to-end test harness
