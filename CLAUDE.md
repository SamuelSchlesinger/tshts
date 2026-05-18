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

## Dependencies

- **ratatui/crossterm** - Terminal UI
- **serde/serde_json** - .tshts file serialization
- **csv** - CSV import/export
- **reqwest** (blocking) - GET() formula function
- **portable-pty/vt100** (dev) - PTY-based end-to-end test harness
