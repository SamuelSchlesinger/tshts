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

## Dependencies

- **ratatui/crossterm** - Terminal UI
- **serde/serde_json** - .tshts file serialization
- **csv** - CSV import/export
- **reqwest** (blocking) - GET() formula function
