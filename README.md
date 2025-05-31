# TSHTS - Terminal Spreadsheet

An efficient, lightweight terminal-based spreadsheet application built in Rust. TSHTS brings the power of spreadsheet calculations to your command line with an intuitive interface and essential formula support.

![License](https://img.shields.io/badge/license-MIT-blue.svg)
![Rust](https://img.shields.io/badge/language-Rust-orange.svg)

![TSHTS Screenshot](screenshot.png)

## 🚀 Quick Start

### Installation

```bash
# Clone the repository
git clone https://github.com/SamuelSchlesinger/tshts.git
cd tshts

# Build and run
cargo run --release
```

### First Steps

1. **Navigate**: Use arrow keys or `hjkl` to move between cells
2. **Edit**: Press `Enter` or `F2` to edit a cell
3. **Formula**: Start with `=` for formulas (e.g., `=A1+B1`, `=SUM(A1:A10)`)
4. **Save**: Press `Ctrl+S` to save your spreadsheet
5. **Help**: Press `F1` or `?` for comprehensive help

## ✨ Key Features

### 🧮 Powerful Formula Engine
- **Arithmetic Operations**: `+`, `-`, `*`, `/`, `**` (power), `%` (modulo)
- **Comparison Operators**: `<`, `>`, `<=`, `>=`, `=`, `<>` (not equal)
- **Essential Functions**: `SUM`, `AVERAGE`, `MIN`, `MAX`, `IF`, `AND`, `OR`, `NOT`, `ABS`, `SQRT`, `ROUND`
- **Function-Based Logic**: All logical operations use clean function syntax
- **Cell References**: Standard notation (A1, B2, AA123, etc.)
- **Range Support**: Use ranges like `A1:C3` in functions
- **Circular Reference Detection**: AST-based analysis prevents infinite loops

### 📊 Smart Interface
- **Auto-sizing Columns**: Columns automatically adjust to content width
- **Manual Resize**: Use `=` to auto-resize current column, `+` for all columns
- **Scrolling Viewport**: Navigate large spreadsheets smoothly
- **Visual Selection**: Clear indication of current cell
- **Status Messages**: Real-time feedback for operations

### 💾 File Management
- **JSON Format**: Human-readable `.tshts` files
- **Save/Load**: `Ctrl+S` to save, `Ctrl+O` to load
- **Error Handling**: Graceful handling of file operations
- **Auto-backup**: Preserves data integrity

## 🎯 For Early Adopters

### Why Choose TSHTS?

**Performance**: Built in Rust for good performance and memory efficiency. Ideal for spreadsheet tasks in terminal environments.

**Portability**: Runs anywhere Rust runs - Linux, macOS, Windows. No GUI dependencies, perfect for servers and remote work.

**Developer-Friendly**: Clean architecture with comprehensive documentation and tests. Easy to extend and customize.

**Modern Workflow**: Integrates seamlessly with version control, automation scripts, and command-line workflows.

### Current Capabilities

TSHTS already supports the core functionality needed for most spreadsheet tasks:

- ✅ Full formula evaluation with 15+ operators and functions
- ✅ Cell references and range operations
- ✅ File persistence with JSON format
- ✅ Responsive terminal UI with keyboard shortcuts
- ✅ Comprehensive error handling and user feedback
- ✅ Auto-sizing and manual column width adjustment

### Roadmap & Contributing

We're actively developing TSHTS with these upcoming features:

- 📅 **Charts & Visualization**: Basic terminal-based charts
- 📅 **Import/Export**: CSV, Excel format support
- 📅 **Scripting**: Lua/Python integration for custom functions
- 📅 **Collaboration**: Real-time sharing capabilities
- 📅 **Plugins**: Extension system for custom functionality

**Want to contribute?** Check our [issues](https://github.com/yourusername/tshts/issues) for good first contributions. We welcome:
- Bug reports and feature requests
- Documentation improvements
- Performance optimizations
- New formula functions
- Platform-specific enhancements

## 📖 Formula Reference

### Basic Arithmetic
```
=2+3          → 8
=A1*B1        → Multiplies values in A1 and B1
=2**3         → 8 (2 to the power of 3)
=10%3         → 1 (10 modulo 3)
```

### Cell References
```
=A1           → Value from cell A1
=A1+B1        → Sum of A1 and B1
=SUM(A1:A10)  → Sum of range A1 through A10
=AVERAGE(B1:B5) → Average of B1 through B5
```

### Functions
```
=SUM(A1,B1,C1)        → Sum of individual cells
=AVERAGE(A1:A10)      → Average of range
=MIN(A1:C3)           → Minimum value in range
=MAX(A1:C3)           → Maximum value in range
=IF(A1>10,1,0)        → Conditional logic
=AND(A1>0,B1<10)      → Logical AND
=OR(A1=0,B1=0)        → Logical OR
=NOT(A1>5)            → Logical NOT
```

### Comparisons
```
=A1<B1        → 1 if A1 less than B1, 0 otherwise
=A1>=B1       → 1 if A1 greater than or equal to B1
=A1<>B1       → 1 if A1 not equal to B1
```

## ⌨️ Keyboard Shortcuts

### Navigation
- **Arrow Keys** / **hjkl**: Move cell selection
- **Page Up/Down**: Fast scrolling
- **Home**: Jump to cell A1

### Editing
- **Enter** / **F2**: Start editing current cell
- **Esc**: Cancel editing
- **Enter**: Confirm editing

### File Operations
- **Ctrl+S**: Save spreadsheet
- **Ctrl+O**: Load spreadsheet
- **q**: Quit application (in normal mode)

### View
- **F1** / **?**: Show/hide help
- **=**: Auto-resize current column
- **+**: Auto-resize all columns
- **-** / **_**: Manually adjust column width

## 🔧 Configuration

TSHTS uses sensible defaults but can be customized:

### Default Settings
- **Grid Size**: 100 rows × 26 columns
- **Column Width**: 8 characters (auto-adjusting)
- **File Format**: JSON (.tshts extension)
- **Default Filename**: `spreadsheet.tshts`

### File Format
TSHTS saves files in a clean JSON format that's both human-readable and version-control friendly:

```json
{
  "cells": [
    [0, 0, {"value": "Hello", "formula": null}],
    [1, 1, {"value": "42", "formula": "=6*7"}]
  ],
  "rows": 100,
  "cols": 26,
  "column_widths": {"0": 15},
  "default_column_width": 8
}
```

## 🏗️ Architecture

TSHTS follows clean architecture principles:

```
src/
├── domain/          # Core business logic
│   ├── models.rs    # Data structures (Spreadsheet, CellData)
│   └── services.rs  # Formula evaluation engine
├── application/     # Application state management
│   └── state.rs     # App state and mode handling
├── infrastructure/ # External integrations
│   └── persistence.rs # File I/O operations
├── presentation/   # User interface
│   ├── ui.rs       # Terminal rendering
│   └── input.rs    # Keyboard input handling
└── main.rs         # Application entry point
```

This modular design makes TSHTS:
- **Testable**: Each layer can be tested independently
- **Maintainable**: Clear separation of concerns
- **Extensible**: Easy to add new features
- **Portable**: Domain logic independent of UI framework

## 🧪 Testing

TSHTS has comprehensive test coverage:

```bash
# Run all tests
cargo test

# Run with coverage
cargo test --release

# Run specific test modules
cargo test domain::
cargo test formula_evaluator
```

Test categories:
- **Unit Tests**: Individual component functionality
- **Integration Tests**: File I/O and persistence
- **Formula Tests**: Expression evaluation correctness
- **UI Tests**: Application state management

## 📋 System Requirements

- **Operating System**: Linux, macOS, Windows
- **Terminal**: Any terminal with basic cursor support
- **Rust**: 1.70+ (for building from source)
- **Memory**: Small memory footprint
- **Storage**: Minimal (files are compressed JSON)

## 🤝 Getting Help

- **In-App Help**: Press `F1` or `?` for comprehensive help
- **Issues**: [GitHub Issues](https://github.com/yourusername/tshts/issues)
- **Discussions**: [GitHub Discussions](https://github.com/yourusername/tshts/discussions)
- **Documentation**: This README and inline code documentation

## 📜 License

MIT License - see [LICENSE](LICENSE) file for details.

## 🙏 Acknowledgments

Built with:
- [Ratatui](https://github.com/ratatui-org/ratatui) - Terminal UI framework
- [Crossterm](https://github.com/crossterm-rs/crossterm) - Cross-platform terminal manipulation
- [Serde](https://github.com/serde-rs/serde) - Serialization framework

Special thanks to the Rust community for creating an ecosystem that makes building fast, reliable terminal applications a joy.

---

**Ready to supercharge your terminal workflow?** Give TSHTS a try and experience the power of spreadsheets without leaving your command line!
