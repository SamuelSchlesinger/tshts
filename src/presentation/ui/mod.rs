//! Terminal-UI rendering layer (ratatui).
//!
//! mod.rs has the top-level dispatcher (`render_ui`) plus the
//! conditional-format style merger and the TerminalColor → ratatui::Color
//! adapter. Concern-specific rendering lives in submodules:
//!   - header   — top-row header + formula bar
//!   - grid     — spreadsheet body (rows, cells, validation)
//!   - status_bar — mode chip, status line, vim pending preview
//!   - popups   — help / chart / autocomplete / recent-files / suggestions
//!   - help     — help text body, section-jump anchors, search helpers

#![allow(unused_imports)]
use crate::application::{App, AppMode};
use crate::domain::{CellStyle, Spreadsheet, TerminalColor, format_cell_value};
use ratatui::{
    layout::{Constraint, Direction, Layout, Rect},
    style::{Color, Modifier, Style},
    text::{Line, Span},
    widgets::{Block, Borders, Cell, Clear, Paragraph, Row, Table, Tabs},
    Frame,
};

mod header;
mod grid;
mod status_bar;
mod popups;
mod help;

pub(crate) use help::get_help_text;
pub use help::{help_section_offset, find_help_match, HELP_SECTIONS};
pub use grid::cell_at;

// Bring submodule render functions into scope for the dispatcher below.
use header::{render_header, render_formula_bar};
use grid::render_spreadsheet;
use status_bar::{render_status_bar, mode_label, render_vim_pending};
use popups::{
    render_chart_popup, render_command_suggestions,
    render_function_autocomplete, render_recent_files,
};
use help::render_help_popup;

fn layer_for_render(base: CellStyle, cf: CellStyle) -> CellStyle {
    CellStyle {
        bold: base.bold || cf.bold,
        underline: base.underline || cf.underline,
        fg_color: cf.fg_color.or(base.fg_color),
        bg_color: cf.bg_color.or(base.bg_color),
    }
}

fn terminal_color_to_ratatui(color: &TerminalColor) -> Color {
    match color {
        TerminalColor::Black => Color::Black,
        TerminalColor::Red => Color::Red,
        TerminalColor::Green => Color::Green,
        TerminalColor::Yellow => Color::Yellow,
        TerminalColor::Blue => Color::Blue,
        TerminalColor::Magenta => Color::Magenta,
        TerminalColor::Cyan => Color::Cyan,
        TerminalColor::White => Color::White,
        TerminalColor::DarkGray => Color::DarkGray,
        TerminalColor::LightRed => Color::LightRed,
        TerminalColor::LightGreen => Color::LightGreen,
        TerminalColor::LightYellow => Color::LightYellow,
        TerminalColor::LightBlue => Color::LightBlue,
        TerminalColor::LightMagenta => Color::LightMagenta,
        TerminalColor::LightCyan => Color::LightCyan,
    }
}

pub fn render_ui(f: &mut Frame, app: &mut App) {
    let chunks = Layout::default()
        .direction(Direction::Vertical)
        .constraints([
            Constraint::Length(1), // header (filename, sheet tabs, filter status)
            Constraint::Length(1), // formula bar
            Constraint::Min(0),    // spreadsheet grid
            Constraint::Length(3), // status / mode line
        ])
        .split(f.area());

    render_header(f, app, chunks[0]);
    render_formula_bar(f, app, chunks[1]);
    render_spreadsheet(f, app, chunks[2]);
    render_status_bar(f, app, chunks[3]);

    if matches!(app.mode, AppMode::Help) {
        render_help_popup(f, app);
    }
    if matches!(app.mode, AppMode::CommandPalette) {
        render_command_suggestions(f, app, chunks[3]);
    }
    if matches!(app.mode, AppMode::Editing) {
        render_function_autocomplete(f, app, chunks[3]);
    }
    if matches!(app.mode, AppMode::LoadFile) {
        render_recent_files(f, chunks[3]);
    }
    if let Some(_) = &app.chart_popup {
        render_chart_popup(f, app);
    }
}

fn caret(s: &str, char_pos: usize) -> String {
    let mut out = String::with_capacity(s.len() + 3);
    for (i, c) in s.chars().enumerate() {
        if i == char_pos {
            out.push('\u{2581}'); // ▁ — low-line underscore, renders as caret
        }
        out.push(c);
    }
    if char_pos >= s.chars().count() {
        out.push('\u{2581}');
    }
    out
}

#[cfg(test)]
mod help_tests {
    use super::*;

    #[test]
    fn each_section_anchor_exists() {
        let text = get_help_text();
        for (key, title) in HELP_SECTIONS {
            assert!(
                text.lines().any(|l| l.trim() == *title),
                "section {} ('{}') missing from help text",
                key,
                title
            );
        }
    }

    #[test]
    fn help_section_offset_returns_valid_line() {
        for (key, _) in HELP_SECTIONS {
            let off = help_section_offset(*key);
            assert!(off.is_some(), "section key '{}' returned None", key);
        }
    }

    #[test]
    fn help_search_finds_known_token() {
        // "VLOOKUP" lives in the lookup section.
        let found = find_help_match("VLOOKUP", 0);
        assert!(found.is_some());
    }

    #[test]
    fn help_search_wraps_when_no_match_below() {
        // Tokens near the top should be findable from a high offset.
        let found = find_help_match("BASIC", 9999);
        assert!(found.is_some());
    }

    // ----- Agent probes: mode_label & render_vim_pending -----
    // These exercise the private helpers used by the status bar.

    use crate::application::{App, AppMode, VimOperator, VisualKind};

    #[test]
    fn agent_mode_label_exhaustive_nonempty() {
        // Construct each variant and verify mode_label returns a non-empty,
        // distinguishable chip.
        let cases: Vec<(&str, AppMode)> = vec![
            ("Normal", AppMode::Normal),
            ("Editing", AppMode::Editing),
            ("Visual/Cell", AppMode::Visual { kind: VisualKind::Cell }),
            ("Visual/Row", AppMode::Visual { kind: VisualKind::Row }),
            ("Visual/Block", AppMode::Visual { kind: VisualKind::Block }),
            ("Help", AppMode::Help),
            ("SaveAs", AppMode::SaveAs),
            ("LoadFile", AppMode::LoadFile),
            ("ExportCsv", AppMode::ExportCsv),
            ("ImportCsv", AppMode::ImportCsv),
            ("Search", AppMode::Search),
            ("GoToCell", AppMode::GoToCell),
            ("FindReplace", AppMode::FindReplace),
            ("CommandPalette", AppMode::CommandPalette),
            ("ConfirmDiscard", AppMode::ConfirmDiscard),
        ];
        let mut labels = Vec::new();
        for (name, m) in &cases {
            let lbl = mode_label(m);
            assert!(!lbl.is_empty(), "mode_label empty for {}", name);
            labels.push(lbl);
        }
        // All visual variants should give distinct labels (a smoke test for
        // requirement #10).
        let v_cell = mode_label(&AppMode::Visual { kind: VisualKind::Cell });
        let v_row = mode_label(&AppMode::Visual { kind: VisualKind::Row });
        let v_block = mode_label(&AppMode::Visual { kind: VisualKind::Block });
        assert_ne!(v_cell, v_row);
        assert_ne!(v_cell, v_block);
        assert_ne!(v_row, v_block);
    }

    #[test]
    fn agent_vim_pending_format() {
        let mut app = App::default();
        // Idle: empty.
        assert_eq!(render_vim_pending(&app), "");
        // Just a count: " [5]"
        app.vim_count = Some(5);
        assert_eq!(render_vim_pending(&app), " [5]");
        // Count + op
        app.vim_pending_op = Some(VimOperator::Delete);
        assert_eq!(render_vim_pending(&app), " [5d]");
        // Count + op + awaiting-g
        app.vim_awaiting_g = true;
        assert_eq!(render_vim_pending(&app), " [5dg]");
        // Just awaiting g
        app.vim_count = None;
        app.vim_pending_op = None;
        assert_eq!(render_vim_pending(&app), " [g]");
    }
}

