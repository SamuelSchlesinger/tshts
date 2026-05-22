//! Submodule of `ui` — see ui/mod.rs.

#![allow(unused_imports)]
use crate::application::{App, AppMode};
use crate::domain::{CellStyle, NumberFormat, Spreadsheet, TerminalColor, format_cell_value};
use ratatui::{
    layout::{Constraint, Direction, Layout, Rect},
    style::{Color, Modifier, Style},
    text::{Line, Span},
    widgets::{Block, Borders, Cell, Clear, Paragraph, Row, Table, Tabs},
    Frame,
};
use super::{caret, layer_for_render, terminal_color_to_ratatui};
use unicode_width::UnicodeWidthChar;

pub(super) fn render_status_bar(f: &mut Frame, app: &mut App, area: Rect) {
    // Pre-compute the stats up-front (mutable borrow) so we can hold immutable
    // borrows of app fields below for the format string.
    let stats = app.get_selection_stats();
    let mode_label = mode_label(&app.mode);
    // The pending-op / count preview shown vim-style on the right.
    let pending = render_vim_pending(app);
    let input_text = match app.mode {
        AppMode::Normal => {
            if let Some(ref status) = app.status_message {
                format!("{} {}", mode_label, status)
            } else {
                let filename = app.filename.as_deref().unwrap_or("unsaved");
                let comment_info = {
                    let cell = app.workbook.current_sheet().get_cell(app.selected_row, app.selected_col);
                    if let Some(ref comment) = cell.comment {
                        // Truncate by display width so a long comment can't
                        // overflow the status bar and corrupt rendering.
                        // Newlines also break the single-line bar, so collapse
                        // them. unicode-width gives us proper width for
                        // CJK / emoji.
                        const MAX_WIDTH: usize = 40;
                        let mut width = 0;
                        let mut truncated = String::new();
                        let mut clipped = false;
                        for ch in comment.chars() {
                            if ch == '\n' || ch == '\r' {
                                truncated.push(' ');
                                width += 1;
                            } else {
                                let w = UnicodeWidthChar::width(ch).unwrap_or(0);
                                if width + w > MAX_WIDTH {
                                    clipped = true;
                                    break;
                                }
                                truncated.push(ch);
                                width += w;
                            }
                        }
                        if clipped {
                            truncated.push('…');
                        }
                        format!(" | Comment: {}", truncated)
                    } else {
                        String::new()
                    }
                };
                format!(
                    "{} File: {}{} | :w save | :q quit | F1/? help{}",
                    mode_label, filename, comment_info, pending
                )
            }
        }
        AppMode::Visual { .. } => {
            let selection_info = if let Some(((start_row, start_col), (end_row, end_col))) = app.get_selection_range() {
                let rows = end_row - start_row + 1;
                let cols = end_col - start_col + 1;
                let stats_str = if let Some((sum, avg, count)) = stats {
                    format!(" | SUM={} AVG={:.2} COUNT={}", sum, avg, count)
                } else {
                    String::new()
                };
                format!(" {}x{} cells{}", rows, cols, stats_str)
            } else {
                String::new()
            };
            format!("{}{} | y yank · d delete · c change · Esc cancel{}", mode_label, selection_info, pending)
        }
        AppMode::Editing => format!(
            "{} {} (Enter save · Esc cancel · Tab next col)",
            mode_label,
            caret(&app.input, app.cursor_position)
        ),
        AppMode::Help => format!("{} ↑↓/jk: scroll | PgUp/PgDn: fast scroll | Home: top | Esc/q: close help", mode_label),
        AppMode::SaveAs => format!(
            "{} Save as: {} (Enter save · Esc cancel)",
            mode_label,
            caret(&app.filename_input, app.cursor_position)
        ),
        AppMode::LoadFile => format!(
            "{} Load file: {} (Enter load · Esc cancel)",
            mode_label,
            caret(&app.filename_input, app.cursor_position)
        ),
        AppMode::ExportCsv => format!(
            "{} Export CSV as: {} (Enter export · Esc cancel)",
            mode_label,
            caret(&app.filename_input, app.cursor_position)
        ),
        AppMode::ImportCsv => format!(
            "{} Import CSV from: {} (Enter import · Esc cancel)",
            mode_label,
            caret(&app.filename_input, app.cursor_position)
        ),
        AppMode::Search => {
            let results_info = if app.search_results.is_empty() {
                if app.search_query.is_empty() {
                    "".to_string()
                } else {
                    " (no results)".to_string()
                }
            } else {
                format!(" ({}/{} results)", app.search_result_index + 1, app.search_results.len())
            };
            format!(
                "{} Search: {}{} (Enter finish · Esc cancel · ↑↓ navigate)",
                mode_label,
                caret(&app.search_query, app.cursor_position),
                results_info
            )
        }
        AppMode::GoToCell => format!(
            "{} Go to cell: {} (Enter go · Esc cancel)",
            mode_label,
            caret(&app.goto_cell_input, app.cursor_position)
        ),
        AppMode::FindReplace => {
            let field = if app.find_replace_on_replace { "Replace" } else { "Find" };
            let results_info = if app.find_replace_results.is_empty() {
                if app.find_replace_search.is_empty() { String::new() } else { " (no results)".to_string() }
            } else {
                format!(" ({}/{})", app.find_replace_index + 1, app.find_replace_results.len())
            };
            format!("{} [{}] Find: {} | Replace: {}{} (Tab: switch · Enter: do · Esc: close)",
                mode_label, field, app.find_replace_search, app.find_replace_replace, results_info)
        }
        AppMode::CommandPalette => {
            format!(
                "{} :{} (Enter execute · Esc cancel)",
                mode_label,
                caret(&app.command_input, app.cursor_position)
            )
        }
        AppMode::ConfirmDiscard => {
            format!(
                "{} {}",
                mode_label,
                app.status_message.clone().unwrap_or_else(|| "Confirm?".to_string())
            )
        }
    };

    let input = Paragraph::new(input_text)
        .block(Block::default().borders(Borders::ALL).title("Status"))
        .style(match app.mode {
            AppMode::Normal => Style::default(),
            AppMode::Editing => Style::default().fg(Color::Green),
            AppMode::Visual { .. } => Style::default().fg(Color::LightMagenta),
            AppMode::Help => Style::default().fg(Color::Cyan),
            AppMode::SaveAs => Style::default().fg(Color::Yellow),
            AppMode::LoadFile => Style::default().fg(Color::Yellow),
            AppMode::ExportCsv => Style::default().fg(Color::Magenta),
            AppMode::ImportCsv => Style::default().fg(Color::Green),
            AppMode::Search => Style::default().fg(Color::LightYellow),
            AppMode::GoToCell => Style::default().fg(Color::Cyan),
            AppMode::FindReplace => Style::default().fg(Color::LightYellow),
            AppMode::CommandPalette => Style::default().fg(Color::Magenta),
            AppMode::ConfirmDiscard => Style::default().fg(Color::Red),
        });
    f.render_widget(input, area);
}

pub(super) fn mode_label(mode: &AppMode) -> &'static str {
    match mode {
        AppMode::Normal => "-- NORMAL --",
        AppMode::Editing => "-- INSERT --",
        AppMode::Visual { kind: crate::application::VisualKind::Cell } => "-- VISUAL --",
        AppMode::Visual { kind: crate::application::VisualKind::Row } => "-- VISUAL LINE --",
        AppMode::Visual { kind: crate::application::VisualKind::Block } => "-- VISUAL BLOCK --",
        AppMode::Help => "-- HELP --",
        AppMode::SaveAs => "-- SAVE --",
        AppMode::LoadFile => "-- OPEN --",
        AppMode::ExportCsv => "-- CSV EXPORT --",
        AppMode::ImportCsv => "-- CSV IMPORT --",
        AppMode::Search => "-- SEARCH --",
        AppMode::GoToCell => "-- GOTO --",
        AppMode::FindReplace => "-- FIND/REPLACE --",
        AppMode::CommandPalette => "-- COMMAND --",
        AppMode::ConfirmDiscard => "-- CONFIRM --",
    }
}

pub(super) fn render_vim_pending(app: &App) -> String {
    let count = app.vim_count.map(|c| c.to_string()).unwrap_or_default();
    let op = app.vim_pending_op.map(|o| match o {
        crate::application::VimOperator::Delete => "d",
        crate::application::VimOperator::Yank => "y",
        crate::application::VimOperator::Change => "c",
    }).unwrap_or("");
    let g = if app.vim_awaiting_g { "g" } else { "" };
    let buf = format!("{}{}{}", count, op, g);
    if buf.is_empty() { String::new() } else { format!(" [{}]", buf) }
}

