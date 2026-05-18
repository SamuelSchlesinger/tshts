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

pub(super) fn render_chart_popup(f: &mut Frame, app: &App) {
    let Some(chart) = app.chart_popup.as_ref() else { return; };
    let area = f.area();
    let popup = Rect {
        x: area.width / 6,
        y: area.height / 6,
        width: area.width * 2 / 3,
        height: area.height * 2 / 3,
    };
    f.render_widget(Clear, popup);
    let inner_h = popup.height.saturating_sub(3) as usize;
    let inner_w = popup.width.saturating_sub(4) as usize;
    // Pull values fresh from the source range so the chart reflects any
    // edits since the popup was opened.
    let ((sr, sc), (er, ec)) = chart.source;
    let mut values: Vec<f64> = Vec::with_capacity((er - sr + 1) * (ec - sc + 1));
    let sheet = app.workbook.current_sheet();
    for r in sr..=er {
        for c in sc..=ec {
            values.push(sheet.get_cell(r, c).value.parse::<f64>().unwrap_or(0.0));
        }
    }
    let min = values.iter().cloned().fold(f64::INFINITY, f64::min);
    let max = values.iter().cloned().fold(f64::NEG_INFINITY, f64::max);
    let range = if max > min { max - min } else { 1.0 };

    let body = match chart.kind {
        crate::application::ChartKind::Bar | crate::application::ChartKind::Sparkline => {
            // Vertical bar chart: one column per value, scaled to inner_h rows.
            let blocks = ['\u{2581}', '\u{2582}', '\u{2583}', '\u{2584}', '\u{2585}', '\u{2586}', '\u{2587}', '\u{2588}'];
            let n = values.len();
            let col_w = (inner_w / n.max(1)).max(1);
            let mut rows_buf: Vec<Vec<char>> = vec![vec![' '; inner_w]; inner_h];
            for (i, v) in values.iter().enumerate() {
                let height = ((v - min) / range * inner_h as f64) as usize;
                for r in 0..height {
                    let block_idx = if r == height - 1 && height > 0 {
                        (((v - min) / range * (inner_h * 8) as f64) as usize % 8).min(7)
                    } else {
                        7
                    };
                    let row = inner_h - 1 - r;
                    let start = i * col_w;
                    for c in 0..col_w.min(2) {
                        if start + c < inner_w {
                            rows_buf[row][start + c] = blocks[block_idx];
                        }
                    }
                }
            }
            rows_buf
                .into_iter()
                .map(|row| row.into_iter().collect::<String>())
                .collect::<Vec<_>>()
                .join("\n")
        }
        crate::application::ChartKind::Line => {
            // Simple ASCII line plot.
            let n = values.len();
            let mut grid: Vec<Vec<char>> = vec![vec![' '; inner_w]; inner_h];
            for (i, v) in values.iter().enumerate() {
                let x = if n > 1 {
                    (i * (inner_w - 1)) / (n - 1)
                } else {
                    0
                };
                let y_norm = (v - min) / range;
                let y = inner_h - 1 - (y_norm * (inner_h - 1) as f64) as usize;
                if y < inner_h && x < inner_w {
                    grid[y][x] = '\u{2022}';
                }
            }
            grid.into_iter()
                .map(|r| r.into_iter().collect::<String>())
                .collect::<Vec<_>>()
                .join("\n")
        }
    };

    let widget = Paragraph::new(body)
        .block(
            Block::default()
                .borders(Borders::ALL)
                .title(format!("{} (min={:.2}, max={:.2})", chart.title, min, max))
                .style(Style::default().fg(Color::Cyan)),
        )
        .style(Style::default().fg(Color::Green));
    f.render_widget(widget, popup);
}

pub(super) fn render_function_autocomplete(f: &mut Frame, app: &App, status_area: Rect) {
    // Find the partial identifier the user is typing right before the cursor.
    let pre: String = app.input.chars().take(app.cursor_position).collect();
    // Walk backwards over alpha/underscore chars to get the current token.
    let token: String = pre
        .chars()
        .rev()
        .take_while(|c| c.is_ascii_alphabetic() || *c == '_' || *c == '.')
        .collect::<String>()
        .chars()
        .rev()
        .collect();
    if token.len() < 2 {
        return;
    }
    let upper = token.to_uppercase();
    let matches: Vec<&'static str> = crate::domain::builtin_function_names()
        .into_iter()
        .filter(|n| n.starts_with(&upper))
        .take(6)
        .collect();
    if matches.is_empty() {
        return;
    }
    let width = (matches.iter().map(|s| s.len()).max().unwrap_or(20) as u16 + 4).max(15);
    let height = (matches.len() as u16 + 2).min(8);
    let popup = Rect {
        x: status_area.x,
        y: status_area.y.saturating_sub(height),
        width: width.min(status_area.width),
        height,
    };
    f.render_widget(Clear, popup);
    let text = matches.join("\n");
    let widget = Paragraph::new(text)
        .block(
            Block::default()
                .borders(Borders::ALL)
                .title("functions")
                .style(Style::default().fg(Color::Cyan)),
        )
        .style(Style::default().fg(Color::Cyan).bg(Color::Black));
    f.render_widget(widget, popup);
}

pub(super) fn render_recent_files(f: &mut Frame, status_area: Rect) {
    let recents = crate::infrastructure::recent::load();
    if recents.is_empty() {
        return;
    }
    let max_show = recents.len().min(8);
    let shown = &recents[..max_show];
    let width = (shown.iter().map(|s| s.len()).max().unwrap_or(20) as u16 + 4).max(30);
    let height = (shown.len() as u16 + 2).min(10);
    let popup = Rect {
        x: status_area.x,
        y: status_area.y.saturating_sub(height),
        width: width.min(status_area.width),
        height,
    };
    f.render_widget(Clear, popup);
    let text = shown.join("\n");
    let widget = Paragraph::new(text)
        .block(
            Block::default()
                .borders(Borders::ALL)
                .title("recent — Tab to cycle")
                .style(Style::default().fg(Color::Cyan)),
        )
        .style(Style::default().fg(Color::White).bg(Color::Black));
    f.render_widget(widget, popup);
}

pub(super) fn render_command_suggestions(f: &mut Frame, app: &App, status_area: Rect) {
    // (read-only access here; mutable on the status bar to write the cache.)
    let suggestions = app.command_suggestions(8);
    if suggestions.is_empty() {
        return;
    }
    let height = (suggestions.len() as u16 + 2).min(10);
    let width = (suggestions.iter().map(|s| s.len()).max().unwrap_or(20) as u16 + 4).max(20);
    let popup = Rect {
        x: status_area.x,
        y: status_area.y.saturating_sub(height),
        width: width.min(status_area.width),
        height,
    };
    f.render_widget(Clear, popup);
    let text = suggestions.join("\n");
    let widget = Paragraph::new(text)
        .block(Block::default().borders(Borders::ALL).title(":suggestions"))
        .style(Style::default().fg(Color::Cyan).bg(Color::Black));
    f.render_widget(widget, popup);
}

