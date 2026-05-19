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

pub(super) fn render_formula_bar(f: &mut Frame, app: &App, area: Rect) {
    let cell = app.workbook.current_sheet().get_cell(app.selected_row, app.selected_col);
    let location = if app.r1c1_mode {
        format!("R{}C{}", app.selected_row + 1, app.selected_col + 1)
    } else {
        format!(
            "{}{}",
            crate::domain::Spreadsheet::column_label(app.selected_col),
            app.selected_row + 1
        )
    };
    let content = if let Some(formula) = cell.formula {
        formula
    } else {
        cell.value
    };
    let text = format!("  {} │ {}", location, content);
    let bar = Paragraph::new(text).style(Style::default().fg(Color::White).bg(Color::DarkGray));
    f.render_widget(bar, area);
}

pub(super) fn render_header(f: &mut Frame, app: &App, area: Rect) {
    let mut tabs = String::new();
    for (i, name) in app.workbook.sheet_names.iter().enumerate() {
        if i == app.workbook.active_sheet {
            tabs.push_str(&format!("[{}]", name));
        } else {
            tabs.push_str(&format!(" {} ", name));
        }
        if i < app.workbook.sheet_names.len() - 1 {
            tabs.push('|');
        }
    }
    let filter_indicator = if let Some(col) = app.filter_column {
        let label = Spreadsheet::column_label(col);
        let hidden = app.hidden_rows.len();
        if let Some(ref v) = app.filter_value {
            format!(" | Filter: {}~\"{}\" ({} hidden)", label, v, hidden)
        } else {
            format!(" | Filter: {} ({} hidden)", label, hidden)
        }
    } else {
        String::new()
    };
    let dirty_marker = if app.dirty { " *" } else { "" };
    let validation_indicator = if app.validations.is_empty() {
        String::new()
    } else {
        let mut cols: Vec<usize> = app.validations.keys().copied().collect();
        cols.sort();
        let names: Vec<String> = cols
            .iter()
            .map(|c| Spreadsheet::column_label(*c))
            .collect();
        format!(" | Validate: {}", names.join(","))
    };
    let hidden_cols_indicator = if app.hidden_cols.is_empty() {
        String::new()
    } else {
        format!(" | Hidden cols: {}", app.hidden_cols.len())
    };
    let iter_indicator = if app.iterative_calc {
        " | Iter".to_string()
    } else {
        String::new()
    };

    let header = Paragraph::new(format!(
        "tshts{} | {}{}{}{}{}",
        dirty_marker,
        tabs,
        filter_indicator,
        validation_indicator,
        hidden_cols_indicator,
        iter_indicator,
    ))
    .style(Style::default().fg(Color::Cyan));
    f.render_widget(header, area);
}


pub(super) fn render_sheet_tabs(f: &mut Frame, app: &App, area: Rect) {
    let titles: Vec<Line> = app
        .workbook
        .sheet_names
        .iter()
        .enumerate()
        .map(|(i, name)| {
            let is_active = i == app.workbook.active_sheet;
            let prefix = Span::styled(" ", Style::default());
            let label = if is_active {
                Span::styled(
                    name.to_string(),
                    Style::default()
                        .fg(Color::Black)
                        .bg(Color::LightCyan)
                        .add_modifier(Modifier::BOLD),
                )
            } else {
                Span::styled(name.to_string(), Style::default().fg(Color::Gray))
            };
            Line::from(vec![prefix, label, Span::raw(" ")])
        })
        .collect();

    let tabs = Tabs::new(titles)
        .select(app.workbook.active_sheet)
        .divider(Span::styled("│", Style::default().fg(Color::DarkGray)))
        .style(Style::default())
        .highlight_style(Style::default());
    f.render_widget(tabs, area);
}
