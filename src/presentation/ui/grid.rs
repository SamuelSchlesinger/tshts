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

pub(super) fn build_row<'a>(app: &App, row: usize, visible_cols: usize, is_frozen: bool) -> Row<'a> {
    let row_number_style = if row == app.selected_row {
        Style::default().bg(Color::LightBlue).fg(Color::Black)
    } else if is_frozen {
        Style::default().fg(Color::Yellow).add_modifier(Modifier::BOLD)
    } else {
        Style::default().fg(Color::Yellow)
    };
    let mut cells = vec![Cell::from(format!("{}", row + 1)).style(row_number_style)];

    // Build the list of columns to render: frozen first, then scrolled body.
    // Hidden columns are skipped.
    let total_cols = app.workbook.current_sheet().cols;
    let frozen_c = app.frozen_cols.min(total_cols);
    let mut col_list: Vec<usize> = (0..frozen_c)
        .filter(|c| !app.hidden_cols.contains(c))
        .collect();
    let body_start = app.scroll_col.max(frozen_c);
    let mut shown = col_list.len();
    let mut c = body_start;
    while shown < visible_cols && c < total_cols {
        if !app.hidden_cols.contains(&c) {
            col_list.push(c);
            shown += 1;
        }
        c += 1;
    }

    for col in col_list.into_iter() {
        let col_is_frozen = col < frozen_c;
        let cell_data = app.workbook.current_sheet().get_cell(row, col);
        let has_comment = cell_data.comment.is_some();
        let has_formula = cell_data.formula.is_some();
        let raw_value = cell_data.value.clone();
        let cell_value = if cell_data.value.is_empty() {
            if has_comment {
                "\u{25c6}".to_string()
            } else if has_formula {
                "\u{00b7}".to_string()
            } else {
                " ".to_string()
            }
        } else {
            let val = if let Some(ref fmt) = cell_data.format {
                format_cell_value(&cell_data.value, fmt)
            } else {
                cell_data.value
            };
            let mut s = val;
            if has_formula {
                s.push('\u{00b7}');
            }
            if has_comment {
                s.push('\u{25c6}');
            }
            s
        };

        // Start with the cell's own style, then layer any conditional-format
        // rule on top (rules win when they fire).
        let cell_style = cell_data.format.as_ref().map(|f| f.style.clone());
        let cf_style = app.workbook.current_sheet().conditional_style_for(row, col);
        let effective = match (cell_style, cf_style) {
            (Some(base), Some(cf)) => Some(layer_for_render(base, cf)),
            (Some(s), None) => Some(s),
            (None, Some(cf)) => Some(cf),
            (None, None) => None,
        };
        let base_style = if let Some(ref s) = effective {
            let mut style = Style::default();
            let mut modifiers = Modifier::empty();
            if s.bold {
                modifiers |= Modifier::BOLD;
            }
            if s.underline {
                modifiers |= Modifier::UNDERLINED;
            }
            if !modifiers.is_empty() {
                style = style.add_modifier(modifiers);
            }
            if let Some(ref fg) = s.fg_color {
                style = style.fg(terminal_color_to_ratatui(fg));
            }
            if let Some(ref bg) = s.bg_color {
                style = style.bg(terminal_color_to_ratatui(bg));
            }
            style
        } else {
            Style::default()
        };

        let is_error = raw_value.starts_with('#') && raw_value.ends_with(['!', '?']);
        let is_hyperlink =
            raw_value.starts_with("http://") || raw_value.starts_with("https://");

        let mut style = if row == app.selected_row && col == app.selected_col {
            base_style.bg(Color::Blue).fg(Color::White)
        } else if app.is_cell_selected(row, col) {
            base_style.bg(Color::LightBlue).fg(Color::Black)
        } else if app.search_results.contains(&(row, col)) {
            if matches!(app.mode, AppMode::Search)
                && app.search_results.get(app.search_result_index) == Some(&(row, col))
            {
                base_style.bg(Color::Yellow).fg(Color::Black)
            } else {
                base_style.bg(Color::DarkGray).fg(Color::White)
            }
        } else if is_frozen || col_is_frozen {
            base_style.bg(Color::Rgb(40, 40, 60))
        } else {
            base_style
        };
        if is_error {
            style = style.fg(Color::Red).add_modifier(Modifier::BOLD);
        } else if is_hyperlink && !(row == app.selected_row && col == app.selected_col) {
            style = style.fg(Color::Cyan).add_modifier(Modifier::UNDERLINED);
        }
        // Data validation: if this column has a rule and the cell value
        // fails it, mark with a red bottom-modified style.
        if !raw_value.is_empty()
            && let Some(predicate) = app.validations.get(&col)
                && !validation_passes(app, &raw_value, predicate) {
                    style = style.bg(Color::Rgb(80, 20, 20)).fg(Color::White);
                }
        // Spill ghosts render dimmer so the user can tell them apart from
        // editable cells. Doesn't fire on the cursor cell — the cursor
        // highlight wins.
        if cell_data.spill_anchor.is_some()
            && !(row == app.selected_row && col == app.selected_col)
        {
            style = style.add_modifier(Modifier::DIM | Modifier::ITALIC);
        }

        cells.push(Cell::from(cell_value).style(style));
    }
    Row::new(cells).height(1)
}

pub(super) fn validation_passes(app: &App, value: &str, predicate: &str) -> bool {
    let token = if value.parse::<f64>().is_ok() {
        value.to_string()
    } else {
        format!("\"{}\"", value.replace('"', "\"\""))
    };
    let bound = predicate.replace('_', &token);
    let formula = format!("={}", bound);
    let evaluator = crate::domain::FormulaEvaluator::for_workbook(
        &app.workbook,
        app.workbook.current_sheet(),
        &app.workbook.named_ranges,
    );
    let result = evaluator.evaluate_formula(&formula);
    !matches!(result.as_str(), "0" | "FALSE" | "")
        && !result.starts_with('#')
}

pub(super) fn render_spreadsheet(f: &mut Frame, app: &mut App, area: Rect) {
    // Reserve one row for the header. saturating_sub guards against a 0-row
    // area (e.g. terminal resized to a single line) which would otherwise
    // underflow to usize::MAX and crash the next loop iteration.
    let visible_rows = (area.height as usize).saturating_sub(1);

    let mut total_width = 4;
    let mut visible_cols = 0;
    let available_width = area.width as usize;
    let total_cols = app.workbook.current_sheet().cols;
    let frozen_c = app.frozen_cols.min(total_cols);

    // Reserve width for frozen cols first (skipping hidden ones).
    for col in 0..frozen_c {
        if app.hidden_cols.contains(&col) {
            continue;
        }
        let col_width = app.workbook.current_sheet().get_column_width(col);
        if total_width + col_width + 1 > available_width {
            break;
        }
        total_width += col_width + 1;
        visible_cols += 1;
    }
    // Then scrolled body cols.
    let body_start = app.scroll_col.max(frozen_c);
    for col in body_start..total_cols {
        if app.hidden_cols.contains(&col) {
            continue;
        }
        let col_width = app.workbook.current_sheet().get_column_width(col);
        if total_width + col_width + 1 > available_width {
            break;
        }
        total_width += col_width + 1;
        visible_cols += 1;
    }

    app.update_viewport_size(visible_rows, visible_cols);

    // Build the same column list used in build_row so headers line up.
    let mut header_cols: Vec<usize> = (0..frozen_c)
        .filter(|c| !app.hidden_cols.contains(c))
        .collect();
    let mut shown = header_cols.len();
    let mut c = body_start;
    while shown < visible_cols && c < total_cols {
        if !app.hidden_cols.contains(&c) {
            header_cols.push(c);
            shown += 1;
        }
        c += 1;
    }

    let mut headers = vec![Cell::from("")];
    for col in &header_cols {
        let header_style = if *col == app.selected_col {
            Style::default().bg(Color::LightBlue).fg(Color::Black)
        } else if *col < frozen_c {
            Style::default().fg(Color::Yellow).add_modifier(Modifier::BOLD)
        } else {
            Style::default().fg(Color::Yellow)
        };
        let label = if app.r1c1_mode {
            format!("C{}", *col + 1)
        } else {
            Spreadsheet::column_label(*col)
        };
        headers.push(Cell::from(label).style(header_style));
    }

    let header_row = Row::new(headers).height(1);
    let mut rows = vec![header_row];
    let mut rendered_rows = 0;

    // Frozen rows: always rendered at the top regardless of scroll position.
    let frozen = app.frozen_rows.min(app.workbook.current_sheet().rows);
    for frow in 0..frozen {
        if app.hidden_rows.contains(&frow) {
            continue;
        }
        if rendered_rows >= visible_rows {
            break;
        }
        rendered_rows += 1;
        rows.push(build_row(app, frow, visible_cols, true));
    }

    // Scrolled body starts after frozen rows.
    let body_start = app.scroll_row.max(frozen);
    let mut row = body_start;
    while row < app.workbook.current_sheet().rows && rendered_rows < visible_rows {
        if app.hidden_rows.contains(&row) {
            row += 1;
            continue;
        }
        rendered_rows += 1;
        rows.push(build_row(app, row, visible_cols, false));
        row += 1;
    }

    let mut widths = vec![Constraint::Length(4)];
    for col in &header_cols {
        widths.push(Constraint::Length(
            app.workbook.current_sheet().get_column_width(*col) as u16,
        ));
    }

    // Record the rendered column rects so mouse hit-testing can decode
    // clicks even with custom column widths or hidden columns.
    // Layout: 1-cell border, 4-cell row-label, 1-cell spacing, then each
    // column with its width followed by a 1-cell spacing.
    let mut col_rects: Vec<(usize, u16, u16)> = Vec::with_capacity(header_cols.len());
    let mut x_cursor = area.x + 1 /* left border */ + 4 /* row label */ + 1 /* spacing */;
    for col in &header_cols {
        let w = app.workbook.current_sheet().get_column_width(*col) as u16;
        col_rects.push((*col, x_cursor, x_cursor + w));
        x_cursor += w + 1; // +1 for column-spacing
    }
    app.last_col_rects = col_rects;
    // First data-row Y: area.y is the top border, +1 is the column-letter
    // row, +2 is the first data row.
    app.last_grid_top_y = area.y + 2;

    let table = Table::new(rows, widths)
        .block(Block::default().borders(Borders::ALL).title("Spreadsheet"))
        .column_spacing(1);

    f.render_widget(table, area);
}

pub fn cell_at(app: &App, x: usize, y: usize) -> Option<(usize, usize)> {
    let grid_top = app.last_grid_top_y as usize;
    if y < grid_top {
        return None;
    }
    let xs = x as u16;
    let col = app
        .last_col_rects
        .iter()
        .find(|(_, lo, hi)| xs >= *lo && xs < *hi)
        .map(|(c, _, _)| *c)?;

    let row_offset = y - grid_top;
    let frozen_r = app.frozen_rows.min(app.workbook.current_sheet().rows);
    if row_offset < frozen_r {
        return Some((row_offset, col));
    }
    let row = app.scroll_row.max(frozen_r) + (row_offset - frozen_r);
    if row >= app.workbook.current_sheet().rows {
        return None;
    }
    Some((row, col))
}

