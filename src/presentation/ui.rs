use crate::application::{App, AppMode};
use crate::domain::{CellStyle, Spreadsheet, TerminalColor, format_cell_value};

/// Cell style + conditional-format style merger (UI side).
/// Booleans OR; colors prefer conditional override.
fn layer_for_render(base: CellStyle, cf: CellStyle) -> CellStyle {
    CellStyle {
        bold: base.bold || cf.bold,
        underline: base.underline || cf.underline,
        fg_color: cf.fg_color.or(base.fg_color),
        bg_color: cf.bg_color.or(base.bg_color),
    }
}
use ratatui::{
    layout::{Constraint, Direction, Layout, Rect},
    style::{Color, Modifier, Style},
    widgets::{Block, Borders, Cell, Clear, Paragraph, Row, Table},
    Frame,
};

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

fn render_chart_popup(f: &mut Frame, app: &App) {
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

fn render_function_autocomplete(f: &mut Frame, app: &App, status_area: Rect) {
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

fn render_recent_files(f: &mut Frame, status_area: Rect) {
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

fn render_command_suggestions(f: &mut Frame, app: &App, status_area: Rect) {
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

fn render_formula_bar(f: &mut Frame, app: &App, area: Rect) {
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

fn render_header(f: &mut Frame, app: &App, area: Rect) {
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

fn build_row<'a>(app: &App, row: usize, visible_cols: usize, is_frozen: bool) -> Row<'a> {
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
        if !raw_value.is_empty() {
            if let Some(predicate) = app.validations.get(&col) {
                if !validation_passes(app, &raw_value, predicate) {
                    style = style.bg(Color::Rgb(80, 20, 20)).fg(Color::White);
                }
            }
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

/// Evaluate a per-column validation predicate against `value`. Returns false
/// only when the predicate evaluates to a truthy false explicitly. Empty
/// predicate or unparseable formula → passes (don't mark good cells bad).
fn validation_passes(app: &App, value: &str, predicate: &str) -> bool {
    let token = if value.parse::<f64>().is_ok() {
        value.to_string()
    } else {
        format!("\"{}\"", value.replace('"', "\"\""))
    };
    let bound = predicate.replace('_', &token);
    let formula = format!("={}", bound);
    let evaluator = crate::domain::FormulaEvaluator::with_workbook(
        &app.workbook,
        app.workbook.current_sheet(),
        &app.workbook.named_ranges,
    );
    let result = evaluator.evaluate_formula(&formula);
    !matches!(result.as_str(), "0" | "FALSE" | "")
        && !result.starts_with('#')
}

fn render_spreadsheet(f: &mut Frame, app: &mut App, area: Rect) {
    let visible_rows = area.height as usize - 1;

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

/// Map terminal (x, y) to a sheet (row, col) using the rectangles the last
/// frame recorded. Handles custom column widths, hidden columns, and frozen
/// rows/cols correctly.
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

/// Splits a string at `char_pos` and joins with a `▏` marker so the user sees
/// where the cursor is. Used by status-bar input fields and the editing line.
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

fn render_status_bar(f: &mut Frame, app: &mut App, area: Rect) {
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
                let filename = app.filename.as_ref().map(|f| f.as_str()).unwrap_or("unsaved");
                let comment_info = {
                    let cell = app.workbook.current_sheet().get_cell(app.selected_row, app.selected_col);
                    if let Some(ref comment) = cell.comment {
                        format!(" | Comment: {}", comment)
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

/// Returns the vim-style mode chip, e.g. "-- NORMAL --" or "-- VISUAL LINE --".
fn mode_label(mode: &AppMode) -> &'static str {
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

/// Vim-style pending-op preview, e.g. ` [5d]` while the user has typed `5d` and
/// the pending operator is waiting for a motion. Empty string when idle.
fn render_vim_pending(app: &App) -> String {
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

fn render_help_popup(f: &mut Frame, app: &App) {
    let area = f.area();
    let popup_area = Rect {
        x: area.width / 10,
        y: area.height / 10,
        width: area.width * 4 / 5,
        height: area.height * 4 / 5,
    };

    f.render_widget(Clear, popup_area);

    let help_text = get_help_text();
    let help_lines: Vec<&str> = help_text.lines().collect();
    let visible_height = popup_area.height.saturating_sub(2) as usize;
    let scroll = app.help_scroll;
    let start_line = scroll.min(help_lines.len().saturating_sub(visible_height));
    let end_line = (start_line + visible_height).min(help_lines.len());
    let visible_text = help_lines[start_line..end_line].join("\n");

    let title = if app.help_search_active {
        format!(
            "Help (Line {}/{}) — /{}",
            start_line + 1,
            help_lines.len(),
            app.help_search
        )
    } else if !app.help_search.is_empty() {
        format!(
            "Help (Line {}/{}) — /{} (press n for next)",
            start_line + 1,
            help_lines.len(),
            app.help_search
        )
    } else {
        format!(
            "Help (Line {}/{}) — 1-9: jump, /: search, ↑↓: scroll, Esc: close",
            start_line + 1,
            help_lines.len()
        )
    };

    let help_widget = Paragraph::new(visible_text)
        .block(
            Block::default()
                .borders(Borders::ALL)
                .title(title)
                .style(Style::default().fg(Color::Cyan)),
        )
        .style(Style::default().fg(Color::White));
    f.render_widget(help_widget, popup_area);
}

/// Line offset of the section identified by `key` (one of '1'..='9').
pub fn help_section_offset(key: char) -> Option<usize> {
    let text = get_help_text();
    let target = HELP_SECTIONS.iter().find(|(k, _)| *k == key)?.1;
    text.lines().position(|l| l.trim() == target)
}

/// Find the first line at or after `from` containing `needle` (case-insensitive).
pub fn find_help_match(needle: &str, from: usize) -> Option<usize> {
    if needle.is_empty() {
        return None;
    }
    let text = get_help_text();
    let needle_lc = needle.to_lowercase();
    let lines: Vec<&str> = text.lines().collect();
    let len = lines.len();
    // Search from `from` forward, then wrap.
    for i in 0..len {
        let idx = (from + i) % len;
        if lines[idx].to_lowercase().contains(&needle_lc) {
            return Some(idx);
        }
    }
    None
}

/// Section markers used to drive `1-9` jump navigation in the help popup.
/// Each tuple is (digit-key, section title, prefix to search for in the doc).
pub const HELP_SECTIONS: &[(char, &str)] = &[
    ('1', "=== BASIC & OPERATORS ==="),
    ('2', "=== NUMERIC FUNCTIONS ==="),
    ('3', "=== STRING FUNCTIONS ==="),
    ('4', "=== LOGICAL FUNCTIONS ==="),
    ('5', "=== LOOKUP / CONDITIONAL ==="),
    ('6', "=== DATE FUNCTIONS ==="),
    ('7', "=== FILE OPERATIONS ==="),
    ('8', "=== NAVIGATION SHORTCUTS ==="),
    ('9', "=== ADVANCED ==="),
    ('0', "=== VIM MODE ==="),
];

pub(crate) fn get_help_text() -> String {
    r#"TSHTS EXPRESSION LANGUAGE REFERENCE (v0.2)

Press 0-9 to jump to a section (0 = Vim Mode). Press / to search within this help.

=== WHAT'S NEW IN 0.2 ===
• Cross-sheet refs: Sheet2!A1, 'My Sheet'!B5:B10, 3-D Sheet1:Sheet3!A1
• Excel error types: #DIV/0!, #REF!, #VALUE!, #NAME?, #NUM!, #N/A, #NULL!, #SPILL!
• 2-D arrays with shape: =A1:C3 * 2 broadcasts; VLOOKUP/HLOOKUP/XLOOKUP/INDEX
• Array literals: {1,2,3;4,5,6}
• Dynamic arrays: SEQUENCE, FILTER, SORT, UNIQUE, TRANSPOSE, SUMPRODUCT
• LET local bindings, LAMBDA functions, MAP/REDUCE/BYROW/BYCOL/SCAN/MAKEARRAY
• Named LAMBDAs: `:name DOUBLE LAMBDA(x, x*2)` then =DOUBLE(7) → 14
• Tables: `:table create A1:D100 name=Sales` + structured refs Table1[Col]
• Pivots: `:pivot SOURCE TARGET row=COL value=COL agg=sum|count|avg|min|max`
  (auto-refreshes via formulas)
• Charts: `:chart bar A1:A10`, `:chart line`, popup with auto-scale
• Goal Seek: `:goalseek TARGET EXPECTED INPUT`
• Iterative calc: `:iterative on/off` for intentional circular refs
• Data validation: `:validate <COL> "_ > 0"` flags violators in red
• Frozen rows AND columns; hyperlinks (http://) open in browser on Enter
• Mouse: click-to-select, scroll wheel
• .xlsx import + export (auto-detects extension)

=== BASIC & OPERATORS ===
• All formulas start with = (equals sign)
• Cell references use column letter + row number (A1, B2, Z99, AA1, etc.)
• Numbers can be integers or decimals (42, 3.14, -5.5)
• Strings use double quotes ("Hello World", "", "Quote""Test")
• Case insensitive for functions and cell references
• Supports both numbers and strings with automatic type conversion

=== ARITHMETIC OPERATORS ===
+       Addition                    =5+3 → 8, =A1+B1
-       Subtraction                 =10-3 → 7, =A1-5
*       Multiplication              =4*3 → 12, =A1*B1
/       Division                    =15/3 → 5, =A1/B1
**      Exponentiation              =2**3 → 8, =A1**2
^       Power (same as **)          =3^2 → 9, =A1^B1
%       Modulo (remainder)          =10%3 → 1, =A1%B1

=== STRING OPERATORS ===
&       Concatenation               ="Hello" & " " & "World" → Hello World
                                   ="Number: " & 42 → Number: 42
""      String literals             ="Hello World", =""
                                   Use "" for quotes: "Quote""Test" → Quote"Test

=== COMPARISON OPERATORS ===
<       Less than                   =A1<B1 → 1 or 0 (works with numbers only)
>       Greater than                =A1>B1 → 1 or 0 (works with numbers only)
<=      Less than or equal          =A1<=B1 → 1 or 0 (works with numbers only)
>=      Greater than or equal       =A1>=B1 → 1 or 0 (works with numbers only)
=       Equal                       =A1=B1 → 1 or 0 (works with strings and numbers)
<>      Not equal                   =A1<>B1 → 1 or 0 (works with strings and numbers)

Note: Comparisons return 1 for true, 0 for false

=== NUMERIC FUNCTIONS ===
SUM(...)        Sum of values           =SUM(A1,B1,C1) or =SUM(A1:C1)
AVERAGE(...)    Average of values       =AVERAGE(A1:A10)
MIN(...)        Minimum value           =MIN(A1,B1,C1,5)
MAX(...)        Maximum value           =MAX(A1:C3)
ABS(value)      Absolute value          =ABS(-5) → 5
SQRT(value)     Square root             =SQRT(16) → 4
ROUND(num)      Round to integer        =ROUND(3.14) → 3
ROUND(num,places) Round to decimals     =ROUND(3.14159,2) → 3.14

=== STRING FUNCTIONS ===
LEN(text)       String length           =LEN("Hello") → 5
UPPER(text)     Convert to uppercase    =UPPER("hello") → HELLO
LOWER(text)     Convert to lowercase    =LOWER("WORLD") → world
TRIM(text)      Remove leading/trailing spaces  =TRIM("  hi  ") → hi
LEFT(text,num)  First N characters      =LEFT("Hello World",5) → Hello
RIGHT(text,num) Last N characters       =RIGHT("Hello World",5) → World
MID(text,start,len) Substring           =MID("Hello World",6,5) → World
FIND(search,text) Find position         =FIND("lo","Hello") → 3
FIND(search,text,start) Find from pos   =FIND("l","Hello",2) → 3
CONCAT(...)     Concatenate values      =CONCAT("A","B","C") → ABC

=== WEB FUNCTIONS ===
GET(url)        Fetch content from URL  =GET("https://api.example.com/data")
                                       =GET("https://raw.githubusercontent.com/...")

Note: String functions use 0-based indexing (positions start at 0)

=== LOGICAL FUNCTIONS ===
IF(cond,true,false) Conditional         =IF(A1>5,"High","Low")
                                       =IF(A1="Hello","Found","Not Found")
AND(...)        All values true         =AND(A1>0,B1<10)
OR(...)         Any value true          =OR(A1="",A1="N/A")
NOT(value)      Logical not             =NOT(A1>5)
TRUE() / FALSE() Boolean literals
ISBLANK, ISNUMBER, ISTEXT, TYPE          Type tests
COUNT(...), COUNTA(...)                  Counters

Note: For logical tests: 0 and empty strings are false, everything else is true

=== LOOKUP / CONDITIONAL ===
SUMIF(range, criteria, [sum_range])  Sum values matching criteria
COUNTIF(range, criteria)             Count matches; supports ">5", "<=10", "*glob*"
AVERAGEIF(range, criteria, [avg])    Mean of matches
VLOOKUP(value, range, col_index, [exact])
INDEX(range, row, [col])             1-based; col defaults to 1
MATCH(value, range, [type])          type 0 = exact, 1 = approx (default)
INDIRECT(ref_text)                   Build cell ref at eval time: =INDIRECT("A"&ROW())
OFFSET(base, rows, cols, [h], [w])   Range/value offset from base

=== DATE FUNCTIONS ===
TODAY()                              Days since 1899-12-30 (Excel serial)
NOW()                                Days + fractional time-of-day
DATE(year, month, day)               Construct serial
YEAR(serial), MONTH(serial), DAY(serial)
Tip: format a date column with `:format number 0` to see the serial integer.

=== ADVANCED ===
Named ranges: `:name MyRange A1:B10`, then use `MyRange` in formulas.
              `:names` lists; `:unname X` removes.
Conditional formatting: `:cf A "_ > 100" bg=red bold` — `_` binds to cell value.
                       `:cf list`, `:cf clear` to manage rules.
Search options: `:regex on/off`, `:case on/off`.
Auto-save:   `:autosave on/off` — 30s idle window, requires a known filename.
Cache:       `:cache clear` clears GET() cache; F5 recomputes all formulas.
CSV append:  `:import-append PATH` adds rows below current data.
Recent files: tab in load/save dialog cycles through recent + cwd matches.

=== CELL RANGES ===
A1:C3           Rectangle from A1 to C3
A1:A10          Column A, rows 1-10
B2:D2           Row 2, columns B-D

=== TYPE CONVERSION ===
• Numbers in strings are automatically converted: "123" + 1 → 124
• Numbers in string operations: 42 & " items" → "42 items"
• Invalid strings become 0 in math: "hello" + 1 → 1
• String comparisons are case-sensitive: "Hello" <> "hello" → 1

=== FORMULA EXAMPLES ===

Numeric Examples:
=A1+B1*2        Math with precedence
=IF(A1>0,A1*2,0) Conditional calculation
=SUM(A1:A5)/5   Same as AVERAGE(A1:A5)
=MAX(A1:C3)     Largest in 3x3 range
=A1**2+B1**2    Pythagorean calculation

String Examples:
=UPPER(A1) & " - " & LOWER(B1)     Combined formatted text
=IF(LEN(A1)>0,A1,"Empty")          Check for non-empty strings
=LEFT(A1,FIND(" ",A1)-1)           Extract first word
="Hello " & A1 & ", you scored " & B1 & "%"   Dynamic messages
=IF(AND(LEN(A1)>3,A1<>""),"Valid","Invalid")  Validate input

Mixed Type Examples:
="Total: " & SUM(A1:A10) & " items"   Numeric result with description
=IF(AVERAGE(A1:A10)>50,"PASS","FAIL") Grade based on average
=CONCAT("Value: ",A1," Total: ",SUM(B1:B5))  Dynamic labels

=== SEARCH FUNCTIONALITY ===
/               Start text search across all cells
                Search is case-insensitive and searches both cell values and formulas
                Live search: results update as you type
                ↑↓          Navigate through search results while searching
                Enter       Finish search and return to normal mode
                Esc         Cancel search and return to normal mode
                n/N         Navigate search results in normal mode (after search)

=== FILE OPERATIONS ===
Ctrl+S          Save in place (or Save As if file is new)
Ctrl+O          Load spreadsheet from file (prompts to confirm if dirty)
Ctrl+E          Export spreadsheet to CSV file
Ctrl+L          Import data from CSV file (prompts to confirm if dirty)
                Files are saved as "spreadsheet.tshts" in JSON format
                CSV exports contain only cell values (not formulas)
                CSV imports replace current spreadsheet data

Open at startup: pass a filename on the command line: `tshts foo.tshts`.

=== NAVIGATION SHORTCUTS ===
TSHTS uses vim-style modes. In NORMAL mode, single letters trigger commands;
to type into a cell, enter INSERT mode first (see VIM MODE section below).
The classic Excel/Sheets Ctrl-bindings still work alongside the vim layer.

F1 or ?         Show this help (scroll with ↑↓, PgUp/PgDn, Home)
Enter / F2 / i  Start editing the current cell (INSERT mode)
Arrow keys      Navigate cells (h / j / k / l also work)
Shift+arrows    Quick range-select (or use `v` to enter VISUAL mode)
+ key           Auto-resize all columns to fit content
- / _ keys      Manually shrink/grow column width
Ctrl+Z / u      Undo last action
Ctrl+Y / Ctrl+R Redo last undone action
Ctrl+G          Go to cell reference
Ctrl+C / X      Copy / Cut (uses system clipboard for plain text)
Ctrl+V          Enter VISUAL BLOCK mode (paste is `p` or `P`)
Ctrl+D          Autofill the selection from its top-left cell
Ctrl+B / U      Toggle bold / underline on selection
Ctrl+Home/End   Jump to A1 / last cell with data
Ctrl+PgUp/PgDn  Switch to previous / next sheet
/               Start text search (live)
n / N           Next / previous search result
:               Ex-command palette (`:q` quit, `:w` save, `:wq` save+quit, ...)
F5              Recalculate all formulas (refresh RAND / GET caches)
Esc             Clear selection / search highlights / pending op / status
q  or  :q       Quit (prompts if there are unsaved changes)
:q!             Force quit, discarding changes

=== VIM MODE ===
TSHTS speaks vim. There are four primary modes — the current one is shown
as `-- NORMAL --`, `-- INSERT --`, `-- VISUAL --`, or `-- COMMAND --` at
the left of the status bar. Pending operators and counts also show there
(e.g. `[5d]` while you're typing `5dd`).

--- Mode transitions ---
i / a           Enter INSERT mode (edit current cell)
I               INSERT with cursor at start of the cell's text
A               INSERT with cursor at end of the cell's text
o               Open a new row below, enter INSERT mode
O               Move to the row above, enter INSERT mode
s               Substitute cell: clear it and enter INSERT
S               Substitute row: clear current row, enter INSERT at col 0
Enter / F2      Enter INSERT mode without clearing
Esc             Return to NORMAL from any mode (also cancels pending op)
v               Enter VISUAL (cell-granularity selection)
V               Enter VISUAL LINE (whole-row selection)
Ctrl+V          Enter VISUAL BLOCK (rectangular selection)
:               Enter COMMAND (ex-command palette)

--- Motions (work in NORMAL and VISUAL) ---
h j k l         Left / down / up / right
Arrow keys      Same as h/j/k/l
0 / Home        First column of the current row
$ / End         Last column with data in the current row
^               First column with data in the current row
gg              Jump to first row
G               Jump to last row with data
NG              Jump to row N (e.g. `42G`)
PgUp / PgDn     Page up / down
Tab / Shift+Tab Move right / left (Excel-style)

--- Operators (NORMAL mode) ---
Operators set a pending state shown in the status bar; press a motion
next to apply, or repeat the operator key for a whole-row operation.

d{motion}       Delete the range from cursor to motion target
y{motion}       Yank (copy) the range
c{motion}       Change: delete and enter INSERT at the start
dd / yy / cc    Whole-row delete / yank / change
x               Delete current cell (no motion needed)
p               Paste after / below cursor
P               Paste before / above cursor
u               Undo
Ctrl+R          Redo

--- Counts ---
Prefix any motion or operator with a count.
  5j            Move down 5 rows
  3l            Move right 3 columns
  10G           Jump to row 10
  3dd           Delete 3 rows starting here
  2yj           Yank current row + 2 rows below it

--- VISUAL mode ---
Once in visual mode, motions extend the selection. Status bar shows the
size and SUM / AVG / COUNT of the current selection. Then:
  y             Yank the selection
  d  or  x      Delete the selection
  c             Change: delete and enter INSERT at the top-left
  p             Paste over the selection
  v / Esc       Exit VISUAL back to NORMAL

--- COMMAND mode (`:`) ---
:q              Quit (prompts on unsaved changes)
:q!             Force quit, discard changes
:w              Save (in place if known, else prompts)
:w <filename>   Save as <filename> (.xlsx auto-detected)
:wq  or  :x     Save and quit
:wq!  or  :x!   Save and force quit
:e <filename>   Open another file
Every command in this help (`:sort asc`, `:freeze`, `:format number`,
`:cf …`, `:name …`, etc.) is also reachable here. Type and press Tab
to cycle through suggestions.

=== ABSOLUTE REFERENCES ===
$A$1            Both row and column absolute
$A1             Column A is absolute; row shifts with autofill/paste
A$1             Column A shifts; row 1 is absolute
Use `$` to anchor part of a reference during autofill or paste.

=== SELECTION AND AUTOFILL ===
Shift+arrows    Select a range of cells by holding Shift and using arrow keys
Ctrl+D          Autofill: Copy formula from top-left cell of selection to all
                selected cells, automatically adjusting cell references
                Example: =SUM(B4:B6) becomes =SUM(C4:C6), =SUM(D4:D6), etc.
                when dragged right, or =SUM(B5:B7), =SUM(B6:B8), etc. when
                dragged down. Works with any formula containing cell references.

=== HELP NAVIGATION ===
↑↓ or j/k       Scroll help text up/down one line
Page Up/Down    Scroll help text up/down 5 lines
Home            Jump to top of help text
Esc/F1/?/q      Close this help window

=== ERROR HANDLING ===
#ERROR          Displayed when formula evaluation fails
                Common causes: division by zero, invalid functions,
                circular references, invalid FIND operations

Note: Your spreadsheet is automatically saved when you use Ctrl+S.
Use Ctrl+O to load the saved spreadsheet on next session."#.to_string()
}