//! Submodule of `state` — see state/mod.rs.

use super::*;

impl App {
    pub fn start_command_palette(&mut self) {
        self.mode = AppMode::CommandPalette;
        self.command_input.clear();
        self.cursor_position = 0;
        self.status_message = None;
    }

    pub fn execute_command(&mut self) {
        // Handle vim-style ex commands that take filenames first — they need
        // case-preserved arguments. Note `:wq` / `:x` / `:wq!` / `:x!` save +
        // quit, and `:q!` skips the dirty check.
        let trimmed = self.command_input.trim().to_string();
        let lower = trimmed.to_lowercase();
        let close_palette = |s: &mut Self| {
            s.mode = AppMode::Normal;
            s.command_input.clear();
            s.cursor_position = 0;
        };
        // Empty input: just close the palette silently (vim convention).
        if trimmed.is_empty() {
            close_palette(self);
            return;
        }
        // Quit / save-quit variants — exact-match on lowercased string.
        match lower.as_str() {
            "q" => {
                close_palette(self);
                self.request_quit();
                return;
            }
            "q!" | "quit!" => {
                close_palette(self);
                self.should_quit = true;
                return;
            }
            "w" => {
                close_palette(self);
                self.save_in_place_or_prompt();
                return;
            }
            "wq" | "x" => {
                close_palette(self);
                self.save_in_place_or_prompt();
                if !self.dirty {
                    self.should_quit = true;
                }
                return;
            }
            "wq!" | "x!" => {
                close_palette(self);
                self.save_in_place_or_prompt();
                self.should_quit = true;
                return;
            }
            "e" | "edit" => {
                close_palette(self);
                self.request_load_file();
                return;
            }
            _ => {}
        }
        // `:w <filename>` and `:e <filename>` preserve case.
        if let Some(rest) = trimmed.strip_prefix("w ").or_else(|| trimmed.strip_prefix("W ")) {
            let name = rest.trim().to_string();
            if !name.is_empty() {
                let result = if name.to_lowercase().ends_with(".xlsx") {
                    crate::infrastructure::xlsx::save_xlsx(&self.workbook, &name)
                        .map(|_| name.clone())
                } else {
                    crate::infrastructure::FileRepository::save_workbook(&self.workbook, &name)
                };
                self.set_save_result(result);
            } else {
                self.status_message = Some("Usage: w <filename>".to_string());
            }
            close_palette(self);
            return;
        }
        if let Some(rest) = trimmed.strip_prefix("e ").or_else(|| trimmed.strip_prefix("E ")) {
            let name = rest.trim().to_string();
            if name.is_empty() {
                self.status_message = Some("Usage: e <filename>".to_string());
                close_palette(self);
                return;
            }
            // `:e <file>` loads the typed file directly. Going through
            // request_load_file/start_load_file would clobber `filename_input`
            // with the currently-open file's name (or the "spreadsheet.tshts"
            // default), defeating the purpose. Vim's `:e` is also a direct
            // load; users who want a dirty-guard can save first.
            let result = if name.to_lowercase().ends_with(".xlsx") {
                crate::infrastructure::xlsx::load_xlsx(&name)
                    .map(|wb| (wb, name.clone()))
            } else {
                crate::infrastructure::FileRepository::load_workbook(&name)
            };
            self.set_load_workbook_result(result);
            close_palette(self);
            return;
        }
        // `:export <file>` writes the current sheet to CSV.
        if let Some(rest) = trimmed
            .strip_prefix("export ")
            .or_else(|| trimmed.strip_prefix("EXPORT "))
        {
            let name = rest.trim().to_string();
            if name.is_empty() {
                self.status_message = Some("Usage: export <filename>".to_string());
            } else {
                let result = crate::domain::CsvExporter::export_to_csv(
                    self.workbook.current_sheet(),
                    &name,
                );
                self.set_csv_export_result(result);
            }
            close_palette(self);
            return;
        }
        // `:import <file>` replaces the current sheet from CSV. We could
        // wire this through request_csv_import for a dirty-check prompt, but
        // start_csv_import would then clobber the typed filename — so for
        // now, mirror `:e <file>`'s pattern and import directly. Users who
        // want a dirty-guard can `:w` first.
        if let Some(rest) = trimmed
            .strip_prefix("import ")
            .or_else(|| trimmed.strip_prefix("IMPORT "))
        {
            let name = rest.trim().to_string();
            if name.is_empty() {
                self.status_message = Some("Usage: import <filename>".to_string());
            } else {
                let result = crate::domain::CsvExporter::import_from_csv(&name);
                self.set_csv_import_result(result);
            }
            close_palette(self);
            return;
        }
        if trimmed.starts_with("rename ") || trimmed.starts_with("RENAME ") {
            let name = trimmed[7..].trim().to_string();
            if !name.is_empty() {
                if self.workbook.rename_sheet(name.clone()) {
                    self.dirty = true;
                    crate::infrastructure::autosave::mark_dirty();
                    self.status_message = Some(format!("Renamed sheet to '{}'", name));
                } else {
                    self.status_message = Some(format!(
                        "Rename rejected: '{}' is empty or a duplicate",
                        name
                    ));
                }
            } else {
                self.status_message = Some("Usage: rename <name>".to_string());
            }
            self.mode = AppMode::Normal;
            self.command_input.clear();
            self.cursor_position = 0;
            return;
        }

        let cmd = trimmed.to_lowercase();
        let parts: Vec<&str> = cmd.split_whitespace().collect();
        match parts.as_slice() {
            ["ir"] | ["insert", "row"] => self.insert_row(),
            ["dr"] | ["delete", "row"] => self.delete_row(),
            ["ic"] | ["insert", "col"] | ["insert", "column"] => self.insert_col(),
            ["dc"] | ["delete", "col"] | ["delete", "column"] => self.delete_col(),
            ["replace"] | ["find-replace"] | ["find", "replace"] => {
                self.start_find_replace();
            }
            ["sort", "asc"] => self.sort_column_asc(),
            ["sort", "desc"] => self.sort_column_desc(),
            ["freeze"] => {
                // Clamp to the active sheet's bounds so a stale cursor past
                // the last row/col doesn't produce an unreachable frozen pane.
                let sheet = self.workbook.current_sheet();
                let max_r = sheet.rows.saturating_sub(1);
                let max_c = sheet.cols.saturating_sub(1);
                self.frozen_rows = self.selected_row.min(max_r);
                self.frozen_cols = self.selected_col.min(max_c);
                self.status_message = Some(format!("Frozen {} rows, {} cols", self.frozen_rows, self.frozen_cols));
            }
            ["unfreeze"] => {
                self.frozen_rows = 0;
                self.frozen_cols = 0;
                self.status_message = Some("Unfrozen all panes".to_string());
            }
            ["format", "general"] => self.set_selection_format(NumberFormat::General),
            ["format", "number"] => self.set_selection_format(NumberFormat::Number { decimals: 2, thousands_sep: false }),
            ["format", "number", d] => {
                if let Ok(decimals) = d.parse::<u32>() {
                    self.set_selection_format(NumberFormat::Number { decimals, thousands_sep: false });
                } else {
                    self.status_message = Some("Invalid decimal count".to_string());
                }
            }
            ["format", "currency"] => self.set_selection_format(NumberFormat::Currency { symbol: "$".to_string(), decimals: 2 }),
            ["format", "currency", sym] => self.set_selection_format(NumberFormat::Currency { symbol: sym.to_string(), decimals: 2 }),
            ["format", "percent"] | ["format", "percentage"] => self.set_selection_format(NumberFormat::Percentage { decimals: 1 }),
            ["format", "percent", d] | ["format", "percentage", d] => {
                if let Ok(decimals) = d.parse::<u32>() {
                    self.set_selection_format(NumberFormat::Percentage { decimals });
                } else {
                    self.status_message = Some("Invalid decimal count".to_string());
                }
            }
            ["bold"] => self.toggle_bold(),
            ["underline"] => self.toggle_underline(),
            ["color", color_name] => {
                if *color_name == "none" || *color_name == "default" {
                    self.set_selection_fg_color(None);
                } else if let Some(c) = TerminalColor::from_name(color_name) {
                    self.set_selection_fg_color(Some(c));
                } else {
                    self.status_message = Some(format!("Unknown color: {}", color_name));
                }
            }
            ["bg", color_name] => {
                if *color_name == "none" || *color_name == "default" {
                    self.set_selection_bg_color(None);
                } else if let Some(c) = TerminalColor::from_name(color_name) {
                    self.set_selection_bg_color(Some(c));
                } else {
                    self.status_message = Some(format!("Unknown color: {}", color_name));
                }
            }
            ["sheet", "new"] | ["new", "sheet"] | ["addsheet"] => {
                let name = format!("Sheet{}", self.workbook.sheets.len() + 1);
                self.workbook.add_sheet(name.clone());
                self.workbook.active_sheet = self.workbook.sheets.len() - 1;
                self.selected_row = 0;
                self.selected_col = 0;
                self.scroll_row = 0;
                self.scroll_col = 0;
                self.dirty = true;
                crate::infrastructure::autosave::mark_dirty();
                self.status_message = Some(format!("Added sheet '{}'", name));
            }
            ["sheet", "delete"] | ["delsheet"] => {
                let name = self.workbook.sheet_names[self.workbook.active_sheet].clone();
                if self.workbook.remove_sheet(self.workbook.active_sheet) {
                    self.selected_row = 0;
                    self.selected_col = 0;
                    self.scroll_row = 0;
                    self.scroll_col = 0;
                    self.dirty = true;
                    crate::infrastructure::autosave::mark_dirty();
                    // remove_sheet may have shifted the active sheet (the
                    // workbook re-anchors it). Drop search/find-replace state
                    // since their (row, col) results belonged to the old
                    // active sheet.
                    self.invalidate_cross_sheet_state();
                    self.status_message = Some(format!("Deleted sheet '{}'", name));
                } else {
                    self.status_message = Some("Cannot delete the last sheet".to_string());
                }
            }
            ["sheet", "next"] | ["sn"] => {
                self.switch_next_sheet();
            }
            ["sheet", "prev"] | ["sp"] => {
                self.switch_prev_sheet();
            }
            ["comment", ..] => {
                // Preserve case for the comment text — the user wrote it for
                // a human to read later. `parts` is lowercased so we recover
                // from `trimmed`.
                let preserved = trimmed
                    .strip_prefix("comment ")
                    .or_else(|| trimmed.strip_prefix("COMMENT "))
                    .or_else(|| trimmed.strip_prefix("Comment "))
                    .unwrap_or("")
                    .trim()
                    .to_string();
                let text_lc = parts[1..].join(" ");
                if text_lc == "clear" || text_lc == "none" {
                    self.set_cell_comment(None);
                } else if !preserved.is_empty() {
                    self.set_cell_comment(Some(preserved));
                } else {
                    self.set_cell_comment(Some(text_lc));
                }
            }
            ["filter", column_name] => {
                if let Some(col) = Spreadsheet::parse_column_label(column_name) {
                    self.apply_filter(col, None);
                } else {
                    self.status_message = Some(format!("Invalid column: {}", column_name));
                }
            }
            ["filter", column_name, ..] => {
                if let Some(col) = Spreadsheet::parse_column_label(column_name) {
                    let criteria = parts[2..].join(" ");
                    self.apply_filter(col, Some(criteria));
                } else {
                    self.status_message = Some(format!("Invalid column: {}", column_name));
                }
            }
            ["unfilter"] | ["clearfilter"] | ["clear", "filter"] => {
                self.clear_filter();
            }
            ["recalc"] | ["refresh"] => self.recalc_all(),
            ["cache", "clear"] => {
                crate::infrastructure::fetcher::clear_cache();
                self.status_message = Some("GET cache cleared".to_string());
            }
            ["net", "on"] | ["network", "on"] => {
                crate::infrastructure::fetcher::set_network_enabled(true);
                // Clear any cached "network disabled" errors so the next
                // recalc actually attempts the fetch.
                crate::infrastructure::fetcher::clear_cache();
                self.status_message = Some("Network GET enabled for this session".to_string());
                self.recalc_all();
            }
            ["net", "off"] | ["network", "off"] => {
                crate::infrastructure::fetcher::set_network_enabled(false);
                crate::infrastructure::fetcher::clear_cache();
                self.status_message = Some("Network GET disabled".to_string());
            }
            ["net", "status"] | ["network", "status"] => {
                let on = crate::infrastructure::fetcher::network_enabled();
                self.status_message = Some(format!(
                    "Network GET: {}",
                    if on { "enabled" } else { "disabled (run :net on)" }
                ));
            }
            ["clipboard", "clear"] => {
                crate::infrastructure::sidecar::clear();
                self.clipboard = None;
                self.status_message = Some("Clipboard cleared".to_string());
            }
            ["import-append", path] | ["append", path] => {
                let path = path.to_string();
                match crate::domain::CsvExporter::append_from_csv(
                    self.workbook.current_sheet_mut(),
                    &path,
                ) {
                    Ok(n) => {
                        self.dirty = true;
                        self.status_message =
                            Some(format!("Appended {} row(s) from {}", n, path));
                    }
                    Err(e) => {
                        self.status_message = Some(format!("Append failed: {}", e));
                    }
                }
            }
            ["hide", "row"] => {
                let hidden = self.selected_row;
                self.hidden_rows.insert(hidden);
                self.status_message = Some(format!("Hid row {}", hidden + 1));
                // Move cursor off the hidden row — otherwise the formula bar
                // keeps showing hidden content and the user can't escape.
                let max_row = self.workbook.current_sheet().rows.saturating_sub(1);
                let mut next = self.selected_row + 1;
                while next <= max_row && self.hidden_rows.contains(&next) {
                    next += 1;
                }
                if next <= max_row {
                    self.selected_row = next;
                } else {
                    // No visible row below — try above.
                    let mut prev = hidden;
                    while prev > 0 && self.hidden_rows.contains(&prev) {
                        prev -= 1;
                    }
                    if !self.hidden_rows.contains(&prev) {
                        self.selected_row = prev;
                    }
                }
                self.ensure_cursor_visible();
            }
            ["show", "rows"] | ["unhide", "rows"] => {
                self.hidden_rows.clear();
                self.status_message = Some("All rows shown".to_string());
            }
            ["hide", "col"] => {
                let hidden = self.selected_col;
                self.hidden_cols.insert(hidden);
                self.status_message = Some(format!(
                    "Hid column {}",
                    crate::domain::Spreadsheet::column_label(hidden)
                ));
                // Move cursor off the hidden col — same reason as hide row.
                let max_col = self.workbook.current_sheet().cols.saturating_sub(1);
                let mut next = self.selected_col + 1;
                while next <= max_col && self.hidden_cols.contains(&next) {
                    next += 1;
                }
                if next <= max_col {
                    self.selected_col = next;
                } else {
                    let mut prev = hidden;
                    while prev > 0 && self.hidden_cols.contains(&prev) {
                        prev -= 1;
                    }
                    if !self.hidden_cols.contains(&prev) {
                        self.selected_col = prev;
                    }
                }
                self.ensure_cursor_visible();
            }
            ["hide", "col", col_name] => {
                if let Some(c) = Spreadsheet::parse_column_label(col_name) {
                    self.hidden_cols.insert(c);
                    self.status_message = Some(format!("Hid column {}", col_name));
                    // If the cursor sits on the just-hidden column, advance it.
                    if self.selected_col == c {
                        let max_col = self.workbook.current_sheet().cols.saturating_sub(1);
                        let mut next = c + 1;
                        while next <= max_col && self.hidden_cols.contains(&next) {
                            next += 1;
                        }
                        if next <= max_col {
                            self.selected_col = next;
                        } else if c > 0 {
                            let mut prev = c - 1;
                            while prev > 0 && self.hidden_cols.contains(&prev) {
                                prev -= 1;
                            }
                            if !self.hidden_cols.contains(&prev) {
                                self.selected_col = prev;
                            }
                        }
                        self.ensure_cursor_visible();
                    }
                } else {
                    self.status_message = Some(format!("Invalid column: {}", col_name));
                }
            }
            ["show", "cols"] | ["unhide", "cols"] => {
                self.hidden_cols.clear();
                self.status_message = Some("All columns shown".to_string());
            }
            ["show", "all"] => {
                self.hidden_rows.clear();
                self.hidden_cols.clear();
                self.status_message = Some("All rows and columns shown".to_string());
            }
            // `:table create A1:D10 name=Sales`
            ["table", "create", range_str, opts @ ..] => {
                if let Some((start, end)) = parse_range(range_str) {
                    let mut name = format!("Table{}", self.workbook.current_sheet().tables.len() + 1);
                    // Recover case-preserved opts from the original command so
                    // `name=Mine` is stored as "Mine", not "mine".
                    let original_opts: Vec<&str> = trimmed
                        .split_whitespace()
                        .skip(3)
                        .collect();
                    for (i, opt) in opts.iter().enumerate() {
                        if opt.starts_with("name=") {
                            // Pull the same position from the original (case-preserved) tokens.
                            if let Some(orig) = original_opts.get(i)
                                && let Some(n) = orig.strip_prefix("name=") {
                                    name = n.to_string();
                                    continue;
                                }
                            // Fallback: use lowercased opt.
                            if let Some(n) = opt.strip_prefix("name=") {
                                name = n.to_string();
                            }
                        }
                    }
                    let sheet = self.workbook.current_sheet();
                    let cols = end.1 - start.1 + 1;
                    let mut headers = Vec::with_capacity(cols);
                    for c in start.1..=end.1 {
                        let h = sheet.get_cell(start.0, c).value;
                        headers.push(if h.is_empty() {
                            crate::domain::Spreadsheet::column_label(c)
                        } else {
                            h
                        });
                    }
                    let table = crate::domain::Table {
                        name: name.clone(),
                        top_row: start.0,
                        left_col: start.1,
                        bottom_row: end.0,
                        right_col: end.1,
                        headers,
                    };
                    self.workbook.current_sheet_mut().tables.push(table);
                    // Also register each column as a named range so
                    // `Table1[Col1]` works via existing name resolution.
                    let sheet = self.workbook.current_sheet();
                    let table = sheet.tables.last().unwrap().clone();
                    let body_top = table.top_row + 1; // skip header row
                    for (i, header) in table.headers.iter().enumerate() {
                        let key = format!("{}[{}]", table.name.to_uppercase(), header.to_uppercase());
                        let value = format!(
                            "{}:{}",
                            crate::domain::Spreadsheet::format_cell_reference(body_top, table.left_col + i, false, false),
                            crate::domain::Spreadsheet::format_cell_reference(table.bottom_row, table.left_col + i, false, false),
                        );
                        self.workbook.named_ranges.insert(key.clone(), value.clone());
                        for s in &mut self.workbook.sheets {
                            s.named_ranges.insert(key.clone(), value.clone());
                        }
                    }
                    self.dirty = true;
                    self.status_message = Some(format!("Created table '{}'", name));
                } else {
                    self.status_message = Some(format!("Invalid range: {}", range_str));
                }
            }
            // `:pivot RANGE TARGET row=COL value=COL agg=sum|count|avg|min|max`
            // `:chart bar A1:A10 [title=Sales]`
            ["chart", kind, range_str, opts @ ..] => {
                if let Some((s, e)) = parse_range(range_str) {
                    let mut title = format!("{} chart", kind);
                    for o in opts {
                        if let Some(t) = o.strip_prefix("title=") {
                            title = t.to_string();
                        }
                    }
                    let chart_kind = match *kind {
                        "bar" => ChartKind::Bar,
                        "line" => ChartKind::Line,
                        _ => ChartKind::Sparkline,
                    };
                    self.chart_popup = Some(ChartPopup {
                        title,
                        source: (s, e),
                        kind: chart_kind,
                    });
                    self.status_message = Some("Chart shown (Esc to close)".to_string());
                } else {
                    self.status_message = Some("chart: bad range".to_string());
                }
            }
            ["iterative", "on"] => {
                self.iterative_calc = true;
                for s in &mut self.workbook.sheets {
                    s.iterative_calc = true;
                }
                self.status_message = Some(format!(
                    "Iterative calc: on (max {} iters, eps {})",
                    self.workbook.current_sheet().iter_max,
                    self.workbook.current_sheet().iter_epsilon
                ));
            }
            ["iterative", "off"] => {
                self.iterative_calc = false;
                for s in &mut self.workbook.sheets {
                    s.iterative_calc = false;
                }
                self.status_message = Some("Iterative calc: off".to_string());
            }
            ["iterative", "max", n] => {
                if let Ok(v) = n.parse::<usize>() {
                    for s in &mut self.workbook.sheets {
                        s.iter_max = v;
                    }
                    self.status_message = Some(format!("Iterative max = {}", v));
                } else {
                    self.status_message = Some("iterative max: bad number".to_string());
                }
            }
            ["iterative", "epsilon", n] => {
                if let Ok(v) = n.parse::<f64>() {
                    for s in &mut self.workbook.sheets {
                        s.iter_epsilon = v;
                    }
                    self.status_message = Some(format!("Iterative epsilon = {}", v));
                } else {
                    self.status_message = Some("iterative epsilon: bad number".to_string());
                }
            }
            ["r1c1", "on"] => {
                self.r1c1_mode = true;
                self.status_message = Some("R1C1 reference mode: on".to_string());
            }
            ["r1c1", "off"] => {
                self.r1c1_mode = false;
                self.status_message = Some("R1C1 reference mode: off".to_string());
            }
            // `:validate A "_ > 0"` — predicate with `_` bound to the cell value.
            ["validate", col_name, rest @ ..] if !rest.is_empty() => {
                if let Some(col) = Spreadsheet::parse_column_label(col_name) {
                    let predicate = rest.join(" ");
                    self.validations.insert(col, predicate.clone());
                    self.status_message = Some(format!("Validation set on col {}", col_name));
                } else {
                    self.status_message = Some(format!("Invalid column: {}", col_name));
                }
            }
            ["validate", "clear"] => {
                self.validations.clear();
                self.status_message = Some("Validations cleared".to_string());
            }
            ["pivot", source, target, opts @ ..] => {
                if let (Some((s, e)), Some(t)) = (
                    parse_range(source),
                    crate::domain::Spreadsheet::parse_cell_reference(target),
                ) {
                    let mut row_col: Option<usize> = None;
                    let mut val_col: Option<usize> = None;
                    let mut agg = "sum".to_string();
                    for o in opts {
                        if let Some(c) = o.strip_prefix("row=") {
                            row_col = crate::domain::Spreadsheet::parse_column_label(c);
                        } else if let Some(c) = o.strip_prefix("value=") {
                            val_col = crate::domain::Spreadsheet::parse_column_label(c);
                        } else if let Some(a) = o.strip_prefix("agg=") {
                            agg = a.to_string();
                        }
                    }
                    let row_col = match row_col {
                        Some(c) => c,
                        None => {
                            self.status_message = Some("pivot: row=COL required".to_string());
                            return;
                        }
                    };
                    let val_col = val_col.unwrap_or(row_col);
                    // Collect distinct row-keys preserving sort order. The
                    // pivot values are written as formulas (`=SUMIF(...)`)
                    // so they auto-update when source data changes.
                    let mut keys: std::collections::BTreeSet<String> =
                        std::collections::BTreeSet::new();
                    {
                        let sheet = self.workbook.current_sheet();
                        for r in s.0..=e.0 {
                            let k = sheet.get_cell(r, row_col).value;
                            if !k.is_empty() {
                                keys.insert(k);
                            }
                        }
                    }
                    let row_col_label = crate::domain::Spreadsheet::column_label(row_col);
                    let val_col_label = crate::domain::Spreadsheet::column_label(val_col);
                    let source_range = format!(
                        "{}{}:{}{}",
                        row_col_label,
                        s.0 + 1,
                        row_col_label,
                        e.0 + 1
                    );
                    let sum_range = format!(
                        "{}{}:{}{}",
                        val_col_label,
                        s.0 + 1,
                        val_col_label,
                        e.0 + 1
                    );
                    let agg_formula = |key: &str| -> String {
                        let key_esc = key.replace('"', "\"\"");
                        match agg.as_str() {
                            "count" => format!(
                                "=COUNTIF({}, \"{}\")",
                                source_range, key_esc
                            ),
                            "avg" | "average" => format!(
                                "=AVERAGEIF({}, \"{}\", {})",
                                source_range, key_esc, sum_range
                            ),
                            "min" => format!(
                                "=MIN(IF({}=\"{}\", {}))",
                                source_range, key_esc, sum_range
                            ),
                            "max" => format!(
                                "=MAX(IF({}=\"{}\", {}))",
                                source_range, key_esc, sum_range
                            ),
                            _ => format!(
                                "=SUMIF({}, \"{}\", {})",
                                source_range, key_esc, sum_range
                            ),
                        }
                    };
                    let mut rows: Vec<(usize, usize, CellData)> = Vec::new();
                    rows.push((
                        t.0,
                        t.1,
                        CellData {
                            value: "Key".to_string(),
                            formula: None,
                            format: None,
                            comment: None,
                        spill_anchor: None,
                        },
                    ));
                    rows.push((
                        t.0,
                        t.1 + 1,
                        CellData {
                            value: agg.clone(),
                            formula: None,
                            format: None,
                            comment: None,
                        spill_anchor: None,
                        },
                    ));
                    for (i, k) in keys.iter().enumerate() {
                        rows.push((
                            t.0 + 1 + i,
                            t.1,
                            CellData {
                                value: k.clone(),
                                formula: None,
                                format: None,
                                comment: None,
                            spill_anchor: None,
                            },
                        ));
                        let formula = agg_formula(k);
                        // Evaluate immediately so the cell shows its initial
                        // value; set_many will rebuild deps so it tracks.
                        let evaluator = FormulaEvaluator::with_workbook(
                            &self.workbook,
                            self.workbook.current_sheet(),
                            &self.workbook.named_ranges,
                        );
                        let initial = evaluator.evaluate_formula(&formula);
                        rows.push((
                            t.0 + 1 + i,
                            t.1 + 1,
                            CellData {
                                value: initial,
                                formula: Some(formula),
                                format: None,
                                comment: None,
                            spill_anchor: None,
                            },
                        ));
                    }
                    self.set_many_with_undo(rows);
                    self.status_message = Some(format!(
                        "Pivot written to {}{}: {} groups (auto-refreshes via formulas)",
                        crate::domain::Spreadsheet::column_label(t.1),
                        t.0 + 1,
                        keys.len()
                    ));
                } else {
                    self.status_message = Some("pivot: bad source range or target".to_string());
                }
            }
            // `:goalseek TARGET_CELL EXPECTED INPUT_CELL` — bisect input until target = expected.
            ["goalseek", target, expected, input] => {
                let target_pos = crate::domain::Spreadsheet::parse_cell_reference(target);
                let input_pos = crate::domain::Spreadsheet::parse_cell_reference(input);
                let expected_v: f64 = expected.parse().unwrap_or(0.0);
                if let (Some(t), Some(i)) = (target_pos, input_pos) {
                    // Snapshot the original input cell before bisection so we
                    // can either undo the final write or restore on failure
                    // without leaking ~80 intermediate values into undo history.
                    let original_cell = self.workbook.current_sheet().get_cell(i.0, i.1);
                    let mut lo = -1e9_f64;
                    let mut hi = 1e9_f64;
                    let mut result: Option<f64> = None;
                    for _ in 0..80 {
                        let mid = (lo + hi) / 2.0;
                        let cell = CellData {
                            value: mid.to_string(),
                            formula: None,
                            format: None,
                            comment: None,
                            spill_anchor: None,
                        };
                        self.workbook.current_sheet_mut().set_cell(i.0, i.1, cell);
                        let cur: f64 = self.workbook.current_sheet().get_cell(t.0, t.1).value.parse().unwrap_or(0.0);
                        if (cur - expected_v).abs() < 1e-6 {
                            result = Some(mid);
                            break;
                        }
                        if cur < expected_v {
                            lo = mid;
                        } else {
                            hi = mid;
                        }
                    }
                    if let Some(v) = result {
                        // Restore the cell to its pre-bisection state, then
                        // perform the final write through set_cell_with_undo
                        // so the recorded `old_cell` is the *original* value
                        // (not the last bisection step). The intermediate
                        // bisection writes left cross-sheet dependents holding
                        // stale `mid` values; the final set_cell_with_undo
                        // propagates the converged value to them.
                        self.workbook
                            .current_sheet_mut()
                            .set_cell(i.0, i.1, original_cell);
                        let final_cell = CellData {
                            value: v.to_string(),
                            formula: None,
                            format: None,
                            comment: None,
                            spill_anchor: None,
                        };
                        self.set_cell_with_undo(i.0, i.1, final_cell);
                        self.status_message = Some(format!("Goal seek: {} = {:.6}", input, v));
                    } else {
                        // No-op net effect: put the original cell back, no
                        // undo entry needed. Propagate so cross-sheet
                        // dependents see the *original* value, not the last
                        // bisection step.
                        self.workbook
                            .current_sheet_mut()
                            .set_cell(i.0, i.1, original_cell);
                        self.propagate_cell_change(i.0, i.1);
                        self.status_message = Some("Goal seek did not converge".to_string());
                    }
                } else {
                    self.status_message = Some("goalseek: bad cell reference".to_string());
                }
            }
            // `:trace` — show what cells the current cell's formula depends on.
            ["trace"] | ["trace", "precedents"] => {
                let cell = self
                    .workbook
                    .current_sheet()
                    .get_cell(self.selected_row, self.selected_col);
                if let Some(formula) = cell.formula {
                    let evaluator = FormulaEvaluator::new(self.workbook.current_sheet());
                    let refs = evaluator.extract_cell_references(&formula);
                    if refs.is_empty() {
                        self.status_message = Some("(no precedents)".to_string());
                    } else {
                        let refs_str: Vec<String> = refs
                            .iter()
                            .map(|(r, c)| {
                                format!(
                                    "{}{}",
                                    crate::domain::Spreadsheet::column_label(*c),
                                    r + 1
                                )
                            })
                            .collect();
                        self.status_message = Some(format!("Precedents: {}", refs_str.join(", ")));
                    }
                } else {
                    self.status_message = Some("(cell has no formula)".to_string());
                }
            }
            ["trace", "dependents"] => {
                let pos = (self.selected_row, self.selected_col);
                let deps = self.workbook.current_sheet().dependents.get(&pos).cloned();
                match deps {
                    Some(set) if !set.is_empty() => {
                        let s: Vec<String> = set
                            .iter()
                            .map(|(r, c)| {
                                format!(
                                    "{}{}",
                                    crate::domain::Spreadsheet::column_label(*c),
                                    r + 1
                                )
                            })
                            .collect();
                        self.status_message = Some(format!("Dependents: {}", s.join(", ")));
                    }
                    _ => self.status_message = Some("(no dependents)".to_string()),
                }
            }
            ["table", "list"] => {
                let sheet = self.workbook.current_sheet();
                if sheet.tables.is_empty() {
                    self.status_message = Some("(no tables on this sheet)".to_string());
                } else {
                    let names: Vec<String> = sheet
                        .tables
                        .iter()
                        .map(|t| format!("{}({}x{})", t.name, t.bottom_row - t.top_row + 1, t.headers.len()))
                        .collect();
                    self.status_message = Some(names.join(", "));
                }
            }
            ["cf", "clear"] => {
                let sheet_idx = self.workbook.active_sheet;
                let old = self.workbook.current_sheet().conditional_formats.clone();
                let n = old.len();
                if n > 0 {
                    self.workbook.current_sheet_mut().conditional_formats.clear();
                    self.workbook.current_sheet_mut().cf_cache.borrow_mut().clear();
                    self.record_action(UndoAction::ConditionalFormatsReplaced {
                        sheet_idx,
                        old,
                        new: Vec::new(),
                    });
                }
                self.status_message = Some(format!("Cleared {} conditional format rule(s)", n));
            }
            ["cf", "list"] => {
                let rules = &self.workbook.current_sheet().conditional_formats;
                if rules.is_empty() {
                    self.status_message = Some("(no conditional formats)".to_string());
                } else {
                    let entries: Vec<String> = rules.iter().enumerate().map(|(i, r)| {
                        format!("#{} col {} when {}", i, crate::domain::Spreadsheet::column_label(r.column), r.predicate)
                    }).collect();
                    self.status_message = Some(entries.join("  |  "));
                }
            }
            // `:cf <col> <predicate> [bg=COLOR] [fg=COLOR] [bold] [underline]`
            // The predicate may use `_` for the cell value. To pass a spaced
            // predicate, wrap in quotes — but the palette splits on whitespace,
            // so we accept the predicate as a single token. Workaround: join
            // all middle tokens that don't look like style keys.
            ["cf", col_name, rest @ ..] if !rest.is_empty() => {
                if let Some(col) = Spreadsheet::parse_column_label(col_name) {
                    let mut predicate_parts = Vec::new();
                    let mut style = crate::domain::CellStyle::default();
                    for tok in rest {
                        if let Some(c) = tok.strip_prefix("bg=") {
                            style.bg_color = TerminalColor::from_name(c);
                        } else if let Some(c) = tok.strip_prefix("fg=") {
                            style.fg_color = TerminalColor::from_name(c);
                        } else if *tok == "bold" {
                            style.bold = true;
                        } else if *tok == "underline" {
                            style.underline = true;
                        } else {
                            predicate_parts.push(*tok);
                        }
                    }
                    if predicate_parts.is_empty() {
                        self.status_message = Some(
                            "Usage: cf <col> <predicate> [bg=color] [fg=color] [bold] [underline]"
                                .to_string(),
                        );
                    } else {
                        let predicate = predicate_parts.join(" ");
                        let sheet_idx = self.workbook.active_sheet;
                        let old = self.workbook.current_sheet().conditional_formats.clone();
                        let mut new = old.clone();
                        new.push(crate::domain::ConditionalFormat {
                            column: col,
                            predicate: predicate.clone(),
                            style,
                        });
                        {
                            let sheet = self.workbook.current_sheet_mut();
                            sheet.conditional_formats = new.clone();
                            sheet.cf_cache.borrow_mut().clear();
                        }
                        self.record_action(UndoAction::ConditionalFormatsReplaced {
                            sheet_idx,
                            old,
                            new,
                        });
                        self.status_message = Some(format!(
                            "Added cf for col {}: {}",
                            crate::domain::Spreadsheet::column_label(col),
                            predicate
                        ));
                    }
                } else {
                    self.status_message = Some(format!("Invalid column: {}", col_name));
                }
            }
            ["name", name, rest @ ..] if !rest.is_empty() => {
                // Join the remaining tokens so values with spaces (e.g.
                // `LAMBDA(x, x*2)`) survive intact.
                let value = rest.join(" ");
                self.workbook.set_name(name, &value);
                self.dirty = true;
                self.status_message = Some(format!("Named '{}' = {}", name, value));
            }
            ["unname", name] => {
                if self.workbook.remove_name(name) {
                    self.dirty = true;
                    self.status_message = Some(format!("Removed name '{}'", name));
                } else {
                    self.status_message = Some(format!("No such name: {}", name));
                }
            }
            ["names"] => {
                if self.workbook.named_ranges.is_empty() {
                    self.status_message = Some("(no named ranges)".to_string());
                } else {
                    let mut entries: Vec<String> = self
                        .workbook
                        .named_ranges
                        .iter()
                        .map(|(k, v)| format!("{}={}", k, v))
                        .collect();
                    entries.sort();
                    self.status_message = Some(entries.join("  "));
                }
            }
            ["autosave", "on"] => {
                crate::infrastructure::autosave::enable();
                self.status_message =
                    Some("Auto-save enabled (30s idle, writes to current filename)".to_string());
            }
            ["autosave", "off"] => {
                crate::infrastructure::autosave::disable();
                self.status_message = Some("Auto-save disabled".to_string());
            }
            ["regex", "on"] => {
                self.search_regex = true;
                self.status_message = Some("Regex search: on".to_string());
            }
            ["regex", "off"] => {
                self.search_regex = false;
                self.status_message = Some("Regex search: off".to_string());
            }
            ["case", "on"] | ["case-sensitive", "on"] => {
                self.search_case_sensitive = true;
                self.status_message = Some("Case-sensitive search: on".to_string());
            }
            ["case", "off"] | ["case-sensitive", "off"] => {
                self.search_case_sensitive = false;
                self.status_message = Some("Case-sensitive search: off".to_string());
            }
            _ => {
                self.status_message = Some(format!("Unknown command: {}", self.command_input));
            }
        }
        self.mode = AppMode::Normal;
        self.command_input.clear();
        self.cursor_position = 0;
    }

    pub fn command_suggestions(&self, max: usize) -> Vec<&'static str> {
        const ALL: &[&str] = &[
            "q", "q!", "w", "wq", "x", "e ",
            "export ", "import ", "import-append ",
            "ir", "dr", "ic", "dc",
            "insert row", "delete row", "insert col", "delete col",
            "sort asc", "sort desc",
            "freeze", "unfreeze",
            "format general", "format number", "format currency", "format percent",
            "bold", "underline",
            "color red", "color green", "color blue", "color yellow", "color none",
            "bg red", "bg green", "bg blue", "bg yellow", "bg none",
            "sheet new", "sheet delete", "sheet next", "sheet prev",
            "rename ",
            "comment ",
            "filter ", "unfilter",
            "hide row", "show rows",
            "recalc", "cache clear",
            "regex on", "regex off",
            "case on", "case off",
            "autosave on", "autosave off",
            "import-append ",
            "name ", "unname ", "names",
            "cf ", "cf clear", "cf list",
            "table create ", "table list",
            "pivot ", "goalseek ",
            "trace", "trace dependents",
        ];
        let q = self.command_input.trim().to_lowercase();
        if q.is_empty() {
            return ALL.iter().take(max).copied().collect();
        }
        ALL.iter()
            .filter(|c| c.starts_with(&q))
            .take(max)
            .copied()
            .collect()
    }

}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::domain::CellData;

    #[test]
    fn test_command_palette_insert_row() {
        let mut app = App::default();
        app.set_cell_with_undo(0, 0, CellData { value: "A1".to_string(), formula: None, format: None, comment: None, spill_anchor: None });
        app.set_cell_with_undo(1, 0, CellData { value: "A2".to_string(), formula: None, format: None, comment: None, spill_anchor: None });
        let orig_rows = app.workbook.current_sheet().rows;

        app.selected_row = 1;
        app.start_command_palette();
        app.command_input = "ir".to_string();
        app.execute_command();

        assert_eq!(app.workbook.current_sheet().rows, orig_rows + 1);
        assert!(matches!(app.mode, AppMode::Normal));
    }

    #[test]
    fn test_command_palette_delete_row() {
        let mut app = App::default();
        app.set_cell_with_undo(0, 0, CellData { value: "A1".to_string(), formula: None, format: None, comment: None, spill_anchor: None });
        app.set_cell_with_undo(1, 0, CellData { value: "A2".to_string(), formula: None, format: None, comment: None, spill_anchor: None });
        let orig_rows = app.workbook.current_sheet().rows;

        app.selected_row = 0;
        app.start_command_palette();
        app.command_input = "dr".to_string();
        app.execute_command();

        assert_eq!(app.workbook.current_sheet().rows, orig_rows - 1);
        assert_eq!(app.workbook.current_sheet().get_cell(0, 0).value, "A2"); // Shifted up
    }

    #[test]
    fn test_command_palette_format_currency() {
        let mut app = App::default();
        app.set_cell_with_undo(0, 0, CellData { value: "1234.5".to_string(), formula: None, format: None, comment: None, spill_anchor: None });

        app.selected_row = 0;
        app.selected_col = 0;
        app.start_command_palette();
        app.command_input = "format currency".to_string();
        app.execute_command();

        let cell = app.workbook.current_sheet().get_cell(0, 0);
        assert!(cell.format.is_some());
        assert!(matches!(cell.format.unwrap().number_format, NumberFormat::Currency { .. }));
    }

    #[test]
    fn test_command_palette_format_percentage() {
        let mut app = App::default();
        app.set_cell_with_undo(0, 0, CellData { value: "0.5".to_string(), formula: None, format: None, comment: None, spill_anchor: None });

        app.selected_row = 0;
        app.selected_col = 0;
        app.start_command_palette();
        app.command_input = "format percent 2".to_string();
        app.execute_command();

        let cell = app.workbook.current_sheet().get_cell(0, 0);
        assert!(cell.format.is_some());
        assert!(matches!(cell.format.unwrap().number_format, NumberFormat::Percentage { decimals: 2 }));
    }

    #[test]
    fn test_command_palette_unknown_command() {
        let mut app = App::default();
        app.start_command_palette();
        app.command_input = "foobar".to_string();
        app.execute_command();

        assert!(app.status_message.as_ref().unwrap().contains("Unknown command"));
        assert!(matches!(app.mode, AppMode::Normal));
    }

    #[test]
    fn test_command_bold() {
        let mut app = App::default();
        app.set_cell_with_undo(0, 0, CellData { value: "test".to_string(), formula: None, format: None, comment: None, spill_anchor: None });
        app.selected_row = 0;
        app.selected_col = 0;

        app.start_command_palette();
        app.command_input = "bold".to_string();
        app.execute_command();

        let cell = app.workbook.current_sheet().get_cell(0, 0);
        assert!(cell.format.as_ref().unwrap().style.bold);
    }

    #[test]
    fn test_command_underline() {
        let mut app = App::default();
        app.set_cell_with_undo(0, 0, CellData { value: "test".to_string(), formula: None, format: None, comment: None, spill_anchor: None });
        app.selected_row = 0;
        app.selected_col = 0;

        app.start_command_palette();
        app.command_input = "underline".to_string();
        app.execute_command();

        let cell = app.workbook.current_sheet().get_cell(0, 0);
        assert!(cell.format.as_ref().unwrap().style.underline);
    }

    #[test]
    fn test_command_color() {
        let mut app = App::default();
        app.set_cell_with_undo(0, 0, CellData { value: "test".to_string(), formula: None, format: None, comment: None, spill_anchor: None });
        app.selected_row = 0;
        app.selected_col = 0;

        app.start_command_palette();
        app.command_input = "color red".to_string();
        app.execute_command();

        let cell = app.workbook.current_sheet().get_cell(0, 0);
        assert_eq!(cell.format.as_ref().unwrap().style.fg_color, Some(TerminalColor::Red));
    }

    #[test]
    fn test_command_bg_color() {
        let mut app = App::default();
        app.set_cell_with_undo(0, 0, CellData { value: "test".to_string(), formula: None, format: None, comment: None, spill_anchor: None });
        app.selected_row = 0;
        app.selected_col = 0;

        app.start_command_palette();
        app.command_input = "bg blue".to_string();
        app.execute_command();

        let cell = app.workbook.current_sheet().get_cell(0, 0);
        assert_eq!(cell.format.as_ref().unwrap().style.bg_color, Some(TerminalColor::Blue));
    }

    #[test]
    fn test_command_color_none_clears() {
        let mut app = App::default();
        app.set_cell_with_undo(0, 0, CellData { value: "test".to_string(), formula: None, format: None, comment: None, spill_anchor: None });
        app.selected_row = 0;
        app.selected_col = 0;

        // Set color first
        app.set_selection_fg_color(Some(TerminalColor::Red));
        assert_eq!(app.workbook.current_sheet().get_cell(0, 0).format.as_ref().unwrap().style.fg_color, Some(TerminalColor::Red));

        // Clear via command
        app.start_command_palette();
        app.command_input = "color none".to_string();
        app.execute_command();

        let cell = app.workbook.current_sheet().get_cell(0, 0);
        assert_eq!(cell.format.as_ref().unwrap().style.fg_color, None);
    }

    #[test]
    fn test_add_sheet_command() {
        let mut app = App::default();
        app.start_command_palette();
        app.command_input = "sheet new".to_string();
        app.execute_command();

        assert_eq!(app.workbook.sheets.len(), 2);
        assert_eq!(app.workbook.active_sheet, 1); // Switched to new sheet
        assert_eq!(app.workbook.sheet_names[1], "Sheet2");
    }

    #[test]
    fn test_delete_sheet_command() {
        let mut app = App::default();
        // Add a second sheet
        app.workbook.add_sheet("Sheet2".to_string());
        app.workbook.active_sheet = 1;

        app.start_command_palette();
        app.command_input = "sheet delete".to_string();
        app.execute_command();

        assert_eq!(app.workbook.sheets.len(), 1);
        assert_eq!(app.workbook.active_sheet, 0);
    }

    #[test]
    fn test_rename_sheet_command() {
        let mut app = App::default();
        app.start_command_palette();
        app.command_input = "rename Revenue".to_string();
        app.execute_command();

        assert_eq!(app.workbook.sheet_names[0], "Revenue");
    }

    #[test]
    fn test_comment_command() {
        let mut app = App::default();
        app.set_cell_with_undo(0, 0, CellData { value: "Data".to_string(), formula: None, format: None, comment: None, spill_anchor: None });

        app.start_command_palette();
        app.command_input = "comment Test note".to_string();
        app.execute_command();

        // Comment text preserves case (it's user-facing prose).
        let cell = app.workbook.current_sheet().get_cell(0, 0);
        assert_eq!(cell.comment, Some("Test note".to_string()));
    }

    #[test]
    fn test_comment_clear_command() {
        let mut app = App::default();
        app.set_cell_with_undo(0, 0, CellData { value: "Data".to_string(), formula: None, format: None, comment: Some("note".to_string()), spill_anchor: None });

        app.start_command_palette();
        app.command_input = "comment clear".to_string();
        app.execute_command();

        let cell = app.workbook.current_sheet().get_cell(0, 0);
        assert_eq!(cell.comment, None);
    }

    #[test]
    fn test_filter_command() {
        let mut app = App::default();
        app.set_cell_with_undo(0, 0, CellData { value: "Yes".to_string(), formula: None, format: None, comment: None, spill_anchor: None });
        app.set_cell_with_undo(1, 0, CellData { value: "No".to_string(), formula: None, format: None, comment: None, spill_anchor: None });
        app.set_cell_with_undo(2, 0, CellData { value: "Yes".to_string(), formula: None, format: None, comment: None, spill_anchor: None });

        app.start_command_palette();
        app.command_input = "filter a yes".to_string();
        app.execute_command();

        assert!(!app.hidden_rows.contains(&0));
        assert!(app.hidden_rows.contains(&1));
        assert!(!app.hidden_rows.contains(&2));
    }

    #[test]
    fn test_unfilter_command() {
        let mut app = App::default();
        app.hidden_rows.insert(1);
        app.filter_column = Some(0);

        app.start_command_palette();
        app.command_input = "unfilter".to_string();
        app.execute_command();

        assert!(app.hidden_rows.is_empty());
        assert_eq!(app.filter_column, None);
    }

}
