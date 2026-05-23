//! `.xlsx` import and export.
//!
//! Reading uses `calamine` for robustness (handles real Excel quirks).
//! Writing is hand-rolled: a minimal-but-Excel-compliant package consisting
//! of `[Content_Types].xml`, `_rels/.rels`, `xl/workbook.xml`, sheet files,
//! and an optional shared-strings table.
//!
//! We persist cell formulas where present (Excel reads tshts-saved formulas
//! correctly for the operator/function subset that overlaps). Sheet
//! structure, number formats, and named ranges round-trip; cell formatting
//! (colors, bold) is a TODO follow-up.

use std::io::Write;

use crate::domain::{CellData, CellFormat, CellStyle, NumberFormat, Spreadsheet, TerminalColor, Workbook};

/// Hard caps on imported sheet geometry. Calamine reports the declared
/// dimensions of the file, which a hostile workbook can inflate to Excel's
/// `XFD1048576` to make us allocate a 16 GB scrollback region. We clamp to
/// values that still cover any realistic spreadsheet.
const MAX_IMPORT_ROWS: usize = 200_000;
const MAX_IMPORT_COLS: usize = 1_024;

/// Reject obviously-hostile .xlsx files before handing them to calamine.
/// A 1 GiB on-disk file would force calamine to load gigabytes into memory
/// (the SharedStrings table alone can be huge in pathological cases).
/// 256 MiB covers any realistic workbook by a wide margin.
const MAX_XLSX_FILE_BYTES: u64 = 256 * 1024 * 1024;

/// Cap on the number of cells we'll materialize from a single sheet. Once
/// past this, we stop adding cells (the user gets a partial import rather
/// than the process OOMing). This is a second line of defense beyond the
/// row/col geometry caps — a single sparsely-defined sheet could still
/// reference a billion (row, col) coordinates through used_cells iteration.
const MAX_IMPORT_CELLS_PER_SHEET: usize = 5_000_000;

/// Read an `.xlsx` file into a Workbook.
pub fn load_xlsx(path: &str) -> Result<Workbook, String> {
    use calamine::{open_workbook_auto, Data, Reader};
    // Refuse files that are obviously oversized before letting calamine
    // touch them. This is the cheap zip-bomb defense: a 100 KB hostile zip
    // expanding to 100 GB is the textbook case, and rejecting outsize files
    // up-front beats post-hoc OOM detection.
    if let Ok(meta) = std::fs::metadata(path)
        && meta.len() > MAX_XLSX_FILE_BYTES {
            return Err(format!(
                "xlsx too large: {} bytes (limit {})",
                meta.len(),
                MAX_XLSX_FILE_BYTES
            ));
        }
    let mut wb = open_workbook_auto(path).map_err(|e| format!("xlsx open: {}", e))?;
    let sheet_names = wb.sheet_names().to_vec();
    if sheet_names.is_empty() {
        return Err("xlsx has no sheets".to_string());
    }
    let sheet_count = sheet_names.len() as u32;
    let mut out = Workbook {
        version: crate::domain::models::WORKBOOK_SCHEMA_VERSION,
        sheets: Vec::with_capacity(sheet_names.len()),
        sheet_names: sheet_names.clone(),
        active_sheet: 0,
        named_ranges: std::collections::HashMap::new(),
        cross_sheet_dependents: std::collections::HashMap::new(),
        cross_sheet_dependencies: std::collections::HashMap::new(),
        cells_with_qualified_refs: std::collections::HashSet::new(),
        dirty: std::collections::HashSet::new(),
        sheet_ids: (0..sheet_count).map(crate::domain::models::SheetId).collect(),
        next_sheet_id: sheet_count,
        graph: crate::domain::models::WorkbookGraph::new(),
        cell_purities: std::collections::HashMap::new(),
    };
    for name in &sheet_names {
        let range = wb
            .worksheet_range(name)
            .map_err(|e| format!("xlsx sheet {}: {}", name, e))?;
        let formulas = wb.worksheet_formula(name).ok();
        let mut sheet = Spreadsheet::default();
        let (h, w) = range.get_size();
        // Clamp declared geometry to defensive bounds. A pathological file
        // with one cell at `XFD1048576` would otherwise set rows ≈ 1M, cols
        // ≈ 16K — terminal rendering allocates per cell.
        sheet.rows = h.max(100).min(MAX_IMPORT_ROWS);
        sheet.cols = w.max(26).min(MAX_IMPORT_COLS);
        // Build (row, col) → formula. `used_cells` returns coords relative
        // to the range's `start()`, so we must offset back to absolute.
        let mut formula_map: std::collections::HashMap<(usize, usize), String> =
            std::collections::HashMap::new();
        if let Some(fr) = formulas.as_ref() {
            let (start_r, start_c) = fr.start().unwrap_or((0, 0));
            for (r, c, f) in fr.used_cells() {
                if !f.is_empty() {
                    formula_map.insert(
                        (start_r as usize + r, start_c as usize + c),
                        f.clone(),
                    );
                }
            }
        }
        // Same offset for the value range.
        let (val_r, val_c) = range.start().unwrap_or((0, 0));
        let mut imported_cells = 0usize;
        for (r, c, cell) in range.used_cells() {
            if imported_cells >= MAX_IMPORT_CELLS_PER_SHEET {
                break;
            }
            let value = match cell {
                Data::Empty => continue,
                Data::String(s) => s.clone(),
                Data::Float(f) => f.to_string(),
                Data::Int(i) => i.to_string(),
                // Booleans become Excel's display strings so formulas like
                // `=A1=TRUE` work, instead of the raw 0/1 numerals.
                Data::Bool(b) => if *b { "TRUE".to_string() } else { "FALSE".to_string() },
                // DateTime is an Excel serial. Render as ISO so the cell
                // displays as a date rather than the raw float "45000.0".
                Data::DateTime(d) => {
                    let serial = d.as_f64();
                    let (y, m, day) = crate::domain::parser::serial_to_date_pub(serial);
                    format!("{:04}-{:02}-{:02}", y, m, day)
                }
                Data::DateTimeIso(s) => s.clone(),
                Data::DurationIso(s) => s.clone(),
                Data::Error(e) => format!("#{:?}!", e),
            };
            let abs_r = val_r as usize + r;
            let abs_c = val_c as usize + c;
            // Drop cells outside the clamped sheet geometry so a stray cell
            // at `XFD1048576` can't trip rendering bounds.
            if abs_r >= sheet.rows || abs_c >= sheet.cols {
                continue;
            }
            imported_cells += 1;
            // Strip `_xlfn.` and `_xlfn._xlws.` prefixes Excel adds to
            // modern function names (XLOOKUP, FILTER, etc.) so the
            // formula evaluator recognizes them.
            let formula = formula_map.get(&(abs_r, abs_c)).map(|f| {
                let cleaned = strip_xlfn_prefixes(f);
                format!("={}", cleaned)
            });
            let cd = CellData {
                value,
                formula,
                format: None,
                comment: None,
                spill_anchor: None,
            };
            sheet.cells.insert((abs_r, abs_c), cd);
        }
        sheet.rebuild_dependencies();
        out.sheets.push(sheet);
    }
    // Named ranges (definedNames). Calamine exposes them via `defined_names()`.
    for (name, value) in wb.defined_names() {
        let upper = name.to_uppercase();
        out.named_ranges.insert(upper.clone(), value.clone());
        for s in &mut out.sheets {
            s.named_ranges.insert(upper.clone(), value.clone());
        }
    }
    // Build the cross-sheet dep graph so subsequent edits propagate
    // correctly. Same-sheet graphs were built per-sheet above.
    out.rebuild_cross_sheet_deps();
    // Pre-build the unified workbook graph too so .xlsx files behave
    // identically to .tshts on first recalc. The lazy build inside
    // recalc_via_graph would otherwise run on the first edit.
    out.build_dep_graph_from_scratch();
    Ok(out)
}

/// Strip Excel's `_xlfn.` and `_xlfn._xlws.` prefixes from function names.
/// These prefixes appear in formulas saved by newer Excel versions to
/// disambiguate functions added after 2007. Different writers normalize
/// the case differently (some emit `_xlfn.`, some `_XLFN.`, occasionally
/// mixed like `_Xlfn.`) so we scan case-insensitively.
fn strip_xlfn_prefixes(formula: &str) -> String {
    fn ascii_prefix_match(rest: &str, prefix: &[u8]) -> bool {
        // The prefixes are ASCII. Use byte-level compare so we don't try to
        // slice through a multi-byte UTF-8 boundary when the prefix doesn't
        // match. If the first `prefix.len()` bytes do match ASCII-wise, the
        // slice itself is on a char boundary (each byte is single-byte ASCII).
        let bytes = rest.as_bytes();
        bytes.len() >= prefix.len() && bytes[..prefix.len()].eq_ignore_ascii_case(prefix)
    }
    let mut out = String::with_capacity(formula.len());
    let mut rest = formula;
    while !rest.is_empty() {
        if ascii_prefix_match(rest, b"_xlfn._xlws.") {
            rest = &rest["_xlfn._xlws.".len()..];
        } else if ascii_prefix_match(rest, b"_xlfn.") {
            rest = &rest["_xlfn.".len()..];
        } else {
            let ch = rest.chars().next().unwrap();
            out.push(ch);
            rest = &rest[ch.len_utf8()..];
        }
    }
    out
}

#[cfg(test)]
mod xlfn_tests {
    use super::strip_xlfn_prefixes;

    #[test]
    fn strips_lower_and_upper() {
        assert_eq!(strip_xlfn_prefixes("_xlfn.XLOOKUP(A1,B:B,C:C)"), "XLOOKUP(A1,B:B,C:C)");
        assert_eq!(strip_xlfn_prefixes("_XLFN.XLOOKUP(A1,B:B,C:C)"), "XLOOKUP(A1,B:B,C:C)");
    }

    #[test]
    fn strips_mixed_case() {
        assert_eq!(strip_xlfn_prefixes("_Xlfn.XLOOKUP(A1,B:B,C:C)"), "XLOOKUP(A1,B:B,C:C)");
        assert_eq!(strip_xlfn_prefixes("_xLfN.XLOOKUP(A1,B:B,C:C)"), "XLOOKUP(A1,B:B,C:C)");
    }

    #[test]
    fn strips_compound_prefix() {
        assert_eq!(
            strip_xlfn_prefixes("_xlfn._xlws.FILTER(A:A,B:B>0)"),
            "FILTER(A:A,B:B>0)"
        );
        assert_eq!(
            strip_xlfn_prefixes("_XLFN._XLWS.FILTER(A:A,B:B>0)"),
            "FILTER(A:A,B:B>0)"
        );
    }

    #[test]
    fn preserves_non_xlfn() {
        assert_eq!(strip_xlfn_prefixes("=SUM(A1:A10)"), "=SUM(A1:A10)");
        assert_eq!(strip_xlfn_prefixes(""), "");
    }

    #[test]
    fn preserves_utf8() {
        // Multi-byte UTF-8 must round-trip unchanged. The old byte-then-as-char
        // version would have mangled this.
        assert_eq!(strip_xlfn_prefixes("=CONCAT(\"héllo\",A1)"), "=CONCAT(\"héllo\",A1)");
        assert_eq!(strip_xlfn_prefixes("_xlfn.XLOOKUP(\"日本\",A:A,B:B)"), "XLOOKUP(\"日本\",A:A,B:B)");
    }
}

/// Builds a deduplicated style table for the workbook. Index 0 is always
/// the default (no styling). Each unique non-default `CellFormat` gets an
/// index that we'll emit in `<c s="N">` and that we'll define in
/// `xl/styles.xml`.
fn build_style_table(workbook: &Workbook) -> Vec<CellFormat> {
    let mut styles: Vec<CellFormat> = vec![CellFormat::default()];
    for sheet in &workbook.sheets {
        for cd in sheet.cells.values() {
            if let Some(fmt) = &cd.format
                && !styles.iter().any(|s| s == fmt) {
                    styles.push(fmt.clone());
                }
        }
    }
    styles
}

/// Look up the style-table index for a cell's format. Returns 0 (default)
/// when no format is set.
fn style_index(styles: &[CellFormat], format: Option<&CellFormat>) -> usize {
    match format {
        None => 0,
        Some(fmt) => styles.iter().position(|s| s == fmt).unwrap_or(0),
    }
}

/// Map a `TerminalColor` to an Excel-friendly ARGB hex (FF + RGB).
fn color_to_argb(color: &TerminalColor) -> &'static str {
    match color {
        TerminalColor::Black => "FF000000",
        TerminalColor::Red => "FFC00000",
        TerminalColor::Green => "FF008000",
        TerminalColor::Yellow => "FFFFFF00",
        TerminalColor::Blue => "FF0000FF",
        TerminalColor::Magenta => "FFFF00FF",
        TerminalColor::Cyan => "FF00FFFF",
        TerminalColor::White => "FFFFFFFF",
        TerminalColor::DarkGray => "FF808080",
        TerminalColor::LightRed => "FFFF6060",
        TerminalColor::LightGreen => "FF80FF80",
        TerminalColor::LightYellow => "FFFFFF80",
        TerminalColor::LightBlue => "FF80B0FF",
        TerminalColor::LightMagenta => "FFFF80FF",
        TerminalColor::LightCyan => "FF80FFFF",
    }
}

/// Build the `xl/styles.xml` body. Each non-default style maps to a unique
/// numFmt (if needed) + font + fill + cellXf row.
fn build_styles_xml(styles: &[CellFormat]) -> String {
    let mut numfmts: Vec<String> = Vec::new(); // formatCode strings (index = 164 + i)
    let mut fonts: Vec<&CellStyle> = Vec::new();
    let mut fills: Vec<Option<&TerminalColor>> = vec![None, None]; // 0,1 are reserved by Excel
    let mut xfs: Vec<(usize, usize, usize, usize)> = Vec::new(); // (numFmtId, fontIdx, fillIdx, applyAlignment)

    // Default font (font 0) and fill (fill 0,1) are mandatory.
    fonts.push(&CELL_STYLE_DEFAULT);

    for style in styles {
        // numFmt
        let numfmt_id = match &style.number_format {
            NumberFormat::General => 0u32, // built-in "General"
            NumberFormat::Number { decimals, thousands_sep } => {
                let code = if *thousands_sep {
                    format!("#,##0.{}", "0".repeat(*decimals as usize))
                } else {
                    format!("0.{}", "0".repeat(*decimals as usize))
                };
                // Strip trailing dot if decimals == 0
                let code = if code.ends_with('.') { code.trim_end_matches('.').to_string() } else { code };
                if let Some(i) = numfmts.iter().position(|n| n == &code) {
                    164 + i as u32
                } else {
                    numfmts.push(code);
                    164 + (numfmts.len() - 1) as u32
                }
            }
            NumberFormat::Currency { symbol, decimals } => {
                let code = format!("\"{}\"#,##0.{}", symbol, "0".repeat(*decimals as usize));
                let code = if code.ends_with('.') { code.trim_end_matches('.').to_string() } else { code };
                if let Some(i) = numfmts.iter().position(|n| n == &code) {
                    164 + i as u32
                } else {
                    numfmts.push(code);
                    164 + (numfmts.len() - 1) as u32
                }
            }
            NumberFormat::Percentage { decimals } => {
                let code = if *decimals == 0 {
                    "0%".to_string()
                } else {
                    format!("0.{}%", "0".repeat(*decimals as usize))
                };
                if let Some(i) = numfmts.iter().position(|n| n == &code) {
                    164 + i as u32
                } else {
                    numfmts.push(code);
                    164 + (numfmts.len() - 1) as u32
                }
            }
        };

        // Font: dedupe by (bold, underline, fg_color).
        let font_idx = match fonts.iter().position(|f| {
            f.bold == style.style.bold
                && f.underline == style.style.underline
                && f.fg_color == style.style.fg_color
        }) {
            Some(i) => i,
            None => {
                fonts.push(&style.style);
                fonts.len() - 1
            }
        };

        // Fill: only bg_color counts; pattern is "solid".
        let fill_idx = match fills.iter().position(|f| f == &style.style.bg_color.as_ref()) {
            Some(i) => i,
            None => {
                fills.push(style.style.bg_color.as_ref());
                fills.len() - 1
            }
        };

        xfs.push((numfmt_id as usize, font_idx, fill_idx, 0));
    }

    let mut s = String::from(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
"#,
    );

    // numFmts (user-defined start at 164)
    if !numfmts.is_empty() {
        s.push_str(&format!("<numFmts count=\"{}\">\n", numfmts.len()));
        for (i, code) in numfmts.iter().enumerate() {
            s.push_str(&format!(
                "<numFmt numFmtId=\"{}\" formatCode=\"{}\"/>\n",
                164 + i,
                xml_escape(code)
            ));
        }
        s.push_str("</numFmts>\n");
    }

    // fonts: emit font 0 as default, plus each unique non-default.
    s.push_str(&format!("<fonts count=\"{}\">\n", fonts.len()));
    for (i, f) in fonts.iter().enumerate() {
        s.push_str("<font>");
        s.push_str("<sz val=\"11\"/>");
        if i > 0 && f.bold {
            s.push_str("<b/>");
        }
        if i > 0 && f.underline {
            s.push_str("<u/>");
        }
        if i > 0
            && let Some(c) = &f.fg_color {
                s.push_str(&format!("<color rgb=\"{}\"/>", color_to_argb(c)));
            }
        s.push_str("<name val=\"Calibri\"/>");
        s.push_str("</font>\n");
    }
    s.push_str("</fonts>\n");

    // fills: indices 0 and 1 are reserved (none + gray125). Real fills start at 2.
    s.push_str(&format!("<fills count=\"{}\">\n", fills.len()));
    for (i, fill) in fills.iter().enumerate() {
        match i {
            0 => s.push_str("<fill><patternFill patternType=\"none\"/></fill>\n"),
            1 => s.push_str("<fill><patternFill patternType=\"gray125\"/></fill>\n"),
            _ => match fill {
                Some(c) => s.push_str(&format!(
                    "<fill><patternFill patternType=\"solid\"><fgColor rgb=\"{}\"/></patternFill></fill>\n",
                    color_to_argb(c)
                )),
                None => s.push_str("<fill><patternFill patternType=\"none\"/></fill>\n"),
            },
        }
    }
    s.push_str("</fills>\n");

    // borders: just one default.
    s.push_str("<borders count=\"1\"><border><left/><right/><top/><bottom/><diagonal/></border></borders>\n");

    // cellStyleXfs: required parent for cellXfs.
    s.push_str("<cellStyleXfs count=\"1\"><xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\"/></cellStyleXfs>\n");

    // cellXfs: index 0 is default, then one per style.
    s.push_str(&format!("<cellXfs count=\"{}\">\n", xfs.len()));
    for (i, (numfmt, font, fill, _)) in xfs.iter().enumerate() {
        let apply_num = if *numfmt != 0 { " applyNumberFormat=\"1\"" } else { "" };
        let apply_font = if *font != 0 { " applyFont=\"1\"" } else { "" };
        let apply_fill = if *fill > 1 { " applyFill=\"1\"" } else { "" };
        if i == 0 {
            s.push_str("<xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\" xfId=\"0\"/>\n");
        } else {
            s.push_str(&format!(
                "<xf numFmtId=\"{}\" fontId=\"{}\" fillId=\"{}\" borderId=\"0\" xfId=\"0\"{}{}{}/>\n",
                numfmt, font, fill, apply_num, apply_font, apply_fill
            ));
        }
    }
    s.push_str("</cellXfs>\n");

    s.push_str("</styleSheet>");
    s
}

const CELL_STYLE_DEFAULT: CellStyle = CellStyle {
    bold: false,
    underline: false,
    fg_color: None,
    bg_color: None,
};

/// Write a Workbook as `.xlsx`. Each sheet gets its own XML; formulas, named
/// ranges, and cell formatting (bold/underline/fg+bg colors, number formats)
/// are preserved.
pub fn save_xlsx(workbook: &Workbook, path: &str) -> Result<(), String> {
    // Build the zip into an in-memory buffer, then atomic_write so a crash
    // mid-zip-write can't leave a corrupt half-written .xlsx where the
    // user's previous good file used to be.
    let buf = std::io::Cursor::new(Vec::<u8>::with_capacity(64 * 1024));
    let mut zip = zip::ZipWriter::new(buf);
    // zip 2.x renamed FileOptions::default() ergonomics to SimpleFileOptions
    // (the old name was generic over the compression-type parameter and
    // required turbofish to instantiate).
    let opts = zip::write::SimpleFileOptions::default()
        .compression_method(zip::CompressionMethod::Deflated);
    let styles = build_style_table(workbook);

    // [Content_Types].xml
    zip.start_file("[Content_Types].xml", opts).map_err(|e| e.to_string())?;
    let mut ct = String::from(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
"#,
    );
    for i in 1..=workbook.sheets.len() {
        ct.push_str(&format!(
            "<Override PartName=\"/xl/worksheets/sheet{}.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml\"/>\n",
            i
        ));
    }
    ct.push_str("</Types>");
    zip.write_all(ct.as_bytes()).map_err(|e| e.to_string())?;

    // _rels/.rels
    zip.start_file("_rels/.rels", opts).map_err(|e| e.to_string())?;
    zip.write_all(br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>"#).map_err(|e| e.to_string())?;

    // xl/_rels/workbook.xml.rels — sheets + styles
    zip.start_file("xl/_rels/workbook.xml.rels", opts)
        .map_err(|e| e.to_string())?;
    let mut rels = String::from(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
"#,
    );
    for i in 1..=workbook.sheets.len() {
        rels.push_str(&format!(
            "<Relationship Id=\"rId{}\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" Target=\"worksheets/sheet{}.xml\"/>\n",
            i, i
        ));
    }
    let styles_rid = workbook.sheets.len() + 1;
    rels.push_str(&format!(
        "<Relationship Id=\"rId{}\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles\" Target=\"styles.xml\"/>\n",
        styles_rid
    ));
    rels.push_str("</Relationships>");
    zip.write_all(rels.as_bytes()).map_err(|e| e.to_string())?;

    // xl/styles.xml
    zip.start_file("xl/styles.xml", opts).map_err(|e| e.to_string())?;
    zip.write_all(build_styles_xml(&styles).as_bytes())
        .map_err(|e| e.to_string())?;

    // xl/workbook.xml
    zip.start_file("xl/workbook.xml", opts).map_err(|e| e.to_string())?;
    let mut wb = String::from(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
<sheets>
"#,
    );
    for (i, name) in workbook.sheet_names.iter().enumerate() {
        wb.push_str(&format!(
            "<sheet name=\"{}\" sheetId=\"{}\" r:id=\"rId{}\"/>\n",
            xml_escape(name),
            i + 1,
            i + 1
        ));
    }
    wb.push_str("</sheets>");
    if !workbook.named_ranges.is_empty() {
        wb.push_str("\n<definedNames>\n");
        for (name, value) in &workbook.named_ranges {
            wb.push_str(&format!(
                "<definedName name=\"{}\">{}</definedName>\n",
                xml_escape(name),
                xml_escape(value)
            ));
        }
        wb.push_str("</definedNames>");
    }
    wb.push_str("\n</workbook>");
    zip.write_all(wb.as_bytes()).map_err(|e| e.to_string())?;

    // xl/worksheets/sheet*.xml
    for (i, sheet) in workbook.sheets.iter().enumerate() {
        zip.start_file(format!("xl/worksheets/sheet{}.xml", i + 1), opts)
            .map_err(|e| e.to_string())?;
        let mut buf = String::from(
            r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<sheetData>
"#,
        );
        // Group by row for correct sheetData ordering.
        let mut rows: std::collections::BTreeMap<usize, Vec<(usize, &CellData)>> =
            std::collections::BTreeMap::new();
        for (&(r, c), cd) in &sheet.cells {
            rows.entry(r).or_default().push((c, cd));
        }
        for (r, mut cells) in rows {
            cells.sort_by_key(|(c, _)| *c);
            buf.push_str(&format!("<row r=\"{}\">", r + 1));
            for (c, cd) in cells {
                let cell_ref = format!("{}{}", Spreadsheet::column_label(c), r + 1);
                let is_num = cd.value.parse::<f64>().is_ok();
                let s_idx = style_index(&styles, cd.format.as_ref());
                let s_attr = if s_idx == 0 { String::new() } else { format!(" s=\"{}\"", s_idx) };
                if let Some(formula) = cd.formula.as_ref().and_then(|f| f.strip_prefix('=')) {
                    let ty = if is_num { "n" } else { "str" };
                    buf.push_str(&format!(
                        "<c r=\"{}\"{} t=\"{}\"><f>{}</f><v>{}</v></c>",
                        cell_ref,
                        s_attr,
                        ty,
                        xml_escape(formula),
                        xml_escape(&cd.value)
                    ));
                } else if is_num {
                    // Numeric branch is currently safe (parse::<f64>().is_ok()
                    // excludes XML metacharacters) but escape defensively so a
                    // future loosening of `is_num` can't introduce injection.
                    buf.push_str(&format!(
                        "<c r=\"{}\"{}><v>{}</v></c>",
                        cell_ref, s_attr, xml_escape(&cd.value)
                    ));
                } else {
                    buf.push_str(&format!(
                        "<c r=\"{}\"{} t=\"inlineStr\"><is><t>{}</t></is></c>",
                        cell_ref,
                        s_attr,
                        xml_escape(&cd.value)
                    ));
                }
            }
            buf.push_str("</row>\n");
        }
        buf.push_str("</sheetData>\n</worksheet>");
        zip.write_all(buf.as_bytes()).map_err(|e| e.to_string())?;
    }

    let cursor = zip.finish().map_err(|e| e.to_string())?;
    let bytes = cursor.into_inner();
    crate::infrastructure::atomic::atomic_write(path, &bytes).map_err(|e| e.to_string())?;
    Ok(())
}

fn xml_escape(s: &str) -> String {
    s.replace('&', "&amp;")
        .replace('<', "&lt;")
        .replace('>', "&gt;")
        .replace('"', "&quot;")
        .replace('\'', "&apos;")
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn strip_xlfn_prefix() {
        assert_eq!(strip_xlfn_prefixes("XLOOKUP(1, A:A, B:B)"), "XLOOKUP(1, A:A, B:B)");
        assert_eq!(strip_xlfn_prefixes("_xlfn.XLOOKUP(1, A:A, B:B)"), "XLOOKUP(1, A:A, B:B)");
        assert_eq!(strip_xlfn_prefixes("_xlfn._xlws.FILTER(A1:A5, A1:A5>0)"), "FILTER(A1:A5, A1:A5>0)");
        assert_eq!(
            strip_xlfn_prefixes("SUM(_xlfn.MAP(A1:A3, _xlfn.LAMBDA(x, x*2)))"),
            "SUM(MAP(A1:A3, LAMBDA(x, x*2)))"
        );
    }

    #[test]
    fn xlsx_roundtrip_named_ranges() {
        let mut wb = Workbook::default();
        wb.sheets[0].set_cell(0, 0, CellData {
            value: "5".to_string(), formula: None, format: None, comment: None,
        spill_anchor: None,
        });
        wb.set_name("MYVAL", "A1");

        let tmp = tempfile::NamedTempFile::new().unwrap();
        let path = tmp.path().with_extension("xlsx");
        let path_str = path.to_str().unwrap();
        save_xlsx(&wb, path_str).unwrap();
        let loaded = load_xlsx(path_str).unwrap();

        // Calamine returns names possibly prefixed with the sheet name; we
        // just check the value was carried across.
        let has_myval = loaded.named_ranges.keys().any(|k| k.to_uppercase().contains("MYVAL"));
        assert!(has_myval, "named range MYVAL missing from loaded workbook: keys = {:?}", loaded.named_ranges.keys().collect::<Vec<_>>());
    }

    #[test]
    #[ignore = "requires python3 + openpyxl; run with `cargo test --release xlsx_opens_with_openpyxl -- --ignored`"]
    fn xlsx_opens_with_openpyxl() {
        // Validate tshts xlsx against a real Excel reader. Verifies the
        // file opens, the values are present, formulas are recognized, and
        // bold styling survived the round-trip.
        let mut wb = Workbook::default();
        wb.sheets[0].set_cell(0, 0, CellData {
            value: "Bold red".to_string(),
            formula: None,
            format: Some(CellFormat {
                number_format: NumberFormat::General,
                style: CellStyle {
                    bold: true,
                    underline: false,
                    fg_color: Some(TerminalColor::Red),
                    bg_color: None,
                },
            }),
            comment: None,
            spill_anchor: None,
        });
        wb.sheets[0].set_cell(0, 1, CellData {
            value: "42".to_string(),
            formula: None,
            format: None,
            comment: None,
            spill_anchor: None,
        });
        wb.sheets[0].set_cell(0, 2, CellData {
            value: "84".to_string(),
            formula: Some("=B1*2".to_string()),
            format: None,
            comment: None,
            spill_anchor: None,
        });

        let tmp = tempfile::NamedTempFile::new().unwrap();
        let path = tmp.path().with_extension("xlsx");
        save_xlsx(&wb, path.to_str().unwrap()).expect("save_xlsx");

        let script = format!(
            r#"
import openpyxl, sys
wb = openpyxl.load_workbook(r'{}')
ws = wb.active
assert ws['A1'].value == 'Bold red', f'A1 got {{ws["A1"].value!r}}'
assert ws['A1'].font.bold is True, 'A1 not bold'
assert ws['B1'].value == 42, f'B1 got {{ws["B1"].value!r}}'
# Formula cell: C1 should have a formula AND a cached value of 84.
assert ws['C1'].value == '=B1*2', f'C1 formula got {{ws["C1"].value!r}}'
print('OK')
"#,
            path.to_str().unwrap()
        );
        let output = std::process::Command::new("python3")
            .arg("-c")
            .arg(&script)
            .output()
            .expect("python3 not installed");
        let stdout = String::from_utf8_lossy(&output.stdout);
        let stderr = String::from_utf8_lossy(&output.stderr);
        assert!(
            output.status.success() && stdout.trim() == "OK",
            "openpyxl validation failed.\nstdout: {}\nstderr: {}",
            stdout,
            stderr
        );
    }

    #[test]
    fn xlsx_archive_has_expected_parts() {
        // Save a workbook, then crack open the zip to verify the parts an
        // Excel/LibreOffice reader expects are all present and contain
        // the right XML.
        use std::io::Read;
        let mut wb = Workbook::default();
        wb.sheets[0].set_cell(0, 0, CellData {
            value: "Hello".to_string(),
            formula: None,
            format: Some(CellFormat {
                number_format: NumberFormat::General,
                style: CellStyle {
                    bold: true,
                    underline: false,
                    fg_color: Some(TerminalColor::Red),
                    bg_color: None,
                },
            }),
            comment: None,
            spill_anchor: None,
        });
        wb.set_name("MYRANGE", "A1:A3");
        let tmp = tempfile::NamedTempFile::new().unwrap();
        let path = tmp.path().with_extension("xlsx");
        let path_str = path.to_str().unwrap();
        save_xlsx(&wb, path_str).expect("save_xlsx");

        let file = std::fs::File::open(path_str).expect("open xlsx");
        let mut zip = zip::ZipArchive::new(file).expect("valid zip");
        let names: Vec<String> = (0..zip.len())
            .map(|i| zip.by_index(i).unwrap().name().to_string())
            .collect();
        for part in &[
            "[Content_Types].xml",
            "_rels/.rels",
            "xl/_rels/workbook.xml.rels",
            "xl/workbook.xml",
            "xl/styles.xml",
            "xl/worksheets/sheet1.xml",
        ] {
            assert!(
                names.iter().any(|n| n == part),
                "xlsx missing part {}; present: {:?}",
                part,
                names
            );
        }
        let mut styles = String::new();
        zip.by_name("xl/styles.xml").unwrap().read_to_string(&mut styles).unwrap();
        assert!(styles.contains("<b/>"), "styles.xml missing <b/>");
        assert!(styles.contains("FFC00000"), "styles.xml missing red ARGB");
        let mut wbxml = String::new();
        zip.by_name("xl/workbook.xml").unwrap().read_to_string(&mut wbxml).unwrap();
        assert!(wbxml.contains("MYRANGE"), "workbook.xml missing named range");
        let mut sheet1 = String::new();
        zip.by_name("xl/worksheets/sheet1.xml").unwrap().read_to_string(&mut sheet1).unwrap();
        assert!(sheet1.contains(" s=\""), "sheet1.xml missing per-cell style attribute");
    }

    #[test]
    fn xlsx_writes_styled_cells_without_panic() {
        // Verify a styled cell survives save_xlsx (validates the styles.xml
        // we emit). Reading styles back is not implemented yet, but the
        // value should round-trip and the file must be valid xlsx.
        let mut wb = Workbook::default();
        let bold_red = CellFormat {
            number_format: NumberFormat::General,
            style: CellStyle {
                bold: true,
                underline: false,
                fg_color: Some(TerminalColor::Red),
                bg_color: None,
            },
        };
        wb.sheets[0].set_cell(0, 0, CellData {
            value: "BOLD".to_string(),
            formula: None,
            format: Some(bold_red),
            comment: None,
        spill_anchor: None,
        });
        let currency = CellFormat {
            number_format: NumberFormat::Currency {
                symbol: "$".to_string(),
                decimals: 2,
            },
            style: CellStyle::default(),
        };
        wb.sheets[0].set_cell(1, 0, CellData {
            value: "1234.5".to_string(),
            formula: None,
            format: Some(currency),
            comment: None,
        spill_anchor: None,
        });

        let tmp = tempfile::NamedTempFile::new().unwrap();
        let path = tmp.path().with_extension("xlsx");
        let path_str = path.to_str().unwrap();
        save_xlsx(&wb, path_str).expect("save_xlsx should not fail with styles");
        let loaded = load_xlsx(path_str).expect("load_xlsx should re-read the file");
        assert_eq!(loaded.sheets[0].get_cell(0, 0).value, "BOLD");
        assert_eq!(loaded.sheets[0].get_cell(1, 0).value, "1234.5");
    }

    #[test]
    fn xlsx_roundtrip_basic() {
        let mut wb = Workbook::default();
        wb.sheets[0].set_cell(
            0,
            0,
            CellData {
                value: "Hello".to_string(),
                formula: None,
                format: None,
                comment: None,
            spill_anchor: None,
            },
        );
        wb.sheets[0].set_cell(
            0,
            1,
            CellData {
                value: "42".to_string(),
                formula: None,
                format: None,
                comment: None,
            spill_anchor: None,
            },
        );
        wb.sheets[0].set_cell(
            1,
            0,
            CellData {
                value: "84".to_string(),
                formula: Some("=B1*2".to_string()),
                format: None,
                comment: None,
            spill_anchor: None,
            },
        );

        let tmp = tempfile::NamedTempFile::new().unwrap();
        let path = tmp.path().with_extension("xlsx");
        let path_str = path.to_str().unwrap();
        save_xlsx(&wb, path_str).unwrap();

        let loaded = load_xlsx(path_str).unwrap();
        assert_eq!(loaded.sheets.len(), 1);
        assert_eq!(loaded.sheets[0].get_cell(0, 0).value, "Hello");
        assert_eq!(loaded.sheets[0].get_cell(0, 1).value, "42");
        // Formula round-trips.
        assert_eq!(
            loaded.sheets[0].get_cell(1, 0).formula.as_deref(),
            Some("=B1*2")
        );
    }
}
