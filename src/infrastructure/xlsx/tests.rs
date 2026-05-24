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
