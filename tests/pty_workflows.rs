//! Multi-step user workflows driven through a real PTY.

// Test names use uppercase letters to mirror vim keys (e.g. `..._via_S`).
#![allow(non_snake_case)]

mod common;

use common::Harness;
use std::time::Duration;

fn quit_force(h: &mut Harness) {
    h.send_text(":q!");
    h.send_enter();
    assert_eq!(h.wait_for_exit(Duration::from_secs(3)), Some(0));
}

// ----- Persistence round-trip -----

#[test]
fn workflow_xlsx_roundtrip() {
    let dir = tempfile::tempdir().unwrap();
    let path = dir.path().join("book.xlsx");
    let path_str = path.to_str().unwrap();
    let mut h = Harness::new();
    h.send_text("ifoo"); h.send_enter();
    h.send_text("ibar"); h.send_enter();
    h.send_text(":w ");
    h.send_text(path_str);
    h.send_enter();
    h.assert_contains("Saved");
    h.send_text(":q");
    h.send_enter();
    assert_eq!(h.wait_for_exit(Duration::from_secs(3)), Some(0));
    assert!(path.exists(), ".xlsx file should exist after :w");
    // Reopen via cli arg.
    let mut h2 = Harness::with_args(&[path_str]);
    h2.assert_contains("foo");
    h2.assert_contains("bar");
    quit_force(&mut h2);
}

#[test]
fn workflow_csv_export() {
    let dir = tempfile::tempdir().unwrap();
    let path = dir.path().join("out.csv");
    let path_str = path.to_str().unwrap();
    let mut h = Harness::new();
    h.send_text("ihi"); h.send_enter();
    h.send_text("ibye"); h.send_enter();
    // Use the palette command since Ctrl+E may collide with terminal byte conventions.
    h.send_text(":export ");
    h.send_text(path_str);
    h.send_enter();
    // Allow time for the file write.
    std::thread::sleep(Duration::from_millis(300));
    quit_force(&mut h);
    if path.exists() {
        let body = std::fs::read_to_string(&path).unwrap();
        assert!(body.contains("hi") && body.contains("bye"),
            "exported CSV should contain values, got: {}", body);
    } else {
        // :export isn't a palette command — but Ctrl+E is the documented binding.
        panic!("CSV export via :export <path> did not produce a file; check that the command is wired.");
    }
}

// ----- Charts -----

#[test]
fn workflow_chart_bar_opens_popup() {
    let mut h = Harness::new();
    h.send_text("i1"); h.send_enter();
    h.send_text("i2"); h.send_enter();
    h.send_text("i3"); h.send_enter();
    h.send_text(":chart bar A1:A3");
    h.send_enter();
    // Chart popup should be visible with min/max indicators.
    h.assert_contains("min=");
    h.assert_contains("max=");
    h.send_esc();
    // Popup should close.
    h.assert_absent("min=");
    quit_force(&mut h);
}

// ----- Sheet workflows -----

#[test]
fn workflow_multi_sheet_creation_and_switch() {
    let mut h = Harness::new();
    h.assert_contains("Sheet1");
    h.send_text(":sheet new"); h.send_enter();
    h.assert_contains("Sheet2");
    h.send_text(":sheet new"); h.send_enter();
    h.assert_contains("Sheet3");
    // Type a value on Sheet3
    h.send_text("ithree"); h.send_enter();
    h.assert_contains("three");
    // Switch back to Sheet1 — should NOT show "three"
    h.send_text(":sheet prev"); h.send_enter();
    h.send_text(":sheet prev"); h.send_enter();
    h.assert_absent("three");
    quit_force(&mut h);
}

// ----- Auto-fill -----

#[test]
fn workflow_ctrl_d_autofills_formula() {
    let mut h = Harness::new();
    // A1 = 10, A2 = 20, A3 = 30
    h.send_text("i10"); h.send_enter();
    h.send_text("i20"); h.send_enter();
    h.send_text("i30"); h.send_enter();
    // Move to B1, enter =A1*2, then select B1:B3 with shift+down arrows
    h.send_text("kkk"); // back to A1
    h.send_arrow(common::Arrow::Right); // B1
    // Split `i` from `=A1*2` so the Insert-mode transition has time to settle.
    h.send_text("i");
    h.send_text("=A1*2");
    h.send_enter(); // commit B1 = 20
    h.send_text("k"); // back to B1
    h.send_text("vjj"); // visual cell select B1:B3
    h.send_ctrl('d');
    // After autofill, B1=20, B2=40, B3=60 — at least 40 and 60 should be visible.
    h.assert_contains("40");
    h.assert_contains("60");
    quit_force(&mut h);
}

// ----- Formula error display -----

#[test]
fn workflow_div_by_zero_displays_error() {
    let mut h = Harness::new();
    h.send_text("i=1/0"); h.send_enter();
    h.send_text("k");
    h.assert_contains("#DIV/0!");
    quit_force(&mut h);
}

#[test]
fn workflow_unknown_function_displays_name_error() {
    let mut h = Harness::new();
    h.send_text("i=NOPE(1)"); h.send_enter();
    h.send_text("k");
    h.assert_contains("#NAME?");
    quit_force(&mut h);
}

// ----- Long-form value editing flows -----

#[test]
fn workflow_overwrite_existing_cell_via_S() {
    let mut h = Harness::new();
    h.send_text("ifoo"); h.send_enter();
    h.send_text("k");
    h.assert_contains("foo");
    // S = substitute row: clear current row and enter Insert at col 0.
    h.send_text("S");
    h.send_text("baz");
    h.send_enter();
    h.assert_contains("baz");
    h.assert_absent("foo");
    quit_force(&mut h);
}

// ----- Help search -----

#[test]
fn workflow_help_search_finds_term() {
    let mut h = Harness::new();
    h.send_text("?");
    h.assert_contains("-- HELP --");
    h.send_text("/");
    h.send_text("XLOOKUP");
    // The help popup should scroll to a line mentioning XLOOKUP.
    h.assert_contains("XLOOKUP");
    h.send_enter(); // commit search
    h.send_esc();
    quit_force(&mut h);
}

// ----- Format / style -----

#[test]
fn workflow_bold_via_palette_then_ctrl_b() {
    let mut h = Harness::new();
    h.send_text("ibold-me"); h.send_enter();
    h.send_text("k");
    h.send_ctrl('b'); // toggle bold
    // Bold is hard to verify visually through vt100 contents alone; we just
    // confirm no crash and value persists.
    h.assert_contains("bold-me");
    quit_force(&mut h);
}

// ----- Frozen panes -----

#[test]
fn workflow_freeze_then_scroll_keeps_header() {
    let mut h = Harness::new();
    h.send_text("iHEADER"); h.send_enter();
    h.send_text("k");
    // Move to A2, freeze 1 row.
    h.send_arrow(common::Arrow::Down);
    h.send_text(":freeze");
    h.send_enter();
    h.assert_contains("Frozen");
    // Move way down — the header should still be visible.
    h.send_text("50j");
    h.assert_contains("HEADER");
    quit_force(&mut h);
}

// ----- Visual block delete column -----

#[test]
fn workflow_visual_block_delete() {
    let mut h = Harness::new();
    h.send_text("ix"); h.send_enter();
    h.send_text("iy"); h.send_enter();
    h.send_text("iz"); h.send_enter();
    h.send_text("kkk"); // back to A1
    h.send_ctrl('v'); // visual block
    h.send_text("jj");
    h.send_text("d");
    h.assert_absent("x");
    h.assert_absent("y");
    h.assert_absent("z");
    quit_force(&mut h);
}
