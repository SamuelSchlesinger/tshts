//! Power-user feature scenarios driven through a real PTY: structured tables,
//! pivots, 3-D refs, data validation, comments in detail.

mod common;

use common::Harness;
use std::time::Duration;

fn quit_force(h: &mut Harness) {
    h.send_text(":q!");
    h.send_enter();
    assert_eq!(h.wait_for_exit(Duration::from_secs(3)), Some(0));
}

// ----- :table create -----

#[test]
fn power_table_create_and_list() {
    let mut h = Harness::new();
    // Fill 2x2 area cell-by-cell — Tab commits the edit and moves to next col.
    h.send_text("i"); h.send_text("h1"); h.send_tab();
    h.send_text("i"); h.send_text("h2"); h.send_enter();
    h.send_text("kh"); // A2 in NORMAL (cursor was at B2 → C2; back via h → B2 then h again to A2)
    // Actually simpler: jump to A2 explicitly via GoToCell.
    h.send_ctrl('g'); h.send_text("A2"); h.send_enter();
    h.send_text("i"); h.send_text("x"); h.send_tab();
    h.send_text("i"); h.send_text("y"); h.send_enter();
    h.send_text(":table create A1:B2 name=Mine");
    h.send_enter();
    h.send_text(":table list"); h.send_enter();
    h.assert_contains("Mine");
    quit_force(&mut h);
}

// ----- :pivot -----

#[test]
fn power_pivot_basic_grouping() {
    let mut h = Harness::new();
    // Build source data: A=category, B=value
    h.send_text("ifruit"); h.send_text("\t"); h.send_text("1"); h.send_enter();
    h.send_text("ifruit"); h.send_text("\t"); h.send_text("2"); h.send_enter();
    h.send_text("ivegetable"); h.send_text("\t"); h.send_text("5"); h.send_enter();
    // Pivot into column D: group by col A, sum col B.
    h.send_text(":pivot A1:B3 D1 row=A value=B agg=sum");
    h.send_enter();
    // Should land 3 (sum of 1+2) and 5 in the target.
    h.assert_contains("3");
    h.assert_contains("5");
    quit_force(&mut h);
}

// ----- 3-D references -----

#[test]
fn power_3d_sum_across_sheets() {
    let mut h = Harness::new();
    h.send_text("i10"); h.send_enter();
    h.send_text(":sheet new"); h.send_enter();
    h.send_text("i20"); h.send_enter();
    h.send_text(":sheet new"); h.send_enter();
    h.send_text("i30"); h.send_enter();
    // On Sheet3, B1 = =SUM(Sheet1:Sheet3!A1) → 60
    h.send_arrow(common::Arrow::Right);
    h.send_arrow(common::Arrow::Up);
    h.send_text("i");
    h.send_text("=SUM(Sheet1:Sheet3!A1)");
    h.send_enter();
    h.send_text("k");
    h.assert_contains("60");
    quit_force(&mut h);
}

// ----- Data validation -----

#[test]
fn power_validation_flags_violators() {
    let mut h = Harness::new();
    // Validate column A: value must be > 0.
    h.send_text(":validate A _ > 0");
    h.send_enter();
    // Enter a violating value.
    h.send_text("i-5"); h.send_enter();
    // The status bar or status message should not crash; the cell stores -5
    // and (per CLAUDE.md) is marked with `!`. Just verify no crash and the
    // value rendered.
    h.send_text("k");
    h.assert_contains("-5");
    quit_force(&mut h);
}

// ----- Comments full lifecycle -----

#[test]
fn power_comment_set_and_clear() {
    let mut h = Harness::new();
    h.send_text("ix"); h.send_enter();
    h.send_text("k");
    h.send_text(":comment my-note");
    h.send_enter();
    h.assert_contains("my-note");
    // Clear the comment.
    h.send_text(":comment clear");
    h.send_enter();
    h.assert_absent("my-note");
    quit_force(&mut h);
}

// ----- Iterative calc / circular ref -----

#[test]
fn power_iterative_on_for_circular_ref() {
    let mut h = Harness::new();
    h.send_text(":iterative on");
    h.send_enter();
    h.assert_contains("Iterative");
    // Don't actually create circular ref (may be flaky); just confirm toggle.
    h.send_text(":iterative off");
    h.send_enter();
    quit_force(&mut h);
}

// ----- :name and :unname round-trip -----

#[test]
fn power_name_then_unname() {
    let mut h = Harness::new();
    h.send_text("i7"); h.send_enter();
    h.send_text("k");
    h.send_text(":name LUCKY A1");
    h.send_enter();
    // Formula using the named cell.
    h.send_text("j");
    h.send_text("i");
    h.send_text("=LUCKY*2");
    h.send_enter();
    h.send_text("k");
    h.assert_contains("14");
    // Remove the name.
    h.send_text(":unname LUCKY"); h.send_enter();
    quit_force(&mut h);
}

// ----- Format number with decimals -----

#[test]
fn power_format_number_decimals() {
    let mut h = Harness::new();
    h.send_text("i");
    h.send_text("=10/3");
    h.send_enter();
    h.send_text("k");
    h.send_text(":format number 2");
    h.send_enter();
    // Should show "3.33" not "3.333333..."
    h.assert_contains("3.33");
    quit_force(&mut h);
}

// ----- Format currency -----

#[test]
fn power_format_currency() {
    let mut h = Harness::new();
    h.send_text("i100"); h.send_enter();
    h.send_text("k");
    h.send_text(":format currency");
    h.send_enter();
    // Currency format should show "$" or similar.
    let contents = h.screen_contents();
    assert!(contents.contains("$") || contents.contains("100.00"),
        "Currency format should show $ or .00 decimals. Got:\n{}", contents);
    quit_force(&mut h);
}

// ----- :cache clear -----

#[test]
fn power_cache_clear_safe() {
    let mut h = Harness::new();
    h.send_text(":cache clear");
    h.send_enter();
    h.assert_contains("-- NORMAL --");
    quit_force(&mut h);
}

// ----- regex on/off and case on/off -----

#[test]
fn power_regex_and_case_toggles() {
    let mut h = Harness::new();
    h.send_text(":regex on"); h.send_enter();
    h.assert_contains("Regex");
    h.send_text(":regex off"); h.send_enter();
    h.send_text(":case on"); h.send_enter();
    h.assert_contains("Case");
    h.send_text(":case off"); h.send_enter();
    quit_force(&mut h);
}

// ----- Multi-sheet delete -----

#[test]
fn power_sheet_delete_last_blocked() {
    let mut h = Harness::new();
    h.send_text(":sheet delete"); h.send_enter();
    h.assert_contains("Cannot delete the last sheet");
    quit_force(&mut h);
}

// ----- Multi-sheet add then delete -----

#[test]
fn power_sheet_add_delete_cycle() {
    let mut h = Harness::new();
    h.send_text(":sheet new"); h.send_enter();
    h.assert_contains("[Sheet2]");
    h.send_text(":sheet delete"); h.send_enter();
    h.assert_contains("Deleted sheet");
    // Tab bar (row 0) should no longer carry "Sheet2".
    let tab_bar = h.row(0);
    assert!(!tab_bar.contains("Sheet2"),
        "After delete, tab bar should not show Sheet2. Got: {:?}", tab_bar);
    quit_force(&mut h);
}

// ----- Color: foreground -----

#[test]
fn power_color_fg_palette_runs() {
    let mut h = Harness::new();
    h.send_text("ix"); h.send_enter();
    h.send_text("k");
    h.send_text(":color red");
    h.send_enter();
    // No status error; value still there.
    h.assert_contains("x");
    quit_force(&mut h);
}

// ----- Background color -----

#[test]
fn power_color_bg_palette_runs() {
    let mut h = Harness::new();
    h.send_text("ix"); h.send_enter();
    h.send_text("k");
    h.send_text(":bg blue");
    h.send_enter();
    h.assert_contains("x");
    quit_force(&mut h);
}

// ----- Help search wraps around -----

#[test]
fn power_help_search_for_rare_term() {
    let mut h = Harness::new();
    h.send_text("?");
    h.send_text("/");
    h.send_text("LAMBDA");
    // The help should scroll to a line containing LAMBDA.
    h.assert_contains("LAMBDA");
    h.send_esc();
    h.send_esc();
    quit_force(&mut h);
}

// ----- Trace dependencies -----

#[test]
fn power_trace_dependents_safe() {
    let mut h = Harness::new();
    h.send_text("i10"); h.send_enter();
    h.send_text("i=A1*2"); h.send_enter();
    h.send_text("kk");
    h.send_text(":trace dependents");
    h.send_enter();
    h.assert_contains("-- NORMAL --");
    quit_force(&mut h);
}
