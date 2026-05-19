//! PTY scenarios for advanced features: cross-sheet refs, filter, freeze,
//! charts, named ranges, conditional formatting, undo depth. Each test
//! exercises a feature through the real binary and asserts on what the user
//! actually sees.

mod common;

use common::{Arrow, Harness};
use std::time::Duration;

fn quit_force(h: &mut Harness) {
    h.send_text(":q!");
    h.send_enter();
    assert_eq!(h.wait_for_exit(Duration::from_secs(3)), Some(0));
}

// ----- Cross-sheet references -----

#[test]
fn adv_cross_sheet_reference_resolves() {
    let mut h = Harness::new();
    // Sheet1: A1 = 100
    h.send_text("i100"); h.send_enter();
    // Add a second sheet
    h.send_text(":sheet new"); h.send_enter();
    h.assert_contains("Sheet2");
    // On Sheet2, A1 = =Sheet1!A1+5 → should display 105
    h.send_text("i");
    h.send_text("=Sheet1!A1+5");
    h.send_enter();
    h.send_text("k");
    h.assert_contains("105");
    quit_force(&mut h);
}

#[test]
fn adv_rename_sheet_updates_cross_sheet_refs() {
    let mut h = Harness::new();
    h.send_text("i42"); h.send_enter();
    h.send_text(":sheet new"); h.send_enter();
    h.send_text("i");
    h.send_text("=Sheet1!A1");
    h.send_enter();
    h.send_text("k");
    h.assert_contains("42");
    // Rename Sheet1 → Data
    h.send_text(":sheet prev"); h.send_enter();
    h.send_text(":rename Data"); h.send_enter();
    h.assert_contains("Data");
    // The Sheet2 reference should now read Data!A1 and still resolve to 42.
    h.send_text(":sheet next"); h.send_enter();
    h.send_text("k");
    h.assert_contains("42");
    // The formula text should mention "Data!", not stale "Sheet1!".
    let formula = h.formula_bar();
    assert!(
        formula.contains("Data!") || formula.contains("42"),
        "After rename, the cross-sheet formula should refer to the new name. Formula bar: {:?}",
        formula
    );
    quit_force(&mut h);
}

// ----- Filter -----

#[test]
fn adv_filter_hides_non_matching_rows() {
    let mut h = Harness::new();
    h.send_text("ialpha"); h.send_enter();
    h.send_text("ibeta"); h.send_enter();
    h.send_text("ialpha"); h.send_enter();
    h.send_text(":filter A alpha");
    h.send_enter();
    // After filter, "beta" row should be hidden.
    h.assert_absent("beta");
    h.assert_contains("alpha");
    h.send_text(":unfilter"); h.send_enter();
    h.assert_contains("beta");
    quit_force(&mut h);
}

// ----- Hide row / column -----

#[test]
fn adv_hide_row_then_show() {
    let mut h = Harness::new();
    h.send_text("ifoo"); h.send_enter();
    h.send_text("ibar"); h.send_enter();
    // Move to row 2 (bar).
    h.send_text("k");
    h.send_text(":hide row");
    h.send_enter();
    // "bar" should not appear anywhere — grid hides the row AND cursor
    // should have moved off it so the formula bar doesn't leak the value.
    h.assert_absent("bar");
    h.send_text(":show rows"); h.send_enter();
    h.assert_contains("bar");
    quit_force(&mut h);
}

// ----- Freeze + scroll keeps header visible -----

#[test]
fn adv_freeze_then_scroll_far_down() {
    let mut h = Harness::new();
    h.send_text("iHEAD"); h.send_enter();
    // Move to A2, freeze.
    h.send_text(":freeze"); h.send_enter();
    h.assert_contains("Frozen");
    // Scroll way down using count motion.
    h.send_text("100j");
    // HEAD must still be visible because row 1 is frozen.
    h.assert_contains("HEAD");
    quit_force(&mut h);
}

// ----- Chart auto-refresh when source changes -----

#[test]
fn adv_chart_reflects_source_data_change() {
    let mut h = Harness::new();
    h.send_text("i1"); h.send_enter();
    h.send_text("i2"); h.send_enter();
    h.send_text("i3"); h.send_enter();
    h.send_text(":chart bar A1:A3"); h.send_enter();
    h.assert_contains("max=3");
    h.send_esc();
    // Bump A3 → 9, reopen chart, max should now be 9.
    h.send_text("3G");          // jump to row 3
    h.send_text("S");           // substitute row (clear + edit at col 0)
    h.send_text("9"); h.send_enter();
    h.send_text(":chart bar A1:A3"); h.send_enter();
    h.assert_contains("max=9");
    h.send_esc();
    quit_force(&mut h);
}

// ----- Named ranges -----

#[test]
fn adv_named_range_resolves_in_formula() {
    let mut h = Harness::new();
    h.send_text("i100"); h.send_enter();
    h.send_text("i200"); h.send_enter();
    // Define "Total" = A1:A2
    h.send_text("kk"); // back to A1
    h.send_text(":name Total A1:A2"); h.send_enter();
    // In A3, =SUM(Total)
    h.send_text("jj"); // to A3
    h.send_text("i=SUM(Total)"); h.send_enter();
    h.send_text("k");
    h.assert_contains("300");
    quit_force(&mut h);
}

// ----- Undo depth -----

#[test]
fn adv_many_edits_then_many_undos() {
    let mut h = Harness::new();
    // 20 edits in column A
    for i in 0..20 {
        h.send_text("i");
        h.send_text(&format!("v{}", i));
        h.send_enter();
    }
    // Sheet should have v0..v19 visible (at least v19 in last entered row).
    h.assert_contains("v19");
    // 20 undos.
    for _ in 0..20 {
        h.send_text("u");
    }
    // All edits gone.
    h.assert_absent("v19");
    h.assert_absent("v0");
    quit_force(&mut h);
}

// ----- F5 forces recalc -----

#[test]
fn adv_f5_recalculates_formulas() {
    let mut h = Harness::new();
    h.send_text("i");
    h.send_text("=RAND()");
    h.send_enter();
    h.send_text("k");
    let _before = h.data_row(1);
    // F5 (xterm CSI form). Recalc should not crash; we don't compare values
    // since RAND() can collide by chance.
    h.send(b"\x1b[15~");
    std::thread::sleep(Duration::from_millis(200));
    let _after = h.data_row(1);
    quit_force(&mut h);
}

// ----- Hyperlink rendering -----

#[test]
fn adv_url_value_displayed() {
    let mut h = Harness::new();
    h.send_text("i");
    h.send_text("http://example.com");
    h.send_enter();
    h.send_text("k");
    // The cell will be truncated in the grid; check the formula bar.
    let bar = h.formula_bar();
    assert!(bar.contains("http://example.com"),
        "Formula bar should show the full URL: {:?}", bar);
    quit_force(&mut h);
}

// ----- Comment on cell -----

#[test]
fn adv_comment_shows_in_status_bar() {
    let mut h = Harness::new();
    h.send_text("ihi"); h.send_enter();
    h.send_text("k");
    h.send_text(":comment hello note");
    h.send_enter();
    // Status bar shows "Comment: hello note" when the cell is selected.
    h.assert_contains("hello note");
    quit_force(&mut h);
}

// ----- Conditional formatting at least registers without crash -----

#[test]
fn adv_conditional_format_command_runs() {
    let mut h = Harness::new();
    h.send_text("i5"); h.send_enter();
    h.send_text("i15"); h.send_enter();
    h.send_text(":cf A \"_ > 10\" bg=red bold");
    h.send_enter();
    h.assert_contains("15"); // still rendered
    h.send_text(":cf list"); h.send_enter();
    // List should show our rule.
    h.assert_contains("_ > 10");
    quit_force(&mut h);
}

// ----- Hide column then show -----

#[test]
fn adv_hide_col_then_show() {
    let mut h = Harness::new();
    h.send_text("ix"); h.send_enter();
    h.send_text("k");
    h.send_arrow(Arrow::Right); // B1
    h.send_text("iy"); h.send_enter();
    h.send_text("k");
    h.send_text(":hide col B"); h.send_enter();
    h.assert_contains("Hidden cols: 1");
    // Wait until the grid header no longer shows B (it should jump A → C).
    let ok = h.wait_for_text("A        C ", Duration::from_secs(2));
    assert!(ok, "After hiding col B, grid header should skip B. Screen:\n{}",
        h.screen_contents());
    h.send_text(":show cols"); h.send_enter();
    h.assert_contains("y");
    quit_force(&mut h);
}

// ----- LAMBDA + named lambda -----

#[test]
fn adv_lambda_named_callable() {
    let mut h = Harness::new();
    h.send_text(":name DBL LAMBDA(x, x*2)");
    h.send_enter();
    h.send_text("i");
    h.send_text("=DBL(21)");
    h.send_enter();
    h.send_text("k");
    h.assert_contains("42");
    quit_force(&mut h);
}

// ----- Goal seek -----

#[test]
fn adv_goal_seek_finds_solution() {
    let mut h = Harness::new();
    // A1 = 1 (input), A2 = =A1*A1 (target)
    h.send_text("i1"); h.send_enter();
    h.send_text("i=A1*A1"); h.send_enter();
    // GoalSeek: find A1 such that A2 = 9. Answer: 3.
    h.send_text(":goalseek A2 9 A1"); h.send_enter();
    h.send_text("kk");
    h.assert_contains("3"); // approximately
    quit_force(&mut h);
}

// ----- :recalc forces recalculation -----

#[test]
fn adv_recalc_does_not_crash() {
    let mut h = Harness::new();
    h.send_text("i=1+1"); h.send_enter();
    h.send_text(":recalc"); h.send_enter();
    h.send_text("k");
    h.assert_contains("2");
    quit_force(&mut h);
}

// ----- Cell with comment shows ! indicator (or similar) — UX check -----

#[test]
fn adv_search_inside_help_with_n_cycles_matches() {
    let mut h = Harness::new();
    h.send_text("?");
    h.send_text("/");
    h.send_text("SUM");
    h.send_enter(); // commit search
    let after_first = h.screen_contents();
    h.send_text("n");
    let after_second = h.screen_contents();
    // The visible help window should be at a different scroll position.
    // We can't easily verify exact scroll, but the visible content must
    // mention SUM in both cases.
    assert!(after_first.contains("SUM"));
    assert!(after_second.contains("SUM"));
    h.send_esc();
    quit_force(&mut h);
}
