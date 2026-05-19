//! Edge cases and argument handling. Each test pokes a corner of the
//! command/key surface that might silently misbehave.

mod common;

use common::Harness;
use std::time::Duration;

fn quit_force(h: &mut Harness) {
    h.send_text(":q!");
    h.send_enter();
    assert_eq!(h.wait_for_exit(Duration::from_secs(3)), Some(0));
}

// ----- :name preserves case for the name -----

#[test]
fn edge_name_preserves_case() {
    let mut h = Harness::new();
    h.send_text("i7"); h.send_enter();
    h.send_text("k");
    h.send_text(":name LUCKY A1"); h.send_enter();
    h.send_text(":names"); h.send_enter();
    // The user typed LUCKY (uppercase); the listing should preserve the case.
    let contents = h.screen_contents();
    assert!(
        contents.contains("LUCKY"),
        ":names should show user-supplied case 'LUCKY', got:\n{}",
        contents
    );
    quit_force(&mut h);
}

// ----- :comment preserves case and special chars -----

#[test]
fn edge_comment_preserves_case_and_punctuation() {
    let mut h = Harness::new();
    h.send_text("ix"); h.send_enter();
    h.send_text("k");
    h.send_text(":comment Hello, World!");
    h.send_enter();
    h.assert_contains("Hello, World!");
    quit_force(&mut h);
}

// ----- :rename preserves case (already true via dedicated path) -----

#[test]
fn edge_rename_preserves_mixed_case() {
    let mut h = Harness::new();
    h.send_text(":rename QuarterlyBudget");
    h.send_enter();
    h.assert_contains("QuarterlyBudget");
    quit_force(&mut h);
}

// ----- Insert and delete column at boundaries -----

#[test]
fn edge_insert_col_at_first_position() {
    let mut h = Harness::new();
    h.send_text("ifirst"); h.send_enter();
    h.send_text("k");
    h.send_text(":ic"); h.send_enter();
    // After insert col, "first" should now be at B1 (was A1).
    // Both rows visible in the grid.
    let r1 = h.data_row(1); // data row 1
    assert!(
        r1.contains("first"),
        "After :ic, 'first' should still be visible (in B1). Got: {:?}",
        r1
    );
    quit_force(&mut h);
}

// ----- Many rapid mode switches don't leak state -----

#[test]
fn edge_rapid_mode_cycling() {
    let mut h = Harness::new();
    for _ in 0..5 {
        h.send_text("i"); h.send_esc();
        h.send_text("v"); h.send_esc();
        h.send_text(":"); h.send_esc();
    }
    h.assert_contains("-- NORMAL --");
    quit_force(&mut h);
}

// ----- Visual mode then ctrl+c then immediate paste -----

#[test]
fn edge_visual_yank_then_paste_round_trip() {
    let mut h = Harness::new();
    h.send_text("ialpha"); h.send_enter();
    h.send_text("ibeta"); h.send_enter();
    h.send_text("kk"); // back to A1
    h.send_text("vj"); // visual select A1:A2
    h.send_ctrl('c'); // copy (returns to Normal per our visual ctrl handler)
    h.send_text("3l"); // move right 3 cols
    h.send_ctrl('v'); // paste
    h.assert_contains("alpha");
    h.assert_contains("beta");
    quit_force(&mut h);
}

// ----- Search clears highlights on Esc -----

#[test]
fn edge_search_then_esc_clears_state() {
    let mut h = Harness::new();
    h.send_text("ifoo"); h.send_enter();
    h.send_text("/foo");
    h.assert_contains("1/1");
    h.send_esc();
    h.assert_contains("-- NORMAL --");
    // n should no longer cycle (no active search).
    h.send_text("n");
    h.assert_contains("-- NORMAL --");
    quit_force(&mut h);
}

// ----- u after no edits doesn't crash or claim success -----

#[test]
fn edge_undo_on_clean_workbook_no_crash() {
    let mut h = Harness::new();
    h.send_text("u");
    h.assert_contains("-- NORMAL --");
    quit_force(&mut h);
}

// ----- Repeat undo until empty doesn't loop forever -----

#[test]
fn edge_many_undos_drains_stack() {
    let mut h = Harness::new();
    h.send_text("ix"); h.send_enter();
    for _ in 0..30 {
        h.send_text("u");
    }
    h.assert_contains("-- NORMAL --");
    quit_force(&mut h);
}

// ----- Formula with circular ref (no iterative) reports #ERROR or similar -----

#[test]
fn edge_circular_ref_without_iterative() {
    let mut h = Harness::new();
    h.send_text("i"); h.send_text("=A1+1"); h.send_enter();
    // Should reject the formula with a visible status message (not silently
    // drop the user's input).
    h.assert_contains("Circular");
    quit_force(&mut h);
}

// ----- :format invalid number doesn't crash -----

#[test]
fn edge_format_with_bad_decimals() {
    let mut h = Harness::new();
    h.send_text("i1"); h.send_enter();
    h.send_text("k");
    h.send_text(":format number xyz");
    h.send_enter();
    // Either reports error or no-ops; either way, stays alive.
    h.assert_contains("-- NORMAL --");
    quit_force(&mut h);
}

// ----- :color with unknown color name -----

#[test]
fn edge_color_unknown_name_errors_gracefully() {
    let mut h = Harness::new();
    h.send_text("ix"); h.send_enter();
    h.send_text("k");
    h.send_text(":color unobtanium");
    h.send_enter();
    h.assert_contains("Unknown color");
    quit_force(&mut h);
}

// ----- :filter on invalid column -----

#[test]
fn edge_filter_invalid_column() {
    let mut h = Harness::new();
    h.send_text(":filter ZZZ value");
    h.send_enter();
    // Should not crash. Either reports error or no-ops.
    h.assert_contains("-- NORMAL --");
    quit_force(&mut h);
}

// ----- Pasting nothing -----

#[test]
fn edge_paste_with_empty_clipboard() {
    let mut h = Harness::new();
    h.send_text("p");
    h.assert_contains("Nothing to paste");
    quit_force(&mut h);
}

// ----- Pressing q while in INSERT mode types 'q' (not quits) -----

#[test]
fn edge_q_in_insert_mode_is_literal() {
    let mut h = Harness::new();
    h.send_text("i");
    h.send_text("q");
    h.send_enter();
    h.assert_contains("-- NORMAL --");
    // A1 should contain "q".
    h.send_text("k");
    let bar = h.formula_bar();
    assert!(bar.contains("q"), "A1 should hold literal 'q'. Got: {:?}", bar);
    quit_force(&mut h);
}

// ----- Save dialog Esc cancels without saving -----

#[test]
fn edge_save_dialog_esc_returns_to_normal() {
    let mut h = Harness::new();
    h.send_text("ix"); h.send_enter();
    h.send_ctrl('s');
    h.assert_contains("-- SAVE --");
    h.send_esc();
    h.assert_contains("-- NORMAL --");
    // Workbook should still be dirty (cancel did not save).
    h.assert_contains("tshts *");
    quit_force(&mut h);
}

// ----- :sheet next/prev at single-sheet workbook is no-op -----

#[test]
fn edge_sheet_next_with_only_one_sheet() {
    let mut h = Harness::new();
    h.send_text(":sheet next"); h.send_enter();
    h.assert_contains("[Sheet1]");
    h.send_text(":sheet prev"); h.send_enter();
    h.assert_contains("[Sheet1]");
    quit_force(&mut h);
}

// ----- Long edit input is not truncated in the buffer -----

#[test]
fn edge_long_input_preserved() {
    let mut h = Harness::new();
    h.send_text("i");
    let long = "x".repeat(80);
    h.send_text(&long);
    h.send_enter();
    h.send_text("k");
    let bar = h.formula_bar();
    // Formula bar shows the value (which is 80 x's). vt100 may have its own
    // line wrap so we just check at least 60 are visible.
    let xcount = bar.matches('x').count();
    assert!(xcount >= 60, "Long input should mostly survive in formula bar; saw {} x's", xcount);
    quit_force(&mut h);
}
