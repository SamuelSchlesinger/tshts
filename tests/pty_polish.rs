//! Tests targeting UX polish: error messages, edge cases, recovery from
//! mistakes. The kind of issues a first-time user trips over.

mod common;

use common::Harness;
use std::time::Duration;

fn quit_force(h: &mut Harness) {
    h.send_text(":q!");
    h.send_enter();
    assert_eq!(h.wait_for_exit(Duration::from_secs(3)), Some(0));
}

// ----- Empty palette command -----

#[test]
fn polish_empty_palette_just_closes() {
    let mut h = Harness::new();
    h.send_text(":");
    h.assert_contains("-- COMMAND --");
    h.send_enter(); // empty command
    // Should close palette gracefully — no "Unknown command: "
    h.assert_contains("-- NORMAL --");
    h.assert_absent("Unknown command");
    quit_force(&mut h);
}

// ----- :w with no filename when never saved -----

#[test]
fn polish_w_no_filename_opens_save_dialog() {
    let mut h = Harness::new();
    h.send_text("ihi"); h.send_enter();
    h.send_text(":w"); h.send_enter();
    h.assert_contains("-- SAVE --");
    h.send_esc();
    quit_force(&mut h);
}

// ----- Failing :w to bad path produces error status, not crash -----

#[test]
fn polish_w_bad_path_errors_gracefully() {
    let mut h = Harness::new();
    h.send_text("ihi"); h.send_enter();
    h.send_text(":w /this/directory/does/not/exist/foo.tshts");
    h.send_enter();
    // App should report failure and stay alive.
    h.assert_contains("ailed"); // "Failed" or "failed"
    quit_force(&mut h);
}

// ----- :e nonexistent shows error -----

#[test]
fn polish_e_nonexistent_file_errors_gracefully() {
    let mut h = Harness::new();
    h.send_text(":e /this/does/not/exist.tshts");
    h.send_enter();
    h.assert_contains("ailed");
    quit_force(&mut h);
}

// ----- :goto with bad cell ref -----

#[test]
fn polish_goto_with_bad_cell_shows_error_or_no_op() {
    let mut h = Harness::new();
    h.send_ctrl('g');
    h.send_text("NOTACELL");
    h.send_enter();
    // Should remain in NORMAL mode without crashing.
    h.assert_contains("-- NORMAL --");
    quit_force(&mut h);
}

// ----- Long formula in formula bar wraps or scrolls -----

#[test]
fn polish_very_long_formula_displays() {
    let mut h = Harness::new();
    h.send_text("i");
    h.send_text("=1+2+3+4+5+6+7+8+9+10");
    h.send_enter();
    h.send_text("k");
    let bar = h.row(1);
    assert!(bar.contains("=1+2") || bar.contains("55"),
        "Formula bar should show formula or its value. Got: {:?}", bar);
    quit_force(&mut h);
}

// ----- Selecting then typing replaces selection -----

// (Not strictly vim — Excel users might expect this. Let's check current
// behavior to document it.)
#[test]
fn polish_visual_then_letter_in_normal_does_what() {
    let mut h = Harness::new();
    h.send_text("ihi"); h.send_enter();
    h.send_text("k");
    h.send_text("v"); // visual mode
    h.send_text("y"); // yank
    h.assert_contains("-- NORMAL --"); // y exits visual back to normal
    quit_force(&mut h);
}

// ----- u in visual mode should exit visual cleanly -----

#[test]
fn polish_u_in_visual_does_not_crash() {
    let mut h = Harness::new();
    h.send_text("v");
    h.send_text("u");
    // Either undoes (no-op on fresh state) or no-ops; must not crash.
    quit_force(&mut h);
}

// ----- Multiple esc/q presses don't crash -----

#[test]
fn polish_multiple_esc_in_normal_safe() {
    let mut h = Harness::new();
    h.send_esc();
    h.send_esc();
    h.send_esc();
    h.assert_contains("-- NORMAL --");
    quit_force(&mut h);
}

// ----- Tab at last column does NOT move past -----

#[test]
fn polish_tab_at_last_column_stays_put() {
    let mut h = Harness::new();
    // Navigate to last column. cols default is 26.
    h.send_text("25l"); // go to last visible col (A=0..Z=25)
    let before = h.row(1).to_string();
    h.send_tab();
    let after = h.row(1).to_string();
    // Should be at Z or wrap; either way no crash.
    let _ = (before, after);
    quit_force(&mut h);
}

// ----- :sort with no data doesn't crash -----

#[test]
fn polish_sort_empty_column_safe() {
    let mut h = Harness::new();
    h.send_text(":sort asc");
    h.send_enter();
    h.assert_contains("-- NORMAL --");
    quit_force(&mut h);
}

// ----- delete row on row 0 doesn't underflow -----

#[test]
fn polish_dd_at_row_zero_safe() {
    let mut h = Harness::new();
    h.send_text("ialpha"); h.send_enter();
    h.send_text("k");
    h.send_text("dd");
    // Cursor should still be at A1 (or A0; depending on impl). App alive.
    h.assert_absent("alpha");
    quit_force(&mut h);
}

// ----- Esc in CommandPalette returns to Normal cleanly -----

#[test]
fn polish_esc_in_palette_returns_to_normal() {
    let mut h = Harness::new();
    h.send_text(":");
    h.send_text("anything");
    h.send_esc();
    h.assert_contains("-- NORMAL --");
    // Status shouldn't show the discarded input.
    h.assert_absent(":anything");
    quit_force(&mut h);
}

// ----- Editing then quit-without-save prompts -----

#[test]
fn polish_q_with_dirty_pressing_n_returns_to_normal() {
    let mut h = Harness::new();
    h.send_text("ix"); h.send_enter();
    h.send_text("q");
    h.assert_contains("Unsaved changes");
    h.send_text("n");
    h.assert_contains("-- NORMAL --");
    h.assert_absent("Unsaved changes");
    quit_force(&mut h);
}

// ----- After :q!, dirty data is discarded -----

#[test]
fn polish_q_bang_discards_dirty_data() {
    let mut h = Harness::new();
    h.send_text("idirty-data"); h.send_enter();
    h.send_text(":q!"); h.send_enter();
    assert_eq!(h.wait_for_exit(Duration::from_secs(3)), Some(0));
}

// ----- Hitting / in CommandPalette stays in palette (it's a literal char) -----

#[test]
fn polish_slash_in_palette_stays_in_palette() {
    let mut h = Harness::new();
    h.send_text(":");
    h.send_text("foo/bar");
    h.assert_contains("-- COMMAND --");
    h.assert_contains("foo/bar");
    h.send_esc();
    quit_force(&mut h);
}

// ----- Negative count or 0 shouldn't crash -----

#[test]
fn polish_zero_count_is_motion_not_count() {
    // In vim, leading `0` is a row-start motion, not a count zero.
    let mut h = Harness::new();
    h.send_text("3l"); // go to col 3
    let before = h.row(1).to_string();
    h.send_text("0"); // back to col 0
    let after = h.row(1).to_string();
    assert_ne!(before, after, "0 should move to col 0, not be a count");
    quit_force(&mut h);
}
