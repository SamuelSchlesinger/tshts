//! UX-focused PTY scenarios. Each test asks: "does the user actually get the
//! right feedback?" not just "does the underlying state change?"

mod common;

use common::{Arrow, Harness};
use std::time::Duration;

fn quit_force(h: &mut Harness) {
    h.send_text(":q!");
    h.send_enter();
    assert_eq!(h.wait_for_exit(Duration::from_secs(3)), Some(0));
}

// ----- Dirty-flag indicator -----

#[test]
fn ux_dirty_flag_appears_after_edit() {
    let mut h = Harness::new();
    // Before edit, no dirty indicator.
    assert!(
        !h.screen_contents().contains("tshts *"),
        "clean workbook should not show dirty marker"
    );
    h.send_text("i");
    h.send_text("x");
    h.send_enter();
    h.assert_contains("tshts *"); // dirty marker
    quit_force(&mut h);
}

#[test]
fn ux_dirty_flag_clears_after_save() {
    let dir = tempfile::tempdir().unwrap();
    let path = dir.path().join("clean.tshts");
    let mut h = Harness::new();
    h.send_text("ihello");
    h.send_enter();
    h.assert_contains("tshts *");
    h.send_text(":w ");
    h.send_text(path.to_str().unwrap());
    h.send_enter();
    h.assert_contains("Saved");
    // After save, dirty indicator should be gone.
    h.assert_absent("tshts *");
    quit_force(&mut h);
}

// ----- Status messages: appear, then clear at the right moment -----

#[test]
fn ux_yank_shows_status_message() {
    let mut h = Harness::new();
    h.send_text("ix");
    h.send_enter();
    h.send_text("k");
    h.send_text("yy");
    h.assert_contains("Yanked");
    quit_force(&mut h);
}

#[test]
fn ux_delete_shows_status_message() {
    let mut h = Harness::new();
    h.send_text("ix");
    h.send_enter();
    h.send_text("k");
    h.send_text("dd");
    h.assert_contains("Deleted");
    quit_force(&mut h);
}

// ----- Edit lifecycle -----

#[test]
fn ux_esc_cancels_edit_leaves_cell_unchanged() {
    let mut h = Harness::new();
    h.send_text("ifirst");
    h.send_enter(); // commit "first"
    h.send_text("k"); // back to A1
    h.assert_contains("first");
    h.send_text("isecond"); // start editing again (vim `i` puts cursor at start)
    h.send_esc(); // cancel
    // A1 should STILL contain "first", not "second" or anything garbled.
    let row1 = h.data_row(1); // row 4 in the rendered grid is data row 1
    assert!(
        row1.contains("first"),
        "Esc must restore previous cell value, got row: {:?}",
        row1
    );
    // Workbook should still be dirty from the first commit, but the cancelled
    // edit must NOT have added a new dirty mark — actually it stays dirty
    // because we never saved. We just check the cell content.
    quit_force(&mut h);
}

#[test]
fn ux_enter_commits_then_moves_cursor_down() {
    let mut h = Harness::new();
    h.send_text("i");
    h.send_text("hi");
    h.send_enter();
    // Cursor should now be at A2.
    let formula_bar = h.formula_bar();
    assert!(
        formula_bar.contains("A2"),
        "After commit, cursor should advance to A2. Formula bar: {:?}",
        formula_bar
    );
    quit_force(&mut h);
}

#[test]
fn ux_tab_in_editing_commits_then_moves_right() {
    let mut h = Harness::new();
    h.send_text("ihi");
    h.send_tab();
    let formula_bar = h.formula_bar();
    assert!(
        formula_bar.contains("B1"),
        "After Tab commit, cursor should advance to B1. Formula bar: {:?}",
        formula_bar
    );
    quit_force(&mut h);
}

// ----- Search -----

#[test]
fn ux_search_finds_value_and_navigates() {
    let mut h = Harness::new();
    // Plant a few values.
    h.send_text("ialpha"); h.send_enter();
    h.send_text("ibeta"); h.send_enter();
    h.send_text("ialpha"); h.send_enter();
    // Search for alpha.
    h.send_text("/alpha");
    h.assert_contains("-- SEARCH --");
    h.assert_contains("1/2 results"); // first of two matches
    h.send_arrow(Arrow::Down); // next match within search
    h.assert_contains("2/2");
    h.send_enter(); // close search
    quit_force(&mut h);
}

#[test]
fn ux_search_no_results_shows_message() {
    let mut h = Harness::new();
    h.send_text("ifoo"); h.send_enter();
    h.send_text("/bar");
    h.assert_contains("no results");
    h.send_esc();
    quit_force(&mut h);
}

// ----- GoToCell -----

#[test]
fn ux_ctrl_g_then_b5_lands_at_b5() {
    let mut h = Harness::new();
    h.send_ctrl('g');
    h.assert_contains("-- GOTO --");
    h.send_text("B5");
    h.send_enter();
    let formula_bar = h.formula_bar();
    assert!(
        formula_bar.contains("B5"),
        "After GoToCell B5, formula bar should show B5. Got: {:?}",
        formula_bar
    );
    quit_force(&mut h);
}

// ----- Command palette -----

#[test]
fn ux_palette_unknown_command_shows_error_status() {
    let mut h = Harness::new();
    h.send_text(":");
    h.assert_contains("-- COMMAND --");
    h.send_text("zzzzz");
    h.send_enter();
    h.assert_contains("Unknown command");
    quit_force(&mut h);
}

#[test]
fn ux_palette_insert_row_command() {
    let mut h = Harness::new();
    h.send_text("ihere"); h.send_enter();
    h.send_text("k"); // back to A1
    h.send_text(":ir");
    h.send_enter();
    // After inserting a row, "here" should now be at row 2.
    // Row 4 in the rendered grid is data row 1 (header is row 3).
    let r1 = h.data_row(1);
    let r2 = h.data_row(2);
    assert!(
        !r1.contains("here"),
        "After :ir, row 1 should be the new blank row. Got: {:?}",
        r1
    );
    assert!(
        r2.contains("here"),
        "After :ir, 'here' should be at row 2. Got: {:?}",
        r2
    );
    quit_force(&mut h);
}

// ----- Sheet operations -----

#[test]
fn ux_sheet_new_adds_tab() {
    let mut h = Harness::new();
    // Default has Sheet1.
    h.assert_contains("[Sheet1]");
    h.send_text(":sheet new");
    h.send_enter();
    h.assert_contains("Sheet2");
    quit_force(&mut h);
}

#[test]
fn ux_sheet_rename() {
    let mut h = Harness::new();
    h.send_text(":rename Budget");
    h.send_enter();
    h.assert_contains("Budget");
    h.assert_absent("[Sheet1]");
    quit_force(&mut h);
}

// ----- Undo/redo -----

#[test]
fn ux_undo_after_dd_restores_row() {
    let mut h = Harness::new();
    h.send_text("ihi"); h.send_enter();
    h.send_text("k"); // back to A1
    h.assert_contains("hi");
    h.send_text("dd");
    h.assert_absent("hi");
    h.send_text("u");
    h.assert_contains("hi");
    quit_force(&mut h);
}

#[test]
fn ux_redo_after_undo() {
    let mut h = Harness::new();
    h.send_text("ihi"); h.send_enter();
    h.send_text("k");
    h.send_text("dd");
    h.send_text("u");
    h.assert_contains("hi");
    h.send_ctrl('r'); // vim redo
    h.assert_absent("hi");
    quit_force(&mut h);
}

// ----- Help -----

#[test]
fn ux_help_jump_to_section_zero_for_vim_mode() {
    let mut h = Harness::new();
    h.send_text("?");
    h.assert_contains("-- HELP --");
    h.send_text("0"); // jump to section 0 = Vim Mode
    h.assert_contains("VIM MODE");
    h.send_esc();
    quit_force(&mut h);
}

// ----- Formula bar shows current cell content -----

#[test]
fn ux_formula_bar_shows_selected_cell_value() {
    let mut h = Harness::new();
    h.send_text("ihello"); h.send_enter();
    h.send_text("k"); // back to A1
    let formula_bar = h.formula_bar();
    assert!(
        formula_bar.contains("hello"),
        "Formula bar should show selected cell value. Got: {:?}",
        formula_bar
    );
    quit_force(&mut h);
}

#[test]
fn ux_formula_bar_shows_full_text_when_grid_truncates() {
    let mut h = Harness::new();
    // 20-char value will be truncated in the grid (col width ~8).
    h.send_text("ithis-is-a-long-value-here");
    h.send_enter();
    h.send_text("k");
    let formula_bar = h.formula_bar();
    assert!(
        formula_bar.contains("this-is-a-long-value-here"),
        "Formula bar should show full value even when grid clips it. Got: {:?}",
        formula_bar
    );
    quit_force(&mut h);
}

// ----- Pending op preview -----

#[test]
fn ux_pending_op_shows_in_status_bar() {
    let mut h = Harness::new();
    h.send_text("5");
    h.assert_contains("[5]");
    h.send_text("d");
    h.assert_contains("[5d]");
    h.send_esc(); // cancel
    h.assert_absent("[5d]");
    quit_force(&mut h);
}

// ----- Visual mode stats -----

#[test]
fn ux_visual_mode_shows_stats() {
    let mut h = Harness::new();
    h.send_text("i1"); h.send_enter();
    h.send_text("i2"); h.send_enter();
    h.send_text("i3"); h.send_enter();
    h.send_text("kkk"); // back to A1
    h.send_text("vjj"); // select A1:A3
    // SUM=6, AVG=2, COUNT=3
    h.assert_contains("SUM=6");
    h.send_esc();
    quit_force(&mut h);
}
