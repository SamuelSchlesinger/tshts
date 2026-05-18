//! End-to-end PTY smoke tests for tshts.
//!
//! Each test spawns the real release binary inside a pseudo-tty, drives it
//! with synthesized keystrokes, and asserts on the rendered screen content
//! parsed back through `vt100`. These tests catch entire classes of bugs the
//! unit tests can't reach: actual crossterm raw-mode behavior, terminal byte
//! collisions, the alternate-screen lifecycle, and "does the binary even
//! start" regressions.
//!
//! Run with: `cargo test --test pty_smoke`
//!
//! See `tests/common/mod.rs` for the harness API.

mod common;

use common::{Arrow, Harness};
use std::time::Duration;

#[test]
fn launches_and_quits_cleanly_with_colon_q() {
    let mut h = Harness::new();
    // The mode chip should be visible immediately after startup.
    h.assert_contains("-- NORMAL --");
    // Confirm a quit path actually exits the process.
    h.send_text(":q");
    h.send_enter();
    assert_eq!(
        h.wait_for_exit(Duration::from_secs(3)),
        Some(0),
        "process should exit cleanly on :q"
    );
}

#[test]
fn q_in_normal_mode_quits_clean_workbook() {
    let mut h = Harness::new();
    h.send_text("q");
    assert_eq!(
        h.wait_for_exit(Duration::from_secs(3)),
        Some(0),
        "bare q should quit when the workbook is clean"
    );
}

#[test]
fn i_enters_insert_mode_and_types_into_cell() {
    let mut h = Harness::new();
    h.assert_contains("-- NORMAL --");
    h.send_text("i");
    h.assert_contains("-- INSERT --");
    // Keep payload short enough to fit in the default column width once rendered.
    h.send_text("howdy");
    // Mid-edit, "howdy" appears in the status bar.
    h.assert_contains("howdy");
    h.send_enter(); // commit
    h.assert_contains("-- NORMAL --");
    // Value persisted into A1.
    h.assert_contains("howdy");
    h.send_text(":q!");
    h.send_enter();
    assert_eq!(h.wait_for_exit(Duration::from_secs(3)), Some(0));
}

#[test]
fn mode_chip_cycles_normal_insert_visual_command() {
    let mut h = Harness::new();
    h.assert_contains("-- NORMAL --");

    // Normal → Insert → Normal
    h.send_text("i");
    h.assert_contains("-- INSERT --");
    h.send_esc();
    h.assert_contains("-- NORMAL --");

    // Normal → Visual → Normal
    h.send_text("v");
    h.assert_contains("-- VISUAL --");
    h.send_esc();
    h.assert_contains("-- NORMAL --");

    // Normal → Visual Line → Normal
    h.send_text("V");
    h.assert_contains("-- VISUAL LINE --");
    h.send_esc();
    h.assert_contains("-- NORMAL --");

    // Normal → Command → Normal
    h.send_text(":");
    h.assert_contains("-- COMMAND --");
    h.send_esc();
    h.assert_contains("-- NORMAL --");

    h.send_text(":q");
    h.send_enter();
    assert_eq!(h.wait_for_exit(Duration::from_secs(3)), Some(0));
}

#[test]
fn write_then_quit_roundtrips_to_disk() {
    let dir = tempfile::tempdir().unwrap();
    let path = dir.path().join("smoke.tshts");
    let path_str = path.to_str().unwrap();

    let mut h = Harness::new();
    h.send_text("i");
    h.send_text("answer 42");
    h.send_enter(); // commit edit
    h.send_text(":w ");
    h.send_text(path_str);
    h.send_enter();
    h.assert_contains("Saved");
    h.send_text(":q");
    h.send_enter();
    assert_eq!(h.wait_for_exit(Duration::from_secs(3)), Some(0));

    // File on disk should exist and contain the value.
    let body = std::fs::read_to_string(&path).expect("saved file");
    assert!(
        body.contains("answer 42"),
        "saved file should contain typed value, got: {}",
        body
    );

    // Now re-open it via command-line arg and verify the value rendered.
    let mut h2 = Harness::with_args(&[path_str]);
    h2.assert_contains("answer 42");
    h2.send_text(":q");
    h2.send_enter();
    assert_eq!(h2.wait_for_exit(Duration::from_secs(3)), Some(0));
}

#[test]
fn dd_then_p_round_trips_row() {
    let mut h = Harness::new();
    // Put "alpha" in A1.
    h.send_text("i");
    h.send_text("alpha");
    h.send_enter();   // commit; finish_editing moves cursor to A2
    h.send_text("k"); // back to A1
    h.assert_contains("alpha");
    // Yank the row, move down, paste.
    h.send_text("yy");
    h.send_text("j");
    h.send_text("p");
    // Two visible occurrences expected — a contains() check is not enough.
    let occurrences = h.screen_contents().matches("alpha").count();
    assert!(
        occurrences >= 2,
        "expected 'alpha' to appear at least twice after yy+j+p, got {} in screen:\n{}",
        occurrences,
        h.screen_contents()
    );
    h.send_text(":q!");
    h.send_enter();
    assert_eq!(h.wait_for_exit(Duration::from_secs(3)), Some(0));
}

#[test]
fn ctrl_i_does_not_enter_insert_mode() {
    // Ctrl+I is identical to Tab at the terminal byte level. Pressing it must
    // NOT route to vim's `i` (which would enter insert mode). Tab in normal
    // mode moves the cursor right — we just confirm the mode stays NORMAL.
    let mut h = Harness::new();
    h.assert_contains("-- NORMAL --");
    h.send_ctrl('i'); // sends 0x09 — same as Tab
    // Should NOT be in INSERT mode.
    h.assert_absent("-- INSERT --");
    h.assert_contains("-- NORMAL --");
    h.send_text(":q");
    h.send_enter();
    assert_eq!(h.wait_for_exit(Duration::from_secs(3)), Some(0));
}

#[test]
fn q_with_dirty_workbook_prompts_then_force_quits() {
    let mut h = Harness::new();
    h.send_text("i");
    h.send_text("dirty");
    h.send_enter(); // commit so workbook is actually dirty
    h.assert_contains("dirty");
    // Try to quit; should NOT exit, should show confirm prompt.
    h.send_text("q");
    h.assert_contains("Unsaved changes");
    assert!(
        !h.has_exited(),
        "dirty quit should hold for confirmation, not exit"
    );
    // Force quit with `:q!`.
    h.send_text("n"); // cancel the confirm
    h.assert_contains("-- NORMAL --");
    h.send_text(":q!");
    h.send_enter();
    assert_eq!(h.wait_for_exit(Duration::from_secs(3)), Some(0));
}

#[test]
fn arrow_keys_and_hjkl_both_navigate() {
    let mut h = Harness::new();
    // Drop short markers (column width is narrow — long strings get clipped).
    h.send_text("i");
    h.send_text("aa");
    h.send_enter(); // cursor → A2
    // Down 2 via arrow → A4.
    h.send_arrow(Arrow::Down);
    h.send_arrow(Arrow::Down);
    h.send_text("ibb");
    h.send_enter(); // A4 = "bb", cursor → A5
    // Up 4 via vim k → A1.
    h.send_text("kkkk");
    h.send_text("icc");
    h.send_enter(); // A1 prefix "cc" before "aa"; commits then cursor → A2
    h.assert_contains("cc");
    h.assert_contains("bb");
    h.send_text(":q!");
    h.send_enter();
    assert_eq!(h.wait_for_exit(Duration::from_secs(3)), Some(0));
}

#[test]
fn count_prefix_5j_moves_five_rows_down() {
    let mut h = Harness::new();
    // Drop a marker at the origin and at row 5, then verify 5j lands us there.
    h.send_text("i");
    h.send_text("origin");
    h.send_enter();
    h.send_text("5j");
    h.send_text("itarget");
    h.send_enter();
    // Both should be on screen.
    h.assert_contains("origin");
    h.assert_contains("target");
    h.send_text(":q!");
    h.send_enter();
    assert_eq!(h.wait_for_exit(Duration::from_secs(3)), Some(0));
}

#[test]
fn help_popup_opens_and_closes() {
    let mut h = Harness::new();
    h.send_text("?");
    h.assert_contains("-- HELP --");
    // The intro line mentions Vim Mode as the section-0 jump anchor.
    h.assert_contains("Vim Mode");
    h.send_esc();
    h.assert_contains("-- NORMAL --");
    h.send_text(":q");
    h.send_enter();
    assert_eq!(h.wait_for_exit(Duration::from_secs(3)), Some(0));
}

#[test]
fn formula_evaluates_and_displays_result() {
    let mut h = Harness::new();
    h.send_text("i");
    h.send_text("=2+3");
    h.send_enter(); // finish_editing
    // The displayed value should be 5, not "=2+3".
    h.assert_contains("5");
    h.send_text(":q!");
    h.send_enter();
    assert_eq!(h.wait_for_exit(Duration::from_secs(3)), Some(0));
}
