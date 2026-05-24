//! Detailed Insert-mode cursor behavior driven through a real PTY.

// Test names use uppercase letters to mirror vim keys (e.g. `insert_A_*`).
#![allow(non_snake_case)]

mod common;

use common::{Arrow, Harness};
use std::time::Duration;

fn quit_force(h: &mut Harness) {
    h.send_text(":q!");
    h.send_enter();
    assert_eq!(h.wait_for_exit(Duration::from_secs(3)), Some(0));
}

// ----- Cursor keys mid-edit -----

#[test]
fn insert_left_arrow_moves_cursor() {
    let mut h = Harness::new();
    h.send_text("i");
    h.send_text("abcde");
    h.send_arrow(Arrow::Left);
    h.send_arrow(Arrow::Left);
    // Now insert 'X' between c and d.
    h.send_text("X");
    h.send_enter();
    h.send_text("k");
    let bar = h.formula_bar();
    assert!(bar.contains("abcXde"),
        "Left-arrow should reposition the cursor mid-edit. Got: {:?}", bar);
    quit_force(&mut h);
}

#[test]
fn insert_home_jumps_to_start() {
    let mut h = Harness::new();
    h.send_text("i");
    h.send_text("world");
    h.send(b"\x1b[H"); // Home
    h.send_text("hello ");
    h.send_enter();
    h.send_text("k");
    let bar = h.formula_bar();
    assert!(bar.contains("hello world"),
        "Home should move cursor to start. Got: {:?}", bar);
    quit_force(&mut h);
}

#[test]
fn insert_end_jumps_to_end() {
    let mut h = Harness::new();
    h.send_text("i");
    h.send_text("hello");
    h.send_arrow(Arrow::Left);
    h.send_arrow(Arrow::Left);
    h.send_arrow(Arrow::Left);
    h.send(b"\x1b[F"); // End
    h.send_text("!");
    h.send_enter();
    h.send_text("k");
    let bar = h.formula_bar();
    assert!(bar.contains("hello!"),
        "End should move cursor past last char. Got: {:?}", bar);
    quit_force(&mut h);
}

#[test]
fn insert_backspace_deletes_left_of_cursor() {
    let mut h = Harness::new();
    h.send_text("i");
    h.send_text("hello");
    h.send_backspace();
    h.send_enter();
    h.send_text("k");
    let bar = h.formula_bar();
    assert!(bar.contains("hell"),
        "Backspace should remove last char. Got: {:?}", bar);
    quit_force(&mut h);
}

#[test]
fn insert_delete_key_removes_right_of_cursor() {
    let mut h = Harness::new();
    h.send_text("i");
    h.send_text("hello");
    h.send_arrow(Arrow::Left);
    h.send_arrow(Arrow::Left);
    // Now between 'l' and 'l'. Delete should remove the right 'l'.
    h.send(b"\x1b[3~"); // Delete key
    h.send_enter();
    h.send_text("k");
    let bar = h.formula_bar();
    assert!(bar.contains("helo"),
        "Delete should remove char to the right. Got: {:?}", bar);
    quit_force(&mut h);
}

// ----- Multi-byte input -----

#[test]
fn insert_unicode_chars_preserved() {
    let mut h = Harness::new();
    h.send_text("i");
    h.send_text("café"); // 'é' is multi-byte UTF-8
    h.send_enter();
    h.send_text("k");
    let bar = h.formula_bar();
    assert!(bar.contains("café"), "Unicode should round-trip. Got: {:?}", bar);
    quit_force(&mut h);
}

// ----- Editing in formula -----

#[test]
fn insert_edit_existing_formula() {
    let mut h = Harness::new();
    h.send_text("i");
    h.send_text("=1+1");
    h.send_enter();
    h.send_text("k");
    // Re-edit. With `A` cursor is at end.
    h.send_text("A");
    // Type "+1" to extend.
    h.send_text("+1");
    h.send_enter();
    h.send_text("k");
    // Result should be 3.
    h.assert_contains("3");
    quit_force(&mut h);
}

// ----- Continued editing across multiple cells via Tab/Enter -----

#[test]
fn insert_chain_tab_then_enter() {
    let mut h = Harness::new();
    h.send_text("ia"); h.send_tab();   // A1=a, cursor B1
    h.send_text("ib"); h.send_tab();   // B1=b, cursor C1
    h.send_text("ic"); h.send_enter(); // C1=c, cursor C2
    // GoToCell A1 to verify.
    h.send_ctrl('g'); h.send_text("A1"); h.send_enter();
    h.assert_contains("a");
    h.assert_contains("b");
    h.assert_contains("c");
    quit_force(&mut h);
}

// ----- Backspace at position 0 doesn't underflow -----

#[test]
fn insert_backspace_at_position_zero_safe() {
    let mut h = Harness::new();
    h.send_text("i");
    h.send_backspace();
    h.send_backspace();
    h.send_backspace();
    h.send_enter();
    quit_force(&mut h);
}

// ----- Typing into a cell with existing content via `i` (cursor at start) -----

#[test]
fn insert_i_prepends_to_existing_content() {
    let mut h = Harness::new();
    h.send_text("iworld"); h.send_enter();
    h.send_text("k");
    h.send_text("i");        // i = cursor at start
    h.send_text("hello ");
    h.send_enter();
    h.send_text("k");
    let bar = h.formula_bar();
    assert!(bar.contains("hello world"),
        "`i` should put cursor at start, typing prepends. Got: {:?}", bar);
    quit_force(&mut h);
}

// ----- `A` appends to existing content -----

#[test]
fn insert_A_appends_to_existing_content() {
    let mut h = Harness::new();
    h.send_text("ihello"); h.send_enter();
    h.send_text("k");
    h.send_text("A");        // A = cursor at end
    h.send_text(" world");
    h.send_enter();
    h.send_text("k");
    let bar = h.formula_bar();
    assert!(bar.contains("hello world"),
        "`A` should put cursor at end, typing appends. Got: {:?}", bar);
    quit_force(&mut h);
}

// ----- Esc while editing leaves prior cell content intact -----

#[test]
fn insert_esc_leaves_prior_value_intact() {
    let mut h = Harness::new();
    h.send_text("ikeep"); h.send_enter();
    h.send_text("k");
    h.send_text("A");
    h.send_text("discard-this");
    h.send_esc();
    h.send_text("k"); // motion to refresh formula bar
    let bar = h.formula_bar();
    assert!(bar.contains("keep") && !bar.contains("discard"),
        "Esc must restore original value. Got: {:?}", bar);
    quit_force(&mut h);
}
