//! `/` search must keep the matched cell on screen: while the query is being
//! typed, when it's finalized with Enter, and when `n`/`N` navigates to it —
//! even if the user scrolled or moved away in between. Far jumps center the
//! match and Esc restores the original view, as in Vim/Neovim.
//!
//! Regression: the grid reported `height - 1` visible rows, but it's a
//! bordered Table so only `height - 3` render. Hits in the last two
//! "visible" rows were clipped and never scrolled into view. Horizontally,
//! scrolling by a stale column count left hits in wide-column regions past
//! the right edge.

mod common;

use common::Harness;
use std::time::{Duration, Instant};

// Harness layout (30 rows): row 4 = column letters, rows 5..=25 = the 21 data
// rows that actually render, row 26 = grid bottom border.
const FIRST_DATA_ROW: u16 = 5;
const LAST_DATA_ROW: u16 = 25;
/// The 11th of the 21 data rows — where a centered cursor row renders.
const MIDDLE_DATA_ROW: u16 = 15;

fn quit_force(h: &mut Harness) {
    h.send_text(":q!");
    h.send_enter();
    assert_eq!(h.wait_for_exit(Duration::from_secs(3)), Some(0));
}

fn goto(h: &mut Harness, cell: &str) {
    h.send_ctrl('g');
    h.send_text(cell);
    h.send_enter();
}

/// Is `needle` rendered inside the grid body? (The header row and formula
/// bar also echo the cell ref / value, so they're excluded.)
fn grid_shows(h: &Harness, needle: &str) -> bool {
    (FIRST_DATA_ROW..=LAST_DATA_ROW).any(|r| h.row(r).contains(needle))
}

fn dump_grid(h: &Harness) -> String {
    (0..30).map(|r| format!("{:2}|{}\n", r, h.row(r))).collect()
}

/// Poll until the grid body shows (or stops showing) `needle`.
fn wait_grid(h: &Harness, needle: &str, want: bool) -> bool {
    let deadline = Instant::now() + Duration::from_secs(3);
    while Instant::now() < deadline {
        if grid_shows(h, needle) == want {
            return true;
        }
        std::thread::sleep(Duration::from_millis(30));
    }
    false
}

fn assert_on_screen(h: &Harness, needle: &str, when: &str) {
    assert!(
        wait_grid(h, needle, true),
        "{when}: match '{needle}' should be on screen\n{}",
        dump_grid(h)
    );
}

fn assert_off_screen(h: &Harness, needle: &str, when: &str) {
    assert!(
        wait_grid(h, needle, false),
        "{when}: setup expects '{needle}' to be off screen\n{}",
        dump_grid(h)
    );
}

#[test]
fn search_hit_below_bottom_edge_is_scrolled_into_view() {
    let mut h = Harness::new();
    // A23 sits two rows past the 21 rows that actually render — exactly the
    // band the old off-by-two viewport height thought was visible.
    goto(&mut h, "A23");
    h.send_text("izqx");
    h.send_enter();
    goto(&mut h, "A1");
    assert_off_screen(&h, "zqx", "back at A1");

    // While typing the query.
    h.send_text("/zq");
    assert_on_screen(&h, "zqx", "typing the query");

    // After finalizing.
    h.send_text("x");
    h.send_enter();
    h.assert_contains("-- NORMAL --");
    assert_on_screen(&h, "zqx", "after Enter");

    // Move away with the keyboard, then `n` must bring it back.
    h.send_text("gg");
    assert_off_screen(&h, "zqx", "after gg");
    h.send_text("n");
    assert_on_screen(&h, "zqx", "after n");

    // Scroll the view away with the mouse wheel (cursor stays on the
    // match); `N` must bring it back too.
    for _ in 0..25 {
        h.send(b"\x1b[<65;10;10M"); // SGR wheel-down
    }
    assert_off_screen(&h, "zqx", "after wheel scroll");
    h.send_text("N");
    assert_on_screen(&h, "zqx", "after N");

    quit_force(&mut h);
}

#[test]
fn search_hit_in_wide_columns_is_scrolled_into_view() {
    let mut h = Harness::new();
    // Widen T..W to 30 cells each. From A (12 narrow columns on screen) a
    // jump to W used to scroll by that stale count of 12, leaving W clipped
    // past the right edge.
    for col in ["T1", "U1", "V1", "W1"] {
        goto(&mut h, col);
        h.send_text(&"_".repeat(22));
    }
    goto(&mut h, "W3");
    h.send_text("izqw");
    h.send_enter();
    goto(&mut h, "A1");
    assert_off_screen(&h, "zqw", "back at A1");

    h.send_text("/zqw");
    assert_on_screen(&h, "zqw", "typing the query");
    h.send_enter();
    assert_on_screen(&h, "zqw", "after Enter");

    h.send_text("0");
    assert_off_screen(&h, "zqw", "after 0");
    h.send_text("n");
    assert_on_screen(&h, "zqw", "after n");

    quit_force(&mut h);
}

/// Poll until screen row `row` starts with the grid border + `label`
/// (the row-number column).
fn wait_row_label(h: &Harness, row: u16, label: &str) -> bool {
    let want = format!("\u{2502}{label} ");
    let deadline = Instant::now() + Duration::from_secs(3);
    while Instant::now() < deadline {
        if h.row(row).starts_with(&want) {
            return true;
        }
        std::thread::sleep(Duration::from_millis(30));
    }
    false
}

#[test]
fn search_far_hit_is_centered() {
    let mut h = Harness::new();
    goto(&mut h, "A80");
    h.send_text("izqc");
    h.send_enter();
    goto(&mut h, "A1");
    assert_off_screen(&h, "zqc", "back at A1");

    // A far jump centers the match (Vim's update_topline), rather than
    // parking it on the bottom edge.
    h.send_text("/zqc");
    assert!(
        wait_row_label(&h, MIDDLE_DATA_ROW, "80"),
        "typing the query: row 80 should be centered\n{}",
        dump_grid(&h)
    );
    h.send_enter();
    h.send_text("gg");
    assert_off_screen(&h, "zqc", "after gg");
    h.send_text("n");
    assert!(
        wait_row_label(&h, MIDDLE_DATA_ROW, "80"),
        "after n: row 80 should be centered\n{}",
        dump_grid(&h)
    );
    quit_force(&mut h);
}

#[test]
fn search_esc_restores_cursor_and_view() {
    let mut h = Harness::new();
    goto(&mut h, "A80");
    h.send_text("izqe");
    h.send_enter();
    // Far jump to A50 centers it: rows 40..=60 on screen.
    goto(&mut h, "A50");
    assert!(
        wait_row_label(&h, FIRST_DATA_ROW, "40"),
        "setup: expected rows 40.. on screen\n{}",
        dump_grid(&h)
    );

    h.send_text("/zqe");
    assert_on_screen(&h, "zqe", "typing the query");
    assert!(!h.row(FIRST_DATA_ROW).starts_with("\u{2502}40 "), "view moved");

    h.send_esc();
    h.assert_contains("-- NORMAL --");
    assert!(
        wait_row_label(&h, FIRST_DATA_ROW, "40"),
        "Esc should restore the original view\n{}",
        dump_grid(&h)
    );
    assert!(h.formula_bar().contains("A50"), "Esc should restore the cursor: {}", h.formula_bar());
    quit_force(&mut h);
}
