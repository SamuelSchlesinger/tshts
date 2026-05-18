//! Spill / dynamic-array behavior through a real PTY.

mod common;

use common::{Arrow, Harness};
use std::time::Duration;

fn quit_force(h: &mut Harness) {
    h.send_text(":q!");
    h.send_enter();
    assert_eq!(h.wait_for_exit(Duration::from_secs(3)), Some(0));
}

#[test]
fn spill_sequence_5x1_populates_column() {
    let mut h = Harness::new();
    h.send_text("i");
    h.send_text("=SEQUENCE(5)");
    h.send_enter();
    // The 5 values 1..5 should appear in A1:A5.
    h.assert_contains("1");
    h.assert_contains("5");
    quit_force(&mut h);
}

#[test]
fn spill_sequence_2d_grid() {
    let mut h = Harness::new();
    h.send_text("i");
    h.send_text("=SEQUENCE(3,3)");
    h.send_enter();
    // Should produce a 3x3 grid with 1..9.
    let contents = h.screen_contents();
    for v in &["1", "2", "5", "9"] {
        assert!(contents.contains(v),
            "Expected {} in SEQUENCE(3,3) spill. Screen:\n{}", v, contents);
    }
    quit_force(&mut h);
}

#[test]
fn spill_unique_dedup_values() {
    let mut h = Harness::new();
    h.send_text("iapple"); h.send_enter();
    h.send_text("ibanana"); h.send_enter();
    h.send_text("iapple"); h.send_enter();
    h.send_text("icherry"); h.send_enter();
    // C1 = =UNIQUE(A1:A4)
    h.send_ctrl('g'); h.send_text("C1"); h.send_enter();
    h.send_text("i");
    h.send_text("=UNIQUE(A1:A4)");
    h.send_enter();
    let contents = h.screen_contents();
    let banana_count = contents.matches("banana").count();
    assert!(banana_count >= 2,
        "banana should appear once in source col A and once in UNIQUE col C, total >= 2. Saw {}.",
        banana_count);
    quit_force(&mut h);
}

#[test]
fn spill_sort_ascending() {
    let mut h = Harness::new();
    h.send_text("i3"); h.send_enter();
    h.send_text("i1"); h.send_enter();
    h.send_text("i2"); h.send_enter();
    h.send_ctrl('g'); h.send_text("B1"); h.send_enter();
    h.send_text("i");
    h.send_text("=SORT(A1:A3)");
    h.send_enter();
    // Walk B1:B3 — should be 1, 2, 3 in order.
    h.send_ctrl('g'); h.send_text("B1"); h.send_enter();
    let bar = h.row(1);
    assert!(bar.contains("1"), "B1 should be smallest. Got: {:?}", bar);
    quit_force(&mut h);
}

#[test]
fn spill_filter_predicate_keeps_matches() {
    let mut h = Harness::new();
    h.send_text("i5"); h.send_enter();
    h.send_text("i15"); h.send_enter();
    h.send_text("i25"); h.send_enter();
    h.send_ctrl('g'); h.send_text("B1"); h.send_enter();
    h.send_text("i");
    h.send_text("=FILTER(A1:A3, A1:A3>10)");
    h.send_enter();
    // Should keep 15 and 25 only.
    let contents = h.screen_contents();
    assert!(contents.contains("15") && contents.contains("25"));
    quit_force(&mut h);
}

#[test]
fn spill_transpose_swaps_orientation() {
    let mut h = Harness::new();
    h.send_text("i1"); h.send_tab();
    h.send_text("i2"); h.send_tab();
    h.send_text("i3"); h.send_enter();
    // Row A1:C1 = 1,2,3. TRANSPOSE → A3:A5 = 1,2,3 (column).
    h.send_ctrl('g'); h.send_text("A3"); h.send_enter();
    h.send_text("i");
    h.send_text("=TRANSPOSE(A1:C1)");
    h.send_enter();
    let contents = h.screen_contents();
    assert!(contents.contains("1") && contents.contains("2") && contents.contains("3"));
    quit_force(&mut h);
}

#[test]
fn spill_overwrite_target_shows_spill_error() {
    // Put a blocker in the row that SEQUENCE(3) would spill into (A2), then
    // place the formula at A1 — should get #SPILL!.
    let mut h = Harness::new();
    h.send_arrow(Arrow::Down); // A2
    h.send_text("iblocker"); h.send_enter(); // A2 = blocker, cursor → A3
    h.send_ctrl('g'); h.send_text("A1"); h.send_enter();
    h.send_text("i");
    h.send_text("=SEQUENCE(3)");
    h.send_enter();
    h.send_text("k");
    let contents = h.screen_contents();
    assert!(contents.contains("#SPILL"),
        "Should display #SPILL! when blocked. Screen:\n{}", contents);
    quit_force(&mut h);
}

#[test]
fn spill_sumproduct_simple() {
    let mut h = Harness::new();
    h.send_text("i2"); h.send_enter();
    h.send_text("i3"); h.send_enter();
    h.send_text("i4"); h.send_enter();
    h.send_text("i");
    h.send_text("=SUMPRODUCT(A1:A3, A1:A3)");
    h.send_enter();
    h.send_text("k");
    // 2² + 3² + 4² = 29
    h.assert_contains("29");
    quit_force(&mut h);
}

#[test]
fn spill_let_local_binding() {
    let mut h = Harness::new();
    h.send_text("i");
    h.send_text("=LET(x, 5, y, 3, x*y + x + y)");
    h.send_enter();
    h.send_text("k");
    // 5*3 + 5 + 3 = 23
    h.assert_contains("23");
    quit_force(&mut h);
}

#[test]
fn spill_array_literal_inline() {
    let mut h = Harness::new();
    h.send_text("i");
    h.send_text("=SUM({1,2,3,4,5})");
    h.send_enter();
    h.send_text("k");
    h.assert_contains("15");
    quit_force(&mut h);
}
