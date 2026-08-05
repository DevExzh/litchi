//! Focused invariants for the contextual slide selector.

use super::Key;

#[test]
fn key_conversions_keep_name_and_order_selectors_distinct() {
    assert_eq!(Key::from("Title"), Key::Name("Title"));
    assert_eq!(Key::from(3_usize), Key::Index(3));
    assert_ne!(Key::from("3"), Key::from(3_usize));
}

#[test]
fn key_is_copyable_for_checked_graph_queries() {
    let key = Key::Name("Summary");
    let copied = key;
    assert_eq!(copied, Key::Name("Summary"));
    let index = Key::Index(0);
    assert_eq!(index, Key::Index(0));
}
