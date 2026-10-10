#![allow(
    clippy::unwrap_used,
    clippy::expect_used,
    reason = "test assertions panic on failure by design"
)]

//! Focused invariants for the contextual slide selector.

use super::Selector;

#[test]
fn selector_conversions_keep_name_and_order_selectors_distinct() {
    assert_eq!(Selector::from("Title"), Selector::Name("Title".into()));
    assert_eq!(
        Selector::from(3_usize),
        Selector::Position(litchi_core::Position::new(3))
    );
    assert_ne!(Selector::from("3"), Selector::from(3_usize));
}

#[test]
fn selector_accepts_owned_names_and_checked_positions() {
    let owned = Selector::from(String::from("Summary"));
    assert!(matches!(owned, Selector::Name(std::borrow::Cow::Owned(_))));
    let name = String::from("Summary");
    let borrowed = Selector::from(&name);
    assert!(matches!(
        borrowed,
        Selector::Name(std::borrow::Cow::Borrowed(_))
    ));
    assert_eq!(borrowed, owned);
    let position = litchi_core::Position::new(0);
    assert_eq!(Selector::from(position), Selector::Position(position));
}
