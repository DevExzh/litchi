//! Shared archive-free interactive table-lock semantics.
//!
//! Numbers intentionally uses the canonical iWork table lock state so
//! Keynote, Numbers, and Pages expose one compatible semantic vocabulary.
//! Native package adapters retain all producer-specific representations.

/// The canonical interactive editing state of an iWork table.
pub use litchi_iwa_common::table::lock::State;

#[cfg(test)]
mod tests {
    use std::mem::size_of;

    use super::State;

    #[test]
    fn lock_state_is_a_compact_closed_value() {
        assert_eq!(size_of::<State>(), 1);
        assert_eq!(State::default(), State::Unlocked);
        assert_eq!(State::from_locked(false), State::Unlocked);
        assert_eq!(State::from_locked(true), State::Locked);
        assert!(!State::Unlocked.is_locked());
        assert!(State::Locked.is_locked());
    }

    #[test]
    fn lock_state_has_value_traits() {
        fn assert_traits<T: Clone + Copy + Eq + std::hash::Hash + std::fmt::Debug>() {}

        assert_traits::<State>();
    }
}
