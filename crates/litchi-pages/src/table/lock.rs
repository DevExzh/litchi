//! Archive-free interactive lock semantics for Pages body tables.

/// Canonical iWork interactive table-lock state.
pub use litchi_iwa_common::table::lock::State;

/// Semantic spelling used by Pages callers that prefer an application-owned
/// type name. It carries no native object or wire representation.
pub type BodyTableLockState = State;

#[cfg(test)]
mod tests {
    use std::mem::size_of;

    use super::State;

    #[test]
    fn body_table_lock_state_is_compact_and_closed() {
        assert_eq!(size_of::<State>(), 1);
        assert_eq!(State::default(), State::Unlocked);
        assert_eq!(State::from_locked(false), State::Unlocked);
        assert_eq!(State::from_locked(true), State::Locked);
        assert!(!State::Unlocked.is_locked());
        assert!(State::Locked.is_locked());
    }

    #[test]
    fn body_table_lock_state_has_value_traits() {
        fn assert_traits<T: Clone + Copy + Eq + std::hash::Hash + std::fmt::Debug>() {}

        assert_traits::<State>();
    }
}
