//! Archive-free interactive table-lock semantics.

/// Interactive editing state of a native iWork table.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, Default)]
pub enum State {
    /// The table can be selected and edited interactively.
    #[default]
    Unlocked,
    /// The table is protected from interactive editing.
    Locked,
}

impl State {
    /// Construct a lock state from its native boolean representation.
    #[must_use]
    pub const fn from_locked(locked: bool) -> Self {
        if locked { Self::Locked } else { Self::Unlocked }
    }

    /// Return whether the table is protected from interactive editing.
    #[must_use]
    pub const fn is_locked(self) -> bool {
        matches!(self, Self::Locked)
    }
}

#[cfg(test)]
mod tests {
    use std::mem::size_of;

    use super::State;

    #[test]
    fn state_is_a_compact_closed_value() {
        assert_eq!(size_of::<State>(), 1);
        assert_eq!(State::from_locked(false), State::Unlocked);
        assert_eq!(State::from_locked(true), State::Locked);
        assert!(!State::Unlocked.is_locked());
        assert!(State::Locked.is_locked());
    }
}
