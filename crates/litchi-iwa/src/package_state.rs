//! Shared indexed storage for immutable iWork package snapshots.

use std::collections::HashMap;

/// Package entries plus an index for validated name lookups.
///
/// The state is kept behind an `Arc` by both the mutable package and immutable
/// snapshot types. Structural edits clone this state once, then rebuild the
/// small name index; read-only clones never duplicate either the entry bytes
/// or the index.
#[derive(Debug, Clone, Default)]
pub(crate) struct PackageState {
    pub(crate) entries: Vec<(String, Vec<u8>)>,
    positions: HashMap<String, usize>,
}

impl PackageState {
    pub(crate) fn from_entries(entries: Vec<(String, Vec<u8>)>) -> Self {
        let mut state = Self {
            entries,
            positions: HashMap::new(),
        };
        state.rebuild_positions();
        state
    }

    pub(crate) fn position(&self, name: &str) -> Option<usize> {
        self.positions.get(name).copied()
    }

    pub(crate) fn rebuild_positions(&mut self) {
        self.positions.clear();
        self.positions.reserve(self.entries.len());
        for (position, (name, _)) in self.entries.iter().enumerate() {
            let previous = self.positions.insert(name.clone(), position);
            debug_assert!(previous.is_none(), "package entry names must be unique");
        }
    }
}
