//! Shared indexed storage for immutable iWork package snapshots.

use std::collections::HashMap;
use std::sync::{Arc, Mutex};

use crate::archive::Archive;

/// Package entries plus an index for validated name lookups.
///
/// The state is kept behind an `Arc` by both the mutable package and immutable
/// snapshot types. Structural edits clone this state once, then rebuild the
/// small name index; read-only clones never duplicate either the entry bytes,
/// the index, or the single bounded parsed-archive cache.
#[derive(Debug, Default)]
pub(crate) struct PackageState {
    pub(crate) entries: Vec<(String, Vec<u8>)>,
    positions: HashMap<String, usize>,
    parsed_archive: Mutex<Option<(String, Arc<Archive>)>>,
}

impl Clone for PackageState {
    fn clone(&self) -> Self {
        let parsed_archive = self
            .parsed_archive
            .lock()
            .unwrap_or_else(|poisoned| poisoned.into_inner())
            .clone();
        Self {
            entries: self.entries.clone(),
            positions: self.positions.clone(),
            parsed_archive: Mutex::new(parsed_archive),
        }
    }
}

impl PackageState {
    pub(crate) fn from_entries(entries: Vec<(String, Vec<u8>)>) -> Self {
        let mut state = Self {
            entries,
            positions: HashMap::new(),
            parsed_archive: Mutex::new(None),
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
        self.clear_parsed_archive();
    }

    pub(crate) fn cached_archive(&self, name: &str) -> Option<Arc<Archive>> {
        self.parsed_archive
            .lock()
            .unwrap_or_else(|poisoned| poisoned.into_inner())
            .as_ref()
            .filter(|(cached_name, _)| cached_name == name)
            .map(|(_, archive)| Arc::clone(archive))
    }

    pub(crate) fn cache_archive(&self, name: &str, archive: Arc<Archive>) -> Arc<Archive> {
        let mut cached = self
            .parsed_archive
            .lock()
            .unwrap_or_else(|poisoned| poisoned.into_inner());
        if let Some((cached_name, cached_archive)) = cached.as_ref()
            && cached_name == name
        {
            return Arc::clone(cached_archive);
        }
        *cached = Some((name.to_owned(), Arc::clone(&archive)));
        archive
    }

    pub(crate) fn invalidate_archive(&mut self, name: &str) {
        let cached = self
            .parsed_archive
            .get_mut()
            .unwrap_or_else(|poisoned| poisoned.into_inner());
        if cached
            .as_ref()
            .is_some_and(|(cached_name, _)| cached_name == name)
        {
            *cached = None;
        }
    }

    fn clear_parsed_archive(&mut self) {
        let cached = self
            .parsed_archive
            .get_mut()
            .unwrap_or_else(|poisoned| poisoned.into_inner());
        *cached = None;
    }
}
