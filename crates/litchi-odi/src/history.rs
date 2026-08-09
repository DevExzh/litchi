//! Bounded, source-checked history shared by flat and packaged ODI snapshots.
#![allow(
    clippy::arbitrary_source_item_ordering,
    reason = "The artifact contract is introduced before the generic history that uses it."
)]

use crate::{FlatImage, Image};
use litchi_core::{Error, Result};

const MAX_HISTORY_STATES: usize = 1_024;

/// A byte-identifiable immutable ODI artifact suitable for [`History`].
pub trait HistoryArtifact: Clone {
    /// Returns the exact serialized artifact bytes.
    fn history_bytes(&self) -> &[u8];
}

impl HistoryArtifact for FlatImage {
    fn history_bytes(&self) -> &[u8] {
        self.as_bytes()
    }
}

impl HistoryArtifact for Image {
    fn history_bytes(&self) -> &[u8] {
        self.as_bytes()
    }
}

/// Bounded undo/redo history for either flat XML or package snapshots.
///
/// Every transition is checked against the exact current source bytes. New
/// transitions after an undo discard the abandoned redo branch.
pub struct History<T: HistoryArtifact> {
    states: Vec<T>,
    cursor: usize,
    capacity: usize,
}

impl<T: HistoryArtifact> History<T> {
    /// Creates a history retaining at most `capacity` immutable states.
    ///
    /// # Errors
    ///
    /// Returns an error unless capacity is between 1 and 1024 inclusive.
    pub fn new(initial: T, capacity: usize) -> Result<Self> {
        if !(1..=MAX_HISTORY_STATES).contains(&capacity) {
            return Err(Error::InvalidFormat(
                "ODI history capacity must be between 1 and 1024".to_string(),
            ));
        }
        Ok(Self {
            states: vec![initial],
            cursor: 0,
            capacity,
        })
    }

    /// Returns the current immutable snapshot.
    #[must_use]
    pub fn current(&self) -> &T {
        &self.states[self.cursor]
    }

    /// Returns the number of retained states.
    #[must_use]
    pub fn len(&self) -> usize {
        self.states.len()
    }

    /// Returns whether no states are retained. A valid history is never empty.
    #[must_use]
    pub fn is_empty(&self) -> bool {
        self.states.is_empty()
    }

    /// Records a source-bound transition and returns whether bytes changed.
    ///
    /// # Errors
    ///
    /// Returns an error if `source` is not byte-identical to the current state.
    pub fn record(&mut self, source: &T, target: T) -> Result<bool> {
        if self.current().history_bytes() != source.history_bytes() {
            return Err(Error::InvalidFormat(
                "stale ODI history transition source".to_string(),
            ));
        }
        if source.history_bytes() == target.history_bytes() {
            return Ok(false);
        }
        self.states.truncate(self.cursor + 1);
        if self.states.len() == self.capacity {
            self.states.remove(0);
            self.cursor = self.cursor.saturating_sub(1);
        }
        self.states.push(target);
        self.cursor = self.states.len() - 1;
        Ok(true)
    }

    /// Moves to the preceding state, if retained.
    pub fn undo(&mut self) -> Option<&T> {
        if self.cursor == 0 {
            return None;
        }
        self.cursor -= 1;
        Some(self.current())
    }

    /// Moves to the following state, if retained.
    pub fn redo(&mut self) -> Option<&T> {
        if self.cursor + 1 >= self.states.len() {
            return None;
        }
        self.cursor += 1;
        Some(self.current())
    }
}
