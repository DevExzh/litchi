//! Bounded, source-checked history shared by flat and packaged ODI snapshots.
use crate::{Commit, FlatImage, FlatImageCommit, Image};
use litchi_core::{Error, Result};

const MAX_HISTORY_STATES: usize = 1_024;

/// A byte-identifiable immutable ODI artifact suitable for [`History`].
pub trait HistoryArtifact: Clone {
    /// Returns the exact serialized artifact bytes.
    fn history_bytes(&self) -> &[u8];
}

/// A validated commit whose source and target can be recorded atomically.
pub trait CommittedTransition<T: HistoryArtifact> {
    /// Returns the exact source snapshot.
    fn source(&self) -> &T;
    /// Returns the validated target snapshot.
    fn target(&self) -> &T;
}

impl CommittedTransition<Image> for Commit {
    fn source(&self) -> &Image {
        self.source()
    }

    fn target(&self) -> &Image {
        self.image()
    }
}

impl CommittedTransition<FlatImage> for FlatImageCommit {
    fn source(&self) -> &FlatImage {
        self.source()
    }

    fn target(&self) -> &FlatImage {
        self.snapshot()
    }
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
    byte_budget: usize,
    stored_bytes: usize,
}

impl<T: HistoryArtifact> History<T> {
    /// Creates a history retaining at most `capacity` immutable states.
    ///
    /// # Errors
    ///
    /// Returns an error unless capacity is between 1 and 1024 inclusive.
    pub fn new(initial: T, capacity: usize) -> Result<Self> {
        Self::with_byte_budget(initial, capacity, usize::MAX)
    }

    /// Creates a history bounded by both state count and serialized bytes.
    pub fn with_byte_budget(initial: T, capacity: usize, byte_budget: usize) -> Result<Self> {
        if !(1..=MAX_HISTORY_STATES).contains(&capacity) {
            return Err(Error::InvalidFormat(
                "ODI history capacity must be between 1 and 1024".to_string(),
            ));
        }
        let initial_bytes = initial.history_bytes().len();
        if byte_budget == 0 || initial_bytes > byte_budget {
            return Err(Error::InvalidFormat(
                "ODI history byte budget cannot retain its initial state".to_string(),
            ));
        }
        Ok(Self {
            states: vec![initial],
            cursor: 0,
            capacity,
            byte_budget,
            stored_bytes: initial_bytes,
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

    /// Returns the sum of exact serialized bytes retained by the history.
    #[must_use]
    pub const fn stored_bytes(&self) -> usize {
        self.stored_bytes
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
        let target_bytes = target.history_bytes().len();
        if target_bytes > self.byte_budget {
            return Err(Error::InvalidFormat(
                "ODI history target exceeds the byte budget".to_string(),
            ));
        }
        if self.capacity < 2
            || self
                .current()
                .history_bytes()
                .len()
                .saturating_add(target_bytes)
                > self.byte_budget
        {
            return Err(Error::InvalidFormat(
                "ODI history byte budget cannot retain the commit transition".to_string(),
            ));
        }
        for removed in self.states.drain(self.cursor + 1..) {
            self.stored_bytes = self
                .stored_bytes
                .checked_sub(removed.history_bytes().len())
                .ok_or_else(|| Error::InvalidFormat("ODI history byte underflow".to_string()))?;
        }
        while self.states.len() >= self.capacity
            || self.stored_bytes.saturating_add(target_bytes) > self.byte_budget
        {
            let removed = self.states.remove(0);
            self.stored_bytes = self
                .stored_bytes
                .checked_sub(removed.history_bytes().len())
                .ok_or_else(|| Error::InvalidFormat("ODI history byte underflow".to_string()))?;
            self.cursor = self.cursor.saturating_sub(1);
        }
        self.stored_bytes = self
            .stored_bytes
            .checked_add(target_bytes)
            .ok_or_else(|| Error::InvalidFormat("ODI history byte overflow".to_string()))?;
        self.states.push(target);
        self.cursor = self.states.len() - 1;
        Ok(true)
    }

    /// Records a validated commit together with its exact source binding.
    pub fn record_commit(&mut self, commit: &impl CommittedTransition<T>) -> Result<bool> {
        self.record(commit.source(), commit.target().clone())
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
