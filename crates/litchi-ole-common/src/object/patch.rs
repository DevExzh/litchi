//! Reversible snapshot commits for target-selected OLE edits.

use super::snapshot::Snapshot;
use crate::Error;
use std::sync::Arc;

/// A deterministic whole-artifact replacement from one OLE snapshot to the
/// next.
///
/// The common object layer deliberately keeps patches at the artifact
/// boundary. Host crates retain responsibility for semantic operations and
/// dependency closure; this value only makes the already-validated before and
/// after snapshots explicit and safely replayable.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Patch {
    before: Arc<[u8]>,
    after: Arc<[u8]>,
}

impl Patch {
    pub(crate) fn new(before: Vec<u8>, after: Vec<u8>) -> Self {
        Self {
            before: before.into(),
            after: after.into(),
        }
    }

    /// Bytes required as the source of this patch.
    #[must_use]
    pub fn before(&self) -> &[u8] {
        &self.before
    }

    /// Bytes produced by this patch.
    #[must_use]
    pub fn after(&self) -> &[u8] {
        &self.after
    }

    /// Whether the edit did not alter the serialized OLE artifact.
    #[must_use]
    pub fn is_noop(&self) -> bool {
        self.before == self.after
    }

    /// Returns the exact inverse replacement.
    #[must_use]
    pub fn inverse(&self) -> Self {
        Self {
            before: Arc::clone(&self.after),
            after: Arc::clone(&self.before),
        }
    }

    /// Applies the patch only to its expected source snapshot.
    ///
    /// A source mismatch is a typed conflict rather than a last-writer-wins
    /// replacement, which keeps the common layer safe for snapshot joins.
    ///
    /// # Errors
    ///
    /// Returns an error when `source` does not exactly match the artifact this
    /// patch was created from.
    #[inline]
    pub fn apply(&self, source: &[u8]) -> Result<Vec<u8>, Error> {
        if source != self.before.as_ref() {
            return Err(Error::PatchConflict);
        }
        Ok(self.after.as_ref().to_vec())
    }
}

/// The validated result of an OLE object edit: the post-edit snapshot, its
/// reversible artifact patch and diagnostics.
pub type Commit = litchi_core::edit::Commit<Snapshot, Patch>;

/// Content-free diagnostics of one OLE object commit.
pub use litchi_core::edit::Diagnostics;

#[cfg(test)]
#[allow(
    clippy::unwrap_used,
    clippy::expect_used,
    reason = "tests use concise assertions while exercising fallible patch application"
)]
mod tests {
    use super::Patch;

    #[test]
    fn patch_requires_its_source_and_round_trips_inverse() {
        let patch = Patch::new(b"before".to_vec(), b"after".to_vec());
        assert!(!patch.is_noop());
        assert_eq!(patch.apply(b"before").unwrap(), b"after");
        assert!(patch.apply(b"other").is_err());
        assert_eq!(patch.inverse().apply(b"after").unwrap(), b"before");
    }

    #[test]
    fn equal_artifacts_are_a_noop() {
        let patch = Patch::new(b"same".to_vec(), b"same".to_vec());
        assert!(patch.is_noop());
    }
}
