//! Reversible snapshot commits for target-selected OLE edits.

use super::snapshot::Snapshot;
use litchi_cfb::OleError;
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
    pub fn apply(&self, source: &[u8]) -> Result<Vec<u8>, OleError> {
        if source != self.before.as_ref() {
            return Err(OleError::InvalidFormat(
                "OLE patch source snapshot does not match".into(),
            ));
        }
        Ok(self.after.as_ref().to_vec())
    }
}

/// The validated result of an OLE object edit.
#[derive(Debug, Clone)]
pub struct Commit {
    snapshot: Snapshot,
    patch: Patch,
}

impl Commit {
    pub(crate) fn new(snapshot: Snapshot, patch: Patch) -> Self {
        Self { snapshot, patch }
    }

    /// The immutable post-edit snapshot.
    #[must_use]
    pub fn snapshot(&self) -> &Snapshot {
        &self.snapshot
    }

    /// The reversible artifact patch.
    #[must_use]
    pub fn patch(&self) -> &Patch {
        &self.patch
    }

    /// Consumes the commit and returns its post-edit snapshot.
    #[must_use]
    pub fn into_snapshot(self) -> Snapshot {
        self.snapshot
    }

    /// Consumes the commit and returns its patch.
    #[must_use]
    pub fn into_patch(self) -> Patch {
        self.patch
    }

    /// Splits the commit into its snapshot and reversible patch.
    #[must_use]
    pub fn into_parts(self) -> (Snapshot, Patch) {
        (self.snapshot, self.patch)
    }
}

#[cfg(test)]
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
