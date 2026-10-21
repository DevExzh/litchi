//! Reversible snapshot commits for target-selected OLE edits.

use super::options::OpenOptions;
use super::snapshot::Package;
use crate::Error;
use std::fmt;
use std::sync::Arc;

/// A deterministic whole-artifact replacement from one OLE snapshot to the
/// next.
///
/// The common object layer deliberately keeps patches at the artifact
/// boundary. Host crates retain responsibility for semantic operations and
/// dependency closure; this value only makes the already-validated before and
/// after snapshots explicit and safely replayable.
#[derive(Clone, PartialEq, Eq)]
pub struct Patch {
    before: Arc<Vec<u8>>,
    after: Arc<Vec<u8>>,
}

impl fmt::Debug for Patch {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("Patch")
            .field("before_len", &self.before.len())
            .field("after_len", &self.after.len())
            .field("is_noop", &self.is_noop())
            .finish()
    }
}

impl Patch {
    #[cfg(test)]
    pub(crate) fn new(before: Vec<u8>, after: Vec<u8>) -> Self {
        Self::new_shared(Arc::new(before), Arc::new(after))
    }

    pub(crate) fn new_shared(before: Arc<Vec<u8>>, after: Arc<Vec<u8>>) -> Self {
        Self { before, after }
    }

    /// Bytes required as the source of this patch.
    #[must_use]
    pub fn before(&self) -> &[u8] {
        self.before.as_slice()
    }

    /// Bytes produced by this patch.
    #[must_use]
    pub fn after(&self) -> &[u8] {
        self.after.as_slice()
    }

    /// Whether the edit did not alter the serialized OLE artifact.
    #[must_use]
    pub fn is_noop(&self) -> bool {
        self.before.as_slice() == self.after.as_slice()
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
    #[cfg(test)]
    fn apply_bytes(&self, source: &[u8]) -> Result<Vec<u8>, Error> {
        if source != self.before.as_slice() {
            return Err(Error::PatchConflict);
        }
        Ok(self.after.as_ref().to_vec())
    }

    /// Applies this patch to a package with matching recorded content.
    pub fn apply(&self, source: &Package) -> Result<Commit, Error> {
        // A committed Package retains the original source allocation for
        // exact no-op output, but its current captured state may have changed.
        // Compare the published content, rather than only that retained
        // ingress allocation, so a patch cannot be replayed onto its own
        // after-package or another changed branch.
        if source.current_bytes() != self.before.as_slice() {
            return Err(Error::PatchConflict);
        }
        if self.is_noop() {
            return Ok(Commit::new(
                source.clone(),
                self.clone(),
                Diagnostics::UNCHANGED,
            ));
        }
        let package = Package::open_with_options(
            self.after.as_ref().clone(),
            OpenOptions::new()
                .targets(source.targets().clone())
                .limits(source.limits())
                .cfb_limits(source.cfb_limits()),
        )?;
        Ok(Commit::new(package, self.clone(), Diagnostics::CHANGED))
    }
}

/// The validated result of an OLE object edit: the post-edit snapshot, its
/// reversible artifact patch and diagnostics.
pub type Commit = litchi_core::edit::Commit<Package, Patch>;

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
        assert_eq!(patch.apply_bytes(b"before").unwrap(), b"after");
        assert!(patch.apply_bytes(b"other").is_err());
        assert_eq!(patch.inverse().apply_bytes(b"after").unwrap(), b"before");
    }

    #[test]
    fn equal_artifacts_are_a_noop() {
        let patch = Patch::new(b"same".to_vec(), b"same".to_vec());
        assert!(patch.is_noop());
    }

    #[test]
    fn debug_reports_shape_without_artifact_bytes() {
        let patch = Patch::new(
            b"private-before-payload".to_vec(),
            b"private-after-payload".to_vec(),
        );
        let debug = format!("{patch:?}");
        assert_eq!(
            debug,
            "Patch { before_len: 22, after_len: 21, is_noop: false }"
        );
    }
}
