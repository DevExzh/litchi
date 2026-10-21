//! Immutable source-preserving link snapshots.

use super::{Edit, Limits, Link, Patch, Revision, validation};
use crate::Error;
use std::sync::Arc;

/// An immutable, cheaply clonable snapshot of one OLEDS `\x01Ole` stream.
///
/// The snapshot retains the exact source allocation and projects the existing
/// typed [`Link`] model. Unknown fields and bytes after the understood OLEDS
/// grammar remain attached to the snapshot and are replayed by a no-op edit.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Snapshot {
    pub(crate) link: Link,
    pub(crate) revision: Revision,
    pub(crate) limits: Limits,
}

impl Snapshot {
    /// Parses and retains a bounded OLEDS link stream.
    ///
    /// # Errors
    ///
    /// Returns an error when the stream is malformed or exceeds the OLEDS
    /// metadata limit.
    pub fn parse(bytes: &[u8], limits: Limits) -> Result<Self, Error> {
        Self::parse_shared(Arc::<[u8]>::from(bytes), limits)
    }

    /// Parses an already shared OLEDS link stream without copying its bytes.
    ///
    /// # Errors
    ///
    /// Returns an error when the stream is malformed or exceeds the OLEDS
    /// metadata limit.
    pub fn parse_shared(bytes: Arc<[u8]>, limits: Limits) -> Result<Self, Error> {
        let limits = limits.validate()?;
        let link = Link::parse_shared(bytes, limits)?;
        Self::from_link(link, limits)
    }

    /// Captures an existing parsed link as a source-preserving snapshot.
    ///
    /// # Errors
    ///
    /// Returns an error when the parsed link has inconsistent field offsets or
    /// otherwise violates the OLEDS layout invariants.
    pub fn from_link(link: Link, limits: Limits) -> Result<Self, Error> {
        let limits = limits.validate()?;
        validation::validate(&link)?;
        if link.bytes().len() > limits.max_bytes {
            return Err(Error::Limit {
                kind: crate::LimitKind::StreamBytes,
                observed: link.bytes().len() as u64,
                maximum: limits.max_bytes as u64,
            });
        }
        let revision = Revision::of(link.bytes());
        Ok(Self {
            link,
            revision,
            limits,
        })
    }

    /// Borrows the complete typed link projection.
    #[must_use]
    pub const fn link(&self) -> &Link {
        &self.link
    }

    /// Borrows the exact source bytes retained by this snapshot.
    #[must_use]
    pub fn bytes(&self) -> &[u8] {
        self.link.bytes()
    }

    /// Returns shared ownership of the exact source allocation.
    #[must_use]
    pub fn bytes_shared(&self) -> Arc<[u8]> {
        self.link.bytes_shared()
    }

    /// Returns the deterministic identity of the exact source bytes.
    #[must_use]
    pub const fn revision(&self) -> Revision {
        self.revision
    }

    /// Returns the source fingerprint as a compact scalar.
    #[must_use]
    pub const fn fingerprint(&self) -> u64 {
        self.revision.value()
    }

    /// Limits retained for subsequent edits and replay.
    #[must_use]
    pub const fn limits(&self) -> Limits {
        self.limits
    }

    /// Starts an isolated typed edit based on this snapshot.
    #[must_use]
    pub fn edit(&self) -> Edit {
        Edit::new(self.clone())
    }

    pub(crate) fn patch_to(&self, after: &Snapshot) -> Patch {
        Patch::new(self, after)
    }
}

impl std::ops::Deref for Snapshot {
    type Target = Link;

    fn deref(&self) -> &Self::Target {
        self.link()
    }
}
