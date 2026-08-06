//! Immutable source-preserving link snapshots.

use super::{Link, Patch, Revision, Transaction, validation};
use litchi_cfb::OleError;
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
}

impl Snapshot {
    /// Parses and retains a bounded OLEDS link stream.
    pub fn parse(bytes: &[u8]) -> Result<Self, OleError> {
        Self::parse_shared(Arc::<[u8]>::from(bytes))
    }

    /// Parses an already shared OLEDS link stream without copying its bytes.
    pub fn parse_shared(bytes: Arc<[u8]>) -> Result<Self, OleError> {
        let link = Link::parse_shared(bytes)?;
        Self::from_link(link)
    }

    /// Captures an existing parsed link as a source-preserving snapshot.
    pub fn from_link(link: Link) -> Result<Self, OleError> {
        validation::validate(&link)?;
        let revision = Revision::of(link.bytes());
        Ok(Self { link, revision })
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

    /// Starts an isolated typed edit based on this snapshot.
    #[must_use]
    pub fn edit(&self) -> Transaction {
        Transaction::new(self.clone())
    }

    pub(crate) fn patch_to(&self, after: Snapshot) -> Patch {
        Patch::new(self.clone(), after)
    }
}

impl std::ops::Deref for Snapshot {
    type Target = Link;

    fn deref(&self) -> &Self::Target {
        self.link()
    }
}
