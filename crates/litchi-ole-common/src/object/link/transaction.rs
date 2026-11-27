//! Failure-atomic typed link transactions.

use super::{Link, Patch, Snapshot, Times, validation};
use crate::{Error, property_set::Guid};

/// A deterministic identity for one exact serialized OLEDS link stream.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct Revision(u64);

impl Revision {
    pub(crate) fn of(bytes: &[u8]) -> Self {
        let mut value = 0xcbf2_9ce4_8422_2325u64;
        for byte in bytes {
            value ^= u64::from(*byte);
            value = value.wrapping_mul(0x0000_0100_0000_01b3);
        }
        Self(value)
    }

    /// Returns the raw source fingerprint.
    #[must_use]
    pub const fn value(self) -> u64 {
        self.0
    }

    /// Alias for [`Self::value`].
    #[must_use]
    pub const fn fingerprint(self) -> u64 {
        self.value()
    }
}

/// An isolated typed edit over one source link snapshot.
#[derive(Debug, Clone)]
pub struct Edit {
    source: Snapshot,
    candidate: Link,
}

impl Edit {
    pub(crate) fn new(source: Snapshot) -> Self {
        Self {
            candidate: source.link.clone(),
            source,
        }
    }

    /// Borrows the source snapshot used for this transaction.
    #[must_use]
    pub const fn source(&self) -> &Snapshot {
        &self.source
    }

    /// Borrows the projected typed link value.
    #[must_use]
    pub const fn link(&self) -> &Link {
        &self.candidate
    }

    /// Whether any typed field differs from the source snapshot.
    #[must_use]
    fn has_changes(&self) -> bool {
        self.candidate != self.source.link
    }

    /// Updates the cache hint while retaining every other flags bit.
    pub fn set_cache_hint(&mut self, enabled: bool) -> &mut Self {
        self.candidate.set_cache_hint(enabled);
        self
    }

    /// Replaces the raw flags while preserving the embedded/linked layout.
    ///
    /// # Errors
    ///
    /// Returns an error when `flags` changes the embedded/linked layout.
    pub fn set_flags(&mut self, flags: u32) -> Result<&mut Self, Error> {
        self.candidate.set_flags(flags)?;
        Ok(self)
    }

    /// Replaces the implementation-specific OLEDS update option.
    pub fn set_link_update_option(&mut self, value: u32) -> &mut Self {
        self.candidate.set_link_update_option(value);
        self
    }

    /// Replaces the linked object's class identifier.
    ///
    /// # Errors
    ///
    /// Returns an error when the link has no class identifier field in its
    /// wire layout.
    pub fn set_class_id(&mut self, value: Guid) -> Result<&mut Self, Error> {
        self.candidate.set_class_id(value)?;
        Ok(self)
    }

    /// Replaces the linked object's FILETIME group.
    ///
    /// # Errors
    ///
    /// Returns an error when the link has no timestamp group in its wire
    /// layout.
    pub fn set_times(&mut self, value: Times) -> Result<&mut Self, Error> {
        self.candidate.set_times(value)?;
        Ok(self)
    }

    /// Applies a custom typed edit using the existing inert [`Link`] setters.
    ///
    /// # Errors
    ///
    /// Returns an error when `edit` fails or leaves the candidate with an
    /// invalid OLEDS layout.
    pub fn update<F>(&mut self, edit: F) -> Result<&mut Self, Error>
    where
        F: FnOnce(&mut Link) -> Result<(), Error>,
    {
        let mut candidate = self.candidate.clone();
        edit(&mut candidate)?;
        validation::validate(&candidate)?;
        self.candidate = candidate;
        Ok(self)
    }

    /// Validates and publishes the candidate without mutating its source.
    ///
    /// # Errors
    ///
    /// Returns an error when the candidate has an invalid OLEDS layout.
    pub fn commit(self) -> Result<Commit, Error> {
        let snapshot = self.materialize()?;
        let patch = self.source.patch_to(&snapshot);
        let diagnostics = Diagnostics::new(!patch.is_noop());
        Ok(Commit::new(snapshot, patch, diagnostics))
    }

    fn materialize(&self) -> Result<Snapshot, Error> {
        validation::validate(&self.candidate)?;
        if !self.has_changes() {
            return Ok(self.source.clone());
        }
        let bytes = self.candidate.to_bytes(self.source.limits())?.into();
        Snapshot::parse_shared(bytes, self.source.limits())
    }
}

/// A published edit: the new snapshot, its reversible patch and
/// diagnostics.
pub type Commit = litchi_core::edit::Commit<Snapshot, Patch>;

/// Content-free diagnostics of one commit.
pub use litchi_core::edit::Diagnostics;

/// Runs one typed link edit and publishes it atomically.
///
/// # Errors
///
/// Returns an error when `edit` fails or leaves the link with an invalid OLEDS
/// layout.
pub fn update<F>(snapshot: &Snapshot, edit: F) -> Result<Commit, Error>
where
    F: FnOnce(&mut Edit) -> Result<(), Error>,
{
    let mut transaction = snapshot.edit();
    edit(&mut transaction)?;
    transaction.commit()
}
