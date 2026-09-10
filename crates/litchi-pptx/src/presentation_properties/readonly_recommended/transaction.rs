//! Detached source-checked edits for `readonlyRecommended`.

use std::sync::Arc;

use litchi_opc::OpcPackage;

use super::codec;
use crate::{Error, Result};

/// Stable fingerprint of the exact presentation-properties XML source.
pub type Revision = u64;

/// An immutable typed view bound to one exact presentation-properties source.
#[derive(Clone, Debug)]
pub struct Snapshot {
    pub(crate) source_part_name: String,
    pub(crate) source_xml: Arc<Vec<u8>>,
    pub(crate) located: codec::Located,
    pub(crate) revision: Revision,
}

impl Snapshot {
    /// Load the optional source-bound presentation-properties owner.
    ///
    /// # Errors
    ///
    /// Returns an error if the owner or its package graph is malformed.
    pub fn load(package: &OpcPackage) -> Result<Option<Self>> {
        super::package::load_snapshot(package)
    }

    /// Alias emphasizing the source-bound read.
    ///
    /// # Errors
    ///
    /// Returns an error if the owner or its package graph is malformed.
    pub fn read(package: &OpcPackage) -> Result<Option<Self>> {
        Self::load(package)
    }

    /// Parse a complete `p:presentationPr` source and retain its exact bytes.
    ///
    /// # Errors
    ///
    /// Returns an error if the source is malformed, ambiguous, or exceeds the
    /// bounded presentation-properties policy.
    pub fn from_xml(source: impl AsRef<[u8]>) -> Result<Self> {
        let source = source.as_ref();
        if source.len() > codec::MAX_BYTES {
            return Err(Error::Limit {
                resource: "presentation-properties readonlyRecommended",
                limit: codec::MAX_BYTES,
            });
        }
        Self::from_xml_owned(source.to_vec())
    }

    /// Parse owned presentation-properties XML without another source copy.
    ///
    /// # Errors
    ///
    /// Returns an error if the source is malformed, ambiguous, or exceeds the
    /// bounded presentation-properties policy.
    pub fn from_xml_owned(source: Vec<u8>) -> Result<Self> {
        let source_xml = Arc::new(source);
        let located = codec::locate(source_xml.as_slice())?;
        Ok(Self::from_located(String::new(), source_xml, located))
    }

    pub(crate) fn from_part(source_part_name: String, source_xml: Arc<Vec<u8>>) -> Result<Self> {
        let located = codec::locate(source_xml.as_slice())?;
        Ok(Self::from_located(source_part_name, source_xml, located))
    }

    fn from_located(
        source_part_name: String,
        source_xml: Arc<Vec<u8>>,
        located: codec::Located,
    ) -> Self {
        let revision = fingerprint(source_xml.as_slice());
        Self {
            source_part_name,
            source_xml,
            located,
            revision,
        }
    }

    /// Return the owning presentation-properties part name, when loaded from
    /// an OPC package.
    #[inline]
    #[must_use]
    pub fn source_part_name(&self) -> &str {
        &self.source_part_name
    }

    /// Alias using the generic OPC owner vocabulary.
    #[inline]
    #[must_use]
    pub fn part_name(&self) -> &str {
        self.source_part_name()
    }

    /// Return the optional typed recommendation. `None` means the extension is
    /// absent; it is distinct from `Some(false)`.
    #[inline]
    #[must_use]
    pub const fn value(&self) -> Option<bool> {
        self.located.value
    }

    /// Alias emphasizing the XML element's semantic name.
    #[inline]
    #[must_use]
    pub const fn readonly_recommended(&self) -> Option<bool> {
        self.value()
    }

    /// Alias using the conventional Rust spelling.
    #[inline]
    #[must_use]
    pub const fn read_only_recommended(&self) -> Option<bool> {
        self.value()
    }

    /// Borrow the exact source XML captured by this snapshot.
    #[inline]
    #[must_use]
    pub fn source_xml(&self) -> &[u8] {
        self.source_xml.as_slice()
    }

    /// Return the source fingerprint used for stale-source checks.
    #[inline]
    #[must_use]
    pub const fn revision(&self) -> Revision {
        self.revision
    }

    /// Start an atomic detached edit.
    #[inline]
    #[must_use]
    pub fn edit(&self) -> Transaction {
        Transaction {
            original: self.clone(),
            working: self.value(),
        }
    }

    pub(crate) fn same_source(&self, other: &Self) -> bool {
        self.source_part_name == other.source_part_name
            && self.source_xml.as_slice() == other.source_xml.as_slice()
            && self.revision == other.revision
            && self.value() == other.value()
    }

    pub(crate) fn source_arc(&self) -> &Arc<Vec<u8>> {
        &self.source_xml
    }
}

/// A bounded edit staged against one presentation-properties source.
#[derive(Clone, Debug)]
pub struct Transaction {
    original: Snapshot,
    working: Option<bool>,
}

impl Transaction {
    /// Borrow the projected optional value.
    #[inline]
    #[must_use]
    pub const fn value(&self) -> Option<bool> {
        self.working
    }

    /// Alias for [`Self::value`].
    #[inline]
    #[must_use]
    pub const fn snapshot(&self) -> Option<bool> {
        self.value()
    }

    /// Return whether the staged semantic value differs from the source.
    #[inline]
    #[must_use]
    pub fn is_changed(&self) -> bool {
        self.original.value() != self.working
    }

    /// Set or clear the optional recommendation.
    ///
    /// `None` removes the recognized extension while preserving the enclosing
    /// extension list and every unrelated source byte.
    pub fn set(&mut self, value: Option<bool>) -> Result<bool> {
        let changed = self.working != value;
        self.working = value;
        Ok(changed)
    }

    /// Set the recommendation to a concrete boolean value.
    pub fn set_readonly_recommended(&mut self, value: bool) -> Result<bool> {
        self.set(Some(value))
    }

    /// Alias using the conventional Rust spelling.
    pub fn set_read_only_recommended(&mut self, value: bool) -> Result<bool> {
        self.set_readonly_recommended(value)
    }

    /// Remove the recognized recommendation.
    pub fn remove(&mut self) -> Result<bool> {
        self.set(None)
    }

    /// Validate and consume this edit into a reversible source patch.
    ///
    /// # Errors
    ///
    /// Returns an error if the source splice or typed readback fails.
    pub fn commit(self) -> Result<Commit> {
        if !self.is_changed() {
            let patch = Patch::new(self.original.clone(), self.original.clone());
            return Ok(Commit {
                snapshot: self.original,
                patch,
                changed: false,
            });
        }
        let updated = codec::rewrite(
            self.original.source_xml.as_slice(),
            &self.original.located,
            self.working,
        )?;
        let updated = Arc::new(updated);
        let located = codec::locate(updated.as_slice())?;
        if located.value != self.working {
            return Err(invalid(
                "readonlyRecommended serialization did not round-trip",
            ));
        }
        let snapshot = Snapshot::from_located(
            self.original.source_part_name.clone(),
            Arc::clone(&updated),
            located,
        );
        let patch = Patch::new(self.original, snapshot.clone());
        Ok(Commit {
            snapshot,
            patch,
            changed: true,
        })
    }
}

/// A committed source-preserving recommendation edit.
#[derive(Clone, Debug)]
pub struct Commit {
    snapshot: Snapshot,
    patch: Patch,
    changed: bool,
}

impl Commit {
    #[inline]
    #[must_use]
    pub const fn is_changed(&self) -> bool {
        self.changed
    }

    #[inline]
    #[must_use]
    pub const fn changed(&self) -> bool {
        self.changed
    }

    #[inline]
    #[must_use]
    pub fn snapshot(&self) -> &Snapshot {
        &self.snapshot
    }

    #[inline]
    #[must_use]
    pub const fn before_value(&self) -> Option<bool> {
        self.patch.before.value()
    }

    #[inline]
    #[must_use]
    pub const fn after_value(&self) -> Option<bool> {
        self.patch.after.value()
    }

    #[inline]
    #[must_use]
    pub fn patch(&self) -> &Patch {
        &self.patch
    }

    #[must_use]
    pub fn into_patch(self) -> Patch {
        self.patch
    }
}

/// A reversible, source-checked replacement of one properties XML source.
#[derive(Clone, Debug)]
pub struct Patch {
    before: Snapshot,
    after: Snapshot,
}

impl Patch {
    fn new(before: Snapshot, after: Snapshot) -> Self {
        Self { before, after }
    }

    #[inline]
    #[must_use]
    pub fn before(&self) -> &Snapshot {
        &self.before
    }

    #[inline]
    #[must_use]
    pub fn after(&self) -> &Snapshot {
        &self.after
    }

    #[inline]
    #[must_use]
    pub fn is_empty(&self) -> bool {
        self.before.same_source(&self.after)
    }

    #[inline]
    #[must_use]
    pub fn is_changed(&self) -> bool {
        !self.is_empty()
    }

    #[inline]
    #[must_use]
    pub const fn expected_revision(&self) -> Revision {
        self.before.revision
    }

    #[must_use]
    pub fn inverse(&self) -> Self {
        Self {
            before: self.after.clone(),
            after: self.before.clone(),
        }
    }

    /// Apply this patch to an exact presentation-properties XML buffer.
    ///
    /// # Errors
    ///
    /// Returns an error when the target is stale or the published source does
    /// not reparse to the committed typed value.
    pub fn apply(&self, target: &mut Vec<u8>) -> Result<Snapshot> {
        if target.as_slice() != self.before.source_xml.as_slice() {
            return Err(invalid("presentation-properties source is stale"));
        }
        if self.is_empty() {
            return Ok(self.before.clone());
        }
        let located = codec::locate(self.after.source_xml.as_slice())?;
        if located.value != self.after.value() {
            return Err(invalid(
                "published readonlyRecommended source differs from the patch",
            ));
        }
        let result = Snapshot::from_located(
            self.after.source_part_name.clone(),
            Arc::clone(&self.after.source_xml),
            located,
        );
        if !result.same_source(&self.after) {
            return Err(invalid(
                "published readonlyRecommended source differs from the patch",
            ));
        }
        *target = self.after.source_xml.as_ref().clone();
        Ok(result)
    }

    /// Alias emphasizing the detached source target.
    pub fn apply_to(&self, target: &mut Vec<u8>) -> Result<Snapshot> {
        self.apply(target)
    }

    /// Apply this patch atomically to its owning OPC package part.
    ///
    /// # Errors
    ///
    /// Returns an error when the package source is stale or the publication
    /// policy rejects the mutation.
    pub fn apply_package(&self, target: &mut OpcPackage) -> Result<Snapshot> {
        super::package::apply_patch(target, self)
    }
}

fn fingerprint(bytes: &[u8]) -> Revision {
    let mut hash = 0xcbf2_9ce4_8422_2325u64;
    for byte in bytes {
        hash ^= u64::from(*byte);
        hash = hash.wrapping_mul(0x100_0000_01b3);
    }
    hash
}

fn invalid(message: impl Into<String>) -> Error {
    Error::Invalid(message.into())
}
