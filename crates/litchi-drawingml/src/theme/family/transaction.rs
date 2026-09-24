//! Immutable source snapshots and source-checked Theme Family edits.

use std::sync::Arc;

use crate::{Error, Result};

use super::{Family, codec};

/// An immutable, cheap-to-share Theme Family source snapshot.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Snapshot {
    xml: Arc<[u8]>,
    value: Family,
}

impl Snapshot {
    /// Parse and retain a bounded Theme Family fragment.
    ///
    /// Validation precedes the source copy. Already shared callers can use
    /// [`Self::from_shared_xml`] to retain their existing allocation.
    pub fn from_xml(xml: impl AsRef<[u8]>) -> Result<Self> {
        let value = codec::read(xml.as_ref())?;
        let source = value.source_state().ok_or_else(|| {
            Error::Invalid("theme family reader did not retain its source".into())
        })?;
        let xml = Arc::clone(&source.xml);
        Ok(Self { xml, value })
    }

    /// Parse and retain an already shared Theme Family fragment allocation.
    pub fn from_shared_xml(xml: Arc<[u8]>) -> Result<Self> {
        let value = codec::read_shared(Arc::clone(&xml))?;
        Ok(Self { xml, value })
    }

    /// Create a source snapshot from a typed family.
    ///
    /// Detached values use canonical XML. Parsed values retain their source
    /// and apply any staged scalar changes without discarding opaque markup.
    pub fn new(value: Family) -> Result<Self> {
        let xml = codec::write(&value)?;
        Self::from_xml(xml)
    }

    /// Borrow the exact source bytes retained by this snapshot.
    #[must_use]
    pub fn xml_bytes(&self) -> &[u8] {
        &self.xml
    }

    /// Borrow the typed semantic family projection.
    #[must_use]
    pub const fn value(&self) -> &Family {
        &self.value
    }

    /// Start an isolated detached edit.
    #[must_use]
    pub fn edit(&self) -> Edit {
        Edit {
            before: self.clone(),
            staged: self.value.detached(),
        }
    }
}

/// A detached Theme Family edit that is not visible until commit.
#[derive(Debug, Clone)]
pub struct Edit {
    before: Snapshot,
    staged: Family,
}

impl Edit {
    /// Borrow the immutable source snapshot used for stale and inverse checks.
    #[must_use]
    pub const fn before(&self) -> &Snapshot {
        &self.before
    }

    /// Borrow the currently staged semantic family.
    #[must_use]
    pub const fn value(&self) -> &Family {
        &self.staged
    }

    /// Replace the complete typed family projection.
    pub fn replace(&mut self, value: Family) -> Result<&mut Self> {
        // `Family` setters validate every modeled scalar before publication.
        // Keep this operation detached and defer source patching/read-back to
        // commit so replacement does not serialize a throwaway whole value.
        self.staged = value.detached();
        Ok(self)
    }

    /// Set the applied theme name.
    pub fn set_name(&mut self, name: impl AsRef<str>) -> Result<&mut Self> {
        self.staged.set_name(name)?;
        Ok(self)
    }

    /// Set the applied theme GUID.
    pub fn set_id(&mut self, id: impl AsRef<str>) -> Result<&mut Self> {
        self.staged.set_id(id)?;
        Ok(self)
    }

    /// Set the applied variant GUID.
    pub fn set_variant_id(&mut self, variant_id: impl AsRef<str>) -> Result<&mut Self> {
        self.staged.set_variant_id(variant_id)?;
        Ok(self)
    }

    /// Whether staged semantic fields differ from the source snapshot.
    #[must_use]
    pub fn is_changed(&self) -> bool {
        !self.before.value.semantic_eq(&self.staged)
    }

    /// Validate and publish the edit with an exact reversible source patch.
    pub fn commit(self) -> Result<Commit> {
        if !self.is_changed() {
            return Ok(Commit {
                snapshot: self.before.clone(),
                patch: Patch {
                    before: self.before.clone(),
                    after: self.before.clone(),
                },
                changed: false,
            });
        }

        let mut source_value = self.before.value.clone();
        if source_value.name() != self.staged.name() {
            source_value.set_name(self.staged.name())?;
        }
        if source_value.id() != self.staged.id() {
            source_value.set_id(self.staged.id().as_str())?;
        }
        if source_value.variant_id() != self.staged.variant_id() {
            source_value.set_variant_id(self.staged.variant_id().as_str())?;
        }
        let xml = codec::write(&source_value)?;
        let snapshot = Snapshot::from_xml(xml)?;
        if snapshot.value() != &self.staged {
            return Err(Error::Invalid(
                "theme family edit read-back did not match staged fields".into(),
            ));
        }
        let patch = Patch {
            before: self.before.clone(),
            after: snapshot.clone(),
        };
        Ok(Commit {
            snapshot,
            patch,
            changed: true,
        })
    }
}

/// A successful Theme Family publication.
#[derive(Debug, Clone)]
pub struct Commit {
    snapshot: Snapshot,
    patch: Patch,
    changed: bool,
}

impl Commit {
    /// Whether the publication changed any bytes.
    #[must_use]
    pub const fn changed(&self) -> bool {
        self.changed
    }

    /// Borrow the resulting snapshot.
    #[must_use]
    pub const fn snapshot(&self) -> &Snapshot {
        &self.snapshot
    }

    /// Move the resulting snapshot out of this commit.
    #[must_use]
    pub fn into_snapshot(self) -> Snapshot {
        self.snapshot
    }

    /// Borrow the exact reversible patch.
    #[must_use]
    pub const fn patch(&self) -> &Patch {
        &self.patch
    }

    /// Move the reversible patch out of this commit.
    #[must_use]
    pub fn into_patch(self) -> Patch {
        self.patch
    }
}

/// An exact-source-checked and reversible Theme Family patch.
///
/// This in-memory patch shares its validated source and target snapshots.
/// Clone and inversion do not copy XML; application to independently reopened
/// bytes may compare the complete fragment. Durable serialization, composition,
/// and three-way merging are not provided by this fragment owner.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Patch {
    before: Snapshot,
    after: Snapshot,
}

impl Patch {
    /// Borrow the exact source bytes required by this patch.
    #[must_use]
    pub fn before_xml(&self) -> &[u8] {
        self.before.xml_bytes()
    }

    /// Borrow the exact bytes produced by this patch.
    #[must_use]
    pub fn after_xml(&self) -> &[u8] {
        self.after.xml_bytes()
    }

    /// Return the inverse patch.
    #[must_use]
    pub fn inverse(self) -> Self {
        Self {
            before: self.after,
            after: self.before,
        }
    }

    /// Apply this patch only to the exact source snapshot from which it came.
    pub fn apply(&self, source: &Snapshot) -> Result<Snapshot> {
        if !Arc::ptr_eq(&source.xml, &self.before.xml)
            && source.xml.as_ref() != self.before.xml.as_ref()
        {
            return Err(Error::Invalid(
                "theme family patch source does not match its byte precondition".into(),
            ));
        }
        if Arc::ptr_eq(&self.before.xml, &self.after.xml) {
            return Ok(source.clone());
        }
        Ok(self.after.clone())
    }
}
