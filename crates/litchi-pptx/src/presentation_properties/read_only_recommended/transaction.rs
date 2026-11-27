//! Detached source-checked edits for `readonlyRecommended`.

use std::sync::Arc;

use litchi_opc::OpcPackage;

use super::codec;
use crate::presentation_properties::Limits;
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
    pub(crate) limits: Limits,
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

    /// Parse a complete `p:presentationPr` source and retain its exact bytes.
    ///
    /// # Errors
    ///
    /// Returns an error if the source is malformed, ambiguous, or exceeds the
    /// bounded presentation-properties policy.
    pub fn from_xml(source: impl AsRef<[u8]>, limits: Limits) -> Result<Self> {
        let source = source.as_ref();
        if source.len() > limits.max_input_bytes() {
            return Err(Error::Limit {
                kind: crate::opened::LimitKind::Codec(crate::opened::CodecLimitKind::XmlBytes),
                observed: source.len(),
                maximum: limits.max_input_bytes(),
            });
        }
        Self::from_xml_owned(source.to_vec(), limits)
    }

    /// Parse owned presentation-properties XML without another source copy.
    ///
    /// # Errors
    ///
    /// Returns an error if the source is malformed, ambiguous, or exceeds the
    /// bounded presentation-properties policy.
    pub fn from_xml_owned(source: Vec<u8>, limits: Limits) -> Result<Self> {
        let source_xml = Arc::new(source);
        let located = codec::locate(source_xml.as_slice(), limits)?;
        Ok(Self::from_located(
            String::new(),
            source_xml,
            located,
            limits,
        ))
    }

    pub(crate) fn from_part(
        source_part_name: String,
        source_xml: Arc<Vec<u8>>,
        limits: Limits,
    ) -> Result<Self> {
        let located = codec::locate(source_xml.as_slice(), limits)?;
        Ok(Self::from_located(
            source_part_name,
            source_xml,
            located,
            limits,
        ))
    }

    fn from_located(
        source_part_name: String,
        source_xml: Arc<Vec<u8>>,
        located: codec::Located,
        limits: Limits,
    ) -> Self {
        let revision = fingerprint(source_xml.as_slice());
        Self {
            source_part_name,
            source_xml,
            located,
            revision,
            limits,
        }
    }

    /// Return the owning presentation-properties part name, when loaded from
    /// an OPC package.
    #[inline]
    #[must_use]
    pub fn source_part_name(&self) -> &str {
        &self.source_part_name
    }

    /// Return the optional typed recommendation. `None` means the extension is
    /// absent; it is distinct from `Some(false)`.
    #[inline]
    #[must_use]
    pub const fn value(&self) -> Option<bool> {
        self.located.value
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
    pub(crate) fn edit(&self) -> Draft {
        Draft {
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
pub(crate) struct Draft {
    original: Snapshot,
    working: Option<bool>,
}

impl Draft {
    fn is_changed(&self) -> bool {
        self.original.value() != self.working
    }

    /// Set or clear the optional recommendation.
    ///
    /// `None` removes the recognized extension while preserving the enclosing
    /// extension list and every unrelated source byte.
    pub(crate) fn set(&mut self, value: Option<bool>) -> Result<bool> {
        let changed = self.working != value;
        self.working = value;
        Ok(changed)
    }

    /// Set the recommendation.
    pub(crate) fn set_read_only_recommended(&mut self, value: bool) -> Result<bool> {
        self.set(Some(value))
    }

    /// Remove the recognized recommendation.
    pub(crate) fn remove(&mut self) -> Result<bool> {
        self.set(None)
    }

    /// Validate and consume this edit into a reversible source patch.
    ///
    /// # Errors
    ///
    /// Returns an error if the source splice or typed readback fails.
    pub(crate) fn commit(self) -> Result<Prepared> {
        if !self.is_changed() {
            let patch = Change::new(self.original.clone(), self.original.clone());
            return Ok(Prepared {
                snapshot: self.original,
                patch,
                changed: false,
            });
        }
        let updated = codec::rewrite(
            self.original.source_xml.as_slice(),
            &self.original.located,
            self.working,
            self.original.limits,
        )?;
        let updated = Arc::new(updated);
        let located = codec::locate(updated.as_slice(), self.original.limits)?;
        if located.value != self.working {
            return Err(invalid(
                "readonlyRecommended serialization did not round-trip",
            ));
        }
        let snapshot = Snapshot::from_located(
            self.original.source_part_name.clone(),
            Arc::clone(&updated),
            located,
            self.original.limits,
        );
        let patch = Change::new(self.original, snapshot.clone());
        Ok(Prepared {
            snapshot,
            patch,
            changed: true,
        })
    }
}

/// A committed source-preserving recommendation edit.
#[derive(Clone, Debug)]
pub(crate) struct Prepared {
    #[cfg_attr(
        not(test),
        expect(dead_code, reason = "read by the module's source-level tests")
    )]
    snapshot: Snapshot,
    patch: Change,
    #[cfg_attr(
        not(test),
        expect(dead_code, reason = "read by the module's source-level tests")
    )]
    changed: bool,
}

impl Prepared {
    #[cfg(test)]
    fn is_changed(&self) -> bool {
        self.changed
    }

    /// The outcome of this edit or commit.
    #[must_use]
    #[cfg(test)]
    pub(crate) fn diagnostics(&self) -> crate::diagnostics::Diagnostics {
        crate::diagnostics::Diagnostics::new(self.is_changed())
    }

    #[inline]
    #[must_use]
    #[cfg(test)]
    pub(crate) fn snapshot(&self) -> &Snapshot {
        &self.snapshot
    }

    #[inline]
    #[must_use]
    pub(crate) const fn before_value(&self) -> Option<bool> {
        self.patch.before.value()
    }

    #[inline]
    #[must_use]
    #[cfg(test)]
    pub(crate) const fn after_value(&self) -> Option<bool> {
        self.patch.after.value()
    }

    #[inline]
    #[must_use]
    pub(crate) fn patch(&self) -> &Change {
        &self.patch
    }
}

/// A reversible, source-checked replacement of one properties XML source.
#[derive(Clone, Debug)]
pub(crate) struct Change {
    before: Snapshot,
    after: Snapshot,
}

impl Change {
    fn new(before: Snapshot, after: Snapshot) -> Self {
        Self { before, after }
    }

    #[inline]
    #[must_use]
    pub(crate) fn before(&self) -> &Snapshot {
        &self.before
    }

    #[inline]
    #[must_use]
    pub(crate) fn after(&self) -> &Snapshot {
        &self.after
    }

    #[inline]
    #[must_use]
    pub(crate) fn is_noop(&self) -> bool {
        self.before.same_source(&self.after)
    }

    #[must_use]
    #[cfg(test)]
    pub(crate) fn inverse(&self) -> Self {
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
    #[cfg(test)]
    pub(crate) fn apply(&self, target: &mut Vec<u8>) -> Result<Snapshot> {
        if target.as_slice() != self.before.source_xml.as_slice() {
            return Err(invalid("presentation-properties source is stale"));
        }
        if self.is_noop() {
            return Ok(self.before.clone());
        }
        let located = codec::locate(self.after.source_xml.as_slice(), self.after.limits)?;
        if located.value != self.after.value() {
            return Err(invalid(
                "published readonlyRecommended source differs from the patch",
            ));
        }
        let result = Snapshot::from_located(
            self.after.source_part_name.clone(),
            Arc::clone(&self.after.source_xml),
            located,
            self.after.limits,
        );
        if !result.same_source(&self.after) {
            return Err(invalid(
                "published readonlyRecommended source differs from the patch",
            ));
        }
        *target = self.after.source_xml.as_ref().clone();
        Ok(result)
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
