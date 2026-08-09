#![expect(
    clippy::shadow_reuse,
    reason = "parser bindings are intentionally refined after validation"
)]
//! Source-preserving OpenType snapshots, transactions, and patches.

use std::sync::Arc;

use crate::error::{Error, Result};

use super::codec;
use super::model::{Ligatures, NumForm, NumSpacing, OnOff, OpenType, StyleSet, StyleSetId};

/// An immutable complete run/rPr snapshot with its typed OpenType projection.
#[derive(Debug, Clone)]
pub struct Snapshot {
    xml: Arc<[u8]>,
    value: OpenType,
    fingerprint: u64,
}

impl Snapshot {
    /// Parse and retain a bounded source fragment.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation cannot be completed.
    pub fn from_xml(xml: impl Into<Vec<u8>>) -> Result<Self> {
        let xml = xml.into();
        let value = codec::parse(&xml)?;
        Ok(Self {
            fingerprint: fingerprint(&xml),
            xml: Arc::from(xml.into_boxed_slice()),
            value,
        })
    }

    /// Borrow the exact source bytes retained by this snapshot.
    #[must_use]
    pub fn xml_bytes(&self) -> &[u8] {
        &self.xml
    }

    /// Borrow the typed OpenType projection.
    #[must_use]
    pub const fn open_type(&self) -> &OpenType {
        &self.value
    }

    /// Start a source-checked transaction.
    #[must_use]
    pub fn edit(&self) -> Transaction {
        Transaction {
            base: self.clone(),
            next: self.value.clone(),
        }
    }
}

/// A detached OpenType edit against one immutable source snapshot.
#[derive(Debug, Clone)]
pub struct Transaction {
    base: Snapshot,
    next: OpenType,
}

impl Transaction {
    /// Borrow the projected OpenType value.
    #[must_use]
    pub const fn open_type(&self) -> &OpenType {
        &self.next
    }

    pub fn set_ligatures(&mut self, value: Option<Ligatures>) -> &mut Self {
        self.next.ligatures = value;
        self
    }

    pub fn set_num_form(&mut self, value: Option<NumForm>) -> &mut Self {
        self.next.num_form = value;
        self
    }

    pub fn set_num_spacing(&mut self, value: Option<NumSpacing>) -> &mut Self {
        self.next.num_spacing = value;
        self
    }

    pub fn set_cntxt_alts(&mut self, value: Option<OnOff>) -> &mut Self {
        self.next.cntxt_alts = value;
        self
    }

    ///
    /// # Errors
    ///
    /// Returns an error if the operation cannot be completed.
    pub fn set_style_set(&mut self, value: StyleSet) -> Result<&mut Self> {
        self.next.set_style_set(value)?;
        Ok(self)
    }

    pub fn remove_style_set(&mut self, id: StyleSetId) -> &mut Self {
        self.next.remove_style_set(id);
        self
    }

    /// Validate and publish without changing the source snapshot.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation cannot be completed.
    pub fn commit(self) -> Result<Commit> {
        self.next.validate()?;
        let xml = codec::rewrite(self.base.xml_bytes(), &self.next)?;
        let snapshot = Snapshot::from_xml(xml)?;
        Ok(Commit {
            patch: Patch {
                before: self.base.value,
                after: self.next,
                before_fingerprint: self.base.fingerprint,
                after_fingerprint: snapshot.fingerprint,
            },
            snapshot,
        })
    }
}

/// A successful source-preserving publication.
#[derive(Debug, Clone)]
pub struct Commit {
    snapshot: Snapshot,
    patch: Patch,
}

impl Commit {
    #[must_use]
    pub const fn snapshot(&self) -> &Snapshot {
        &self.snapshot
    }

    #[must_use]
    pub fn into_snapshot(self) -> Snapshot {
        self.snapshot
    }

    #[must_use]
    pub const fn patch(&self) -> &Patch {
        &self.patch
    }

    #[must_use]
    pub fn into_patch(self) -> Patch {
        self.patch
    }
}

/// A reversible, source-checked OpenType patch.
#[derive(Debug, Clone)]
pub struct Patch {
    before: OpenType,
    after: OpenType,
    before_fingerprint: u64,
    after_fingerprint: u64,
}

impl Patch {
    #[must_use]
    pub const fn before(&self) -> &OpenType {
        &self.before
    }

    #[must_use]
    pub const fn after(&self) -> &OpenType {
        &self.after
    }

    #[must_use]
    pub fn inverse(&self) -> Self {
        Self {
            before: self.after.clone(),
            after: self.before.clone(),
            before_fingerprint: self.after_fingerprint,
            after_fingerprint: self.before_fingerprint,
        }
    }

    /// Apply only to the exact source snapshot captured by the transaction.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation cannot be completed.
    pub fn apply(&self, source: &Snapshot) -> Result<Snapshot> {
        if source.fingerprint != self.before_fingerprint || source.value != self.before {
            return Err(Error::InvalidFormat(
                "OpenType patch source does not match its precondition".into(),
            ));
        }
        let xml = codec::rewrite(source.xml_bytes(), &self.after)?;
        Snapshot::from_xml(xml)
    }
}

fn fingerprint(bytes: &[u8]) -> u64 {
    // Stable bounded FNV-1a source identity; the full source comparison is
    // still performed by `Patch::apply` through the snapshot fingerprint and
    // semantic precondition.
    let mut hash = 0xcbf2_9ce4_8422_2325_u64;
    for byte in bytes {
        hash ^= u64::from(*byte);
        hash = hash.wrapping_mul(0x0100_0000_01b3);
    }
    hash
}
