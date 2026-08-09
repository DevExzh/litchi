#![expect(
    clippy::shadow_reuse,
    reason = "parser bindings are intentionally refined after validation"
)]
//! Failure-atomic snapshots, edits, and reversible footnote-layout patches.

use std::sync::Arc;

use crate::error::{Error, Result};

use super::codec::{self, Context};
use super::model::Layout;
use super::validation;

/// An immutable, cheaply clonable `sectPr` snapshot.
#[derive(Debug, Clone)]
pub struct Snapshot {
    xml: Arc<[u8]>,
    value: Option<Layout>,
    context: Arc<Context>,
}

impl Snapshot {
    /// Parse and retain a bounded section-property snapshot.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation cannot be completed.
    pub fn from_xml(xml: impl Into<Vec<u8>>) -> Result<Self> {
        Self::from_xml_with_context(xml, Context::default())
    }

    /// Parse a detached section with namespace state inherited from its
    /// owning document part. The inherited state is not added to `xml_bytes`.
    pub(crate) fn from_xml_with_context(xml: impl Into<Vec<u8>>, context: Context) -> Result<Self> {
        let xml = xml.into();
        let value = codec::read_with_context(&xml, &context)?.value;
        Ok(Self {
            xml: Arc::from(xml.into_boxed_slice()),
            value,
            context: Arc::new(context),
        })
    }

    /// Return the authored section XML without copying it.
    #[must_use]
    pub fn xml_bytes(&self) -> &[u8] {
        &self.xml
    }

    /// Return the direct Word 2012 layout; absence remains observable.
    #[must_use]
    pub const fn layout(&self) -> Option<Layout> {
        self.value
    }

    /// Alias for callers reasoning in terms of the XML property name.
    #[must_use]
    pub const fn footnote_columns(&self) -> Option<Layout> {
        self.layout()
    }

    /// Start an isolated edit based on this snapshot.
    #[must_use]
    pub fn edit(&self) -> Transaction {
        Transaction {
            base: self.clone(),
            next: self.value,
        }
    }
}

/// A section-property edit that has not yet been published.
#[derive(Debug, Clone)]
pub struct Transaction {
    base: Snapshot,
    next: Option<Layout>,
}

impl Transaction {
    /// Return the projected layout in this transaction.
    #[must_use]
    pub const fn layout(&self) -> Option<Layout> {
        self.next
    }

    /// Set or remove the direct Word 2012 footnote layout.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation cannot be completed.
    pub fn set_layout(&mut self, value: Option<Layout>) -> Result<&mut Self> {
        validation::validate_layout(value)?;
        self.next = value;
        Ok(self)
    }

    /// Alias using the XML property vocabulary.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation cannot be completed.
    pub fn set_footnote_columns(&mut self, value: Option<Layout>) -> Result<&mut Self> {
        self.set_layout(value)
    }

    /// Remove the direct footnote layout marker.
    #[must_use]
    pub fn clear(&mut self) -> &mut Self {
        self.next = None;
        self
    }

    /// Validate and publish the edit without changing the source snapshot.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation cannot be completed.
    pub fn commit(self) -> Result<Commit> {
        if self.next == self.base.value {
            return Ok(Commit {
                patch: Patch::new(
                    self.base.value,
                    self.next,
                    self.base.xml.clone(),
                    self.base.xml.clone(),
                    self.base.context.clone(),
                    self.base.context.clone(),
                ),
                snapshot: self.base,
            });
        }
        let xml = codec::rewrite_with_context(
            self.base.xml_bytes(),
            self.next,
            self.base.context.as_ref(),
        )?;
        let snapshot = Snapshot::from_xml_with_context(xml, self.base.context.as_ref().clone())?;
        Ok(Commit {
            patch: Patch::new(
                self.base.value,
                self.next,
                self.base.xml.clone(),
                snapshot.xml.clone(),
                self.base.context.clone(),
                snapshot.context.clone(),
            ),
            snapshot,
        })
    }
}

/// A successful publication containing the new snapshot and reversible patch.
#[derive(Debug, Clone)]
pub struct Commit {
    snapshot: Snapshot,
    patch: Patch,
}

impl Commit {
    /// Borrow the published snapshot.
    #[must_use]
    pub const fn snapshot(&self) -> &Snapshot {
        &self.snapshot
    }

    /// Move the published snapshot out of the commit.
    #[must_use]
    pub fn into_snapshot(self) -> Snapshot {
        self.snapshot
    }

    /// Borrow the reversible patch.
    #[must_use]
    pub const fn patch(&self) -> &Patch {
        &self.patch
    }

    /// Move the reversible patch out of the commit.
    #[must_use]
    pub fn into_patch(self) -> Patch {
        self.patch
    }
}

/// A lineage-independent, preconditioned reversible layout patch.
#[derive(Debug, Clone, Eq, PartialEq)]
pub struct Patch {
    before: Option<Layout>,
    after: Option<Layout>,
    before_xml: Arc<[u8]>,
    after_xml: Arc<[u8]>,
    before_context: Arc<Context>,
    after_context: Arc<Context>,
}

impl Patch {
    fn new(
        before: Option<Layout>,
        after: Option<Layout>,
        before_xml: Arc<[u8]>,
        after_xml: Arc<[u8]>,
        before_context: Arc<Context>,
        after_context: Arc<Context>,
    ) -> Self {
        Self {
            before,
            after,
            before_xml,
            after_xml,
            before_context,
            after_context,
        }
    }

    /// Return the expected source state.
    #[must_use]
    pub const fn before(&self) -> Option<Layout> {
        self.before
    }

    /// Return the state produced by this patch.
    #[must_use]
    pub const fn after(&self) -> Option<Layout> {
        self.after
    }

    /// Return the inverse operation.
    #[must_use]
    pub fn inverse(&self) -> Self {
        Self {
            before: self.after,
            after: self.before,
            before_xml: self.after_xml.clone(),
            after_xml: self.before_xml.clone(),
            before_context: self.after_context.clone(),
            after_context: self.before_context.clone(),
        }
    }

    /// Apply the patch only when the target has the exact expected source
    /// bytes, inherited namespace context, and semantic state.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation cannot be completed.
    pub fn apply(&self, source: &Snapshot) -> Result<Snapshot> {
        if source.value != self.before
            || source.xml.as_ref() != self.before_xml.as_ref()
            || source.context.as_ref() != self.before_context.as_ref()
        {
            return Err(Error::InvalidFormat(
                "footnote-columns patch source snapshot does not match its precondition".into(),
            ));
        }
        if self.before == self.after {
            return Ok(source.clone());
        }
        Snapshot::from_xml_with_context(
            self.after_xml.to_vec(),
            self.after_context.as_ref().clone(),
        )
    }
}
