//! Immutable, source-bound transactions for packaged ODT documents.
//!
//! This is the safe package-level mutation boundary. It intentionally stages
//! only operations that keep the authoritative XML snapshot intact: callers
//! get exact no-op bytes, source-checked reversible patches, and a complete
//! compact-XML audit before publication. Broader structural `MutableDocument`
//! operations remain available separately while their opaque-content
//! preservation contracts are migrated to this transaction surface.

use crate::{Document, mutable::MutableDocument};
use litchi_core::{Error, Result};
use std::sync::Arc;

const MAX_PACKAGE_BYTES: usize = 64 * 1024 * 1024;
const MAX_OPERATIONS: usize = 1_024;

/// Immutable, validated ODT package snapshot.
#[derive(Clone)]
pub struct Snapshot {
    bytes: Arc<Vec<u8>>,
}

impl Snapshot {
    /// Opens and retains an ODT package as an immutable snapshot.
    pub fn from_bytes(bytes: Vec<u8>) -> Result<Self> {
        ensure_package_size(bytes.len(), "ODT transaction input")?;
        // Validate a bounded copy so the retained snapshot is the exact
        // caller-provided package, not a writer-normalized representation.
        Document::from_bytes(copy_bytes(&bytes)?)?;
        Ok(Self {
            bytes: Arc::new(bytes),
        })
    }

    /// Captures the exact bytes backing an already validated document.
    pub fn from_document(document: &Document) -> Result<Self> {
        let bytes = copy_bytes(document.original_bytes())?;
        Self::from_bytes(bytes)
    }

    /// Returns the exact package bytes represented by this snapshot.
    #[must_use]
    pub fn as_bytes(&self) -> &[u8] {
        self.bytes.as_slice()
    }

    /// Reopens this immutable snapshot for semantic inspection.
    pub fn document(&self) -> Result<Document> {
        Document::from_bytes(copy_bytes(self.as_bytes())?)
    }

    /// Starts a detached, failure-atomic package edit.
    #[must_use]
    pub fn edit(&self) -> Edit {
        Edit {
            source: self.clone(),
            operations: Vec::new(),
        }
    }
}

/// Selector used by packaged-document edit operations.
#[derive(Clone, Debug, PartialEq, Eq)]
#[non_exhaustive]
pub enum ParagraphSelector {
    /// The zero-based paragraph position in semantic document order.
    Index(usize),
    /// The unique paragraph whose extracted semantic text equals this value.
    ExactText(String),
}

impl ParagraphSelector {
    /// Select a paragraph by semantic document order.
    #[must_use]
    pub const fn at(index: usize) -> Self {
        Self::Index(index)
    }

    /// Select the one paragraph matching `text` exactly.
    #[must_use]
    pub fn exact_text(text: impl Into<String>) -> Self {
        Self::ExactText(text.into())
    }
}

/// Detached, lossless packaged-ODT edit.
pub struct Edit {
    source: Snapshot,
    operations: Vec<Operation>,
}

impl Edit {
    /// Stages a line break in the selected paragraph without rebuilding its
    /// inline markup. Scripts, forms, DDE declarations, and other active
    /// content are retained as inert package data and are never executed.
    pub fn append_line_break(&mut self, selector: ParagraphSelector) -> Result<&mut Self> {
        if self.operations.len() >= MAX_OPERATIONS {
            return Err(Error::InvalidFormat(format!(
                "ODT transaction exceeds {MAX_OPERATIONS} staged operations"
            )));
        }
        let index = resolve_paragraph(&self.source.document()?, &selector)?;
        self.operations
            .try_reserve(1)
            .map_err(|source| Error::Allocation {
                resource: "ODT transaction operations",
                source,
            })?;
        self.operations.push(Operation::AppendLineBreak { index });
        Ok(self)
    }

    /// Validates every staged operation and publishes one immutable snapshot.
    pub fn commit(self) -> Result<Commit> {
        if self.operations.is_empty() {
            return Ok(Commit::new(self.source.clone(), self.source));
        }

        let document = self.source.document()?;
        let mut mutable = MutableDocument::from_document(document)?;
        for operation in &self.operations {
            match operation {
                Operation::AppendLineBreak { index } => mutable.append_line_break(*index)?,
            }
        }
        let bytes = mutable.to_bytes()?;
        ensure_package_size(bytes.len(), "ODT transaction output")?;
        audit_compact_xml(&bytes)?;
        let after = Snapshot::from_bytes(bytes)?;
        // Reparse after the compact audit so commit never returns an invalid
        // candidate even when a writer implementation changes independently.
        after.document()?;
        Ok(Commit::new(self.source, after))
    }
}

#[derive(Clone)]
enum Operation {
    AppendLineBreak { index: usize },
}

/// A validated packaged-ODT transaction result.
pub struct Commit {
    snapshot: Snapshot,
    patch: Patch,
}

impl Commit {
    fn new(before: Snapshot, snapshot: Snapshot) -> Self {
        Self {
            patch: Patch {
                before,
                after: snapshot.clone(),
            },
            snapshot,
        }
    }

    /// Returns the committed immutable snapshot.
    #[must_use]
    pub const fn snapshot(&self) -> &Snapshot {
        &self.snapshot
    }

    /// Returns the exact source-checked reversible patch.
    #[must_use]
    pub const fn patch(&self) -> &Patch {
        &self.patch
    }

    /// Consumes the commit and returns its immutable snapshot.
    #[must_use]
    pub fn into_snapshot(self) -> Snapshot {
        self.snapshot
    }
}

/// Exact-byte, source-checked package patch.
#[derive(Clone)]
pub struct Patch {
    before: Snapshot,
    after: Snapshot,
}

impl Patch {
    /// Applies this patch only to the exact snapshot from which it was made.
    pub fn apply(&self, source: &Snapshot) -> Result<Snapshot> {
        if source.as_bytes() != self.before.as_bytes() {
            return Err(Error::InvalidFormat(
                "ODT patch source does not match its expected snapshot".to_string(),
            ));
        }
        Ok(self.after.clone())
    }

    /// Returns the patch that restores the exact source bytes.
    #[must_use]
    pub fn inverse(&self) -> Self {
        Self {
            before: self.after.clone(),
            after: self.before.clone(),
        }
    }
}

fn copy_bytes(source: &[u8]) -> Result<Vec<u8>> {
    ensure_package_size(source.len(), "ODT transaction package")?;
    let mut bytes = Vec::new();
    bytes
        .try_reserve_exact(source.len())
        .map_err(|source| Error::Allocation {
            resource: "ODT transaction package",
            source,
        })?;
    bytes.extend_from_slice(source);
    Ok(bytes)
}

fn ensure_package_size(size: usize, scope: &str) -> Result<()> {
    if size > MAX_PACKAGE_BYTES {
        return Err(Error::InvalidFormat(format!(
            "{scope} exceeds the {MAX_PACKAGE_BYTES}-byte package limit"
        )));
    }
    Ok(())
}

fn resolve_paragraph(document: &Document, selector: &ParagraphSelector) -> Result<usize> {
    let paragraphs = document.paragraphs()?;
    match selector {
        ParagraphSelector::Index(index) if *index < paragraphs.len() => Ok(*index),
        ParagraphSelector::Index(index) => Err(Error::InvalidFormat(format!(
            "paragraph selector index {index} is out of bounds (found {})",
            paragraphs.len()
        ))),
        ParagraphSelector::ExactText(text) => {
            let mut selected = None;
            for (index, paragraph) in paragraphs.into_iter().enumerate() {
                if paragraph.text()? == *text {
                    if selected.replace(index).is_some() {
                        return Err(Error::InvalidFormat(
                            "paragraph text selector is ambiguous".to_string(),
                        ));
                    }
                }
            }
            selected.ok_or_else(|| {
                Error::InvalidFormat("paragraph text selector did not match".to_string())
            })
        },
    }
}

fn audit_compact_xml(bytes: &[u8]) -> Result<()> {
    let package = crate::core::OwnedPackage::from_bytes(copy_bytes(bytes)?)?;
    for path in package.files()? {
        if path.ends_with(".xml") {
            let xml = package.get_file(&path)?;
            let limits = litchi_odf_common::compact_xml::Limits::new(MAX_PACKAGE_BYTES, 4_096)
                .map_err(Error::from)?;
            litchi_odf_common::compact_xml::validate_with_limits(&xml, limits)
                .map_err(Error::from)?;
        }
    }
    Ok(())
}
