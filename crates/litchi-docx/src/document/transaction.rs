//! Source-preserving main-document snapshots, edits, and reversible patches.

mod durable;

use std::sync::Arc;

use litchi_core::Position;
use litchi_core::xml::escape_xml;
use quick_xml::events::Event;
use quick_xml::name::{Namespace, ResolveResult};
use quick_xml::reader::NsReader;
use quick_xml::{Reader, XmlVersion};
use thiserror::Error;

use crate::namespace::{
    STRICT_WORDPROCESSINGML_NAMESPACE, WORDPROCESSINGML_NAMESPACE, is_wordprocessing_namespace,
};
use crate::paragraph::Paragraph;

pub use durable::{Composition, History, JoinError, PreparedEdit, ThreeWayError, ThreeWayPlan};
pub use litchi_core::patch::{
    CompositionLimits, HistoryLimits, MergeChoice, SubEditConflict, SubEditJoinFailure,
    ThreeWayMergeFailure,
};

const MAX_DOCUMENT_XML_BYTES: usize = 32 * 1024 * 1024;
const MAX_DOCUMENT_DEPTH: usize = 256;
const MAX_DOCUMENT_NODES: usize = 1_000_000;
const MAX_OPERATIONS: usize = 4_096;
const MAX_REPLACEMENT_TEXT_BYTES: usize = 16 * 1024 * 1024;

/// Result returned by main-document transaction operations.
pub type TransactionResult<T> = Result<T, TransactionError>;

/// A typed reason why a paragraph operation cannot be represented safely.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[non_exhaustive]
pub enum Refusal {
    /// The selected owner contains unsupported structural content.
    ComplexContent,
    /// The selected owner has no editable text element.
    ComplexRun,
    /// The selected hyperlink does not exist.
    HyperlinkNotFound,
    /// The selected direct paragraph run does not exist.
    RunNotFound,
    /// The selected simple field does not exist.
    FieldNotFound,
    /// The selected tracked insertion or deletion does not exist.
    RevisionNotFound,
    /// The selected direct inline content control does not exist.
    ContentControlNotFound,
    /// The selected table, row, or cell does not exist.
    CellNotFound,
    /// The requested text needs structural run elements such as `w:tab` or
    /// `w:br`, which this focused text operation does not synthesize.
    StructuralText,
}

impl std::fmt::Display for Refusal {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        formatter.write_str(match self {
            Self::ComplexContent => "selected owner contains unsupported structural content",
            Self::ComplexRun => "selected content has no editable Word text",
            Self::HyperlinkNotFound => "direct paragraph hyperlink was not found",
            Self::RunNotFound => "direct paragraph run was not found",
            Self::FieldNotFound => "direct simple field was not found",
            Self::RevisionNotFound => "direct tracked revision was not found",
            Self::ContentControlNotFound => "direct inline content control was not found",
            Self::CellNotFound => "table cell was not found",
            Self::StructuralText => "text requires structural WordprocessingML elements",
        })
    }
}

/// Direct tracked-revision wrapper selected for inert text replacement.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[non_exhaustive]
pub enum RevisionKind {
    /// `w:ins` tracked insertion content.
    Insertion,
    /// `w:del` tracked deletion content.
    Deletion,
}

/// Dependency-checked paragraph payload prepared for one exact receiving
/// document. Package planning rewrites relationship references to dependencies
/// already proven present in the receiver before constructing this value.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ParagraphTransfer {
    target: Arc<Vec<u8>>,
    fragment: Arc<Vec<u8>>,
    dependency_digest: Arc<str>,
}

impl ParagraphTransfer {
    pub(crate) fn new(target: Arc<Vec<u8>>, fragment: Vec<u8>, dependency_digest: String) -> Self {
        Self {
            target,
            fragment: Arc::new(fragment),
            dependency_digest: dependency_digest.into(),
        }
    }

    /// Exact compact paragraph XML retained by the plan.
    #[must_use]
    pub fn xml_bytes(&self) -> &[u8] {
        self.fragment.as_slice()
    }
}

impl RevisionKind {
    const fn local_name(self) -> &'static [u8] {
        match self {
            Self::Insertion => b"ins",
            Self::Deletion => b"del",
        }
    }
}

/// Typed reason why a cross-package paragraph dependency closure cannot be
/// represented without copying or guessing package resources.
#[derive(Debug, Clone, PartialEq, Eq)]
#[non_exhaustive]
pub enum TransferRefusal {
    /// A relationship reference in the donor paragraph is dangling.
    MissingDonorRelationship(String),
    /// The receiving package has no semantically equivalent dependency edge.
    MissingEquivalentDependency {
        /// OPC relationship type URI.
        relationship_type: String,
        /// Exact external or relative internal target reference.
        target: String,
    },
    /// Multiple receiver edges are equivalent, so choosing an ID would guess.
    AmbiguousEquivalentDependency {
        /// OPC relationship type URI.
        relationship_type: String,
        /// Exact external or relative internal target reference.
        target: String,
    },
    /// The selected donor paragraph could not be represented as compact XML.
    InvalidParagraphXml,
}

impl std::fmt::Display for TransferRefusal {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::MissingDonorRelationship(identifier) => {
                write!(formatter, "donor relationship {identifier} is missing")
            },
            Self::MissingEquivalentDependency {
                relationship_type,
                target,
            } => write!(
                formatter,
                "receiver lacks dependency {relationship_type} -> {target}"
            ),
            Self::AmbiguousEquivalentDependency {
                relationship_type,
                target,
            } => write!(
                formatter,
                "receiver dependency {relationship_type} -> {target} is ambiguous"
            ),
            Self::InvalidParagraphXml => {
                formatter.write_str("donor paragraph XML is not transferable")
            },
        }
    }
}

/// A main-document transaction failure.
#[derive(Debug, Error)]
#[non_exhaustive]
pub enum TransactionError {
    /// The underlying DOCX document or package is invalid.
    #[error(transparent)]
    Document(#[from] crate::Error),
    /// A checked paragraph position is outside the projected document.
    #[error("paragraph position {position} is out of bounds for length {len}")]
    OutOfBounds {
        /// Requested zero-based paragraph position.
        position: usize,
        /// Projected direct-body paragraph count.
        len: usize,
    },
    /// The selected paragraph cannot be changed without guessing how to
    /// rewrite dependent or structured content.
    #[error("paragraph {position} edit refused: {reason}")]
    Refused {
        /// Selected zero-based paragraph position.
        position: usize,
        /// Stable refusal category.
        reason: Refusal,
    },
    /// A configured transaction resource ceiling was exceeded.
    #[error("document transaction {resource} limit exceeded: {actual} > {max}")]
    Limit {
        /// Bounded resource.
        resource: &'static str,
        /// Maximum accepted value.
        max: usize,
        /// Observed or requested value.
        actual: usize,
    },
    /// The patch target no longer has the exact source bytes captured by the
    /// edit.
    #[error("document patch source is stale")]
    StaleSource,
    /// A semantic durable operation's expected value did not match source.
    #[error("document patch semantic precondition does not match")]
    SemanticPrecondition,
    /// A common durable patch could not be constructed or decoded.
    #[error(transparent)]
    Durable(#[from] litchi_core::patch::PatchError),
    /// A common disjoint-composition bound or identifier was invalid.
    #[error(transparent)]
    Composition(#[from] litchi_core::patch::CompositionError),
    /// A durable patch used an unsupported or malformed DOCX vocabulary.
    #[error("invalid DOCX durable patch: {0}")]
    InvalidDurable(String),
    /// A cross-package transfer dependency could not be closed safely.
    #[error("paragraph transfer refused: {0}")]
    Transfer(TransferRefusal),
}

/// An immutable, cheaply clonable snapshot of the main document XML.
#[derive(Debug, Clone)]
pub struct Snapshot {
    xml: Arc<Vec<u8>>,
    paragraphs: Arc<[Range]>,
    tables: Arc<[Range]>,
    content_end: u32,
    conformance: Conformance,
}

impl Snapshot {
    /// Parse and retain one bounded `WordprocessingML` main document.
    ///
    /// # Errors
    ///
    /// Returns a typed document or resource-limit error when the XML is
    /// malformed, unsupported, or exceeds the transaction bounds.
    pub fn from_xml(source_xml: impl Into<Vec<u8>>) -> TransactionResult<Self> {
        let xml = source_xml.into();
        if xml.len() > MAX_DOCUMENT_XML_BYTES {
            return Err(TransactionError::Limit {
                resource: "XML bytes",
                max: MAX_DOCUMENT_XML_BYTES,
                actual: xml.len(),
            });
        }
        let layout = scan_document(&xml)?;
        Ok(Self {
            xml: Arc::new(xml),
            paragraphs: layout.paragraphs.into(),
            tables: layout.tables.into(),
            content_end: layout.content_end,
            conformance: layout.conformance,
        })
    }

    /// Borrow the exact main-document XML bytes.
    #[must_use]
    pub fn xml_bytes(&self) -> &[u8] {
        self.xml.as_slice()
    }

    /// Return the number of direct main-body paragraphs.
    #[must_use]
    pub fn paragraph_count(&self) -> usize {
        self.paragraphs.len()
    }

    /// Borrow one direct main-body paragraph through a checked position.
    #[must_use]
    pub fn paragraph(&self, position: Position) -> Option<Paragraph> {
        self.paragraphs.get(position.get()).map(|range| {
            Paragraph::from_arc_range(Arc::clone(&self.xml), range.start, range.length)
        })
    }

    /// Return all direct main-body paragraphs without copying their XML.
    #[must_use]
    pub fn paragraphs(&self) -> Vec<Paragraph> {
        self.paragraphs
            .iter()
            .map(|range| {
                Paragraph::from_arc_range(Arc::clone(&self.xml), range.start, range.length)
            })
            .collect()
    }

    /// Return the number of direct main-body tables.
    #[must_use]
    pub fn table_count(&self) -> usize {
        self.tables.len()
    }

    /// Start an isolated edit whose selectors resolve against its projected
    /// state.
    #[must_use]
    pub fn edit(&self) -> Edit {
        Edit {
            base: self.clone(),
            projected: self.clone(),
            operations: Vec::new(),
            replacement_text_bytes: 0,
        }
    }

    fn same_source(&self, other: &Self) -> bool {
        self.xml.as_slice() == other.xml.as_slice()
    }
}

/// A semantic main-document operation recorded in a reversible patch.
#[derive(Debug, Clone, PartialEq, Eq)]
#[non_exhaustive]
pub enum Operation {
    /// Replace the complete text of one direct-body paragraph.
    ReplaceParagraphText {
        /// Projected paragraph position at the time of the operation.
        position: Position,
        /// Text required before applying the operation.
        before: String,
        /// Text produced by the operation.
        after: String,
    },
    /// Replace text in one direct hyperlink while retaining its relationships.
    ReplaceHyperlinkText {
        /// Direct-body paragraph position.
        paragraph: Position,
        /// Direct hyperlink position within that paragraph.
        hyperlink: Position,
        /// Text required before applying the operation.
        before: String,
        /// Text produced by the operation.
        after: String,
    },
    /// Replace text in one direct run inside an otherwise rich paragraph.
    ReplaceRunText {
        /// Direct-body paragraph position.
        paragraph: Position,
        /// Direct run position within the paragraph.
        run: Position,
        /// Text required before applying the operation.
        before: String,
        /// Text produced by the operation.
        after: String,
    },
    /// Replace the displayed result of one direct `w:fldSimple` field.
    ReplaceSimpleFieldText {
        /// Direct-body paragraph position.
        paragraph: Position,
        /// Direct simple-field position within the paragraph.
        field: Position,
        /// Text required before applying the operation.
        before: String,
        /// Text produced by the operation.
        after: String,
    },
    /// Replace inert text inside one direct tracked revision wrapper.
    ReplaceRevisionText {
        /// Direct-body paragraph position.
        paragraph: Position,
        /// Tracked wrapper family.
        kind: RevisionKind,
        /// Position among direct wrappers of the selected family.
        revision: Position,
        /// Text required before applying the operation.
        before: String,
        /// Text produced by the operation.
        after: String,
    },
    /// Replace text in one direct inline content control.
    ReplaceContentControlText {
        /// Direct-body paragraph position.
        paragraph: Position,
        /// Direct inline content-control position within the paragraph.
        control: Position,
        /// Text required before applying the operation.
        before: String,
        /// Text produced by the operation.
        after: String,
    },
    /// Replace text in one basic direct-body table cell.
    ReplaceCellText {
        /// Direct-body table position.
        table: Position,
        /// Direct row position in the table.
        row: Position,
        /// Direct cell position in the row.
        cell: Position,
        /// Text required before applying the operation.
        before: String,
        /// Text produced by the operation.
        after: String,
    },
    /// Replace one direct paragraph inside a rich or multi-paragraph cell.
    ReplaceCellParagraphText {
        /// Direct-body table position.
        table: Position,
        /// Direct row position in the table.
        row: Position,
        /// Direct cell position in the row.
        cell: Position,
        /// Direct paragraph position in the cell.
        paragraph: Position,
        /// Text required before applying the operation.
        before: String,
        /// Text produced by the operation.
        after: String,
    },
    /// Insert one compact plain-text paragraph.
    InsertParagraph {
        /// Projected insertion position at the time of the operation.
        position: Position,
        /// Inserted inert text.
        text: String,
    },
    /// Remove a paragraph previously inserted by the inverse patch.
    RemoveParagraph {
        /// Projected paragraph position at the time of the operation.
        position: Position,
        /// Removed inert text.
        text: String,
    },
    /// Insert one dependency-checked compact paragraph fragment.
    InsertTransferredParagraph {
        /// Projected insertion position.
        position: Position,
        /// Compact paragraph XML with receiver-local relationship references.
        xml: Arc<Vec<u8>>,
        /// Exact receiver relationship/resource inventory required at publish.
        dependency_digest: Arc<str>,
    },
    /// Remove the exact transferred paragraph fragment.
    RemoveTransferredParagraph {
        /// Projected paragraph position.
        position: Position,
        /// Exact compact paragraph XML expected at the position.
        xml: Arc<Vec<u8>>,
        /// Exact receiver relationship/resource inventory required at publish.
        dependency_digest: Arc<str>,
    },
}

impl Operation {
    fn inverse(&self) -> Self {
        match self {
            Self::ReplaceParagraphText {
                position,
                before,
                after,
            } => Self::ReplaceParagraphText {
                position: *position,
                before: after.clone(),
                after: before.clone(),
            },
            Self::ReplaceHyperlinkText {
                paragraph,
                hyperlink,
                before,
                after,
            } => Self::ReplaceHyperlinkText {
                paragraph: *paragraph,
                hyperlink: *hyperlink,
                before: after.clone(),
                after: before.clone(),
            },
            Self::ReplaceRunText {
                paragraph,
                run,
                before,
                after,
            } => Self::ReplaceRunText {
                paragraph: *paragraph,
                run: *run,
                before: after.clone(),
                after: before.clone(),
            },
            Self::ReplaceSimpleFieldText {
                paragraph,
                field,
                before,
                after,
            } => Self::ReplaceSimpleFieldText {
                paragraph: *paragraph,
                field: *field,
                before: after.clone(),
                after: before.clone(),
            },
            Self::ReplaceRevisionText {
                paragraph,
                kind,
                revision,
                before,
                after,
            } => Self::ReplaceRevisionText {
                paragraph: *paragraph,
                kind: *kind,
                revision: *revision,
                before: after.clone(),
                after: before.clone(),
            },
            Self::ReplaceContentControlText {
                paragraph,
                control,
                before,
                after,
            } => Self::ReplaceContentControlText {
                paragraph: *paragraph,
                control: *control,
                before: after.clone(),
                after: before.clone(),
            },
            Self::ReplaceCellText {
                table,
                row,
                cell,
                before,
                after,
            } => Self::ReplaceCellText {
                table: *table,
                row: *row,
                cell: *cell,
                before: after.clone(),
                after: before.clone(),
            },
            Self::ReplaceCellParagraphText {
                table,
                row,
                cell,
                paragraph,
                before,
                after,
            } => Self::ReplaceCellParagraphText {
                table: *table,
                row: *row,
                cell: *cell,
                paragraph: *paragraph,
                before: after.clone(),
                after: before.clone(),
            },
            Self::InsertParagraph { position, text } => Self::RemoveParagraph {
                position: *position,
                text: text.clone(),
            },
            Self::RemoveParagraph { position, text } => Self::InsertParagraph {
                position: *position,
                text: text.clone(),
            },
            Self::InsertTransferredParagraph {
                position,
                xml,
                dependency_digest,
            } => Self::RemoveTransferredParagraph {
                position: *position,
                xml: Arc::clone(xml),
                dependency_digest: Arc::clone(dependency_digest),
            },
            Self::RemoveTransferredParagraph {
                position,
                xml,
                dependency_digest,
            } => Self::InsertTransferredParagraph {
                position: *position,
                xml: Arc::clone(xml),
                dependency_digest: Arc::clone(dependency_digest),
            },
        }
    }
}

/// A staged main-document edit.
#[derive(Debug, Clone)]
pub struct Edit {
    base: Snapshot,
    projected: Snapshot,
    operations: Vec<Operation>,
    replacement_text_bytes: usize,
}

impl Edit {
    /// Borrow the immutable source snapshot.
    #[must_use]
    pub const fn source(&self) -> &Snapshot {
        &self.base
    }

    /// Borrow the current projected snapshot.
    #[must_use]
    pub const fn projected(&self) -> &Snapshot {
        &self.projected
    }

    /// Replace all text in a direct-body paragraph while retaining run
    /// boundaries, formatting, drawings, and unknown run XML.
    ///
    /// Replacement characters are assigned to the existing text slots in
    /// order: each slot keeps up to its original character count and the final
    /// slot receives any remainder. Direct hyperlinks and other paragraph
    /// wrappers use their focused operations and are refused here.
    ///
    /// # Errors
    ///
    /// Returns a typed refusal, checked-position error, resource-limit error,
    /// or malformed-document error without changing the projected snapshot.
    pub fn replace_paragraph_text(
        &mut self,
        position: Position,
        authored_text: impl Into<String>,
    ) -> TransactionResult<&mut Self> {
        self.reserve_operation()?;
        let text = authored_text.into();
        validate_authored_text(&text).map_err(|reason| TransactionError::Refused {
            position: position.get(),
            reason,
        })?;
        let replacement_text_bytes = self.checked_text_total(text.len())?;
        let range = self.range(position)?;
        let paragraph_start = usize::try_from(range.start).map_err(|_conversion_error| {
            crate::Error::InvalidFormat("paragraph offset does not fit usize".into())
        })?;
        let paragraph_end = paragraph_start
            .checked_add(usize::try_from(range.length).map_err(|_conversion_error| {
                crate::Error::InvalidFormat("paragraph length does not fit usize".into())
            })?)
            .ok_or_else(|| crate::Error::InvalidFormat("paragraph range overflow".into()))?;
        let paragraph = self
            .projected
            .xml_bytes()
            .get(paragraph_start..paragraph_end)
            .ok_or_else(|| crate::Error::InvalidFormat("paragraph range is outside XML".into()))?;
        let owner =
            scan_text_owner(paragraph, b"p").map_err(|reason| TransactionError::Refused {
                position: position.get(),
                reason,
            })?;
        if owner.text == text {
            return Ok(self);
        }
        let replacement = rewrite_text_owner(paragraph, &owner, &text)?;
        let xml = replace_range(
            self.projected.xml_bytes(),
            paragraph_start,
            paragraph_end,
            &replacement,
        )?;
        let candidate = Snapshot::from_xml(xml)?;
        let readback = candidate
            .paragraph(position)
            .ok_or(TransactionError::OutOfBounds {
                position: position.get(),
                len: candidate.paragraph_count(),
            })?
            .text()?;
        if readback != text {
            return Err(crate::Error::InvalidFormat(
                "document text edit failed semantic readback".into(),
            )
            .into());
        }
        self.operations.push(Operation::ReplaceParagraphText {
            position,
            before: owner.text,
            after: text,
        });
        self.replacement_text_bytes = replacement_text_bytes;
        self.projected = candidate;
        Ok(self)
    }

    /// Replace all text in one direct paragraph hyperlink while leaving its
    /// anchor, tooltip, relationship, target frame, and unknown XML untouched.
    ///
    /// # Errors
    ///
    /// Returns a checked selector/refusal, resource-limit, or malformed XML
    /// error without changing the projected snapshot.
    pub fn replace_hyperlink_text(
        &mut self,
        paragraph: Position,
        hyperlink: Position,
        authored_text: impl Into<String>,
    ) -> TransactionResult<&mut Self> {
        self.reserve_operation()?;
        let text = authored_text.into();
        validate_authored_text(&text).map_err(|reason| TransactionError::Refused {
            position: paragraph.get(),
            reason,
        })?;
        let replacement_text_bytes = self.checked_text_total(text.len())?;
        let paragraph_range = self.range(paragraph)?;
        let paragraph_start = checked_start(paragraph_range, "paragraph")?;
        let paragraph_end = checked_end(paragraph_range, "paragraph")?;
        let paragraph_xml = checked_slice(
            self.projected.xml_bytes(),
            paragraph_start,
            paragraph_end,
            "paragraph",
        )?;
        let hyperlink_range = select_direct_child(
            paragraph_xml,
            b"p",
            b"hyperlink",
            hyperlink,
            Refusal::HyperlinkNotFound,
        )
        .map_err(|reason| TransactionError::Refused {
            position: paragraph.get(),
            reason,
        })?;
        let hyperlink_start = checked_relative_start(paragraph_start, hyperlink_range)?;
        let hyperlink_end = checked_relative_end(paragraph_start, hyperlink_range)?;
        let hyperlink_xml = checked_slice(
            self.projected.xml_bytes(),
            hyperlink_start,
            hyperlink_end,
            "hyperlink",
        )?;
        let owner = scan_text_owner(hyperlink_xml, b"hyperlink").map_err(|reason| {
            TransactionError::Refused {
                position: paragraph.get(),
                reason,
            }
        })?;
        if owner.text == text {
            return Ok(self);
        }
        let replacement = rewrite_text_owner(hyperlink_xml, &owner, &text)?;
        let xml = replace_range(
            self.projected.xml_bytes(),
            hyperlink_start,
            hyperlink_end,
            &replacement,
        )?;
        let candidate = Snapshot::from_xml(xml)?;
        let actual = selected_hyperlink_text(&candidate, paragraph, hyperlink)?;
        if actual != text {
            return Err(crate::Error::InvalidFormat(
                "document hyperlink edit failed semantic readback".into(),
            )
            .into());
        }
        self.operations.push(Operation::ReplaceHyperlinkText {
            paragraph,
            hyperlink,
            before: owner.text,
            after: text,
        });
        self.replacement_text_bytes = replacement_text_bytes;
        self.projected = candidate;
        Ok(self)
    }

    /// Replace text in one direct run while retaining its `w:rPr`, drawings,
    /// and opaque run children. This provides a checked edit seam inside a
    /// paragraph that also contains fields, hyperlinks, or revision wrappers.
    ///
    /// # Errors
    ///
    /// Returns a checked selector/refusal, resource-limit, or malformed XML
    /// error without changing the projected snapshot.
    pub fn replace_run_text(
        &mut self,
        paragraph: Position,
        run: Position,
        authored_text: impl Into<String>,
    ) -> TransactionResult<&mut Self> {
        self.replace_direct_paragraph_owner_text(
            paragraph,
            run,
            b"r",
            Refusal::RunNotFound,
            authored_text.into(),
            |before, after| Operation::ReplaceRunText {
                paragraph,
                run,
                before,
                after,
            },
        )
    }

    /// Replace the displayed text of one direct simple field while preserving
    /// its instruction, dirty/lock state, formatting, and opaque XML.
    ///
    /// Field instructions remain inert and are never evaluated or refreshed.
    ///
    /// # Errors
    ///
    /// Returns a checked selector/refusal, resource-limit, or malformed XML
    /// error without changing the projected snapshot.
    pub fn replace_simple_field_text(
        &mut self,
        paragraph: Position,
        field: Position,
        authored_text: impl Into<String>,
    ) -> TransactionResult<&mut Self> {
        self.replace_direct_paragraph_owner_text(
            paragraph,
            field,
            b"fldSimple",
            Refusal::FieldNotFound,
            authored_text.into(),
            |before, after| Operation::ReplaceSimpleFieldText {
                paragraph,
                field,
                before,
                after,
            },
        )
    }

    /// Replace inert text inside one direct tracked insertion or deletion.
    /// Revision metadata and the wrapper itself remain byte-preserved.
    ///
    /// # Errors
    ///
    /// Returns a checked selector/refusal, resource-limit, or malformed XML
    /// error without changing the projected snapshot.
    pub fn replace_revision_text(
        &mut self,
        paragraph: Position,
        kind: RevisionKind,
        revision: Position,
        authored_text: impl Into<String>,
    ) -> TransactionResult<&mut Self> {
        self.replace_direct_paragraph_owner_text(
            paragraph,
            revision,
            kind.local_name(),
            Refusal::RevisionNotFound,
            authored_text.into(),
            |before, after| Operation::ReplaceRevisionText {
                paragraph,
                kind,
                revision,
                before,
                after,
            },
        )
    }

    /// Replace text inside one direct inline content control while retaining
    /// `w:sdtPr`, data binding, lock state, wrapper metadata, and opaque XML.
    ///
    /// This operation supports an inline `w:sdtContent` whose direct children
    /// are runs; block controls and nested control structures are refused.
    ///
    /// # Errors
    ///
    /// Returns a checked selector/refusal, resource-limit, or malformed XML
    /// error without changing the projected snapshot.
    pub fn replace_content_control_text(
        &mut self,
        paragraph: Position,
        control: Position,
        authored_text: impl Into<String>,
    ) -> TransactionResult<&mut Self> {
        self.reserve_operation()?;
        let text = authored_text.into();
        validate_authored_text(&text).map_err(|reason| TransactionError::Refused {
            position: paragraph.get(),
            reason,
        })?;
        let replacement_text_bytes = self.checked_text_total(text.len())?;
        let content = select_content_control_content(&self.projected, paragraph, control)?;
        let content_xml = checked_slice(
            self.projected.xml_bytes(),
            content.0,
            content.1,
            "content control content",
        )?;
        let owner = scan_text_owner(content_xml, b"sdtContent").map_err(|reason| {
            TransactionError::Refused {
                position: paragraph.get(),
                reason,
            }
        })?;
        if owner.text == text {
            return Ok(self);
        }
        let replacement = rewrite_text_owner(content_xml, &owner, &text)?;
        let candidate = Snapshot::from_xml(replace_range(
            self.projected.xml_bytes(),
            content.0,
            content.1,
            &replacement,
        )?)?;
        let actual = selected_content_control_text(&candidate, paragraph, control)?;
        if actual != text {
            return Err(crate::Error::InvalidFormat(
                "document content-control edit failed semantic readback".into(),
            )
            .into());
        }
        self.operations.push(Operation::ReplaceContentControlText {
            paragraph,
            control,
            before: owner.text,
            after: text,
        });
        self.replacement_text_bytes = replacement_text_bytes;
        self.projected = candidate;
        Ok(self)
    }

    /// Replace text in a basic direct-body table cell.
    ///
    /// The supported cell contains one direct paragraph. Its existing runs,
    /// formatting, cell properties, drawings, and unknown run XML remain in
    /// place; nested tables, controls, and multiple cell paragraphs are
    /// refused.
    ///
    /// # Errors
    ///
    /// Returns a checked selector/refusal, resource-limit, or malformed XML
    /// error without changing the projected snapshot.
    pub fn replace_table_cell_text(
        &mut self,
        table: Position,
        row: Position,
        cell: Position,
        authored_text: impl Into<String>,
    ) -> TransactionResult<&mut Self> {
        self.reserve_operation()?;
        let text = authored_text.into();
        validate_authored_text(&text).map_err(|reason| TransactionError::Refused {
            position: table.get(),
            reason,
        })?;
        let replacement_text_bytes = self.checked_text_total(text.len())?;
        let cell_selection = select_cell(&self.projected, table, row, cell)?;
        let paragraph_range = single_cell_paragraph(cell_selection.xml)?;
        let paragraph_start = checked_relative_start(cell_selection.start, paragraph_range)?;
        let paragraph_end = checked_relative_end(cell_selection.start, paragraph_range)?;
        let paragraph_xml = checked_slice(
            self.projected.xml_bytes(),
            paragraph_start,
            paragraph_end,
            "table cell paragraph",
        )?;
        let owner =
            scan_text_owner(paragraph_xml, b"p").map_err(|reason| TransactionError::Refused {
                position: table.get(),
                reason,
            })?;
        if owner.text == text {
            return Ok(self);
        }
        let replacement = rewrite_text_owner(paragraph_xml, &owner, &text)?;
        let xml = replace_range(
            self.projected.xml_bytes(),
            paragraph_start,
            paragraph_end,
            &replacement,
        )?;
        let candidate = Snapshot::from_xml(xml)?;
        let actual = selected_cell_text(&candidate, table, row, cell)?;
        if actual != text {
            return Err(crate::Error::InvalidFormat(
                "document table-cell edit failed semantic readback".into(),
            )
            .into());
        }
        self.operations.push(Operation::ReplaceCellText {
            table,
            row,
            cell,
            before: owner.text,
            after: text,
        });
        self.replacement_text_bytes = replacement_text_bytes;
        self.projected = candidate;
        Ok(self)
    }

    /// Replace one direct paragraph in a rich or multi-paragraph table cell.
    /// Other paragraphs, nested tables, cell properties, and opaque cell XML
    /// remain untouched.
    ///
    /// # Errors
    ///
    /// Returns a checked selector/refusal, resource-limit, or malformed XML
    /// error without changing the projected snapshot.
    pub fn replace_table_cell_paragraph_text(
        &mut self,
        table: Position,
        row: Position,
        cell: Position,
        paragraph: Position,
        authored_text: impl Into<String>,
    ) -> TransactionResult<&mut Self> {
        self.reserve_operation()?;
        let text = authored_text.into();
        validate_authored_text(&text).map_err(|reason| TransactionError::Refused {
            position: table.get(),
            reason,
        })?;
        let replacement_text_bytes = self.checked_text_total(text.len())?;
        let cell_selection = select_cell(&self.projected, table, row, cell)?;
        let paragraph_range = select_direct_child(
            cell_selection.xml,
            b"tc",
            b"p",
            paragraph,
            Refusal::CellNotFound,
        )
        .map_err(|reason| TransactionError::Refused {
            position: table.get(),
            reason,
        })?;
        let paragraph_start = checked_relative_start(cell_selection.start, paragraph_range)?;
        let paragraph_end = checked_relative_end(cell_selection.start, paragraph_range)?;
        let paragraph_xml = checked_slice(
            self.projected.xml_bytes(),
            paragraph_start,
            paragraph_end,
            "table cell paragraph",
        )?;
        let owner =
            scan_text_owner(paragraph_xml, b"p").map_err(|reason| TransactionError::Refused {
                position: table.get(),
                reason,
            })?;
        if owner.text == text {
            return Ok(self);
        }
        let replacement = rewrite_text_owner(paragraph_xml, &owner, &text)?;
        let xml = replace_range(
            self.projected.xml_bytes(),
            paragraph_start,
            paragraph_end,
            &replacement,
        )?;
        let candidate = Snapshot::from_xml(xml)?;
        let actual = selected_cell_paragraph_text(&candidate, table, row, cell, paragraph)?;
        if actual != text {
            return Err(crate::Error::InvalidFormat(
                "document table-cell paragraph edit failed semantic readback".into(),
            )
            .into());
        }
        self.operations.push(Operation::ReplaceCellParagraphText {
            table,
            row,
            cell,
            paragraph,
            before: owner.text,
            after: text,
        });
        self.replacement_text_bytes = replacement_text_bytes;
        self.projected = candidate;
        Ok(self)
    }

    /// Insert a compact plain-text paragraph at a projected zero-based
    /// position. `position == paragraph_count()` appends before the body-final
    /// section properties.
    ///
    /// # Errors
    ///
    /// Returns a typed refusal, checked-position error, resource-limit error,
    /// or malformed-document error without changing the projected snapshot.
    pub fn insert_paragraph(
        &mut self,
        position: Position,
        authored_text: impl Into<String>,
    ) -> TransactionResult<&mut Self> {
        self.reserve_operation()?;
        let text = authored_text.into();
        validate_authored_text(&text).map_err(|reason| TransactionError::Refused {
            position: position.get(),
            reason,
        })?;
        let replacement_text_bytes = self.checked_text_total(text.len())?;
        let count = self.projected.paragraph_count();
        if position.get() > count {
            return Err(TransactionError::OutOfBounds {
                position: position.get(),
                len: count,
            });
        }
        let offset = if position.get() == count {
            usize::try_from(self.projected.content_end).map_err(|_conversion_error| {
                crate::Error::InvalidFormat("document insertion offset does not fit usize".into())
            })?
        } else {
            usize::try_from(self.range(position)?.start).map_err(|_conversion_error| {
                crate::Error::InvalidFormat("paragraph offset does not fit usize".into())
            })?
        };
        let paragraph = plain_paragraph(self.projected.conformance, &text);
        let xml = replace_range(
            self.projected.xml_bytes(),
            offset,
            offset,
            paragraph.as_bytes(),
        )?;
        let candidate = Snapshot::from_xml(xml)?;
        let readback = candidate
            .paragraph(position)
            .ok_or(TransactionError::OutOfBounds {
                position: position.get(),
                len: candidate.paragraph_count(),
            })?
            .text()?;
        let expected_count = count.checked_add(1).ok_or(TransactionError::Limit {
            resource: "paragraphs",
            max: usize::MAX,
            actual: usize::MAX,
        })?;
        if readback != text || candidate.paragraph_count() != expected_count {
            return Err(crate::Error::InvalidFormat(
                "document paragraph insertion failed semantic readback".into(),
            )
            .into());
        }
        self.operations
            .push(Operation::InsertParagraph { position, text });
        self.replacement_text_bytes = replacement_text_bytes;
        self.projected = candidate;
        Ok(self)
    }

    /// Insert one non-mutating dependency-checked paragraph transfer plan.
    ///
    /// The plan is receiver-specific: every relationship reference was mapped
    /// to an equivalent relationship already owned by the exact target
    /// package. This method never copies or guesses package dependencies.
    ///
    /// # Errors
    ///
    /// Returns a stale-plan, position, operation-bound, or XML validation
    /// error without changing the projected snapshot.
    pub fn insert_paragraph_transfer(
        &mut self,
        position: Position,
        plan: &ParagraphTransfer,
    ) -> TransactionResult<&mut Self> {
        if plan.target.as_slice() != self.base.xml_bytes() {
            return Err(TransactionError::StaleSource);
        }
        self.insert_transferred_paragraph(
            position,
            Arc::clone(&plan.fragment),
            Arc::clone(&plan.dependency_digest),
        )
    }

    fn apply_operation(&mut self, operation: &Operation) -> TransactionResult<&mut Self> {
        match operation {
            Operation::ReplaceParagraphText {
                position,
                before,
                after,
            } => {
                let actual = self
                    .projected
                    .paragraph(*position)
                    .ok_or(TransactionError::OutOfBounds {
                        position: position.get(),
                        len: self.projected.paragraph_count(),
                    })?
                    .text()?;
                if &actual != before {
                    return Err(TransactionError::SemanticPrecondition);
                }
                self.replace_paragraph_text(*position, after.clone())
            },
            Operation::ReplaceHyperlinkText {
                paragraph,
                hyperlink,
                before,
                after,
            } => {
                if selected_hyperlink_text(&self.projected, *paragraph, *hyperlink)? != *before {
                    return Err(TransactionError::SemanticPrecondition);
                }
                self.replace_hyperlink_text(*paragraph, *hyperlink, after.clone())
            },
            Operation::ReplaceRunText {
                paragraph,
                run,
                before,
                after,
            } => {
                if selected_direct_paragraph_owner_text(
                    &self.projected,
                    *paragraph,
                    *run,
                    b"r",
                    Refusal::RunNotFound,
                )? != *before
                {
                    return Err(TransactionError::SemanticPrecondition);
                }
                self.replace_run_text(*paragraph, *run, after.clone())
            },
            Operation::ReplaceSimpleFieldText {
                paragraph,
                field,
                before,
                after,
            } => {
                if selected_direct_paragraph_owner_text(
                    &self.projected,
                    *paragraph,
                    *field,
                    b"fldSimple",
                    Refusal::FieldNotFound,
                )? != *before
                {
                    return Err(TransactionError::SemanticPrecondition);
                }
                self.replace_simple_field_text(*paragraph, *field, after.clone())
            },
            Operation::ReplaceRevisionText {
                paragraph,
                kind,
                revision,
                before,
                after,
            } => {
                if selected_direct_paragraph_owner_text(
                    &self.projected,
                    *paragraph,
                    *revision,
                    kind.local_name(),
                    Refusal::RevisionNotFound,
                )? != *before
                {
                    return Err(TransactionError::SemanticPrecondition);
                }
                self.replace_revision_text(*paragraph, *kind, *revision, after.clone())
            },
            Operation::ReplaceContentControlText {
                paragraph,
                control,
                before,
                after,
            } => {
                if selected_content_control_text(&self.projected, *paragraph, *control)? != *before
                {
                    return Err(TransactionError::SemanticPrecondition);
                }
                self.replace_content_control_text(*paragraph, *control, after.clone())
            },
            Operation::ReplaceCellText {
                table,
                row,
                cell,
                before,
                after,
            } => {
                if selected_cell_text(&self.projected, *table, *row, *cell)? != *before {
                    return Err(TransactionError::SemanticPrecondition);
                }
                self.replace_table_cell_text(*table, *row, *cell, after.clone())
            },
            Operation::ReplaceCellParagraphText {
                table,
                row,
                cell,
                paragraph,
                before,
                after,
            } => {
                if selected_cell_paragraph_text(&self.projected, *table, *row, *cell, *paragraph)?
                    != *before
                {
                    return Err(TransactionError::SemanticPrecondition);
                }
                self.replace_table_cell_paragraph_text(
                    *table,
                    *row,
                    *cell,
                    *paragraph,
                    after.clone(),
                )
            },
            Operation::InsertParagraph { position, text } => {
                self.insert_paragraph(*position, text.clone())
            },
            Operation::RemoveParagraph { position, text } => {
                self.remove_plain_paragraph(*position, text)
            },
            Operation::InsertTransferredParagraph {
                position,
                xml,
                dependency_digest,
            } => self.insert_transferred_paragraph(
                *position,
                Arc::clone(xml),
                Arc::clone(dependency_digest),
            ),
            Operation::RemoveTransferredParagraph {
                position,
                xml,
                dependency_digest,
            } => self.remove_transferred_paragraph(*position, xml, dependency_digest),
        }
    }

    fn remove_plain_paragraph(
        &mut self,
        position: Position,
        expected_text: &str,
    ) -> TransactionResult<&mut Self> {
        self.reserve_operation()?;
        let range = self.range(position)?;
        let start = checked_start(range, "paragraph")?;
        let end = checked_end(range, "paragraph")?;
        let source = checked_slice(self.projected.xml_bytes(), start, end, "paragraph")?;
        let expected = plain_paragraph(self.projected.conformance, expected_text);
        if source != expected.as_bytes() {
            return Err(TransactionError::SemanticPrecondition);
        }
        let xml = replace_range(self.projected.xml_bytes(), start, end, &[])?;
        let candidate = Snapshot::from_xml(xml)?;
        if candidate.paragraph_count().checked_add(1) != Some(self.projected.paragraph_count()) {
            return Err(crate::Error::InvalidFormat(
                "document paragraph removal failed semantic readback".into(),
            )
            .into());
        }
        self.operations.push(Operation::RemoveParagraph {
            position,
            text: expected_text.to_owned(),
        });
        self.projected = candidate;
        Ok(self)
    }

    fn insert_transferred_paragraph(
        &mut self,
        position: Position,
        xml: Arc<Vec<u8>>,
        dependency_digest: Arc<str>,
    ) -> TransactionResult<&mut Self> {
        self.reserve_operation()?;
        let count = self.projected.paragraph_count();
        if position.get() > count {
            return Err(TransactionError::OutOfBounds {
                position: position.get(),
                len: count,
            });
        }
        let offset = if position.get() == count {
            usize::try_from(self.projected.content_end).map_err(|_error| {
                crate::Error::InvalidFormat("document insertion offset does not fit usize".into())
            })?
        } else {
            checked_start(self.range(position)?, "paragraph")?
        };
        let candidate = Snapshot::from_xml(replace_range(
            self.projected.xml_bytes(),
            offset,
            offset,
            xml.as_slice(),
        )?)?;
        let inserted = candidate
            .paragraph(position)
            .ok_or(TransactionError::OutOfBounds {
                position: position.get(),
                len: candidate.paragraph_count(),
            })?;
        if inserted.xml_bytes() != xml.as_slice() {
            return Err(crate::Error::InvalidFormat(
                "transferred paragraph failed exact readback".into(),
            )
            .into());
        }
        self.operations.push(Operation::InsertTransferredParagraph {
            position,
            xml,
            dependency_digest,
        });
        self.projected = candidate;
        Ok(self)
    }

    fn remove_transferred_paragraph(
        &mut self,
        position: Position,
        xml: &Arc<Vec<u8>>,
        dependency_digest: &Arc<str>,
    ) -> TransactionResult<&mut Self> {
        self.reserve_operation()?;
        let range = self.range(position)?;
        let start = checked_start(range, "paragraph")?;
        let end = checked_end(range, "paragraph")?;
        let source = checked_slice(self.projected.xml_bytes(), start, end, "paragraph")?;
        if source != xml.as_slice() {
            return Err(TransactionError::SemanticPrecondition);
        }
        let candidate =
            Snapshot::from_xml(replace_range(self.projected.xml_bytes(), start, end, &[])?)?;
        self.operations.push(Operation::RemoveTransferredParagraph {
            position,
            xml: Arc::clone(xml),
            dependency_digest: Arc::clone(dependency_digest),
        });
        self.projected = candidate;
        Ok(self)
    }

    fn replace_direct_paragraph_owner_text(
        &mut self,
        paragraph: Position,
        owner: Position,
        child_name: &[u8],
        missing: Refusal,
        text: String,
        operation: impl FnOnce(String, String) -> Operation,
    ) -> TransactionResult<&mut Self> {
        self.reserve_operation()?;
        validate_authored_text(&text).map_err(|reason| TransactionError::Refused {
            position: paragraph.get(),
            reason,
        })?;
        let replacement_text_bytes = self.checked_text_total(text.len())?;
        let paragraph_range = self.range(paragraph)?;
        let paragraph_start = checked_start(paragraph_range, "paragraph")?;
        let paragraph_end = checked_end(paragraph_range, "paragraph")?;
        let paragraph_xml = checked_slice(
            self.projected.xml_bytes(),
            paragraph_start,
            paragraph_end,
            "paragraph",
        )?;
        let owner_range = select_direct_child(paragraph_xml, b"p", child_name, owner, missing)
            .map_err(|reason| TransactionError::Refused {
                position: paragraph.get(),
                reason,
            })?;
        let owner_start = checked_relative_start(paragraph_start, owner_range)?;
        let owner_end = checked_relative_end(paragraph_start, owner_range)?;
        let owner_xml = checked_slice(
            self.projected.xml_bytes(),
            owner_start,
            owner_end,
            "paragraph text owner",
        )?;
        let scanned =
            scan_text_owner(owner_xml, child_name).map_err(|reason| TransactionError::Refused {
                position: paragraph.get(),
                reason,
            })?;
        if scanned.text == text {
            return Ok(self);
        }
        let replacement = rewrite_text_owner(owner_xml, &scanned, &text)?;
        let xml = replace_range(
            self.projected.xml_bytes(),
            owner_start,
            owner_end,
            &replacement,
        )?;
        let candidate = Snapshot::from_xml(xml)?;
        let actual = selected_direct_paragraph_owner_text(
            &candidate, paragraph, owner, child_name, missing,
        )?;
        if actual != text {
            return Err(crate::Error::InvalidFormat(
                "document owner text edit failed semantic readback".into(),
            )
            .into());
        }
        self.operations.push(operation(scanned.text, text));
        self.replacement_text_bytes = replacement_text_bytes;
        self.projected = candidate;
        Ok(self)
    }

    /// Validate and publish the projected snapshot without changing the
    /// source snapshot.
    ///
    /// # Errors
    ///
    /// Reserved for commit-time document validation failures.
    pub fn commit(self) -> TransactionResult<Commit> {
        let projected = if self.base.same_source(&self.projected) {
            self.projected
        } else {
            let source = std::str::from_utf8(self.projected.xml_bytes()).map_err(|error| {
                crate::Error::InvalidFormat(format!(
                    "changed main-document XML is not UTF-8: {error}"
                ))
            })?;
            let compact = crate::writer::doc::compact_changed_document_xml(source)?;
            Snapshot::from_xml(compact.into_bytes())?
        };
        let diagnostics = Diagnostics {
            operations: self.operations.len(),
            changed: !self.base.same_source(&projected),
        };
        let patch = Patch {
            before: self.base,
            after: projected.clone(),
            operations: self.operations.into(),
        };
        Ok(Commit {
            snapshot: projected,
            patch,
            diagnostics,
        })
    }

    fn range(&self, position: Position) -> TransactionResult<Range> {
        self.projected
            .paragraphs
            .get(position.get())
            .copied()
            .ok_or(TransactionError::OutOfBounds {
                position: position.get(),
                len: self.projected.paragraph_count(),
            })
    }

    fn reserve_operation(&self) -> TransactionResult<()> {
        if self.operations.len() >= MAX_OPERATIONS {
            return Err(TransactionError::Limit {
                resource: "operations",
                max: MAX_OPERATIONS,
                actual: self.operations.len().saturating_add(1),
            });
        }
        Ok(())
    }

    fn checked_text_total(&self, bytes: usize) -> TransactionResult<usize> {
        let actual =
            self.replacement_text_bytes
                .checked_add(bytes)
                .ok_or(TransactionError::Limit {
                    resource: "replacement text bytes",
                    max: MAX_REPLACEMENT_TEXT_BYTES,
                    actual: usize::MAX,
                })?;
        if actual > MAX_REPLACEMENT_TEXT_BYTES {
            return Err(TransactionError::Limit {
                resource: "replacement text bytes",
                max: MAX_REPLACEMENT_TEXT_BYTES,
                actual,
            });
        }
        Ok(actual)
    }
}

/// Diagnostics for one successful main-document commit.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct Diagnostics {
    operations: usize,
    changed: bool,
}

impl Diagnostics {
    /// Number of semantic operations in the commit.
    #[must_use]
    pub const fn operations(self) -> usize {
        self.operations
    }

    /// Whether the commit changed the exact main-document bytes.
    #[must_use]
    pub const fn changed(self) -> bool {
        self.changed
    }
}

/// A successful main-document publication.
#[derive(Debug, Clone)]
pub struct Commit {
    snapshot: Snapshot,
    patch: Patch,
    diagnostics: Diagnostics,
}

impl Commit {
    /// Borrow the published snapshot.
    #[must_use]
    pub const fn snapshot(&self) -> &Snapshot {
        &self.snapshot
    }

    /// Borrow the reversible patch.
    #[must_use]
    pub const fn patch(&self) -> &Patch {
        &self.patch
    }

    /// Return content-free commit diagnostics.
    #[must_use]
    pub const fn diagnostics(&self) -> Diagnostics {
        self.diagnostics
    }

    /// Move the snapshot and patch out of the commit.
    #[must_use]
    pub fn into_parts(self) -> (Snapshot, Patch) {
        (self.snapshot, self.patch)
    }
}

/// A reversible, exact-source-checked main-document patch.
#[derive(Debug, Clone)]
pub struct Patch {
    before: Snapshot,
    after: Snapshot,
    operations: Arc<[Operation]>,
}

impl Patch {
    /// Exact immutable source required by this patch.
    #[must_use]
    pub const fn source(&self) -> &Snapshot {
        &self.before
    }

    /// Exact immutable target produced by this patch.
    #[must_use]
    pub const fn target(&self) -> &Snapshot {
        &self.after
    }

    /// Borrow the semantic operations in staging order.
    #[must_use]
    pub fn operations(&self) -> &[Operation] {
        &self.operations
    }

    /// Whether this patch changes the exact main-document bytes.
    #[must_use]
    pub fn changed(&self) -> bool {
        !self.before.same_source(&self.after)
    }

    /// Return the exact inverse patch.
    #[must_use]
    pub fn inverse(&self) -> Self {
        Self {
            before: self.after.clone(),
            after: self.before.clone(),
            operations: self
                .operations
                .iter()
                .rev()
                .map(Operation::inverse)
                .collect::<Vec<_>>()
                .into(),
        }
    }

    /// Apply only when the target has the exact source document bytes.
    ///
    /// # Errors
    ///
    /// Returns [`TransactionError::StaleSource`] when `source` does not match
    /// the exact bytes against which this patch was produced.
    pub fn apply(&self, source: &Snapshot) -> TransactionResult<Snapshot> {
        if !source.same_source(&self.before) {
            return Err(TransactionError::StaleSource);
        }
        Ok(if self.changed() {
            self.after.clone()
        } else {
            source.clone()
        })
    }
}

#[derive(Debug, Clone, Copy)]
struct Range {
    start: u32,
    length: u32,
}

struct Layout {
    paragraphs: Vec<Range>,
    tables: Vec<Range>,
    content_end: u32,
    conformance: Conformance,
}

#[derive(Debug, Clone, Copy)]
enum Conformance {
    Transitional,
    Strict,
}

impl Conformance {
    const fn namespace(self) -> &'static str {
        match self {
            Self::Transitional => "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
            Self::Strict => "http://purl.oclc.org/ooxml/wordprocessingml/main",
        }
    }
}

struct TextOwner {
    slots: Vec<TextSlot>,
    text: String,
}

struct TextSlot {
    start: usize,
    end: usize,
    prefix: Vec<u8>,
    local_name: Vec<u8>,
    characters: usize,
}

enum FragmentPrefix {
    Unseen,
    Unprefixed,
    Prefixed(Vec<u8>),
}

impl FragmentPrefix {
    fn from_name(name: quick_xml::name::QName<'_>) -> Self {
        name.prefix().map_or(Self::Unprefixed, |prefix| {
            Self::Prefixed(prefix.into_inner().to_vec())
        })
    }
}

struct CellSelection<'a> {
    xml: &'a [u8],
    start: usize,
}

fn scan_document(xml: &[u8]) -> TransactionResult<Layout> {
    let mut reader = NsReader::from_reader(xml);
    let mut paragraphs = Vec::new();
    let mut tables = Vec::new();
    let mut body_depth = None;
    let mut body_end = None;
    let mut final_section_start = None;
    let mut pending = None::<(bool, bool, bool, usize)>;
    let mut conformance = None;
    let mut saw_document = false;
    let mut depth = 0usize;
    let mut nodes = 0usize;

    loop {
        let event_start =
            usize::try_from(reader.buffer_position()).map_err(|_conversion_error| {
                crate::Error::InvalidFormat("document offset does not fit usize".into())
            })?;
        let raw_event = reader
            .read_event()
            .map_err(|error| crate::Error::Xml(error.to_string()))?
            .into_owned();
        let resolver = reader.resolver().clone();
        let (namespace, event) = resolver.resolve_event(raw_event);
        let event_end = usize::try_from(reader.buffer_position()).map_err(|_conversion_error| {
            crate::Error::InvalidFormat("document offset does not fit usize".into())
        })?;

        if matches!(event, Event::Start(_) | Event::Empty(_)) {
            nodes = nodes.checked_add(1).ok_or_else(|| {
                crate::Error::InvalidFormat("document element counter overflow".into())
            })?;
            if nodes > MAX_DOCUMENT_NODES {
                return Err(TransactionError::Limit {
                    resource: "XML elements",
                    max: MAX_DOCUMENT_NODES,
                    actual: nodes,
                });
            }
        }

        match event {
            Event::Start(element) => {
                depth = depth.checked_add(1).ok_or_else(|| {
                    crate::Error::InvalidFormat("document XML nesting is too deep".into())
                })?;
                if depth > MAX_DOCUMENT_DEPTH {
                    return Err(TransactionError::Limit {
                        resource: "XML depth",
                        max: MAX_DOCUMENT_DEPTH,
                        actual: depth,
                    });
                }
                let is_word = is_wordprocessing_namespace(&namespace);
                let local = element.local_name();
                if depth == 1 && is_word && local.as_ref() == b"document" {
                    saw_document = true;
                }
                if is_word && local.as_ref() == b"body" {
                    if depth != 2 || !saw_document {
                        return Err(crate::Error::InvalidFormat(
                            "WordprocessingML body is not a direct child of the document root"
                                .into(),
                        )
                        .into());
                    }
                    if body_depth.is_some() || body_end.is_some() {
                        return Err(crate::Error::InvalidFormat(
                            "main document contains multiple bodies".into(),
                        )
                        .into());
                    }
                    body_depth = Some(depth);
                    conformance = conformance_from_namespace(&namespace);
                } else if body_depth.is_some_and(|body| depth == body + 1) {
                    let is_paragraph = is_word && local.as_ref() == b"p";
                    let is_table = is_word && local.as_ref() == b"tbl";
                    let is_section = is_word && local.as_ref() == b"sectPr";
                    if final_section_start.is_some() {
                        return Err(crate::Error::InvalidFormat(
                            "body-final section properties are not the final body child".into(),
                        )
                        .into());
                    }
                    pending = Some((is_paragraph, is_table, is_section, event_start));
                }
            },
            Event::Empty(element) => {
                let child_depth = depth.checked_add(1).ok_or_else(|| {
                    crate::Error::InvalidFormat("document XML nesting is too deep".into())
                })?;
                if body_depth.is_some_and(|body| child_depth == body + 1) {
                    let is_word = is_wordprocessing_namespace(&namespace);
                    let local = element.local_name();
                    if final_section_start.is_some() {
                        return Err(crate::Error::InvalidFormat(
                            "body-final section properties are not the final body child".into(),
                        )
                        .into());
                    }
                    if is_word && local.as_ref() == b"p" {
                        paragraphs.push(checked_range(event_start, event_end)?);
                    }
                    if is_word && local.as_ref() == b"tbl" {
                        tables.push(checked_range(event_start, event_end)?);
                    }
                    if is_word && local.as_ref() == b"sectPr" {
                        final_section_start = Some(event_start);
                    }
                }
            },
            Event::End(element) => {
                if let Some((is_paragraph, is_table, is_section, start)) = pending
                    && body_depth.is_some_and(|body| depth == body + 1)
                {
                    if is_paragraph {
                        paragraphs.push(checked_range(start, event_end)?);
                    }
                    if is_table {
                        tables.push(checked_range(start, event_end)?);
                    }
                    if is_section {
                        final_section_start = Some(start);
                    }
                    pending = None;
                }
                if body_depth == Some(depth)
                    && is_wordprocessing_namespace(&namespace)
                    && element.local_name().as_ref() == b"body"
                {
                    body_end = Some(event_start);
                    body_depth = None;
                }
                depth = depth.checked_sub(1).ok_or_else(|| {
                    crate::Error::InvalidFormat("invalid document XML nesting".into())
                })?;
            },
            Event::DocType(_) => {
                return Err(crate::Error::InvalidFormat(
                    "DTD declarations are forbidden in a Word main document".into(),
                )
                .into());
            },
            Event::PI(_) => {
                return Err(crate::Error::InvalidFormat(
                    "processing instructions are forbidden in a Word main document".into(),
                )
                .into());
            },
            Event::Eof if depth != 0 || pending.is_some() => {
                return Err(crate::Error::InvalidFormat(
                    "unterminated Word main document XML".into(),
                )
                .into());
            },
            Event::Eof => break,
            Event::Text(_)
            | Event::CData(_)
            | Event::Comment(_)
            | Event::Decl(_)
            | Event::GeneralRef(_) => {},
        }
    }
    let body_end_offset = body_end.ok_or_else(|| {
        crate::Error::InvalidFormat("main document has no WordprocessingML body".into())
    })?;
    let document_conformance = conformance.ok_or_else(|| {
        crate::Error::InvalidFormat("main document body has no supported namespace".into())
    })?;
    if !saw_document {
        return Err(crate::Error::InvalidFormat(
            "main document has no WordprocessingML document root".into(),
        )
        .into());
    }
    let content_end = u32::try_from(final_section_start.unwrap_or(body_end_offset)).map_err(
        |_conversion_error| {
            crate::Error::InvalidFormat("document insertion offset exceeds u32".into())
        },
    )?;
    Ok(Layout {
        paragraphs,
        tables,
        content_end,
        conformance: document_conformance,
    })
}

fn conformance_from_namespace(namespace: &ResolveResult<'_>) -> Option<Conformance> {
    match namespace {
        ResolveResult::Bound(Namespace(uri)) if *uri == WORDPROCESSINGML_NAMESPACE => {
            Some(Conformance::Transitional)
        },
        ResolveResult::Bound(Namespace(uri)) if *uri == STRICT_WORDPROCESSINGML_NAMESPACE => {
            Some(Conformance::Strict)
        },
        ResolveResult::Bound(_) | ResolveResult::Unbound | ResolveResult::Unknown(_) => None,
    }
}

fn checked_range(start: usize, end: usize) -> TransactionResult<Range> {
    Ok(Range {
        start: u32::try_from(start).map_err(|_conversion_error| {
            crate::Error::InvalidFormat("paragraph offset exceeds u32".into())
        })?,
        length: u32::try_from(
            end.checked_sub(start)
                .ok_or_else(|| crate::Error::InvalidFormat("paragraph range underflow".into()))?,
        )
        .map_err(|_conversion_error| {
            crate::Error::InvalidFormat("paragraph length exceeds u32".into())
        })?,
    })
}

fn scan_text_owner(xml: &[u8], root_name: &[u8]) -> Result<TextOwner, Refusal> {
    let mut reader = NsReader::from_reader(xml);
    let mut fragment_prefix = FragmentPrefix::Unseen;
    let mut root_depth = None;
    let mut run_depth = None;
    let mut slots = Vec::new();
    let mut open_text = None::<(usize, Vec<u8>, Vec<u8>)>;
    let mut text = String::new();
    let mut depth = 0usize;
    let mut saw_root = false;

    loop {
        let event_start = usize::try_from(reader.buffer_position())
            .map_err(|_conversion_error| Refusal::ComplexContent)?;
        let raw_event = reader
            .read_event()
            .map_err(|_xml_error| Refusal::ComplexContent)?
            .into_owned();
        let resolver = reader.resolver().clone();
        let (namespace, event) = resolver.resolve_event(raw_event);
        let event_end = usize::try_from(reader.buffer_position())
            .map_err(|_conversion_error| Refusal::ComplexContent)?;

        match event {
            Event::Start(element) => {
                depth = depth.checked_add(1).ok_or(Refusal::ComplexContent)?;
                if root_depth.is_none() {
                    fragment_prefix = FragmentPrefix::from_name(element.name());
                    if saw_root
                        || !is_transaction_fragment_word_name(
                            &namespace,
                            element.name(),
                            root_name,
                            &fragment_prefix,
                        )
                    {
                        return Err(Refusal::ComplexContent);
                    }
                    root_depth = Some(depth);
                    if root_name == b"r" {
                        run_depth = Some(depth);
                    }
                    saw_root = true;
                } else if run_depth.is_some_and(|run| depth == run + 1) {
                    if is_owner_text_element(
                        root_name,
                        &namespace,
                        element.name(),
                        &fragment_prefix,
                    ) {
                        let prefix = element
                            .name()
                            .prefix()
                            .map_or_else(Vec::new, |value| value.into_inner().to_vec());
                        open_text =
                            Some((event_start, prefix, element.local_name().as_ref().to_vec()));
                    } else if is_structural_run_text(&namespace, element.name(), &fragment_prefix) {
                        return Err(Refusal::ComplexRun);
                    }
                } else if root_depth.is_some_and(|root| depth == root + 1) {
                    if is_transaction_fragment_word_name(
                        &namespace,
                        element.name(),
                        b"r",
                        &fragment_prefix,
                    ) {
                        run_depth = Some(depth);
                    } else if is_fragment_word_element(&namespace, &fragment_prefix)
                        && !(root_name == b"p"
                            && is_transaction_fragment_word_name(
                                &namespace,
                                element.name(),
                                b"pPr",
                                &fragment_prefix,
                            ))
                    {
                        return Err(Refusal::ComplexContent);
                    }
                } else if open_text.is_some() {
                    return Err(Refusal::ComplexRun);
                }
            },
            Event::Empty(element) => {
                let child_depth = depth.checked_add(1).ok_or(Refusal::ComplexContent)?;
                if run_depth.is_some_and(|run| child_depth == run + 1) {
                    if is_owner_text_element(
                        root_name,
                        &namespace,
                        element.name(),
                        &fragment_prefix,
                    ) {
                        let prefix = element
                            .name()
                            .prefix()
                            .map_or_else(Vec::new, |value| value.into_inner().to_vec());
                        slots.push(TextSlot {
                            start: event_start,
                            end: event_end,
                            prefix,
                            local_name: element.local_name().as_ref().to_vec(),
                            characters: 0,
                        });
                    } else if is_structural_run_text(&namespace, element.name(), &fragment_prefix) {
                        return Err(Refusal::ComplexRun);
                    }
                } else if root_depth.is_some_and(|root| child_depth == root + 1) {
                    if is_transaction_fragment_word_name(
                        &namespace,
                        element.name(),
                        b"r",
                        &fragment_prefix,
                    ) {
                        continue;
                    }
                    if is_fragment_word_element(&namespace, &fragment_prefix)
                        && !(root_name == b"p"
                            && is_transaction_fragment_word_name(
                                &namespace,
                                element.name(),
                                b"pPr",
                                &fragment_prefix,
                            ))
                    {
                        return Err(Refusal::ComplexContent);
                    }
                }
            },
            Event::End(element) => {
                if open_text.is_some()
                    && is_owner_text_element(
                        root_name,
                        &namespace,
                        element.name(),
                        &fragment_prefix,
                    )
                {
                    let (start, prefix, local_name) =
                        open_text.take().ok_or(Refusal::ComplexRun)?;
                    let value = decode_text_fragment(
                        xml.get(start..event_end).ok_or(Refusal::ComplexRun)?,
                    )?;
                    let characters = value.chars().count();
                    text.push_str(&value);
                    slots.push(TextSlot {
                        start,
                        end: event_end,
                        prefix,
                        local_name,
                        characters,
                    });
                }
                if run_depth == Some(depth)
                    && is_transaction_fragment_word_name(
                        &namespace,
                        element.name(),
                        b"r",
                        &fragment_prefix,
                    )
                {
                    run_depth = None;
                }
                depth = depth.checked_sub(1).ok_or(Refusal::ComplexContent)?;
            },
            Event::Text(event_text) if open_text.is_none() => {
                if !event_text.as_ref().iter().all(u8::is_ascii_whitespace) {
                    return Err(if run_depth.is_some() {
                        Refusal::ComplexRun
                    } else {
                        Refusal::ComplexContent
                    });
                }
            },
            Event::CData(_) | Event::GeneralRef(_) if open_text.is_none() => {
                return Err(if run_depth.is_some() {
                    Refusal::ComplexRun
                } else {
                    Refusal::ComplexContent
                });
            },
            Event::Eof => break,
            Event::Text(_)
            | Event::CData(_)
            | Event::Comment(_)
            | Event::Decl(_)
            | Event::PI(_)
            | Event::DocType(_)
            | Event::GeneralRef(_) => {},
        }
    }
    if !saw_root || depth != 0 || open_text.is_some() {
        return Err(Refusal::ComplexContent);
    }
    if slots.is_empty() {
        return Err(Refusal::ComplexRun);
    }
    Ok(TextOwner { slots, text })
}

fn is_owner_text_element(
    root_name: &[u8],
    namespace: &ResolveResult<'_>,
    name: quick_xml::name::QName<'_>,
    fragment_prefix: &FragmentPrefix,
) -> bool {
    is_transaction_fragment_word_name(
        namespace,
        name,
        if root_name == b"del" {
            b"delText"
        } else {
            b"t"
        },
        fragment_prefix,
    )
}

fn decode_text_fragment(xml: &[u8]) -> Result<String, Refusal> {
    let mut reader = Reader::from_reader(xml);
    let mut value = String::new();
    let mut depth = 0usize;
    loop {
        match reader.read_event().map_err(|_error| Refusal::ComplexRun)? {
            Event::Start(_) => {
                depth = depth.checked_add(1).ok_or(Refusal::ComplexRun)?;
                if depth > 1 {
                    return Err(Refusal::ComplexRun);
                }
            },
            Event::Empty(_) => return Err(Refusal::ComplexRun),
            Event::Text(text) if depth == 1 => {
                let decoded = text
                    .xml_content(XmlVersion::Explicit1_0)
                    .map_err(|_error| Refusal::ComplexRun)?;
                let unescaped =
                    quick_xml::escape::unescape(&decoded).map_err(|_error| Refusal::ComplexRun)?;
                value.push_str(&unescaped);
            },
            Event::CData(text) if depth == 1 => {
                let decoded = text
                    .xml_content(XmlVersion::Explicit1_0)
                    .map_err(|_error| Refusal::ComplexRun)?;
                value.push_str(&decoded);
            },
            Event::GeneralRef(reference) if depth == 1 => {
                value.push_str(
                    &litchi_ooxml_common::xml::decode_xml_reference(&reference)
                        .map_err(|_error| Refusal::ComplexRun)?,
                );
            },
            Event::End(_) => {
                depth = depth.checked_sub(1).ok_or(Refusal::ComplexRun)?;
            },
            Event::Eof => {
                if depth == 0 {
                    break;
                }
                return Err(Refusal::ComplexRun);
            },
            Event::Text(_)
            | Event::CData(_)
            | Event::Comment(_)
            | Event::Decl(_)
            | Event::PI(_)
            | Event::DocType(_)
            | Event::GeneralRef(_) => {},
        }
    }
    Ok(value)
}

fn rewrite_text_owner(xml: &[u8], owner: &TextOwner, text: &str) -> TransactionResult<Vec<u8>> {
    let characters = text.chars().collect::<Vec<_>>();
    let mut cursor = 0usize;
    let mut replacements = Vec::with_capacity(owner.slots.len());
    for (index, slot) in owner.slots.iter().enumerate() {
        let remaining = characters.len().saturating_sub(cursor);
        let count = if index + 1 == owner.slots.len() {
            remaining
        } else {
            slot.characters.min(remaining)
        };
        let value = characters[cursor..cursor + count]
            .iter()
            .collect::<String>();
        cursor = cursor.saturating_add(count);
        replacements.push((
            slot.start,
            slot.end,
            text_element_named(&slot.prefix, &slot.local_name, &value).into_bytes(),
        ));
    }
    replace_ranges(xml, &replacements)
}

fn is_transaction_fragment_word_name(
    namespace: &ResolveResult<'_>,
    name: quick_xml::name::QName<'_>,
    local_name: &[u8],
    fragment_prefix: &FragmentPrefix,
) -> bool {
    name.local_name().as_ref() == local_name && is_fragment_word_element(namespace, fragment_prefix)
}

fn is_fragment_word_element(
    namespace: &ResolveResult<'_>,
    fragment_prefix: &FragmentPrefix,
) -> bool {
    if is_wordprocessing_namespace(namespace) {
        return true;
    }
    match namespace {
        ResolveResult::Unknown(prefix) => matches!(
            fragment_prefix,
            FragmentPrefix::Prefixed(candidate) if candidate.as_slice() == prefix.as_slice()
        ),
        ResolveResult::Unbound => matches!(fragment_prefix, FragmentPrefix::Unprefixed),
        ResolveResult::Bound(_) => false,
    }
}

fn is_structural_run_text(
    namespace: &ResolveResult<'_>,
    name: quick_xml::name::QName<'_>,
    fragment_prefix: &FragmentPrefix,
) -> bool {
    [
        b"tab".as_slice(),
        b"br".as_slice(),
        b"cr".as_slice(),
        b"noBreakHyphen".as_slice(),
        b"softHyphen".as_slice(),
        b"instrText".as_slice(),
        b"delText".as_slice(),
        b"fldChar".as_slice(),
    ]
    .into_iter()
    .any(|local| is_transaction_fragment_word_name(namespace, name, local, fragment_prefix))
}

fn select_direct_child(
    xml: &[u8],
    root_name: &[u8],
    child_name: &[u8],
    position: Position,
    missing: Refusal,
) -> Result<Range, Refusal> {
    let mut reader = NsReader::from_reader(xml);
    let mut fragment_prefix = FragmentPrefix::Unseen;
    let mut root_depth = None;
    let mut capture = None::<(usize, usize)>;
    let mut depth = 0usize;
    let mut index = 0usize;
    let mut saw_root = false;
    loop {
        let start = usize::try_from(reader.buffer_position()).map_err(|_error| missing)?;
        let raw_event = reader.read_event().map_err(|_error| missing)?.into_owned();
        let resolver = reader.resolver().clone();
        let (namespace, event) = resolver.resolve_event(raw_event);
        let end = usize::try_from(reader.buffer_position()).map_err(|_error| missing)?;
        match event {
            Event::Start(element) => {
                depth = depth.checked_add(1).ok_or(missing)?;
                if root_depth.is_none() {
                    fragment_prefix = FragmentPrefix::from_name(element.name());
                    if saw_root
                        || !is_transaction_fragment_word_name(
                            &namespace,
                            element.name(),
                            root_name,
                            &fragment_prefix,
                        )
                    {
                        return Err(missing);
                    }
                    saw_root = true;
                    root_depth = Some(depth);
                } else if root_depth.is_some_and(|root| depth == root + 1)
                    && is_transaction_fragment_word_name(
                        &namespace,
                        element.name(),
                        child_name,
                        &fragment_prefix,
                    )
                {
                    if index == position.get() {
                        capture = Some((start, depth));
                    }
                    index = index.checked_add(1).ok_or(missing)?;
                }
            },
            Event::Empty(element) => {
                let child_depth = depth.checked_add(1).ok_or(missing)?;
                if root_depth.is_some_and(|root| child_depth == root + 1)
                    && is_transaction_fragment_word_name(
                        &namespace,
                        element.name(),
                        child_name,
                        &fragment_prefix,
                    )
                {
                    if index == position.get() {
                        return checked_range(start, end).map_err(|_error| missing);
                    }
                    index = index.checked_add(1).ok_or(missing)?;
                }
            },
            Event::End(_) => {
                if let Some((capture_start, capture_depth)) = capture
                    && depth == capture_depth
                {
                    return checked_range(capture_start, end).map_err(|_error| missing);
                }
                depth = depth.checked_sub(1).ok_or(missing)?;
            },
            Event::Eof if depth != 0 || capture.is_some() => return Err(missing),
            Event::Eof => break,
            Event::Text(_)
            | Event::CData(_)
            | Event::Comment(_)
            | Event::Decl(_)
            | Event::PI(_)
            | Event::DocType(_)
            | Event::GeneralRef(_) => {},
        }
    }
    Err(missing)
}

fn single_cell_paragraph(xml: &[u8]) -> TransactionResult<Range> {
    validate_basic_cell_children(xml).map_err(|reason| TransactionError::Refused {
        position: 0,
        reason,
    })?;
    let first = select_direct_child(xml, b"tc", b"p", Position::new(0), Refusal::ComplexContent)
        .map_err(|reason| TransactionError::Refused {
            position: 0,
            reason,
        })?;
    if select_direct_child(xml, b"tc", b"p", Position::new(1), Refusal::CellNotFound).is_ok() {
        return Err(TransactionError::Refused {
            position: 0,
            reason: Refusal::ComplexContent,
        });
    }
    Ok(first)
}

fn validate_basic_cell_children(xml: &[u8]) -> Result<(), Refusal> {
    let mut reader = NsReader::from_reader(xml);
    let mut fragment_prefix = FragmentPrefix::Unseen;
    let mut root_depth = None;
    let mut depth = 0usize;
    let mut saw_root = false;
    loop {
        let raw_event = reader
            .read_event()
            .map_err(|_error| Refusal::ComplexContent)?
            .into_owned();
        let resolver = reader.resolver().clone();
        let (namespace, event) = resolver.resolve_event(raw_event);
        match event {
            Event::Start(element) => {
                depth = depth.checked_add(1).ok_or(Refusal::ComplexContent)?;
                if root_depth.is_none() {
                    fragment_prefix = FragmentPrefix::from_name(element.name());
                    if saw_root
                        || !is_transaction_fragment_word_name(
                            &namespace,
                            element.name(),
                            b"tc",
                            &fragment_prefix,
                        )
                    {
                        return Err(Refusal::ComplexContent);
                    }
                    saw_root = true;
                    root_depth = Some(depth);
                } else if root_depth.is_some_and(|root| depth == root + 1)
                    && is_fragment_word_element(&namespace, &fragment_prefix)
                    && !is_transaction_fragment_word_name(
                        &namespace,
                        element.name(),
                        b"tcPr",
                        &fragment_prefix,
                    )
                    && !is_transaction_fragment_word_name(
                        &namespace,
                        element.name(),
                        b"p",
                        &fragment_prefix,
                    )
                {
                    return Err(Refusal::ComplexContent);
                }
            },
            Event::Empty(element) => {
                let child_depth = depth.checked_add(1).ok_or(Refusal::ComplexContent)?;
                if root_depth.is_some_and(|root| child_depth == root + 1)
                    && is_fragment_word_element(&namespace, &fragment_prefix)
                    && !is_transaction_fragment_word_name(
                        &namespace,
                        element.name(),
                        b"tcPr",
                        &fragment_prefix,
                    )
                    && !is_transaction_fragment_word_name(
                        &namespace,
                        element.name(),
                        b"p",
                        &fragment_prefix,
                    )
                {
                    return Err(Refusal::ComplexContent);
                }
            },
            Event::End(_) => {
                depth = depth.checked_sub(1).ok_or(Refusal::ComplexContent)?;
            },
            Event::Text(text)
                if root_depth.is_some_and(|root| depth == root)
                    && !text.as_ref().iter().all(u8::is_ascii_whitespace) =>
            {
                return Err(Refusal::ComplexContent);
            },
            Event::CData(_) | Event::GeneralRef(_)
                if root_depth.is_some_and(|root| depth == root) =>
            {
                return Err(Refusal::ComplexContent);
            },
            Event::Eof => break,
            Event::Text(_)
            | Event::CData(_)
            | Event::Comment(_)
            | Event::Decl(_)
            | Event::PI(_)
            | Event::DocType(_)
            | Event::GeneralRef(_) => {},
        }
    }
    if saw_root && depth == 0 {
        Ok(())
    } else {
        Err(Refusal::ComplexContent)
    }
}

fn select_cell(
    snapshot: &Snapshot,
    table: Position,
    row: Position,
    cell: Position,
) -> TransactionResult<CellSelection<'_>> {
    let table_range =
        snapshot
            .tables
            .get(table.get())
            .copied()
            .ok_or(TransactionError::Refused {
                position: table.get(),
                reason: Refusal::CellNotFound,
            })?;
    let table_start = checked_start(table_range, "table")?;
    let table_end = checked_end(table_range, "table")?;
    let table_xml = checked_slice(snapshot.xml_bytes(), table_start, table_end, "table")?;
    let row_range = select_direct_child(table_xml, b"tbl", b"tr", row, Refusal::CellNotFound)
        .map_err(|reason| TransactionError::Refused {
            position: table.get(),
            reason,
        })?;
    let row_start = checked_relative_start(table_start, row_range)?;
    let row_end = checked_relative_end(table_start, row_range)?;
    let row_xml = checked_slice(snapshot.xml_bytes(), row_start, row_end, "table row")?;
    let cell_range = select_direct_child(row_xml, b"tr", b"tc", cell, Refusal::CellNotFound)
        .map_err(|reason| TransactionError::Refused {
            position: table.get(),
            reason,
        })?;
    let cell_start = checked_relative_start(row_start, cell_range)?;
    let cell_end = checked_relative_end(row_start, cell_range)?;
    Ok(CellSelection {
        xml: checked_slice(snapshot.xml_bytes(), cell_start, cell_end, "table cell")?,
        start: cell_start,
    })
}

fn selected_hyperlink_text(
    snapshot: &Snapshot,
    paragraph: Position,
    hyperlink: Position,
) -> TransactionResult<String> {
    let paragraph_range =
        snapshot
            .paragraphs
            .get(paragraph.get())
            .copied()
            .ok_or(TransactionError::OutOfBounds {
                position: paragraph.get(),
                len: snapshot.paragraph_count(),
            })?;
    let start = checked_start(paragraph_range, "paragraph")?;
    let end = checked_end(paragraph_range, "paragraph")?;
    let paragraph_xml = checked_slice(snapshot.xml_bytes(), start, end, "paragraph")?;
    let range = select_direct_child(
        paragraph_xml,
        b"p",
        b"hyperlink",
        hyperlink,
        Refusal::HyperlinkNotFound,
    )
    .map_err(|reason| TransactionError::Refused {
        position: paragraph.get(),
        reason,
    })?;
    let child_start = checked_relative_start(start, range)?;
    let child_end = checked_relative_end(start, range)?;
    scan_text_owner(
        checked_slice(snapshot.xml_bytes(), child_start, child_end, "hyperlink")?,
        b"hyperlink",
    )
    .map(|scanned| scanned.text)
    .map_err(|reason| TransactionError::Refused {
        position: paragraph.get(),
        reason,
    })
}

fn selected_direct_paragraph_owner_text(
    snapshot: &Snapshot,
    paragraph: Position,
    owner: Position,
    child_name: &[u8],
    missing: Refusal,
) -> TransactionResult<String> {
    let paragraph_range =
        snapshot
            .paragraphs
            .get(paragraph.get())
            .copied()
            .ok_or(TransactionError::OutOfBounds {
                position: paragraph.get(),
                len: snapshot.paragraph_count(),
            })?;
    let paragraph_start = checked_start(paragraph_range, "paragraph")?;
    let paragraph_end = checked_end(paragraph_range, "paragraph")?;
    let paragraph_xml = checked_slice(
        snapshot.xml_bytes(),
        paragraph_start,
        paragraph_end,
        "paragraph",
    )?;
    let owner_range = select_direct_child(paragraph_xml, b"p", child_name, owner, missing)
        .map_err(|reason| TransactionError::Refused {
            position: paragraph.get(),
            reason,
        })?;
    let start = checked_relative_start(paragraph_start, owner_range)?;
    let end = checked_relative_end(paragraph_start, owner_range)?;
    scan_text_owner(
        checked_slice(snapshot.xml_bytes(), start, end, "paragraph text owner")?,
        child_name,
    )
    .map(|scanned| scanned.text)
    .map_err(|reason| TransactionError::Refused {
        position: paragraph.get(),
        reason,
    })
}

fn select_content_control_content(
    snapshot: &Snapshot,
    paragraph: Position,
    control: Position,
) -> TransactionResult<(usize, usize)> {
    let paragraph_range =
        snapshot
            .paragraphs
            .get(paragraph.get())
            .copied()
            .ok_or(TransactionError::OutOfBounds {
                position: paragraph.get(),
                len: snapshot.paragraph_count(),
            })?;
    let paragraph_start = checked_start(paragraph_range, "paragraph")?;
    let paragraph_end = checked_end(paragraph_range, "paragraph")?;
    let paragraph_xml = checked_slice(
        snapshot.xml_bytes(),
        paragraph_start,
        paragraph_end,
        "paragraph",
    )?;
    let control_range = select_direct_child(
        paragraph_xml,
        b"p",
        b"sdt",
        control,
        Refusal::ContentControlNotFound,
    )
    .map_err(|reason| TransactionError::Refused {
        position: paragraph.get(),
        reason,
    })?;
    let control_start = checked_relative_start(paragraph_start, control_range)?;
    let control_end = checked_relative_end(paragraph_start, control_range)?;
    let control_xml = checked_slice(
        snapshot.xml_bytes(),
        control_start,
        control_end,
        "content control",
    )?;
    let content_range = select_direct_child(
        control_xml,
        b"sdt",
        b"sdtContent",
        Position::new(0),
        Refusal::ComplexContent,
    )
    .map_err(|reason| TransactionError::Refused {
        position: paragraph.get(),
        reason,
    })?;
    if select_direct_child(
        control_xml,
        b"sdt",
        b"sdtContent",
        Position::new(1),
        Refusal::ComplexContent,
    )
    .is_ok()
    {
        return Err(TransactionError::Refused {
            position: paragraph.get(),
            reason: Refusal::ComplexContent,
        });
    }
    Ok((
        checked_relative_start(control_start, content_range)?,
        checked_relative_end(control_start, content_range)?,
    ))
}

fn selected_content_control_text(
    snapshot: &Snapshot,
    paragraph: Position,
    control: Position,
) -> TransactionResult<String> {
    let (start, end) = select_content_control_content(snapshot, paragraph, control)?;
    scan_text_owner(
        checked_slice(snapshot.xml_bytes(), start, end, "content control content")?,
        b"sdtContent",
    )
    .map(|owner| owner.text)
    .map_err(|reason| TransactionError::Refused {
        position: paragraph.get(),
        reason,
    })
}

fn selected_cell_text(
    snapshot: &Snapshot,
    table: Position,
    row: Position,
    cell: Position,
) -> TransactionResult<String> {
    let selection = select_cell(snapshot, table, row, cell)?;
    let paragraph = single_cell_paragraph(selection.xml)?;
    let start = checked_relative_start(selection.start, paragraph)?;
    let end = checked_relative_end(selection.start, paragraph)?;
    scan_text_owner(
        checked_slice(snapshot.xml_bytes(), start, end, "table cell paragraph")?,
        b"p",
    )
    .map(|owner| owner.text)
    .map_err(|reason| TransactionError::Refused {
        position: table.get(),
        reason,
    })
}

fn selected_cell_paragraph_text(
    snapshot: &Snapshot,
    table: Position,
    row: Position,
    cell: Position,
    paragraph: Position,
) -> TransactionResult<String> {
    let selection = select_cell(snapshot, table, row, cell)?;
    let paragraph_range =
        select_direct_child(selection.xml, b"tc", b"p", paragraph, Refusal::CellNotFound).map_err(
            |reason| TransactionError::Refused {
                position: table.get(),
                reason,
            },
        )?;
    let start = checked_relative_start(selection.start, paragraph_range)?;
    let end = checked_relative_end(selection.start, paragraph_range)?;
    scan_text_owner(
        checked_slice(snapshot.xml_bytes(), start, end, "table cell paragraph")?,
        b"p",
    )
    .map(|owner| owner.text)
    .map_err(|reason| TransactionError::Refused {
        position: table.get(),
        reason,
    })
}

fn checked_start(range: Range, resource: &'static str) -> TransactionResult<usize> {
    usize::try_from(range.start).map_err(|_error| {
        crate::Error::InvalidFormat(format!("{resource} offset does not fit usize")).into()
    })
}

fn checked_end(range: Range, resource: &'static str) -> TransactionResult<usize> {
    checked_start(range, resource)?
        .checked_add(usize::try_from(range.length).map_err(|_error| {
            crate::Error::InvalidFormat(format!("{resource} length does not fit usize"))
        })?)
        .ok_or_else(|| {
            crate::Error::InvalidFormat(format!("{resource} range overflows usize")).into()
        })
}

fn checked_relative_start(base: usize, range: Range) -> TransactionResult<usize> {
    base.checked_add(usize::try_from(range.start).map_err(|_error| {
        crate::Error::InvalidFormat("relative XML offset does not fit usize".into())
    })?)
    .ok_or_else(|| crate::Error::InvalidFormat("relative XML offset overflows".into()).into())
}

fn checked_relative_end(base: usize, range: Range) -> TransactionResult<usize> {
    checked_relative_start(base, range)?
        .checked_add(usize::try_from(range.length).map_err(|_error| {
            crate::Error::InvalidFormat("relative XML length does not fit usize".into())
        })?)
        .ok_or_else(|| crate::Error::InvalidFormat("relative XML range overflows".into()).into())
}

fn checked_slice<'a>(
    source: &'a [u8],
    start: usize,
    end: usize,
    resource: &'static str,
) -> TransactionResult<&'a [u8]> {
    source.get(start..end).ok_or_else(|| {
        crate::Error::InvalidFormat(format!("{resource} range is outside document XML")).into()
    })
}

fn replace_ranges(
    source: &[u8],
    replacements: &[(usize, usize, Vec<u8>)],
) -> TransactionResult<Vec<u8>> {
    let mut output = source.to_vec();
    for (start, end, replacement) in replacements.iter().rev() {
        output = replace_range(&output, *start, *end, replacement)?;
    }
    Ok(output)
}

fn validate_authored_text(text: &str) -> Result<(), Refusal> {
    if text.contains(['\t', '\n', '\r'])
        || text.chars().any(|character| {
            !matches!(character, '\u{20}'..='\u{D7FF}' | '\u{E000}'..='\u{FFFD}' | '\u{10000}'..='\u{10FFFF}')
        })
    {
        return Err(Refusal::StructuralText);
    }
    Ok(())
}

fn text_element(prefix: &[u8], text: &str) -> String {
    text_element_named(prefix, b"t", text)
}

fn text_element_named(prefix: &[u8], local_name: &[u8], text: &str) -> String {
    let prefix_text = String::from_utf8_lossy(prefix);
    let local_name_text = String::from_utf8_lossy(local_name);
    let name = if prefix_text.is_empty() {
        local_name_text.into_owned()
    } else {
        format!("{prefix_text}:{local_name_text}")
    };
    if text.is_empty() {
        return format!("<{name}/>");
    }
    let preserve = text.chars().next().is_some_and(char::is_whitespace)
        || text.chars().next_back().is_some_and(char::is_whitespace);
    if preserve {
        format!(
            "<{name} xml:space=\"preserve\">{}</{name}>",
            escape_xml(text)
        )
    } else {
        format!("<{name}>{}</{name}>", escape_xml(text))
    }
}

fn plain_paragraph(conformance: Conformance, text: &str) -> String {
    if text.is_empty() {
        return format!("<w:p xmlns:w=\"{}\"/>", conformance.namespace());
    }
    format!(
        "<w:p xmlns:w=\"{}\"><w:r>{}</w:r></w:p>",
        conformance.namespace(),
        text_element(b"w", text)
    )
}

fn replace_range(
    source: &[u8],
    start: usize,
    end: usize,
    replacement: &[u8],
) -> TransactionResult<Vec<u8>> {
    if start > end || end > source.len() {
        return Err(crate::Error::InvalidFormat("invalid document rewrite range".into()).into());
    }
    let capacity = source
        .len()
        .checked_sub(end - start)
        .and_then(|size| size.checked_add(replacement.len()))
        .ok_or(TransactionError::Limit {
            resource: "projected XML bytes",
            max: MAX_DOCUMENT_XML_BYTES,
            actual: usize::MAX,
        })?;
    if capacity > MAX_DOCUMENT_XML_BYTES {
        return Err(TransactionError::Limit {
            resource: "projected XML bytes",
            max: MAX_DOCUMENT_XML_BYTES,
            actual: capacity,
        });
    }
    let mut output = Vec::new();
    output
        .try_reserve_exact(capacity)
        .map_err(|allocation_error| crate::Error::Allocation {
            resource: "document transaction XML",
            source: allocation_error,
        })?;
    output.extend_from_slice(&source[..start]);
    output.extend_from_slice(replacement);
    output.extend_from_slice(&source[end..]);
    Ok(output)
}

#[cfg(test)]
mod tests {
    use super::*;

    const WORD: &str = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

    fn document(body: &str) -> Vec<u8> {
        format!("<w:document xmlns:w=\"{WORD}\"><w:body>{body}<w:sectPr/></w:body></w:document>")
            .into_bytes()
    }

    fn durable_limits() -> litchi_core::patch::PatchLimits {
        litchi_core::patch::PatchLimits::new(
            litchi_core::patch::BlobLimits::new(1, MAX_DOCUMENT_XML_BYTES, MAX_DOCUMENT_XML_BYTES),
            1024 * 1024,
            32,
            8,
            256 * 1024,
            512 * 1024,
        )
    }

    #[test]
    fn length_changing_text_edit_preserves_formatting_and_is_reversible() {
        let source = Snapshot::from_xml(document(
            "<w:p w:rsidR=\"1\"><w:pPr><w:keepNext/></w:pPr><w:r><w:rPr><w:b/></w:rPr><w:t>old</w:t></w:r></w:p>",
        ))
        .unwrap();
        let mut edit = source.edit();
        edit.replace_paragraph_text(Position::new(0), " longer & text ")
            .unwrap();
        let commit = edit.commit().unwrap();

        assert_eq!(
            commit
                .snapshot()
                .paragraph(Position::new(0))
                .unwrap()
                .text()
                .unwrap(),
            " longer & text "
        );
        assert!(std::str::from_utf8(commit.snapshot().xml_bytes()).unwrap().contains("<w:pPr><w:keepNext/></w:pPr><w:r><w:rPr><w:b/></w:rPr><w:t xml:space=\"preserve\"> longer &amp; text </w:t></w:r>"));
        assert_eq!(commit.diagnostics().operations(), 1);
        assert!(commit.diagnostics().changed());

        let restored = commit.patch().inverse().apply(commit.snapshot()).unwrap();
        assert_eq!(restored.xml_bytes(), source.xml_bytes());
        assert!(matches!(commit.patch().apply(&restored), Ok(_)));
    }

    #[test]
    fn insertion_uses_projected_checked_positions_and_strict_namespace() {
        let strict = "http://purl.oclc.org/ooxml/wordprocessingml/main";
        let xml = format!(
            "<s:document xmlns:s=\"{strict}\"><s:body><s:p><s:r><s:t>A</s:t></s:r></s:p><s:sectPr/></s:body></s:document>"
        );
        let source = Snapshot::from_xml(xml.into_bytes()).unwrap();
        let mut edit = source.edit();
        edit.insert_paragraph(Position::new(0), "B")
            .unwrap()
            .insert_paragraph(Position::new(2), " C ")
            .unwrap();
        let commit = edit.commit().unwrap();

        let text = commit
            .snapshot()
            .paragraphs()
            .into_iter()
            .map(|paragraph| paragraph.text().unwrap())
            .collect::<Vec<_>>();
        assert_eq!(text, ["B", "A", " C "]);
        let xml = std::str::from_utf8(commit.snapshot().xml_bytes()).unwrap();
        assert!(xml.contains(&format!(
            "<w:p xmlns:w=\"{strict}\"><w:r><w:t>B</w:t></w:r></w:p>"
        )));
        assert!(!xml.contains('\n'));
    }

    #[test]
    fn refuses_complex_content_and_stale_patch_sources() {
        let source = Snapshot::from_xml(document(
            "<w:p><w:hyperlink w:anchor=\"a\"><w:r><w:t>linked</w:t></w:r></w:hyperlink></w:p><w:p><w:r><w:t>plain</w:t></w:r></w:p>",
        ))
        .unwrap();
        let mut refused = source.edit();
        assert!(matches!(
            refused.replace_paragraph_text(Position::new(0), "no"),
            Err(TransactionError::Refused {
                reason: Refusal::ComplexContent,
                ..
            })
        ));

        let mut edit = source.edit();
        edit.replace_paragraph_text(Position::new(1), "changed")
            .unwrap();
        let commit = edit.commit().unwrap();
        let stale = Snapshot::from_xml(document("<w:p><w:r><w:t>other</w:t></w:r></w:p>")).unwrap();
        assert!(matches!(
            commit.patch().apply(&stale),
            Err(TransactionError::StaleSource)
        ));
    }

    #[test]
    fn exact_noop_shares_snapshot_bytes_and_records_no_operation() {
        let source = Snapshot::from_xml(document("<w:p><w:r><w:t>same</w:t></w:r></w:p>")).unwrap();
        let mut edit = source.edit();
        edit.replace_paragraph_text(Position::new(0), "same")
            .unwrap();
        let commit = edit.commit().unwrap();

        assert!(!commit.patch().changed());
        assert!(commit.patch().operations().is_empty());
        assert!(Arc::ptr_eq(&source.xml, &commit.snapshot().xml));
    }

    #[test]
    fn multi_run_hyperlink_and_cell_edits_preserve_formatting_and_unknown_xml() {
        let source = Snapshot::from_xml(document(
            "<w:p w:rsidR=\"01\"><w:pPr><w:keepNext/></w:pPr><w:r><w:rPr><w:b/></w:rPr><w:t>Bold</w:t></w:r><w:r><w:rPr><w:i/></w:rPr><w:drawing><x:opaque xmlns:x=\"urn:test\"/></w:drawing><w:t>tail</w:t></w:r></w:p><w:p><w:hyperlink r:id=\"rId9\" w:tooltip=\"tip\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\"><w:r><w:rPr><w:u/></w:rPr><w:t>link</w:t></w:r><w:r><w:t> text</w:t></w:r></w:hyperlink></w:p><w:tbl><w:tblPr><w:tblStyle w:val=\"Grid\"/></w:tblPr><w:tr><w:trPr><w:cantSplit/></w:trPr><w:tc><w:tcPr><w:shd w:fill=\"FFFF00\"/></w:tcPr><w:p><w:r><w:rPr><w:b/></w:rPr><w:t>cell</w:t></w:r><w:r><z:keep xmlns:z=\"urn:test\"/><w:t> tail</w:t></w:r></w:p></w:tc></w:tr></w:tbl>",
        ))
        .unwrap();
        let mut edit = source.edit();
        edit.replace_paragraph_text(Position::new(0), "Reformatted")
            .unwrap()
            .replace_hyperlink_text(Position::new(1), Position::new(0), "new target")
            .unwrap()
            .replace_table_cell_text(
                Position::new(0),
                Position::new(0),
                Position::new(0),
                "updated cell",
            )
            .unwrap();
        let commit = edit.commit().unwrap();
        let xml = std::str::from_utf8(commit.snapshot().xml_bytes()).unwrap();

        assert_eq!(
            commit
                .snapshot()
                .paragraph(Position::new(0))
                .unwrap()
                .text()
                .unwrap(),
            "Reformatted"
        );
        assert_eq!(
            selected_hyperlink_text(commit.snapshot(), Position::new(1), Position::new(0)).unwrap(),
            "new target"
        );
        assert_eq!(
            selected_cell_text(
                commit.snapshot(),
                Position::new(0),
                Position::new(0),
                Position::new(0)
            )
            .unwrap(),
            "updated cell"
        );
        for retained in [
            "<w:pPr><w:keepNext/></w:pPr>",
            "<w:rPr><w:b/></w:rPr>",
            "<w:rPr><w:i/></w:rPr>",
            "<x:opaque xmlns:x=\"urn:test\"/>",
            "r:id=\"rId9\"",
            "w:tooltip=\"tip\"",
            "<w:shd w:fill=\"FFFF00\"/>",
            "<z:keep xmlns:z=\"urn:test\"/>",
        ] {
            assert!(xml.contains(retained), "missing retained XML: {retained}");
        }
        assert_eq!(commit.patch().operations().len(), 3);
        assert_eq!(
            commit
                .patch()
                .inverse()
                .apply(commit.snapshot())
                .unwrap()
                .xml_bytes(),
            source.xml_bytes()
        );
    }

    #[test]
    fn rich_owner_edits_are_compact_durable_and_exactly_reversible() {
        let source = Snapshot::from_xml(
            format!(
                "<w:document xmlns:w=\"{WORD}\">\n  <w:body>\n    <w:p><w:r><w:rPr><w:b/></w:rPr><w:t>direct</w:t><x:keep xmlns:x=\"urn:test\"/></w:r><w:fldSimple w:instr=\" AUTHOR \"><w:r><w:rPr><w:i/></w:rPr><w:t>field</w:t></w:r></w:fldSimple><w:ins w:id=\"7\" w:author=\"A\"><w:r><w:t>added</w:t></w:r></w:ins><w:del w:id=\"8\" w:author=\"A\"><w:r><w:delText>gone &amp; old</w:delText></w:r></w:del><w:sdt><w:sdtPr><w:tag w:val=\"kept\"/></w:sdtPr><w:sdtContent><w:r><w:rPr><w:smallCaps/></w:rPr><w:t>control</w:t></w:r></w:sdtContent></w:sdt></w:p>\n    <w:tbl><w:tr><w:tc><w:tcPr><w:shd w:fill=\"00FF00\"/></w:tcPr><w:p><w:r><w:t>first</w:t></w:r></w:p><w:tbl><w:tr><w:tc><w:p><w:r><w:t>nested kept</w:t></w:r></w:p></w:tc></w:tr></w:tbl><w:p><w:r><w:rPr><w:u/></w:rPr><w:t>second</w:t></w:r></w:p></w:tc></w:tr></w:tbl>\n    <w:sectPr/>\n  </w:body>\n</w:document>"
            )
            .into_bytes(),
        )
        .unwrap();
        let mut edit = source.edit();
        edit.replace_run_text(Position::new(0), Position::new(0), " run & entity ")
            .unwrap()
            .replace_simple_field_text(Position::new(0), Position::new(0), "new field")
            .unwrap()
            .replace_revision_text(
                Position::new(0),
                RevisionKind::Insertion,
                Position::new(0),
                "new insertion",
            )
            .unwrap()
            .replace_content_control_text(Position::new(0), Position::new(0), "new control")
            .unwrap()
            .replace_revision_text(
                Position::new(0),
                RevisionKind::Deletion,
                Position::new(0),
                "new deletion",
            )
            .unwrap()
            .replace_table_cell_paragraph_text(
                Position::new(0),
                Position::new(0),
                Position::new(0),
                Position::new(1),
                "rich cell paragraph",
            )
            .unwrap();
        let commit = edit.commit().unwrap();
        let xml = std::str::from_utf8(commit.snapshot().xml_bytes()).unwrap();
        assert!(!xml.contains("\n  "));
        for retained in [
            "<w:rPr><w:b/></w:rPr>",
            "<x:keep xmlns:x=\"urn:test\"/>",
            "w:instr=\" AUTHOR \"",
            "<w:ins w:id=\"7\" w:author=\"A\">",
            "<w:del w:id=\"8\" w:author=\"A\">",
            "<w:delText>new deletion</w:delText>",
            "<w:tag w:val=\"kept\"/>",
            "<w:rPr><w:smallCaps/></w:rPr>",
            "<w:t>new control</w:t>",
            "<w:t>nested kept</w:t>",
            "<w:rPr><w:u/></w:rPr>",
            "<w:t xml:space=\"preserve\"> run &amp; entity </w:t>",
        ] {
            assert!(xml.contains(retained), "missing retained XML: {retained}");
        }

        let durable = commit.patch().to_durable(durable_limits()).unwrap();
        let wire = durable.to_deterministic_json().unwrap();
        let decoded =
            litchi_core::patch::Patch::<litchi_core::patch::Reversible>::from_deterministic_json(
                &wire,
                durable_limits(),
            )
            .unwrap();
        let applied = source.apply_durable(&decoded).unwrap();
        assert_eq!(applied.xml_bytes(), commit.snapshot().xml_bytes());
        assert_eq!(
            applied
                .apply_durable(&decoded.inverse())
                .unwrap()
                .xml_bytes(),
            source.xml_bytes()
        );
    }

    #[test]
    fn three_way_planning_keeps_disjoint_rich_owners_and_resolves_overlap() {
        let source = Snapshot::from_xml(document(
            "<w:p><w:r><w:t>run</w:t></w:r><w:fldSimple w:instr=\" AUTHOR \"><w:r><w:t>field</w:t></w:r></w:fldSimple></w:p>",
        ))
        .unwrap();
        let limits = CompositionLimits::new(8, 8, 32, 8);

        let mut run = source.edit();
        run.replace_run_text(Position::new(0), Position::new(0), "changed run")
            .unwrap();
        let mut left = source.compose(limits);
        left.join(run.prepare(limits, "run").unwrap()).unwrap();
        let mut field = source.edit();
        field
            .replace_simple_field_text(Position::new(0), Position::new(0), "changed field")
            .unwrap();
        let mut right = source.compose(limits);
        right.join(field.prepare(limits, "field").unwrap()).unwrap();
        let plan = source.plan_three_way(left, right).unwrap();
        assert!(plan.is_clean());
        assert_eq!(
            plan.finish()
                .unwrap()
                .commit()
                .unwrap()
                .patch()
                .operations()
                .len(),
            2
        );

        let branch = |identifier: &str, value: &str| {
            let mut edit = source.edit();
            edit.replace_simple_field_text(Position::new(0), Position::new(0), value)
                .unwrap();
            let mut composition = source.compose(limits);
            composition
                .join(edit.prepare(limits, identifier).unwrap())
                .unwrap();
            composition
        };
        let mut conflict = source
            .plan_three_way(branch("left", "left"), branch("right", "right"))
            .unwrap();
        assert!(!conflict.is_clean());
        assert!(conflict.conflicts().len() >= 1);
        conflict.resolve(MergeChoice::Left);
        let merged = conflict.finish().unwrap().commit().unwrap();
        assert!(
            std::str::from_utf8(merged.snapshot().xml_bytes())
                .unwrap()
                .contains(">left</w:t>")
        );
    }

    #[test]
    fn durable_patch_is_deterministic_stale_checked_and_reversible() {
        let source = Snapshot::from_xml(document(
            "<w:p><w:r><w:t>one</w:t></w:r><w:r><w:t> two</w:t></w:r></w:p><w:tbl><w:tr><w:tc><w:p><w:r><w:t>cell</w:t></w:r></w:p></w:tc></w:tr></w:tbl>",
        ))
        .unwrap();
        let mut edit = source.edit();
        edit.replace_paragraph_text(Position::new(0), "durable text")
            .unwrap()
            .replace_table_cell_text(
                Position::new(0),
                Position::new(0),
                Position::new(0),
                "durable cell",
            )
            .unwrap();
        let commit = edit.commit().unwrap();
        let durable = commit.patch().to_durable(durable_limits()).unwrap();
        let first = durable.to_deterministic_json().unwrap();
        let second = durable.to_deterministic_json().unwrap();
        assert_eq!(first, second);
        let decoded =
            litchi_core::patch::Patch::<litchi_core::patch::Reversible>::from_deterministic_json(
                &first,
                durable_limits(),
            )
            .unwrap();

        let applied = source.apply_durable(&decoded).unwrap();
        assert_eq!(applied.xml_bytes(), commit.snapshot().xml_bytes());
        let restored = applied.apply_durable(&decoded.inverse()).unwrap();
        assert_eq!(restored.xml_bytes(), source.xml_bytes());
        let stale = Snapshot::from_xml(document("<w:p><w:r><w:t>other</w:t></w:r></w:p>")).unwrap();
        assert!(matches!(
            stale.apply_durable(&decoded),
            Err(TransactionError::StaleSource)
        ));
    }

    #[test]
    fn disjoint_composition_and_bounded_history_are_deterministic() {
        let source = Snapshot::from_xml(document(
            "<w:p><w:r><w:t>body</w:t></w:r></w:p><w:tbl><w:tr><w:tc><w:p><w:r><w:t>cell</w:t></w:r></w:p></w:tc></w:tr></w:tbl>",
        ))
        .unwrap();
        let limits = CompositionLimits::new(8, 8, 32, 8);
        let mut paragraph = source.edit();
        paragraph
            .replace_paragraph_text(Position::new(0), "changed body")
            .unwrap();
        let paragraph = paragraph.prepare(limits, "b-paragraph").unwrap();
        let mut cell = source.edit();
        cell.replace_table_cell_text(
            Position::new(0),
            Position::new(0),
            Position::new(0),
            "changed cell",
        )
        .unwrap();
        let cell = cell.prepare(limits, "a-cell").unwrap();
        let mut composition = source.compose(limits);
        composition.join(paragraph).unwrap().join(cell).unwrap();
        let commit = composition.commit().unwrap();
        assert_eq!(commit.patch().operations().len(), 2);
        assert!(matches!(
            commit.patch().operations()[0],
            Operation::ReplaceCellText { .. }
        ));

        let mut left = source.edit();
        left.replace_paragraph_text(Position::new(0), "left")
            .unwrap();
        let left = left.prepare(limits, "left").unwrap();
        let mut right = source.edit();
        right
            .replace_paragraph_text(Position::new(0), "right")
            .unwrap();
        let right = right.prepare(limits, "right").unwrap();
        let mut overlap = source.compose(limits);
        overlap.join(left).unwrap();
        assert!(matches!(
            overlap.join(right).unwrap_err().failure(),
            SubEditJoinFailure::Overlap(_)
        ));

        let mut indexed = source.edit();
        indexed
            .replace_run_text(Position::new(0), Position::new(0), "indexed")
            .unwrap();
        let indexed = indexed.prepare(limits, "indexed-owner").unwrap();
        let mut appended = source.edit();
        appended
            .insert_paragraph(Position::new(source.paragraph_count()), "appended")
            .unwrap();
        let appended = appended.prepare(limits, "append-only").unwrap();
        let mut append_composition = source.compose(limits);
        append_composition
            .join(indexed)
            .unwrap()
            .join(appended)
            .unwrap();
        append_composition.commit().unwrap();

        let mut indexed = source.edit();
        indexed
            .replace_run_text(Position::new(0), Position::new(0), "indexed")
            .unwrap();
        let indexed = indexed.prepare(limits, "indexed-owner").unwrap();
        let mut prefixed = source.edit();
        prefixed
            .insert_paragraph(Position::new(0), "prefixed")
            .unwrap();
        let prefixed = prefixed.prepare(limits, "prefix-insert").unwrap();
        let mut prefix_overlap = source.compose(limits);
        prefix_overlap.join(indexed).unwrap();
        assert!(matches!(
            prefix_overlap.join(prefixed).unwrap_err().failure(),
            SubEditJoinFailure::Overlap(_)
        ));

        let budget = u64::try_from(commit.snapshot().xml_bytes().len()).unwrap();
        let mut history = source.history(HistoryLimits::new(1, budget));
        history.record(commit).unwrap();
        assert!(history.can_undo());
        assert!(history.undo());
        assert_eq!(history.current().xml_bytes(), source.xml_bytes());
        assert!(history.redo());
    }

    #[test]
    fn adversarial_complex_cells_and_structural_run_text_are_atomic_refusals() {
        let source = Snapshot::from_xml(document(
            "<w:p><w:r><w:t>safe</w:t><w:br/></w:r></w:p><w:tbl><w:tr><w:tc><w:p><w:r><w:t>one</w:t></w:r></w:p><w:p><w:r><w:t>two</w:t></w:r></w:p></w:tc></w:tr></w:tbl>",
        ))
        .unwrap();
        let mut edit = source.edit();
        assert!(
            edit.replace_paragraph_text(Position::new(0), "unsafe flatten")
                .is_err()
        );
        assert!(
            edit.replace_table_cell_text(
                Position::new(0),
                Position::new(0),
                Position::new(0),
                "unsafe flatten",
            )
            .is_err()
        );
        assert_eq!(edit.projected().xml_bytes(), source.xml_bytes());
    }
}
