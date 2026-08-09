//! Source-preserving main-document snapshots, edits, and reversible patches.

use std::sync::Arc;

use litchi_core::Position;
use litchi_core::xml::escape_xml;
use quick_xml::events::Event;
use quick_xml::name::{Namespace, ResolveResult};
use quick_xml::reader::NsReader;
use thiserror::Error;

use crate::namespace::{
    STRICT_WORDPROCESSINGML_NAMESPACE, WORDPROCESSINGML_NAMESPACE, is_wordprocessing_namespace,
};
use crate::paragraph::{Paragraph, is_fragment_word_name};

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
    /// The paragraph has multiple runs or non-text inline content.
    ComplexContent,
    /// The selected run has multiple text nodes or another run child.
    ComplexRun,
    /// The requested text needs structural run elements such as `w:tab` or
    /// `w:br`, which this focused text operation does not synthesize.
    StructuralText,
}

impl std::fmt::Display for Refusal {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        formatter.write_str(match self {
            Self::ComplexContent => "paragraph contains non-simple inline content",
            Self::ComplexRun => "paragraph run contains non-simple text content",
            Self::StructuralText => "text requires structural WordprocessingML elements",
        })
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
}

/// An immutable, cheaply clonable snapshot of the main document XML.
#[derive(Debug, Clone)]
pub struct Snapshot {
    xml: Arc<Vec<u8>>,
    paragraphs: Arc<[Range]>,
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
    /// Replace the complete text of one simple paragraph.
    ReplaceText {
        /// Projected paragraph position at the time of the operation.
        position: Position,
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
}

impl Operation {
    fn inverse(&self) -> Self {
        match self {
            Self::ReplaceText {
                position,
                before,
                after,
            } => Self::ReplaceText {
                position: *position,
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

    /// Replace all text in a simple direct-body paragraph.
    ///
    /// The paragraph may retain `w:pPr` and its single run may retain `w:rPr`.
    /// Hyperlinks, fields, revisions, controls, bookmarks, multiple runs, and
    /// structural run content are refused rather than flattened.
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
        let simple =
            scan_simple_paragraph(paragraph).map_err(|reason| TransactionError::Refused {
                position: position.get(),
                reason,
            })?;
        if simple.text == text {
            return Ok(self);
        }
        let replacement = text_element(&simple.prefix, &text);
        let start = paragraph_start
            .checked_add(simple.start)
            .ok_or_else(|| crate::Error::InvalidFormat("text range overflow".into()))?;
        let end = paragraph_start
            .checked_add(simple.end)
            .ok_or_else(|| crate::Error::InvalidFormat("text range overflow".into()))?;
        let xml = replace_range(
            self.projected.xml_bytes(),
            start,
            end,
            replacement.as_bytes(),
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
        self.operations.push(Operation::ReplaceText {
            position,
            before: simple.text,
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

    /// Validate and publish the projected snapshot without changing the
    /// source snapshot.
    ///
    /// # Errors
    ///
    /// Reserved for commit-time document validation failures.
    pub fn commit(self) -> TransactionResult<Commit> {
        let diagnostics = Diagnostics {
            operations: self.operations.len(),
            changed: !self.base.same_source(&self.projected),
        };
        let patch = Patch {
            before: self.base,
            after: self.projected.clone(),
            operations: self.operations.into(),
        };
        Ok(Commit {
            snapshot: self.projected,
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

struct SimpleParagraph {
    start: usize,
    end: usize,
    prefix: Vec<u8>,
    text: String,
}

fn scan_document(xml: &[u8]) -> TransactionResult<Layout> {
    let mut reader = NsReader::from_reader(xml);
    let mut paragraphs = Vec::new();
    let mut body_depth = None;
    let mut body_end = None;
    let mut final_section_start = None;
    let mut pending = None::<(bool, bool, usize)>;
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
                    let is_section = is_word && local.as_ref() == b"sectPr";
                    if final_section_start.is_some() {
                        return Err(crate::Error::InvalidFormat(
                            "body-final section properties are not the final body child".into(),
                        )
                        .into());
                    }
                    pending = Some((is_paragraph, is_section, event_start));
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
                    if is_word && local.as_ref() == b"sectPr" {
                        final_section_start = Some(event_start);
                    }
                }
            },
            Event::End(element) => {
                if let Some((is_paragraph, is_section, start)) = pending
                    && body_depth.is_some_and(|body| depth == body + 1)
                {
                    if is_paragraph {
                        paragraphs.push(checked_range(start, event_end)?);
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

fn scan_simple_paragraph(xml: &[u8]) -> Result<SimpleParagraph, Refusal> {
    let mut reader = NsReader::from_reader(xml);
    let mut fragment_prefix: Option<Option<Vec<u8>>> = None;
    let mut root_depth = None;
    let mut run_depth = None;
    let mut text_range = None::<(usize, usize, Vec<u8>)>;
    let mut open_text = None::<(usize, Vec<u8>)>;
    let mut depth = 0usize;
    let mut runs = 0usize;
    let mut texts = 0usize;

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
                    fragment_prefix = Some(
                        element
                            .name()
                            .prefix()
                            .map(|prefix| prefix.into_inner().to_vec()),
                    );
                    if !is_fragment_word_name(&namespace, element.name(), b"p", &fragment_prefix) {
                        return Err(Refusal::ComplexContent);
                    }
                    root_depth = Some(depth);
                } else if root_depth.is_some_and(|root| depth == root + 1) {
                    if is_fragment_word_name(&namespace, element.name(), b"r", &fragment_prefix) {
                        runs = runs.checked_add(1).ok_or(Refusal::ComplexContent)?;
                        if runs != 1 {
                            return Err(Refusal::ComplexContent);
                        }
                        run_depth = Some(depth);
                    } else if !is_fragment_word_name(
                        &namespace,
                        element.name(),
                        b"pPr",
                        &fragment_prefix,
                    ) {
                        return Err(Refusal::ComplexContent);
                    }
                } else if run_depth.is_some_and(|run| depth == run + 1) {
                    if is_fragment_word_name(&namespace, element.name(), b"t", &fragment_prefix) {
                        texts = texts.checked_add(1).ok_or(Refusal::ComplexRun)?;
                        if texts != 1 {
                            return Err(Refusal::ComplexRun);
                        }
                        let prefix = element
                            .name()
                            .prefix()
                            .map_or_else(Vec::new, |value| value.into_inner().to_vec());
                        open_text = Some((event_start, prefix));
                    } else if !is_fragment_word_name(
                        &namespace,
                        element.name(),
                        b"rPr",
                        &fragment_prefix,
                    ) {
                        return Err(Refusal::ComplexRun);
                    }
                } else if open_text.is_some() {
                    return Err(Refusal::ComplexRun);
                }
            },
            Event::Empty(element) => {
                let child_depth = depth.checked_add(1).ok_or(Refusal::ComplexContent)?;
                if root_depth.is_some_and(|root| child_depth == root + 1) {
                    if is_fragment_word_name(&namespace, element.name(), b"r", &fragment_prefix) {
                        return Err(Refusal::ComplexRun);
                    }
                    if !is_fragment_word_name(&namespace, element.name(), b"pPr", &fragment_prefix)
                    {
                        return Err(Refusal::ComplexContent);
                    }
                } else if run_depth.is_some_and(|run| child_depth == run + 1) {
                    if is_fragment_word_name(&namespace, element.name(), b"t", &fragment_prefix) {
                        texts = texts.checked_add(1).ok_or(Refusal::ComplexRun)?;
                        if texts != 1 {
                            return Err(Refusal::ComplexRun);
                        }
                        let prefix = element
                            .name()
                            .prefix()
                            .map_or_else(Vec::new, |value| value.into_inner().to_vec());
                        text_range = Some((event_start, event_end, prefix));
                    } else if !is_fragment_word_name(
                        &namespace,
                        element.name(),
                        b"rPr",
                        &fragment_prefix,
                    ) {
                        return Err(Refusal::ComplexRun);
                    }
                }
            },
            Event::End(element) => {
                if open_text.is_some()
                    && is_fragment_word_name(&namespace, element.name(), b"t", &fragment_prefix)
                {
                    let (start, prefix) = open_text.take().ok_or(Refusal::ComplexRun)?;
                    text_range = Some((start, event_end, prefix));
                }
                if run_depth == Some(depth)
                    && is_fragment_word_name(&namespace, element.name(), b"r", &fragment_prefix)
                {
                    run_depth = None;
                }
                depth = depth.checked_sub(1).ok_or(Refusal::ComplexContent)?;
            },
            Event::Text(text) if open_text.is_none() => {
                if !text.as_ref().iter().all(u8::is_ascii_whitespace) {
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
    if runs != 1 || texts != 1 || depth != 0 || open_text.is_some() {
        return Err(if runs == 1 {
            Refusal::ComplexRun
        } else {
            Refusal::ComplexContent
        });
    }
    let (start, end, prefix) = text_range.ok_or(Refusal::ComplexRun)?;
    let text = Paragraph::new(xml.to_vec())
        .text()
        .map_err(|_text_error| Refusal::ComplexRun)?;
    Ok(SimpleParagraph {
        start,
        end,
        prefix,
        text,
    })
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
    let prefix_text = String::from_utf8_lossy(prefix);
    let name = if prefix_text.is_empty() {
        "t".to_owned()
    } else {
        format!("{prefix_text}:t")
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
}
