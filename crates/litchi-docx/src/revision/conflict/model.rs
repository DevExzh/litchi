#![expect(
    clippy::shadow_reuse,
    reason = "parser bindings are intentionally refined after validation"
)]
//! Typed, source-coordinate conflict markup values.

use std::sync::Arc;

use crate::{Error, Result};

use super::{codec, validation};

/// The conflict annotation operation.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum Kind {
    /// A concurrent insertion.
    Insert,
    /// A concurrent deletion.
    Delete,
}

/// The XML scope in which a conflict annotation occurs.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum Scope {
    /// Inline run content.
    Inline,
    /// A property container.
    Property,
}

/// A `WordprocessingML` markup identifier.
///
/// Word accepts the complete signed `i32` lexical domain except its reserved
/// sentinel `-1`; the type preserves that domain without narrowing it.
#[derive(Debug, Clone, Copy, PartialEq, Eq, PartialOrd, Ord, Hash)]
pub struct Id(i32);

impl Id {
    /// Validate and construct an identifier.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation cannot be completed.
    pub fn new(value: i32) -> Result<Self> {
        validation::validate_id(value)?;
        Ok(Self(value))
    }

    /// Return the parsed numeric value.
    #[must_use]
    pub const fn get(self) -> i32 {
        self.0
    }
}

/// Author and date metadata common to conflict elements.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Metadata {
    /// The stable markup identifier.
    pub id: Id,
    /// The declared author.
    pub author: String,
    /// Optional lexical date retained as supplied by the XML source.
    pub date: Option<String>,
}

impl Metadata {
    /// Construct validated metadata.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation cannot be completed.
    pub fn new(id: Id, author: String, date: Option<String>) -> Result<Self> {
        validation::validate_metadata(id, &author, date.as_deref())?;
        Ok(Self { id, author, date })
    }
}

/// A half-open byte range into the exact XML source retained by [`Snapshot`].
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct Span {
    /// Inclusive start byte offset.
    start: usize,
    /// Exclusive end byte offset.
    end: usize,
}

impl Span {
    /// Construct a checked half-open source span.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation cannot be completed.
    pub fn new(start: usize, end: usize) -> Result<Self> {
        if start > end {
            return Err(Error::Invalid("conflict span starts after its end".into()));
        }
        Ok(Self { start, end })
    }

    /// Return the span length in bytes.
    #[must_use]
    pub const fn len(self) -> usize {
        self.end - self.start
    }

    /// Whether the span is empty.
    #[must_use]
    pub const fn is_empty(self) -> bool {
        self.start == self.end
    }

    /// Return the inclusive source offset.
    #[must_use]
    pub const fn start(self) -> usize {
        self.start
    }

    /// Return the exclusive source offset.
    #[must_use]
    pub const fn end(self) -> usize {
        self.end
    }
}

/// The exact source location of one XML attribute and its value.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct AttributeSpan {
    /// Full lexical attribute span.
    pub attribute: Span,
    /// Exact lexical value span, excluding quotes.
    pub value: Span,
}

/// An inert conflict annotation with exact source coordinates.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Conflict {
    /// Insertion or deletion annotation kind.
    pub kind: Kind,
    /// Inline or property annotation scope.
    pub scope: Scope,
    /// Declared markup metadata.
    pub metadata: Metadata,
    /// Whole element span.
    pub span: Span,
    /// Opening-tag span.
    pub start_tag: Span,
    /// Optional exact `w:id` attribute span.
    pub id_span: Option<AttributeSpan>,
    /// Inner-content span.
    pub content: Span,
    /// Exact ordered text and CDATA source segments inside this conflict.
    pub(crate) text: Arc<[Span]>,
    /// Optional exact author attribute span.
    pub author_span: Option<AttributeSpan>,
    /// Optional exact date attribute span.
    pub date_span: Option<AttributeSpan>,
}

impl Conflict {
    /// Borrow exact text and CDATA source segments in document order.
    #[must_use]
    pub fn text_spans(&self) -> &[Span] {
        &self.text
    }

    /// Return the smallest source span bounding all text segments.
    ///
    /// This is a convenience coordinate only; callers that need exact text
    /// ownership must use [`Self::text_spans`], because markup can separate
    /// consecutive text segments.
    #[must_use]
    pub fn text_extent(&self) -> Option<Span> {
        let first = self.text.first()?;
        let last = self.text.last()?;
        Span::new(first.start(), last.end()).ok()
    }

    /// Replace parsed text segments before the inventory is published.
    pub(crate) fn set_text_spans(&mut self, text: Arc<[Span]>) {
        self.text = text;
    }
}

/// One paired range boundary pair.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Range {
    /// Insertion or deletion range kind.
    pub kind: Kind,
    /// Shared range metadata.
    pub metadata: Metadata,
    /// Exact start-boundary span.
    pub start_span: Span,
    /// Exact end-boundary span.
    pub end_span: Span,
    /// Optional exact `w:id` attribute on the start boundary.
    pub start_id_span: Option<AttributeSpan>,
    /// Optional exact `w:id` attribute on the end boundary.
    pub end_id_span: Option<AttributeSpan>,
    /// Optional exact `w:author` attribute on the start boundary.
    pub author_span: Option<AttributeSpan>,
    /// Optional exact `w:date` attribute on the start boundary.
    pub date_span: Option<AttributeSpan>,
}

/// Conflict markup in source order.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct Inventory {
    /// Inline and property conflict elements in source order.
    pub conflicts: Vec<Conflict>,
    /// Paired range conflicts in start-boundary source order.
    pub ranges: Vec<Range>,
}

/// Resource bounds for inert conflict-markup processing.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct Limits {
    /// Maximum retained XML source bytes.
    pub max_source_bytes: usize,
    /// Maximum conflict elements.
    pub max_conflicts: usize,
    /// Maximum paired range annotations.
    pub max_ranges: usize,
    /// Maximum aggregate text bytes described by annotations.
    pub max_text_bytes: usize,
    /// Maximum exact text and CDATA segments retained for one story.
    pub max_text_segments: usize,
    /// Maximum XML nesting depth.
    pub max_depth: usize,
    /// Maximum attributes on a conflict element.
    pub max_attributes: usize,
    /// Maximum XML reader events accepted from a source.
    pub max_events: usize,
    /// Maximum aggregate metadata bytes retained after XML decoding.
    pub max_metadata_bytes: usize,
    /// Maximum simultaneously unmatched conflict ranges.
    pub max_open_ranges: usize,
    /// Maximum encoded bytes of one attribute value.
    pub max_attribute_bytes: usize,
    /// Maximum bytes produced by an edited source.
    pub max_output_bytes: usize,
    /// Maximum `WordprocessingML` stories inspected in one package operation.
    pub max_stories: usize,
    /// Maximum aggregate source bytes across inspected stories.
    pub max_total_story_bytes: usize,
    /// Maximum aggregate conflict elements across inspected stories.
    pub max_total_conflicts: usize,
    /// Maximum aggregate decoded conflict metadata bytes across stories.
    pub max_total_metadata_bytes: usize,
    /// Maximum exact text and CDATA segments retained across stories.
    pub max_total_text_segments: usize,
    /// Maximum aggregate paired ranges across inspected stories.
    pub max_total_ranges: usize,
    /// Maximum relationships retained for one story.
    pub max_relationships_per_story: usize,
    /// Maximum relationships retained across all inspected stories.
    pub max_total_relationships: usize,
    /// Maximum canonical topology bytes retained for package stale checks.
    pub max_topology_bytes: usize,
}

impl Default for Limits {
    fn default() -> Self {
        Self {
            max_source_bytes: 16 * 1024 * 1024,
            max_conflicts: 100_000,
            max_ranges: 100_000,
            max_text_bytes: 8 * 1024 * 1024,
            max_text_segments: 200_000,
            max_depth: 256,
            max_attributes: 64,
            max_events: 500_000,
            max_metadata_bytes: 8 * 1024 * 1024,
            max_open_ranges: 100_000,
            max_attribute_bytes: 16 * 1024,
            max_output_bytes: 32 * 1024 * 1024,
            max_stories: 128,
            max_total_story_bytes: 128 * 1024 * 1024,
            max_total_conflicts: 500_000,
            max_total_metadata_bytes: 32 * 1024 * 1024,
            max_total_text_segments: 1_000_000,
            max_total_ranges: 500_000,
            max_relationships_per_story: 4_096,
            max_total_relationships: 32_768,
            max_topology_bytes: 16 * 1024 * 1024,
        }
    }
}

impl Limits {
    /// Validate finite, hard-capped resource limits.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation cannot be completed.
    pub fn validate(self) -> Result<Self> {
        validation::validate_limits(self)?;
        Ok(self)
    }
}

/// Package binding used only by crate-private package integration.
#[derive(Debug, Clone, PartialEq, Eq)]
pub(crate) struct Binding {
    pub(crate) part: String,
    pub(crate) content_type: String,
    pub(crate) topology: Arc<[u8]>,
}

impl Binding {
    pub(crate) fn new(part: String, content_type: String, topology: Arc<[u8]>) -> Self {
        Self {
            part,
            content_type,
            topology,
        }
    }

    pub(crate) fn part(&self) -> &str {
        &self.part
    }

    pub(crate) fn content_type(&self) -> &str {
        &self.content_type
    }

    pub(crate) fn topology(&self) -> &[u8] {
        &self.topology
    }
}

/// A cheaply clonable owner for detached XML or an OPC-native part blob.
#[derive(Debug, Clone)]
pub(crate) enum Source {
    Detached(Arc<[u8]>),
    Blob(Arc<Vec<u8>>),
}

impl Source {
    pub(crate) fn as_slice(&self) -> &[u8] {
        match self {
            Self::Detached(source) => source,
            Self::Blob(source) => source,
        }
    }

    pub(crate) fn len(&self) -> usize {
        self.as_slice().len()
    }
}

impl AsRef<[u8]> for Source {
    fn as_ref(&self) -> &[u8] {
        self.as_slice()
    }
}

#[derive(Debug, Clone)]
struct SnapshotInner {
    source: Source,
    inventory: Arc<Inventory>,
    limits: Limits,
    binding: Option<Binding>,
}

/// An immutable, source-preserving parsed conflict snapshot.
#[derive(Debug, Clone)]
pub struct Snapshot {
    inner: Arc<SnapshotInner>,
}

impl Snapshot {
    /// Parse a conflict-markup XML source using default bounded limits.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation cannot be completed.
    pub fn from_xml(source: impl Into<Vec<u8>>) -> Result<Self> {
        Self::from_xml_with_limits(source, Limits::default())
    }

    /// Parse a conflict-markup XML source using explicit bounded limits.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation cannot be completed.
    pub fn from_xml_with_limits(source: impl Into<Vec<u8>>, limits: Limits) -> Result<Self> {
        Self::from_arc_with_limits(Arc::from(source.into().into_boxed_slice()), limits)
    }

    /// Parse a shared source buffer without copying it.
    pub(crate) fn from_arc_with_limits(source: Arc<[u8]>, limits: Limits) -> Result<Self> {
        Self::from_source_with_limits(Source::Detached(source), limits)
    }

    /// Parse an OPC-native part blob without copying it.
    #[allow(
        dead_code,
        reason = "writer helper is retained for package integration"
    )] // Package integration consumes this without copying an OPC part blob.
    pub(crate) fn from_blob_with_limits(source: Arc<Vec<u8>>, limits: Limits) -> Result<Self> {
        Self::from_source_with_limits(Source::Blob(source), limits)
    }

    pub(crate) fn from_source_with_limits(source: Source, limits: Limits) -> Result<Self> {
        Self::from_source_with_parse_and_retained_limits(source, limits, limits)
    }

    /// Parse with temporary aggregate residuals while retaining the caller's
    /// stable per-story policy in the resulting snapshot.
    pub(crate) fn from_source_with_parse_and_retained_limits(
        source: Source,
        parse_limits: Limits,
        retained_limits: Limits,
    ) -> Result<Self> {
        let parse_limits = parse_limits.validate()?;
        let retained_limits = retained_limits.validate()?;
        let inventory = Arc::new(codec::parse(source.as_slice(), parse_limits)?);
        Ok(Self {
            inner: Arc::new(SnapshotInner {
                source,
                inventory,
                limits: retained_limits,
                binding: None,
            }),
        })
    }

    /// Borrow the exact retained XML source.
    #[must_use]
    pub fn source(&self) -> &[u8] {
        self.inner.source.as_slice()
    }

    /// Borrow the exact retained XML source under shared ownership.
    pub(crate) fn source_owner(&self) -> Source {
        self.inner.source.clone()
    }

    /// Borrow parsed annotations in source order.
    #[must_use]
    pub fn inventory(&self) -> &Inventory {
        &self.inner.inventory
    }

    /// Return the resource limits that governed this snapshot.
    #[must_use]
    pub fn limits(&self) -> Limits {
        self.inner.limits
    }

    /// Start a failure-atomic conflict-markup edit.
    #[must_use]
    pub fn edit(&self) -> super::transaction::Transaction {
        super::transaction::Transaction::new(self.clone())
    }

    pub(crate) fn with_binding(mut self, binding: Binding) -> Self {
        Arc::make_mut(&mut self.inner).binding = Some(binding);
        self
    }

    pub(crate) fn binding(&self) -> Option<&Binding> {
        self.inner.binding.as_ref()
    }
}
