use std::fmt;
use std::ops::Range;
use std::str::Utf8Error;
use std::sync::Arc;

use thiserror::Error;

use super::parse;
use super::transaction::{Edit, Patch};

/// The Markdown grammar used to interpret an exact source snapshot.
#[derive(Clone, Copy, Debug, Default, PartialEq, Eq)]
#[non_exhaustive]
pub enum Dialect {
    /// `CommonMark` 0.31.2 without syntax extensions.
    #[default]
    CommonMark,
    /// `CommonMark` plus GFM tables, task lists, strikethrough, alerts, and
    /// compatible footnote definitions.
    GitHubFlavored,
}

/// Resource limits applied before and during Markdown parsing.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub struct ReadLimits {
    /// Maximum retained UTF-8 source length.
    pub max_source_bytes: usize,
    /// Maximum length of one physical source line, excluding its line ending.
    pub max_line_bytes: usize,
    /// Maximum parser events produced by the source.
    pub max_events: usize,
    /// Maximum number of top-level semantic blocks.
    pub max_blocks: usize,
    /// Maximum simultaneous block and inline container nesting.
    pub max_nesting_depth: usize,
    /// Maximum structural operations in one atomic edit.
    pub max_operations: usize,
}

impl ReadLimits {
    /// Safe defaults for an ordinary Markdown document.
    pub const DEFAULT: Self = Self {
        max_source_bytes: 16 * 1024 * 1024,
        max_line_bytes: 1024 * 1024,
        max_events: 1_000_000,
        max_blocks: 100_000,
        max_nesting_depth: 256,
        max_operations: 10_000,
    };

    pub(crate) fn validate(self) -> Result<(), Error> {
        for (name, value) in [
            ("max_source_bytes", self.max_source_bytes),
            ("max_line_bytes", self.max_line_bytes),
            ("max_events", self.max_events),
            ("max_blocks", self.max_blocks),
            ("max_nesting_depth", self.max_nesting_depth),
            ("max_operations", self.max_operations),
        ] {
            if value == 0 {
                return Err(Error::InvalidLimit { name });
            }
        }
        Ok(())
    }
}

impl Default for ReadLimits {
    fn default() -> Self {
        Self::DEFAULT
    }
}

/// A top-level Markdown block classification.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
#[non_exhaustive]
pub enum BlockKind {
    /// An ordinary paragraph.
    Paragraph,
    /// An ATX or Setext heading.
    Heading {
        /// `CommonMark` heading level in the inclusive range `1..=6`.
        level: u8,
    },
    /// A block quote, including a GFM alert quote.
    BlockQuote,
    /// An indented or fenced code block.
    CodeBlock {
        /// Whether the source uses a backtick or tilde fence.
        fenced: bool,
    },
    /// A raw HTML block.
    Html,
    /// An ordered or unordered list, including every nested item.
    List {
        /// The first ordered-list value, or `None` for an unordered list.
        start: Option<u64>,
    },
    /// One ordered or unordered list item.
    ListItem,
    /// A GFM footnote definition.
    FootnoteDefinition,
    /// A GFM pipe table.
    Table,
    /// A GFM table header group.
    TableHead,
    /// A GFM table row.
    TableRow,
    /// A GFM table cell.
    TableCell,
    /// A thematic break.
    ThematicBreak,
    /// A `CommonMark` link reference definition.
    LinkDefinition,
}

/// A lossless inline-node classification in parser preorder.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
#[non_exhaustive]
pub enum InlineKind {
    /// Literal text; lexical escapes remain available through [`Inline::source`].
    Text,
    /// Emphasized content.
    Emphasis,
    /// Strongly emphasized content.
    Strong,
    /// GFM strikethrough content.
    Strikethrough,
    /// An inline code span.
    Code,
    /// A link of any supported source form.
    Link,
    /// An image.
    Image,
    /// Raw inline HTML.
    Html,
    /// A GFM footnote reference.
    FootnoteReference,
    /// A soft source line break.
    SoftBreak,
    /// An explicit hard line break.
    HardBreak,
    /// A GFM task-list checkbox marker.
    TaskListMarker {
        /// Whether the source marker is checked.
        checked: bool,
    },
}

/// A reference-graph edge or definition.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
#[non_exhaustive]
pub enum ReferenceKind {
    /// A link destination or reference use.
    Link,
    /// An image destination or reference use.
    Image,
    /// A footnote use.
    Footnote,
    /// A link reference definition.
    LinkDefinition,
    /// A footnote definition.
    FootnoteDefinition,
}

/// A borrowed exact-source top-level block.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct Block<'snapshot> {
    pub(crate) source: &'snapshot str,
    pub(crate) record: &'snapshot BlockRecord,
}

impl<'snapshot> Block<'snapshot> {
    /// Semantic classification of this block.
    #[must_use]
    pub const fn kind(&self) -> BlockKind {
        self.record.kind
    }

    /// Exact byte range in the complete snapshot source.
    #[must_use]
    pub fn range(&self) -> Range<usize> {
        self.record.range.clone()
    }

    /// Exact Markdown source for this block.
    #[must_use]
    pub fn source(&self) -> &'snapshot str {
        &self.source[self.record.range.clone()]
    }

    /// Iterate over inline nodes contained by this block in parser preorder.
    #[must_use]
    pub fn inlines(&self) -> Inlines<'snapshot> {
        Inlines {
            source: self.source,
            records: self.record.inlines.iter(),
        }
    }

    /// Iterate over nested block nodes in parser preorder.
    #[must_use]
    pub fn descendants(&self) -> NestedBlocks<'snapshot> {
        NestedBlocks {
            source: self.source,
            records: self.record.descendants.iter(),
        }
    }
}

/// A borrowed nested block node within one top-level block.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct NestedBlock<'snapshot> {
    source: &'snapshot str,
    record: &'snapshot NestedBlockRecord,
}

impl<'snapshot> NestedBlock<'snapshot> {
    /// Parser nesting depth relative to the top-level block.
    #[must_use]
    pub const fn depth(&self) -> usize {
        self.record.depth
    }

    /// Semantic classification of this nested block.
    #[must_use]
    pub const fn kind(&self) -> BlockKind {
        self.record.kind
    }

    /// Exact byte range in the complete snapshot source.
    #[must_use]
    pub fn range(&self) -> Range<usize> {
        self.record.range.clone()
    }

    /// Exact Markdown source represented by this node.
    #[must_use]
    pub fn source(&self) -> &'snapshot str {
        &self.source[self.record.range.clone()]
    }
}

/// Iterator over nested block nodes in parser preorder.
#[derive(Clone, Debug)]
pub struct NestedBlocks<'snapshot> {
    source: &'snapshot str,
    records: std::slice::Iter<'snapshot, NestedBlockRecord>,
}

impl ExactSizeIterator for NestedBlocks<'_> {}

impl<'snapshot> Iterator for NestedBlocks<'snapshot> {
    type Item = NestedBlock<'snapshot>;

    fn next(&mut self) -> Option<Self::Item> {
        self.records.next().map(|record| NestedBlock {
            source: self.source,
            record,
        })
    }

    fn size_hint(&self) -> (usize, Option<usize>) {
        self.records.size_hint()
    }
}

/// A borrowed lossless inline node.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct Inline<'snapshot> {
    source: &'snapshot str,
    record: &'snapshot InlineRecord,
}

impl<'snapshot> Inline<'snapshot> {
    /// Parser nesting depth relative to the containing block.
    #[must_use]
    pub const fn depth(&self) -> usize {
        self.record.depth
    }

    /// Inline classification.
    #[must_use]
    pub const fn kind(&self) -> InlineKind {
        self.record.kind
    }

    /// Exact byte range in the complete source.
    #[must_use]
    pub fn range(&self) -> Range<usize> {
        self.record.range.clone()
    }

    /// Exact inline source including delimiters and escapes.
    #[must_use]
    pub fn source(&self) -> &'snapshot str {
        &self.source[self.record.range.clone()]
    }
}

/// Iterator over inline nodes in parser preorder.
#[derive(Clone, Debug)]
pub struct Inlines<'snapshot> {
    source: &'snapshot str,
    records: std::slice::Iter<'snapshot, InlineRecord>,
}

impl ExactSizeIterator for Inlines<'_> {}

impl<'snapshot> Iterator for Inlines<'snapshot> {
    type Item = Inline<'snapshot>;

    fn next(&mut self) -> Option<Self::Item> {
        self.records.next().map(|record| Inline {
            source: self.source,
            record,
        })
    }

    fn size_hint(&self) -> (usize, Option<usize>) {
        self.records.size_hint()
    }
}

/// A borrowed reference-graph entry.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct Reference<'snapshot> {
    source: &'snapshot str,
    record: &'snapshot ReferenceRecord,
}

impl Reference<'_> {
    /// Resolved destination for links/images, if any.
    #[must_use]
    pub fn destination(&self) -> Option<&str> {
        self.record.destination.as_deref()
    }

    /// Normalized reference or footnote label, if any.
    #[must_use]
    pub fn label(&self) -> Option<&str> {
        self.record.label.as_deref()
    }

    /// Reference role.
    #[must_use]
    pub const fn kind(&self) -> ReferenceKind {
        self.record.kind
    }

    /// Exact source range of this use or definition.
    #[must_use]
    pub fn range(&self) -> Range<usize> {
        self.record.range.clone()
    }

    /// Exact Markdown source represented by this reference use or definition.
    #[must_use]
    pub fn source(&self) -> &str {
        &self.source[self.record.range.clone()]
    }

    /// Optional interpreted source title.
    #[must_use]
    pub fn title(&self) -> Option<&str> {
        self.record.title.as_deref()
    }
}

/// Iterator over the source-ordered reference graph.
#[derive(Clone, Debug)]
pub struct References<'snapshot> {
    source: &'snapshot str,
    records: std::slice::Iter<'snapshot, ReferenceRecord>,
}

impl ExactSizeIterator for References<'_> {}

impl<'snapshot> Iterator for References<'snapshot> {
    type Item = Reference<'snapshot>;

    fn next(&mut self) -> Option<Self::Item> {
        self.records.next().map(|record| Reference {
            source: self.source,
            record,
        })
    }

    fn size_hint(&self) -> (usize, Option<usize>) {
        self.records.size_hint()
    }
}

/// Iterator over top-level Markdown blocks in source order.
#[derive(Clone, Debug)]
pub struct Blocks<'snapshot> {
    source: &'snapshot str,
    records: std::slice::Iter<'snapshot, BlockRecord>,
}

impl DoubleEndedIterator for Blocks<'_> {
    fn next_back(&mut self) -> Option<Self::Item> {
        self.records.next_back().map(|record| Block {
            source: self.source,
            record,
        })
    }
}

impl Blocks<'_> {
    /// Number of blocks not yet yielded.
    #[must_use]
    pub fn len(&self) -> usize {
        self.records.len()
    }

    /// Whether no blocks remain.
    #[must_use]
    pub fn is_empty(&self) -> bool {
        self.records.len() == 0
    }
}

impl ExactSizeIterator for Blocks<'_> {}

impl<'snapshot> Iterator for Blocks<'snapshot> {
    type Item = Block<'snapshot>;

    fn next(&mut self) -> Option<Self::Item> {
        self.records.next().map(|record| Block {
            source: self.source,
            record,
        })
    }

    fn size_hint(&self) -> (usize, Option<usize>) {
        self.records.size_hint()
    }
}

/// An immutable, cheap-to-clone exact Markdown snapshot.
#[derive(Clone)]
pub struct Snapshot {
    pub(crate) state: Arc<State>,
}

impl Snapshot {
    /// Apply a reversible patch only to its exact immutable before-image.
    ///
    /// # Errors
    ///
    /// Returns [`Error::PatchConflict`] for any different source, dialect, or
    /// read policy. The patch target is fully reparsed before publication.
    pub fn apply(&self, patch: &Patch) -> Result<super::Commit, Error> {
        super::transaction::apply(self, patch)
    }

    /// Borrow a block by zero-based source position.
    #[must_use]
    pub fn block(&self, position: usize) -> Option<Block<'_>> {
        self.state.blocks.get(position).map(|record| Block {
            source: &self.state.source,
            record,
        })
    }

    /// Iterate over top-level blocks in source order.
    #[must_use]
    pub fn blocks(&self) -> Blocks<'_> {
        Blocks {
            source: &self.state.source,
            records: self.state.blocks.iter(),
        }
    }

    /// Grammar used to interpret this snapshot.
    #[must_use]
    pub fn dialect(&self) -> Dialect {
        self.state.dialect
    }

    /// Start a bounded edit against this immutable snapshot.
    #[must_use]
    pub const fn edit(&self) -> Edit<'_> {
        Edit::new(self)
    }

    /// Read a `CommonMark` snapshot with safe default limits.
    ///
    /// # Errors
    ///
    /// Returns a typed UTF-8, input, allocation, or resource-limit error.
    pub fn read(source: &str) -> Result<Self, Error> {
        Self::read_with(source, Dialect::CommonMark, ReadLimits::DEFAULT)
    }

    /// Read UTF-8 bytes as `CommonMark` with safe default limits.
    ///
    /// # Errors
    ///
    /// Returns [`Error::Utf8`] for a non-UTF-8 byte sequence, plus the errors
    /// documented by [`Self::read`].
    pub fn read_bytes(source: &[u8]) -> Result<Self, Error> {
        let utf8_source = std::str::from_utf8(source).map_err(Error::Utf8)?;
        Self::read(utf8_source)
    }

    /// Read an exact source with an explicit grammar and resource policy.
    ///
    /// # Errors
    ///
    /// Returns a typed UTF-8 input, allocation, or resource-limit error. NUL is
    /// refused because `CommonMark` requires it to be replaced with U+FFFD;
    /// silently doing so would violate exact-source semantics.
    pub fn read_with(source: &str, dialect: Dialect, limits: ReadLimits) -> Result<Self, Error> {
        parse::read(source, dialect, limits)
    }

    /// Exact complete UTF-8 Markdown source.
    #[must_use]
    pub fn source(&self) -> &str {
        &self.state.source
    }

    /// Limits retained for future edits and patch application.
    #[must_use]
    pub fn limits(&self) -> ReadLimits {
        self.state.limits
    }

    /// Iterate over resolved uses and definitions in deterministic source order.
    #[must_use]
    pub fn references(&self) -> References<'_> {
        References {
            source: &self.state.source,
            records: self.state.references.iter(),
        }
    }

    /// Check which source-ranged features cannot be represented by a target.
    ///
    /// This is a read-only preflight; it never renders or mutates the snapshot.
    ///
    /// # Errors
    ///
    /// Returns an allocation error if the bounded issue report cannot be built.
    pub fn preflight_projection(
        &self,
        capabilities: ProjectionCapabilities,
    ) -> Result<ProjectionPreflight, Error> {
        let mut issues = Vec::new();
        for block in self.blocks() {
            if block.kind() == BlockKind::Table && !capabilities.tables {
                push_projection_issue(&mut issues, ProjectionIssueKind::Table, block.range())?;
            }
            if block.kind() == BlockKind::FootnoteDefinition && !capabilities.footnotes {
                push_projection_issue(&mut issues, ProjectionIssueKind::Footnote, block.range())?;
            }
            if matches!(block.kind(), BlockKind::Html) && !capabilities.raw_html {
                push_projection_issue(&mut issues, ProjectionIssueKind::RawHtml, block.range())?;
            }
            for inline in block.inlines() {
                let issue_kind = match inline.kind() {
                    InlineKind::Html if !capabilities.raw_html => {
                        Some(ProjectionIssueKind::RawHtml)
                    },
                    InlineKind::TaskListMarker { .. } if !capabilities.task_lists => {
                        Some(ProjectionIssueKind::TaskList)
                    },
                    InlineKind::FootnoteReference if !capabilities.footnotes => {
                        Some(ProjectionIssueKind::Footnote)
                    },
                    InlineKind::Text
                    | InlineKind::Emphasis
                    | InlineKind::Strong
                    | InlineKind::Strikethrough
                    | InlineKind::Code
                    | InlineKind::Link
                    | InlineKind::Image
                    | InlineKind::Html
                    | InlineKind::FootnoteReference
                    | InlineKind::SoftBreak
                    | InlineKind::HardBreak
                    | InlineKind::TaskListMarker { .. } => None,
                };
                if let Some(kind) = issue_kind {
                    push_projection_issue(&mut issues, kind, inline.range())?;
                }
            }
        }
        Ok(ProjectionPreflight {
            issues: issues.into_boxed_slice(),
        })
    }

    /// Build and fully validate a reference-aware append into `destination`.
    ///
    /// Link and footnote definitions required by the selected block are
    /// included recursively when absent from the destination. Neither snapshot
    /// is mutated by this preflight.
    ///
    /// # Errors
    ///
    /// Returns a typed missing-block, dialect, dependency-conflict, allocation,
    /// limit, or candidate-validation error.
    pub fn preflight_transfer_block(
        &self,
        position: usize,
        destination: &Self,
    ) -> Result<super::TransferPlan, Error> {
        super::transaction::preflight_transfer_block(self, position, destination)
    }
}

/// Features supported by a projection target.
#[derive(Clone, Copy, Debug, Default, PartialEq, Eq)]
#[expect(
    clippy::struct_excessive_bools,
    reason = "the public projection contract exposes four independently selectable features"
)]
pub struct ProjectionCapabilities {
    /// Preserve raw block and inline HTML.
    pub raw_html: bool,
    /// Preserve GFM tables.
    pub tables: bool,
    /// Preserve GFM task-list markers.
    pub task_lists: bool,
    /// Preserve footnote uses and definitions.
    pub footnotes: bool,
}

impl ProjectionCapabilities {
    /// A target capable of retaining every feature checked by the preflight.
    pub const LOSSLESS: Self = Self {
        raw_html: true,
        tables: true,
        task_lists: true,
        footnotes: true,
    };
}

/// A feature that a projection target declared unsupported.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
#[non_exhaustive]
pub enum ProjectionIssueKind {
    /// Raw block or inline HTML.
    RawHtml,
    /// A GFM pipe table.
    Table,
    /// A GFM task-list marker.
    TaskList,
    /// A footnote use or definition.
    Footnote,
}

/// One exact source-ranged projection issue.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct ProjectionIssue {
    kind: ProjectionIssueKind,
    range: Range<usize>,
}

impl ProjectionIssue {
    /// Unsupported feature class.
    #[must_use]
    pub const fn kind(&self) -> ProjectionIssueKind {
        self.kind
    }

    /// Exact range requiring target-specific handling.
    #[must_use]
    pub fn range(&self) -> Range<usize> {
        self.range.clone()
    }
}

/// Complete, non-mutating projection preflight result.
#[derive(Clone, Debug, Default, PartialEq, Eq)]
pub struct ProjectionPreflight {
    issues: Box<[ProjectionIssue]>,
}

impl ProjectionPreflight {
    /// Source-ordered unsupported features.
    #[must_use]
    pub const fn issues(&self) -> &[ProjectionIssue] {
        &self.issues
    }

    /// Whether the declared target can preserve all checked features.
    #[must_use]
    pub fn is_lossless(&self) -> bool {
        self.issues.is_empty()
    }
}

impl fmt::Debug for Snapshot {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("Snapshot")
            .field("source_bytes", &self.state.source.len())
            .field("blocks", &self.state.blocks.len())
            .field("dialect", &self.state.dialect)
            .finish()
    }
}

impl Eq for Snapshot {}

impl PartialEq for Snapshot {
    fn eq(&self, other: &Self) -> bool {
        self.state.dialect == other.state.dialect && self.state.source == other.state.source
    }
}

/// A typed Markdown reader or editor refusal.
#[derive(Debug, Error)]
#[non_exhaustive]
pub enum Error {
    /// A read-limit field was configured as zero.
    #[error("Markdown read limit '{name}' must be nonzero")]
    InvalidLimit {
        /// Stable field name.
        name: &'static str,
    },
    /// The complete source is larger than policy permits.
    #[error("Markdown source has {actual} bytes; limit is {limit}")]
    SourceTooLarge {
        /// Observed source bytes.
        actual: usize,
        /// Configured maximum.
        limit: usize,
    },
    /// One physical source line is larger than policy permits.
    #[error("Markdown line {line} has {actual} bytes; limit is {limit}")]
    LineTooLong {
        /// One-based physical line number.
        line: usize,
        /// Observed bytes excluding CR/LF.
        actual: usize,
        /// Configured maximum.
        limit: usize,
    },
    /// A NUL would require lossy `CommonMark` preprocessing.
    #[error("Markdown source contains NUL at byte {offset}")]
    NullByte {
        /// Exact offending byte offset.
        offset: usize,
    },
    /// Parsing produced too many events.
    #[error("Markdown parser event count exceeds limit {limit}")]
    EventLimitExceeded {
        /// Configured maximum.
        limit: usize,
    },
    /// Parsing produced too many top-level blocks.
    #[error("Markdown block count exceeds limit {limit}")]
    BlockLimitExceeded {
        /// Configured maximum.
        limit: usize,
    },
    /// Container nesting exceeds policy.
    #[error("Markdown nesting depth exceeds limit {limit} at byte {offset}")]
    NestingLimitExceeded {
        /// Configured maximum.
        limit: usize,
        /// Source byte offset at which the limit was exceeded.
        offset: usize,
    },
    /// A selected top-level block does not exist.
    #[error("Markdown snapshot has no block at position {position}")]
    BlockNotFound {
        /// Requested zero-based position.
        position: usize,
    },
    /// A selected nested block does not exist.
    #[error("Markdown block {block_position} has no nested block at position {nested_position}")]
    NestedBlockNotFound {
        /// Top-level block position.
        block_position: usize,
        /// Nested block position in parser preorder.
        nested_position: usize,
    },
    /// A selected inline node does not exist.
    #[error("Markdown block {block_position} has no inline at position {inline_position}")]
    InlineNotFound {
        /// Top-level block position.
        block_position: usize,
        /// Inline position in parser preorder.
        inline_position: usize,
    },
    /// A nested edit escaped or changed its enclosing top-level block.
    #[error("Markdown nested edit changed its top-level structural boundary")]
    StructuralBoundaryChanged,
    /// A transfer was requested between different Markdown dialects.
    #[error("Markdown transfer requires matching source and destination dialects")]
    TransferDialectMismatch,
    /// A destination defines a transferred dependency differently.
    #[error("Markdown destination has a conflicting definition for '{label}'")]
    TransferDependencyConflict {
        /// Normalized link or footnote label.
        label: String,
    },
    /// A block replacement or append was not exactly one parsed block.
    #[error("Markdown replacement must contain exactly one top-level block; found {actual}")]
    ReplacementBlockCount {
        /// Parsed replacement block count.
        actual: usize,
    },
    /// The bounded edit already contains an operation.
    #[error("Markdown edit already has a staged operation")]
    OperationAlreadyStaged,
    /// Two operations target the same immutable base block.
    #[error("Markdown edit already targets block position {position}")]
    OverlappingOperation {
        /// Conflicting zero-based base position.
        position: usize,
    },
    /// The transaction would exceed its operation budget.
    #[error("Markdown edit operation count exceeds limit {limit}")]
    OperationLimitExceeded {
        /// Configured maximum.
        limit: usize,
    },
    /// A referenced definition would be removed while a use remains.
    #[error("Markdown definition '{label}' remains referenced outside the edit closure")]
    ReferenceDependency {
        /// Normalized link or footnote label.
        label: String,
    },
    /// A bounded history cannot retain another commit.
    #[error("Markdown history {resource} exceeds limit {limit}")]
    HistoryLimitExceeded {
        /// Stable bounded resource.
        resource: &'static str,
        /// Configured maximum.
        limit: usize,
    },
    /// A durable patch envelope exceeds its caller-selected byte limit.
    #[error("Markdown patch envelope has {actual} bytes; limit is {limit}")]
    PatchEnvelopeTooLarge {
        /// Observed bytes.
        actual: usize,
        /// Configured maximum.
        limit: usize,
    },
    /// A durable patch envelope is malformed or unsupported.
    #[error("invalid Markdown patch envelope: {reason}")]
    InvalidPatchEnvelope {
        /// Stable parse or version failure.
        reason: String,
    },
    /// Commit was requested without staging an operation.
    #[error("Markdown edit has no staged operation")]
    NoStagedOperation,
    /// A patch was applied to a different exact snapshot.
    #[error("Markdown patch does not match the exact source snapshot")]
    PatchConflict,
    /// Input bytes were not valid UTF-8.
    #[error("Markdown source is not UTF-8: {0}")]
    Utf8(Utf8Error),
    /// A required allocation failed.
    #[error("failed to allocate {resource}: {source}")]
    Allocation {
        /// Stable resource description.
        resource: &'static str,
        /// Allocation failure reported by the standard library.
        #[source]
        source: std::collections::TryReserveError,
    },
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub(crate) struct BlockRecord {
    pub(crate) kind: BlockKind,
    pub(crate) range: Range<usize>,
    pub(crate) inlines: Box<[InlineRecord]>,
    pub(crate) descendants: Box<[NestedBlockRecord]>,
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub(crate) struct NestedBlockRecord {
    pub(crate) kind: BlockKind,
    pub(crate) range: Range<usize>,
    pub(crate) depth: usize,
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub(crate) struct InlineRecord {
    pub(crate) kind: InlineKind,
    pub(crate) range: Range<usize>,
    pub(crate) depth: usize,
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub(crate) struct ReferenceRecord {
    pub(crate) kind: ReferenceKind,
    pub(crate) range: Range<usize>,
    pub(crate) label: Option<String>,
    pub(crate) destination: Option<String>,
    pub(crate) title: Option<String>,
}

#[derive(Debug)]
pub(crate) struct State {
    pub(crate) source: Arc<str>,
    pub(crate) blocks: Box<[BlockRecord]>,
    pub(crate) references: Box<[ReferenceRecord]>,
    pub(crate) dialect: Dialect,
    pub(crate) limits: ReadLimits,
}

fn push_projection_issue(
    issues: &mut Vec<ProjectionIssue>,
    kind: ProjectionIssueKind,
    range: Range<usize>,
) -> Result<(), Error> {
    issues.try_reserve(1).map_err(|source| Error::Allocation {
        resource: "Markdown projection preflight",
        source,
    })?;
    issues.push(ProjectionIssue { kind, range });
    Ok(())
}
