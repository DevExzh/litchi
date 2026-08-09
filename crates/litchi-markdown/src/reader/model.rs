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
}

impl ReadLimits {
    /// Safe defaults for an ordinary Markdown document.
    pub const DEFAULT: Self = Self {
        max_source_bytes: 16 * 1024 * 1024,
        max_line_bytes: 1024 * 1024,
        max_events: 1_000_000,
        max_blocks: 100_000,
        max_nesting_depth: 256,
    };

    pub(crate) fn validate(self) -> Result<(), Error> {
        for (name, value) in [
            ("max_source_bytes", self.max_source_bytes),
            ("max_line_bytes", self.max_line_bytes),
            ("max_events", self.max_events),
            ("max_blocks", self.max_blocks),
            ("max_nesting_depth", self.max_nesting_depth),
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
    /// A GFM footnote definition.
    FootnoteDefinition,
    /// A GFM pipe table.
    Table,
    /// A thematic break.
    ThematicBreak,
    /// A `CommonMark` link reference definition.
    LinkDefinition,
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
    /// A block replacement or append was not exactly one parsed block.
    #[error("Markdown replacement must contain exactly one top-level block; found {actual}")]
    ReplacementBlockCount {
        /// Parsed replacement block count.
        actual: usize,
    },
    /// The bounded edit already contains an operation.
    #[error("Markdown edit already has a staged operation")]
    OperationAlreadyStaged,
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
}

#[derive(Debug)]
pub(crate) struct State {
    pub(crate) source: Arc<str>,
    pub(crate) blocks: Box<[BlockRecord]>,
    pub(crate) dialect: Dialect,
    pub(crate) limits: ReadLimits,
}
