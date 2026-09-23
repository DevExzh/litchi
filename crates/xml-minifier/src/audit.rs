//! Bounded verification of the repository's compact XML output contract.
//!
//! Character data and CDATA are never normalized. Plain spaces inside the
//! document element remain content; whitespace outside the document element,
//! and whitespace-only nodes containing CR, LF, or tab outside an inherited
//! `xml:space="preserve"` scope, are classified as structural formatting.

#![allow(
    clippy::arbitrary_source_item_ordering,
    reason = "semantic API types precede their streaming implementation and package submodule"
)]

use core::{fmt, mem::size_of};
use quick_xml::Reader;
use quick_xml::XmlVersion;
use quick_xml::encoding::Decoder;
use quick_xml::events::{BytesStart, Event};
use std::io::{self, BufRead, Read};

/// Finite resource budgets for one XML document.
#[derive(Clone, Copy, Debug, Eq, PartialEq)]
pub struct Limits {
    attributes: usize,
    bytes: usize,
    depth: usize,
    events: usize,
    text_bytes: usize,
    token_bytes: usize,
}

impl Limits {
    /// Hard ceiling for aggregate attributes.
    pub const ATTRIBUTE_CEILING: usize = 1_000_000;
    /// Hard ceiling for one XML document in bytes.
    pub const BYTE_CEILING: usize = 256 * 1024 * 1024;
    /// Hard ceiling for element nesting.
    pub const DEPTH_CEILING: usize = 4_096;
    /// Hard ceiling for parser events.
    pub const EVENT_CEILING: usize = 4_000_000;
    /// Hard ceiling for aggregate character-data bytes.
    pub const TEXT_BYTE_CEILING: usize = 256 * 1024 * 1024;
    /// Hard ceiling for one lexical token in bytes.
    pub const TOKEN_BYTE_CEILING: usize = 64 * 1024 * 1024;

    /// Creates an explicit limit profile.
    ///
    /// # Errors
    ///
    /// Returns [`ConfigError`] when any requested value exceeds its immutable
    /// hard ceiling.
    pub fn new(
        max_bytes: usize,
        max_depth: usize,
        max_events: usize,
        max_attributes: usize,
        max_token_bytes: usize,
        max_text_bytes: usize,
    ) -> Result<Self, ConfigError> {
        let limits = Self {
            attributes: max_attributes,
            bytes: max_bytes,
            depth: max_depth,
            events: max_events,
            text_bytes: max_text_bytes,
            token_bytes: max_token_bytes,
        };
        limits.check()?;
        Ok(limits)
    }

    /// Starts a safe fallible limit builder from the default profile.
    #[must_use]
    pub fn builder() -> Builder {
        Builder::default()
    }

    /// Returns the immutable hard ceiling for `resource`.
    #[must_use]
    pub const fn ceiling(resource: Resource) -> usize {
        match resource {
            Resource::Attributes => Self::ATTRIBUTE_CEILING,
            Resource::Bytes => Self::BYTE_CEILING,
            Resource::Depth => Self::DEPTH_CEILING,
            Resource::Events => Self::EVENT_CEILING,
            Resource::TextBytes => Self::TEXT_BYTE_CEILING,
            Resource::TokenBytes => Self::TOKEN_BYTE_CEILING,
        }
    }

    /// Narrows one resource without permitting an increase.
    #[must_use]
    pub const fn narrow(mut self, resource: Resource, maximum: usize) -> Self {
        match resource {
            Resource::Attributes => self.attributes = minimum(self.attributes, maximum),
            Resource::Bytes => self.bytes = minimum(self.bytes, maximum),
            Resource::Depth => self.depth = minimum(self.depth, maximum),
            Resource::Events => self.events = minimum(self.events, maximum),
            Resource::TextBytes => self.text_bytes = minimum(self.text_bytes, maximum),
            Resource::TokenBytes => self.token_bytes = minimum(self.token_bytes, maximum),
        }
        self
    }

    const fn bounded(
        bytes: usize,
        depth: usize,
        events: usize,
        attributes: usize,
        token_bytes: usize,
        text_bytes: usize,
    ) -> Self {
        Self {
            attributes,
            bytes,
            depth,
            events,
            text_bytes,
            token_bytes,
        }
    }

    fn check(self) -> Result<(), ConfigError> {
        for resource in [
            Resource::Bytes,
            Resource::Depth,
            Resource::Events,
            Resource::Attributes,
            Resource::TokenBytes,
            Resource::TextBytes,
        ] {
            let requested = self.value(resource);
            let ceiling = Self::ceiling(resource);
            if requested > ceiling {
                return Err(ConfigError {
                    ceiling,
                    requested,
                    resource,
                });
            }
        }
        Ok(())
    }

    const fn value(self, resource: Resource) -> usize {
        match resource {
            Resource::Attributes => self.attributes,
            Resource::Bytes => self.bytes,
            Resource::Depth => self.depth,
            Resource::Events => self.events,
            Resource::TextBytes => self.text_bytes,
            Resource::TokenBytes => self.token_bytes,
        }
    }

    /// Maximum number of attributes in the document.
    #[must_use]
    pub const fn max_attributes(self) -> usize {
        self.attributes
    }

    /// Maximum input size in bytes.
    #[must_use]
    pub const fn max_bytes(self) -> usize {
        self.bytes
    }

    /// Maximum element nesting depth.
    #[must_use]
    pub const fn max_depth(self) -> usize {
        self.depth
    }

    /// Maximum number of parser events.
    #[must_use]
    pub const fn max_events(self) -> usize {
        self.events
    }

    /// Maximum aggregate character-data bytes.
    #[must_use]
    pub const fn max_text_bytes(self) -> usize {
        self.text_bytes
    }

    /// Maximum bytes in one lexical token.
    #[must_use]
    pub const fn max_token_bytes(self) -> usize {
        self.token_bytes
    }

    /// Returns a checked upper bound for the streaming auditor's dynamic
    /// parser buffers.
    ///
    /// The bound includes the reusable event buffer (`max_token_bytes + 1`),
    /// quick-xml's retained open-element names (at most one token-sized name
    /// per admitted depth, with geometric `Vec` capacity), its open-name
    /// indexes, the bounded lexical capture and exposed source window, the
    /// conservative spare token window, the geometric attribute-name
    /// `Range<usize>` tracker and large-tag hash prefilter (including a
    /// fourfold capacity/control-byte factor and an eight-entry minimum), two
    /// transient token-sized decoded and normalized attribute-value buffers, and this
    /// auditor's inherited `xml:space` stack. Counting the reusable event
    /// buffer, these dynamic token windows are bounded by six token capacities.
    /// Attribute duplicate-check scratch is per lexical event, rather than an
    /// aggregate document allocation. The parser cannot expose more attribute
    /// entries in one event than the bounded token window, so its range and hash
    /// capacities use `min(max_attributes, max_token_bytes + 1)`. The aggregate
    /// attribute counter still uses `max_attributes` and therefore retains its
    /// document-wide acceptance policy.
    /// The three-byte BOM probe fits the conservative 12-byte fixed allowance
    /// retained from the earlier streaming auditor. The caller's
    /// `BufRead` storage, other fixed-size parser values, allocator metadata,
    /// and error strings are outside this bound.
    /// `None` means the checked arithmetic could not represent the envelope in
    /// `usize`. This is a size envelope; it does not make quick-xml's internal
    /// allocations fallible.
    #[must_use]
    pub fn streaming_memory_upper_bound(self) -> Option<usize> {
        let token = self.token_bytes.checked_add(1)?;
        let levels = self.depth.checked_add(1)?;
        let event_and_attribute_scratch = token.checked_mul(6)?;
        let open_names = levels.checked_mul(token)?.checked_mul(2)?.max(8);
        let open_indexes = self
            .depth
            .checked_add(1)?
            .checked_mul(size_of::<usize>())?
            .checked_mul(2)?
            .max(4 * size_of::<usize>());
        let spaces = self
            .depth
            .checked_add(1)?
            .checked_mul(size_of::<Space>())?
            .checked_mul(2)?
            .max(8 * size_of::<Space>());
        let per_event_attributes = self.attributes.min(token);
        let attribute_entries = per_event_attributes.checked_add(1)?;
        let attribute_range_size = size_of::<std::ops::Range<usize>>();
        let attribute_ranges = attribute_entries
            .checked_mul(attribute_range_size)?
            .checked_mul(2)?
            .max(4 * attribute_range_size);
        let attribute_hash_entry = size_of::<u64>().checked_add(size_of::<u8>())?;
        let attribute_hashes = attribute_entries
            .checked_mul(attribute_hash_entry)?
            .checked_mul(4)?
            .max(attribute_hash_entry.checked_mul(8)?);
        event_and_attribute_scratch
            .checked_add(open_names)?
            .checked_add(open_indexes)?
            .checked_add(spaces)
            .and_then(|total| total.checked_add(attribute_ranges))
            .and_then(|total| total.checked_add(attribute_hashes))
            .and_then(|total| total.checked_add(12))
    }
}

impl Default for Limits {
    fn default() -> Self {
        Self::bounded(
            32 * 1024 * 1024,
            256,
            1_000_000,
            250_000,
            4 * 1024 * 1024,
            16 * 1024 * 1024,
        )
    }
}

/// Fallible construction of one bounded [`Limits`] profile.
#[derive(Clone, Copy, Debug, Default, Eq, PartialEq)]
pub struct Builder {
    limits: Limits,
}

impl Builder {
    /// Sets the aggregate attribute limit.
    ///
    /// # Errors
    ///
    /// Returns [`ConfigError`] when `maximum` exceeds the immutable ceiling.
    pub fn attributes(self, maximum: usize) -> Result<Self, ConfigError> {
        self.setting(Resource::Attributes, maximum)
    }

    /// Builds the checked profile without allocation.
    #[must_use]
    pub const fn build(self) -> Limits {
        self.limits
    }

    /// Sets the input-byte limit.
    ///
    /// # Errors
    ///
    /// Returns [`ConfigError`] when `maximum` exceeds the immutable ceiling.
    pub fn bytes(self, maximum: usize) -> Result<Self, ConfigError> {
        self.setting(Resource::Bytes, maximum)
    }

    /// Sets the element-depth limit.
    ///
    /// # Errors
    ///
    /// Returns [`ConfigError`] when `maximum` exceeds the immutable ceiling.
    pub fn depth(self, maximum: usize) -> Result<Self, ConfigError> {
        self.setting(Resource::Depth, maximum)
    }

    /// Sets the parser-event limit.
    ///
    /// # Errors
    ///
    /// Returns [`ConfigError`] when `maximum` exceeds the immutable ceiling.
    pub fn events(self, maximum: usize) -> Result<Self, ConfigError> {
        self.setting(Resource::Events, maximum)
    }

    /// Sets one typed resource limit.
    ///
    /// # Errors
    ///
    /// Returns [`ConfigError`] when `maximum` exceeds the resource's immutable
    /// ceiling.
    pub fn limit(self, resource: Resource, maximum: usize) -> Result<Self, ConfigError> {
        self.setting(resource, maximum)
    }

    /// Sets the aggregate character-data limit.
    ///
    /// # Errors
    ///
    /// Returns [`ConfigError`] when `maximum` exceeds the immutable ceiling.
    pub fn text_bytes(self, maximum: usize) -> Result<Self, ConfigError> {
        self.setting(Resource::TextBytes, maximum)
    }

    /// Sets the single-token byte limit.
    ///
    /// # Errors
    ///
    /// Returns [`ConfigError`] when `maximum` exceeds the immutable ceiling.
    pub fn token_bytes(self, maximum: usize) -> Result<Self, ConfigError> {
        self.setting(Resource::TokenBytes, maximum)
    }

    fn setting(mut self, resource: Resource, maximum: usize) -> Result<Self, ConfigError> {
        let ceiling = Limits::ceiling(resource);
        if maximum > ceiling {
            return Err(ConfigError {
                ceiling,
                requested: maximum,
                resource,
            });
        }
        match resource {
            Resource::Attributes => self.limits.attributes = maximum,
            Resource::Bytes => self.limits.bytes = maximum,
            Resource::Depth => self.limits.depth = maximum,
            Resource::Events => self.limits.events = maximum,
            Resource::TextBytes => self.limits.text_bytes = maximum,
            Resource::TokenBytes => self.limits.token_bytes = maximum,
        }
        Ok(self)
    }
}

/// Invalid limit configuration above an immutable hard ceiling.
#[derive(Clone, Copy, Debug, Eq, PartialEq)]
pub struct ConfigError {
    ceiling: usize,
    requested: usize,
    resource: Resource,
}

impl ConfigError {
    /// Immutable ceiling that was exceeded.
    #[must_use]
    pub const fn ceiling(self) -> usize {
        self.ceiling
    }

    /// Requested value.
    #[must_use]
    pub const fn requested(self) -> usize {
        self.requested
    }

    /// Resource whose ceiling was exceeded.
    #[must_use]
    pub const fn resource(self) -> Resource {
        self.resource
    }
}

impl fmt::Display for ConfigError {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        write!(
            formatter,
            "XML {:?} limit {} exceeds hard ceiling {}",
            self.resource, self.requested, self.ceiling
        )
    }
}

impl std::error::Error for ConfigError {}

/// A resource governed by [`Limits`].
#[derive(Clone, Copy, Debug, Eq, PartialEq)]
#[non_exhaustive]
pub enum Resource {
    /// Total attribute count.
    Attributes,
    /// Input byte length.
    Bytes,
    /// Element nesting depth.
    Depth,
    /// Parser event count.
    Events,
    /// Aggregate character-data bytes.
    TextBytes,
    /// One lexical token's byte length.
    TokenBytes,
}

/// A lexically provable compactness defect.
#[derive(Clone, Copy, Debug, Eq, PartialEq)]
#[non_exhaustive]
pub enum Kind {
    /// Provably structural whitespace outside `xml:space="preserve"`.
    FormattingWhitespace,
    /// A whitespace-only text node contains only spaces, so a schema-neutral
    /// auditor cannot prove whether it is content or indentation.
    AmbiguousWhitespace,
    /// Attribute boundaries do not use exactly one ASCII space.
    AttributeSeparation,
    /// Whitespace occurs immediately before `>`, `/>`, or an end-tag close.
    WhitespaceBeforeClose,
}

/// Location and category of a compactness defect.
#[derive(Clone, Copy, Debug, Eq, PartialEq)]
pub struct Violation {
    kind: Kind,
    offset: usize,
}

impl Violation {
    /// The stable defect category.
    #[must_use]
    pub const fn kind(self) -> Kind {
        self.kind
    }

    /// Zero-based byte offset in the original XML.
    #[must_use]
    pub const fn offset(self) -> usize {
        self.offset
    }
}

/// Failure to parse or verify a compact XML document.
#[derive(Debug, Eq, PartialEq)]
#[non_exhaustive]
pub enum Error {
    /// A finite audit budget was exceeded.
    Limit {
        /// Governed resource.
        resource: Resource,
        /// Configured inclusive limit.
        limit: usize,
        /// First observed value beyond the limit.
        actual: usize,
        /// Byte offset at which accounting failed.
        offset: usize,
    },
    /// The input was not UTF-8 XML.
    Encoding {
        /// First invalid UTF-8 byte.
        valid_up_to: usize,
    },
    /// XML parsing or document-structure validation failed.
    Malformed {
        /// Parser byte offset.
        offset: usize,
        /// Bounded-by-input parser diagnostic.
        detail: String,
    },
    /// XML is valid but violates the compact output contract.
    NotCompact(Violation),
    /// DTD and DOCTYPE declarations are ineligible for compact package XML.
    Doctype {
        /// Zero-based byte offset of the declaration.
        offset: usize,
    },
    /// A bounded audit buffer could not reserve its configured capacity.
    Allocation,
}

impl Error {
    fn malformed(offset: usize, detail: impl Into<String>) -> Self {
        Self::Malformed {
            offset,
            detail: detail.into(),
        }
    }
}

impl fmt::Display for Error {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Self::Limit {
                resource,
                limit,
                actual,
                offset,
            } => write!(
                formatter,
                "XML {resource:?} limit {limit} exceeded by {actual} at byte {offset}"
            ),
            Self::Encoding { valid_up_to } => {
                write!(formatter, "XML is not UTF-8 at byte {valid_up_to}")
            },
            Self::Malformed { offset, detail } => {
                write!(formatter, "malformed XML at byte {offset}: {detail}")
            },
            Self::NotCompact(violation) => write!(
                formatter,
                "noncompact XML {:?} at byte {}",
                violation.kind, violation.offset
            ),
            Self::Doctype { offset } => {
                write!(
                    formatter,
                    "DTD and DOCTYPE are not allowed at byte {offset}"
                )
            },
            Self::Allocation => {
                formatter.write_str("could not reserve the bounded XML depth stack")
            },
        }
    }
}

impl std::error::Error for Error {}

/// Failure returned by the streaming XML auditor.
///
/// A source failure remains an [`Input`](StreamError::Input) error so callers
/// can distinguish a
/// transport or decompression failure from XML validation. Resource windows,
/// parser failures, encoding failures, and compactness defects are returned as
/// [`Audit`](StreamError::Audit) errors. Streaming validation reports the first
/// failure observed
/// while consuming the source; this can differ from the error ordering of the
/// slice API when a later source byte has not yet been read.
#[derive(Debug)]
#[non_exhaustive]
pub enum StreamError {
    /// The supplied source returned an I/O failure.
    Input(io::Error),
    /// XML parsing, compactness, encoding, or finite-budget failure.
    Audit(Error),
}

impl fmt::Display for StreamError {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Self::Input(source) => write!(formatter, "XML input stream failed: {source}"),
            Self::Audit(source) => source.fmt(formatter),
        }
    }
}

impl std::error::Error for StreamError {
    fn source(&self) -> Option<&(dyn std::error::Error + 'static)> {
        match self {
            Self::Input(source) => Some(source),
            Self::Audit(source) => Some(source),
        }
    }
}

impl From<Error> for StreamError {
    fn from(source: Error) -> Self {
        Self::Audit(source)
    }
}

/// Accounting summary for a verified compact XML document.
#[derive(Clone, Copy, Debug, Eq, PartialEq)]
#[must_use]
pub struct Report {
    attributes: usize,
    bytes: usize,
    events: usize,
    max_depth: usize,
    text_bytes: usize,
}

impl Report {
    /// Total attributes parsed.
    #[must_use]
    pub const fn attributes(self) -> usize {
        self.attributes
    }

    /// Input bytes parsed.
    #[must_use]
    pub const fn bytes(self) -> usize {
        self.bytes
    }

    /// Parser events, including EOF.
    #[must_use]
    pub const fn events(self) -> usize {
        self.events
    }

    /// Greatest element nesting depth.
    #[must_use]
    pub const fn max_depth(self) -> usize {
        self.max_depth
    }

    /// Aggregate character-data bytes.
    #[must_use]
    pub const fn text_bytes(self) -> usize {
        self.text_bytes
    }
}

#[derive(Clone, Copy, Debug, Eq, PartialEq)]
enum Space {
    Default,
    Preserve,
}

/// Which contracts one audit pass asserts.
///
/// `require_compact` selects whether provable compactness defects are
/// refused; every structural, encoding, DOCTYPE and finite-budget check runs
/// under every policy.
#[derive(Clone, Copy, Debug, Eq, PartialEq)]
struct Policy {
    reject_ambiguous_space: bool,
    require_compact: bool,
}

impl Policy {
    /// Authored bytes: compact, and whitespace-only space runs are ambiguous.
    const AUTHORED: Self = Self {
        reject_ambiguous_space: true,
        require_compact: true,
    };
    /// The historical default: compact, ambiguous space runs admitted.
    const COMPACT: Self = Self {
        reject_ambiguous_space: false,
        require_compact: true,
    };
    /// Producer bytes: structure and budgets only, no compactness contract.
    const SOURCE: Self = Self {
        reject_ambiguous_space: false,
        require_compact: false,
    };
}

struct State {
    ambiguous_space_offset: Option<usize>,
    attributes: usize,
    depth: usize,
    events: usize,
    max_depth: usize,
    roots: usize,
    spaces: Vec<Space>,
    text_bytes: usize,
    text_run_has_explicit_content: bool,
}

impl State {
    fn new() -> Self {
        Self {
            ambiguous_space_offset: None,
            attributes: 0,
            depth: 0,
            events: 0,
            max_depth: 0,
            roots: 0,
            spaces: Vec::new(),
            text_bytes: 0,
            text_run_has_explicit_content: false,
        }
    }

    /// The state inside `depth` open elements of a document whose element
    /// has been seen, with none of the open elements' spaces on the stack:
    /// an end tag that would close one of them is refused as unexpected.
    fn within(depth: usize) -> Self {
        Self {
            depth,
            max_depth: depth,
            roots: 1,
            ..Self::new()
        }
    }

    fn current_space(&self) -> Space {
        self.spaces.last().copied().unwrap_or(Space::Default)
    }
}

/// Parses `input` and rejects the first compactness or resource defect.
///
/// This function is an auditor, not a postprocessor: it never changes input
/// and therefore cannot silently rewrite opaque or mixed-content XML.
///
/// # Errors
///
/// Returns [`Error`] for invalid UTF-8, malformed XML, a finite resource-limit
/// breach, or the first compactness violation.
pub fn verify(input: &[u8], limits: Limits) -> Result<Report, Error> {
    verify_with_policy(input, limits, Policy::COMPACT)
}

/// Verifies XML that is about to be published from authored or changed bytes.
///
/// Unlike [`verify`], this refuses whitespace-only text nodes made entirely of
/// spaces outside `xml:space="preserve"`. Such nodes can be semantic in mixed
/// content, but a schema-neutral publication boundary cannot distinguish them
/// from formatting indentation. Authors that require those spaces must make
/// the preservation intent explicit with `xml:space="preserve"`.
///
/// # Errors
///
/// Returns [`Error`] for every failure reported by [`verify`], or
/// [`Kind::AmbiguousWhitespace`] when authored whitespace cannot be classified.
pub fn verify_authored(input: &[u8], limits: Limits) -> Result<Report, Error> {
    verify_with_policy(input, limits, Policy::AUTHORED)
}

/// Verifies XML a package already holds and that this library did not author.
///
/// This is the publication audit for *source* bytes: the payload a producer
/// wrote, which publication either republishes or replaces. It performs every
/// structural and finite-budget check [`verify`] performs — UTF-8, well-formed
/// XML, exactly one document element, character data only inside it, no DTD or
/// DOCTYPE, and each [`Limits`] budget — and does **not** assert this
/// repository's compact output contract. Indentation and line endings between
/// elements, a line ending after the XML declaration, attribute separators of
/// any length or kind, and whitespace before a tag close are accepted as the
/// producer spelled them, so no [`Error::NotCompact`] is ever returned.
///
/// Compactness is a contract on what this library *emits*; it is not a
/// property of an arbitrary conforming document, and asserting it on bytes the
/// library did not write refuses ordinary producer output. Authored bytes
/// still go through [`verify_authored`].
///
/// # Errors
///
/// Returns [`Error`] for invalid UTF-8, malformed XML, a DTD or DOCTYPE
/// declaration, or a finite resource-limit breach. Never returns
/// [`Error::NotCompact`].
pub fn verify_source(input: &[u8], limits: Limits) -> Result<Report, Error> {
    verify_with_policy(input, limits, Policy::SOURCE)
}

/// How [`verify_source_replacement`] established the replacement's verdict.
///
/// Every variant means the same thing about the replacement: it passed every
/// check [`verify_source`] makes, under the same limits. They differ only in
/// which of its bytes had to be scanned again to prove it, and are reported
/// for diagnostics and tests. They carry byte offsets, never content.
#[derive(Clone, Debug, Eq, PartialEq)]
#[non_exhaustive]
pub enum ReplacementProof {
    /// The replacement is byte-identical to the original, so it has the
    /// original's verdict.
    Identical,
    /// Outside one element of the original, which is not the document
    /// element, the replacement is byte-identical to the original. Only the
    /// replacement bytes that took that element's place were scanned, in the
    /// parser state the original's audit had reached at the element.
    Window {
        /// The replaced element's span in the original.
        original: core::ops::Range<usize>,
        /// The replacement bytes that took its place.
        replacement: core::ops::Range<usize>,
    },
    /// The replacement was scanned completely, as [`verify_source`] scans it.
    Complete,
}

/// The first failure [`verify_source_replacement`] found, and on which side.
#[derive(Debug, Eq, PartialEq)]
#[non_exhaustive]
pub enum ReplacementError {
    /// The original failed with the error [`verify_source`] reports for it;
    /// the replacement was not examined.
    Original(Error),
    /// The original passed and the replacement failed, with the error
    /// [`verify_source`] reports for it.
    Replacement(Error),
}

impl ReplacementError {
    /// The audit failure, from whichever side it came.
    #[must_use]
    pub fn into_error(self) -> Error {
        match self {
            Self::Original(error) | Self::Replacement(error) => error,
        }
    }
}

impl fmt::Display for ReplacementError {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Self::Original(error) => write!(formatter, "original XML: {error}"),
            Self::Replacement(error) => write!(formatter, "replacement XML: {error}"),
        }
    }
}

impl std::error::Error for ReplacementError {
    fn source(&self) -> Option<&(dyn std::error::Error + 'static)> {
        match self {
            Self::Original(error) | Self::Replacement(error) => Some(error),
        }
    }
}

/// Verifies the original bytes of a Part a package already holds and the
/// replacement that will be published in their place, under the source
/// policy.
///
/// The verdict is exactly that of the two audits it replaces, in their order:
/// [`ReplacementError::Original`] carrying the error of
/// `verify_source(original, limits)` when that fails, otherwise
/// [`ReplacementError::Replacement`] carrying the error of
/// `verify_source(replacement, limits)` when that fails, otherwise `Ok`.
///
/// The original is always scanned completely. The replacement usually differs
/// from it only locally, and is then not scanned again where it repeats the
/// original byte for byte: the audit of the original also locates the
/// innermost element, other than the document element, that covers every
/// byte in which the two differ. Outside that element the replacement is the
/// original's own bytes, so they tokenize as the original's did and were
/// checked, in the same parser state, moments earlier. Only the replacement
/// bytes that took the element's place are scanned, starting in the state the
/// original reached there: the same open-element depth, inside the document
/// element, with no enclosing element closable from inside. They must be
/// balanced and must begin and end with markup, so the parse outside them is
/// unchanged; each aggregate budget is then re-totalled for the whole
/// replacement, and the replacement's length is checked against
/// [`Limits::max_bytes`]. UTF-8 validity composes, because the window starts
/// and ends at an ASCII delimiter.
///
/// The shortcut only ever decides that the replacement passes. When the
/// difference is not covered by such an element, the window is larger than
/// half the replacement, or any window check fails, the replacement is scanned
/// completely, so every error is the one [`verify_source`] reports, with its
/// offset. [`ReplacementProof`] says which path was taken.
///
/// # Errors
///
/// Returns [`ReplacementError`] naming the failing side, carrying the
/// [`Error`] that side's [`verify_source`] call returns.
pub fn verify_source_replacement(
    original: &[u8],
    replacement: &[u8],
    limits: Limits,
) -> Result<ReplacementProof, ReplacementError> {
    if original == replacement {
        return verify_source(original, limits)
            .map(|_report| ReplacementProof::Identical)
            .map_err(ReplacementError::Original);
    }
    let prefix = common_prefix_len(original, replacement);
    let suffix = common_suffix_len(&original[prefix..], &replacement[prefix..]);
    let mut search =
        WindowSearch::new(replacement, original.len(), prefix, original.len() - suffix);
    // Every window contains the replacement's differing bytes, and a window
    // larger than half the replacement is never used: when the difference
    // alone is that large, do not look for one.
    search.stopped = (replacement.len() - suffix - prefix).saturating_mul(2) > replacement.len();
    let report = verify_observed(original, limits, Policy::SOURCE, &mut search)
        .map_err(ReplacementError::Original)?;
    if let Some(proof) = search
        .found
        .and_then(|window| window.prove(report, replacement, limits))
    {
        return Ok(proof);
    }
    verify_source(replacement, limits)
        .map(|_report| ReplacementProof::Complete)
        .map_err(ReplacementError::Replacement)
}

/// Verifies one XML document from a caller-owned buffered source.
///
/// The source is consumed incrementally. The reusable parser event buffer is
/// admitted at `max_token_bytes + 1` bytes, and the guarded source refuses to
/// expose another byte before quick-xml can grow that buffer beyond the
/// configured token window. The total input, event count, depth, attributes,
/// and character-data budgets retain the same meanings as [`verify`].
///
/// `quick_xml` retains the names of open elements for end-tag matching. The
/// checked [`Limits::streaming_memory_upper_bound`] helper accounts for that
/// `(depth + 1) * (max_token_bytes + 1)` dynamic envelope in addition to the
/// event buffer and this auditor's `xml:space` stack. The extra admitted name
/// accounts for quick-xml recording a start tag before this auditor refuses an
/// over-limit depth.
///
/// # Errors
///
/// Returns [`StreamError::Input`] for an I/O failure from `reader`, and
/// [`StreamError::Audit`] for XML, compactness, encoding, or finite-budget
/// failures. Unlike [`verify`], UTF-8 and other failures are observed in
/// source order rather than after an upfront whole-input check.
pub fn verify_reader<R: BufRead>(reader: R, limits: Limits) -> Result<Report, StreamError> {
    verify_reader_with_policy(reader, limits, false)
}

/// Verifies authored XML from a caller-owned buffered source.
///
/// This has the same streaming and bounded-memory behavior as
/// [`verify_reader`], while applying the stricter authored rule from
/// [`verify_authored`]: whitespace-only text runs outside
/// `xml:space="preserve"` are rejected as ambiguous.
pub fn verify_authored_reader<R: BufRead>(
    reader: R,
    limits: Limits,
) -> Result<Report, StreamError> {
    verify_reader_with_policy(reader, limits, true)
}

fn verify_reader_with_policy<R: BufRead>(
    reader: R,
    limits: Limits,
    reject_ambiguous_space: bool,
) -> Result<Report, StreamError> {
    // The streaming auditor has no source-policy caller; it keeps the compact
    // contract exactly as it stood.
    let policy = Policy {
        reject_ambiguous_space,
        require_compact: true,
    };
    let token_window = limits
        .token_bytes
        .checked_add(1)
        .ok_or(StreamError::Audit(Error::Allocation))?;
    let mut buffer = Vec::new();
    buffer
        .try_reserve_exact(token_window)
        .map_err(|_allocation| StreamError::Audit(Error::Allocation))?;

    let max_total = u64::try_from(limits.bytes).unwrap_or(u64::MAX);
    let mut guarded =
        GuardedBufRead::new(reader, max_total, limits.token_bytes).map_err(StreamError::Input)?;
    guarded
        .try_reserve_capture(token_window)
        .map_err(|_allocation| StreamError::Audit(Error::Allocation))?;
    let saw_bom = guarded.saw_bom();
    let mut reader = Reader::from_reader(guarded);
    reader.config_mut().trim_text(false);
    let mut state = State::new();

    loop {
        buffer.clear();
        reader.get_mut().begin_token();
        let start_u64 = reader.buffer_position();
        let physical_start_u64 = reader.get_ref().position();
        let bom_bytes = if saw_bom { 3 } else { 0 };
        let start = usize::try_from(start_u64)
            .unwrap_or(usize::MAX)
            .saturating_add(bom_bytes);
        let physical_start = usize::try_from(physical_start_u64)
            .unwrap_or(usize::MAX)
            .max(bom_bytes);
        let event = reader
            .read_event_into(&mut buffer)
            .map_err(|error| map_stream_reader_error(error, start))?;
        let end_u64 = reader.buffer_position();
        let token_bytes = end_u64
            .checked_sub(start_u64)
            .and_then(|value| usize::try_from(value).ok())
            .ok_or_else(|| {
                StreamError::Audit(Error::malformed(
                    start,
                    "parser position moved backwards or exceeded addressable input",
                ))
            })?;

        state.events = checked_add(state.events, 1, Resource::Events, limits.events, start)
            .map_err(StreamError::Audit)?;
        check_limit(Resource::TokenBytes, limits.token_bytes, token_bytes, start)
            .map_err(StreamError::Audit)?;

        let raw = reader.get_ref().captured();

        // `read_event_into` borrows the reusable buffer. Checking UTF-8 here
        // catches code points split across arbitrary source chunks after the
        // parser has assembled the complete bounded event.
        let event_bytes: &[u8] = &event;
        std::str::from_utf8(event_bytes).map_err(|error| {
            StreamError::Audit(Error::Encoding {
                valid_up_to: event_encoding_offset(
                    &event,
                    physical_start,
                    error.valid_up_to(),
                    token_bytes,
                ),
            })
        })?;

        match event {
            Event::Start(tag) => {
                finish_text_run(&mut state, policy.reject_ambiguous_space)
                    .map_err(StreamError::Audit)?;
                check_start(raw, false, start, policy.require_compact)
                    .map_err(StreamError::Audit)?;
                let space = inspect_attributes(
                    &tag,
                    reader.decoder(),
                    state.current_space(),
                    &mut state,
                    limits,
                    start,
                )
                .map_err(StreamError::Audit)?;
                enter_element(&mut state, limits, start).map_err(StreamError::Audit)?;
                state
                    .spaces
                    .try_reserve(1)
                    .map_err(|_allocation| StreamError::Audit(Error::Allocation))?;
                state.spaces.push(space);
            },
            Event::Empty(tag) => {
                finish_text_run(&mut state, policy.reject_ambiguous_space)
                    .map_err(StreamError::Audit)?;
                check_start(raw, true, start, policy.require_compact)
                    .map_err(StreamError::Audit)?;
                inspect_attributes(
                    &tag,
                    reader.decoder(),
                    state.current_space(),
                    &mut state,
                    limits,
                    start,
                )
                .map_err(StreamError::Audit)?;
                enter_empty(&mut state, limits, start).map_err(StreamError::Audit)?;
            },
            Event::End(_) => {
                finish_text_run(&mut state, policy.reject_ambiguous_space)
                    .map_err(StreamError::Audit)?;
                check_end(raw, start, policy.require_compact).map_err(StreamError::Audit)?;
                if state.depth == 0 || state.spaces.pop().is_none() {
                    return Err(StreamError::Audit(Error::malformed(
                        start,
                        "unexpected end element",
                    )));
                }
                state.depth -= 1;
            },
            Event::Text(text) => {
                let bytes = text.as_ref();
                check_character_context(state.depth, bytes, start).map_err(StreamError::Audit)?;
                charge_text(&mut state, limits, bytes.len(), start).map_err(StreamError::Audit)?;
                let whitespace = is_xml_whitespace(bytes);
                if policy.require_compact
                    && ((state.depth == 0 && whitespace)
                        || (is_structural_whitespace(bytes)
                            && state.current_space() != Space::Preserve))
                {
                    return Err(StreamError::Audit(Error::NotCompact(Violation {
                        kind: Kind::FormattingWhitespace,
                        offset: start,
                    })));
                }
                if policy.reject_ambiguous_space && state.current_space() != Space::Preserve {
                    if whitespace && !state.text_run_has_explicit_content {
                        state.ambiguous_space_offset.get_or_insert(start);
                    } else if !whitespace {
                        state.ambiguous_space_offset = None;
                        state.text_run_has_explicit_content = true;
                    }
                }
            },
            Event::CData(data) => {
                if state.depth == 0 {
                    return Err(StreamError::Audit(Error::malformed(
                        start,
                        "CDATA outside the document element",
                    )));
                }
                charge_text(&mut state, limits, data.as_ref().len(), start)
                    .map_err(StreamError::Audit)?;
                state.ambiguous_space_offset = None;
                state.text_run_has_explicit_content = true;
            },
            Event::GeneralRef(reference) => {
                check_character_context(state.depth, reference.as_ref(), start)
                    .map_err(StreamError::Audit)?;
                charge_text(&mut state, limits, raw.len(), start).map_err(StreamError::Audit)?;
                state.ambiguous_space_offset = None;
                state.text_run_has_explicit_content = true;
            },
            Event::Decl(_) => {
                finish_text_run(&mut state, policy.reject_ambiguous_space)
                    .map_err(StreamError::Audit)?;
                check_declaration(raw, start, policy.require_compact)
                    .map_err(StreamError::Audit)?;
            },
            Event::Comment(_) | Event::PI(_) => {
                finish_text_run(&mut state, policy.reject_ambiguous_space)
                    .map_err(StreamError::Audit)?;
            },
            Event::DocType(_) => {
                finish_text_run(&mut state, policy.reject_ambiguous_space)
                    .map_err(StreamError::Audit)?;
                return Err(StreamError::Audit(Error::Doctype { offset: start }));
            },
            Event::Eof => {
                finish_text_run(&mut state, policy.reject_ambiguous_space)
                    .map_err(StreamError::Audit)?;
                break;
            },
        }
    }

    let final_offset = usize::try_from(reader.get_ref().position()).unwrap_or(usize::MAX);
    if state.depth != 0 {
        return Err(StreamError::Audit(Error::malformed(
            final_offset,
            "unclosed document element",
        )));
    }
    if state.roots != 1 {
        return Err(StreamError::Audit(Error::malformed(
            final_offset,
            "XML must contain exactly one document element",
        )));
    }

    Ok(Report {
        attributes: state.attributes,
        bytes: final_offset,
        events: state.events,
        max_depth: state.max_depth,
        text_bytes: state.text_bytes,
    })
}

fn event_encoding_offset(
    event: &Event<'_>,
    start: usize,
    event_offset: usize,
    token_bytes: usize,
) -> usize {
    let prefix = match event {
        Event::Start(_) | Event::Empty(_) => 1,
        Event::End(_) | Event::Decl(_) | Event::PI(_) => 2,
        Event::GeneralRef(_) => 1,
        Event::Comment(_) => 4,
        Event::CData(_) => 9,
        Event::DocType(content) => token_bytes
            .saturating_sub(content.as_ref().len())
            .saturating_sub(1),
        Event::Text(_) | Event::Eof => 0,
    };
    start.saturating_add(prefix).saturating_add(event_offset)
}

/// A `BufRead` facade that exposes at most the remaining total and per-event
/// windows. It retains only the current token's consumed bytes for raw lexical
/// checks; the caller's reader remains the owner of the input stream.
struct GuardedBufRead<R> {
    inner: R,
    // Holding this fixed prefix makes UTF-8 BOM handling independent of the
    // source's chunk size without allocating a second input buffer.
    prefix: [u8; 3],
    prefix_len: usize,
    prefix_pos: usize,
    saw_bom: bool,
    total: u64,
    token: usize,
    max_total: u64,
    max_token: usize,
    max_token_window: usize,
    captured: Vec<u8>,
    exposed: Vec<u8>,
    exposed_pos: usize,
    exposed_prefix: bool,
}

impl<R: BufRead> GuardedBufRead<R> {
    fn new(inner: R, max_total: u64, max_token: usize) -> io::Result<Self> {
        let max_token_window = max_token.checked_add(1).ok_or_else(|| {
            io::Error::new(
                io::ErrorKind::InvalidInput,
                "XML token window overflows usize",
            )
        })?;
        let mut guarded = Self {
            inner,
            prefix: [0; 3],
            prefix_len: 0,
            prefix_pos: 0,
            saw_bom: false,
            total: 0,
            token: 0,
            max_total,
            max_token,
            max_token_window,
            captured: Vec::new(),
            exposed: Vec::new(),
            exposed_pos: 0,
            exposed_prefix: false,
        };

        // Read at most three bytes up front so a BOM split over arbitrary
        // `BufRead` chunks is recognized consistently. Bytes that do not form
        // a BOM remain pending source and are counted when the parser consumes
        // them.
        let prefix_limit = max_total.min(3) as usize;
        while guarded.prefix_len < prefix_limit {
            let available = loop {
                match guarded.inner.fill_buf() {
                    Ok(available) => break available,
                    Err(error) if error.kind() == io::ErrorKind::Interrupted => continue,
                    Err(error) => return Err(error),
                }
            };
            if available.is_empty() {
                break;
            }
            let count = available
                .len()
                .min(prefix_limit.saturating_sub(guarded.prefix_len));
            guarded.prefix[guarded.prefix_len..guarded.prefix_len + count]
                .copy_from_slice(&available[..count]);
            guarded.inner.consume(count);
            guarded.prefix_len += count;
        }
        if guarded.prefix_len == 3 && guarded.prefix == [0xEF, 0xBB, 0xBF] {
            guarded.saw_bom = true;
        }
        Ok(guarded)
    }

    const fn saw_bom(&self) -> bool {
        self.saw_bom
    }

    fn try_reserve_capture(&mut self, capacity: usize) -> Result<(), ()> {
        self.captured.try_reserve_exact(capacity).map_err(|_| ())?;
        self.exposed.try_reserve_exact(capacity).map_err(|_| ())
    }

    fn begin_token(&mut self) {
        self.token = 0;
        self.captured.clear();
    }

    const fn position(&self) -> u64 {
        self.total
    }

    fn pending(&self) -> &[u8] {
        &self.prefix[self.prefix_pos..self.prefix_len]
    }

    fn captured(&self) -> &[u8] {
        &self.captured
    }
}

impl<R: BufRead> Read for GuardedBufRead<R> {
    fn read(&mut self, output: &mut [u8]) -> io::Result<usize> {
        if output.is_empty() {
            return Ok(0);
        }
        let available = self.fill_buf()?;
        let count = available.len().min(output.len());
        output[..count].copy_from_slice(&available[..count]);
        self.consume(count);
        Ok(count)
    }
}

impl<R: BufRead> BufRead for GuardedBufRead<R> {
    fn fill_buf(&mut self) -> io::Result<&[u8]> {
        if self.exposed_pos < self.exposed.len() {
            return Ok(&self.exposed[self.exposed_pos..]);
        }
        self.exposed.clear();
        self.exposed_pos = 0;

        if self.token > self.max_token {
            return Err(window_token_error(self.token, self.max_token));
        }
        if self.total >= self.max_total {
            if !self.pending().is_empty() {
                return Err(window_total_error(
                    self.total.saturating_add(1),
                    self.max_total,
                ));
            }
            let available = self.inner.fill_buf()?;
            if available.is_empty() {
                return Ok(available);
            }
            return Err(window_total_error(
                self.total.saturating_add(1),
                self.max_total,
            ));
        }

        // Expose the complete initial BOM to quick-xml even when the token
        // ceiling is smaller than three. It is framing, not an XML token;
        // consume charges it to total bytes only. Keeping it in the parser's
        // input also ensures a second BOM remains ordinary character data.
        if self.saw_bom && self.prefix_pos == 0 {
            self.exposed.extend_from_slice(&self.prefix);
            self.exposed_prefix = true;
            return Ok(&self.exposed);
        }

        let total_remaining = usize::try_from(self.max_total - self.total).unwrap_or(usize::MAX);
        let token = self.token;
        let token_remaining = self.max_token_window.saturating_sub(token);
        if token_remaining == 0 {
            // The lookahead byte has already been consumed into quick-xml's
            // event buffer. Returning it again would permit a zero-progress
            // loop on malformed input.
            return Err(window_token_error(token, self.max_token));
        }

        let pending = !self.pending().is_empty();
        if pending {
            let start = self.prefix_pos;
            let available = &self.prefix[start..self.prefix_len];
            if available.is_empty() {
                return Ok(available);
            }
            let visible = available.len().min(total_remaining).min(token_remaining);
            self.exposed.extend_from_slice(&available[..visible]);
        } else {
            let mut exposed = std::mem::take(&mut self.exposed);
            let available = match self.inner.fill_buf() {
                Ok(available) => available,
                Err(error) => {
                    self.exposed = exposed;
                    return Err(error);
                },
            };
            if available.is_empty() {
                self.exposed = exposed;
                return Ok(&self.exposed);
            }
            let visible = available.len().min(total_remaining).min(token_remaining);
            exposed.extend_from_slice(&available[..visible]);
            self.exposed = exposed;
        }
        self.exposed_prefix = pending;
        Ok(&self.exposed)
    }

    fn consume(&mut self, amount: usize) {
        // quick-xml consumes only bytes returned by `fill_buf`; saturating the
        // accounting keeps a hostile/incorrect source from causing a panic.
        let available = self.exposed.len().saturating_sub(self.exposed_pos);
        let bom = self.saw_bom && self.exposed_prefix && self.prefix_pos < 3;
        let amount = if bom {
            amount.min(available)
        } else {
            amount
                .min(available)
                .min(self.max_token_window.saturating_sub(self.token))
        };
        if amount == 0 {
            return;
        }
        if !bom {
            self.captured
                .extend_from_slice(&self.exposed[self.exposed_pos..self.exposed_pos + amount]);
            self.token = self.token.saturating_add(amount);
        }
        if self.exposed_prefix {
            self.prefix_pos = self.prefix_pos.saturating_add(amount);
        } else {
            self.inner.consume(amount);
        }
        self.exposed_pos += amount;
        self.total = self.total.saturating_add(amount as u64);
    }
}

#[derive(Debug)]
enum WindowError {
    Total { observed: u64, limit: u64 },
    Token { observed: usize, limit: usize },
}

impl fmt::Display for WindowError {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Self::Total { observed, limit } => write!(
                formatter,
                "XML input byte window exceeded: observed {observed}, limit {limit}"
            ),
            Self::Token { observed, limit } => write!(
                formatter,
                "XML token window exceeded: observed {observed}, limit {limit}"
            ),
        }
    }
}

impl std::error::Error for WindowError {}

fn window_total_error(observed: u64, limit: u64) -> io::Error {
    io::Error::other(WindowError::Total { observed, limit })
}

fn window_token_error(observed: usize, limit: usize) -> io::Error {
    io::Error::other(WindowError::Token { observed, limit })
}

fn map_stream_reader_error(error: quick_xml::Error, offset: usize) -> StreamError {
    match error {
        quick_xml::Error::Io(source) => {
            if let Some(window) = source
                .get_ref()
                .and_then(|inner| inner.downcast_ref::<WindowError>())
            {
                return StreamError::Audit(window.to_audit_error(offset));
            }
            StreamError::Input(io::Error::new(source.kind(), source))
        },
        quick_xml::Error::Encoding(_) => StreamError::Audit(Error::Encoding {
            valid_up_to: offset,
        }),
        other => StreamError::Audit(Error::malformed(offset, other.to_string())),
    }
}

impl WindowError {
    fn to_audit_error(&self, offset: usize) -> Error {
        match self {
            Self::Total { observed, limit } => Error::Limit {
                resource: Resource::Bytes,
                limit: usize::try_from(*limit).unwrap_or(usize::MAX),
                actual: usize::try_from(*observed).unwrap_or(usize::MAX),
                offset,
            },
            Self::Token { observed, limit } => Error::Limit {
                resource: Resource::TokenBytes,
                limit: *limit,
                actual: *observed,
                offset,
            },
        }
    }
}

fn verify_with_policy(input: &[u8], limits: Limits, policy: Policy) -> Result<Report, Error> {
    verify_observed(input, limits, policy, &mut ())
}

/// The slice auditor, with an observer of the tokens it accepts.
///
/// The observer only reads the auditor's state, so every entry point reaches
/// the verdict [`verify_with_policy`] reaches; the unit observer compiles
/// away.
fn verify_observed<O: Observer>(
    input: &[u8],
    limits: Limits,
    policy: Policy,
    observer: &mut O,
) -> Result<Report, Error> {
    check_limit(Resource::Bytes, limits.bytes, input.len(), 0)?;
    let xml = std::str::from_utf8(input).map_err(|error| Error::Encoding {
        valid_up_to: error.valid_up_to(),
    })?;
    // quick-xml excludes the leading UTF-8 BOM from buffer_position(),
    // while raw lexical spans and diagnostics address the original input.
    let bom_bytes = if input.starts_with(b"\xEF\xBB\xBF") {
        3
    } else {
        0
    };
    let mut state = State::new();
    scan(input, xml, bom_bytes, limits, policy, &mut state, observer)?;

    if state.depth != 0 {
        return Err(Error::malformed(input.len(), "unclosed document element"));
    }
    if state.roots != 1 {
        return Err(Error::malformed(
            input.len(),
            "XML must contain exactly one document element",
        ));
    }

    Ok(Report {
        attributes: state.attributes,
        bytes: input.len(),
        events: state.events,
        max_depth: state.max_depth,
        text_bytes: state.text_bytes,
    })
}

/// The slice auditor's token loop: every token from the start of `xml` to
/// its EOF, checked from `state` onward.
///
/// `input` is the same text as `xml` plus a leading byte-order mark of
/// `bom_bytes` bytes, which quick-xml consumes without counting it; raw spans
/// and diagnostics address `input`.
fn scan<O: Observer>(
    input: &[u8],
    xml: &str,
    bom_bytes: usize,
    limits: Limits,
    policy: Policy,
    state: &mut State,
    observer: &mut O,
) -> Result<(), Error> {
    let mut reader = Reader::from_str(xml);
    reader.config_mut().trim_text(false);

    loop {
        let start = position(&reader).saturating_add(bom_bytes);
        let event = reader.read_event().map_err(|error| {
            Error::malformed(
                position(&reader).saturating_add(bom_bytes),
                error.to_string(),
            )
        })?;
        let end = position(&reader).saturating_add(bom_bytes);
        let raw = input
            .get(start..end)
            .ok_or_else(|| Error::malformed(start, "parser position escaped input"))?;
        let listening = observer.listening();
        let before = if listening {
            Counters::of(state)
        } else {
            Counters::default()
        };

        state.events = checked_add(state.events, 1, Resource::Events, limits.events, start)?;
        check_limit(Resource::TokenBytes, limits.token_bytes, raw.len(), start)?;

        let token = match event {
            Event::Start(tag) => {
                finish_text_run(state, policy.reject_ambiguous_space)?;
                check_start(raw, false, start, policy.require_compact)?;
                let space = inspect_attributes(
                    &tag,
                    reader.decoder(),
                    state.current_space(),
                    state,
                    limits,
                    start,
                )?;
                enter_element(state, limits, start)?;
                state
                    .spaces
                    .try_reserve(1)
                    .map_err(|_allocation| Error::Allocation)?;
                state.spaces.push(space);
                Token::Start
            },
            Event::Empty(tag) => {
                finish_text_run(state, policy.reject_ambiguous_space)?;
                check_start(raw, true, start, policy.require_compact)?;
                inspect_attributes(
                    &tag,
                    reader.decoder(),
                    state.current_space(),
                    state,
                    limits,
                    start,
                )?;
                enter_empty(state, limits, start)?;
                Token::Empty
            },
            Event::End(_) => {
                finish_text_run(state, policy.reject_ambiguous_space)?;
                check_end(raw, start, policy.require_compact)?;
                if state.depth == 0 || state.spaces.pop().is_none() {
                    return Err(Error::malformed(start, "unexpected end element"));
                }
                state.depth -= 1;
                Token::End
            },
            Event::Text(text) => {
                let bytes = text.as_ref();
                check_character_context(state.depth, bytes, start)?;
                charge_text(state, limits, bytes.len(), start)?;
                let whitespace = is_xml_whitespace(bytes);
                if policy.require_compact
                    && ((state.depth == 0 && whitespace)
                        || (is_structural_whitespace(bytes)
                            && state.current_space() != Space::Preserve))
                {
                    return Err(Error::NotCompact(Violation {
                        kind: Kind::FormattingWhitespace,
                        offset: start,
                    }));
                }
                if policy.reject_ambiguous_space && state.current_space() != Space::Preserve {
                    if whitespace && !state.text_run_has_explicit_content {
                        state.ambiguous_space_offset.get_or_insert(start);
                    } else if !whitespace {
                        state.ambiguous_space_offset = None;
                        state.text_run_has_explicit_content = true;
                    }
                }
                Token::Character
            },
            Event::CData(data) => {
                if state.depth == 0 {
                    return Err(Error::malformed(
                        start,
                        "CDATA outside the document element",
                    ));
                }
                charge_text(state, limits, data.as_ref().len(), start)?;
                state.ambiguous_space_offset = None;
                state.text_run_has_explicit_content = true;
                Token::Markup
            },
            Event::GeneralRef(reference) => {
                check_character_context(state.depth, reference.as_ref(), start)?;
                charge_text(state, limits, raw.len(), start)?;
                state.ambiguous_space_offset = None;
                state.text_run_has_explicit_content = true;
                Token::Character
            },
            Event::Decl(_) => {
                finish_text_run(state, policy.reject_ambiguous_space)?;
                check_declaration(raw, start, policy.require_compact)?;
                Token::Markup
            },
            Event::Comment(_) | Event::PI(_) => {
                finish_text_run(state, policy.reject_ambiguous_space)?;
                Token::Markup
            },
            Event::DocType(_) => {
                finish_text_run(state, policy.reject_ambiguous_space)?;
                return Err(Error::Doctype { offset: start });
            },
            Event::Eof => {
                finish_text_run(state, policy.reject_ambiguous_space)?;
                break;
            },
        };
        if listening {
            observer.accepted(token, start, end, before, state);
        }
    }
    Ok(())
}

/// Aggregate budgets the slice auditor has charged, as it counts them.
#[derive(Clone, Copy, Debug, Default, Eq, PartialEq)]
struct Counters {
    attributes: usize,
    events: usize,
    text_bytes: usize,
}

impl Counters {
    const fn of(state: &State) -> Self {
        Self {
            attributes: state.attributes,
            events: state.events,
            text_bytes: state.text_bytes,
        }
    }

    /// What was charged between `earlier` and `self`.
    fn since(self, earlier: Self) -> Option<Self> {
        Some(Self {
            attributes: self.attributes.checked_sub(earlier.attributes)?,
            events: self.events.checked_sub(earlier.events)?,
            text_bytes: self.text_bytes.checked_sub(earlier.text_bytes)?,
        })
    }

    /// These totals with `removed` taken out and `added` put in.
    fn replaced(self, removed: Self, added: Self) -> Option<Self> {
        Some(Self {
            attributes: self
                .attributes
                .checked_sub(removed.attributes)?
                .checked_add(added.attributes)?,
            events: self
                .events
                .checked_sub(removed.events)?
                .checked_add(added.events)?,
            text_bytes: self
                .text_bytes
                .checked_sub(removed.text_bytes)?
                .checked_add(added.text_bytes)?,
        })
    }

    const fn within(self, limits: Limits) -> bool {
        self.attributes <= limits.attributes
            && self.events <= limits.events
            && self.text_bytes <= limits.text_bytes
    }
}

/// The lexical class of one token the slice auditor accepted.
#[derive(Clone, Copy, Debug, Eq, PartialEq)]
enum Token {
    /// A start tag; its element is now open.
    Start,
    /// An end tag; its element is now closed.
    End,
    /// An empty-element tag.
    Empty,
    /// Other markup delimited by `<` and `>`: an XML declaration, a comment,
    /// a processing instruction or a CDATA section.
    Markup,
    /// Character data or a general reference.
    Character,
}

/// A reader of the tokens the slice auditor accepts.
///
/// An observer sees each token after it passed every check and updated the
/// state, and only reads that state, so it cannot change a verdict.
trait Observer {
    /// Whether the observer still wants tokens. Once it answers `false`, the
    /// scan charges it nothing but this question.
    fn listening(&self) -> bool {
        true
    }

    /// `before` holds the counters as they stood before `token` was charged;
    /// `state` is the auditor's state after it.
    fn accepted(&mut self, token: Token, start: usize, end: usize, before: Counters, state: &State);
}

impl Observer for () {
    #[inline(always)]
    fn listening(&self) -> bool {
        false
    }

    #[inline(always)]
    fn accepted(
        &mut self,
        _token: Token,
        _start: usize,
        _end: usize,
        _before: Counters,
        _state: &State,
    ) {
    }
}

/// Watches the audit of an original payload for the innermost element that
/// covers every byte in which a replacement differs from it.
///
/// Elements are considered as they close, innermost first. The first one that
/// is not the document element, starts at or before the first differing byte,
/// ends at or after the start of the common tail, and whose replacement bytes
/// begin with `<` and end with `>`, is the window.
struct WindowSearch<'r> {
    replacement: &'r [u8],
    original_len: usize,
    /// `original[..prefix] == replacement[..prefix]`.
    prefix: usize,
    /// `original[tail..]` equals the same number of trailing replacement bytes.
    tail: usize,
    /// Open elements that started at or before `prefix`: their start offsets
    /// and the counters before their start tags.
    open: Vec<(usize, Counters)>,
    /// Open elements that started after `prefix`. None of them can cover the
    /// first differing byte.
    untracked: usize,
    found: Option<Window>,
    stopped: bool,
}

impl<'r> WindowSearch<'r> {
    const fn new(replacement: &'r [u8], original_len: usize, prefix: usize, tail: usize) -> Self {
        Self {
            replacement,
            original_len,
            prefix,
            tail,
            open: Vec::new(),
            untracked: 0,
            found: None,
            stopped: false,
        }
    }

    /// Consider the element `original[start..end]`, which has just closed,
    /// leaving `state.depth` elements open around it.
    fn consider(&mut self, start: usize, end: usize, before: Counters, state: &State) {
        // Depth zero is the document element, whose window would be the
        // whole document; an element ending before the common tail does not
        // cover every difference.
        if state.depth == 0 || end < self.tail {
            return;
        }
        let Some(window_end) = end
            .checked_add(self.replacement.len())
            .and_then(|total| total.checked_sub(self.original_len))
        else {
            return;
        };
        if window_end <= start
            || self.replacement.get(start) != Some(&b'<')
            || self.replacement.get(window_end - 1) != Some(&b'>')
        {
            return;
        }
        self.stopped = true;
        self.found = Counters::of(state).since(before).map(|removed| Window {
            original: start..end,
            replacement: start..window_end,
            depth: state.depth,
            removed,
        });
    }
}

impl Observer for WindowSearch<'_> {
    #[inline]
    fn listening(&self) -> bool {
        !self.stopped
    }

    #[inline]
    fn accepted(
        &mut self,
        token: Token,
        start: usize,
        end: usize,
        before: Counters,
        state: &State,
    ) {
        match token {
            Token::Start if start <= self.prefix => {
                if self.open.try_reserve(1).is_err() {
                    self.stopped = true;
                    return;
                }
                self.open.push((start, before));
            },
            Token::Start => self.untracked += 1,
            Token::End if self.untracked > 0 => self.untracked -= 1,
            Token::End => match self.open.pop() {
                Some((opened, counters)) => self.consider(opened, end, counters, state),
                None => self.stopped = true,
            },
            Token::Empty if start <= self.prefix => self.consider(start, end, before, state),
            Token::Empty | Token::Markup | Token::Character => {},
        }
    }
}

/// One element of an original payload and the replacement bytes that took
/// its place.
#[derive(Debug)]
struct Window {
    original: core::ops::Range<usize>,
    replacement: core::ops::Range<usize>,
    /// Elements open around the window; at least one, the document element.
    depth: usize,
    /// Counters the original element was charged.
    removed: Counters,
}

impl Window {
    /// Prove the replacement's source-policy verdict from the original's, or
    /// return `None` so the caller audits the replacement completely.
    ///
    /// Only success is decided here. Every failure, and every doubt, is left
    /// to the complete audit, which reports the canonical error and offset.
    fn prove(
        self,
        original: Report,
        replacement: &[u8],
        limits: Limits,
    ) -> Option<ReplacementProof> {
        // A window of more than half the payload saves nothing, and a failure
        // inside it would then be scanned twice.
        if replacement.len() > limits.bytes
            || self.replacement.len().checked_mul(2)? > replacement.len()
        {
            return None;
        }
        let bytes = replacement.get(self.replacement.clone())?;
        let added = audit_window(bytes, limits, self.depth)?;
        let totals = Counters {
            attributes: original.attributes,
            events: original.events,
            text_bytes: original.text_bytes,
        }
        .replaced(self.removed, added)?;
        if !totals.within(limits) {
            return None;
        }
        Some(ReplacementProof::Window {
            original: self.original,
            replacement: self.replacement,
        })
    }
}

/// Audit `bytes` under the source policy as the content that replaced one
/// element with `depth` elements open around it: the document element is
/// open, and no element opened outside the window may be closed inside it.
///
/// Returns the counters the window is charged, excluding its own EOF event,
/// or `None` when the window fails a check, is not balanced, or does not begin
/// and end with markup.
fn audit_window(bytes: &[u8], limits: Limits, depth: usize) -> Option<Counters> {
    if depth == 0 || bytes.first() != Some(&b'<') {
        return None;
    }
    let xml = std::str::from_utf8(bytes).ok()?;
    let mut state = State::within(depth);
    let mut last = LastToken::default();
    scan(bytes, xml, 0, limits, Policy::SOURCE, &mut state, &mut last).ok()?;
    if state.depth != depth || !last.is_markup_ending_at(bytes.len()) {
        return None;
    }
    let charged = Counters::of(&state);
    Some(Counters {
        events: charged.events.checked_sub(1)?,
        ..charged
    })
}

/// Remembers the last token a scan accepted.
#[derive(Default)]
struct LastToken {
    token: Option<Token>,
    end: usize,
}

impl LastToken {
    fn is_markup_ending_at(&self, end: usize) -> bool {
        self.end == end && matches!(self.token, Some(Token::End | Token::Empty | Token::Markup))
    }
}

impl Observer for LastToken {
    fn accepted(
        &mut self,
        token: Token,
        _start: usize,
        end: usize,
        _before: Counters,
        _state: &State,
    ) {
        self.token = Some(token);
        self.end = end;
    }
}

/// Bytes compared per slice comparison before the scan narrows down.
const COMPARE_BLOCK: usize = 4096;
/// Bytes compared per slice comparison inside the first unequal block.
const COMPARE_LINE: usize = 64;

/// Length of the longest common prefix of `left` and `right`.
fn common_prefix_len(left: &[u8], right: &[u8]) -> usize {
    let limit = left.len().min(right.len());
    let mut matched = 0;
    for step in [COMPARE_BLOCK, COMPARE_LINE] {
        while matched + step <= limit
            && left[matched..matched + step] == right[matched..matched + step]
        {
            matched += step;
        }
    }
    matched
        + left[matched..limit]
            .iter()
            .zip(&right[matched..limit])
            .take_while(|(left, right)| left == right)
            .count()
}

/// Length of the longest common suffix of `left` and `right`.
fn common_suffix_len(left: &[u8], right: &[u8]) -> usize {
    let limit = left.len().min(right.len());
    let mut matched = 0;
    for step in [COMPARE_BLOCK, COMPARE_LINE] {
        while matched + step <= limit
            && left[left.len() - matched - step..left.len() - matched]
                == right[right.len() - matched - step..right.len() - matched]
        {
            matched += step;
        }
    }
    matched
        + left[..left.len() - matched]
            .iter()
            .rev()
            .zip(right[..right.len() - matched].iter().rev())
            .take(limit - matched)
            .take_while(|(left, right)| left == right)
            .count()
}

fn finish_text_run(state: &mut State, reject_ambiguous_space: bool) -> Result<(), Error> {
    if reject_ambiguous_space && let Some(offset) = state.ambiguous_space_offset.take() {
        return Err(Error::NotCompact(Violation {
            kind: Kind::AmbiguousWhitespace,
            offset,
        }));
    }
    state.ambiguous_space_offset = None;
    state.text_run_has_explicit_content = false;
    Ok(())
}

fn charge_text(
    state: &mut State,
    limits: Limits,
    amount: usize,
    offset: usize,
) -> Result<(), Error> {
    state.text_bytes = checked_add(
        state.text_bytes,
        amount,
        Resource::TextBytes,
        limits.text_bytes,
        offset,
    )?;
    Ok(())
}

fn check_character_context(depth: usize, bytes: &[u8], offset: usize) -> Result<(), Error> {
    if depth == 0 && !is_xml_whitespace(bytes) {
        return Err(Error::malformed(
            offset,
            "character data outside the document element",
        ));
    }
    Ok(())
}

fn enter_element(state: &mut State, limits: Limits, offset: usize) -> Result<(), Error> {
    if state.depth == 0 {
        if state.roots != 0 {
            return Err(Error::malformed(offset, "multiple document elements"));
        }
        state.roots = 1;
    }
    state.depth = checked_add(state.depth, 1, Resource::Depth, limits.depth, offset)?;
    state.max_depth = state.max_depth.max(state.depth);
    Ok(())
}

fn enter_empty(state: &mut State, limits: Limits, offset: usize) -> Result<(), Error> {
    if state.depth == 0 {
        if state.roots != 0 {
            return Err(Error::malformed(offset, "multiple document elements"));
        }
        state.roots = 1;
    }
    let depth = checked_add(state.depth, 1, Resource::Depth, limits.depth, offset)?;
    state.max_depth = state.max_depth.max(depth);
    Ok(())
}

fn inspect_attributes(
    tag: &BytesStart<'_>,
    decoder: Decoder,
    inherited: Space,
    state: &mut State,
    limits: Limits,
    offset: usize,
) -> Result<Space, Error> {
    let mut space = inherited;
    for attribute_result in tag.attributes() {
        let attribute =
            attribute_result.map_err(|error| Error::malformed(offset, error.to_string()))?;
        state.attributes = checked_add(
            state.attributes,
            1,
            Resource::Attributes,
            limits.attributes,
            offset,
        )?;
        if attribute.key.as_ref() == b"xml:space" {
            let value = attribute
                .decoded_and_normalized_value(XmlVersion::Implicit1_0, decoder)
                .map_err(|error| Error::malformed(offset, error.to_string()))?;
            space = match value.as_ref() {
                "default" => Space::Default,
                "preserve" => Space::Preserve,
                _ => {
                    return Err(Error::malformed(
                        offset,
                        "xml:space must be 'default' or 'preserve'",
                    ));
                },
            };
        }
    }
    Ok(space)
}

fn check_declaration(raw: &[u8], offset: usize, compact: bool) -> Result<(), Error> {
    let Some(inner) = raw
        .strip_prefix(b"<?")
        .and_then(|value| value.strip_suffix(b"?>"))
    else {
        return Err(Error::malformed(offset, "invalid XML declaration boundary"));
    };
    check_attribute_layout(inner, offset + 2, compact)
}

fn check_end(raw: &[u8], offset: usize, compact: bool) -> Result<(), Error> {
    let Some(inner) = raw
        .strip_prefix(b"</")
        .and_then(|value| value.strip_suffix(b">"))
    else {
        return Err(Error::malformed(offset, "invalid end-tag boundary"));
    };
    if let Some(index) = inner.iter().position(|byte| is_space(*byte)) {
        let trailing = inner[index..].iter().all(|byte| is_space(*byte));
        if !compact {
            // `</name >` is well-formed XML that a producer may write; only an
            // end tag carrying markup after its name stays refused.
            if trailing {
                return Ok(());
            }
            return Err(Error::malformed(
                offset + 2 + index,
                "end tag must not carry markup after its name",
            ));
        }
        let kind = if trailing {
            Kind::WhitespaceBeforeClose
        } else {
            Kind::AttributeSeparation
        };
        return Err(Error::NotCompact(Violation {
            kind,
            offset: offset + 2 + index,
        }));
    }
    Ok(())
}

fn check_start(raw: &[u8], empty: bool, offset: usize, compact: bool) -> Result<(), Error> {
    let Some(without_open) = raw.strip_prefix(b"<") else {
        return Err(Error::malformed(offset, "invalid start-tag boundary"));
    };
    let inner = if empty {
        without_open.strip_suffix(b"/>")
    } else {
        without_open.strip_suffix(b">")
    }
    .ok_or_else(|| Error::malformed(offset, "invalid start-tag close"))?;
    check_attribute_layout(inner, offset + 1, compact)
}

fn check_attribute_layout(inner: &[u8], offset: usize, compact: bool) -> Result<(), Error> {
    let Some(mut cursor) = inner.iter().position(|byte| is_space(*byte)) else {
        return Ok(());
    };

    loop {
        let separator = cursor;
        while cursor < inner.len() && is_space(inner[cursor]) {
            cursor += 1;
        }
        if cursor == inner.len() {
            if !compact {
                // Whitespace before the tag close is a producer spelling.
                return Ok(());
            }
            return Err(Error::NotCompact(Violation {
                kind: Kind::WhitespaceBeforeClose,
                offset: offset + separator,
            }));
        }
        if compact && (cursor != separator + 1 || inner[separator] != b' ') {
            return Err(Error::NotCompact(Violation {
                kind: Kind::AttributeSeparation,
                offset: offset + separator,
            }));
        }

        let name_start = cursor;
        while cursor < inner.len() && !is_space(inner[cursor]) && inner[cursor] != b'=' {
            cursor += 1;
        }
        if cursor == name_start {
            return Err(Error::malformed(offset + cursor, "missing attribute name"));
        }
        if cursor == inner.len() || is_space(inner[cursor]) {
            if compact {
                return Err(Error::NotCompact(Violation {
                    kind: Kind::AttributeSeparation,
                    offset: offset + cursor,
                }));
            }
            // XML permits whitespace around `=`; the attribute must still have
            // one, and a name with no value stays refused.
            while cursor < inner.len() && is_space(inner[cursor]) {
                cursor += 1;
            }
            if cursor == inner.len() || inner[cursor] != b'=' {
                return Err(Error::malformed(
                    offset + cursor,
                    "attribute name must be followed by '='",
                ));
            }
        }
        cursor += 1;
        if cursor == inner.len() || is_space(inner[cursor]) {
            if compact {
                return Err(Error::NotCompact(Violation {
                    kind: Kind::AttributeSeparation,
                    offset: offset + cursor,
                }));
            }
            while cursor < inner.len() && is_space(inner[cursor]) {
                cursor += 1;
            }
            if cursor == inner.len() {
                return Err(Error::malformed(
                    offset + cursor,
                    "attribute value must be quoted",
                ));
            }
        }

        let quote = inner[cursor];
        if quote != b'\'' && quote != b'"' {
            return Err(Error::malformed(
                offset + cursor,
                "attribute value must be quoted",
            ));
        }
        cursor += 1;
        while cursor < inner.len() && inner[cursor] != quote {
            cursor += 1;
        }
        if cursor == inner.len() {
            return Err(Error::malformed(
                offset + cursor,
                "unterminated attribute value",
            ));
        }
        cursor += 1;
        if cursor == inner.len() {
            return Ok(());
        }
        if !is_space(inner[cursor]) {
            return Err(Error::malformed(
                offset + cursor,
                "missing attribute separator",
            ));
        }
    }
}

fn check_limit(
    resource: Resource,
    limit: usize,
    actual: usize,
    offset: usize,
) -> Result<(), Error> {
    if actual > limit {
        return Err(Error::Limit {
            resource,
            limit,
            actual,
            offset,
        });
    }
    Ok(())
}

fn checked_add(
    current: usize,
    amount: usize,
    resource: Resource,
    limit: usize,
    offset: usize,
) -> Result<usize, Error> {
    let actual = current.saturating_add(amount);
    check_limit(resource, limit, actual, offset)?;
    Ok(actual)
}

fn is_xml_whitespace(bytes: &[u8]) -> bool {
    !bytes.is_empty() && bytes.iter().all(|byte| is_space(*byte))
}

fn is_structural_whitespace(bytes: &[u8]) -> bool {
    is_xml_whitespace(bytes)
        && bytes
            .iter()
            .any(|byte| matches!(byte, b'\t' | b'\n' | b'\r'))
}

const fn is_space(byte: u8) -> bool {
    matches!(byte, b' ' | b'\t' | b'\n' | b'\r')
}

const fn minimum(left: usize, right: usize) -> usize {
    if left < right { left } else { right }
}

fn position(reader: &Reader<&[u8]>) -> usize {
    match usize::try_from(reader.buffer_position()) {
        Ok(value) => value,
        Err(_) => usize::MAX,
    }
}

/// Bounded verification over named XML package parts.
pub mod package {
    use super::{Error as DocumentError, Limits as DocumentLimits, verify as verify_document};
    use core::fmt;

    /// Returns whether a package member must be treated as XML.
    ///
    /// Package conventions identify XML both lexically (`.xml`, `.rels`, and
    /// `.rdf`) and through a manifest/content-type media type. Parameters are
    /// ignored and media type matching is ASCII case-insensitive.
    #[must_use]
    pub fn is_xml_part(name: &str, media_type: &str) -> bool {
        is_xml_name(name) || is_xml_media_type(media_type)
    }

    /// Returns whether a media type denotes XML according to the XML media
    /// type registrations and structured syntax suffix convention.
    #[must_use]
    pub fn is_xml_media_type(media_type: &str) -> bool {
        let essence = media_type
            .split_once(';')
            .map_or(media_type, |(value, _parameters)| value)
            .trim();
        essence.eq_ignore_ascii_case("application/xml")
            || essence.eq_ignore_ascii_case("text/xml")
            || essence
                .get(essence.len().saturating_sub(4)..)
                .is_some_and(|suffix| suffix.eq_ignore_ascii_case("+xml"))
    }

    fn is_xml_name(name: &str) -> bool {
        let leaf = name.rsplit('/').next().unwrap_or(name);
        if leaf.eq_ignore_ascii_case("[Content_Types].xml") {
            return true;
        }
        leaf.rsplit_once('.').is_some_and(|(_, extension)| {
            extension.eq_ignore_ascii_case("xml")
                || extension.eq_ignore_ascii_case("rels")
                || extension.eq_ignore_ascii_case("rdf")
        })
    }

    /// A borrowed named XML package member.
    #[derive(Clone, Copy, Debug)]
    pub struct Part<'a> {
        bytes: &'a [u8],
        name: &'a str,
    }

    impl<'a> Part<'a> {
        /// Creates a borrowed part without copying its name or payload.
        #[must_use]
        pub const fn new(name: &'a str, bytes: &'a [u8]) -> Self {
            Self { bytes, name }
        }

        /// Part payload.
        #[must_use]
        pub const fn bytes(self) -> &'a [u8] {
            self.bytes
        }

        /// Archive-relative diagnostic name.
        #[must_use]
        pub const fn name(self) -> &'a str {
            self.name
        }
    }

    /// Finite aggregate and per-document package audit limits.
    #[derive(Clone, Copy, Debug, Eq, PartialEq)]
    pub struct Limits {
        document: DocumentLimits,
        bytes: usize,
        parts: usize,
    }

    impl Limits {
        /// Creates an explicit package profile.
        #[must_use]
        pub const fn new(document: DocumentLimits, max_parts: usize, max_bytes: usize) -> Self {
            Self {
                document,
                bytes: max_bytes,
                parts: max_parts,
            }
        }

        /// Per-document limits.
        #[must_use]
        pub const fn document(self) -> DocumentLimits {
            self.document
        }

        /// Maximum aggregate XML payload bytes.
        #[must_use]
        pub const fn max_bytes(self) -> usize {
            self.bytes
        }

        /// Maximum named XML parts.
        #[must_use]
        pub const fn max_parts(self) -> usize {
            self.parts
        }
    }

    impl Default for Limits {
        fn default() -> Self {
            Self::new(DocumentLimits::default(), 65_536, 256 * 1024 * 1024)
        }
    }

    /// Aggregate package audit accounting.
    #[derive(Clone, Copy, Debug, Eq, PartialEq)]
    #[must_use]
    pub struct Report {
        attributes: usize,
        bytes: usize,
        events: usize,
        max_depth: usize,
        parts: usize,
        text_bytes: usize,
    }

    impl Report {
        /// Aggregate attributes.
        #[must_use]
        pub const fn attributes(self) -> usize {
            self.attributes
        }

        /// Aggregate XML bytes.
        #[must_use]
        pub const fn bytes(self) -> usize {
            self.bytes
        }

        /// Aggregate parser events.
        #[must_use]
        pub const fn events(self) -> usize {
            self.events
        }

        /// Greatest per-document depth.
        #[must_use]
        pub const fn max_depth(self) -> usize {
            self.max_depth
        }

        /// Number of audited parts.
        #[must_use]
        pub const fn parts(self) -> usize {
            self.parts
        }

        /// Aggregate character-data bytes.
        #[must_use]
        pub const fn text_bytes(self) -> usize {
            self.text_bytes
        }
    }

    /// Package-level audit failure borrowing the failing part name.
    #[derive(Debug)]
    #[non_exhaustive]
    pub enum Error<'a> {
        /// Aggregate package budget exceeded.
        Limit {
            /// Resource name (`parts` or `bytes`).
            resource: &'static str,
            /// Inclusive limit.
            limit: usize,
            /// First value beyond the limit.
            actual: usize,
        },
        /// One named XML part failed verification.
        Part {
            /// Borrowed archive-relative name.
            name: &'a str,
            /// Typed document failure.
            source: DocumentError,
        },
    }

    impl fmt::Display for Error<'_> {
        fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
            match self {
                Self::Limit {
                    resource,
                    limit,
                    actual,
                } => write!(
                    formatter,
                    "XML package {resource} limit {limit} exceeded by {actual}"
                ),
                Self::Part { name, source } => write!(formatter, "XML part '{name}': {source}"),
            }
        }
    }

    impl std::error::Error for Error<'_> {
        fn source(&self) -> Option<&(dyn std::error::Error + 'static)> {
            match self {
                Self::Limit { .. } => None,
                Self::Part { source, .. } => Some(source),
            }
        }
    }

    /// Verifies a sequence of generated OOXML/ODF or referenced XML assets.
    ///
    /// Parts are consumed incrementally and payloads remain borrowed.
    ///
    /// # Errors
    ///
    /// Returns [`Error::Limit`] for an aggregate package-budget breach or
    /// [`Error::Part`] with the borrowed name and typed document failure.
    pub fn verify<'a, I>(parts: I, limits: Limits) -> Result<Report, Error<'a>>
    where
        I: IntoIterator<Item = Part<'a>>,
    {
        let mut report = Report {
            attributes: 0,
            bytes: 0,
            events: 0,
            max_depth: 0,
            parts: 0,
            text_bytes: 0,
        };

        for part in parts {
            report.parts = package_add(report.parts, 1, "parts", limits.parts)?;
            report.bytes = package_add(report.bytes, part.bytes.len(), "bytes", limits.bytes)?;
            let item =
                verify_document(part.bytes, limits.document).map_err(|source| Error::Part {
                    name: part.name,
                    source,
                })?;
            report.attributes = report.attributes.saturating_add(item.attributes());
            report.events = report.events.saturating_add(item.events());
            report.max_depth = report.max_depth.max(item.max_depth());
            report.text_bytes = report.text_bytes.saturating_add(item.text_bytes());
        }
        Ok(report)
    }

    fn package_add<'a>(
        current: usize,
        amount: usize,
        resource: &'static str,
        limit: usize,
    ) -> Result<usize, Error<'a>> {
        let actual = current.saturating_add(amount);
        if actual > limit {
            return Err(Error::Limit {
                resource,
                limit,
                actual,
            });
        }
        Ok(actual)
    }
}

#[cfg(test)]
mod stream_tests {
    use super::*;

    struct Chunked<'a> {
        input: &'a [u8],
        position: usize,
        chunk: usize,
    }

    impl<'a> Chunked<'a> {
        fn new(input: &'a [u8], chunk: usize) -> Self {
            Self {
                input,
                position: 0,
                chunk: chunk.max(1),
            }
        }
    }

    impl Read for Chunked<'_> {
        fn read(&mut self, output: &mut [u8]) -> io::Result<usize> {
            let available = self.fill_buf()?;
            let amount = available.len().min(output.len());
            output[..amount].copy_from_slice(&available[..amount]);
            self.consume(amount);
            Ok(amount)
        }
    }

    impl BufRead for Chunked<'_> {
        fn fill_buf(&mut self) -> io::Result<&[u8]> {
            let end = self
                .position
                .saturating_add(self.chunk)
                .min(self.input.len());
            Ok(&self.input[self.position..end])
        }

        fn consume(&mut self, amount: usize) {
            self.position = self.position.saturating_add(amount).min(self.input.len());
        }
    }

    struct Failing;

    impl Read for Failing {
        fn read(&mut self, _output: &mut [u8]) -> io::Result<usize> {
            Err(io::Error::other("test source failure"))
        }
    }

    impl BufRead for Failing {
        fn fill_buf(&mut self) -> io::Result<&[u8]> {
            Err(io::Error::other("test source failure"))
        }

        fn consume(&mut self, _amount: usize) {}
    }

    fn limits(input: &[u8]) -> Limits {
        Limits::new(input.len(), 32, 256, 256, 256, input.len()).expect("finite test limits")
    }

    #[test]
    fn reader_matches_slice_report_across_one_byte_chunks() {
        let xml = b"<?xml version=\"1.0\"?><root xml:space=\"preserve\">\n<child>e\xC3\xA9</child><![CDATA[  ]]>&amp;&#32;<?keep x?></root>";
        let expected = verify(xml, limits(xml)).expect("slice XML is valid");
        let actual = verify_reader(Chunked::new(xml, 1), limits(xml))
            .expect("one-byte chunks must preserve XML semantics");
        assert_eq!(actual, expected);
    }

    #[test]
    fn reader_matches_slice_bom_admission() {
        let xml = b"\xEF\xBB\xBF<?xml version=\"1.0\"?><root/>";
        let expected = verify(xml, limits(xml)).expect("a UTF-8 BOM is valid XML");
        let actual = verify_reader(Chunked::new(xml, 1), limits(xml))
            .expect("streaming XML must preserve the slice framing rule");
        assert_eq!(actual, expected);
        assert_eq!(actual.bytes(), xml.len());
    }

    #[test]
    fn authored_reader_preserves_authored_whitespace_policy() {
        let xml = b"<root> <child/></root>";
        let error = verify_authored_reader(Chunked::new(xml, 2), limits(xml))
            .expect_err("ambiguous authored whitespace must be refused");
        assert!(matches!(
            error,
            StreamError::Audit(Error::NotCompact(violation))
                if violation.kind() == Kind::AmbiguousWhitespace
        ));

        let preserved = b"<root xml:space=\"preserve\"> <child/></root>";
        let _report = verify_authored_reader(Chunked::new(preserved, 1), limits(preserved))
            .expect("explicit xml:space must preserve the text run");
    }

    #[test]
    fn token_window_fails_before_unbounded_parser_growth() {
        let xml = b"<root>12345678901234567</root>";
        let profile = Limits::new(xml.len(), 8, 64, 64, 16, xml.len()).unwrap();
        let error = verify_reader(Chunked::new(xml, 1), profile)
            .expect_err("the text event exceeds the token window");
        assert!(matches!(
            error,
            StreamError::Audit(Error::Limit {
                resource: Resource::TokenBytes,
                limit: 16,
                actual: 17,
                ..
            })
        ));
    }

    #[test]
    fn total_window_fails_with_audit_limit() {
        let xml = b"<root/>";
        let profile = Limits::new(xml.len() - 1, 8, 64, 64, 64, 64).unwrap();
        let error = verify_reader(Chunked::new(xml, 1), profile)
            .expect_err("the final byte exceeds the total window");
        assert!(matches!(
            error,
            StreamError::Audit(Error::Limit {
                resource: Resource::Bytes,
                limit,
                actual,
                ..
            }) if limit == xml.len() - 1 && actual == xml.len()
        ));
    }

    #[test]
    fn source_failures_are_distinguished_from_audit_failures() {
        let error = verify_reader(Failing, Limits::default())
            .expect_err("the source error must be retained");
        assert!(
            matches!(error, StreamError::Input(source) if source.to_string() == "test source failure")
        );
    }

    #[test]
    fn streamed_encoding_offsets_include_markup_prefixes() {
        let xml = b"<root a=\"\xFF\"/>";
        let error = verify_reader(Chunked::new(xml, 1), limits(xml))
            .expect_err("invalid UTF-8 in an attribute must be rejected");
        assert!(matches!(
            error,
            StreamError::Audit(Error::Encoding { valid_up_to: 9 })
        ));
    }

    #[test]
    fn streaming_memory_bound_accounts_for_depth_and_token_windows() {
        let limits = Limits::new(1024, 4, 32, 32, 16, 1024).unwrap();
        let bound = limits
            .streaming_memory_upper_bound()
            .expect("finite profile must have a representable envelope");
        assert!(bound >= (limits.max_token_bytes() + 1) * 2);
        let deep = Limits::new(1024, 8, 32, 32, 16, 1024).unwrap();
        assert!(deep.streaming_memory_upper_bound().unwrap() > bound);
    }

    #[test]
    fn streaming_memory_bound_caps_per_event_attribute_scratch_at_token_window() {
        let token_bytes = 4096;
        let aggregate =
            Limits::new(1024, 4, 32, Limits::ATTRIBUTE_CEILING, token_bytes, 1024).unwrap();
        let token_cap = Limits::new(1024, 4, 32, token_bytes + 1, token_bytes, 1024).unwrap();

        assert_eq!(
            aggregate.streaming_memory_upper_bound(),
            token_cap.streaming_memory_upper_bound(),
            "aggregate attribute accounting must not size one-event scratch"
        );
    }

    #[test]
    fn aggregate_attribute_budget_remains_independent_of_token_scratch() {
        let mut xml = b"<r>".to_vec();
        for _ in 0..16 {
            xml.extend_from_slice(b"<x a=\"1\"/>");
        }
        xml.extend_from_slice(b"</r>");

        let limits = Limits::new(xml.len(), 2, 64, 16, 10, 0).unwrap();
        let report = verify_reader(Chunked::new(&xml, 1), limits)
            .expect("aggregate attributes across events must remain accepted");
        assert_eq!(report.attributes(), 16);
    }

    #[test]
    fn streaming_memory_bound_reports_checked_arithmetic_overflow() {
        let overflowing = Limits::bounded(0, 0, 0, usize::MAX, usize::MAX, 0);
        assert_eq!(
            overflowing.streaming_memory_upper_bound(),
            None,
            "token-window arithmetic must fail closed on usize overflow"
        );
    }
}
