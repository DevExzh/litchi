//! Positional, source-backed edits of one ordinary Unicode main-story
//! paragraph in a legacy Word document.
//!
//! This owner is deliberately smaller than the ordinary, owned body editor.
//! It does not rebuild a FIB, CLX, or any property table.  A transaction is
//! admitted only when one complete paragraph (including its paragraph mark)
//! is represented by one uncompressed Unicode piece and the replacement has
//! exactly the same UTF-16 width.  Publication is then a checked equal-length
//! splice in the existing `WordDocument` stream.

use crate::{sprm::parse_sprms, sprm_operations as ops};
use litchi_cfb::{
    ArtifactFingerprint, OleError, OverlayError, PublishReport, SameLengthStreamSplice,
    SharedOleFile, SharedOleFileLimits, StreamSpliceLimits, ValidatedOverlayPlan,
    consts::{STGTY_ROOT, STGTY_STORAGE, STGTY_STREAM},
};
use litchi_core::{OwnedSource, Position, ReadAt, SourceVersion};
use std::{
    fmt,
    io::{self, Write},
    sync::Arc,
};

const WORD_DOCUMENT: &str = "WordDocument";
const TABLE_0: &str = "0Table";
const TABLE_1: &str = "1Table";

const FIB_MAGIC_WORD97: u16 = 0xA5EC;
const FIB_MAGIC_WORD95: u16 = 0xA5DC;
const WORD97_NFIB: u16 = 0x00C1;
const FIB_BASE_PREFIX: usize = 154;
const FIB_POINTER_COUNT: usize = 152;
const FIB_POINTERS: usize = 154;
const FIB_CCP_TEXT: usize = 76;
const FIB_CCP_MACRO: usize = 88;
const FIB_FC_MIN: usize = 24;
const FIB_FC_MAC: usize = 28;
const FIB_FLAGS: usize = 10;
const FIB_NFIB: usize = 2;
const FIB_FLAG_ENCRYPTED: u16 = 0x0100;
const FIB_FLAG_TABLE_1: u16 = 0x0200;
const FIB_FLAG_OBFUSCATED: u16 = 0x8000;

const PLCFBTE_CHPX: usize = 12;
const PLCFBTE_PAPX: usize = 13;
const PLCFFLD_MOM: usize = 16;
const DOP: usize = 31;
const CLX: usize = 33;
const STTBFRMARK: usize = 51;
const DGG_INFO: usize = 50;
const PROTECTION_POINTERS: [usize; 4] = [141, 142, 143, 144];

const DOP_MINIMUM: usize = 84;
const FKP_PAGE_BYTES: usize = 512;
const FKP_PAGE_MASK: u32 = 0x003F_FFFF;
const MAX_FKP_RUNS: usize = 101;
const SCAN_CHUNK_UNITS: usize = 4_096;

const DEFAULT_MAX_FIB_BYTES: usize = 64 * 1024;
const DEFAULT_MAX_CLX_BYTES: usize = 16 * 1024 * 1024;
const DEFAULT_MAX_TABLE_BYTES: usize = 16 * 1024 * 1024;
const DEFAULT_MAX_PREFIX_UNITS: usize = 16 * 1024 * 1024;
const DEFAULT_MAX_REPLACEMENT_UNITS: usize = 1024 * 1024;
const DEFAULT_MAX_PIECES: usize = 65_536;

/// A refusal reason specific to the narrow source-backed paragraph contract.
#[non_exhaustive]
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum Refusal {
    /// The selected paragraph position does not exist.
    ParagraphNotFound,
    /// The paragraph has no replaceable text bytes.
    EmptyParagraph,
    /// The replacement changes the encoded UTF-16 unit count.
    LengthChange,
    /// The selected text is in an ANSI/compressed piece.
    CompressedPiece,
    /// The paragraph crosses a text-piece boundary.
    CrossPiece,
    /// A structural character is present in the selected paragraph.
    StructuralContent,
    /// A field marker or field property is present.
    Field,
    /// Drawing, picture, or object data is present.
    Drawing,
    /// Revision metadata or tracked-revision properties are present.
    Revision,
    /// The CLX contains a fast-save property run.
    FastSave,
    /// A piece carries an unresolved property modifier.
    Prm,
    /// The FIB declares encryption or obfuscation.
    Encrypted,
    /// Document or range protection is active.
    Protected,
    /// A macro-bearing storage or stream is present.
    Macro,
    /// A signature-bearing storage or stream is present.
    Signed,
    /// The source topology is not uniquely resolvable by this owner.
    AmbiguousTopology,
    /// The source is a valid OLE container but outside this owner’s policy.
    UnsupportedSource,
}

impl fmt::Display for Refusal {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        let message = match self {
            Self::ParagraphNotFound => "DOC paragraph position was not found",
            Self::EmptyParagraph => "DOC paragraph has no replaceable text",
            Self::LengthChange => "source-backed DOC paragraph replacement changes UTF-16 width",
            Self::CompressedPiece => "source-backed DOC paragraph is in a compressed piece",
            Self::CrossPiece => "source-backed DOC paragraph crosses text pieces",
            Self::StructuralContent => "DOC paragraph contains unsupported structural content",
            Self::Field => "DOC field content is outside the source-backed closure",
            Self::Drawing => "DOC drawing or object content is outside the source-backed closure",
            Self::Revision => "DOC revision content is outside the source-backed closure",
            Self::FastSave => "DOC fast-save CLX property runs are unsupported",
            Self::Prm => "DOC piece property modifiers are unsupported",
            Self::Encrypted => "encrypted or obfuscated DOC sources are unsupported",
            Self::Protected => "protected DOC sources are unsupported",
            Self::Macro => "macro-bearing DOC sources are unsupported",
            Self::Signed => "signed DOC sources are unsupported",
            Self::AmbiguousTopology => "DOC source topology is ambiguous",
            Self::UnsupportedSource => "DOC source is outside the safe source-backed closure",
        };
        formatter.write_str(message)
    }
}

/// Failure returned by a source-backed paragraph operation.
#[non_exhaustive]
#[derive(Debug)]
pub enum Error {
    /// The validated CFB reader rejected the source.
    Ole(OleError),
    /// The common equal-length CFB planner rejected the operation.
    Overlay(OverlayError),
    /// The format-level closure refused the operation.
    Refused(Refusal),
    /// A bounded source or table range was malformed.
    InvalidData(String),
    /// A configured finite limit was exceeded.
    Limit(String),
    /// A bounded allocation could not be reserved.
    Allocation {
        /// Bounded resource name.
        resource: &'static str,
        /// Allocation failure returned by the standard library.
        source: std::collections::TryReserveError,
    },
    /// A positional source or sink reported I/O failure.
    Io(io::Error),
}

impl fmt::Display for Error {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Self::Ole(error) => error.fmt(formatter),
            Self::Overlay(error) => error.fmt(formatter),
            Self::Refused(error) => error.fmt(formatter),
            Self::InvalidData(message) => write!(formatter, "invalid DOC source data: {message}"),
            Self::Limit(message) => {
                write!(formatter, "DOC source-backed limit exceeded: {message}")
            },
            Self::Allocation { resource, source } => {
                write!(formatter, "could not reserve {resource}: {source}")
            },
            Self::Io(error) => error.fmt(formatter),
        }
    }
}

impl std::error::Error for Error {
    fn source(&self) -> Option<&(dyn std::error::Error + 'static)> {
        match self {
            Self::Ole(error) => Some(error),
            Self::Overlay(error) => Some(error),
            Self::Allocation { source, .. } => Some(source),
            Self::Io(error) => Some(error),
            Self::Refused(_) | Self::InvalidData(_) | Self::Limit(_) => None,
        }
    }
}

impl From<OleError> for Error {
    fn from(error: OleError) -> Self {
        Self::Ole(error)
    }
}

impl From<OverlayError> for Error {
    fn from(error: OverlayError) -> Self {
        Self::Overlay(error)
    }
}

impl From<io::Error> for Error {
    fn from(error: io::Error) -> Self {
        Self::Io(error)
    }
}

/// Result type for the source-backed DOC paragraph owner.
pub type Result<T> = std::result::Result<T, Error>;

/// Finite limits for source-backed DOC parsing and publication.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct SourceLimits {
    max_input_bytes: u64,
    max_fib_bytes: usize,
    max_clx_bytes: usize,
    max_table_bytes: usize,
    max_prefix_units: usize,
    max_replacement_units: usize,
    max_pieces: usize,
    splice_limits: StreamSpliceLimits,
}

impl SourceLimits {
    /// Creates a finite source-backed DOC policy.
    pub fn new(
        max_input_bytes: u64,
        max_fib_bytes: usize,
        max_clx_bytes: usize,
        max_table_bytes: usize,
        max_prefix_units: usize,
        max_replacement_units: usize,
        max_pieces: usize,
        splice_limits: StreamSpliceLimits,
    ) -> Result<Self> {
        let limits = Self {
            max_input_bytes,
            max_fib_bytes,
            max_clx_bytes,
            max_table_bytes,
            max_prefix_units,
            max_replacement_units,
            max_pieces,
            splice_limits,
        };
        limits.validate()?;
        Ok(limits)
    }

    fn validate(self) -> Result<()> {
        if self.max_input_bytes == 0
            || self.max_fib_bytes == 0
            || self.max_clx_bytes == 0
            || self.max_table_bytes == 0
            || self.max_prefix_units == 0
            || self.max_replacement_units == 0
            || self.max_pieces == 0
        {
            return Err(Error::Limit(
                "source-backed DOC limits must be non-zero".into(),
            ));
        }
        if self.max_input_bytes > SharedOleFileLimits::MAX_INPUT_BYTES {
            return Err(Error::Limit(format!(
                "input limit {} exceeds CFB maximum {}",
                self.max_input_bytes,
                SharedOleFileLimits::MAX_INPUT_BYTES
            )));
        }
        if self.max_fib_bytes < FIB_BASE_PREFIX {
            return Err(Error::Limit(format!(
                "FIB limit {} is below the required {FIB_BASE_PREFIX}-byte prefix",
                self.max_fib_bytes
            )));
        }
        Ok(())
    }

    /// Maximum complete CFB input length.
    #[must_use]
    pub const fn max_input_bytes(self) -> u64 {
        self.max_input_bytes
    }

    /// Maximum FIB byte range retained during open.
    #[must_use]
    pub const fn max_fib_bytes(self) -> usize {
        self.max_fib_bytes
    }

    /// Maximum CLX byte range read during open.
    #[must_use]
    pub const fn max_clx_bytes(self) -> usize {
        self.max_clx_bytes
    }

    /// Maximum individual table range read during validation.
    #[must_use]
    pub const fn max_table_bytes(self) -> usize {
        self.max_table_bytes
    }

    /// Maximum number of main-story UTF-16 units scanned for one selector.
    #[must_use]
    pub const fn max_prefix_units(self) -> usize {
        self.max_prefix_units
    }

    /// Maximum UTF-16 units admitted in one replacement.
    #[must_use]
    pub const fn max_replacement_units(self) -> usize {
        self.max_replacement_units
    }

    /// Maximum number of CLX pieces retained.
    #[must_use]
    pub const fn max_pieces(self) -> usize {
        self.max_pieces
    }

    /// Common CFB splice limits used by this owner.
    #[must_use]
    pub const fn splice_limits(self) -> StreamSpliceLimits {
        self.splice_limits
    }

    /// Returns a copy with a different input ceiling.
    #[must_use]
    pub const fn with_max_input_bytes(mut self, value: u64) -> Self {
        self.max_input_bytes = value;
        self
    }

    /// Returns a copy with a different FIB ceiling.
    #[must_use]
    pub const fn with_max_fib_bytes(mut self, value: usize) -> Self {
        self.max_fib_bytes = value;
        self
    }

    /// Returns a copy with a different paragraph-prefix ceiling.
    #[must_use]
    pub const fn with_max_prefix_units(mut self, value: usize) -> Self {
        self.max_prefix_units = value;
        self
    }

    /// Returns a copy with a different replacement ceiling.
    #[must_use]
    pub const fn with_max_replacement_units(mut self, value: usize) -> Self {
        self.max_replacement_units = value;
        self
    }

    /// Returns a copy with different common splice limits.
    #[must_use]
    pub const fn with_splice_limits(mut self, value: StreamSpliceLimits) -> Self {
        self.splice_limits = value;
        self
    }
}

impl Default for SourceLimits {
    fn default() -> Self {
        Self {
            max_input_bytes: SharedOleFileLimits::MAX_INPUT_BYTES,
            max_fib_bytes: DEFAULT_MAX_FIB_BYTES,
            max_clx_bytes: DEFAULT_MAX_CLX_BYTES,
            max_table_bytes: DEFAULT_MAX_TABLE_BYTES,
            max_prefix_units: DEFAULT_MAX_PREFIX_UNITS,
            max_replacement_units: DEFAULT_MAX_REPLACEMENT_UNITS,
            max_pieces: DEFAULT_MAX_PIECES,
            splice_limits: StreamSpliceLimits::default(),
        }
    }
}

/// Options for a source-backed DOC paragraph snapshot.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub struct SourceBackedOptions {
    /// Finite source and publication limits.
    pub limits: SourceLimits,
}

impl SourceBackedOptions {
    /// Creates options with explicit source limits.
    #[must_use]
    pub const fn new(limits: SourceLimits) -> Self {
        Self { limits }
    }

    /// Returns the bounded source policy.
    #[must_use]
    pub const fn limits(self) -> SourceLimits {
        self.limits
    }
}

/// Alias retained for callers that spell the owner as a source transaction.
pub type SourceOptions = SourceBackedOptions;

/// One paragraph selected by its ordinary zero-based position.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct Paragraph {
    position: Position,
    text: String,
}

impl Paragraph {
    /// Position used to resolve this paragraph.
    #[must_use]
    pub const fn position(&self) -> Position {
        self.position
    }

    /// Decoded paragraph text excluding its terminating paragraph mark.
    #[must_use]
    pub fn text(&self) -> &str {
        &self.text
    }
}

/// Descriptive name for [`Paragraph`] in the source-backed API.
pub type SourceParagraph = Paragraph;

#[derive(Clone)]
struct SourceInner {
    source: Arc<dyn ReadAt>,
    shared: Arc<SharedOleFile>,
    version: SourceVersion,
    length: u64,
    fingerprint: ArtifactFingerprint,
    limits: SourceLimits,
    layout: Layout,
}

#[derive(Clone)]
struct Layout {
    table_name: String,
    ccp_text: u32,
    pieces: Vec<Piece>,
    chpx: Option<Pointer>,
    papx: Option<Pointer>,
}

#[derive(Clone, Copy, Debug)]
struct Piece {
    cp_start: u32,
    cp_end: u32,
    fc: u64,
    unicode: bool,
}

#[derive(Clone)]
struct ResolvedParagraph {
    position: Position,
    cp_start: u32,
    cp_end: u32,
    fc: u64,
    expected: Arc<[u8]>,
    text: String,
}

impl fmt::Debug for SourceInner {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("SourceInner")
            .field("length", &self.length)
            .field("version", &self.version)
            .field("fingerprint", &self.fingerprint)
            .finish_non_exhaustive()
    }
}

/// Immutable positional source snapshot for safe Unicode main-story edits.
#[derive(Clone)]
pub struct SourceSnapshot {
    inner: Arc<SourceInner>,
}

impl fmt::Debug for SourceSnapshot {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("SourceSnapshot")
            .field("length", &self.inner.length)
            .field("version", &self.inner.version)
            .field("fingerprint", &self.inner.fingerprint)
            .finish_non_exhaustive()
    }
}

impl SourceSnapshot {
    /// Opens a positional DOC source with the default finite limits.
    pub fn open(source: Arc<dyn ReadAt>) -> Result<Self> {
        Self::open_with_options(source, SourceBackedOptions::default())
    }

    /// Opens a positional DOC source under explicit finite limits.
    pub fn open_with_limits(source: Arc<dyn ReadAt>, limits: SourceLimits) -> Result<Self> {
        Self::open_with_options(source, SourceBackedOptions::new(limits))
    }

    /// Opens a positional DOC source under explicit options.
    pub fn open_with_options(
        source: Arc<dyn ReadAt>,
        options: SourceBackedOptions,
    ) -> Result<Self> {
        let limits = options.limits;
        limits.validate()?;
        let length = source.len()?;
        if length > limits.max_input_bytes {
            return Err(Error::Limit(format!(
                "source length {length} exceeds {}",
                limits.max_input_bytes
            )));
        }
        let shared_limits = SharedOleFileLimits::new(limits.max_input_bytes)?;
        let version = source.version()?;
        let first_shared = Arc::new(SharedOleFile::open_with_limits(
            Arc::clone(&source),
            shared_limits,
        )?);
        ensure_source_identity(&source, version, length)?;
        if first_shared.source_version()? != version {
            return Err(Error::Overlay(OverlayError::SourceChanged {
                expected: version,
                observed: first_shared.source_version()?,
            }));
        }
        let opening_fingerprint = identity_fingerprint(&first_shared, limits)?;
        ensure_source_identity(&source, version, length)?;
        // Reopen after the first complete fingerprint.  This binds the
        // directory/FAT index used below to the same bytes that supplied the
        // opening identity even when a custom adapter lies about its token.
        let shared = Arc::new(SharedOleFile::open_with_limits(
            Arc::clone(&source),
            shared_limits,
        )?);
        ensure_source_identity(&source, version, length)?;
        if shared.source_version()? != version {
            return Err(Error::Overlay(OverlayError::SourceChanged {
                expected: version,
                observed: shared.source_version()?,
            }));
        }
        reject_container_components(&shared)?;
        let (table_name, fib) = read_fib(&shared, limits)?;
        reject_fib_policy(&shared, &table_name, &fib, limits)?;
        let clx_pointer = fib
            .pointer(CLX)
            .ok_or(Error::Refused(Refusal::UnsupportedSource))?;
        if clx_pointer.length == 0 {
            return Err(Error::Refused(Refusal::UnsupportedSource));
        }
        if clx_pointer.length > limits.max_clx_bytes as u64 {
            return Err(Error::Limit(format!(
                "CLX length {} exceeds {}",
                clx_pointer.length, limits.max_clx_bytes
            )));
        }
        let clx = read_table_range(
            &shared,
            &table_name,
            clx_pointer,
            limits.max_clx_bytes,
            "CLX",
        )?;
        let pieces = parse_clx(
            &clx,
            shared.stream_len(&[WORD_DOCUMENT])?,
            fib.ccp_text,
            fib.fc_min,
            fib.fc_mac,
            limits,
        )?;
        let final_fingerprint = identity_fingerprint(&shared, limits)?;
        ensure_source_identity(&source, version, length)?;
        if final_fingerprint != opening_fingerprint {
            return Err(Error::Overlay(OverlayError::SourceFingerprintChanged {
                expected: opening_fingerprint,
                observed: final_fingerprint,
            }));
        }
        let fingerprint = identity_fingerprint(&shared, limits)?;
        ensure_source_identity(&source, version, length)?;
        if fingerprint != final_fingerprint {
            return Err(Error::Overlay(OverlayError::SourceFingerprintChanged {
                expected: final_fingerprint,
                observed: fingerprint,
            }));
        }
        Ok(Self {
            inner: Arc::new(SourceInner {
                source,
                shared,
                version,
                length,
                fingerprint,
                limits,
                layout: Layout {
                    table_name,
                    ccp_text: fib.ccp_text,
                    pieces,
                    chpx: fib.pointer(PLCFBTE_CHPX),
                    papx: fib.pointer(PLCFBTE_PAPX),
                },
            }),
        })
    }

    /// Opens an owned byte vector after validating its CFB/DOC structure.
    pub fn from_bytes(bytes: Vec<u8>) -> Result<Self> {
        Self::open(Arc::new(OwnedSource::new(bytes)))
    }

    /// Parses and validates a borrowed byte slice by taking one owned copy.
    pub fn parse(bytes: &[u8]) -> Result<Self> {
        let maximum = SourceLimits::default().max_input_bytes;
        let length = u64::try_from(bytes.len())
            .map_err(|_error| Error::Limit("borrowed source length does not fit u64".into()))?;
        if length > maximum {
            return Err(Error::Limit(format!(
                "borrowed source length {} exceeds {maximum}",
                bytes.len()
            )));
        }
        let mut owned = Vec::new();
        owned
            .try_reserve_exact(bytes.len())
            .map_err(|source| Error::Allocation {
                resource: "DOC borrowed source bytes",
                source,
            })?;
        owned.extend_from_slice(bytes);
        Self::from_bytes(owned)
    }

    /// Source version captured at open.
    #[must_use]
    pub fn source_version(&self) -> SourceVersion {
        self.inner.version
    }

    /// Complete source-artifact fingerprint captured at open.
    #[must_use]
    pub fn fingerprint(&self) -> ArtifactFingerprint {
        self.inner.fingerprint
    }

    /// Complete source byte length captured at open.
    #[must_use]
    pub fn len(&self) -> u64 {
        self.inner.length
    }

    /// Whether the source artifact is empty.
    #[must_use]
    pub fn is_empty(&self) -> bool {
        self.inner.length == 0
    }

    /// Returns the configured finite source limits.
    #[must_use]
    pub fn limits(&self) -> SourceLimits {
        self.inner.limits
    }

    /// Reads one paragraph using its zero-based main-story position.
    pub fn paragraph(&self, position: Position) -> Result<Paragraph> {
        let resolved = self.resolve(position)?;
        Ok(Paragraph {
            position,
            text: resolved.text,
        })
    }

    /// Starts an isolated source-backed transaction for one paragraph.
    pub fn edit_paragraph(&self, position: Position) -> Result<Transaction> {
        let resolved = self.resolve(position)?;
        Ok(Transaction {
            source: self.clone(),
            resolved,
            replacement: None,
        })
    }

    /// Alias for [`Self::edit_paragraph`].
    pub fn transaction(&self, position: Position) -> Result<Transaction> {
        self.edit_paragraph(position)
    }

    /// Short alias for [`Self::edit_paragraph`].
    pub fn edit(&self, position: Position) -> Result<Transaction> {
        self.edit_paragraph(position)
    }

    fn resolve(&self, position: Position) -> Result<ResolvedParagraph> {
        self.ensure_current()?;
        let resolved = resolve_paragraph(
            &self.inner.shared,
            &self.inner.layout,
            position,
            self.inner.limits,
        )?;
        reject_selected_property_revisions(
            &self.inner.shared,
            &self.inner.layout,
            &resolved,
            self.inner.limits,
        )?;
        self.ensure_current()?;
        Ok(resolved)
    }

    fn ensure_current(&self) -> Result<()> {
        ensure_source_identity(&self.inner.source, self.inner.version, self.inner.length)?;
        let observed = identity_fingerprint(&self.inner.shared, self.inner.limits)?;
        ensure_source_identity(&self.inner.source, self.inner.version, self.inner.length)?;
        if observed != self.inner.fingerprint {
            return Err(Error::Overlay(OverlayError::SourceFingerprintChanged {
                expected: self.inner.fingerprint,
                observed,
            }));
        }
        let confirmed = identity_fingerprint(&self.inner.shared, self.inner.limits)?;
        ensure_source_identity(&self.inner.source, self.inner.version, self.inner.length)?;
        if confirmed != observed {
            return Err(Error::Overlay(OverlayError::SourceFingerprintChanged {
                expected: observed,
                observed: confirmed,
            }));
        }
        Ok(())
    }
}

/// One isolated, equal-UTF-16-width paragraph replacement.
pub struct Transaction {
    source: SourceSnapshot,
    resolved: ResolvedParagraph,
    replacement: Option<String>,
}

impl fmt::Debug for Transaction {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("Transaction")
            .field("position", &self.resolved.position)
            .field("text_units", &self.resolved.text.encode_utf16().count())
            .field("staged", &self.replacement.is_some())
            .finish()
    }
}

impl Transaction {
    /// Source text currently read for the selected paragraph.
    #[must_use]
    pub fn text(&self) -> &str {
        &self.resolved.text
    }

    /// Selected ordinary paragraph position.
    #[must_use]
    pub const fn position(&self) -> Position {
        self.resolved.position
    }

    /// Alias for [`Self::position`].
    #[must_use]
    pub const fn target(&self) -> Position {
        self.position()
    }

    /// Stages a replacement only when UTF-16 width is unchanged.
    pub fn set_text(&mut self, value: impl Into<String>) -> Result<()> {
        let candidate = value.into();
        let units = candidate.encode_utf16().count();
        if units > self.source.inner.limits.max_replacement_units {
            return Err(Error::Limit(format!(
                "replacement units {units} exceed {}",
                self.source.inner.limits.max_replacement_units
            )));
        }
        if units != self.resolved.text.encode_utf16().count() {
            return Err(Error::Refused(Refusal::LengthChange));
        }
        self.replacement = Some(candidate);
        Ok(())
    }

    /// Alias for [`Self::set_text`].
    pub fn replace_paragraph(&mut self, value: impl Into<String>) -> Result<()> {
        self.set_text(value)
    }

    /// Publishes the checked same-length replacement.
    pub fn commit(self) -> Result<Commit> {
        let replacement = self
            .replacement
            .unwrap_or_else(|| self.resolved.text.clone());
        publish(self.source, self.resolved, replacement)
    }

    /// Alias emphasizing the source-backed publication boundary.
    pub fn commit_source_backed(self) -> Result<Commit> {
        self.commit()
    }

    /// Discards staged state and returns the immutable source snapshot.
    #[must_use]
    pub fn rollback(self) -> SourceSnapshot {
        self.source
    }
}

/// Content-free evidence for one source-backed paragraph publication.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct Diagnostics {
    position: Position,
    changed_spans: usize,
    replacement_bytes: u64,
    source_bytes: u64,
    source_fingerprint: ArtifactFingerprint,
    target_fingerprint: ArtifactFingerprint,
    source_version: SourceVersion,
    target_version: SourceVersion,
}

impl Diagnostics {
    /// Selected paragraph position.
    #[must_use]
    pub const fn position(self) -> Position {
        self.position
    }

    /// Alias for [`Self::position`].
    #[must_use]
    pub const fn target(self) -> Position {
        self.position
    }

    /// Number of physical CFB spans changed by the splice.
    #[must_use]
    pub const fn changed_spans(self) -> usize {
        self.changed_spans
    }

    /// Replacement bytes submitted to the CFB planner.
    #[must_use]
    pub const fn replacement_bytes(self) -> u64 {
        self.replacement_bytes
    }

    /// Complete source CFB length.
    #[must_use]
    pub const fn source_bytes(self) -> u64 {
        self.source_bytes
    }

    /// Exact source fingerprint used before publication.
    #[must_use]
    pub const fn source_fingerprint(self) -> ArtifactFingerprint {
        self.source_fingerprint
    }

    /// Exact target fingerprint produced by the plan.
    #[must_use]
    pub const fn target_fingerprint(self) -> ArtifactFingerprint {
        self.target_fingerprint
    }

    /// Source version checked before planning.
    #[must_use]
    pub const fn source_version(self) -> SourceVersion {
        self.source_version
    }

    /// Composed target version.
    #[must_use]
    pub const fn target_version(self) -> SourceVersion {
        self.target_version
    }
}

/// Alias used by the other source-backed format owners.
pub type SourceBackedDiagnostics = Diagnostics;

/// A published source-backed DOC target and its reversible patch.
pub struct Commit {
    snapshot: SourceSnapshot,
    patch: Patch,
    plan: ValidatedOverlayPlan,
    diagnostics: Diagnostics,
}

impl fmt::Debug for Commit {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("Commit")
            .field("snapshot", &self.snapshot)
            .field("diagnostics", &self.diagnostics)
            .finish_non_exhaustive()
    }
}

impl Commit {
    /// Reopened composed target snapshot.
    #[must_use]
    pub const fn snapshot(&self) -> &SourceSnapshot {
        &self.snapshot
    }

    /// Exact-source reversible patch.
    #[must_use]
    pub const fn patch(&self) -> &Patch {
        &self.patch
    }

    /// Content-free publication diagnostics.
    #[must_use]
    pub const fn diagnostics(&self) -> Diagnostics {
        self.diagnostics
    }

    /// Whether the transaction changed no artifact bytes.
    #[must_use]
    pub fn is_noop(&self) -> bool {
        self.plan.is_noop()
    }

    /// Streams the complete validated target to a sequential sink.
    pub fn write_to<W: Write>(&self, writer: &mut W) -> Result<PublishReport> {
        self.plan.write_to(writer).map_err(Error::Overlay)
    }

    /// Alias emphasizing the forward-only stream boundary.
    pub fn publish_to_stream<W: Write>(&self, writer: &mut W) -> Result<PublishReport> {
        self.write_to(writer)
    }

    /// Publishes through the common synced sibling-temp/atomic-rename path.
    pub fn save<P: AsRef<std::path::Path>>(&self, path: P) -> Result<PublishReport> {
        self.plan.save(path).map_err(Error::Overlay)
    }
}

/// A source-checked reversible equal-length DOC paragraph splice.
#[derive(Clone)]
pub struct Patch {
    before: SourceSnapshot,
    after: SourceSnapshot,
    operation: Option<Operation>,
}

impl fmt::Debug for Patch {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("Patch")
            .field("source_version", &self.before.source_version())
            .field("target_version", &self.after.source_version())
            .field("has_operation", &self.operation.is_some())
            .finish()
    }
}

impl Patch {
    /// Exact source snapshot required for forward application.
    #[must_use]
    pub const fn source(&self) -> &SourceSnapshot {
        &self.before
    }

    /// Exact target snapshot produced by forward application.
    #[must_use]
    pub const fn target(&self) -> &SourceSnapshot {
        &self.after
    }

    /// Exact source-artifact fingerprint required for application.
    #[must_use]
    pub fn source_fingerprint(&self) -> ArtifactFingerprint {
        self.before.fingerprint()
    }

    /// Exact target-artifact fingerprint produced by application.
    #[must_use]
    pub fn target_fingerprint(&self) -> ArtifactFingerprint {
        self.after.fingerprint()
    }

    /// Whether this patch changes no bytes.
    #[must_use]
    pub fn is_empty(&self) -> bool {
        self.operation.is_none()
    }

    /// Applies only to the exact source version and fingerprint captured by
    /// this patch.
    pub fn apply(&self, current: &SourceSnapshot) -> Result<SourceSnapshot> {
        current.ensure_current()?;
        if current.source_version() != self.before.source_version() {
            return Err(Error::Overlay(OverlayError::SourceChanged {
                expected: self.before.source_version(),
                observed: current.source_version(),
            }));
        }
        if current.fingerprint() != self.before.fingerprint() {
            return Err(Error::Overlay(OverlayError::SourceFingerprintChanged {
                expected: self.before.fingerprint(),
                observed: current.fingerprint(),
            }));
        }
        let Some(operation) = self.operation.clone() else {
            return Ok(current.clone());
        };
        let commit = publish_operation(current.clone(), operation)?;
        if commit.snapshot.fingerprint() != self.after.fingerprint() {
            return Err(Error::Overlay(OverlayError::TargetFingerprintChanged {
                expected: self.after.fingerprint(),
                observed: commit.snapshot.fingerprint(),
            }));
        }
        Ok(commit.snapshot)
    }

    /// Returns the exact-source-checked inverse patch.
    #[must_use]
    pub fn inverse(&self) -> Self {
        Self {
            before: self.after.clone(),
            after: self.before.clone(),
            operation: self.operation.as_ref().map(Operation::inverse),
        }
    }
}

/// Compatibility aliases matching the names used by other source owners.
pub type SourceTransaction = Transaction;
pub type SourceBackedCommit = Commit;
pub type SourceBackedPatch = Patch;

#[derive(Clone)]
struct Operation {
    position: Position,
    cp_start: u32,
    cp_end: u32,
    fc: u64,
    expected: Arc<[u8]>,
    replacement: Arc<[u8]>,
    before_text: String,
    after_text: String,
}

impl Operation {
    fn inverse(&self) -> Self {
        Self {
            position: self.position,
            cp_start: self.cp_start,
            cp_end: self.cp_end,
            fc: self.fc,
            expected: Arc::clone(&self.replacement),
            replacement: Arc::clone(&self.expected),
            before_text: self.after_text.clone(),
            after_text: self.before_text.clone(),
        }
    }
}

fn publish(
    source: SourceSnapshot,
    resolved: ResolvedParagraph,
    replacement: String,
) -> Result<Commit> {
    if replacement == resolved.text {
        return publish_noop(source, resolved.position);
    }
    let replacement_units = replacement.encode_utf16().count();
    if replacement_units > source.inner.limits.max_replacement_units {
        return Err(Error::Limit(format!(
            "replacement units {replacement_units} exceed {}",
            source.inner.limits.max_replacement_units
        )));
    }
    let expected_units = resolved.text.encode_utf16().count();
    if replacement_units != expected_units {
        return Err(Error::Refused(Refusal::LengthChange));
    }
    let replacement_bytes = encode_utf16_le(&replacement)?;
    if replacement_bytes.len() != resolved.expected.len() {
        return Err(Error::InvalidData(
            "UTF-16 replacement width disagrees with the selected source range".into(),
        ));
    }
    let operation = Operation {
        position: resolved.position,
        cp_start: resolved.cp_start,
        cp_end: resolved.cp_end,
        fc: resolved.fc,
        expected: Arc::clone(&resolved.expected),
        replacement: Arc::from(replacement_bytes),
        before_text: resolved.text,
        after_text: replacement,
    };
    publish_operation(source, operation)
}

fn publish_noop(source: SourceSnapshot, position: Position) -> Result<Commit> {
    source.ensure_current()?;
    let plan = source
        .inner
        .shared
        .plan_same_length_stream_splices(Vec::new(), source.inner.limits.splice_limits)?;
    let observed = plan.source_fingerprint();
    if observed != source.fingerprint() {
        return Err(Error::Overlay(OverlayError::SourceFingerprintChanged {
            expected: source.fingerprint(),
            observed,
        }));
    }
    source.ensure_current()?;
    let diagnostics = Diagnostics {
        position,
        changed_spans: plan.changed_spans(),
        replacement_bytes: 0,
        source_bytes: source.len(),
        source_fingerprint: observed,
        target_fingerprint: plan.target_fingerprint(),
        source_version: source.source_version(),
        target_version: source.source_version(),
    };
    let patch = Patch {
        before: source.clone(),
        after: source.clone(),
        operation: None,
    };
    Ok(Commit {
        snapshot: source,
        patch,
        plan,
        diagnostics,
    })
}

fn publish_operation(source: SourceSnapshot, operation: Operation) -> Result<Commit> {
    source.ensure_current()?;
    let current = source.resolve(operation.position)?;
    if current.cp_start != operation.cp_start
        || current.cp_end != operation.cp_end
        || current.fc != operation.fc
        || current.expected.as_ref() != operation.expected.as_ref()
        || current.text != operation.before_text
    {
        return Err(Error::Overlay(OverlayError::PreconditionFailed {
            path: vec![WORD_DOCUMENT.to_string()],
            offset: operation.fc,
        }));
    }
    let plan = source.inner.shared.plan_same_length_stream_splices(
        vec![SameLengthStreamSplice::new(
            vec![WORD_DOCUMENT.to_string()],
            operation.fc,
            Arc::clone(&operation.expected),
            Arc::clone(&operation.replacement),
        )],
        source.inner.limits.splice_limits,
    )?;
    let observed_source = plan.source_fingerprint();
    if observed_source != source.fingerprint() {
        return Err(Error::Overlay(OverlayError::SourceFingerprintChanged {
            expected: source.fingerprint(),
            observed: observed_source,
        }));
    }
    source.ensure_current()?;
    let composed = plan.composed_source()?;
    let candidate_source: Arc<dyn ReadAt> = Arc::new(composed);
    let candidate =
        SourceSnapshot::open_with_limits(Arc::clone(&candidate_source), source.inner.limits)?;
    let candidate_paragraph = candidate.resolve(operation.position)?;
    if candidate_paragraph.cp_start != operation.cp_start
        || candidate_paragraph.cp_end != operation.cp_end
        || candidate_paragraph.fc != operation.fc
        || candidate_paragraph.expected.as_ref() != operation.replacement.as_ref()
        || candidate_paragraph.text != operation.after_text
    {
        return Err(Error::InvalidData(
            "candidate DOC source failed paragraph readback".into(),
        ));
    }
    if candidate.fingerprint() != plan.target_fingerprint() {
        return Err(Error::Overlay(OverlayError::TargetFingerprintChanged {
            expected: plan.target_fingerprint(),
            observed: candidate.fingerprint(),
        }));
    }
    let target_version = candidate.source_version();
    let replacement_bytes = u64::try_from(operation.replacement.len())
        .map_err(|_error| Error::Limit("replacement byte count does not fit u64".into()))?;
    let diagnostics = Diagnostics {
        position: operation.position,
        changed_spans: plan.changed_spans(),
        replacement_bytes,
        source_bytes: source.len(),
        source_fingerprint: observed_source,
        target_fingerprint: plan.target_fingerprint(),
        source_version: source.source_version(),
        target_version,
    };
    let patch = Patch {
        before: source,
        after: candidate.clone(),
        operation: Some(operation),
    };
    Ok(Commit {
        snapshot: candidate,
        patch,
        plan,
        diagnostics,
    })
}

#[derive(Clone, Copy)]
struct Pointer {
    offset: u64,
    length: u64,
}

struct Fib {
    flags: u16,
    ccp_text: u32,
    ccp_macro: u32,
    fc_min: u32,
    fc_mac: u32,
    pointers: Vec<Pointer>,
}

impl Fib {
    fn pointer(&self, index: usize) -> Option<Pointer> {
        self.pointers.get(index).copied()
    }
}

fn read_fib(shared: &SharedOleFile, limits: SourceLimits) -> Result<(String, Fib)> {
    if limits.max_fib_bytes < FIB_BASE_PREFIX {
        return Err(Error::Limit(format!(
            "FIB limit {} is below the required {FIB_BASE_PREFIX}-byte prefix",
            limits.max_fib_bytes
        )));
    }
    let word_len = shared.stream_len(&[WORD_DOCUMENT])?;
    if word_len < FIB_BASE_PREFIX as u64 {
        return Err(Error::InvalidData(
            "WordDocument is shorter than the FIB prefix".into(),
        ));
    }
    let mut prefix = try_zeroed(FIB_BASE_PREFIX, "DOC FIB prefix")?;
    shared.read_stream_range(&[WORD_DOCUMENT], 0, &mut prefix)?;
    let magic = read_u16(&prefix, 0, "FIB magic")?;
    if magic != FIB_MAGIC_WORD97 && magic != FIB_MAGIC_WORD95 {
        return Err(Error::InvalidData(format!(
            "FIB magic 0x{magic:04X} is not a DOC FIB"
        )));
    }
    let nfib = read_u16(&prefix, FIB_NFIB, "FIB nFib")?;
    if nfib < WORD97_NFIB {
        return Err(Error::Refused(Refusal::UnsupportedSource));
    }
    let flags = read_u16(&prefix, FIB_FLAGS, "FIB flags")?;
    if flags & (FIB_FLAG_ENCRYPTED | FIB_FLAG_OBFUSCATED) != 0 {
        return Err(Error::Refused(Refusal::Encrypted));
    }
    if read_u32(&prefix, FIB_CCP_MACRO, "FIB ccpMcr")? != 0 {
        return Err(Error::Refused(Refusal::Macro));
    }
    let count = usize::from(read_u16(&prefix, FIB_POINTER_COUNT, "FIB pointer count")?);
    let pointer_bytes = count
        .checked_mul(8)
        .ok_or_else(|| Error::Limit("FIB pointer array size overflows usize".into()))?;
    let fib_len = FIB_POINTERS
        .checked_add(pointer_bytes)
        .ok_or_else(|| Error::Limit("FIB size overflows usize".into()))?;
    if fib_len > limits.max_fib_bytes {
        return Err(Error::Limit(format!(
            "FIB length {fib_len} exceeds {}",
            limits.max_fib_bytes
        )));
    }
    let fib_len_u64 = u64::try_from(fib_len)
        .map_err(|_error| Error::Limit("FIB length does not fit u64".into()))?;
    if fib_len_u64 > word_len {
        return Err(Error::InvalidData(
            "FIB pointer array exceeds WordDocument".into(),
        ));
    }
    let mut data = try_zeroed(fib_len, "DOC FIB")?;
    shared.read_stream_range(&[WORD_DOCUMENT], 0, &mut data)?;
    let mut pointers = Vec::new();
    pointers
        .try_reserve_exact(count)
        .map_err(|source| Error::Allocation {
            resource: "DOC FIB pointers",
            source,
        })?;
    for index in 0..count {
        let offset = FIB_POINTERS
            .checked_add(
                index
                    .checked_mul(8)
                    .ok_or_else(|| Error::Limit("FIB pointer offset overflows usize".into()))?,
            )
            .ok_or_else(|| Error::Limit("FIB pointer offset overflows usize".into()))?;
        let fc = u64::from(read_u32(&data, offset, "FIB pointer offset")?);
        let length = u64::from(read_u32(
            &data,
            offset
                .checked_add(4)
                .ok_or_else(|| Error::Limit("FIB pointer length offset overflows usize".into()))?,
            "FIB pointer length",
        )?);
        pointers.push(Pointer { offset: fc, length });
    }
    let ccp_text = read_u32(&data, FIB_CCP_TEXT, "FIB ccpText")?;
    let ccp_macro = read_u32(&data, FIB_CCP_MACRO, "FIB ccpMcr")?;
    let fc_min = read_u32(&data, FIB_FC_MIN, "FIB fcMin")?;
    let fc_mac = read_u32(&data, FIB_FC_MAC, "FIB fcMac")?;
    if fc_min > fc_mac {
        return Err(Error::InvalidData("FIB fcMin exceeds fcMac".into()));
    }
    if u64::from(fc_min) < fib_len_u64 {
        return Err(Error::InvalidData(
            "FIB fcMin overlaps the complete FIB".into(),
        ));
    }
    if u64::from(fc_mac) > word_len {
        return Err(Error::InvalidData("FIB fcMac exceeds WordDocument".into()));
    }
    let table_name = if flags & FIB_FLAG_TABLE_1 != 0 {
        TABLE_1.to_string()
    } else {
        TABLE_0.to_string()
    };
    if !shared.exists(&[&table_name]) {
        return Err(Error::Refused(Refusal::AmbiguousTopology));
    }
    Ok((
        table_name,
        Fib {
            flags,
            ccp_text,
            ccp_macro,
            fc_min,
            fc_mac,
            pointers,
        },
    ))
}

fn reject_fib_policy(
    shared: &SharedOleFile,
    table_name: &str,
    fib: &Fib,
    limits: SourceLimits,
) -> Result<()> {
    if fib.flags & (FIB_FLAG_ENCRYPTED | FIB_FLAG_OBFUSCATED) != 0 {
        return Err(Error::Refused(Refusal::Encrypted));
    }
    if fib.ccp_macro != 0 {
        return Err(Error::Refused(Refusal::Macro));
    }
    if fib
        .pointer(PLCFFLD_MOM)
        .is_some_and(|pointer| pointer.length != 0)
    {
        return Err(Error::Refused(Refusal::Field));
    }
    if fib
        .pointer(DGG_INFO)
        .is_some_and(|pointer| pointer.length != 0)
    {
        return Err(Error::Refused(Refusal::Drawing));
    }
    if fib
        .pointer(STTBFRMARK)
        .is_some_and(|pointer| pointer.length != 0)
    {
        return Err(Error::Refused(Refusal::Revision));
    }
    if PROTECTION_POINTERS.iter().copied().any(|index| {
        fib.pointer(index)
            .is_some_and(|pointer| pointer.length != 0)
    }) {
        return Err(Error::Refused(Refusal::Protected));
    }
    if let Some(pointer) = fib.pointer(DOP).filter(|pointer| pointer.length != 0) {
        if pointer.length < DOP_MINIMUM as u64 {
            return Err(Error::InvalidData(
                "DOP is shorter than its protected prefix".into(),
            ));
        }
        let dop = read_table_prefix(shared, table_name, pointer, DOP_MINIMUM, limits)?;
        if let Some(refusal) = classify_dop(&dop) {
            return Err(Error::Refused(refusal));
        }
    }
    Ok(())
}

fn classify_dop(dop: &[u8]) -> Option<Refusal> {
    if dop.len() < DOP_MINIMUM {
        return Some(Refusal::AmbiguousTopology);
    }
    let protected = dop[6] & 0x10 != 0
        || dop[7] & (0x02 | 0x20 | 0x40) != 0
        || i32::from_le_bytes([dop[78], dop[79], dop[80], dop[81]]) != 0;
    if protected {
        return Some(Refusal::Protected);
    }
    if dop[5] & 0x80 != 0 {
        return Some(Refusal::Revision);
    }
    None
}

fn reject_container_components(shared: &SharedOleFile) -> Result<()> {
    let mut saw_word = false;
    let mut saw_table_0 = false;
    let mut saw_table_1 = false;
    for entry in shared.directory_entries() {
        let name = entry.name.to_ascii_lowercase();
        if entry.entry_type == STGTY_STREAM {
            saw_word |= name == WORD_DOCUMENT.to_ascii_lowercase();
            saw_table_0 |= name == TABLE_0.to_ascii_lowercase();
            saw_table_1 |= name == TABLE_1.to_ascii_lowercase();
        }
        if name == "_xmlsignatures" || name == "_signatures" || name.contains("signature") {
            return Err(Error::Refused(Refusal::Signed));
        }
        if name == "macros"
            || name == "vba"
            || name == "_vba_project"
            || name == "dir"
            || name == "project"
            || name == "projectwm"
        {
            return Err(Error::Refused(Refusal::Macro));
        }
        if name == "objectpool" || name == "msodatastore" {
            return Err(Error::Refused(Refusal::Drawing));
        }
        if entry.entry_type == STGTY_STORAGE
            && entry.entry_type != STGTY_ROOT
            && entry.children.iter().any(|child| {
                let child_name = child.name.to_ascii_lowercase();
                child_name == "vba" || child_name == "macros" || child_name.contains("signature")
            })
        {
            return Err(Error::Refused(Refusal::UnsupportedSource));
        }
    }
    if !saw_word || (!saw_table_0 && !saw_table_1) {
        return Err(Error::Refused(Refusal::AmbiguousTopology));
    }
    Ok(())
}

fn parse_clx(
    data: &[u8],
    word_len: u64,
    ccp_text: u32,
    fc_min: u32,
    fc_mac: u32,
    limits: SourceLimits,
) -> Result<Vec<Piece>> {
    if fc_min > fc_mac || u64::from(fc_mac) > word_len {
        return Err(Error::Refused(Refusal::AmbiguousTopology));
    }
    if data.first() == Some(&1) {
        return Err(Error::Refused(Refusal::FastSave));
    }
    if data.first() != Some(&2) || data.len() < 5 {
        return Err(Error::Refused(Refusal::AmbiguousTopology));
    }
    let size = usize::try_from(read_u32(data, 1, "CLX PlcPcd size")?)
        .map_err(|_error| Error::Limit("CLX PlcPcd size does not fit usize".into()))?;
    if size < 4 || size.checked_add(5) != Some(data.len()) || !(size - 4).is_multiple_of(12) {
        return Err(Error::InvalidData("CLX PlcPcd shape is invalid".into()));
    }
    let count = (size - 4) / 12;
    if count == 0 || count > limits.max_pieces {
        return Err(Error::Limit(format!(
            "CLX piece count {count} exceeds {}",
            limits.max_pieces
        )));
    }
    let cp_base: usize = 5;
    let pcd_base = cp_base
        .checked_add(
            (count + 1)
                .checked_mul(4)
                .ok_or_else(|| Error::Limit("CLX CP array size overflows usize".into()))?,
        )
        .ok_or_else(|| Error::Limit("CLX PCD offset overflows usize".into()))?;
    let mut pieces = Vec::new();
    pieces
        .try_reserve_exact(count)
        .map_err(|source| Error::Allocation {
            resource: "DOC CLX pieces",
            source,
        })?;
    let mut previous_fc_end = 0_u64;
    for index in 0..count {
        let cp_offset = cp_base
            .checked_add(
                index
                    .checked_mul(4)
                    .ok_or_else(|| Error::Limit("CLX CP offset overflows usize".into()))?,
            )
            .ok_or_else(|| Error::Limit("CLX CP offset overflows usize".into()))?;
        let start = read_u32(data, cp_offset, "CLX CP start")?;
        let end = read_u32(
            data,
            cp_offset
                .checked_add(4)
                .ok_or_else(|| Error::Limit("CLX CP end offset overflows usize".into()))?,
            "CLX CP end",
        )?;
        if start >= end
            || (index == 0 && start != 0)
            || pieces
                .last()
                .is_some_and(|piece: &Piece| piece.cp_end != start)
        {
            return Err(Error::Refused(Refusal::AmbiguousTopology));
        }
        let pcd_offset = pcd_base
            .checked_add(
                index
                    .checked_mul(8)
                    .ok_or_else(|| Error::Limit("CLX PCD offset overflows usize".into()))?,
            )
            .ok_or_else(|| Error::Limit("CLX PCD offset overflows usize".into()))?;
        let raw = read_u32(
            data,
            pcd_offset
                .checked_add(2)
                .ok_or_else(|| Error::Limit("CLX FC offset overflows usize".into()))?,
            "CLX FC",
        )?;
        if raw & 0x8000_0000 != 0 {
            return Err(Error::Refused(Refusal::AmbiguousTopology));
        }
        let unicode = raw & 0x4000_0000 == 0;
        let encoded_fc = raw & 0x3FFF_FFFF;
        if unicode && encoded_fc & 1 != 0 {
            return Err(Error::Refused(Refusal::AmbiguousTopology));
        }
        if !unicode && encoded_fc & 1 != 0 {
            return Err(Error::Refused(Refusal::CompressedPiece));
        }
        let fc = if unicode {
            u64::from(encoded_fc)
        } else {
            u64::from(encoded_fc / 2)
        };
        let prm = read_u16(
            data,
            pcd_offset
                .checked_add(6)
                .ok_or_else(|| Error::Limit("CLX PRM offset overflows usize".into()))?,
            "CLX PRM",
        )?;
        if prm != 0 {
            return Err(Error::Refused(Refusal::Prm));
        }
        let cp_len = u64::from(end - start);
        let byte_len = cp_len
            .checked_mul(if unicode { 2 } else { 1 })
            .ok_or_else(|| Error::Limit("CLX piece byte length overflows u64".into()))?;
        let fc_end = fc
            .checked_add(byte_len)
            .ok_or_else(|| Error::Limit("CLX piece FC range overflows u64".into()))?;
        if fc < u64::from(fc_min)
            || fc_end > u64::from(fc_mac)
            || fc_end > word_len
            || (index != 0 && fc < previous_fc_end)
        {
            return Err(Error::Refused(Refusal::AmbiguousTopology));
        }
        previous_fc_end = fc_end;
        pieces.push(Piece {
            cp_start: start,
            cp_end: end,
            fc,
            unicode,
        });
    }
    if ccp_text > pieces.last().map_or(0, |piece| piece.cp_end) {
        return Err(Error::Refused(Refusal::AmbiguousTopology));
    }
    Ok(pieces)
}

fn resolve_paragraph(
    shared: &SharedOleFile,
    layout: &Layout,
    position: Position,
    limits: SourceLimits,
) -> Result<ResolvedParagraph> {
    let target = position.get();
    let mut paragraph_index = 0usize;
    let mut prefix_units = 0usize;
    let mut pending: Vec<u16> = Vec::new();
    let mut pending_bytes: Vec<u8> = Vec::new();
    let mut pending_start_cp = 0_u32;
    let mut pending_piece = usize::MAX;
    let scan_chunk_units = limits.max_prefix_units.min(SCAN_CHUNK_UNITS);
    let scan_chunk_bytes = scan_chunk_units
        .checked_mul(2)
        .ok_or_else(|| Error::Limit("DOC scan chunk byte length overflows usize".into()))?;
    let mut bytes = try_zeroed(scan_chunk_bytes, "DOC paragraph scan chunk")?;
    for (piece_index, piece) in layout.pieces.iter().enumerate() {
        let start = piece.cp_start;
        let end = piece.cp_end.min(layout.ccp_text);
        if start >= end {
            continue;
        }
        if !piece.unicode {
            return Err(Error::Refused(Refusal::CompressedPiece));
        }
        let mut cp = start;
        while cp < end {
            let piece_remaining = usize::try_from(end - cp)
                .map_err(|_error| Error::Limit("DOC piece CP range does not fit usize".into()))?;
            let prefix_remaining = limits
                .max_prefix_units
                .checked_sub(prefix_units)
                .ok_or_else(|| Error::Limit("DOC paragraph-prefix count overflows usize".into()))?;
            if prefix_remaining == 0 {
                return Err(Error::Limit(format!(
                    "paragraph-prefix units exceed {}",
                    limits.max_prefix_units
                )));
            }
            let count = piece_remaining.min(prefix_remaining).min(scan_chunk_units);
            let byte_len = count.checked_mul(2).ok_or_else(|| {
                Error::Limit("DOC Unicode scan chunk byte length overflows usize".into())
            })?;
            let fc = piece
                .fc
                .checked_add(
                    u64::from(cp - piece.cp_start)
                        .checked_mul(2)
                        .ok_or_else(|| Error::Limit("DOC piece FC offset overflows u64".into()))?,
                )
                .ok_or_else(|| Error::Limit("DOC piece FC offset overflows u64".into()))?;
            shared.read_stream_range(&[WORD_DOCUMENT], fc, &mut bytes[..byte_len])?;
            prefix_units = prefix_units
                .checked_add(count)
                .ok_or_else(|| Error::Limit("DOC paragraph-prefix count overflows usize".into()))?;
            for (unit_offset, unit_bytes) in bytes[..byte_len].chunks_exact(2).enumerate() {
                let cp = cp
                    .checked_add(u32::try_from(unit_offset).map_err(|_error| {
                        Error::Limit("DOC paragraph CP offset does not fit u32".into())
                    })?)
                    .ok_or_else(|| Error::Limit("DOC paragraph CP overflows u32".into()))?;
                let unit = u16::from_le_bytes([unit_bytes[0], unit_bytes[1]]);
                if pending.is_empty() {
                    pending_start_cp = cp;
                    pending_piece = piece_index;
                }
                reserve_pending(&mut pending, &mut pending_bytes, limits.max_prefix_units)?;
                pending.push(unit);
                pending_bytes.extend_from_slice(unit_bytes);
                if unit != 0x000D {
                    continue;
                }
                if paragraph_index == target {
                    if pending_piece != piece_index {
                        return Err(Error::Refused(Refusal::CrossPiece));
                    }
                    if pending.len() <= 1 {
                        return Err(Error::Refused(Refusal::EmptyParagraph));
                    }
                    let content = &pending[..pending.len() - 1];
                    if content.iter().any(|&value| is_structural_unit(value)) {
                        return Err(Error::Refused(Refusal::StructuralContent));
                    }
                    let text = decode_utf16(content)?;
                    let payload_len = content.len().checked_mul(2).ok_or_else(|| {
                        Error::Limit("DOC paragraph byte length overflows usize".into())
                    })?;
                    let expected = try_arc_bytes(
                        pending_bytes.get(..payload_len).ok_or_else(|| {
                            Error::InvalidData("DOC paragraph bytes are truncated".into())
                        })?,
                        "DOC selected paragraph bytes",
                    )?;
                    let fc = piece
                        .fc
                        .checked_add(
                            u64::from(pending_start_cp - piece.cp_start)
                                .checked_mul(2)
                                .ok_or_else(|| {
                                    Error::Limit("DOC paragraph FC offset overflows u64".into())
                                })?,
                        )
                        .ok_or_else(|| {
                            Error::Limit("DOC paragraph FC offset overflows u64".into())
                        })?;
                    return Ok(ResolvedParagraph {
                        position,
                        cp_start: pending_start_cp,
                        cp_end: cp,
                        fc,
                        expected,
                        text,
                    });
                }
                paragraph_index = paragraph_index
                    .checked_add(1)
                    .ok_or_else(|| Error::Limit("DOC paragraph index overflows usize".into()))?;
                pending.clear();
                pending_bytes.clear();
                pending_piece = usize::MAX;
            }
            cp = cp
                .checked_add(u32::try_from(count).map_err(|_error| {
                    Error::Limit("DOC paragraph CP offset does not fit u32".into())
                })?)
                .ok_or_else(|| Error::Limit("DOC paragraph CP overflows u32".into()))?;
        }
        if !pending.is_empty() && paragraph_index == target {
            return Err(Error::Refused(Refusal::CrossPiece));
        }
    }
    Err(Error::Refused(Refusal::ParagraphNotFound))
}

fn reject_selected_property_revisions(
    shared: &SharedOleFile,
    layout: &Layout,
    paragraph: &ResolvedParagraph,
    limits: SourceLimits,
) -> Result<()> {
    // DOP and STTBFRMark are checked during open.  The FKP check below covers
    // revision SPRMs attached only to the selected physical character range.
    for index in [PLCFBTE_CHPX, PLCFBTE_PAPX] {
        let Some(pointer) = read_pointer_from_layout(layout, index)? else {
            continue;
        };
        if pointer.length == 0 {
            continue;
        }
        if pointer.length > limits.max_table_bytes as u64 {
            return Err(Error::Limit(format!(
                "FKP BTE length {} exceeds {}",
                pointer.length, limits.max_table_bytes
            )));
        }
        let bte = read_table_range(
            shared,
            &layout.table_name,
            pointer,
            limits.max_table_bytes,
            "FKP BTE",
        )?;
        let (fcs, pages) = parse_bte(&bte)?;
        for (range, page_number) in fcs.windows(2).zip(pages) {
            let selected_start = paragraph.fc;
            let selected_end = paragraph
                .fc
                .checked_add(paragraph.expected.len() as u64)
                .ok_or_else(|| Error::Limit("selected DOC FC range overflows u64".into()))?;
            if u64::from(range[1]) <= selected_start || u64::from(range[0]) >= selected_end {
                continue;
            }
            let page_offset = u64::from(page_number & FKP_PAGE_MASK)
                .checked_mul(FKP_PAGE_BYTES as u64)
                .ok_or_else(|| Error::Limit("DOC FKP page offset overflows u64".into()))?;
            let mut page = try_zeroed(FKP_PAGE_BYTES, "DOC FKP page")?;
            shared.read_stream_range(&[WORD_DOCUMENT], page_offset, &mut page)?;
            if page[511] as usize == 0 || page[511] as usize > MAX_FKP_RUNS {
                return Err(Error::Refused(Refusal::AmbiguousTopology));
            }
            if fkp_has_revision_sprm(&page, index == PLCFBTE_CHPX, selected_start, selected_end)? {
                return Err(Error::Refused(Refusal::Revision));
            }
        }
    }
    Ok(())
}

fn read_pointer_from_layout(layout: &Layout, index: usize) -> Result<Option<Pointer>> {
    match index {
        PLCFBTE_CHPX => Ok(layout.chpx),
        PLCFBTE_PAPX => Ok(layout.papx),
        _ => Ok(None),
    }
}

fn fkp_has_revision_sprm(
    page: &[u8],
    chpx: bool,
    selected_start: u64,
    selected_end: u64,
) -> Result<bool> {
    let run_count = page
        .get(511)
        .copied()
        .map(usize::from)
        .ok_or_else(|| Error::InvalidData("DOC FKP page is truncated".into()))?;
    if run_count == 0 || run_count > MAX_FKP_RUNS {
        return Err(Error::Refused(Refusal::AmbiguousTopology));
    }
    let bx_size = if chpx { 1 } else { 13 };
    let bx_offset = (run_count + 1)
        .checked_mul(4)
        .ok_or_else(|| Error::Limit("DOC FKP BX offset overflows usize".into()))?;
    let bx_end = bx_offset
        .checked_add(
            run_count
                .checked_mul(bx_size)
                .ok_or_else(|| Error::Limit("DOC FKP BX range overflows usize".into()))?,
        )
        .ok_or_else(|| Error::Limit("DOC FKP BX range overflows usize".into()))?;
    if bx_end > page.len() {
        return Err(Error::Refused(Refusal::AmbiguousTopology));
    }
    for index in 0..run_count {
        let fc_start = u64::from(read_u32(page, index * 4, "DOC FKP FC")?);
        let fc_end = u64::from(read_u32(page, (index + 1) * 4, "DOC FKP FC")?);
        if fc_start >= fc_end {
            return Err(Error::Refused(Refusal::AmbiguousTopology));
        }
        if fc_end <= selected_start || fc_start >= selected_end {
            continue;
        }
        let bx = usize::from(page[bx_offset + index * bx_size]);
        if bx == 0 {
            continue;
        }
        let property_offset = bx
            .checked_mul(2)
            .ok_or_else(|| Error::Limit("DOC FKP property offset overflows usize".into()))?;
        if property_offset + 1 >= page.len() {
            return Err(Error::Refused(Refusal::AmbiguousTopology));
        }
        let first = usize::from(page[property_offset]);
        let (start, byte_count) = if chpx {
            (
                property_offset
                    .checked_add(1)
                    .ok_or_else(|| Error::Limit("DOC FKP property start overflows usize".into()))?,
                first,
            )
        } else if first == 0 {
            (
                property_offset.checked_add(2).ok_or_else(|| {
                    Error::Limit("DOC PAPX property start overflows usize".into())
                })?,
                usize::from(page[property_offset + 1])
                    .checked_mul(2)
                    .ok_or_else(|| {
                        Error::Limit("DOC PAPX property length overflows usize".into())
                    })?,
            )
        } else {
            (
                property_offset.checked_add(1).ok_or_else(|| {
                    Error::Limit("DOC PAPX property start overflows usize".into())
                })?,
                first
                    .checked_mul(2)
                    .and_then(|value| value.checked_sub(1))
                    .ok_or_else(|| {
                        Error::Limit("DOC PAPX property length overflows usize".into())
                    })?,
            )
        };
        let end = start
            .checked_add(byte_count)
            .ok_or_else(|| Error::Limit("DOC FKP property end overflows usize".into()))?;
        if end > page.len().saturating_sub(1) {
            return Err(Error::Refused(Refusal::AmbiguousTopology));
        }
        let mut grpprl = &page[start..end];
        if !chpx && !grpprl.is_empty() {
            if grpprl.len() < 2 {
                return Err(Error::Refused(Refusal::AmbiguousTopology));
            }
            grpprl = &grpprl[2..];
        }
        let sprms =
            parse_sprms(grpprl).map_err(|_error| Error::Refused(Refusal::AmbiguousTopology))?;
        if sprms.iter().any(|sprm| is_revision_opcode(sprm.opcode)) {
            return Ok(true);
        }
    }
    Ok(false)
}

fn is_revision_opcode(opcode: u16) -> bool {
    matches!(
        opcode,
        ops::SPRM_C_F_RMARK_DEL
            | ops::SPRM_C_F_RMARK
            | ops::SPRM_C_IBST_RMARK
            | ops::SPRM_C_DTTM_RMARK
            | ops::SPRM_C_IDSL_RMARK
            | ops::SPRM_C_RSID_PROP
            | ops::SPRM_C_RSID_TEXT
            | ops::SPRM_C_RSID_RM_DEL
            | ops::SPRM_C_PROP_RMARK
            | ops::SPRM_C_PROP_RMARK_CURRENT
            | ops::SPRM_C_IBST_RMARK_DEL
            | ops::SPRM_C_DTTM_RMARK_DEL
            | ops::SPRM_C_IDSL_RMARK_DEL
            | ops::SPRM_C_WALL
            | ops::SPRM_P_PROP_RMARK
            | ops::SPRM_P_WALL
            | ops::SPRM_P_RSID
            | ops::SPRM_P_PROP_RMARK90
            | ops::SPRM_P_PROP_RMARK_CURRENT
            | ops::SPRM_T_PROP_RMARK
            | ops::SPRM_T_WALL
            | ops::SPRM_T_RSID
    )
}

fn parse_bte(data: &[u8]) -> Result<(Vec<u32>, Vec<u32>)> {
    if data.len() < 12 || !(data.len() - 4).is_multiple_of(8) {
        return Err(Error::Refused(Refusal::AmbiguousTopology));
    }
    let count = (data.len() - 4) / 8;
    let mut fcs = Vec::new();
    fcs.try_reserve_exact(count + 1)
        .map_err(|source| Error::Allocation {
            resource: "DOC FKP FC boundaries",
            source,
        })?;
    let mut pages = Vec::new();
    pages
        .try_reserve_exact(count)
        .map_err(|source| Error::Allocation {
            resource: "DOC FKP page numbers",
            source,
        })?;
    for index in 0..=count {
        fcs.push(read_u32(data, index * 4, "FKP FC boundary")?);
    }
    for index in 0..count {
        pages.push(read_u32(data, (count + 1) * 4 + index * 4, "FKP page")? & FKP_PAGE_MASK);
    }
    if fcs.windows(2).any(|pair| pair[0] >= pair[1]) || pages.contains(&0) {
        return Err(Error::Refused(Refusal::AmbiguousTopology));
    }
    Ok((fcs, pages))
}

fn read_table_range(
    shared: &SharedOleFile,
    table_name: &str,
    pointer: Pointer,
    maximum: usize,
    resource: &'static str,
) -> Result<Vec<u8>> {
    let length = usize::try_from(pointer.length)
        .map_err(|_error| Error::Limit(format!("{resource} length does not fit usize")))?;
    if length > maximum {
        return Err(Error::Limit(format!(
            "{resource} length {length} exceeds {maximum}"
        )));
    }
    let mut data = try_zeroed(length, resource)?;
    shared.read_stream_range(&[table_name], pointer.offset, &mut data)?;
    Ok(data)
}

fn read_table_prefix(
    shared: &SharedOleFile,
    table_name: &str,
    pointer: Pointer,
    length: usize,
    limits: SourceLimits,
) -> Result<Vec<u8>> {
    let requested = length.min(
        usize::try_from(pointer.length)
            .map_err(|_error| Error::Limit("table prefix length does not fit usize".into()))?,
    );
    if requested > limits.max_table_bytes {
        return Err(Error::Limit(format!(
            "table prefix length {requested} exceeds {}",
            limits.max_table_bytes
        )));
    }
    let mut data = try_zeroed(requested, "DOC table prefix")?;
    shared.read_stream_range(&[table_name], pointer.offset, &mut data)?;
    Ok(data)
}

fn ensure_source_identity(
    source: &Arc<dyn ReadAt>,
    expected_version: SourceVersion,
    expected_length: u64,
) -> Result<()> {
    let observed = source.version()?;
    if observed != expected_version {
        return Err(Error::Overlay(OverlayError::SourceChanged {
            expected: expected_version,
            observed,
        }));
    }
    let length = source.len()?;
    if length != expected_length {
        return Err(Error::Overlay(OverlayError::Unavailable {
            reason: format!("source length changed from {expected_length} to {length}"),
        }));
    }
    let observed = source.version()?;
    if observed != expected_version {
        return Err(Error::Overlay(OverlayError::SourceChanged {
            expected: expected_version,
            observed,
        }));
    }
    Ok(())
}

fn identity_fingerprint(
    shared: &SharedOleFile,
    limits: SourceLimits,
) -> Result<ArtifactFingerprint> {
    let plan = shared.plan_same_length_stream_splices(Vec::new(), limits.splice_limits)?;
    Ok(plan.source_fingerprint())
}

fn reserve_pending(
    pending: &mut Vec<u16>,
    pending_bytes: &mut Vec<u8>,
    maximum_units: usize,
) -> Result<()> {
    if pending.len() >= maximum_units {
        return Err(Error::Limit(format!(
            "pending paragraph units exceed {maximum_units}"
        )));
    }
    let required_units = pending
        .len()
        .checked_add(1)
        .ok_or_else(|| Error::Limit("pending paragraph unit length overflows usize".into()))?;
    let required_bytes = required_units
        .checked_mul(2)
        .ok_or_else(|| Error::Limit("pending paragraph byte length overflows usize".into()))?;
    if pending.capacity() >= required_units && pending_bytes.capacity() >= required_bytes {
        return Ok(());
    }

    // Reserve a geometric target so a long paragraph copies its accumulated
    // units only O(log n) times.  The target is always capped by the caller's
    // finite unit limit, and both representations grow toward the same bound.
    let mut target_units = required_units;
    if pending.capacity() < required_units {
        target_units = target_units.max(pending.capacity().checked_mul(2).unwrap_or(usize::MAX));
    }
    let byte_capacity_units = pending_bytes.capacity() / 2;
    if byte_capacity_units < required_units {
        target_units = target_units.max(byte_capacity_units.checked_mul(2).unwrap_or(usize::MAX));
    }
    target_units = target_units.min(maximum_units);
    if target_units < required_units {
        return Err(Error::Limit(format!(
            "pending paragraph units exceed {maximum_units}"
        )));
    }

    if pending.capacity() < target_units {
        let additional_units = target_units - pending.capacity();
        pending
            .try_reserve_exact(additional_units)
            .map_err(|source| Error::Allocation {
                resource: "DOC pending paragraph units",
                source,
            })?;
    }
    let target_bytes = target_units
        .checked_mul(2)
        .ok_or_else(|| Error::Limit("pending paragraph byte growth overflows usize".into()))?;
    if pending_bytes.capacity() < target_bytes {
        let additional_bytes = target_bytes - pending_bytes.capacity();
        pending_bytes
            .try_reserve_exact(additional_bytes)
            .map_err(|source| Error::Allocation {
                resource: "DOC pending paragraph bytes",
                source,
            })?;
    }
    Ok(())
}

fn try_zeroed(length: usize, resource: &'static str) -> Result<Vec<u8>> {
    let mut data = Vec::new();
    data.try_reserve_exact(length)
        .map_err(|source| Error::Allocation { resource, source })?;
    data.resize(length, 0);
    Ok(data)
}

fn read_u16(data: &[u8], offset: usize, resource: &'static str) -> Result<u16> {
    let bytes = data
        .get(
            offset
                ..offset
                    .checked_add(2)
                    .ok_or_else(|| Error::Limit(format!("{resource} offset overflows")))?,
        )
        .ok_or_else(|| Error::InvalidData(format!("{resource} is truncated")))?;
    Ok(u16::from_le_bytes([bytes[0], bytes[1]]))
}

fn read_u32(data: &[u8], offset: usize, resource: &'static str) -> Result<u32> {
    let bytes = data
        .get(
            offset
                ..offset
                    .checked_add(4)
                    .ok_or_else(|| Error::Limit(format!("{resource} offset overflows")))?,
        )
        .ok_or_else(|| Error::InvalidData(format!("{resource} is truncated")))?;
    Ok(u32::from_le_bytes([bytes[0], bytes[1], bytes[2], bytes[3]]))
}

fn encode_utf16_le(value: &str) -> Result<Vec<u8>> {
    let units = value.encode_utf16().count();
    let bytes = units
        .checked_mul(2)
        .ok_or_else(|| Error::Limit("UTF-16 replacement size overflows usize".into()))?;
    let mut output = Vec::new();
    output
        .try_reserve_exact(bytes)
        .map_err(|source| Error::Allocation {
            resource: "DOC UTF-16 replacement",
            source,
        })?;
    for unit in value.encode_utf16() {
        output.extend_from_slice(&unit.to_le_bytes());
    }
    Ok(output)
}

fn decode_utf16(units: &[u16]) -> Result<String> {
    let capacity = units
        .len()
        .checked_mul(3)
        .ok_or_else(|| Error::Limit("UTF-16 paragraph size overflows usize".into()))?;
    let mut output = String::new();
    output
        .try_reserve_exact(capacity)
        .map_err(|source| Error::Allocation {
            resource: "DOC decoded paragraph text",
            source,
        })?;
    for decoded in char::decode_utf16(units.iter().copied()) {
        let character = decoded.map_err(|_error| {
            Error::InvalidData("selected DOC paragraph is not valid UTF-16".into())
        })?;
        output.push(character);
    }
    Ok(output)
}

fn try_arc_bytes(bytes: &[u8], resource: &'static str) -> Result<Arc<[u8]>> {
    let mut owned = Vec::new();
    owned
        .try_reserve_exact(bytes.len())
        .map_err(|source| Error::Allocation { resource, source })?;
    owned.extend_from_slice(bytes);
    Ok(Arc::from(owned))
}

fn is_structural_unit(value: u16) -> bool {
    (value < 0x0020 && value != 0x0009) || value == 0xFFFC
}

#[cfg(test)]
mod tests {
    #![allow(
        clippy::expect_used,
        clippy::unwrap_used,
        reason = "unit-test fixtures use contextual fail-fast assertions"
    )]

    use super::*;
    use std::sync::{
        Mutex,
        atomic::{AtomicBool, AtomicU64, AtomicUsize, Ordering},
    };

    #[test]
    fn utf16_replacement_preserves_surrogate_units() {
        let encoded = encode_utf16_le("a😀").expect("UTF-16 replacement");
        assert_eq!(encoded.len(), 6);
        assert_eq!(u16::from_le_bytes([encoded[2], encoded[3]]), 0xD83D);
        assert_eq!(u16::from_le_bytes([encoded[4], encoded[5]]), 0xDE00);
    }

    #[test]
    fn structural_units_are_closed() {
        assert!(is_structural_unit(0x0013));
        assert!(is_structural_unit(0xFFFC));
        assert!(!is_structural_unit(0x0009));
    }

    #[test]
    fn pending_capacity_grows_geometrically_and_stays_bounded() {
        const MAXIMUM_UNITS: usize = 1 << 16;
        let mut pending = Vec::new();
        let mut pending_bytes = Vec::new();
        let mut previous_capacity = 0;
        let mut growth_count = 0;
        for _ in 0..MAXIMUM_UNITS {
            reserve_pending(&mut pending, &mut pending_bytes, MAXIMUM_UNITS)
                .expect("bounded pending capacity");
            let capacity = pending.capacity();
            if capacity != previous_capacity {
                if previous_capacity != 0 {
                    assert!(capacity >= previous_capacity.saturating_mul(2));
                }
                growth_count += 1;
                previous_capacity = capacity;
            }
            pending.push(0);
            pending_bytes.extend_from_slice(&[0, 0]);
        }
        assert_eq!(pending.len(), MAXIMUM_UNITS);
        assert_eq!(pending_bytes.len(), MAXIMUM_UNITS * 2);
        assert!(growth_count <= 32);
    }

    #[test]
    fn ordinary_fixture_has_a_source_backed_paragraph() {
        let path = std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
            .join("../../test-data/ole/doc/documentProperties.doc");
        let bytes = std::fs::read(path).expect("DOC fixture");
        let source = SourceSnapshot::from_bytes(bytes).expect("source-backed DOC fixture");
        let paragraph = source.paragraph(Position::new(0)).expect("first paragraph");
        assert!(!paragraph.text().is_empty());
    }

    #[test]
    fn fixture_replacement_is_reversible_and_same_width() {
        let path = std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
            .join("../../test-data/ole/doc/documentProperties.doc");
        let bytes = std::fs::read(path).expect("DOC fixture");
        let source = SourceSnapshot::from_bytes(bytes).expect("source-backed DOC fixture");
        let original_fingerprint = source.fingerprint();
        let mut transaction = source
            .edit_paragraph(Position::new(0))
            .expect("first paragraph");
        let original = transaction.text().to_owned();
        let mut replacement = original.clone();
        if let Some((offset, character)) = original
            .char_indices()
            .find(|(_, value)| value.is_ascii_alphabetic())
        {
            let next = if character == 'a' { 'b' } else { 'a' };
            replacement.replace_range(offset..offset + character.len_utf8(), &next.to_string());
        } else {
            return;
        }
        transaction
            .set_text(replacement.clone())
            .expect("same UTF-16 width");
        let commit = transaction.commit().expect("source-backed commit");
        assert!(!commit.is_noop());
        assert_eq!(
            commit
                .snapshot()
                .paragraph(Position::new(0))
                .unwrap()
                .text(),
            replacement
        );
        let restored = commit
            .patch()
            .inverse()
            .apply(commit.snapshot())
            .expect("inverse");
        assert_eq!(restored.fingerprint(), original_fingerprint);
        assert_eq!(
            restored.paragraph(Position::new(0)).unwrap().text(),
            original
        );
    }

    #[test]
    fn exact_noop_keeps_snapshot_and_artifact_identity() {
        let path = std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
            .join("../../test-data/ole/doc/documentProperties.doc");
        let bytes = std::fs::read(path).expect("DOC fixture");
        let original = bytes.clone();
        let source = SourceSnapshot::from_bytes(bytes).expect("source-backed DOC fixture");
        let commit = source
            .edit_paragraph(Position::new(0))
            .expect("first paragraph")
            .commit()
            .expect("no-op commit");
        assert!(commit.is_noop());
        assert!(commit.patch().is_empty());
        assert_eq!(commit.snapshot().fingerprint(), source.fingerprint());
        assert!(Arc::ptr_eq(&commit.snapshot().inner, &source.inner));
        let mut output = Vec::new();
        commit.write_to(&mut output).expect("no-op output");
        assert_eq!(output, original);
    }

    #[test]
    fn length_change_and_partial_sink_are_typed_failures() {
        let path = std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
            .join("../../test-data/ole/doc/documentProperties.doc");
        let bytes = std::fs::read(path).expect("DOC fixture");
        let source = SourceSnapshot::from_bytes(bytes).expect("source-backed DOC fixture");
        let mut transaction = source
            .edit_paragraph(Position::new(0))
            .expect("first paragraph");
        let too_long = format!("{}x", transaction.text());
        assert_eq!(
            transaction.set_text(too_long).unwrap_err().to_string(),
            Refusal::LengthChange.to_string()
        );
        let commit = transaction
            .commit()
            .expect("unchanged transaction is a no-op");
        let mut sink = FailingSink { remaining: 9 };
        let error = commit.write_to(&mut sink).unwrap_err();
        assert!(matches!(
            error,
            Error::Overlay(OverlayError::IncompleteOutput { .. })
        ));
    }

    #[test]
    fn piece_relative_payload_readback_uses_a_later_paragraph() {
        let path = std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
            .join("../../test-data/ole/doc/lists-margins.doc");
        let bytes = std::fs::read(path).expect("DOC fixture");
        let source = SourceSnapshot::from_bytes(bytes).expect("source-backed DOC fixture");
        let mut transaction = source
            .edit_paragraph(Position::new(1))
            .expect("second paragraph");
        let original = transaction.text().to_owned();
        let Some((offset, character)) = original
            .char_indices()
            .find(|(_, value)| value.is_ascii_alphabetic())
        else {
            return;
        };
        let next = if character == 'a' { 'b' } else { 'a' };
        let mut replacement = original.clone();
        replacement.replace_range(offset..offset + character.len_utf8(), &next.to_string());
        transaction
            .set_text(replacement.clone())
            .expect("same width");
        let commit = transaction.commit().expect("second paragraph commit");
        assert_eq!(
            commit
                .snapshot()
                .paragraph(Position::new(1))
                .unwrap()
                .text(),
            replacement
        );
    }

    #[test]
    fn fib_limit_below_fixed_prefix_is_rejected_before_open() {
        let path = std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
            .join("../../test-data/ole/doc/documentProperties.doc");
        let bytes = std::fs::read(path).expect("DOC fixture");
        let limits = SourceLimits::default().with_max_fib_bytes(FIB_BASE_PREFIX - 1);
        assert!(matches!(
            SourceSnapshot::open_with_limits(Arc::new(OwnedSource::new(bytes)), limits),
            Err(Error::Limit(_))
        ));
    }

    #[test]
    fn zero_builder_limits_are_rejected_before_open() {
        let limits = SourceLimits::default().with_max_prefix_units(0);
        assert!(matches!(
            SourceSnapshot::open_with_limits(
                Arc::new(OwnedSource::new(Vec::new())),
                limits,
            ),
            Err(Error::Limit(message)) if message.contains("non-zero")
        ));
    }

    #[test]
    fn fib_fc_min_must_follow_the_complete_fib() {
        let mut writer = crate::writer::Writer::new();
        let text = "x".repeat(SCAN_CHUNK_UNITS * 2);
        writer.add_paragraph(&text).expect("DOC fixture");
        let mut output = io::Cursor::new(Vec::new());
        writer.write_to(&mut output).expect("DOC fixture");
        let mut bytes = output.into_inner();
        let mut file =
            litchi_cfb::OleFile::open(io::Cursor::new(bytes.clone())).expect("CFB fixture");
        let word = file
            .open_stream(&[WORD_DOCUMENT])
            .expect("WordDocument fixture");
        let count = usize::from(read_u16(&word, FIB_POINTER_COUNT, "FIB pointer count").unwrap());
        let fib_len = FIB_POINTERS
            .checked_add(count.checked_mul(8).expect("FIB size"))
            .expect("FIB size");
        let word_entry = file
            .list_directory_entries(&[])
            .expect("CFB directory")
            .into_iter()
            .find(|entry| entry.name == WORD_DOCUMENT)
            .expect("WordDocument entry");
        assert!(!word_entry.is_minifat);
        let sector_size = 1_usize << usize::from(u16::from_le_bytes([bytes[30], bytes[31]]));
        let physical = (u64::from(word_entry.start_sector) + 1)
            .checked_mul(u64::try_from(sector_size).expect("sector size"))
            .and_then(|offset| offset.checked_add(u64::try_from(FIB_FC_MIN).expect("offset")))
            .and_then(|offset| usize::try_from(offset).ok())
            .expect("WordDocument FIB offset");
        let malformed = u32::try_from(fib_len - 1).expect("FIB length");
        bytes[physical..physical + 4].copy_from_slice(&malformed.to_le_bytes());
        assert!(matches!(
            SourceSnapshot::from_bytes(bytes),
            Err(Error::InvalidData(message)) if message.contains("fcMin overlaps")
        ));
    }

    #[test]
    fn clx_fc_bounds_and_unicode_alignment_are_closed() {
        fn one_piece(raw_fc: u32, cp_end: u32) -> Vec<u8> {
            let mut data = vec![2, 16, 0, 0, 0];
            data.extend_from_slice(&0_u32.to_le_bytes());
            data.extend_from_slice(&cp_end.to_le_bytes());
            data.extend_from_slice(&0_u16.to_le_bytes());
            data.extend_from_slice(&raw_fc.to_le_bytes());
            data.extend_from_slice(&0_u16.to_le_bytes());
            data
        }

        let limits = SourceLimits::default();
        assert!(matches!(
            parse_clx(&one_piece(198, 2), 1_024, 2, 200, 300, limits),
            Err(Error::Refused(Refusal::AmbiguousTopology))
        ));
        assert!(matches!(
            parse_clx(&one_piece(298, 2), 1_024, 2, 200, 300, limits),
            Err(Error::Refused(Refusal::AmbiguousTopology))
        ));
        assert!(matches!(
            parse_clx(&one_piece(201, 1), 1_024, 1, 200, 300, limits),
            Err(Error::Refused(Refusal::AmbiguousTopology))
        ));
    }

    #[test]
    fn dop_protection_bit_40_is_not_misclassified_as_revision() {
        let mut dop = vec![0_u8; DOP_MINIMUM];
        dop[7] = 0x40;
        assert_eq!(classify_dop(&dop), Some(Refusal::Protected));
    }

    #[test]
    fn stable_token_mutation_is_rejected_before_paragraph_and_commit_access() {
        let path = std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
            .join("../../test-data/ole/doc/documentProperties.doc");
        let bytes = std::fs::read(path).expect("DOC fixture");
        let hostile = Arc::new(StableMutationSource {
            bytes: Mutex::new(bytes),
            armed: AtomicBool::new(false),
            fired: AtomicBool::new(false),
        });
        let source: Arc<dyn ReadAt> = hostile.clone();
        let snapshot = SourceSnapshot::open(source).expect("source-backed DOC fixture");
        hostile.armed.store(true, Ordering::SeqCst);
        assert!(matches!(
            snapshot.paragraph(Position::new(0)),
            Err(Error::Overlay(
                OverlayError::SourceFingerprintChanged { .. }
            ))
        ));
        assert!(hostile.fired.load(Ordering::SeqCst));

        let bytes = std::fs::read(
            std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
                .join("../../test-data/ole/doc/documentProperties.doc"),
        )
        .expect("DOC fixture");
        let hostile = Arc::new(StableMutationSource {
            bytes: Mutex::new(bytes),
            armed: AtomicBool::new(false),
            fired: AtomicBool::new(false),
        });
        let source: Arc<dyn ReadAt> = hostile.clone();
        let snapshot = SourceSnapshot::open(source).expect("source-backed DOC fixture");
        let transaction = snapshot
            .edit_paragraph(Position::new(0))
            .expect("first paragraph");
        hostile.armed.store(true, Ordering::SeqCst);
        assert!(matches!(
            transaction.commit(),
            Err(Error::Overlay(
                OverlayError::SourceFingerprintChanged { .. }
            ))
        ));
        assert!(hostile.fired.load(Ordering::SeqCst));
    }

    #[test]
    fn first_paragraph_scans_a_giant_piece_in_bounded_chunks() {
        let giant = "x".repeat(SCAN_CHUNK_UNITS * 8);
        let mut writer = crate::writer::Writer::new();
        let combined = format!("a\r{giant}");
        writer.add_paragraph(&combined).expect("giant paragraph");
        let mut output = io::Cursor::new(Vec::new());
        writer.write_to(&mut output).expect("DOC fixture");
        let counted = Arc::new(CountingSource {
            bytes: Arc::from(output.into_inner()),
            total_bytes: AtomicU64::new(0),
            max_read: AtomicUsize::new(0),
        });
        let source: Arc<dyn ReadAt> = counted.clone();
        let snapshot = SourceSnapshot::open(source).expect("source-backed DOC fixture");
        let scan_chunk_cp = u32::try_from(SCAN_CHUNK_UNITS).expect("scan chunk size");
        assert!(
            snapshot
                .inner
                .layout
                .pieces
                .first()
                .is_some_and(|piece| piece.cp_end > scan_chunk_cp)
        );
        counted.total_bytes.store(0, Ordering::SeqCst);
        counted.max_read.store(0, Ordering::SeqCst);
        let paragraph = resolve_paragraph(
            &snapshot.inner.shared,
            &snapshot.inner.layout,
            Position::new(0),
            snapshot.inner.limits,
        )
        .expect("first paragraph");
        assert_eq!(paragraph.text, "a");
        assert!(counted.total_bytes.load(Ordering::SeqCst) < giant.len() as u64);
    }

    struct CountingSource {
        bytes: Arc<[u8]>,
        total_bytes: AtomicU64,
        max_read: AtomicUsize,
    }

    impl ReadAt for CountingSource {
        fn len(&self) -> io::Result<u64> {
            u64::try_from(self.bytes.len())
                .map_err(|_error| io::Error::other("source length does not fit u64"))
        }

        fn read_at(&self, offset: u64, output: &mut [u8]) -> io::Result<usize> {
            let start = usize::try_from(offset)
                .map_err(|_error| io::Error::other("source offset does not fit usize"))?;
            let Some(input) = self.bytes.get(start..) else {
                return Ok(0);
            };
            let count = input.len().min(output.len());
            output[..count].copy_from_slice(&input[..count]);
            self.total_bytes.fetch_add(
                u64::try_from(count)
                    .map_err(|_error| io::Error::other("source read count does not fit u64"))?,
                Ordering::SeqCst,
            );
            self.max_read.fetch_max(count, Ordering::SeqCst);
            Ok(count)
        }

        fn version(&self) -> io::Result<SourceVersion> {
            Ok(SourceVersion::new(0xD0C0_0113, 0))
        }
    }

    struct StableMutationSource {
        bytes: Mutex<Vec<u8>>,
        armed: AtomicBool,
        fired: AtomicBool,
    }

    impl ReadAt for StableMutationSource {
        fn len(&self) -> io::Result<u64> {
            Ok(self.bytes.lock().expect("source lock").len() as u64)
        }

        fn read_at(&self, offset: u64, output: &mut [u8]) -> io::Result<usize> {
            let mut bytes = self.bytes.lock().expect("source lock");
            let start = usize::try_from(offset)
                .map_err(|_error| io::Error::other("source offset does not fit usize"))?;
            let Some(input) = bytes.get(start..) else {
                return Ok(0);
            };
            let count = input.len().min(output.len());
            output[..count].copy_from_slice(&input[..count]);
            if self.armed.load(Ordering::SeqCst) && !self.fired.swap(true, Ordering::SeqCst) {
                if let Some(last) = bytes.last_mut() {
                    *last ^= 0x01;
                }
            }
            Ok(count)
        }

        fn version(&self) -> io::Result<SourceVersion> {
            Ok(SourceVersion::new(0xD0C0_0112, 0))
        }
    }

    struct FailingSink {
        remaining: usize,
    }

    impl Write for FailingSink {
        fn write(&mut self, bytes: &[u8]) -> io::Result<usize> {
            if self.remaining == 0 {
                return Err(io::Error::other("test sink stopped"));
            }
            let accepted = self.remaining.min(bytes.len());
            self.remaining -= accepted;
            Ok(accepted)
        }

        fn flush(&mut self) -> io::Result<()> {
            Ok(())
        }
    }
}
