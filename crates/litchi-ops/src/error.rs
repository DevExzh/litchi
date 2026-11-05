//! Unified typed error with the closed code list, messages, and suggestions.
//!
//! Every failure of the unified operations layer is one [`Error`]. An error
//! carries exactly one [`Code`] from the closed list agreed in
//! `docs/UNIFIED_OPS_API_DESIGN.md` (OfficeCLI's full taxonomy plus the Litchi
//! extensions `unsupported_capability`, `limit_exceeded` and `cancelled`), a
//! human-readable message, an optional suggestion, the format involved where
//! one is known, and the typed leaf error it came from.
//!
//! Leaf errors are never flattened into strings: the original error is kept as
//! a typed [`Source`] and is returned by [`std::error::Error::source`], so
//! callers can still match on the leaf's limit, offset or transaction detail.
//!
//! # Classification
//!
//! Each `From` conversion classifies its leaf error once, on the error path:
//!
//! | Leaf condition | Code |
//! |---|---|
//! | I/O error of kind `NotFound` | [`Code::FileNotFound`] |
//! | Any other I/O error, or a source that changed after open | [`Code::IoError`] |
//! | No Office format recognized | [`Code::UnsupportedFormat`], no detected format |
//! | A recognized format this opener does not handle | [`Code::UnsupportedFormat`], with the detected format |
//! | A recognized format whose cargo feature is off | [`Code::UnsupportedFormat`], with a suggestion |
//! | Malformed package, record, XML or payload | [`Code::CorruptFile`] |
//! | A finite resource budget was exceeded | [`Code::LimitExceeded`] |
//! | The format cannot perform the requested operation | [`Code::UnsupportedCapability`] |
//! | Password, signature or edit protection | [`Code::DocumentProtected`], with a suggestion |
//! | Allocation failure, durability after commit, internal invariant | [`Code::InternalError`] |
//! | Cooperative cancellation | [`Code::Cancelled`] |
//! | An invalid configured limit or execution policy, a wrong password, or a caller's cell reference or range that names no grid cells | [`Code::InvalidArgument`] |
//! | An index outside its collection, among them a caller's cell indices outside the worksheet grid | [`Code::OutOfRange`] |
//! | A named slide, sheet, tag or font that does not exist | [`Code::NotFound`] |
//! | A sheet that is not a worksheet | [`Code::UnsupportedType`] |
//! | Path or selector syntax | [`Code::InvalidPath`] or [`Code::InvalidSelector`]; [`Code::MaxDepthExceeded`] for the step and nesting limits |
//!
//! Leaf enums that are `#[non_exhaustive]` end in a wildcard arm that reports
//! [`Code::CorruptFile`]: a variant the classifier does not name is treated as
//! a malformed input until it is classified explicitly.
//!
//! # Refused additions
//!
//! Every format answers an `add` it cannot perform the same way (record
//! 0269):
//!
//! | Refusal | Code |
//! |---|---|
//! | The element kind cannot be added in this format (yet), whatever the parent | [`Code::UnsupportedElement`] |
//! | The element kind is addable, but not under this parent | [`Code::UnsupportedType`] |
//! | The kind is addable here, but a declared gap (a placement, a copy source, a property family) is pending | [`Code::UnsupportedCapability`], with the follow-up as suggestion |
//!
//! The kind is checked before the parent, so the message names what is
//! missing rather than a parent rule of another kind.

use std::borrow::Cow;
use std::fmt;
use std::io;
use std::sync::Arc;

use litchi_core::{ExecutionError, FileFormat};

/// Result type of the unified operations layer.
pub type Result<T> = std::result::Result<T, Error>;

/// Stable, machine-readable error code.
///
/// The list is closed: every [`Error`] carries exactly one code. The wire
/// spelling returned by [`Code::as_str`] is the `snake_case` OfficeCLI code,
/// so agents can switch on it across Rust, Python and the CLI.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum Code {
    /// A requested element does not exist.
    NotFound,
    /// A path resolved to nothing.
    PathNotFound,
    /// A query matched no element where at least one was required.
    ZeroMatches,
    /// An index or position is outside its collection.
    OutOfRange,
    /// The target of a range operation does not exist.
    RangeTargetNotFound,
    /// A name is already in use.
    DuplicateName,
    /// A property value is invalid for its property.
    InvalidValue,
    /// An input document or payload is invalid.
    InvalidInput,
    /// An argument is invalid.
    InvalidArgument,
    /// A JSON payload is malformed.
    InvalidJson,
    /// An `XPath` expression is malformed.
    InvalidXpath,
    /// A path is malformed.
    InvalidPath,
    /// A selector is malformed.
    InvalidSelector,
    /// The property is not supported on this element.
    UnsupportedProperty,
    /// The path form is not supported here.
    UnsupportedPath,
    /// The element kind is not supported here.
    UnsupportedElement,
    /// The value or element type is not supported here.
    UnsupportedType,
    /// The file format is not recognized, not handled by this view, or not
    /// compiled in.
    UnsupportedFormat,
    /// The command is not supported.
    UnsupportedCommand,
    /// A required property is missing.
    MissingProperty,
    /// A required argument is missing.
    MissingArgument,
    /// The file does not exist.
    FileNotFound,
    /// The file is locked by another process.
    FileLocked,
    /// The destination file already exists.
    FileExists,
    /// The operation requires a file.
    FileRequired,
    /// Reading or writing failed, or the source changed underneath an open
    /// handle.
    IoError,
    /// The file is malformed.
    CorruptFile,
    /// A batch payload is malformed.
    CorruptBatch,
    /// The document is encrypted, signed, or edit-protected.
    DocumentProtected,
    /// The result failed validation.
    ValidationError,
    /// The output would exceed its size ceiling.
    OutputTooLarge,
    /// Nesting exceeded its depth ceiling.
    MaxDepthExceeded,
    /// A merge could not be applied exactly.
    MergeNotExact,
    /// A refresh of derived content failed.
    RefreshFailed,
    /// A bare selector was rejected where a qualified one is required.
    BareSelectorRejected,
    /// An internal failure, such as an allocation failure or a broken
    /// invariant.
    InternalError,
    /// Litchi extension: the format cannot perform the requested operation.
    UnsupportedCapability,
    /// Litchi extension: a finite resource budget was exceeded. Retrying with
    /// the same limits fails again.
    LimitExceeded,
    /// Litchi extension: the operation was cooperatively cancelled.
    Cancelled,
}

impl Code {
    /// Every code, in declaration order.
    pub const ALL: &'static [Self] = &[
        Self::NotFound,
        Self::PathNotFound,
        Self::ZeroMatches,
        Self::OutOfRange,
        Self::RangeTargetNotFound,
        Self::DuplicateName,
        Self::InvalidValue,
        Self::InvalidInput,
        Self::InvalidArgument,
        Self::InvalidJson,
        Self::InvalidXpath,
        Self::InvalidPath,
        Self::InvalidSelector,
        Self::UnsupportedProperty,
        Self::UnsupportedPath,
        Self::UnsupportedElement,
        Self::UnsupportedType,
        Self::UnsupportedFormat,
        Self::UnsupportedCommand,
        Self::MissingProperty,
        Self::MissingArgument,
        Self::FileNotFound,
        Self::FileLocked,
        Self::FileExists,
        Self::FileRequired,
        Self::IoError,
        Self::CorruptFile,
        Self::CorruptBatch,
        Self::DocumentProtected,
        Self::ValidationError,
        Self::OutputTooLarge,
        Self::MaxDepthExceeded,
        Self::MergeNotExact,
        Self::RefreshFailed,
        Self::BareSelectorRejected,
        Self::InternalError,
        Self::UnsupportedCapability,
        Self::LimitExceeded,
        Self::Cancelled,
    ];

    /// The stable `snake_case` wire spelling of the code.
    #[must_use]
    pub const fn as_str(self) -> &'static str {
        match self {
            Self::NotFound => "not_found",
            Self::PathNotFound => "path_not_found",
            Self::ZeroMatches => "zero_matches",
            Self::OutOfRange => "out_of_range",
            Self::RangeTargetNotFound => "range_target_not_found",
            Self::DuplicateName => "duplicate_name",
            Self::InvalidValue => "invalid_value",
            Self::InvalidInput => "invalid_input",
            Self::InvalidArgument => "invalid_argument",
            Self::InvalidJson => "invalid_json",
            Self::InvalidXpath => "invalid_xpath",
            Self::InvalidPath => "invalid_path",
            Self::InvalidSelector => "invalid_selector",
            Self::UnsupportedProperty => "unsupported_property",
            Self::UnsupportedPath => "unsupported_path",
            Self::UnsupportedElement => "unsupported_element",
            Self::UnsupportedType => "unsupported_type",
            Self::UnsupportedFormat => "unsupported_format",
            Self::UnsupportedCommand => "unsupported_command",
            Self::MissingProperty => "missing_property",
            Self::MissingArgument => "missing_argument",
            Self::FileNotFound => "file_not_found",
            Self::FileLocked => "file_locked",
            Self::FileExists => "file_exists",
            Self::FileRequired => "file_required",
            Self::IoError => "io_error",
            Self::CorruptFile => "corrupt_file",
            Self::CorruptBatch => "corrupt_batch",
            Self::DocumentProtected => "document_protected",
            Self::ValidationError => "validation_error",
            Self::OutputTooLarge => "output_too_large",
            Self::MaxDepthExceeded => "max_depth_exceeded",
            Self::MergeNotExact => "merge_not_exact",
            Self::RefreshFailed => "refresh_failed",
            Self::BareSelectorRejected => "bare_selector_rejected",
            Self::InternalError => "internal_error",
            Self::UnsupportedCapability => "unsupported_capability",
            Self::LimitExceeded => "limit_exceeded",
            Self::Cancelled => "cancelled",
        }
    }
}

impl Code {
    /// A short, format-neutral description of the code, used where the
    /// error's own message comes from a leaf error.
    const fn summary(self) -> &'static str {
        match self {
            Self::NotFound => "a requested element does not exist",
            Self::PathNotFound => "a path resolved to nothing",
            Self::ZeroMatches => "a query matched no element",
            Self::OutOfRange => "an index is outside its collection",
            Self::RangeTargetNotFound => "the target of a range operation does not exist",
            Self::DuplicateName => "a name is already in use",
            Self::InvalidValue => "a property value is invalid",
            Self::InvalidInput => "an input document or payload is invalid",
            Self::InvalidArgument => "an argument is invalid",
            Self::InvalidJson => "a JSON payload is malformed",
            Self::InvalidXpath => "an XPath expression is malformed",
            Self::InvalidPath => "a path is malformed",
            Self::InvalidSelector => "a selector is malformed",
            Self::UnsupportedProperty => "the property is not supported on this element",
            Self::UnsupportedPath => "the path form is not supported here",
            Self::UnsupportedElement => "the element kind is not supported here",
            Self::UnsupportedType => "the value or element type is not supported here",
            Self::UnsupportedFormat => "the file format is not supported here",
            Self::UnsupportedCommand => "the command is not supported",
            Self::MissingProperty => "a required property is missing",
            Self::MissingArgument => "a required argument is missing",
            Self::FileNotFound => "the file does not exist",
            Self::FileLocked => "the file is locked by another process",
            Self::FileExists => "the destination file already exists",
            Self::FileRequired => "the operation requires a file",
            Self::IoError => "reading or writing failed, or the source changed",
            Self::CorruptFile => "the file is malformed",
            Self::CorruptBatch => "a batch payload is malformed",
            Self::DocumentProtected => "the document is encrypted, signed, or edit-protected",
            Self::ValidationError => "the result failed validation",
            Self::OutputTooLarge => "the output would exceed its size ceiling",
            Self::MaxDepthExceeded => "nesting exceeded its depth ceiling",
            Self::MergeNotExact => "a merge could not be applied exactly",
            Self::RefreshFailed => "a refresh of derived content failed",
            Self::BareSelectorRejected => "a bare selector was rejected",
            Self::InternalError => "an internal failure occurred",
            Self::UnsupportedCapability => "the format cannot perform the requested operation",
            Self::LimitExceeded => "a resource limit was exceeded",
            Self::Cancelled => "the operation was cancelled",
        }
    }
}

impl fmt::Display for Code {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter.write_str(self.as_str())
    }
}

/// The typed leaf error an [`Error`] was converted from.
///
/// Each format-specific variant exists only when its cargo feature is enabled.
/// ODT, ODS and ODP report [`litchi_core::Error`] and arrive as
/// [`Source::Core`].
#[non_exhaustive]
pub enum Source {
    /// A format-neutral atomic publication failure.
    Publication(litchi_core::atomic::Error),
    /// The error was raised by the unified layer itself.
    None,
    /// A bounded batch JSON ingress error.
    Batch(crate::batch::Error),
    /// A CSV/TSV import decoding or parsing error (record 0266).
    Import(crate::import::Error),
    /// A template merge data error (record 0274).
    Merge(crate::merge::Error),
    /// A raw-layer validation that could not complete (record 0336).
    Raw(crate::raw::Error),
    /// A formula evaluation refusal or bound (record 0266).
    #[cfg(all(
        feature = "eval",
        any(feature = "xlsx", feature = "xlsb", feature = "xls")
    ))]
    Calc(crate::calc::Error),
    /// An explicit transaction admission or replay error.
    Transaction(crate::transaction::Error),
    /// An I/O error.
    Io(io::Error),
    /// A shared Litchi core error.
    Core(litchi_core::Error),
    /// An execution-context refusal: cancellation or a budget charge.
    Execution(ExecutionError),
    /// A path or selector syntax error, with its typed kind and byte offset.
    Selector(litchi_core::selector::ParseError),
    /// An Open Packaging Conventions error.
    #[cfg(any(feature = "docx", feature = "pptx", feature = "xlsx", feature = "xlsb"))]
    Opc(litchi_opc::OpcError),
    /// A Compound File Binary error.
    #[cfg(any(
        feature = "doc",
        feature = "xls",
        feature = "ppt",
        feature = "encryption"
    ))]
    Cfb(litchi_cfb::OleError),
    /// A `WordprocessingML` error.
    #[cfg(feature = "docx")]
    Docx(litchi_docx::Error),
    /// A `WordprocessingML` main-document transaction error.
    #[cfg(feature = "docx")]
    DocxTransaction(litchi_docx::document::Error),
    /// A `WordprocessingML` core-properties transaction error.
    #[cfg(feature = "docx")]
    DocxMetadata(litchi_docx::metadata::Error),
    /// A `PresentationML` error.
    #[cfg(feature = "pptx")]
    Pptx(litchi_pptx::Error),
    /// A `SpreadsheetML` error.
    #[cfg(feature = "xlsx")]
    Xlsx(litchi_xlsx::Error),
    /// A BIFF12 package error.
    #[cfg(feature = "xlsb")]
    Xlsb(litchi_xlsb::package::PackageError),
    /// A BIFF12 record-wire error.
    #[cfg(feature = "xlsb")]
    XlsbRaw(litchi_xlsb::Error),
    /// A legacy Word binary error.
    #[cfg(feature = "doc")]
    Doc(litchi_doc::Error),
    /// A legacy Word main-story semantic edit error.
    #[cfg(feature = "doc")]
    DocEdit(litchi_doc::body_text::Error),
    /// A canonical legacy Word package transaction error.
    #[cfg(feature = "doc")]
    DocTransaction(litchi_doc::package::edit::Error),
    /// A legacy Word package creation/publication error.
    #[cfg(feature = "doc")]
    DocWrite(litchi_doc::writer::WriteError),
    /// A source-backed legacy Word main-story edit error.
    #[cfg(feature = "doc")]
    DocBodyText(litchi_doc::source_backed::Error),
    /// A legacy Word document-property (metadata) transaction error.
    #[cfg(feature = "doc")]
    DocMetadata(litchi_doc::package::metadata::Error),
    /// A legacy Excel document-property (metadata) transaction error.
    #[cfg(feature = "xls")]
    XlsMetadata(litchi_xls::package::metadata::Error),
    /// A legacy PowerPoint document-property (metadata) transaction error.
    #[cfg(feature = "ppt")]
    PptMetadata(litchi_ppt::package::metadata::Error),
    /// A legacy Excel BIFF error.
    #[cfg(feature = "xls")]
    Xls(litchi_xls::Error),
    /// A canonical legacy Excel package or projection error.
    #[cfg(feature = "xls")]
    XlsPackage(litchi_xls::package::Error),
    /// A canonical primary XLS edit or patch error.
    #[cfg(feature = "xls")]
    XlsEdit(litchi_xls::package::edit::Error),
    /// A source-backed legacy Excel error.
    #[cfg(feature = "xls")]
    XlsSource(litchi_xls::SourceBackedError),
    /// A legacy PowerPoint binary error.
    #[cfg(feature = "ppt")]
    Ppt(litchi_ppt::Error),
    /// A canonical primary PPT edit or patch error.
    #[cfg(feature = "ppt")]
    PptEdit(litchi_ppt::package::edit::Error),
    /// A source-backed legacy PowerPoint text-edit error.
    #[cfg(feature = "ppt")]
    PptTextEdit(litchi_ppt::text_edit::Error),
    /// A compound-file signature error, raised while a legacy binary
    /// document's signature is removed after `Signed::unsign`.
    #[cfg(any(feature = "doc", feature = "ppt", feature = "xls"))]
    Sign(litchi_sign::Error),
    /// A Rich Text Format error.
    #[cfg(feature = "rtf")]
    Rtf(litchi_rtf::Error),
    /// A native RTF authoring error.
    #[cfg(feature = "rtf")]
    RtfBuild(litchi_rtf::builder::Error),
    /// A native RTF transaction or patch error.
    #[cfg(feature = "rtf")]
    RtfEdit(litchi_rtf::edit::Error),
    /// An OOXML package-encryption error.
    #[cfg(feature = "encryption")]
    Crypto(litchi_crypto::ooxml::Error),
    /// An ODF package ingress, egress or transaction error (typed limits).
    #[cfg(any(feature = "odt", feature = "ods", feature = "odp", odf_objects))]
    Odf(litchi_odf_common::Error),
    /// An ODP source-backed tail-append or replay error. This route retains
    /// the format-owned freshness progress and nested source diagnostics.
    #[cfg(feature = "odp")]
    OdpTailAppend(litchi_odp::tail_append::Error),
    /// A formula conversion error: Markdown output renders OMML formulas as
    /// LaTeX.
    #[cfg(feature = "formula")]
    Formula(litchi_formula::FormulaError),
    /// A Pages, Keynote or Numbers leaf error (record 0296).
    #[cfg(feature = "iwork-adapter")]
    Iwork(crate::format::iwork::LeafError),
}

// Leaf diagnostic payloads can contain authored document text. Typed access
// remains explicit through `leaf()` and the standard error source chain.
impl fmt::Debug for Source {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter.debug_struct("Source").finish_non_exhaustive()
    }
}

impl Source {
    /// The leaf error as a standard error, or `None` for [`Source::None`].
    fn as_error(&self) -> Option<&(dyn std::error::Error + 'static)> {
        match self {
            Self::Publication(error) => Some(error),
            Self::None => None,
            Self::Batch(error) => Some(error),
            Self::Import(error) => Some(error),
            Self::Merge(error) => Some(error),
            Self::Raw(error) => Some(error),
            #[cfg(all(
                feature = "eval",
                any(feature = "xlsx", feature = "xlsb", feature = "xls")
            ))]
            Self::Calc(error) => Some(error),
            Self::Transaction(error) => Some(error),
            Self::Io(error) => Some(error),
            Self::Core(error) => Some(error),
            Self::Execution(error) => Some(error),
            Self::Selector(error) => Some(error),
            #[cfg(any(feature = "docx", feature = "pptx", feature = "xlsx", feature = "xlsb"))]
            Self::Opc(error) => Some(error),
            #[cfg(any(
                feature = "doc",
                feature = "xls",
                feature = "ppt",
                feature = "encryption"
            ))]
            Self::Cfb(error) => Some(error),
            #[cfg(feature = "docx")]
            Self::Docx(error) => Some(error),
            #[cfg(feature = "docx")]
            Self::DocxTransaction(error) => Some(error),
            #[cfg(feature = "docx")]
            Self::DocxMetadata(error) => Some(error),
            #[cfg(feature = "pptx")]
            Self::Pptx(error) => Some(error),
            #[cfg(feature = "xlsx")]
            Self::Xlsx(error) => Some(error),
            #[cfg(feature = "xlsb")]
            Self::Xlsb(error) => Some(error),
            #[cfg(feature = "xlsb")]
            Self::XlsbRaw(error) => Some(error),
            #[cfg(feature = "doc")]
            Self::Doc(error) => Some(error),
            #[cfg(feature = "doc")]
            Self::DocEdit(error) => Some(error),
            #[cfg(feature = "doc")]
            Self::DocTransaction(error) => Some(error),
            #[cfg(feature = "doc")]
            Self::DocWrite(error) => Some(error),
            #[cfg(feature = "doc")]
            Self::DocBodyText(error) => Some(error),
            #[cfg(feature = "doc")]
            Self::DocMetadata(error) => Some(error),
            #[cfg(feature = "xls")]
            Self::XlsMetadata(error) => Some(error),
            #[cfg(feature = "ppt")]
            Self::PptMetadata(error) => Some(error),
            #[cfg(feature = "xls")]
            Self::Xls(error) => Some(error),
            #[cfg(feature = "xls")]
            Self::XlsPackage(error) => Some(error),
            #[cfg(feature = "xls")]
            Self::XlsEdit(error) => Some(error),
            #[cfg(feature = "xls")]
            Self::XlsSource(error) => Some(error),
            #[cfg(feature = "ppt")]
            Self::Ppt(error) => Some(error),
            #[cfg(feature = "ppt")]
            Self::PptEdit(error) => Some(error),
            #[cfg(feature = "ppt")]
            Self::PptTextEdit(error) => Some(error),
            #[cfg(any(feature = "doc", feature = "ppt", feature = "xls"))]
            Self::Sign(error) => Some(error),
            #[cfg(feature = "rtf")]
            Self::Rtf(error) => Some(error),
            #[cfg(feature = "rtf")]
            Self::RtfBuild(error) => Some(error),
            #[cfg(feature = "rtf")]
            Self::RtfEdit(error) => Some(error),
            #[cfg(feature = "encryption")]
            Self::Crypto(error) => Some(error),
            #[cfg(any(feature = "odt", feature = "ods", feature = "odp", odf_objects))]
            Self::Odf(error) => Some(error),
            #[cfg(feature = "odp")]
            Self::OdpTailAppend(error) => Some(error),
            #[cfg(feature = "formula")]
            Self::Formula(error) => Some(error),
            #[cfg(feature = "iwork-adapter")]
            Self::Iwork(error) => Some(error),
        }
    }
}

/// The error of every unified operation.
///
/// The error is one pointer wide, so `Result<T>` stays small on the success
/// path, and cloning it shares the typed leaf error instead of copying it.
///
/// Its [`Display`](fmt::Display) form is facade-level: the code and its
/// summary for an error converted from a leaf error, or the code and the
/// message for an error the unified layer raised itself, followed by the
/// format, the failing property and batch item, the suggestion, and the
/// properties a warn-policy command skipped before it failed
/// ([`Error::warnings`]), when there are any. The leaf error's own text is
/// [`Error::message`], and the leaf error is [`std::error::Error::source`], so
/// an error-chain printer shows it once, as the cause.
#[derive(Clone)]
pub struct Error(Box<Inner>);

#[derive(Clone)]
struct Inner {
    code: Code,
    message: Cow<'static, str>,
    suggestion: Option<Cow<'static, str>>,
    format: Option<FileFormat>,
    detected: Option<FileFormat>,
    /// The zero-based index of the batch item that failed.
    item: Option<usize>,
    /// The property key that failed, as the caller wrote it.
    property: Option<Cow<'static, str>>,
    /// Shared so that cloning an error never copies or loses its leaf error.
    source: Arc<Source>,
    /// The properties a `batch::Unsupported::Warn` command skipped before
    /// it failed; empty (and unallocated) otherwise.
    warnings: Box<[Error]>,
}

const _: () = assert!(size_of::<Error>() == size_of::<usize>());

impl Error {
    /// Creates an error raised by the unified layer itself, with no leaf
    /// source.
    #[must_use]
    pub fn new(code: Code, message: impl Into<Cow<'static, str>>) -> Self {
        Self(Box::new(Inner {
            code,
            message: message.into(),
            suggestion: None,
            format: None,
            detected: None,
            item: None,
            property: None,
            source: Arc::new(Source::None),
            warnings: Box::default(),
        }))
    }

    /// Reports an input whose format is unsupported. `detected` names the
    /// format detection recognized, or is `None` when no Office format was
    /// recognized.
    #[must_use]
    pub fn unsupported_format(detected: Option<FileFormat>) -> Self {
        let message: Cow<'static, str> = match detected {
            Some(format) => {
                format!("detected {format:?}, which this operation does not handle").into()
            },
            None => "not a recognized Office file".into(),
        };
        let mut error = Self::new(Code::UnsupportedFormat, message);
        error.0.detected = detected;
        error
    }

    /// Reports an iWork input (Pages, Keynote or Numbers) that a caller's ZIP
    /// tail named but this build does not open: the iWork arm of detection
    /// declined it, or its cargo feature is off. The suggestion names the
    /// features.
    #[must_use]
    pub fn iwork_unsupported(detected: FileFormat) -> Self {
        Self::unsupported_format(Some(detected)).with_suggestion(IWORK_SUGGESTION)
    }

    /// Reports a read or edit attempt on an encrypted document that has not
    /// been unlocked. `format` is the format of the encrypted content, or
    /// `None` when it cannot be known before the document is unlocked.
    #[must_use]
    pub fn locked(format: Option<FileFormat>) -> Self {
        let message: Cow<'static, str> = match format {
            Some(format) => format!("the {format:?} document is encrypted").into(),
            None => "the document is encrypted".into(),
        };
        Self::from_class(
            Class::suggest(Code::DocumentProtected, LOCKED_SUGGESTION),
            message,
            format,
            Source::None,
        )
    }

    /// Reports an edit attempt on a signed document, which would invalidate
    /// its signature.
    #[must_use]
    pub fn signed(format: FileFormat) -> Self {
        Self::from_class(
            Class::suggest(Code::DocumentProtected, SIGNED_SUGGESTION),
            format!("editing the signed {format:?} document would invalidate its signature"),
            Some(format),
            Source::None,
        )
    }

    /// Reports a recognized format whose cargo feature is not enabled in this
    /// build.
    #[cfg(any(
        feature = "doc",
        feature = "docx",
        feature = "ppt",
        feature = "pptx",
        feature = "xls",
        feature = "xlsx",
        feature = "xlsb",
        feature = "rtf",
        feature = "odt",
        feature = "ods",
        feature = "odp",
        feature = "pages",
        feature = "keynote",
        feature = "numbers",
        odf_objects
    ))]
    pub(crate) fn feature_disabled(detected: FileFormat, feature: &'static str) -> Self {
        Self::unsupported_format(Some(detected))
            .with_suggestion(format!("enable the `{feature}` cargo feature"))
    }

    /// Sets the suggestion that tells the caller how to recover.
    #[must_use]
    pub fn with_suggestion(mut self, suggestion: impl Into<Cow<'static, str>>) -> Self {
        self.0.suggestion = Some(suggestion.into());
        self
    }

    /// Sets the format of the document the failing operation worked on.
    #[must_use]
    pub fn with_format(mut self, format: FileFormat) -> Self {
        self.0.format = Some(format);
        self
    }

    /// Reclassifies a leaf error that the caller's own argument caused,
    /// keeping its typed leaf and message: [`Code::InvalidArgument`] for
    /// input that does not parse, [`Code::InvalidValue`] for an authored property,
    /// [`Code::DuplicateName`] for a conflicting authored name, and
    /// [`Code::OutOfRange`] for a position outside
    /// its collection, and [`Code::UnsupportedCapability`] for a write the
    /// leaf refuses on a document it already parsed (ODS, record 0277). A
    /// leaf crate reports such input with the same error it
    /// reports for malformed file content, so only the call site that
    /// received the argument can tell the two apart.
    #[cfg(any(
        feature = "xlsx",
        feature = "xlsb",
        feature = "ods",
        feature = "odt",
        feature = "odp",
        odf_objects
    ))]
    pub(crate) fn caused_by_argument(mut self, code: Code) -> Self {
        debug_assert!(matches!(
            code,
            Code::InvalidArgument
                | Code::InvalidValue
                | Code::OutOfRange
                | Code::DuplicateName
                | Code::UnsupportedCapability
        ));
        self.0.code = code;
        self.0.suggestion = None;
        self
    }

    /// The stable error code.
    #[must_use]
    pub fn code(&self) -> Code {
        self.0.code
    }

    /// The human-readable message: the leaf error's own text for an error
    /// converted from one, otherwise the unified layer's message.
    #[must_use]
    pub fn message(&self) -> &str {
        &self.0.message
    }

    /// How the caller can recover, when there is a known way.
    #[must_use]
    pub fn suggestion(&self) -> Option<&str> {
        self.0.suggestion.as_deref()
    }

    /// The format of the document the failing operation worked on, when known.
    #[must_use]
    pub fn format(&self) -> Option<FileFormat> {
        self.0.format
    }

    /// For [`Code::UnsupportedFormat`], the format detection recognized;
    /// `None` when no Office format was recognized.
    #[must_use]
    pub fn detected(&self) -> Option<FileFormat> {
        self.0.detected
    }

    /// Sets the zero-based index of the batch item that failed.
    #[must_use]
    pub fn with_item(mut self, item: usize) -> Self {
        self.0.item = Some(item);
        self
    }

    /// Sets the property key that failed, as the caller wrote it.
    #[must_use]
    pub fn with_property(mut self, property: impl Into<Cow<'static, str>>) -> Self {
        self.0.property = Some(property.into());
        self
    }

    /// The zero-based index of the batch item that failed, when this error was
    /// attributed to one.
    #[must_use]
    pub fn item(&self) -> Option<usize> {
        self.0.item
    }

    /// The property key that failed, as the caller wrote it, when one was
    /// responsible for the error.
    #[must_use]
    pub fn property(&self) -> Option<&str> {
        self.0.property.as_deref()
    }

    /// The properties a command run under
    /// [`Unsupported::Warn`](crate::batch::Unsupported::Warn) skipped before
    /// it failed, in the order they were refused: each the
    /// `unsupported_property` refusal that dropped it, exactly as a
    /// successful command reports them in
    /// [`Applied::warnings`](crate::batch::Applied::warnings).
    ///
    /// When this is not empty, this error is the failure of the command
    /// *without* those properties, and the command as written was refused
    /// first for the first of them. A failure that only the skipping caused
    /// (a chart `add` whose series all came from unsupported keys fails
    /// with `missing_property`) is told apart from a failure of the
    /// caller's own input by this list. Empty for every other error.
    ///
    /// ```
    /// # #[cfg(feature = "xlsx")]
    /// # {
    /// use litchi_core::FileFormat;
    /// use litchi_ops::batch::{Options, Unsupported};
    /// use litchi_ops::{BatchItem, Code, Props, create};
    ///
    /// let file = create(FileFormat::Xlsx)?;
    /// let mut item = BatchItem::new("set");
    /// item.path = Some("/Sheet1/A1".into());
    /// item.props = Some(Props::new().with("sparkle", true).with("font.size", -3));
    /// let options = Options::new().unsupported(Unsupported::Warn);
    /// let error = file.apply_with(&item, &options).unwrap_err();
    /// // The rest of the command failed on its own value ...
    /// assert_eq!(error.code(), Code::InvalidValue);
    /// // ... after `sparkle` was skipped.
    /// assert_eq!(error.warnings()[0].property(), Some("sparkle"));
    /// # }
    /// # Ok::<(), litchi_ops::Error>(())
    /// ```
    #[must_use]
    pub fn warnings(&self) -> &[Error] {
        &self.0.warnings
    }

    /// Attaches the properties a warn-policy command skipped before it
    /// failed.
    pub(crate) fn with_warnings(mut self, warnings: Box<[Error]>) -> Self {
        self.0.warnings = warnings;
        self
    }

    /// The typed leaf error this error was converted from. Clones of an error
    /// share it.
    #[must_use]
    pub fn leaf(&self) -> &Source {
        &self.0.source
    }

    fn from_class(
        class: Class,
        message: impl Into<Cow<'static, str>>,
        format: Option<FileFormat>,
        source: Source,
    ) -> Self {
        Self(Box::new(Inner {
            code: class.code,
            message: message.into(),
            suggestion: class.suggestion,
            format,
            detected: class.detected,
            item: None,
            property: None,
            source: Arc::new(source),
            warnings: Box::default(),
        }))
    }
}

impl fmt::Debug for Error {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("Error")
            .field("code", &self.0.code)
            .field("format", &self.0.format)
            .field("detected", &self.0.detected)
            .field("source", &self.0.source)
            .finish()
    }
}

/// The facade-level rendering: a leaf error's text is not repeated here,
/// because [`std::error::Error::source`] returns the leaf error and
/// error-chain printers show it as the cause.
impl fmt::Display for Error {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter.write_str(self.0.code.as_str())?;
        formatter.write_str(": ")?;
        if matches!(*self.0.source, Source::None) {
            formatter.write_str(&self.0.message)?;
        } else {
            formatter.write_str(self.0.code.summary())?;
            if let Some(format) = self.0.format.or(self.0.detected) {
                write!(formatter, " [{format:?}]")?;
            }
        }
        if let Some(property) = &self.0.property {
            write!(formatter, " (property: {property})")?;
        }
        if let Some(item) = self.0.item {
            write!(formatter, " (item: {item})")?;
        }
        if let Some(suggestion) = &self.0.suggestion {
            write!(formatter, " (suggestion: {suggestion})")?;
        }
        if !self.0.warnings.is_empty() {
            formatter.write_str(" (after skipping unsupported properties: ")?;
            for (index, warning) in self.0.warnings.iter().enumerate() {
                if index > 0 {
                    formatter.write_str(", ")?;
                }
                formatter.write_str(warning.property().unwrap_or("?"))?;
            }
            formatter.write_str(")")?;
        }
        Ok(())
    }
}

impl std::error::Error for Error {
    fn source(&self) -> Option<&(dyn std::error::Error + 'static)> {
        self.0.source.as_error()
    }
}

const LIMIT_SUGGESTION: &str =
    "raise the matching limit in the open options if the input is trusted";
#[cfg(any(feature = "doc", feature = "xls", feature = "ppt", feature = "xlsb"))]
const PASSWORD_SUGGESTION: &str = "unlock the document with its password";
#[cfg(feature = "encryption")]
const LOCKED_SUGGESTION: &str = "unlock it with `Locked::unlock` and its password";
#[cfg(not(feature = "encryption"))]
const LOCKED_SUGGESTION: &str =
    "build `litchi-ops` with the `encryption` feature, then unlock it with `Locked::unlock`";
#[cfg(any(
    feature = "doc",
    feature = "xls",
    feature = "ppt",
    feature = "encryption"
))]
const WRONG_PASSWORD_SUGGESTION: &str = "check the password and try again";
const SIGNED_SUGGESTION: &str = "remove the signature explicitly with `Signed::unsign()` before editing; the first write then publishes the document without it";
const IWORK_SUGGESTION: &str =
    "open Apple iWork packages in a build with the `pages`, `keynote` and `numbers` cargo features";

/// The outcome of classifying one leaf error.
struct Class {
    code: Code,
    suggestion: Option<Cow<'static, str>>,
    detected: Option<FileFormat>,
}

impl Class {
    const fn of(code: Code) -> Self {
        let suggestion = match code {
            Code::LimitExceeded => Some(Cow::Borrowed(LIMIT_SUGGESTION)),
            _ => None,
        };
        Self {
            code,
            suggestion,
            detected: None,
        }
    }

    const fn suggest(code: Code, suggestion: &'static str) -> Self {
        Self {
            code,
            suggestion: Some(Cow::Borrowed(suggestion)),
            detected: None,
        }
    }
}

/// Maps a leaf error to its unified classification.
trait Classify {
    fn classify(&self) -> Class;
}

impl Classify for crate::transaction::Error {
    fn classify(&self) -> Class {
        Class::of(match self {
            Self::PatchConflict => Code::InvalidInput,
            Self::InvalidLimit { .. } => Code::InvalidArgument,
            Self::Limit { .. } => Code::LimitExceeded,
            Self::Allocation { .. } => Code::InternalError,
        })
    }
}

impl Classify for crate::batch::Error {
    fn classify(&self) -> Class {
        Class::of(match self {
            Self::InvalidLimit { .. } => Code::InvalidArgument,
            Self::Limit { .. } => Code::LimitExceeded,
            Self::Malformed { .. } => Code::CorruptBatch,
            Self::Allocation { .. } => Code::InternalError,
        })
    }
}

impl Classify for crate::raw::Error {
    fn classify(&self) -> Class {
        Class::of(self.code())
    }
}

impl Classify for crate::merge::Error {
    fn classify(&self) -> Class {
        Class::of(match self {
            Self::Json { .. } | Self::NotAnObject | Self::NotARecord { .. } => Code::InvalidJson,
            Self::TooDeep => Code::LimitExceeded,
        })
    }
}

impl Classify for crate::import::Error {
    fn classify(&self) -> Class {
        Class::of(match self {
            Self::InvalidOption { .. } => Code::InvalidArgument,
            Self::FieldTooLong { .. } => Code::InvalidValue,
            Self::Encoding { .. }
            | Self::UnterminatedQuote { .. }
            | Self::InvalidCharacter { .. } => Code::InvalidInput,
        })
    }
}

#[cfg(all(
    feature = "eval",
    any(feature = "xlsx", feature = "xlsb", feature = "xls")
))]
impl Classify for crate::calc::Error {
    fn classify(&self) -> Class {
        Class::of(match self {
            Self::NotACell => Code::InvalidArgument,
            Self::Unsupported | Self::Volatile => Code::UnsupportedCapability,
            Self::Limit { .. } => Code::LimitExceeded,
            Self::Suspended => Code::InternalError,
        })
    }
}

impl Classify for io::Error {
    fn classify(&self) -> Class {
        // Format-owned cooperative adapters carry cancellation and budget
        // refusals through `Read`/`Write` as an execution error.
        if let Some(execution) = self
            .get_ref()
            .and_then(|source| source.downcast_ref::<ExecutionError>())
        {
            return execution.classify();
        }
        match self.kind() {
            io::ErrorKind::NotFound => Class::of(Code::FileNotFound),
            io::ErrorKind::AlreadyExists => Class::of(Code::FileExists),
            _ => Class::of(Code::IoError),
        }
    }
}

impl Classify for ExecutionError {
    fn classify(&self) -> Class {
        match self {
            Self::Cancelled => Class::of(Code::Cancelled),
            Self::ResourceLimit(_) => Class::of(Code::LimitExceeded),
            // The remaining variants reject an inconsistent caller policy.
            _ => Class::of(Code::InvalidArgument),
        }
    }
}

/// The one classification of the shared edit refusal (record 0351): every
/// format's `Error::Refused(Refusal)` reports its kind's code, so the same
/// refusal is the same code in every format.
impl Classify for litchi_core::edit::Refusal {
    fn classify(&self) -> Class {
        self.kind().classify()
    }
}

impl Classify for litchi_core::edit::RefusalKind {
    fn classify(&self) -> Class {
        Class::of(refusal_code(*self))
    }
}

/// The one code of each refusal kind (record 0351 §2.4), for callers that
/// map a leaf error to a [`Code`] directly (the iWork leaf errors, record
/// 0361).
pub(crate) const fn refusal_code(kind: litchi_core::edit::RefusalKind) -> Code {
    use litchi_core::edit::RefusalKind;

    match kind {
        RefusalKind::UnsupportedOperation | RefusalKind::UnsupportedContent => {
            Code::UnsupportedCapability
        },
        RefusalKind::UnsupportedTarget => Code::UnsupportedType,
        RefusalKind::UnsupportedProperty => Code::UnsupportedProperty,
        RefusalKind::InvalidValue => Code::InvalidValue,
        RefusalKind::NotFound => Code::NotFound,
        RefusalKind::Protected => Code::DocumentProtected,
        // A kind added later is unsupported until it is classified.
        _ => Code::UnsupportedCapability,
    }
}

/// The grammar's own code decides: `invalid_path` or `invalid_selector` for
/// syntax, `max_depth_exceeded` for the step and nesting limits. The grammar's
/// corrective hint, when it has one, becomes the suggestion.
impl Classify for litchi_core::selector::ParseError {
    fn classify(&self) -> Class {
        use litchi_core::selector::{ErrorCode, Grammar};

        let code = match self.code() {
            ErrorCode::InvalidPath => Code::InvalidPath,
            ErrorCode::InvalidSelector => Code::InvalidSelector,
            ErrorCode::MaxDepthExceeded => Code::MaxDepthExceeded,
            ErrorCode::PathNotFound => Code::PathNotFound,
            _ => match self.grammar() {
                Grammar::Query => Code::InvalidSelector,
                _ => Code::InvalidPath,
            },
        };
        match self.suggestion() {
            Some(suggestion) => Class::suggest(code, suggestion),
            None => Class::of(code),
        }
    }
}

impl Classify for litchi_core::Error {
    fn classify(&self) -> Class {
        match self {
            Self::PatchConflict => Class::of(Code::InvalidInput),
            Self::InvalidLimit { .. } => Class::of(Code::InvalidArgument),
            Self::Io(error) => error.classify(),
            Self::SourceChanged { .. } => Class::of(Code::IoError),
            Self::NotOfficeFile => Class::of(Code::UnsupportedFormat),
            Self::UnexpectedFormat { detected } => Class {
                detected: Some(*detected),
                ..Class::of(Code::UnsupportedFormat)
            },
            Self::FeatureDisabled(feature) => Class {
                code: Code::UnsupportedFormat,
                suggestion: Some(format!("enable the `{feature}` cargo feature").into()),
                detected: None,
            },
            Self::ParseError(_)
            | Self::InvalidFormat(_)
            | Self::CorruptedFile(_)
            | Self::ComponentNotFound(_)
            | Self::XmlError(_)
            | Self::XmlCompactness { .. }
            | Self::InvalidContentType { .. }
            | Self::ZipError(_) => Class::of(Code::CorruptFile),
            Self::ResourceLimit(_) => Class::of(Code::LimitExceeded),
            Self::Unsupported(_) => Class::of(Code::UnsupportedCapability),
            Self::Refused(refusal) => refusal.classify(),
            Self::Allocation { .. } | Self::Committed(_) | Self::Other(_) => {
                Class::of(Code::InternalError)
            },
            _ => Class::of(Code::CorruptFile),
        }
    }
}

#[cfg(any(feature = "docx", feature = "pptx", feature = "xlsx", feature = "xlsb"))]
impl Classify for litchi_drawingml::Error {
    fn classify(&self) -> Class {
        match self {
            Self::Limit { .. } => Class::of(Code::LimitExceeded),
            Self::InvalidLimit { .. } => Class::of(Code::InvalidArgument),
            // A patch made for a different immutable owner is caller input,
            // not a malformed document or a transient I/O failure.
            Self::PatchConflict => Class::of(Code::InvalidInput),
            Self::Io(error) => error.classify(),
            Self::Mce(error) if mce_invalid_limit(error) => Class::of(Code::InvalidArgument),
            Self::Mce(error) if mce_limit(error) => Class::of(Code::LimitExceeded),
            Self::Invalid(_) => Class::of(Code::InvalidValue),
            Self::Xml(_) | Self::Mce(_) | Self::Common(_) => Class::of(Code::CorruptFile),
            _ => Class::of(Code::CorruptFile),
        }
    }
}

/// Whether a shared OOXML error is a typed budget of a shared family
/// (record 0360).
#[cfg(any(feature = "docx", feature = "pptx", feature = "xlsx", feature = "xlsb"))]
fn mce_limit(error: &litchi_ooxml_common::mce::Error) -> bool {
    matches!(error, litchi_ooxml_common::mce::Error::LimitExceeded { .. })
}

/// Whether a shared MCE policy rejected the caller's configured bound before
/// processing started. This is an invalid argument, not an exhausted document
/// resource (record 0360).
#[cfg(any(feature = "docx", feature = "pptx", feature = "xlsx", feature = "xlsb"))]
fn mce_invalid_limit(error: &litchi_ooxml_common::mce::Error) -> bool {
    matches!(error, litchi_ooxml_common::mce::Error::InvalidLimit { .. })
}

/// Whether a shared OOXML error is a typed budget of a shared family,
/// including a typed MCE budget nested in the common wrapper (record 0360).
#[cfg(any(feature = "docx", feature = "pptx", feature = "xlsx", feature = "xlsb"))]
fn common_limit(error: &litchi_ooxml_common::Error) -> bool {
    match error {
        litchi_ooxml_common::Error::Limit { .. } => true,
        litchi_ooxml_common::Error::Mce(error) => mce_limit(error),
        _ => false,
    }
}

#[cfg(any(feature = "docx", feature = "pptx", feature = "xlsx", feature = "xlsb"))]
fn common_invalid_limit(error: &litchi_ooxml_common::Error) -> bool {
    matches!(
        error,
        litchi_ooxml_common::Error::Mce(error) if mce_invalid_limit(error)
    )
}

#[cfg(any(feature = "docx", feature = "pptx", feature = "xlsx", feature = "xlsb"))]
fn classify_common_semantic(error: &litchi_ooxml_common::Error) -> Option<Class> {
    match error {
        litchi_ooxml_common::Error::Refused(refusal) => Some(refusal.classify()),
        litchi_ooxml_common::Error::PatchConflict => Some(Class::of(Code::InvalidInput)),
        litchi_ooxml_common::Error::Signed => {
            Some(Class::suggest(Code::DocumentProtected, SIGNED_SUGGESTION))
        },
        litchi_ooxml_common::Error::SourceChanged { .. } => Some(Class::of(Code::IoError)),
        litchi_ooxml_common::Error::Internal(_) => Some(Class::of(Code::InternalError)),
        _ => None,
    }
}

#[cfg(any(feature = "docx", feature = "pptx", feature = "xlsx", feature = "xlsb"))]
impl Classify for litchi_opc::OpcError {
    fn classify(&self) -> Class {
        match self {
            Self::XmlPublication { source, .. } if source.is_limit() => {
                Class::of(Code::LimitExceeded)
            },
            Self::XmlPublication { source, .. } if source.is_allocation() => {
                Class::of(Code::InternalError)
            },
            Self::IoError(error) => error.classify(),
            Self::PackageNotFound(_) => Class::of(Code::FileNotFound),
            Self::SourceChanged { .. }
            | Self::CentralDirectorySpool { .. }
            | Self::IncompleteOutput { .. } => Class::of(Code::IoError),
            // A durable replay or restore supplied an authenticated artifact
            // proof that does not describe the current/original artifact.
            // The proof is caller-owned input, while a source that changes
            // during the operation remains an I/O freshness failure above.
            Self::SourceArtifactMismatch { .. } => Class::of(Code::InvalidInput),
            Self::Limit { .. }
            | Self::SourcePartSpliceLimit { .. }
            | Self::CentralDirectorySpoolLimitExceeded { .. } => Class::of(Code::LimitExceeded),
            Self::InvalidLimit { .. }
            | Self::InvalidSourcePartSpliceLimit { .. }
            | Self::UnsupportedExecutionAffinity => Class::of(Code::InvalidArgument),
            Self::Cancelled => Class::of(Code::Cancelled),
            Self::Execution(error) => error.classify(),
            Self::SignedSourceRequiresExplicitPolicy => {
                Class::suggest(Code::DocumentProtected, SIGNED_SUGGESTION)
            },
            Self::SourceBackedOverlayUnavailable { .. } | Self::PreservationUnavailable { .. } => {
                Class::of(Code::UnsupportedCapability)
            },
            Self::Allocation { .. }
            | Self::CollectionAllocation { .. }
            | Self::Committed { .. }
            | Self::OperationAccountingOverflow { .. }
            | Self::SourceBackedBatchWorkerPanic { .. }
            | Self::SourceBackedBatchInvariant { .. }
            | Self::ManagedPartDataArcEscape
            | Self::ManagedPackageMaterialization
            | Self::ForeignPartView
            | Self::ParallelRead(_)
            | Self::ValidationReport(_) => Class::of(Code::InternalError),
            _ => Class::of(Code::CorruptFile),
        }
    }
}

#[cfg(any(
    feature = "doc",
    feature = "xls",
    feature = "ppt",
    feature = "encryption"
))]
impl Classify for litchi_cfb::OleError {
    fn classify(&self) -> Class {
        match self {
            Self::Io(error) => error.classify(),
            Self::SourceChanged { .. } => Class::of(Code::IoError),
            Self::NotOleFile => Class::of(Code::UnsupportedFormat),
            Self::LimitExceeded { .. } => Class::of(Code::LimitExceeded),
            Self::InvalidLimit { .. } => Class::of(Code::InvalidArgument),
            Self::InvalidFormat(_)
            | Self::InvalidData(_)
            | Self::CorruptedFile(_)
            | Self::StreamNotFound => Class::of(Code::CorruptFile),
            Self::Allocation { .. } | Self::Committed { .. } => Class::of(Code::InternalError),
        }
    }
}

#[cfg(feature = "docx")]
impl Classify for litchi_docx::Error {
    fn classify(&self) -> Class {
        match self {
            #[cfg(feature = "encryption")]
            Self::Encryption(error) => error.classify(),
            #[cfg(feature = "encryption")]
            Self::EncryptionPolicy { .. } => Class::of(Code::UnsupportedCapability),
            Self::UnsupportedOpenPolicy { .. } => Class::of(Code::UnsupportedCapability),
            Self::Opc(error) => error.classify(),
            Self::Io(error) => error.classify(),
            Self::Drawing(error) => error.classify(),
            Self::Signed => Class::suggest(Code::DocumentProtected, SIGNED_SUGGESTION),
            Self::SourceChanged { .. } => Class::of(Code::IoError),
            // A nonempty accepted prefix means the sink observed a partial
            // publication, regardless of whether the producer or OPC layer
            // raised the secondary cause. A zero-prefix stream preserves the
            // original typed producer/OPC classification.
            Self::ArtifactStream(error) if error.written().is_some_and(|written| written != 0) => {
                Class::of(Code::IoError)
            },
            Self::ArtifactStream(error) => error
                .opc_error()
                .map(|error| error.classify())
                .or_else(|| error.producer_error().map(|error| error.classify()))
                .unwrap_or_else(|| Class::of(Code::InternalError)),
            Self::TailAppend(error) => classify_docx_tail_append(error),
            Self::InkLimit { .. }
            | Self::AltChunkLimit { .. }
            | Self::GlossaryLimit { .. }
            | Self::ContentControlLimit { .. }
            | Self::ConflictLimit { .. }
            | Self::ChartLimit { .. }
            | Self::FontLimit { .. }
            | Self::BibliographyLimit { .. }
            | Self::WebLimit { .. }
            | Self::StoryLimit { .. }
            | Self::NumberingLimit { .. }
            | Self::MailMergeLimit { .. }
            | Self::StylesWithEffectsLimit { .. }
            | Self::SettingsLimit { .. }
            | Self::RevisionLimit { .. }
            | Self::BlockReadLimit { .. }
            | Self::ExternalHyperlinkDetachmentLimit { .. }
            | Self::ExternalHyperlinkRedactionLimit { .. }
            | Self::SectionInventoryLimit { .. } => Class::of(Code::LimitExceeded),
            // The shared OOXML families' typed budgets (record 0360).
            Self::Common(litchi_ooxml_common::Error::Opc(error)) => error.classify(),
            Self::Common(error) if classify_common_semantic(error).is_some() => {
                classify_common_semantic(error).unwrap_or_else(|| Class::of(Code::CorruptFile))
            },
            Self::Common(error) if common_invalid_limit(error) => Class::of(Code::InvalidArgument),
            Self::Common(error) if common_limit(error) => Class::of(Code::LimitExceeded),
            Self::Mce(error) if mce_invalid_limit(error) => Class::of(Code::InvalidArgument),
            Self::Mce(error) if mce_limit(error) => Class::of(Code::LimitExceeded),
            Self::OutOfBounds { .. } => Class::of(Code::OutOfRange),
            Self::Refused(refusal) => refusal.classify(),
            Self::PatchConflict { .. } => Class::of(Code::InvalidInput),
            Self::Allocation { .. }
            | Self::RevisionAllocation { .. }
            | Self::Internal { .. }
            | Self::Other(_) => Class::of(Code::InternalError),
            _ => Class::of(Code::CorruptFile),
        }
    }
}

#[cfg(feature = "docx")]
fn classify_docx_tail_append(error: &litchi_docx::source_backed::tail_append::Error) -> Class {
    use litchi_docx::source_backed::tail_append::Error as TailAppendError;

    match error {
        TailAppendError::Document(error) => error.classify(),
        TailAppendError::Refused(refusal) => refusal.kind().classify(),
        TailAppendError::Limit { .. } => Class::of(Code::LimitExceeded),
        TailAppendError::Allocation { .. } => Class::of(Code::InternalError),
        TailAppendError::Execution(error) => error.classify(),
        TailAppendError::Scan(_) => Class::of(Code::CorruptFile),
        TailAppendError::Io(error) => error.classify(),
        TailAppendError::PatchConflict => Class::of(Code::InvalidInput),
        _ => Class::of(Code::CorruptFile),
    }
}

#[cfg(feature = "docx")]
impl Classify for litchi_docx::document::Error {
    fn classify(&self) -> Class {
        match self {
            Self::Document(error) => error.classify(),
            Self::OutOfBounds { .. } => Class::of(Code::OutOfRange),
            Self::Refused { reason, .. } => reason.kind().classify(),
            Self::Transfer(refusal) => refusal.kind().classify(),
            Self::Limit { .. } => Class::of(Code::LimitExceeded),
            // A patch for another base: the caller supplied it (record 0351).
            Self::PatchConflict => Class::of(Code::InvalidInput),
            Self::SemanticPrecondition => Class::of(Code::InvalidArgument),
            Self::Durable(_) | Self::Composition(_) | Self::InvalidDurable(_) => {
                Class::of(Code::InvalidInput)
            },
            _ => Class::of(Code::CorruptFile),
        }
    }
}

#[cfg(feature = "docx")]
impl Classify for litchi_docx::metadata::Error {
    fn classify(&self) -> Class {
        match self {
            Self::Package(error) => error.classify(),
            Self::Signed => Class::of(Code::DocumentProtected),
            Self::Limit { .. } => Class::of(Code::LimitExceeded),
            Self::PatchConflict => Class::of(Code::InvalidInput),
            // Application and custom properties (record 0342).
            Self::Properties(error) => match error {
                litchi_ooxml_common::properties::Error::Refused(refusal) => refusal.classify(),
                litchi_ooxml_common::properties::Error::Invalid(_) => Class::of(Code::InvalidValue),
                litchi_ooxml_common::properties::Error::Limit { .. } => {
                    Class::of(Code::LimitExceeded)
                },
                litchi_ooxml_common::properties::Error::Signed => {
                    Class::of(Code::DocumentProtected)
                },
                litchi_ooxml_common::properties::Error::PatchConflict => {
                    Class::of(Code::InvalidInput)
                },
                litchi_ooxml_common::properties::Error::SourceChanged { .. } => {
                    Class::of(Code::IoError)
                },
                litchi_ooxml_common::properties::Error::Internal(_) => {
                    Class::of(Code::InternalError)
                },
                _ => Class::of(Code::CorruptFile),
            },
            _ => Class::of(Code::InternalError),
        }
    }
}

#[cfg(feature = "pptx")]
impl Classify for litchi_pptx::Error {
    fn classify(&self) -> Class {
        match self {
            #[cfg(feature = "encryption")]
            Self::Encryption(error) => error.classify(),
            #[cfg(feature = "encryption")]
            Self::EncryptionPolicy { source, .. } => match source {
                litchi_ooxml_common::package_encryption::PolicyError::OrdinaryOutputWouldRemoveEncryption { .. } => Class::of(Code::DocumentProtected),
                litchi_ooxml_common::package_encryption::PolicyError::NoRetainedMode => Class::of(Code::UnsupportedCapability),
            },
            Self::Opc(error) => error.classify(),
            Self::Io(error) => error.classify(),
            Self::Drawing(error) => error.classify(),
            Self::Signed => Class::suggest(Code::DocumentProtected, SIGNED_SUGGESTION),
            Self::SourceChanged { .. } => Class::of(Code::IoError),
            // A patch for another base: the caller supplied it (record 0351).
            Self::PatchConflict => Class::of(Code::InvalidInput),
            // Every well-formed refusal is classified by its shared kind;
            // operation and reason remain content-free diagnostics only.
            Self::Refused(refusal) => refusal.classify(),
            Self::UnsupportedOpenPolicy { .. } => Class::of(Code::InvalidArgument),
            Self::InvalidLimit { .. } => Class::of(Code::InvalidArgument),
            // The core-properties reader reports its OPC failures, among them
            // a source change or cancellation during the read, through the
            // shared OOXML error.
            Self::CustomProperties(litchi_ooxml_common::Error::Opc(error)) => error.classify(),
            Self::CustomProperties(error) if classify_common_semantic(error).is_some() => {
                classify_common_semantic(error).unwrap_or_else(|| Class::of(Code::CorruptFile))
            },
            Self::CustomProperties(error) if common_invalid_limit(error) => {
                Class::of(Code::InvalidArgument)
            },
            Self::CustomProperties(error) if common_limit(error) => {
                Class::of(Code::LimitExceeded)
            },
            Self::MarkupCompatibility(error) if mce_invalid_limit(error) => {
                Class::of(Code::InvalidArgument)
            },
            Self::MarkupCompatibility(error) if mce_limit(error) => {
                Class::of(Code::LimitExceeded)
            },
            Self::Limit { .. } => Class::of(Code::LimitExceeded),
            Self::IndexOutOfBounds { .. }
            | Self::SlideIndexOutOfBounds { .. }
            | Self::FontIndexOutOfBounds { .. } => Class::of(Code::OutOfRange),
            Self::NameNotFound(_) | Self::SlideNameNotFound(_) | Self::SlideLayoutNameNotFound(_) | Self::FontNotFound(_) => {
                Class::of(Code::NotFound)
            },
            Self::Allocation { .. }
            | Self::CapacityOverflow { .. }
            | Self::Internal { .. } => Class::of(Code::InternalError),
            _ => Class::of(Code::CorruptFile),
        }
    }
}

#[cfg(feature = "xlsx")]
fn classify_xlsx_join_source(source: Option<&litchi_xlsx::Error>) -> Class {
    source.map_or_else(|| Class::of(Code::CorruptFile), |source| source.classify())
}

#[cfg(feature = "xlsx")]
impl Classify for litchi_xlsx::Error {
    fn classify(&self) -> Class {
        match self {
            Self::SheetNameDependency(error) => match error {
                litchi_ooxml_common::xml::sheet_name::Error::Blocked(_) => {
                    Class::of(Code::UnsupportedCapability)
                },
                litchi_ooxml_common::xml::sheet_name::Error::Limit(_) => {
                    Class::of(Code::LimitExceeded)
                },
                litchi_ooxml_common::xml::sheet_name::Error::Allocation { .. } => {
                    Class::of(Code::InternalError)
                },
                _ => Class::of(Code::CorruptFile),
            },
            Self::Package(error) => error.classify(),
            Self::SourceChanged { .. } => Class::of(Code::IoError),
            Self::Internal(_) => Class::of(Code::InternalError),
            Self::Common(litchi_ooxml_common::Error::Opc(error)) => error.classify(),
            Self::Common(error) if classify_common_semantic(error).is_some() => {
                classify_common_semantic(error).unwrap_or_else(|| Class::of(Code::CorruptFile))
            },
            Self::Common(error) if common_invalid_limit(error) => {
                Class::of(Code::InvalidArgument)
            },
            Self::Common(error) if common_limit(error) => Class::of(Code::LimitExceeded),
            Self::MarkupCompatibility(error) if mce_invalid_limit(error) => {
                Class::of(Code::InvalidArgument)
            },
            Self::MarkupCompatibility(error) if mce_limit(error) => {
                Class::of(Code::LimitExceeded)
            },
            #[cfg(feature = "encryption")]
            Self::Encryption(error) => error.classify(),
            #[cfg(feature = "encryption")]
            Self::EncryptionPolicy { source, .. } => match source {
                litchi_ooxml_common::package_encryption::PolicyError::OrdinaryOutputWouldRemoveEncryption { .. } => Class::suggest(Code::DocumentProtected, "use SaveOptions::reencrypt for encrypted publication"),
                litchi_ooxml_common::package_encryption::PolicyError::NoRetainedMode => Class::of(Code::UnsupportedCapability),
            },
            Self::ResourceLimit(_) | Self::Limit { .. } => Class::of(Code::LimitExceeded),
            Self::FormControl(litchi_xlsx::form_control::FormControlError::Limit { .. }) => {
                Class::of(Code::LimitExceeded)
            },
            // Typed refusals, patches for another base and invalid configured
            // limits (record 0360).
            Self::Refused(refusal) => refusal.classify(),
            Self::PatchConflict { .. } => Class::of(Code::InvalidInput),
            Self::InvalidLimit { .. } => Class::of(Code::InvalidArgument),
            Self::Signed => Class::suggest(Code::DocumentProtected, SIGNED_SUGGESTION),
            Self::Unsupported { .. } | Self::UnsupportedSelector => {
                Class::of(Code::UnsupportedCapability)
            },
            Self::NotWorksheet { .. } => Class::of(Code::UnsupportedType),
            Self::EditBlocked { reason, .. } => match reason {
                litchi_xlsx::EditBlock::ProtectedSheet => Class::of(Code::DocumentProtected),
                litchi_xlsx::EditBlock::CoveredMerge
                | litchi_xlsx::EditBlock::GroupFormula
                | litchi_xlsx::EditBlock::DataValidation
                | litchi_xlsx::EditBlock::UnknownCell
                | litchi_xlsx::EditBlock::MarkupCompatibility => Class::of(Code::UnsupportedCapability),
                _ => Class::of(Code::CorruptFile),
            },
            Self::MergeEditBlocked { reason, .. } => match reason {
                litchi_xlsx::MergeEditBlock::ProtectedSheet => Class::of(Code::DocumentProtected),
                litchi_xlsx::MergeEditBlock::SingleCell => Class::of(Code::InvalidValue),
                litchi_xlsx::MergeEditBlock::Overlap { .. }
                | litchi_xlsx::MergeEditBlock::TableOverlap { .. }
                | litchi_xlsx::MergeEditBlock::FollowerContent { .. }
                | litchi_xlsx::MergeEditBlock::GroupFormula
                | litchi_xlsx::MergeEditBlock::MarkupCompatibility
                | litchi_xlsx::MergeEditBlock::UnmodeledPayload => Class::of(Code::UnsupportedCapability),
                _ => Class::of(Code::CorruptFile),
            },
            Self::RowEditBlocked { reason, .. } => match reason {
                litchi_xlsx::RowEditBlock::ProtectedSheet => Class::of(Code::DocumentProtected),
                _ => Class::of(Code::CorruptFile),
            },
            Self::ColumnEditBlocked { reason, .. } => match reason {
                litchi_xlsx::ColumnEditBlock::ProtectedSheet => Class::of(Code::DocumentProtected),
                litchi_xlsx::ColumnEditBlock::MarkupCompatibility
                | litchi_xlsx::ColumnEditBlock::StyleNeedsWidth => Class::of(Code::UnsupportedCapability),
                _ => Class::of(Code::CorruptFile),
            },
            Self::TabEditBlocked { reason, .. } => match reason {
                litchi_xlsx::TabEditBlock::ProtectedWorkbook => Class::of(Code::DocumentProtected),
                litchi_xlsx::TabEditBlock::LastVisibleTab | litchi_xlsx::TabEditBlock::NotVisible => {
                    Class::of(Code::InvalidArgument)
                },
                litchi_xlsx::TabEditBlock::SheetLimit | litchi_xlsx::TabEditBlock::ActiveTabLimit => {
                    Class::of(Code::LimitExceeded)
                },
                litchi_xlsx::TabEditBlock::TrackedWorkbook
                | litchi_xlsx::TabEditBlock::MarkupCompatibility => Class::of(Code::UnsupportedCapability),
                _ => Class::of(Code::CorruptFile),
            },
            Self::RenameBlocked { .. } => Class::of(Code::UnsupportedCapability),
            Self::SheetRemoveBlocked { reason, .. } => match reason {
                litchi_xlsx::RemoveBlock::LastSheet => Class::of(Code::InvalidArgument),
                litchi_xlsx::RemoveBlock::MixedEdit
                | litchi_xlsx::RemoveBlock::IncomingReference
                | litchi_xlsx::RemoveBlock::UnmodeledReference
                | litchi_xlsx::RemoveBlock::MarkupCompatibility
                | litchi_xlsx::RemoveBlock::MacroProject
                | litchi_xlsx::RemoveBlock::IncomingRelationship => Class::of(Code::UnsupportedCapability),
                _ => Class::of(Code::CorruptFile),
            },
            // A join keeps the original normalization failure as its typed
            // source.  Preserve the two operation-boundary classes through
            // that wrapper; a join with no retained source keeps the
            // historical malformed-input fallback, while a retained source
            // keeps its own established classification.
            Self::Join(error) => classify_xlsx_join_source(error.normalization_source()),
            Self::Allocation { .. } => Class::of(Code::InternalError),
            // A coordinate or range the leaf read from the file. One the
            // caller passed is reclassified where it was received
            // (`Error::caused_by_argument`).
            Self::Coordinate(_) | Self::Range(_) => Class::of(Code::CorruptFile),
            _ => Class::of(Code::CorruptFile),
        }
    }
}

#[cfg(feature = "xlsb")]
impl Classify for litchi_xlsb::package::PackageError {
    fn classify(&self) -> Class {
        match self {
            Self::SheetNameDependency(error) => match error {
                litchi_ooxml_common::xml::sheet_name::Error::Blocked(_) => {
                    Class::of(Code::UnsupportedCapability)
                },
                litchi_ooxml_common::xml::sheet_name::Error::Limit(_) => {
                    Class::of(Code::LimitExceeded)
                },
                litchi_ooxml_common::xml::sheet_name::Error::Allocation { .. } => {
                    Class::of(Code::InternalError)
                },
                _ => Class::of(Code::CorruptFile),
            },
            Self::Io(error) => error.classify(),
            Self::Opc(error) => error.classify(),
            Self::Common(error) if common_invalid_limit(error) => Class::of(Code::InvalidArgument),
            Self::SourceChanged { .. } => Class::of(Code::IoError),
            Self::Internal(_) => Class::of(Code::InternalError),
            Self::Common(litchi_ooxml_common::Error::Opc(error)) => error.classify(),
            Self::Common(error) if classify_common_semantic(error).is_some() => {
                classify_common_semantic(error).unwrap_or_else(|| Class::of(Code::CorruptFile))
            },
            Self::Common(error) if common_limit(error) => Class::of(Code::LimitExceeded),
            Self::Wire(error) => error.classify(),
            Self::Execution(error) => error.classify(),
            Self::ExternalLinkLimitExceeded { .. } | Self::Limit { .. } => {
                Class::of(Code::LimitExceeded)
            },
            Self::UnsupportedOpenPolicy { .. } => Class::of(Code::InvalidArgument),
            // Typed refusals, patches for another base and invalid configured
            // limits (record 0360b).
            Self::Refused(refusal) => refusal.classify(),
            Self::PatchConflict { .. } => Class::of(Code::InvalidInput),
            Self::InvalidLimit { .. } => Class::of(Code::InvalidArgument),
            Self::Signed => Class::suggest(Code::DocumentProtected, SIGNED_SUGGESTION),
            Self::PasswordProtected => Class::suggest(Code::DocumentProtected, PASSWORD_SUGGESTION),
            Self::UnsupportedFeature(_) => Class::of(Code::UnsupportedCapability),
            Self::WorksheetNotFound(_) => Class::of(Code::NotFound),
            Self::Allocation { .. } | Self::CapacityOverflow { .. } => {
                Class::of(Code::InternalError)
            },
            _ => Class::of(Code::CorruptFile),
        }
    }
}

#[cfg(feature = "xlsb")]
impl Classify for litchi_xlsb::Error {
    fn classify(&self) -> Class {
        match self {
            Self::Io(error) => error.classify(),
            Self::PayloadLimit { .. } | Self::StringLimit { .. } => Class::of(Code::LimitExceeded),
            Self::InvalidLimit { .. } => Class::of(Code::InvalidArgument),
            _ => Class::of(Code::CorruptFile),
        }
    }
}

#[cfg(feature = "doc")]
impl Classify for litchi_doc::Error {
    fn classify(&self) -> Class {
        match self {
            Self::Object(error) => error.classify(),
            Self::Io(error) => error.classify(),
            Self::Ole(error) => error.classify(),
            Self::CustomXml(error) => error.classify(),
            Self::Vba(error) => error.classify(),
            Self::Signature(error) => error.classify(),
            Self::Crypto(error) => error.classify(),
            Self::PatchConflict => Class::of(Code::InvalidInput),
            Self::Limit { .. } => Class::of(Code::LimitExceeded),
            Self::InvalidLimit { .. } => Class::of(Code::InvalidArgument),
            Self::PasswordRequired => Class::suggest(Code::DocumentProtected, PASSWORD_SUGGESTION),
            Self::ProtectionDenied(_) => Class::suggest(
                Code::DocumentProtected,
                "authorize the protected edit explicitly",
            ),
            Self::InvalidPassword => {
                Class::suggest(Code::InvalidArgument, WRONG_PASSWORD_SUGGESTION)
            },
            Self::UnsupportedEncryption(_) => Class::of(Code::UnsupportedCapability),
            Self::UnsupportedVersion { .. } => Class::of(Code::UnsupportedFormat),
            _ => Class::of(Code::CorruptFile),
        }
    }
}

#[cfg(any(feature = "doc", feature = "ppt"))]
impl Classify for litchi_core::patch::PatchError {
    fn classify(&self) -> Class {
        use litchi_core::patch::PatchError;

        match self {
            PatchError::InvalidText { .. }
            | PatchError::InvalidBlobId
            | PatchError::UnsupportedVersion(_) => Class::of(Code::InvalidArgument),
            PatchError::BlobLimit { .. }
            | PatchError::HistoryWeight { .. }
            | PatchError::JsonLimit { .. } => Class::of(Code::LimitExceeded),
            PatchError::Allocation | PatchError::HistoryInvariant => Class::of(Code::InternalError),
            PatchError::BlobBase64(_)
            | PatchError::Json(_)
            | PatchError::NonCanonicalJson
            | PatchError::BlobDigestMismatch
            | PatchError::IncompatibleBlobLimits { .. }
            | PatchError::OperationPairing => Class::of(Code::CorruptFile),
            _ => Class::of(Code::CorruptFile),
        }
    }
}

#[cfg(any(feature = "doc", feature = "ppt"))]
impl Classify for litchi_core::patch::CompositionError {
    fn classify(&self) -> Class {
        use litchi_core::patch::CompositionError;

        match self {
            CompositionError::InvalidId | CompositionError::InvalidEffect => {
                Class::of(Code::InvalidArgument)
            },
            CompositionError::Limit { .. } => Class::of(Code::LimitExceeded),
            _ => Class::of(Code::CorruptFile),
        }
    }
}

#[cfg(feature = "doc")]
impl Classify for litchi_doc::package::edit::Error {
    fn classify(&self) -> Class {
        match self {
            Self::Package(error) => error.classify(),
            Self::Body(error) => error.classify(),
            Self::Limit { .. } => Class::of(Code::LimitExceeded),
            _ => Class::of(Code::CorruptFile),
        }
    }
}

#[cfg(feature = "doc")]
impl Classify for litchi_doc::body_text::Error {
    fn classify(&self) -> Class {
        use litchi_doc::body_text::Refusal;

        match self {
            Self::Invalid(error) => error.classify(),
            // The protected kinds keep their specific suggestions.
            Self::Refused(Refusal::Protected | Refusal::UnknownProtection) => Class::suggest(
                Code::DocumentProtected,
                "authorize the protected edit explicitly",
            ),
            Self::Refused(Refusal::SignedSource) => {
                Class::suggest(Code::DocumentProtected, SIGNED_SUGGESTION)
            },
            Self::Refused(Refusal::Encrypted) => Class::suggest(
                Code::DocumentProtected,
                "native encrypted DOC edits are not supported; edit a decrypted copy",
            ),
            // Not refusals of a well-formed request (record 0362): a missing
            // position is `out_of_range` as DOCX `OutOfBounds` and the RTF
            // positional errors are, a taken resource identifier is a
            // duplicate, an exhausted budget is a limit, and removing the
            // sole paragraph is the DOC family's final-element argument error.
            Self::Refused(
                Refusal::ParagraphNotFound | Refusal::TargetNotFound | Refusal::SectionNotFound,
            ) => Class::of(Code::OutOfRange),
            Self::Refused(Refusal::ResourceCollision { .. }) => Class::of(Code::DuplicateName),
            Self::Refused(Refusal::OperationLimit { .. } | Refusal::ReplacementLimit { .. }) => {
                Class::of(Code::LimitExceeded)
            },
            Self::Refused(Refusal::SoleParagraph) => Class::of(Code::InvalidArgument),
            Self::Refused(Refusal::BookmarkName) => Class::suggest(
                Code::InvalidValue,
                "name the bookmark with 1 through 39 letters, digits, `_`, `.` or `-`, unused by another bookmark",
            ),
            // Every other refusal by its kind alone (record 0351 §2.4).
            Self::Refused(refusal) => refusal.kind().classify(),
            Self::PatchConflict => Class::of(Code::InvalidInput),
            // A three-way semantic overlap remains a caller-resolvable
            // invalid input; stale replay uses the distinct PatchConflict
            // variant above.
            Self::Conflict => Class::of(Code::InvalidInput),
            Self::Composition(error) => error.classify(),
            Self::Durable(error) => error.classify(),
            _ => Class::of(Code::CorruptFile),
        }
    }
}

#[cfg(feature = "doc")]
impl Classify for litchi_doc::writer::WriteError {
    fn classify(&self) -> Class {
        use litchi_doc::writer::WriteError;

        match self {
            WriteError::Io(error) => error.classify(),
            WriteError::InvalidData(_) => Class::of(Code::CorruptFile),
            WriteError::Ole(error) => error.classify(),
            // The VBA leaf is intentionally retained as a typed cause. The
            // facade cannot name its private crate dependency, so unsupported
            // VBA details remain a conservative malformed-write category.
            WriteError::Vba(_) => Class::of(Code::CorruptFile),
            WriteError::Package(error) => error.classify(),
            WriteError::Metadata(error) => error.classify(),
            WriteError::EncryptedMetadata => {
                Class::suggest(Code::UnsupportedCapability, ENCRYPTED_METADATA_SUGGESTION)
            },
        }
    }
}

#[cfg(any(feature = "xls", feature = "ppt"))]
impl Classify for litchi_biff::Error {
    fn classify(&self) -> Class {
        match self {
            Self::InvalidLimit { .. } => Class::of(Code::InvalidArgument),
            Self::LimitExceeded { .. } => Class::of(Code::LimitExceeded),
            Self::InvalidKind { .. } => Class::of(Code::InvalidValue),
            Self::Allocation { .. } => Class::of(Code::InternalError),
            Self::TruncatedHeader { .. }
            | Self::TruncatedPayload { .. }
            | Self::EmptyRecord
            | Self::MultipleRecords { .. }
            | Self::SizeOverflow { .. } => Class::of(Code::CorruptFile),
            _ => Class::of(Code::InternalError),
        }
    }
}

#[cfg(any(feature = "doc", feature = "xls", feature = "ppt"))]
impl Classify for litchi_ole_common::Error {
    fn classify(&self) -> Class {
        match self {
            Self::Cfb(error) => error.classify(),
            Self::Limit { .. } => Class::of(Code::LimitExceeded),
            Self::InvalidLimit { .. } => Class::of(Code::InvalidArgument),
            Self::PatchConflict => Class::of(Code::InvalidInput),
            Self::Refused(refusal) => refusal.classify(),
            // `litchi_ole_common::Error` is non-exhaustive by contract. Keep
            // a future native error conservative until its facade mapping is
            // assigned explicitly.
            _ => Class::of(Code::InternalError),
        }
    }
}

#[cfg(any(feature = "xls", feature = "ppt"))]
impl Classify for litchi_ograph::Error {
    fn classify(&self) -> Class {
        match self {
            Self::InvalidLimit { .. } => Class::of(Code::InvalidArgument),
            Self::Limit { .. } => Class::of(Code::LimitExceeded),
            Self::Cfb(error) => error.classify(),
            Self::Biff(error) => error.classify(),
            Self::Refused(refusal) => refusal.classify(),
            Self::PatchConflict { .. } => Class::of(Code::InvalidInput),
            Self::Allocation { .. } => Class::of(Code::InternalError),
            Self::WrongRecord { .. }
            | Self::InvalidRecordLength { .. }
            | Self::InvalidRecordValue { .. }
            | Self::MissingStream { .. }
            | Self::DuplicateStream { .. }
            | Self::UnexpectedEntry { .. }
            | Self::InvalidWorkbook { .. }
            | Self::InvalidChart { .. }
            | Self::InvalidModel { .. }
            | Self::SizeOverflow { .. } => Class::of(Code::CorruptFile),
            _ => Class::of(Code::InternalError),
        }
    }
}

#[cfg(feature = "xls")]
impl Classify for litchi_xls::Error {
    fn classify(&self) -> Class {
        match self {
            Self::Object(error) => error.classify(),
            Self::Vba(error) => error.classify(),
            Self::CustomXml(error) => error.classify(),
            Self::Signature(error) => error.classify(),
            Self::InvalidLimit { .. } | Self::EncryptionEntropyRequired => {
                Class::of(Code::InvalidArgument)
            },
            Self::Metadata(error) => error.classify(),
            Self::EncryptedMetadata => {
                Class::suggest(Code::UnsupportedCapability, ENCRYPTED_METADATA_SUGGESTION)
            },
            Self::Limit { .. } => Class::of(Code::LimitExceeded),
            Self::Io(error) => error.classify(),
            Self::Cfb(error) => error.classify(),
            Self::Graph(error) => error.classify(),
            Self::PasswordProtected | Self::PasswordRequired => {
                Class::suggest(Code::DocumentProtected, PASSWORD_SUGGESTION)
            },
            Self::WeakEncryptionRequiresExplicitPolicy => Class::suggest(
                Code::DocumentProtected,
                "choose a weak-encryption policy explicitly before saving",
            ),
            Self::InvalidPassword => {
                Class::suggest(Code::InvalidArgument, WRONG_PASSWORD_SUGGESTION)
            },
            Self::Refused(refusal) => refusal.classify(),
            Self::PatchConflict { .. } => Class::of(Code::InvalidInput),
            Self::Internal(_) => Class::of(Code::InternalError),
            Self::Signed => Class::suggest(Code::DocumentProtected, SIGNED_SUGGESTION),
            Self::UnsupportedEncryption(_)
            | Self::UnsupportedFeature(_)
            | Self::SourceBoundWorksheetMutation { .. } => Class::of(Code::UnsupportedCapability),
            Self::UnsupportedBiffVersion(_) => Class::of(Code::UnsupportedFormat),
            Self::StringTooLong { .. } | Self::TooMany { .. } | Self::RecordTooLong { .. } => {
                Class::of(Code::LimitExceeded)
            },
            Self::WorksheetNotFound(_) => Class::of(Code::NotFound),
            Self::Allocation(_) => Class::of(Code::InternalError),
            Self::InvalidRecord { .. }
            | Self::MalformedFilePass(_)
            | Self::InvalidLength { .. }
            | Self::UnexpectedEndOfStream(_)
            | Self::InvalidFormula(_)
            | Self::InvalidCellReference(_)
            | Self::InvalidFormat(_)
            | Self::Encoding(_)
            | Self::InvalidData(_)
            | Self::UnexpectedRecordType { .. }
            | Self::Eof(_) => Class::of(Code::CorruptFile),
            _ => Class::of(Code::InternalError),
        }
    }
}

#[cfg(feature = "xls")]
impl Classify for litchi_xls::package::edit::Error {
    fn classify(&self) -> Class {
        match self {
            Self::Package(error) => error.classify(),
            Self::CellValues(error) => error.classify(),
            Self::Limit { .. } => Class::of(Code::LimitExceeded),
            Self::PatchConflict => Class::of(Code::InvalidInput),
            _ => Class::of(Code::InternalError),
        }
    }
}

#[cfg(feature = "xls")]
impl From<litchi_xls::package::edit::Error> for Error {
    fn from(error: litchi_xls::package::edit::Error) -> Self {
        let class = error.classify();
        // The leaf names the exact refusal (the dependency record, the
        // reader check); the code summary alone would hide it (record 0293).
        let message = error.to_string();
        Self::from_class(
            class,
            message,
            Some(FileFormat::Xls),
            Source::XlsEdit(error),
        )
    }
}

#[cfg(feature = "xls")]
impl Classify for litchi_xls::package::Error {
    fn classify(&self) -> Class {
        match self {
            Self::Reader(error) => error.classify(),
            Self::Source(error) => error.classify(),
            Self::CachedSource(error) => error.as_ref().classify(),
            Self::Cfb(error) => error.classify(),
            Self::Object(error) => error.classify(),
            Self::Io(error) => error.classify(),
            Self::Allocation { .. } => Class::of(Code::InternalError),
            Self::EncryptionPolicy => Class::suggest(
                Code::DocumentProtected,
                "export the original encrypted source explicitly",
            ),
            Self::InvalidData(_) => Class::of(Code::CorruptFile),
            _ => Class::of(Code::InternalError),
        }
    }
}

#[cfg(feature = "xls")]
impl From<litchi_xls::package::Error> for Error {
    fn from(error: litchi_xls::package::Error) -> Self {
        use litchi_xls::package::Error as PackageError;
        match error {
            PackageError::Reader(error) => error.into(),
            PackageError::Source(error) => error.into(),
            PackageError::Cfb(error) => litchi_xls::Error::Cfb(error).into(),
            PackageError::Io(error) => litchi_xls::Error::Io(error).into(),
            error => Self::from_class(
                error.classify(),
                error.to_string(),
                Some(FileFormat::Xls),
                Source::XlsPackage(error),
            ),
        }
    }
}

#[cfg(feature = "xls")]
impl Classify for litchi_xls::SourceBackedError {
    fn classify(&self) -> Class {
        match self {
            Self::Io(error) => error.classify(),
            Self::Cfb(error) => error.classify(),
            Self::Object(error) => error.classify(),
            Self::Parse(error) => error.classify(),
            Self::Execution(error) => error.classify(),
            Self::SourceChanged { .. } => Class::of(Code::IoError),
            Self::ResourceLimit { .. } => Class::of(Code::LimitExceeded),
            Self::EncryptedUnsupported => {
                Class::suggest(Code::DocumentProtected, PASSWORD_SUGGESTION)
            },
            Self::UnsupportedBiffVersion(_) => Class::of(Code::UnsupportedFormat),
            Self::WorksheetNotFound(_) => Class::of(Code::NotFound),
            Self::WorkbookStreamMissing | Self::InvalidData(_) => Class::of(Code::CorruptFile),
            Self::Allocation { .. } => Class::of(Code::InternalError),
        }
    }
}

#[cfg(feature = "ppt")]
impl Classify for litchi_ppt::package::edit::Error {
    fn classify(&self) -> Class {
        match self {
            Self::Package(error) => error.classify(),
            Self::Native(error) => match error {
                litchi_ppt::slide_order::Error::Package(error) => error.classify(),
                litchi_ppt::slide_order::Error::Text(error) => error.classify(),
                litchi_ppt::slide_order::Error::Composition(error) => error.classify(),
                // By kind alone (record 0351 §2.4); a protected refusal is a
                // signed source, which keeps the signature suggestion.
                litchi_ppt::slide_order::Error::Refused(refusal) => match refusal.kind() {
                    litchi_core::edit::RefusalKind::Protected => {
                        Class::suggest(Code::DocumentProtected, SIGNED_SUGGESTION)
                    },
                    kind => kind.classify(),
                },
                _ => Class::of(Code::InternalError),
            },
            Self::Signed => Class::suggest(Code::DocumentProtected, SIGNED_SUGGESTION),
            Self::InvalidLimit { .. } => Class::of(Code::InvalidArgument),
            Self::Limit { .. } => Class::of(Code::LimitExceeded),
            Self::PatchConflict => Class::of(Code::InvalidInput),
            Self::Durable(error) => error.classify(),
            _ => Class::of(Code::InternalError),
        }
    }
}

#[cfg(feature = "ppt")]
impl From<litchi_ppt::package::edit::Error> for Error {
    fn from(error: litchi_ppt::package::edit::Error) -> Self {
        let class = error.classify();
        let message = class.code.summary();
        Self::from_class(
            class,
            message,
            Some(FileFormat::Ppt),
            Source::PptEdit(error),
        )
    }
}

#[cfg(feature = "ppt")]
impl Classify for litchi_ppt::writer::WriteError {
    fn classify(&self) -> Class {
        #[allow(
            unreachable_patterns,
            reason = "additional writer variants depend on PPT leaf feature unification"
        )]
        match self {
            Self::Io(error) => error.classify(),
            Self::Ole(error) => error.classify(),
            Self::Graph(error) => error.classify(),
            Self::InvalidData(_) => Class::of(Code::InvalidArgument),
            #[cfg(feature = "encryption")]
            Self::EncryptionEntropyRequired => Class::of(Code::InvalidArgument),
            Self::Metadata(error) => error.classify(),
            Self::EncryptedMetadata => {
                Class::suggest(Code::UnsupportedCapability, ENCRYPTED_METADATA_SUGGESTION)
            },
            _ => Class::of(Code::UnsupportedCapability),
        }
    }
}

#[cfg(feature = "ppt")]
impl Classify for litchi_ppt::Error {
    fn classify(&self) -> Class {
        match self {
            Self::Authoring(error) => error.classify(),
            Self::Io(error) => error.classify(),
            Self::Ole(error) => error.classify(),
            Self::Object(error) => error.classify(),
            Self::CustomXml(error) => error.classify(),
            #[cfg(feature = "sign")]
            Self::Signature(error) => error.classify(),
            Self::Graph(error) => error.classify(),
            Self::Limit { .. } | Self::ResourceLimit(_) => Class::of(Code::LimitExceeded),
            Self::InvalidLimit { .. } => Class::of(Code::InvalidArgument),
            Self::AllocationFailed(_) => Class::of(Code::InternalError),
            Self::PasswordRequired => Class::suggest(Code::DocumentProtected, PASSWORD_SUGGESTION),
            Self::EncryptionPolicy => Class::suggest(
                Code::DocumentProtected,
                "export the original encrypted source explicitly",
            ),
            Self::InvalidPassword => {
                Class::suggest(Code::InvalidArgument, WRONG_PASSWORD_SUGGESTION)
            },
            Self::UnsupportedEncryption(_) => Class::of(Code::UnsupportedCapability),
            Self::OfficeArt(error) => classify_ppt_officeart(error),
            Self::InvalidFormat(_)
            | Self::StreamNotFound(_)
            | Self::Corrupted(_)
            | Self::MalformedEncryptionHeader(_) => Class::of(Code::CorruptFile),
            _ => Class::of(Code::InternalError),
        }
    }
}

#[cfg(feature = "ppt")]
fn classify_ppt_officeart(error: &litchi_ppt::odraw::Error) -> Class {
    use litchi_ppt::odraw::Error as OfficeArtError;

    match error {
        OfficeArtError::LimitExceeded { .. } | OfficeArtError::ImageLimitExceeded { .. } => {
            Class::of(Code::LimitExceeded)
        },
        OfficeArtError::InvalidLimit { .. } => Class::of(Code::InvalidArgument),
        _ => Class::of(Code::CorruptFile),
    }
}

#[cfg(any(feature = "odt", feature = "ods", feature = "odp", odf_objects))]
impl Classify for litchi_odf_common::Error {
    fn classify(&self) -> Class {
        match self {
            Self::Limit { .. } => Class::of(Code::LimitExceeded),
            Self::InvalidLimit { .. } => Class::of(Code::InvalidArgument),
            Self::Refused(refusal) => refusal.classify(),
            Self::PatchConflict { .. } => Class::of(Code::InvalidInput),
            Self::Signed => Class::of(Code::DocumentProtected),
            Self::Core(error) => error.classify(),
            Self::SourceContentPublication(error) => classify_source_publication(error),
            Self::SourceContentInsertion(error) => classify_source_insertion(error),
            _ => Class::of(Code::InternalError),
        }
    }
}

#[cfg(any(feature = "odt", feature = "ods", feature = "odp", odf_objects))]
fn classify_source_publication(error: &litchi_odf_common::SourceContentPublicationError) -> Class {
    use litchi_odf_common::SourceContentPublicationError as Publication;

    match error {
        Publication::Core(error) => error.classify(),
        Publication::Allocation { .. } => Class::of(Code::InternalError),
        Publication::Unsupported { .. } => Class::of(Code::UnsupportedCapability),
        Publication::Cancelled { .. } => Class::of(Code::Cancelled),
        Publication::LimitExceeded { .. } => Class::of(Code::LimitExceeded),
        Publication::SourceChanged { .. } => Class::of(Code::IoError),
        Publication::Source { source, .. } => source.classify(),
        Publication::Execution { source, .. } => source.classify(),
        Publication::Sink { source, .. } => source.classify(),
        Publication::Archive { source, .. } => classify_zip_error(source),
        _ => Class::of(Code::InternalError),
    }
}

#[cfg(any(feature = "odt", feature = "ods", feature = "odp", odf_objects))]
fn classify_source_insertion(error: &litchi_odf_common::SourceContentInsertionError) -> Class {
    use litchi_odf_common::SourceContentInsertionError as Insertion;

    match error {
        Insertion::Publication { source, .. } => classify_source_publication(source),
        Insertion::Replay { source, .. } => classify_replay(source),
        Insertion::Transport {
            publication,
            replay,
            ..
        } => {
            let publication_class = classify_source_publication(publication);
            if publication_source_changed(publication) {
                publication_class
            } else {
                let replay_class = classify_replay(replay);
                if replay_source_changed(replay) {
                    replay_class
                } else if publication_class.code == Code::IoError {
                    publication_class
                } else if replay_class.code == Code::IoError {
                    replay_class
                } else {
                    publication_class
                }
            }
        },
        _ => Class::of(Code::InternalError),
    }
}

#[cfg(any(feature = "odt", feature = "ods", feature = "odp", odf_objects))]
fn classify_zip_error(error: &soapberry_zip::Error) -> Class {
    use soapberry_zip::ErrorKind;

    match error.kind() {
        ErrorKind::IO(_) | ErrorKind::CentralDirectorySpool { .. } => Class::of(Code::IoError),
        ErrorKind::FileNotFound(_) => Class::of(Code::FileNotFound),
        ErrorKind::LimitExceeded { .. } | ErrorKind::CentralDirectorySpoolLimitExceeded { .. } => {
            Class::of(Code::LimitExceeded)
        },
        ErrorKind::Cancelled => Class::of(Code::Cancelled),
        ErrorKind::UnsupportedPreservation { .. } => Class::of(Code::UnsupportedCapability),
        _ if error.is_content_fault() => Class::of(Code::CorruptFile),
        _ => Class::of(Code::InternalError),
    }
}

#[cfg(any(feature = "odt", feature = "ods", feature = "odp", odf_objects))]
fn classify_replay(
    error: &soapberry_zip::ReplayPublicationError<
        litchi_odf_common::SourceContentInsertionCallbackError,
    >,
) -> Class {
    use soapberry_zip::ReplayPublicationError;

    match error {
        ReplayPublicationError::Archive {
            source,
            callback_error,
            ..
        } => {
            if let Some(callback) = callback_error
                .as_ref()
                .filter(|callback| callback_source_changed(callback))
            {
                classify_callback(callback)
            } else {
                classify_zip_error(source)
            }
        },
        ReplayPublicationError::Callback { source, .. } => classify_callback(source),
        ReplayPublicationError::Limit { .. } => Class::of(Code::LimitExceeded),
        ReplayPublicationError::Sink {
            source,
            callback_error,
            ..
        } => {
            if let Some(callback) = callback_error
                .as_ref()
                .filter(|callback| callback_source_changed(callback))
            {
                classify_callback(callback)
            } else {
                source.classify()
            }
        },
        ReplayPublicationError::NonDeterministic { .. } => Class::of(Code::InternalError),
        _ => Class::of(Code::InternalError),
    }
}

#[cfg(any(feature = "odt", feature = "ods", feature = "odp", odf_objects))]
fn publication_source_changed(error: &litchi_odf_common::SourceContentPublicationError) -> bool {
    use litchi_odf_common::SourceContentPublicationError as Publication;

    matches!(
        error,
        Publication::SourceChanged { .. }
            | Publication::Core(litchi_core::Error::SourceChanged { .. })
            | Publication::Source {
                source: litchi_core::Error::SourceChanged { .. },
                ..
            }
    )
}

#[cfg(any(feature = "odt", feature = "ods", feature = "odp", odf_objects))]
fn replay_source_changed(
    error: &soapberry_zip::ReplayPublicationError<
        litchi_odf_common::SourceContentInsertionCallbackError,
    >,
) -> bool {
    use soapberry_zip::ReplayPublicationError;

    match error {
        ReplayPublicationError::Archive { callback_error, .. }
        | ReplayPublicationError::Sink { callback_error, .. } => {
            callback_error.as_ref().is_some_and(callback_source_changed)
        },
        ReplayPublicationError::Callback { source, .. } => callback_source_changed(source),
        ReplayPublicationError::Limit { .. } | ReplayPublicationError::NonDeterministic { .. } => {
            false
        },
        _ => false,
    }
}

#[cfg(any(feature = "odt", feature = "ods", feature = "odp", odf_objects))]
fn callback_source_changed(error: &litchi_odf_common::SourceContentInsertionCallbackError) -> bool {
    use litchi_odf_common::SourceMemberReaderError;

    match error {
        litchi_odf_common::SourceContentInsertionCallbackError::Member(error) => {
            match error.as_ref() {
                SourceMemberReaderError::Core { error, .. } => {
                    matches!(error, litchi_core::Error::SourceChanged { .. })
                },
                SourceMemberReaderError::Callback(error) => callback_source_changed(error),
                _ => false,
            }
        },
        litchi_odf_common::SourceContentInsertionCallbackError::Io(_)
        | litchi_odf_common::SourceContentInsertionCallbackError::Cancelled
        | litchi_odf_common::SourceContentInsertionCallbackError::Execution(_)
        | litchi_odf_common::SourceContentInsertionCallbackError::Verification { .. } => false,
        _ => false,
    }
}

#[cfg(any(feature = "odt", feature = "ods", feature = "odp", odf_objects))]
fn classify_callback(error: &litchi_odf_common::SourceContentInsertionCallbackError) -> Class {
    use litchi_odf_common::{SourceContentInsertionCallbackError, SourceMemberReaderError};

    match error {
        SourceContentInsertionCallbackError::Member(error) => match error.as_ref() {
            SourceMemberReaderError::Core { error, .. } => error.classify(),
            SourceMemberReaderError::Callback(error) => classify_callback(error),
            _ => Class::of(Code::InternalError),
        },
        SourceContentInsertionCallbackError::Io(error) => error.classify(),
        SourceContentInsertionCallbackError::Cancelled => Class::of(Code::Cancelled),
        SourceContentInsertionCallbackError::Execution(error) => error.classify(),
        SourceContentInsertionCallbackError::Verification { .. } => Class::of(Code::InternalError),
        _ => Class::of(Code::InternalError),
    }
}

#[cfg(feature = "odp")]
impl Classify for litchi_odp::tail_append::Error {
    fn classify(&self) -> Class {
        use litchi_odf_common::SourceMemberReaderError;

        match self {
            Self::Invalid(error) => error.classify(),
            Self::SourceMember(error) => match error {
                SourceMemberReaderError::Core { error, .. }
                | SourceMemberReaderError::Callback(error) => error.classify(),
                _ => Class::of(Code::InternalError),
            },
            Self::Plan(error) => classify_source_publication(error),
            Self::Publication(error) => classify_source_insertion(error),
            Self::SourceChanged { .. } => Class::of(Code::IoError),
            Self::PatchConflict { .. } => Class::of(Code::InvalidInput),
            _ => Class::of(Code::InternalError),
        }
    }
}

#[cfg(feature = "rtf")]
impl Classify for litchi_rtf::builder::Error {
    fn classify(&self) -> Class {
        match self {
            Self::Source(error) => error.classify(),
            Self::Io(error) => error.classify(),
            _ => Class::of(Code::InternalError),
        }
    }
}

#[cfg(feature = "rtf")]
impl Classify for litchi_rtf::Error {
    fn classify(&self) -> Class {
        match self {
            Self::Limit { .. } | Self::LimitExceeded { .. } => Class::of(Code::LimitExceeded),
            Self::InvalidLimit { .. } => Class::of(Code::InvalidArgument),
            Self::AllocationFailed { .. } => Class::of(Code::InternalError),
            Self::UnsupportedCodePage(_) | Self::UnsupportedContent(_) => {
                Class::of(Code::UnsupportedCapability)
            },
            _ => Class::of(Code::CorruptFile),
        }
    }
}

#[cfg(feature = "rtf")]
impl Classify for litchi_rtf::edit::Error {
    fn classify(&self) -> Class {
        match self {
            Self::Refused(refusal) => refusal.kind().classify(),
            Self::Limit { .. } | Self::InputTooLarge { .. } => Class::of(Code::LimitExceeded),
            Self::InvalidLimit { .. } => Class::of(Code::InvalidArgument),
            Self::EmptyParagraphBatch
            | Self::ParagraphBatchOutOfOrder { .. }
            | Self::EmptyParagraphRemovalBatch
            | Self::ParagraphRemovalBatchOutOfOrder { .. }
            | Self::EmptyParagraphAlignmentBatch
            | Self::ParagraphAlignmentBatchOutOfOrder { .. }
            | Self::EmptyParagraphLayoutBatch
            | Self::EmptyParagraphLayoutPatch { .. }
            | Self::ParagraphLayoutBatchOutOfOrder { .. }
            | Self::EmptyRunInsertionBatch
            | Self::RunInsertionBatchOutOfOrder { .. }
            | Self::EmptyRunTextBatch
            | Self::RunTextBatchOutOfOrder { .. }
            | Self::EmptyRunFormatBatch
            | Self::EmptyRunFormatPatch { .. }
            | Self::RunFormatBatchOutOfOrder { .. }
            | Self::Conflict { .. }
            | Self::ReversedSpan { .. }
            | Self::ParagraphSplitOffsetOutOfRange { .. }
            | Self::ParagraphSplitAtEndRequiresBoundary { .. }
            | Self::ParagraphMergeNonAdjacent { .. }
            | Self::PicturePayloadSizeMismatch { .. }
            | Self::EmptyPicturePayloadBatch
            | Self::PicturePayloadBatchOutOfOrder { .. }
            | Self::EmptyPictureRemovalBatch
            | Self::PictureRemovalBatchOutOfOrder { .. }
            | Self::DurablePatch(_) => Class::of(Code::InvalidArgument),
            Self::EmptyRunTextInsertion { .. } => Class::of(Code::InvalidValue),
            Self::SpanOutOfRange { .. }
            | Self::SpanNotOnCharacterBoundary { .. }
            | Self::ParagraphOutOfRange { .. }
            | Self::RunInsertionOutOfRange { .. }
            | Self::RunOutOfRange { .. }
            | Self::DestinationOutOfRange(_)
            | Self::PictureOutOfRange { .. } => Class::of(Code::OutOfRange),
            Self::StalePrecondition(_) => Class::of(Code::InvalidArgument),
            Self::HistorySourceMismatch | Self::PatchConflict => Class::of(Code::InvalidInput),
            Self::History(_) => Class::of(Code::InvalidInput),
            Self::Rtf(error) => error.classify(),
            Self::Write(_) | Self::Internal(_) => Class::of(Code::InternalError),
            _ => Class::of(Code::CorruptFile),
        }
    }
}

#[cfg(feature = "encryption")]
impl Classify for litchi_crypto::ooxml::Error {
    fn classify(&self) -> Class {
        match self {
            Self::Io(error) => error.classify(),
            Self::Password => Class::suggest(Code::InvalidArgument, WRONG_PASSWORD_SUGGESTION),
            Self::Limit { .. } => Class::of(Code::LimitExceeded),
            Self::InvalidLimit(_) | Self::PasswordRequired => Class::of(Code::InvalidArgument),
            Self::Unsupported(_) => Class::of(Code::UnsupportedCapability),
            Self::Allocation(_) | Self::Random(_) => Class::of(Code::InternalError),
            _ => Class::of(Code::CorruptFile),
        }
    }
}

#[cfg(feature = "formula")]
impl Classify for litchi_formula::FormulaError {
    fn classify(&self) -> Class {
        use litchi_formula::{MtefError, MtefWriteError, OmmlError};

        match self {
            Self::Omml(OmmlError::UnsupportedFeature(_)) => Class::of(Code::UnsupportedElement),
            Self::Omml(OmmlError::LimitExceeded { .. }) => Class::of(Code::LimitExceeded),
            Self::Mtef(MtefError::LimitExceeded { .. })
            | Self::MtefWrite(MtefWriteError::LimitExceeded { .. }) => {
                Class::of(Code::LimitExceeded)
            },
            Self::Omml(OmmlError::ArenaAllocationError(_)) => Class::of(Code::InternalError),
            _ => Class::of(Code::CorruptFile),
        }
    }
}

#[cfg(any(feature = "doc", feature = "ppt"))]
impl Classify for litchi_cfb::OverlayError {
    fn classify(&self) -> Class {
        match self {
            Self::Ole(error) => error.classify(),
            Self::Io(error) => error.classify(),
            Self::Unavailable { .. } => Class::of(Code::UnsupportedCapability),
            // The source no longer holds the bytes the edit was planned on.
            Self::SourceChanged { .. }
            | Self::SourceFingerprintChanged { .. }
            | Self::PreconditionFailed { .. }
            | Self::IncompleteOutput { .. } => Class::of(Code::IoError),
            Self::TargetFingerprintChanged { .. }
            | Self::Allocation { .. }
            | Self::Committed { .. } => Class::of(Code::InternalError),
            _ => Class::of(Code::CorruptFile),
        }
    }
}

#[cfg(feature = "doc")]
impl Classify for litchi_doc::source_backed::Error {
    fn classify(&self) -> Class {
        use litchi_doc::source_backed::Refusal;

        match self {
            Self::Ole(error) => error.classify(),
            Self::Overlay(error) => error.classify(),
            Self::Io(error) => error.classify(),
            Self::Refused(Refusal::Signed) => {
                Class::suggest(Code::DocumentProtected, SIGNED_SUGGESTION)
            },
            Self::Refused(Refusal::Encrypted) => {
                Class::suggest(Code::DocumentProtected, PASSWORD_SUGGESTION)
            },
            Self::Refused(Refusal::Protected | Refusal::UnknownProtection) => Class::suggest(
                Code::DocumentProtected,
                "authorize the protected edit explicitly",
            ),
            // A missing position is `out_of_range`, as in the body editor.
            Self::Refused(Refusal::ParagraphNotFound) => Class::of(Code::OutOfRange),
            // Every other refusal by its kind alone (record 0351 §2.4).
            Self::Refused(refusal) => refusal.kind().classify(),
            Self::InvalidData(_) => Class::of(Code::CorruptFile),
            Self::Limit(_) => Class::of(Code::LimitExceeded),
            Self::Allocation { .. } => Class::of(Code::InternalError),
            _ => Class::of(Code::CorruptFile),
        }
    }
}

#[cfg(feature = "ppt")]
impl Classify for litchi_ppt::text_edit::Error {
    fn classify(&self) -> Class {
        match self {
            Self::Package(error) => error.classify(),
            Self::Source(error) => error.classify(),
            // By kind alone (record 0351 §2.4). `UnsupportedSource`, the
            // leaf's one refusal for signed, encrypted and otherwise
            // unsupported containers, is unsupported content: a signed
            // handle is refused before it reaches the leaf.
            Self::Refused(refusal) => refusal.kind().classify(),
            Self::PatchConflict => Class::of(Code::InvalidInput),
            Self::Signed => Class::suggest(Code::DocumentProtected, SIGNED_SUGGESTION),
            _ => Class::of(Code::CorruptFile),
        }
    }
}

#[cfg(any(feature = "doc", feature = "ppt", feature = "xls"))]
impl Classify for litchi_sign::Error {
    fn classify(&self) -> Class {
        match self {
            Self::Cfb(error) => error.classify(),
            Self::Limit(_) => Class::of(Code::LimitExceeded),
            Self::Xml(_) | Self::Container(_) => Class::of(Code::CorruptFile),
            // Encrypted or rights-managed storage, an unsupported legacy
            // signature form, or an algorithm the policy refuses.
            Self::Encrypted | Self::Legacy | Self::Unsupported(_) | Self::Sha1 => {
                Class::of(Code::UnsupportedCapability)
            },
            Self::Key(_) => Class::of(Code::InvalidArgument),
            Self::Sign(_) => Class::of(Code::InternalError),
            _ => Class::of(Code::CorruptFile),
        }
    }
}

#[cfg(any(feature = "doc", feature = "ppt", feature = "xls"))]
impl Classify for litchi_ole_common::custom_xml::Error {
    fn classify(&self) -> Class {
        match self {
            Self::Ole(error) => error.classify(),
            Self::ResourceLimit(_) => Class::of(Code::LimitExceeded),
            Self::Invalid(_) | Self::Xml(_) => Class::of(Code::CorruptFile),
            _ => Class::of(Code::CorruptFile),
        }
    }
}

#[cfg(any(feature = "doc", feature = "xls"))]
impl Classify for litchi_vba::Error {
    fn classify(&self) -> Class {
        match self {
            Self::Cfb(error) => error.classify(),
            Self::Limit { .. } => Class::of(Code::LimitExceeded),
            Self::UnsupportedCodePage(_) => Class::of(Code::UnsupportedCapability),
            Self::InvalidData(_) => Class::of(Code::CorruptFile),
            _ => Class::of(Code::CorruptFile),
        }
    }
}

#[cfg(feature = "doc")]
impl Classify for litchi_crypto::spaces::Error {
    fn classify(&self) -> Class {
        match self {
            Self::PatchConflict => Class::of(Code::InvalidInput),
            Self::Ole(error) => error.classify(),
            Self::InvalidLimit { .. } => Class::of(Code::InvalidArgument),
            Self::LimitExceeded { .. } => Class::of(Code::LimitExceeded),
            Self::Allocation { .. } => Class::of(Code::InternalError),
            Self::Invalid(_) => Class::of(Code::CorruptFile),
            _ => Class::of(Code::CorruptFile),
        }
    }
}

/// Implements `From<leaf>` for [`Error`]: classify, keep the leaf's message,
/// record the leaf's format, and keep the leaf as the typed source.
macro_rules! from_leaf {
    ($(#[$cfg:meta])* $leaf:ty => $variant:ident, $format:expr) => {
        $(#[$cfg])*
        impl From<$leaf> for Error {
            // An error conversion is never the hot path: keep it out of the
            // callers' bodies (record 0353a).
            #[cold]
            #[inline(never)]
            fn from(error: $leaf) -> Self {
                Self::from_class(
                    error.classify(),
                    error.to_string(),
                    $format,
                    Source::$variant(error),
                )
            }
        }
    };
}

from_leaf!(io::Error => Io, None);
from_leaf!(crate::batch::Error => Batch, None);
from_leaf!(crate::import::Error => Import, None);
from_leaf!(crate::merge::Error => Merge, None);
from_leaf!(crate::raw::Error => Raw, None);
from_leaf!(#[cfg(all(feature = "eval", any(feature = "xlsx", feature = "xlsb", feature = "xls")))] crate::calc::Error => Calc, None);
from_leaf!(crate::transaction::Error => Transaction, None);
from_leaf!(litchi_core::Error => Core, None);
from_leaf!(ExecutionError => Execution, None);
from_leaf!(litchi_core::selector::ParseError => Selector, None);
from_leaf!(
    #[cfg(any(
        feature = "docx",
        feature = "pptx",
        feature = "xlsx",
        feature = "xlsb"
    ))]
    litchi_opc::OpcError => Opc, None
);
from_leaf!(
    #[cfg(any(feature = "doc", feature = "xls", feature = "ppt", feature = "encryption"))]
    litchi_cfb::OleError => Cfb, None
);
from_leaf!(#[cfg(feature = "docx")] litchi_docx::Error => Docx, Some(FileFormat::Docx));
from_leaf!(
    #[cfg(feature = "docx")]
    litchi_docx::document::Error => DocxTransaction, Some(FileFormat::Docx)
);
from_leaf!(
    #[cfg(feature = "docx")]
    litchi_docx::metadata::Error => DocxMetadata, Some(FileFormat::Docx)
);
from_leaf!(#[cfg(feature = "pptx")] litchi_pptx::Error => Pptx, Some(FileFormat::Pptx));
from_leaf!(#[cfg(feature = "xlsx")] litchi_xlsx::Error => Xlsx, Some(FileFormat::Xlsx));
// XLSB formula and relationship errors may retain authored values. Keep the
// typed cause, but expose only a content-free classification in diagnostics.
#[cfg(feature = "xlsb")]
impl From<litchi_xlsb::package::PackageError> for Error {
    #[cold]
    #[inline(never)]
    fn from(error: litchi_xlsb::package::PackageError) -> Self {
        let class = error.classify();
        let message = class.code.summary();
        Self::from_class(class, message, Some(FileFormat::Xlsb), Source::Xlsb(error))
    }
}

#[cfg(feature = "xlsb")]
impl From<litchi_xlsb::Error> for Error {
    #[cold]
    #[inline(never)]
    fn from(error: litchi_xlsb::Error) -> Self {
        let class = error.classify();
        let message = class.code.summary();
        Self::from_class(
            class,
            message,
            Some(FileFormat::Xlsb),
            Source::XlsbRaw(error),
        )
    }
}
from_leaf!(#[cfg(feature = "doc")] litchi_doc::Error => Doc, Some(FileFormat::Doc));
#[cfg(feature = "doc")]
impl From<litchi_doc::package::edit::Error> for Error {
    fn from(error: litchi_doc::package::edit::Error) -> Self {
        let class = error.classify();
        let message = class.code.summary();
        Self::from_class(
            class,
            message,
            Some(FileFormat::Doc),
            Source::DocTransaction(error),
        )
    }
}
#[cfg(feature = "doc")]
impl From<litchi_doc::body_text::Error> for Error {
    fn from(error: litchi_doc::body_text::Error) -> Self {
        let class = error.classify();
        let message = class.code.summary();
        Self::from_class(
            class,
            message,
            Some(FileFormat::Doc),
            Source::DocEdit(error),
        )
    }
}
#[cfg(feature = "doc")]
impl From<litchi_doc::writer::WriteError> for Error {
    fn from(error: litchi_doc::writer::WriteError) -> Self {
        let class = error.classify();
        let message = class.code.summary();
        Self::from_class(
            class,
            message,
            Some(FileFormat::Doc),
            Source::DocWrite(error),
        )
    }
}
from_leaf!(
    #[cfg(feature = "doc")]
    litchi_doc::source_backed::Error => DocBodyText, Some(FileFormat::Doc)
);
from_leaf!(#[cfg(feature = "xls")] litchi_xls::Error => Xls, Some(FileFormat::Xls));
from_leaf!(
    #[cfg(feature = "xls")]
    litchi_xls::SourceBackedError => XlsSource, Some(FileFormat::Xls)
);
from_leaf!(#[cfg(feature = "ppt")] litchi_ppt::Error => Ppt, Some(FileFormat::Ppt));
// Legacy OLE metadata (record 0254): the shared property-set conversion and
// the three leaf metadata transaction families.
#[cfg(any(feature = "doc", feature = "xls", feature = "ppt"))]
const ENCRYPTED_METADATA_SUGGESTION: &str = "leave the metadata unset for an encrypted legacy document; its encrypted summary stream is not authored";

#[cfg(any(feature = "doc", feature = "xls", feature = "ppt"))]
impl Classify for litchi_ole_common::property_set::summary_properties::MetadataError {
    fn classify(&self) -> Class {
        match self {
            Self::Unsupported { .. } => Class::suggest(
                Code::UnsupportedCapability,
                "leave identifier, hyperlink_base and unzoned timestamps unset for legacy binary formats",
            ),
            _ => Class::of(Code::InvalidValue),
        }
    }
}

macro_rules! ole_metadata_classify {
    ($(#[$cfg:meta])* $leaf:ty, $protected:ident) => {
        $(#[$cfg])*
        impl Classify for $leaf {
            fn classify(&self) -> Class {
                match self {
                    Self::Package(error) => error.classify(),
                    Self::Metadata(error) => error.classify(),
                    Self::Property(_) => Class::of(Code::InvalidValue),
                    Self::$protected => {
                        Class::suggest(Code::DocumentProtected, ENCRYPTED_METADATA_SUGGESTION)
                    },
                    Self::PatchConflict => Class::of(Code::InvalidInput),
                    _ => Class::of(Code::InternalError),
                }
            }
        }
    };
}

ole_metadata_classify!(
    #[cfg(feature = "doc")]
    litchi_doc::package::metadata::Error,
    Encrypted
);
ole_metadata_classify!(
    #[cfg(feature = "xls")]
    litchi_xls::package::metadata::Error,
    Protected
);
ole_metadata_classify!(
    #[cfg(feature = "ppt")]
    litchi_ppt::package::metadata::Error,
    Encrypted
);
from_leaf!(
    #[cfg(feature = "doc")]
    litchi_doc::package::metadata::Error => DocMetadata, Some(FileFormat::Doc)
);
from_leaf!(
    #[cfg(feature = "xls")]
    litchi_xls::package::metadata::Error => XlsMetadata, Some(FileFormat::Xls)
);
from_leaf!(
    #[cfg(feature = "ppt")]
    litchi_ppt::package::metadata::Error => PptMetadata, Some(FileFormat::Ppt)
);
from_leaf!(
    #[cfg(feature = "ppt")]
    litchi_ppt::text_edit::Error => PptTextEdit, Some(FileFormat::Ppt)
);
from_leaf!(
    #[cfg(any(feature = "doc", feature = "ppt", feature = "xls"))]
    litchi_sign::Error => Sign, None
);
from_leaf!(#[cfg(feature = "rtf")] litchi_rtf::Error => Rtf, Some(FileFormat::Rtf));
from_leaf!(
    #[cfg(any(feature = "odt", feature = "ods", feature = "odp", odf_objects))]
    litchi_odf_common::Error => Odf, None
);
from_leaf!(
    #[cfg(feature = "odp")]
    litchi_odp::tail_append::Error => OdpTailAppend, Some(FileFormat::Odp)
);
from_leaf!(#[cfg(feature = "rtf")] litchi_rtf::builder::Error => RtfBuild, Some(FileFormat::Rtf));
from_leaf!(#[cfg(feature = "rtf")] litchi_rtf::edit::Error => RtfEdit, Some(FileFormat::Rtf));
from_leaf!(
    #[cfg(feature = "encryption")]
    litchi_crypto::ooxml::Error => Crypto, None
);
from_leaf!(#[cfg(feature = "formula")] litchi_formula::FormulaError => Formula, None);
from_leaf!(#[cfg(feature = "iwork-adapter")] crate::format::iwork::LeafError => Iwork, None);

#[cfg(feature = "iwork-adapter")]
impl Classify for crate::format::iwork::LeafError {
    fn classify(&self) -> Class {
        Class::of(self.code())
    }
}

#[cfg(test)]
mod tests;

impl Classify for litchi_core::atomic::Error {
    fn classify(&self) -> Class {
        match self {
            Self::Io(error) => error.classify(),
            Self::Committed { .. } => Class::of(Code::InternalError),
            _ => Class::of(Code::InternalError),
        }
    }
}
from_leaf!(litchi_core::atomic::Error => Publication, None);

/// Classify a typed error for the format-neutral sheet error the formula
/// engine reports: the code selects the variant, and the message is the
/// code's static summary, so no authored content crosses.
#[cfg(any(feature = "xlsx", feature = "xlsb", feature = "xls", feature = "ods"))]
impl From<Error> for litchi_core::sheet::Error {
    #[cold]
    #[inline(never)]
    fn from(error: Error) -> Self {
        use litchi_core::edit::RefusalKind;
        let code = error.code();
        let refused = |kind| Self::refused(kind, "read", code.summary());
        match code {
            Code::NotFound
            | Code::PathNotFound
            | Code::ZeroMatches
            | Code::RangeTargetNotFound
            | Code::FileNotFound => refused(RefusalKind::NotFound),
            Code::OutOfRange
            | Code::InvalidValue
            | Code::InvalidInput
            | Code::InvalidArgument
            | Code::InvalidPath
            | Code::InvalidSelector
            | Code::MissingArgument => refused(RefusalKind::InvalidValue),
            Code::UnsupportedProperty => refused(RefusalKind::UnsupportedProperty),
            Code::UnsupportedType | Code::UnsupportedElement | Code::UnsupportedPath => {
                refused(RefusalKind::UnsupportedTarget)
            },
            Code::UnsupportedFormat | Code::UnsupportedCommand | Code::UnsupportedCapability => {
                refused(RefusalKind::UnsupportedOperation)
            },
            Code::LimitExceeded | Code::OutputTooLarge | Code::MaxDepthExceeded => {
                refused(RefusalKind::UnsupportedContent)
            },
            Code::DocumentProtected => refused(RefusalKind::Protected),
            Code::IoError | Code::FileLocked => Self::Io,
            Code::CorruptFile | Code::CorruptBatch => Self::Malformed(code.summary()),
            _ => Self::Internal(code.summary()),
        }
    }
}
