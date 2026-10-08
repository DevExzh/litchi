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
//!
//! Leaf enums that are `#[non_exhaustive]` end in a wildcard arm that reports
//! [`Code::CorruptFile`]: a variant the classifier does not name is treated as
//! a malformed input until it is classified explicitly.

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
/// Each variant exists only when the cargo feature of its format is enabled.
/// ODT, ODS and ODP report [`litchi_core::Error`] and arrive as
/// [`Source::Core`].
#[derive(Debug)]
#[non_exhaustive]
pub enum Source {
    /// The error was raised by the unified layer itself.
    None,
    /// An I/O error.
    Io(io::Error),
    /// A shared Litchi core error.
    Core(litchi_core::Error),
    /// An execution-context refusal: cancellation or a budget charge.
    Execution(ExecutionError),
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
    /// A source-backed legacy Word main-story edit error.
    #[cfg(feature = "doc")]
    DocBodyText(litchi_doc::body_text::source::Error),
    /// A legacy Excel BIFF error.
    #[cfg(feature = "xls")]
    Xls(litchi_xls::Error),
    /// A source-backed legacy Excel error.
    #[cfg(feature = "xls")]
    XlsSource(litchi_xls::SourceBackedError),
    /// A legacy PowerPoint binary error.
    #[cfg(feature = "ppt")]
    Ppt(litchi_ppt::Error),
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
    /// An OOXML package-encryption error.
    #[cfg(feature = "encryption")]
    Crypto(litchi_crypto::ooxml::Error),
    /// A formula conversion error: Markdown output renders OMML formulas as
    /// LaTeX.
    #[cfg(feature = "formula")]
    Formula(litchi_formula::FormulaError),
}

impl Source {
    /// The leaf error as a standard error, or `None` for [`Source::None`].
    fn as_error(&self) -> Option<&(dyn std::error::Error + 'static)> {
        match self {
            Self::None => None,
            Self::Io(error) => Some(error),
            Self::Core(error) => Some(error),
            Self::Execution(error) => Some(error),
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
            Self::DocBodyText(error) => Some(error),
            #[cfg(feature = "xls")]
            Self::Xls(error) => Some(error),
            #[cfg(feature = "xls")]
            Self::XlsSource(error) => Some(error),
            #[cfg(feature = "ppt")]
            Self::Ppt(error) => Some(error),
            #[cfg(feature = "ppt")]
            Self::PptTextEdit(error) => Some(error),
            #[cfg(any(feature = "doc", feature = "ppt", feature = "xls"))]
            Self::Sign(error) => Some(error),
            #[cfg(feature = "rtf")]
            Self::Rtf(error) => Some(error),
            #[cfg(feature = "encryption")]
            Self::Crypto(error) => Some(error),
            #[cfg(feature = "formula")]
            Self::Formula(error) => Some(error),
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
/// format and the suggestion when there are any. The leaf error's own text is
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
    /// Shared so that cloning an error never copies or loses its leaf error.
    source: Arc<Source>,
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
            source: Arc::new(Source::None),
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

    /// Reports an iWork input (Pages, Keynote or Numbers) given to the unified
    /// opener, which does not handle iWork in this release.
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
        feature = "numbers"
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
    /// input that does not parse, [`Code::OutOfRange`] for a position outside
    /// its collection. A leaf crate reports such input with the same error it
    /// reports for malformed file content, so only the call site that
    /// received the argument can tell the two apart.
    #[cfg(feature = "xlsx")]
    pub(crate) fn caused_by_argument(mut self, code: Code) -> Self {
        debug_assert!(matches!(code, Code::InvalidArgument | Code::OutOfRange));
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
            source: Arc::new(source),
        }))
    }
}

impl fmt::Debug for Error {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("Error")
            .field("code", &self.0.code)
            .field("message", &self.0.message)
            .field("suggestion", &self.0.suggestion)
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
        if let Some(suggestion) = &self.0.suggestion {
            write!(formatter, " (suggestion: {suggestion})")?;
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
    "open Pages, Keynote and Numbers files with `litchi::iwork::Document::open`";

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

impl Classify for litchi_core::Error {
    fn classify(&self) -> Class {
        match self {
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
            Self::Allocation { .. } | Self::Committed(_) | Self::Other(_) => {
                Class::of(Code::InternalError)
            },
            _ => Class::of(Code::CorruptFile),
        }
    }
}

#[cfg(any(feature = "docx", feature = "pptx", feature = "xlsx", feature = "xlsb"))]
impl Classify for litchi_opc::OpcError {
    fn classify(&self) -> Class {
        match self {
            Self::IoError(error) => error.classify(),
            Self::PackageNotFound(_) => Class::of(Code::FileNotFound),
            Self::SourceChanged { .. }
            | Self::CentralDirectorySpool { .. }
            | Self::IncompleteOutput { .. } => Class::of(Code::IoError),
            Self::ReadLimit { .. }
            | Self::SourcePartSpliceLimit { .. }
            | Self::CentralDirectorySpoolLimitExceeded { .. } => Class::of(Code::LimitExceeded),
            Self::InvalidReadLimit { .. }
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
            Self::Opc(error) => error.classify(),
            Self::Io(error) => error.classify(),
            Self::InkLimit { .. }
            | Self::RevisionLimit { .. }
            | Self::ExternalHyperlinkDetachmentLimit { .. }
            | Self::ExternalHyperlinkRedactionLimit { .. }
            | Self::SectionInventoryLimit { .. } => Class::of(Code::LimitExceeded),
            Self::OutOfBounds { .. } => Class::of(Code::OutOfRange),
            Self::UnsafeEdit { .. } => Class::of(Code::UnsupportedCapability),
            Self::Allocation { .. } | Self::RevisionAllocation { .. } | Self::Other(_) => {
                Class::of(Code::InternalError)
            },
            _ => Class::of(Code::CorruptFile),
        }
    }
}

#[cfg(feature = "docx")]
impl Classify for litchi_docx::document::Error {
    fn classify(&self) -> Class {
        match self {
            Self::Document(error) => error.classify(),
            Self::OutOfBounds { .. } => Class::of(Code::OutOfRange),
            Self::Refused { .. } | Self::Transfer(_) => Class::of(Code::UnsupportedCapability),
            Self::Limit { .. } => Class::of(Code::LimitExceeded),
            // The patch target no longer holds the bytes the edit captured:
            // the source changed underneath the handle.
            Self::StaleSource => Class::of(Code::IoError),
            Self::SemanticPrecondition => Class::of(Code::InvalidArgument),
            Self::Durable(_) | Self::Composition(_) | Self::InvalidDurable(_) => {
                Class::of(Code::InvalidInput)
            },
            _ => Class::of(Code::CorruptFile),
        }
    }
}

#[cfg(feature = "pptx")]
impl Classify for litchi_pptx::Error {
    fn classify(&self) -> Class {
        match self {
            Self::Opc(error) => error.classify(),
            Self::Io(error) => error.classify(),
            // The core-properties reader reports its OPC failures, among them
            // a source change or cancellation during the read, through the
            // shared OOXML error.
            Self::CustomProperties(litchi_ooxml_common::Error::Opc(error)) => error.classify(),
            Self::Limit { .. } => Class::of(Code::LimitExceeded),
            Self::IndexOutOfBounds { .. }
            | Self::SlideIndexOutOfBounds { .. }
            | Self::FontIndexOutOfBounds { .. } => Class::of(Code::OutOfRange),
            Self::NameNotFound(_) | Self::SlideNameNotFound(_) | Self::FontNotFound(_) => {
                Class::of(Code::NotFound)
            },
            Self::UnsafeEdit { .. } => Class::of(Code::UnsupportedCapability),
            Self::Allocation { .. } => Class::of(Code::InternalError),
            _ => Class::of(Code::CorruptFile),
        }
    }
}

#[cfg(feature = "xlsx")]
impl Classify for litchi_xlsx::Error {
    fn classify(&self) -> Class {
        match self {
            Self::Package(error) => error.classify(),
            Self::ResourceLimit(_) => Class::of(Code::LimitExceeded),
            Self::Signed => Class::suggest(Code::DocumentProtected, SIGNED_SUGGESTION),
            Self::Unsupported { .. } | Self::UnsupportedSelector => {
                Class::of(Code::UnsupportedCapability)
            },
            Self::NotWorksheet { .. } => Class::of(Code::UnsupportedType),
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
            Self::Io(error) => error.classify(),
            Self::Opc(error) => error.classify(),
            Self::Wire(error) => error.classify(),
            Self::Execution(error) => error.classify(),
            Self::LimitExceeded { .. } | Self::ExternalLinkLimitExceeded { .. } => {
                Class::of(Code::LimitExceeded)
            },
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
            Self::Io(error) => error.classify(),
            Self::Ole(error) => error.classify(),
            Self::ResourceLimit(_) => Class::of(Code::LimitExceeded),
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

#[cfg(feature = "xls")]
impl Classify for litchi_xls::Error {
    fn classify(&self) -> Class {
        match self {
            Self::Io(error) => error.classify(),
            Self::Cfb(error) => error.classify(),
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
            Self::UnsupportedEncryption(_)
            | Self::UnsupportedFeature(_)
            | Self::UnsafeEdit(_)
            | Self::SourceBoundWorksheetMutation { .. } => Class::of(Code::UnsupportedCapability),
            Self::UnsupportedBiffVersion(_) => Class::of(Code::UnsupportedFormat),
            Self::StringTooLong { .. } | Self::TooMany { .. } | Self::RecordTooLong { .. } => {
                Class::of(Code::LimitExceeded)
            },
            Self::WorksheetNotFound(_) => Class::of(Code::NotFound),
            Self::Allocation(_) => Class::of(Code::InternalError),
            Self::Vba(_)
            | Self::Graph(_)
            | Self::InvalidRecord { .. }
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
        }
    }
}

#[cfg(feature = "xls")]
impl Classify for litchi_xls::SourceBackedError {
    fn classify(&self) -> Class {
        match self {
            Self::Io(error) => error.classify(),
            Self::Cfb(error) => error.classify(),
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
impl Classify for litchi_ppt::Error {
    fn classify(&self) -> Class {
        match self {
            Self::Io(error) => error.classify(),
            Self::Ole(error) => error.classify(),
            Self::ResourceLimit(_) => Class::of(Code::LimitExceeded),
            Self::AllocationFailed(_) => Class::of(Code::InternalError),
            Self::PasswordRequired => Class::suggest(Code::DocumentProtected, PASSWORD_SUGGESTION),
            Self::InvalidPassword => {
                Class::suggest(Code::InvalidArgument, WRONG_PASSWORD_SUGGESTION)
            },
            Self::UnsupportedEncryption(_) => Class::of(Code::UnsupportedCapability),
            Self::OfficeArt(_)
            | Self::Graph(_)
            | Self::InvalidFormat(_)
            | Self::StreamNotFound(_)
            | Self::Corrupted(_)
            | Self::MalformedEncryptionHeader(_) => Class::of(Code::CorruptFile),
        }
    }
}

#[cfg(feature = "rtf")]
impl Classify for litchi_rtf::Error {
    fn classify(&self) -> Class {
        match self {
            Self::LimitExceeded { .. } => Class::of(Code::LimitExceeded),
            Self::AllocationFailed { .. } => Class::of(Code::InternalError),
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
        use litchi_formula::OmmlError;

        match self {
            Self::Omml(OmmlError::UnsupportedFeature(_)) => Class::of(Code::UnsupportedElement),
            Self::Omml(OmmlError::DepthLimitExceeded(_)) => Class::of(Code::LimitExceeded),
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
impl Classify for litchi_doc::body_text::source::Error {
    fn classify(&self) -> Class {
        use litchi_doc::body_text::source::Refusal;

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
            Self::Refused(Refusal::ParagraphNotFound) => Class::of(Code::OutOfRange),
            Self::Refused(_) => Class::of(Code::UnsupportedCapability),
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
        use litchi_ppt::text_edit::Refusal;

        match self {
            Self::Package(error) => error.classify(),
            Self::Source(error) => error.classify(),
            Self::Refused(Refusal::SlideNotFound { .. } | Refusal::ShapeNotFound) => {
                Class::of(Code::NotFound)
            },
            // Includes `UnsupportedSource`, the leaf's one refusal for
            // signed, encrypted and otherwise unsupported containers: a
            // signed handle is refused before it reaches the leaf.
            Self::Refused(_) => Class::of(Code::UnsupportedCapability),
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

/// Implements `From<leaf>` for [`Error`]: classify, keep the leaf's message,
/// record the leaf's format, and keep the leaf as the typed source.
macro_rules! from_leaf {
    ($(#[$cfg:meta])* $leaf:ty => $variant:ident, $format:expr) => {
        $(#[$cfg])*
        impl From<$leaf> for Error {
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
from_leaf!(litchi_core::Error => Core, None);
from_leaf!(ExecutionError => Execution, None);
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
from_leaf!(#[cfg(feature = "pptx")] litchi_pptx::Error => Pptx, Some(FileFormat::Pptx));
from_leaf!(#[cfg(feature = "xlsx")] litchi_xlsx::Error => Xlsx, Some(FileFormat::Xlsx));
from_leaf!(
    #[cfg(feature = "xlsb")]
    litchi_xlsb::package::PackageError => Xlsb, Some(FileFormat::Xlsb)
);
from_leaf!(#[cfg(feature = "xlsb")] litchi_xlsb::Error => XlsbRaw, Some(FileFormat::Xlsb));
from_leaf!(#[cfg(feature = "doc")] litchi_doc::Error => Doc, Some(FileFormat::Doc));
from_leaf!(
    #[cfg(feature = "doc")]
    litchi_doc::body_text::source::Error => DocBodyText, Some(FileFormat::Doc)
);
from_leaf!(#[cfg(feature = "xls")] litchi_xls::Error => Xls, Some(FileFormat::Xls));
from_leaf!(
    #[cfg(feature = "xls")]
    litchi_xls::SourceBackedError => XlsSource, Some(FileFormat::Xls)
);
from_leaf!(#[cfg(feature = "ppt")] litchi_ppt::Error => Ppt, Some(FileFormat::Ppt));
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
    #[cfg(feature = "encryption")]
    litchi_crypto::ooxml::Error => Crypto, None
);
from_leaf!(#[cfg(feature = "formula")] litchi_formula::FormulaError => Formula, None);

#[cfg(test)]
mod tests;
