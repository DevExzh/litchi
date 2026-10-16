//! Python exceptions: one `LitchiError` base class that carries the unified
//! error's `code`, `message` and `suggestion`, and one subclass per error
//! family.
//!
//! Every family class also derives from the built-in exception Python code
//! already catches for that kind of failure, so `except OSError` and
//! `except ValueError` keep working:
//!
//! | Class | Built-in base | Codes |
//! |---|---|---|
//! | `LitchiError` | `Exception` | every other code: `limit_exceeded`, `internal_error`, `cancelled`, `validation_error`, `output_too_large`, `max_depth_exceeded`, `merge_not_exact`, `refresh_failed`, and codes added later |
//! | `PathError` | `LookupError` | `not_found`, `path_not_found`, `zero_matches`, `out_of_range`, `range_target_not_found`, `invalid_path`, `invalid_selector`, `invalid_xpath`, `unsupported_path`, `bare_selector_rejected` |
//! | `PropertyError` | `ValueError` | `invalid_value`, `unsupported_property`, `missing_property`, `duplicate_name`, `invalid_argument`, `missing_argument`, `invalid_json`, `corrupt_batch` |
//! | `UnsupportedError` | `ValueError` | `unsupported_format`, `unsupported_capability`, `unsupported_element`, `unsupported_type`, `unsupported_command` |
//! | `IoError` | `OSError` | `file_not_found`, `file_locked`, `file_exists`, `file_required`, `io_error` |
//! | `CorruptFileError` | `ValueError` | `corrupt_file`, `invalid_input` |
//! | `ProtectedError` | `ValueError` | `document_protected` |
//!
//! The classes are defined in Python, because a pyo3 class can extend only
//! one native base and the I/O family must extend both `LitchiError` and
//! `OSError`. They are created once per process and shared by every
//! initialization of the module.

use std::borrow::Cow;
use std::ffi::CStr;

use litchi::{Code, FileFormat};
use pyo3::exceptions::PyRuntimeError;
use pyo3::prelude::*;
use pyo3::sync::PyOnceLock;
use pyo3::types::{PyCode, PyCodeInput, PyCodeMethods, PyDict, PyModule, PyType};

/// The Python source of the exception classes. It must stay valid on
/// Python 3.8, the oldest version the `abi3-py38` build supports.
const CLASSES: &CStr = cr#"
class LitchiError(Exception):
    """Base class of the errors litchi_py raises for a file it cannot open, read,
    edit or save.

    Attributes:
        code: The stable, machine-readable error code, for example
            "unsupported_format". It is the code Rust's ``litchi::Code`` reports.
        message: The human-readable message.
        suggestion: How to recover, or None when there is no known way.
        item: The zero-based index of the batch item or edit operation that
            failed, or None.

    ``str()`` gives the message, followed by the suggestion when there is one.
    Each family subclass also derives from the built-in exception that fits it,
    so ``except OSError`` and ``except ValueError`` keep working.
    """

    _default_code = "internal_error"

    def __init__(self, message, code=None, suggestion=None, item=None):
        super().__init__(message)
        self._code = type(self)._default_code if code is None else code
        self._message = message
        self._suggestion = suggestion
        self._item = item

    @property
    def code(self):
        """The stable error code, for example "unsupported_format"."""
        return self._code

    @property
    def message(self):
        """The human-readable message."""
        return self._message

    @property
    def suggestion(self):
        """How to recover, or None when there is no known way."""
        return self._suggestion

    @property
    def item(self):
        """The zero-based index of the batch item or edit operation that
        failed, or None when the error is not about one."""
        return self._item

    def __str__(self):
        if self._suggestion is None:
            return str(self._message)
        return "{} (suggestion: {})".format(self._message, self._suggestion)

    def __reduce__(self):
        return (type(self), (self._message, self._code, self._suggestion, self._item))


class PathError(LitchiError, LookupError):
    """A path or selector is malformed or resolves to nothing (codes such as
    "path_not_found", "invalid_path" and "out_of_range")."""

    _default_code = "path_not_found"


class PropertyError(LitchiError, ValueError):
    """A property, argument or payload value is invalid, missing or not
    supported (codes such as "invalid_value" and "unsupported_property")."""

    _default_code = "invalid_value"


class UnsupportedError(LitchiError, ValueError):
    """The input's format, or the requested operation, is not supported (codes
    "unsupported_format", "unsupported_capability", "unsupported_element",
    "unsupported_type" and "unsupported_command")."""

    _default_code = "unsupported_capability"


class IoError(LitchiError, OSError):
    """Reading or writing the file failed, or the file changed after it was
    opened (codes "file_not_found", "file_locked", "file_exists",
    "file_required" and "io_error")."""

    _default_code = "io_error"


class CorruptFileError(LitchiError, ValueError):
    """The file is malformed (codes "corrupt_file" and "invalid_input")."""

    _default_code = "corrupt_file"


class ProtectedError(LitchiError, ValueError):
    """The document is encrypted, signed or edit-protected, or a save would
    change its protection (code "document_protected")."""

    _default_code = "document_protected"
"#;

/// The exception families, one Python class each. The discriminant is the
/// class's position in [`Class::ALL`].
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub(crate) enum Class {
    Base = 0,
    Path = 1,
    Property = 2,
    Unsupported = 3,
    Io = 4,
    Corrupt = 5,
    Protected = 6,
}

impl Class {
    /// Every class with its Python name, in discriminant order.
    const ALL: [(Self, &'static str); 7] = [
        (Self::Base, "LitchiError"),
        (Self::Path, "PathError"),
        (Self::Property, "PropertyError"),
        (Self::Unsupported, "UnsupportedError"),
        (Self::Io, "IoError"),
        (Self::Corrupt, "CorruptFileError"),
        (Self::Protected, "ProtectedError"),
    ];

    /// The family of a unified error code.
    pub(crate) fn of(code: Code) -> Self {
        match code {
            Code::NotFound
            | Code::PathNotFound
            | Code::ZeroMatches
            | Code::OutOfRange
            | Code::RangeTargetNotFound
            | Code::InvalidPath
            | Code::InvalidSelector
            | Code::InvalidXpath
            | Code::UnsupportedPath
            | Code::BareSelectorRejected => Self::Path,
            Code::InvalidValue
            | Code::UnsupportedProperty
            | Code::MissingProperty
            | Code::DuplicateName
            | Code::InvalidArgument
            | Code::MissingArgument
            | Code::InvalidJson
            | Code::CorruptBatch => Self::Property,
            Code::UnsupportedFormat
            | Code::UnsupportedCapability
            | Code::UnsupportedElement
            | Code::UnsupportedType
            | Code::UnsupportedCommand => Self::Unsupported,
            Code::FileNotFound
            | Code::FileLocked
            | Code::FileExists
            | Code::FileRequired
            | Code::IoError => Self::Io,
            Code::CorruptFile | Code::InvalidInput => Self::Corrupt,
            Code::DocumentProtected => Self::Protected,
            // `limit_exceeded`, `internal_error`, `cancelled`,
            // `validation_error`, `output_too_large`, `max_depth_exceeded`,
            // `merge_not_exact`, `refresh_failed`, and codes added later.
            _ => Self::Base,
        }
    }
}

/// The created classes, in [`Class::ALL`] order.
struct Classes([Py<PyType>; 7]);

static TYPES: PyOnceLock<Classes> = PyOnceLock::new();

fn classes(py: Python<'_>) -> PyResult<&Classes> {
    TYPES.get_or_try_init(py, || {
        let namespace = PyDict::new(py);
        // The classes report `litchi_py` as their module, where they are
        // published, so they print and pickle under that name.
        namespace.set_item("__name__", "litchi_py")?;
        PyCode::compile(py, CLASSES, c"<litchi_py exceptions>", PyCodeInput::File)?
            .run(Some(&namespace), None)?;
        let class = |index: usize| -> PyResult<Py<PyType>> {
            let name = Class::ALL[index].1;
            let value = namespace.get_item(name)?.ok_or_else(|| {
                PyRuntimeError::new_err(format!("litchi_py did not define {name}"))
            })?;
            Ok(value.cast_into::<PyType>()?.unbind())
        };
        Ok(Classes([
            class(0)?,
            class(1)?,
            class(2)?,
            class(3)?,
            class(4)?,
            class(5)?,
            class(6)?,
        ]))
    })
}

/// Publish the exception classes in the module.
pub fn register(module: &Bound<'_, PyModule>) -> PyResult<()> {
    let py = module.py();
    let classes = classes(py)?;
    for (class, name) in Class::ALL {
        module.add(name, classes.0[class as usize].bind(py))?;
    }
    Ok(())
}

/// The opener a failure came from, for suggestions that name the opener a
/// caller should use instead.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub(crate) enum Opener {
    Document,
    Presentation,
    Workbook,
}

impl Opener {
    /// The opener that reads `format`, if any does. ODM and OTH read as
    /// documents through the ODT adapter (record 0282); the OpenDocument
    /// object formats belong to none of the three families.
    fn of(format: FileFormat) -> Option<Self> {
        match format {
            FileFormat::Doc
            | FileFormat::Docx
            | FileFormat::Rtf
            | FileFormat::Odt
            | FileFormat::Oth
            | FileFormat::Odm
            | FileFormat::Pages => Some(Self::Document),
            FileFormat::Ppt | FileFormat::Pptx | FileFormat::Odp | FileFormat::Keynote => {
                Some(Self::Presentation)
            },
            FileFormat::Xls
            | FileFormat::Xlsx
            | FileFormat::Xlsb
            | FileFormat::Ods
            | FileFormat::Numbers => Some(Self::Workbook),
            _ => None,
        }
    }

    fn name(self) -> &'static str {
        match self {
            Self::Document => "Document.open",
            Self::Presentation => "Presentation.open",
            Self::Workbook => "Workbook.open",
        }
    }
}

/// An error on its way to Python: the unified code, its family class, the
/// message and a suggestion a Python caller can act on.
///
/// Plain Rust data, so it can be built while the GIL is released.
#[derive(Clone, Debug)]
pub(crate) struct Failure {
    class: Class,
    code: Code,
    message: Cow<'static, str>,
    suggestion: Option<Cow<'static, str>>,
    item: Option<usize>,
}

/// Python callers cannot unlock a document yet: the bindings expose no
/// password entry point.
const PROTECTED_SUGGESTION: &str = "remove the password in the application that created the file; litchi_py cannot unlock encrypted documents yet";

/// The Python form of Rust's suggestion for an edit of a signed document.
const SIGNED_SUGGESTION: &str = "remove the signature explicitly with File.unsign() before editing; the first write then publishes the document without it";

impl Failure {
    /// A failure with `code` and `message` and no suggestion.
    pub(crate) fn new(code: Code, message: impl Into<Cow<'static, str>>) -> Self {
        Self {
            class: Class::of(code),
            code,
            message: message.into(),
            suggestion: None,
            item: None,
        }
    }

    /// A read failure of an opened file.
    pub(crate) fn read(error: &litchi::Error) -> Self {
        let mut failure = Self::new(error.code(), error.message().to_owned());
        failure.item = error.item();
        failure.suggestion = match error.code() {
            // Rust's suggestions name Rust APIs: `Signed::unsign`, which
            // `File.unsign()` exposes, and `Locked::unlock`, which the
            // bindings do not expose yet.
            Code::DocumentProtected => match error.suggestion() {
                Some(text) if text.contains("unsign") => Some(SIGNED_SUGGESTION.into()),
                Some(text) if !text.contains("unlock") => Some(text.to_owned().into()),
                _ => Some(PROTECTED_SUGGESTION.into()),
            },
            // The limits are the library defaults; Python offers no open
            // options to raise them.
            Code::LimitExceeded => None,
            _ => error.suggestion().map(|text| text.to_owned().into()),
        };
        failure
    }

    /// A failure of `opener`: an open, or the family check that follows it.
    /// A refusal that detected another family's format names the opener
    /// that reads it.
    pub(crate) fn open(error: &litchi::Error, opener: Opener) -> Self {
        let mut failure = Self::read(error);
        if error.code() == Code::UnsupportedFormat
            && let Some(detected) = error.detected()
        {
            failure.suggestion = match Opener::of(detected) {
                Some(right) if right != opener => {
                    Some(format!("open it with {}", right.name()).into())
                },
                // A format of this family whose reader this build lacks
                // keeps the build hint.
                Some(_) => failure.suggestion,
                None => None,
            };
        }
        failure
    }

    /// A call on a `File` after its `close()`.
    pub(crate) fn closed() -> Self {
        Self::new(Code::FileRequired, "the file is closed")
            .with_suggestion("open it again with litchi_py.open")
    }

    /// An argument a Python caller passed that the call cannot take.
    pub(crate) fn argument(message: impl Into<Cow<'static, str>>) -> Self {
        Self::new(Code::InvalidArgument, message)
    }

    /// The same failure, attributed to the zero-based batch item or edit
    /// operation `item`.
    pub(crate) fn with_item(mut self, item: usize) -> Self {
        self.item = Some(item);
        self
    }

    /// The error code's text.
    pub(crate) fn code(&self) -> &'static str {
        self.code.as_str()
    }

    /// The message.
    pub(crate) fn message(&self) -> &str {
        &self.message
    }

    /// The suggestion, if any.
    pub(crate) fn suggestion(&self) -> Option<&str> {
        self.suggestion.as_deref()
    }

    /// The failure as a unified error, for a Rust closure that must return
    /// one (`litchi::File::edit`).
    pub(crate) fn into_error(self) -> litchi::Error {
        let error = litchi::Error::new(self.code, self.message);
        match self.suggestion {
            Some(suggestion) => error.with_suggestion(suggestion),
            None => error,
        }
    }

    /// The same failure with `suggestion`.
    pub(crate) fn with_suggestion(mut self, suggestion: impl Into<Cow<'static, str>>) -> Self {
        self.suggestion = Some(suggestion.into());
        self
    }

    /// A detected format these bindings have no `FileFormat` value for: a
    /// format a newer Litchi added before the bindings learned it.
    pub(crate) fn unknown_format(format: FileFormat) -> Self {
        Self::new(
            Code::InternalError,
            format!(
                "Litchi detected {format:?}, which this build of litchi_py has no FileFormat value for"
            ),
        )
    }

    /// The Python exception.
    pub(crate) fn into_pyerr(self, py: Python<'_>) -> PyErr {
        let classes = match classes(py) {
            Ok(classes) => classes,
            Err(error) => return error,
        };
        let class = classes.0[self.class as usize].bind(py);
        match class.call1((self.message, self.code.as_str(), self.suggestion, self.item)) {
            Ok(instance) => PyErr::from_value(instance),
            Err(error) => error,
        }
    }
}

impl From<Failure> for PyErr {
    fn from(failure: Failure) -> Self {
        Python::attach(|py| failure.into_pyerr(py))
    }
}

/// Convert a unified read error to its Python exception.
pub(crate) fn to_py_err(error: litchi::Error) -> PyErr {
    Failure::read(&error).into()
}
