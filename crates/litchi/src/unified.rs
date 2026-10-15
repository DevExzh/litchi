//! The facade's entry points into the unified operations API (ADR 0033).
//!
//! `litchi-ops` owns [`File`](crate::File) and its single-pass openers, and
//! since record 0296 (owner decision D1) it opens Apple Pages, Keynote and
//! Numbers packages itself, as `File::Pages`, `File::Keynote` and
//! `File::Numbers`. These openers are `litchi_ops::open` with one facade
//! refinement: an iWork input the unified `File` cannot hold (an
//! app-authored package directory) is refused with a suggestion that names
//! the facade's read-only `litchi::iwork::Document`, or the `iwork` cargo
//! feature that builds it.

use std::path::Path;

#[cfg(all(
    any(feature = "pages", feature = "keynote", feature = "numbers"),
    not(feature = "iwork")
))]
use litchi_core::FileFormat;
#[cfg(all(
    any(feature = "pages", feature = "keynote", feature = "numbers"),
    not(feature = "iwork")
))]
use litchi_ops::Code;
use litchi_ops::{Error, File, OpenOptions, Result};

/// Open the file at `path`, detecting its format from the content.
///
/// One detection pass classifies the file and builds the owner its format
/// keeps: most formats stay source-backed on `unix` and `windows`, reading
/// payloads on demand. An encrypted input opens as [`File::Locked`] and a
/// signed one as [`File::Signed`].
///
/// # Errors
///
/// Returns [`Code::FileNotFound`](crate::Code::FileNotFound) when the path
/// does not exist, [`Code::UnsupportedFormat`](crate::Code::UnsupportedFormat)
/// for an input that is not an Office file (with no
/// [`detected`](Error::detected) format), for an Apple iWork package
/// directory (with the iWork format and a suggestion), and for a format
/// whose cargo feature is off (with the format and a suggestion),
/// [`Code::LimitExceeded`](crate::Code::LimitExceeded) when the input
/// exceeds a read limit, and the format's typed error when its container
/// cannot be read.
///
/// ```
/// # #[cfg(feature = "docx")]
/// # {
/// let path = concat!(
///     env!("CARGO_MANIFEST_DIR"),
///     "/../../test-data/ooxml/docx/documentProperties.docx"
/// );
/// let file = litchi::open(path)?;
/// assert_eq!(file.format(), Some(litchi::FileFormat::Docx));
/// assert_eq!(file.text()?.trim(), "Hello World!");
/// # }
/// # Ok::<(), litchi::Error>(())
/// ```
pub fn open(path: impl AsRef<Path>) -> Result<File> {
    open_with(path, &OpenOptions::default())
}

/// [`open`] with explicit [`OpenOptions`]: input and per-format limits,
/// owner selection, source retention and an execution context.
///
/// # Errors
///
/// As [`open`], plus [`Code::Cancelled`](crate::Code::Cancelled) when the
/// execution context is cancelled before or after the open.
///
/// ```no_run
/// use litchi::{OpenOptions, SourceMode};
///
/// let options = OpenOptions::new()
///     .max_input_bytes(64 << 20)
///     .source_mode(SourceMode::Eager);
/// let file = litchi::open_with("report.docx", &options)?;
/// # let _ = file;
/// # Ok::<(), litchi::Error>(())
/// ```
pub fn open_with(path: impl AsRef<Path>, options: &OpenOptions) -> Result<File> {
    litchi_ops::open_with(path, options).map_err(refuse)
}

/// Open an in-memory input, detecting its format from the content.
///
/// The bytes move into the handle without a copy; a borrowed slice is
/// copied once.
///
/// # Errors
///
/// As [`open`], without the file-system errors.
///
/// ```
/// let error = litchi::open_bytes(b"plain text".to_vec()).unwrap_err();
/// assert_eq!(error.code(), litchi::Code::UnsupportedFormat);
/// assert_eq!(error.detected(), None);
/// ```
pub fn open_bytes(bytes: impl Into<Vec<u8>>) -> Result<File> {
    open_bytes_with(bytes, &OpenOptions::default())
}

/// [`open_bytes`] with explicit [`OpenOptions`].
///
/// # Errors
///
/// As [`open_bytes`], plus [`Code::Cancelled`](crate::Code::Cancelled) when
/// the execution context is cancelled before or after the open.
///
/// ```
/// use litchi::{Code, OpenOptions};
///
/// let options = OpenOptions::new().max_input_bytes(4);
/// let error = litchi::open_bytes_with(b"plain text".to_vec(), &options).unwrap_err();
/// assert_eq!(error.code(), Code::LimitExceeded);
/// ```
pub fn open_bytes_with(bytes: impl Into<Vec<u8>>, options: &OpenOptions) -> Result<File> {
    litchi_ops::open_bytes_with(bytes, options).map_err(refuse)
}

/// Finish a refusal: the suggestion for an iWork package directory names
/// `litchi::iwork`, which only the aggregate `iwork` feature builds, so a
/// build without it says so.
fn refuse(error: Error) -> Error {
    #[cfg(all(
        any(feature = "pages", feature = "keynote", feature = "numbers"),
        not(feature = "iwork")
    ))]
    if error.code() == Code::UnsupportedFormat
        && matches!(
            error.detected(),
            Some(FileFormat::Pages | FileFormat::Keynote | FileFormat::Numbers)
        )
        && error
            .suggestion()
            .is_some_and(|suggestion| suggestion.contains("litchi::iwork"))
    {
        return error.with_suggestion(
            "enable the `iwork` cargo feature and open the file with `litchi::iwork::Document::open`",
        );
    }
    error
}
