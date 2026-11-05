//! The facade's entry points into the unified operations API (ADR 0033).
//!
//! `litchi-ops` owns [`File`](crate::File) and its single-pass openers, and
//! since record 0296 (owner decision D1) it opens Apple Pages, Keynote and
//! Numbers packages itself, as `File::Pages`, `File::Keynote` and
//! `File::Numbers`, package directories included (record 0306). These
//! openers are `litchi_ops::open` and its variants.

use std::path::Path;

use litchi_ops::{File, OpenOptions, Result};

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
/// [`detected`](crate::Error::detected) format) and for a format whose cargo
/// feature is off (with the format and a suggestion),
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
    litchi_ops::open_with(path, options)
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
    litchi_ops::open_bytes_with(bytes, options)
}
