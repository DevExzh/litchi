//! Checked RTF fresh-authoring builder.
//!
//! [`Builder`] is the small authoring surface for a newly created RTF package.
//! It deliberately owns only the plain body text and the two policies that
//! affect the generated source.  The native writer remains the one place that
//! knows how to escape text and declare a character set.

use crate::Package;
use crate::codec::limits::{OpenOptions, SourceLimitKind, SourceLimits};
use crate::write::{Charset, Options, Writer};
use std::fmt;
use std::io::{self, Write};

/// Result of a checked RTF authoring operation.
pub type Result<T> = std::result::Result<T, Error>;

/// Errors returned while constructing a fresh RTF package.
#[derive(Debug)]
#[non_exhaustive]
pub enum Error {
    /// The source profile rejected an input or the generated source.
    Source(crate::Error),
    /// The native writer or package parser reported an I/O failure.
    Io(io::Error),
}

impl fmt::Display for Error {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Self::Source(error) => error.fmt(formatter),
            Self::Io(error) => error.fmt(formatter),
        }
    }
}

impl std::error::Error for Error {
    fn source(&self) -> Option<&(dyn std::error::Error + 'static)> {
        match self {
            Self::Source(error) => Some(error),
            Self::Io(error) => Some(error),
        }
    }
}

impl From<crate::Error> for Error {
    fn from(error: crate::Error) -> Self {
        Self::Source(error)
    }
}

impl From<io::Error> for Error {
    fn from(error: io::Error) -> Self {
        Self::Io(error)
    }
}

/// A fresh RTF authoring configuration.
///
/// `text` replaces the body text currently held by the builder.  Text is
/// copied only after its byte length has been admitted by the selected source
/// profile.  `build` then applies the same profile to the complete encoded
/// source and reopens it through the immutable [`Package`] owner.
///
/// ```
/// use litchi_rtf::builder::Builder;
///
/// let package: litchi_rtf::Package = Builder::new().text("Hello")?.build()?;
/// assert_eq!(package.text(), "Hello");
/// # Ok::<(), litchi_rtf::builder::Error>(())
/// ```
///
/// The builder is an authoring configuration.  It has no direct byte or file
/// publication methods; those belong to the [`Package`] returned by
/// [`Self::build`].
///
/// ```compile_fail
/// use litchi_rtf::Builder;
///
/// let builder = Builder::new();
/// builder.to_bytes();
/// ```
///
/// ```compile_fail
/// use litchi_rtf::Builder;
///
/// let builder = Builder::new();
/// builder.save("output.rtf");
/// ```
pub struct Builder {
    source_limits: SourceLimits,
    charset: Charset,
    text: String,
}

impl Builder {
    /// Start a builder with the production source profile and Windows-1252
    /// ANSI declaration.
    #[must_use]
    pub fn new() -> Self {
        Self {
            source_limits: SourceLimits::default(),
            charset: Charset::WINDOWS_1252,
            text: String::new(),
        }
    }

    /// Select the checked physical source profile used for text admission and
    /// the final strict package reopen.
    #[must_use]
    pub const fn source_limits(mut self, source_limits: SourceLimits) -> Self {
        self.source_limits = source_limits;
        self
    }

    /// Select the checked legacy character set declared in the generated RTF
    /// header.
    #[must_use]
    pub const fn charset(mut self, charset: Charset) -> Self {
        self.charset = charset;
        self
    }

    /// Replace the plain body text.
    ///
    /// The input is checked before the builder allocates its copy.  The
    /// generated RTF is checked again during [`Self::build`], because escaping
    /// control characters and Unicode can expand the encoded source.
    ///
    /// # Errors
    /// Returns [`Error::Source`] when the input text already exceeds the
    /// selected physical source-byte ceiling, or when the copy cannot reserve
    /// its required capacity.
    pub fn text(mut self, text: &str) -> Result<Self> {
        let maximum = self.source_limits.max_input_bytes();
        if text.len() > maximum {
            return Err(Error::Source(crate::Error::Limit {
                kind: SourceLimitKind::InputBytes,
                observed: text.len(),
                maximum,
            }));
        }

        let mut copied = String::new();
        copied.try_reserve_exact(text.len()).map_err(|_error| {
            Error::Source(crate::Error::AllocationFailed {
                resource: "builder text",
                requested: text.len(),
            })
        })?;
        copied.push_str(text);
        self.text = copied;
        Ok(self)
    }

    /// Encode and strictly reopen the newly authored package.
    ///
    /// The body is emitted through the native writer's header and text
    /// primitives, so escaping, newline handling, and Unicode fallback rules
    /// stay identical to the existing writer API.  Publication and filesystem
    /// saving remain responsibilities of [`Package`].
    ///
    /// # Errors
    /// Returns a typed source-limit, writer-I/O, or strict-reopen error.
    pub fn build(self) -> Result<Package> {
        let mut sink = BoundedSink::new(self.source_limits.max_input_bytes());
        let mut writer = Writer::with_options(
            &mut sink,
            Options {
                charset: self.charset,
                ..Options::default()
            },
        );
        let write_result = writer
            .write_document_header()
            .and_then(|()| writer.write_text(&self.text))
            .and_then(|()| writer.write_str("}"))
            .and_then(|()| writer.flush());
        drop(writer);

        if let Err(error) = write_result {
            if let Some(source_error) = sink.take_source_error() {
                return Err(Error::Source(source_error));
            }
            return Err(Error::Io(error));
        }

        let bytes = sink.into_inner();
        Package::from_owned_bytes_with_options(
            bytes,
            OpenOptions::new().source_limits(self.source_limits),
        )
        .map_err(Error::Source)
    }
}

impl Default for Builder {
    fn default() -> Self {
        Self::new()
    }
}

impl fmt::Debug for Builder {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("Builder")
            .field("source_limits", &self.source_limits)
            .field("charset", &self.charset)
            .field("text_bytes", &self.text.len())
            .finish_non_exhaustive()
    }
}

struct BoundedSink {
    bytes: Vec<u8>,
    maximum: usize,
    source_error: Option<crate::Error>,
}

impl BoundedSink {
    fn new(maximum: usize) -> Self {
        Self {
            bytes: Vec::new(),
            maximum,
            source_error: None,
        }
    }

    fn take_source_error(&mut self) -> Option<crate::Error> {
        self.source_error.take()
    }

    fn into_inner(self) -> Vec<u8> {
        self.bytes
    }
}

impl Write for BoundedSink {
    fn write(&mut self, bytes: &[u8]) -> io::Result<usize> {
        let observed = self.bytes.len().saturating_add(bytes.len());
        if observed > self.maximum {
            let error = crate::Error::Limit {
                kind: SourceLimitKind::InputBytes,
                observed,
                maximum: self.maximum,
            };
            self.source_error = Some(error);
            return Err(io::Error::other("RTF source limit exceeded"));
        }

        let capacity = observed
            .max(self.bytes.capacity().saturating_mul(2))
            .max(64)
            .min(self.maximum);
        if observed > self.bytes.capacity()
            && self
                .bytes
                .try_reserve_exact(capacity - self.bytes.len())
                .is_err()
        {
            self.source_error = Some(crate::Error::AllocationFailed {
                resource: "builder output",
                requested: observed,
            });
            return Err(io::Error::other("RTF builder output allocation failed"));
        }
        self.bytes.extend_from_slice(bytes);
        Ok(bytes.len())
    }

    fn flush(&mut self) -> io::Result<()> {
        Ok(())
    }
}

#[cfg(test)]
mod tests {
    use super::{Builder, Error};
    use crate::{Package, SourceLimitKind, SourceLimits, write};

    #[test]
    fn output_growth_is_amortized_and_bounded() {
        use std::io::Write;
        let mut sink = super::BoundedSink::new(1000);
        let mut growths = 0;
        for _ in 0..1000 {
            let before = sink.bytes.capacity();
            sink.write_all(b"x").expect("bounded write");
            growths += usize::from(sink.bytes.capacity() != before);
            assert!(sink.bytes.capacity() <= 1000);
        }
        assert!(growths <= 5);
        assert!(sink.write_all(b"x").is_err());
        assert_eq!(sink.bytes.len(), 1000);
    }

    #[test]
    fn empty_builder_round_trips_through_package_owner() {
        let package: Package = Builder::new().build().expect("empty package");
        assert_eq!(package.text(), "");
        assert_eq!(package.paragraph_count(), 0);
        assert_eq!(package.source_limits(), SourceLimits::default());
    }

    #[test]
    fn builder_uses_native_text_escaping_and_newline_semantics() {
        let package = Builder::new()
            .text("plain\\{}\n\t\r—é")
            .expect("text")
            .build()
            .expect("package");
        assert_eq!(package.text(), "plain\\{}\n\t\r—é");

        let mut bytes = Vec::new();
        package.write_to(&mut bytes).expect("write");
        let encoded = String::from_utf8(bytes).expect("UTF-8 generated RTF");
        assert!(encoded.contains(r"plain\\\{\}\par \tab \'0d\emdash \u233?"));
    }

    #[test]
    fn charset_is_declared_by_native_header_writer() {
        let package = Builder::new()
            .charset(write::Charset::Pc)
            .text("body")
            .expect("text")
            .build()
            .expect("package");
        let mut bytes = Vec::new();
        package.write_to(&mut bytes).expect("write");
        assert!(
            String::from_utf8(bytes)
                .expect("UTF-8 generated RTF")
                .contains(r"\pc")
        );
    }

    #[test]
    fn text_admission_happens_before_copy_and_tightened_build_limit_is_enforced() {
        let limits = SourceLimits::builder()
            .max_input_bytes(1)
            .expect("limit")
            .build()
            .expect("limits");
        let error = Builder::new()
            .source_limits(limits)
            .text("too large")
            .expect_err("text should be rejected");
        assert!(matches!(
            error,
            Error::Source(crate::Error::Limit {
                kind: SourceLimitKind::InputBytes,
                ..
            })
        ));

        let text_limits = SourceLimits::builder()
            .max_input_bytes(32)
            .expect("limit")
            .build()
            .expect("limits");
        let error = Builder::new()
            .text("body")
            .expect("text")
            .source_limits(text_limits)
            .build()
            .expect_err("header plus body should exceed the tiny source budget");
        assert!(matches!(
            error,
            Error::Source(crate::Error::Limit {
                kind: SourceLimitKind::InputBytes,
                ..
            })
        ));
    }

    #[test]
    fn strict_reopen_enforces_non_byte_source_limits() {
        let limits = SourceLimits::builder()
            .max_events(0)
            .expect("event limit")
            .build()
            .expect("profile");
        let error = Builder::new()
            .source_limits(limits)
            .build()
            .expect_err("events denied");
        assert!(matches!(
            error,
            Error::Source(crate::Error::Limit {
                kind: SourceLimitKind::Events,
                maximum: 0,
                ..
            })
        ));
    }

    #[test]
    fn built_package_is_immutable_and_still_editable() {
        let package = Builder::new()
            .text("before")
            .expect("text")
            .build()
            .expect("package");
        let mut edit = package.edit();
        edit.replace_paragraph_text(litchi_core::Position::new(0), "after")
            .expect("replace");
        let changed = edit.commit().expect("commit").package().clone();
        assert_eq!(package.text(), "before");
        assert_eq!(changed.text(), "after");
    }

    #[test]
    fn package_new_uses_the_same_empty_builder_path() {
        let from_package = Package::new().expect("new package");
        let from_builder = Builder::default().build().expect("built package");
        assert_eq!(from_package.text(), from_builder.text());
        assert_eq!(
            from_package.to_bytes().expect("bytes"),
            from_builder.to_bytes().expect("bytes")
        );
    }
}
