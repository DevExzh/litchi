//! Bounded semantic plain-text output to caller-owned sequential sinks.

use std::error::Error as StdError;
use std::fmt;
use std::io::{self, Write};

/// Conservative default ceiling for one semantic text conversion.
pub const DEFAULT_MAX_TEXT_OUTPUT_BYTES: u64 = 64 * 1024 * 1024;
/// Conservative default ceiling for semantic objects emitted by one conversion.
pub const DEFAULT_MAX_TEXT_OUTPUT_OBJECTS: u64 = 1_000_000;

/// Separator used before each object after the first emitted object.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
#[non_exhaustive]
pub enum TextObjectKind {
    /// A Word-processing paragraph or equivalent text block.
    Paragraph,
    /// A presentation slide.
    Slide,
}

/// Explicit policy for bounded semantic plain-text conversion.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub struct TextOutputOptions<'a> {
    paragraph_separator: &'a str,
    slide_separator: &'a str,
    max_output_bytes: u64,
    max_objects: u64,
    include_empty_objects: bool,
}

impl<'a> TextOutputOptions<'a> {
    /// Construct a policy with caller-selected separators and resource ceilings.
    #[must_use]
    pub const fn new(
        paragraph_separator: &'a str,
        slide_separator: &'a str,
        max_output_bytes: u64,
        max_objects: u64,
    ) -> Self {
        Self {
            paragraph_separator,
            slide_separator,
            max_output_bytes,
            max_objects,
            include_empty_objects: true,
        }
    }

    /// Include or omit empty semantic objects.
    #[must_use]
    pub const fn with_empty_objects(mut self, include: bool) -> Self {
        self.include_empty_objects = include;
        self
    }

    /// Separator between emitted paragraphs.
    #[must_use]
    pub const fn paragraph_separator(self) -> &'a str {
        self.paragraph_separator
    }

    /// Separator between emitted slides.
    #[must_use]
    pub const fn slide_separator(self) -> &'a str {
        self.slide_separator
    }

    /// Maximum number of bytes that may be written.
    #[must_use]
    pub const fn max_output_bytes(self) -> u64 {
        self.max_output_bytes
    }

    /// Maximum number of semantic objects that may be completed.
    #[must_use]
    pub const fn max_objects(self) -> u64 {
        self.max_objects
    }

    /// Whether empty semantic objects are emitted and counted.
    #[must_use]
    pub const fn include_empty_objects(self) -> bool {
        self.include_empty_objects
    }

    fn separator(self, kind: TextObjectKind) -> &'a str {
        match kind {
            TextObjectKind::Paragraph => self.paragraph_separator,
            TextObjectKind::Slide => self.slide_separator,
        }
    }
}

impl Default for TextOutputOptions<'static> {
    fn default() -> Self {
        Self::new(
            "\n",
            "\n\n",
            DEFAULT_MAX_TEXT_OUTPUT_BYTES,
            DEFAULT_MAX_TEXT_OUTPUT_OBJECTS,
        )
    }
}

/// Exact progress made in a caller-owned sequential sink.
#[derive(Clone, Copy, Debug, Default, PartialEq, Eq)]
pub struct TextOutputReport {
    bytes_written: u64,
    objects_written: u64,
}

impl TextOutputReport {
    /// Number of bytes accepted by the sink.
    #[must_use]
    pub const fn bytes_written(self) -> u64 {
        self.bytes_written
    }

    /// Number of complete semantic objects written.
    #[must_use]
    pub const fn objects_written(self) -> u64 {
        self.objects_written
    }
}

/// Resource whose caller-selected conversion ceiling was exceeded.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
#[non_exhaustive]
pub enum TextOutputLimitKind {
    /// UTF-8 output bytes, including separators.
    OutputBytes,
    /// Complete semantic objects.
    Objects,
}

/// One rejected output-resource observation.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub struct TextOutputLimit {
    kind: TextOutputLimitKind,
    observed: u64,
    limit: u64,
}

impl TextOutputLimit {
    /// Resource that exceeded its ceiling.
    #[must_use]
    pub const fn kind(self) -> TextOutputLimitKind {
        self.kind
    }

    /// Value required by the rejected operation.
    #[must_use]
    pub const fn observed(self) -> u64 {
        self.observed
    }

    /// Caller-selected ceiling.
    #[must_use]
    pub const fn limit(self) -> u64 {
        self.limit
    }
}

/// Failure from semantic conversion, retaining truthful partial-output progress.
#[derive(Debug)]
#[non_exhaustive]
pub enum TextOutputError<E> {
    /// The format owner rejected or could not decode semantic content.
    Document {
        /// Format-specific source error.
        source: E,
        /// Progress before the source error was observed.
        progress: TextOutputReport,
    },
    /// A caller-selected resource ceiling was exceeded.
    Limit {
        /// Rejected resource observation.
        limit: TextOutputLimit,
        /// Progress before the limit was observed.
        progress: TextOutputReport,
    },
    /// The caller-owned sink failed after accepting the reported bytes.
    Sink {
        /// Sink error.
        source: io::Error,
        /// Exact progress before the failed write.
        progress: TextOutputReport,
    },
    /// A public replay factory returned a different fragment extent during
    /// emission than during preflight.
    ///
    /// This diagnostic is available only when emission begins. A conversion
    /// refused by an empty-object or resource-limit preflight never requests a
    /// second iterator and therefore cannot compare replays.
    NonDeterministicFragments {
        /// Number of fragments observed during preflight.
        preflight_fragments: u64,
        /// Number of fragments observed during emission.
        emitted_fragments: u64,
        /// Joined payload bytes observed during preflight.
        preflight_bytes: u64,
        /// Joined payload bytes observed during emission.
        emitted_bytes: u64,
        /// Exact output accepted before the mismatch was detected.
        progress: TextOutputReport,
    },
}

impl<E> TextOutputError<E> {
    /// Progress completed before the failure.
    #[must_use]
    pub const fn progress(&self) -> TextOutputReport {
        match self {
            Self::Document { progress, .. }
            | Self::Limit { progress, .. }
            | Self::Sink { progress, .. }
            | Self::NonDeterministicFragments { progress, .. } => *progress,
        }
    }

    /// Rejected resource observation, when this is a limit error.
    #[must_use]
    pub const fn limit(&self) -> Option<TextOutputLimit> {
        match self {
            Self::Limit { limit, .. } => Some(*limit),
            Self::Document { .. } | Self::Sink { .. } | Self::NonDeterministicFragments { .. } => {
                None
            },
        }
    }
}

impl<E: fmt::Display> fmt::Display for TextOutputError<E> {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Self::Document { source, progress } => write!(
                formatter,
                "semantic text conversion failed after {} bytes and {} objects: {source}",
                progress.bytes_written, progress.objects_written
            ),
            Self::Limit { limit, progress } => write!(
                formatter,
                "semantic text {:?} limit exceeded (observed {}, limit {}) after {} bytes and {} objects",
                limit.kind,
                limit.observed,
                limit.limit,
                progress.bytes_written,
                progress.objects_written
            ),
            Self::Sink { source, progress } => write!(
                formatter,
                "semantic text sink failed after {} bytes and {} objects: {source}",
                progress.bytes_written, progress.objects_written
            ),
            Self::NonDeterministicFragments {
                preflight_fragments,
                emitted_fragments,
                preflight_bytes,
                emitted_bytes,
                progress,
            } => write!(
                formatter,
                "semantic text fragment replay changed from {preflight_fragments} fragments/{preflight_bytes} bytes to {emitted_fragments} fragments/{emitted_bytes} bytes after {} bytes and {} objects",
                progress.bytes_written, progress.objects_written
            ),
        }
    }
}

impl<E: StdError + 'static> StdError for TextOutputError<E> {
    fn source(&self) -> Option<&(dyn StdError + 'static)> {
        match self {
            Self::Document { source, .. } => Some(source),
            Self::Sink { source, .. } => Some(source),
            Self::Limit { .. } | Self::NonDeterministicFragments { .. } => None,
        }
    }
}

/// Shared bounded writer used by format-owned semantic traversals.
///
/// This is public so concrete format crates can share one limit and partial-
/// output contract. Ordinary callers normally use a format's `write_text_to`
/// method instead.
pub struct SequentialTextWriter<'a, 'output, W: Write + ?Sized> {
    output: &'output mut W,
    options: TextOutputOptions<'a>,
    progress: TextOutputReport,
}

impl<'a, 'output, W: Write + ?Sized> SequentialTextWriter<'a, 'output, W> {
    /// Start writing to a caller-owned sequential sink.
    #[must_use]
    pub const fn new(output: &'output mut W, options: TextOutputOptions<'a>) -> Self {
        Self {
            output,
            options,
            progress: TextOutputReport {
                bytes_written: 0,
                objects_written: 0,
            },
        }
    }

    /// Current exact progress.
    #[must_use]
    pub const fn progress(&self) -> TextOutputReport {
        self.progress
    }

    fn start_object<E>(&mut self, kind: TextObjectKind) -> Result<(), TextOutputError<E>> {
        let observed = self
            .progress
            .objects_written
            .checked_add(1)
            .ok_or_else(|| {
                self.limit_error(
                    TextOutputLimitKind::Objects,
                    u64::MAX,
                    self.options.max_objects,
                )
            })?;
        if observed > self.options.max_objects {
            return Err(self.limit_error(
                TextOutputLimitKind::Objects,
                observed,
                self.options.max_objects,
            ));
        }
        if self.progress.objects_written != 0 {
            self.write_bytes(self.options.separator(kind).as_bytes())?;
        }
        Ok(())
    }

    /// Write one already-borrowed object without an intermediate allocation.
    ///
    /// Empty objects are skipped when requested by the policy.
    ///
    /// # Errors
    /// Returns a typed limit or exact partial sink failure.
    pub fn write_object<E>(
        &mut self,
        kind: TextObjectKind,
        value: &str,
    ) -> Result<(), TextOutputError<E>> {
        if value.is_empty() && !self.options.include_empty_objects {
            return Ok(());
        }
        let separator_bytes = if self.progress.objects_written == 0 {
            0
        } else {
            self.options.separator(kind).len()
        };
        let required = separator_bytes
            .checked_add(value.len())
            .and_then(|value| u64::try_from(value).ok())
            .and_then(|value| self.progress.bytes_written.checked_add(value))
            .unwrap_or(u64::MAX);
        if required > self.options.max_output_bytes {
            return Err(self.limit_error(
                TextOutputLimitKind::OutputBytes,
                required,
                self.options.max_output_bytes,
            ));
        }
        self.start_object(kind)?;
        self.write_bytes(value.as_bytes())?;
        self.progress.objects_written = self.progress.objects_written.saturating_add(1);
        Ok(())
    }

    /// Write one object from replayable borrowed fragments without flattening it.
    ///
    /// `fragments` is called once for resource preflight and a second time only
    /// if emission begins. Every call must construct a fresh iterator over the
    /// same immutable fragment sequence. Deterministic replay is a caller
    /// precondition: the factory must not inspect or mutate advancing external
    /// state. When emission occurs, fragment count and joined byte extent are
    /// compared afterward; that parity check detects extent divergence but is
    /// not a substitute for the deterministic-replay precondition.
    /// `fragment_separator` is written between fragments and contributes to both
    /// emptiness and byte limits. Therefore `[]` and `[""]` are empty, while
    /// `["", ""]` joined by `"-"` emits `"-"` and is not empty.
    ///
    /// # Errors
    /// Returns a typed resource-limit or exact partial sink failure. If
    /// emission occurs and the two iterator extents differ, returns
    /// [`TextOutputError::NonDeterministicFragments`] with exact partial-output
    /// progress.
    pub fn write_joined_object<'fragment, E, I, F>(
        &mut self,
        kind: TextObjectKind,
        mut fragments: F,
        fragment_separator: &str,
    ) -> Result<(), TextOutputError<E>>
    where
        I: Iterator<Item = &'fragment str>,
        F: FnMut() -> I,
    {
        let mut fragment_count = 0usize;
        let mut value_bytes = 0usize;
        for fragment in fragments() {
            if fragment_count != 0 {
                value_bytes = value_bytes
                    .checked_add(fragment_separator.len())
                    .unwrap_or(usize::MAX);
            }
            value_bytes = value_bytes
                .checked_add(fragment.len())
                .unwrap_or(usize::MAX);
            fragment_count = fragment_count.saturating_add(1);
        }
        if value_bytes == 0 && !self.options.include_empty_objects {
            return Ok(());
        }

        let object_separator_bytes = if self.progress.objects_written == 0 {
            0
        } else {
            self.options.separator(kind).len()
        };
        let required = object_separator_bytes
            .checked_add(value_bytes)
            .and_then(|value| u64::try_from(value).ok())
            .and_then(|value| self.progress.bytes_written.checked_add(value))
            .unwrap_or(u64::MAX);
        if required > self.options.max_output_bytes {
            return Err(self.limit_error(
                TextOutputLimitKind::OutputBytes,
                required,
                self.options.max_output_bytes,
            ));
        }

        self.start_object(kind)?;
        let mut emitted_fragments = 0usize;
        let mut emitted_bytes = 0usize;
        for fragment in fragments() {
            if emitted_fragments != 0 {
                emitted_bytes = emitted_bytes
                    .checked_add(fragment_separator.len())
                    .unwrap_or(usize::MAX);
                self.write_bytes(fragment_separator.as_bytes())?;
            }
            emitted_bytes = emitted_bytes
                .checked_add(fragment.len())
                .unwrap_or(usize::MAX);
            emitted_fragments = emitted_fragments.saturating_add(1);
            self.write_bytes(fragment.as_bytes())?;
        }
        if emitted_fragments != fragment_count || emitted_bytes != value_bytes {
            return Err(TextOutputError::NonDeterministicFragments {
                preflight_fragments: u64::try_from(fragment_count).unwrap_or(u64::MAX),
                emitted_fragments: u64::try_from(emitted_fragments).unwrap_or(u64::MAX),
                preflight_bytes: u64::try_from(value_bytes).unwrap_or(u64::MAX),
                emitted_bytes: u64::try_from(emitted_bytes).unwrap_or(u64::MAX),
                progress: self.progress,
            });
        }
        self.progress.objects_written = self.progress.objects_written.saturating_add(1);
        Ok(())
    }

    /// Wrap a format-specific semantic failure with current output progress.
    pub fn document_error<E>(&self, source: E) -> TextOutputError<E> {
        TextOutputError::Document {
            source,
            progress: self.progress,
        }
    }

    /// Finish the conversion and return exact progress.
    #[must_use]
    pub const fn finish(self) -> TextOutputReport {
        self.progress
    }

    fn limit_error<E>(
        &self,
        kind: TextOutputLimitKind,
        observed: u64,
        limit: u64,
    ) -> TextOutputError<E> {
        TextOutputError::Limit {
            limit: TextOutputLimit {
                kind,
                observed,
                limit,
            },
            progress: self.progress,
        }
    }

    fn write_bytes<E>(&mut self, mut bytes: &[u8]) -> Result<(), TextOutputError<E>> {
        let fragment_bytes = u64::try_from(bytes.len()).unwrap_or(u64::MAX);
        let observed = self
            .progress
            .bytes_written
            .checked_add(fragment_bytes)
            .unwrap_or(u64::MAX);
        if observed > self.options.max_output_bytes {
            return Err(self.limit_error(
                TextOutputLimitKind::OutputBytes,
                observed,
                self.options.max_output_bytes,
            ));
        }

        while !bytes.is_empty() {
            match self.output.write(bytes) {
                Ok(0) => {
                    return Err(TextOutputError::Sink {
                        source: io::Error::new(
                            io::ErrorKind::WriteZero,
                            "failed to write semantic text fragment",
                        ),
                        progress: self.progress,
                    });
                },
                Ok(written) if written <= bytes.len() => {
                    self.progress.bytes_written =
                        self.progress.bytes_written.saturating_add(written as u64);
                    bytes = &bytes[written..];
                },
                Ok(_) => {
                    return Err(TextOutputError::Sink {
                        source: io::Error::new(
                            io::ErrorKind::InvalidData,
                            "sink reported writing beyond the supplied semantic text fragment",
                        ),
                        progress: self.progress,
                    });
                },
                Err(error) if error.kind() == io::ErrorKind::Interrupted => {},
                Err(source) => {
                    return Err(TextOutputError::Sink {
                        source,
                        progress: self.progress,
                    });
                },
            }
        }
        Ok(())
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[derive(Default)]
    struct FailAfter {
        accepted: Vec<u8>,
        limit: usize,
    }

    #[derive(Default)]
    struct InterruptedOnce {
        output: Vec<u8>,
        interrupted: bool,
    }

    impl Write for InterruptedOnce {
        fn write(&mut self, bytes: &[u8]) -> io::Result<usize> {
            if !self.interrupted {
                self.interrupted = true;
                return Err(io::Error::new(io::ErrorKind::Interrupted, "retry"));
            }
            self.output.extend_from_slice(bytes);
            Ok(bytes.len())
        }

        fn flush(&mut self) -> io::Result<()> {
            Ok(())
        }
    }

    struct WriteZero;

    impl Write for WriteZero {
        fn write(&mut self, _bytes: &[u8]) -> io::Result<usize> {
            Ok(0)
        }

        fn flush(&mut self) -> io::Result<()> {
            Ok(())
        }
    }

    impl Write for FailAfter {
        fn write(&mut self, bytes: &[u8]) -> io::Result<usize> {
            if self.accepted.len() == self.limit {
                return Err(io::Error::other("injected failure"));
            }
            let remaining = self.limit.saturating_sub(self.accepted.len());
            let accepted = remaining.min(bytes.len());
            self.accepted.extend_from_slice(&bytes[..accepted]);
            Ok(accepted)
        }

        fn flush(&mut self) -> io::Result<()> {
            Ok(())
        }
    }

    #[test]
    fn writes_deterministic_separators_and_counts_complete_objects() {
        let options = TextOutputOptions::new("|", "::", 64, 4);
        let mut output = Vec::new();
        let mut writer = SequentialTextWriter::new(&mut output, options);
        writer
            .write_object::<io::Error>(TextObjectKind::Paragraph, "alpha")
            .unwrap();
        writer
            .write_object::<io::Error>(TextObjectKind::Paragraph, "β")
            .unwrap();
        let report = writer.finish();
        assert_eq!(output, "alpha|β".as_bytes());
        assert_eq!(report.bytes_written(), 8);
        assert_eq!(report.objects_written(), 2);
    }

    #[test]
    fn limits_refuse_the_next_complete_object_before_output() {
        let options = TextOutputOptions::new("|", "::", 64, 1);
        let mut output = Vec::new();
        let mut writer = SequentialTextWriter::new(&mut output, options);
        writer
            .write_object::<io::Error>(TextObjectKind::Paragraph, "first")
            .unwrap();
        let error = writer
            .write_object::<io::Error>(TextObjectKind::Paragraph, "second")
            .unwrap_err();
        assert_eq!(output, b"first");
        assert_eq!(error.progress().bytes_written(), 5);
        assert_eq!(error.progress().objects_written(), 1);
        assert_eq!(error.limit().unwrap().kind(), TextOutputLimitKind::Objects);
    }

    #[test]
    fn reports_exact_multibyte_partial_sink_progress() {
        let options = TextOutputOptions::new("|", "::", 64, 4);
        let mut output = FailAfter {
            accepted: Vec::new(),
            limit: 4,
        };
        let mut writer = SequentialTextWriter::new(&mut output, options);
        let error = writer
            .write_object::<io::Error>(TextObjectKind::Paragraph, "é文")
            .unwrap_err();
        assert_eq!(output.accepted, "é文".as_bytes()[..4]);
        assert_eq!(error.progress().bytes_written(), 4);
        assert_eq!(error.progress().objects_written(), 0);
        assert!(matches!(error, TextOutputError::Sink { .. }));
    }

    #[test]
    fn joined_objects_preflight_multibyte_output_and_empty_policy() {
        let options = TextOutputOptions::new("|", "::", 64, 4).with_empty_objects(false);
        let mut output = Vec::new();
        let mut writer = SequentialTextWriter::new(&mut output, options);
        writer
            .write_joined_object::<io::Error, _, _>(
                TextObjectKind::Slide,
                || ["é", "文"].into_iter(),
                "-",
            )
            .unwrap();
        writer
            .write_joined_object::<io::Error, _, _>(TextObjectKind::Slide, std::iter::empty, "-")
            .unwrap();
        writer
            .write_joined_object::<io::Error, _, _>(TextObjectKind::Slide, || [""].into_iter(), "-")
            .unwrap();
        writer
            .write_joined_object::<io::Error, _, _>(
                TextObjectKind::Slide,
                || ["", ""].into_iter(),
                "-",
            )
            .unwrap();
        let report = writer.finish();
        assert_eq!(output, "é-文::-".as_bytes());
        assert_eq!(report.bytes_written(), 9);
        assert_eq!(report.objects_written(), 2);
    }

    #[test]
    fn joined_byte_limit_refuses_before_any_output() {
        let options = TextOutputOptions::new("|", "::", 5, 4);
        let mut output = Vec::new();
        let mut writer = SequentialTextWriter::new(&mut output, options);
        let error = writer
            .write_joined_object::<io::Error, _, _>(
                TextObjectKind::Paragraph,
                || ["é", "文"].into_iter(),
                "-",
            )
            .unwrap_err();
        assert!(output.is_empty());
        assert_eq!(error.progress(), TextOutputReport::default());
        let limit = error.limit().unwrap();
        assert_eq!(limit.kind(), TextOutputLimitKind::OutputBytes);
        assert_eq!(limit.observed(), 6);
    }

    #[test]
    fn interrupted_writes_retry_and_write_zero_is_truthful() {
        let options = TextOutputOptions::new("|", "::", 64, 4);
        let mut interrupted = InterruptedOnce::default();
        let mut writer = SequentialTextWriter::new(&mut interrupted, options);
        let report = writer
            .write_object::<io::Error>(TextObjectKind::Paragraph, "retry")
            .map(|()| writer.finish())
            .unwrap();
        assert_eq!(interrupted.output, b"retry");
        assert_eq!(report.bytes_written(), 5);

        let mut zero = WriteZero;
        let mut writer = SequentialTextWriter::new(&mut zero, options);
        let error = writer
            .write_object::<io::Error>(TextObjectKind::Paragraph, "zero")
            .unwrap_err();
        assert_eq!(error.progress(), TextOutputReport::default());
        assert!(matches!(
            error,
            TextOutputError::Sink { source, .. }
                if source.kind() == io::ErrorKind::WriteZero
        ));
    }

    #[test]
    fn detects_a_nondeterministic_replay_factory() {
        let options = TextOutputOptions::new("|", "::", 64, 4);
        let first = ["a"];
        let second = ["b", "c"];
        let mut calls = 0usize;
        let mut output = Vec::new();
        let mut writer = SequentialTextWriter::new(&mut output, options);
        let error = writer
            .write_joined_object::<io::Error, _, _>(
                TextObjectKind::Paragraph,
                || {
                    calls += 1;
                    if calls == 1 {
                        first.as_slice().iter().copied()
                    } else {
                        second.as_slice().iter().copied()
                    }
                },
                "-",
            )
            .unwrap_err();
        assert_eq!(output, b"b-c");
        assert_eq!(error.progress().bytes_written(), 3);
        assert_eq!(error.progress().objects_written(), 0);
        assert!(matches!(
            error,
            TextOutputError::NonDeterministicFragments {
                preflight_fragments: 1,
                emitted_fragments: 2,
                preflight_bytes: 1,
                emitted_bytes: 3,
                ..
            }
        ));
    }
}
