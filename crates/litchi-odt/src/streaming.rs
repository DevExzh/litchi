//! Bounded fresh ODT publication from an ordered plain-text paragraph source.
//!
//! Each paragraph is consumed once and encoded into a bounded reusable XML
//! fragment. Spaces, tabs, and line feeds use ODF controls when needed to
//! preserve their meaning; carriage returns require an explicit caller policy
//! and are refused here. Rich content and edits to an existing document belong
//! to the corresponding document APIs.
//!
//! A failure can leave a partial caller sink. The returned error preserves its
//! cause and the number of bytes acknowledged by the sink; discard that output.
//!
//! ```
//! use litchi_core::ExecutionContext;
//! use litchi_odt::streaming::{StreamingLimits, stream_plain_paragraphs_to};
//! # fn example(context: &ExecutionContext) -> Result<(), Box<dyn std::error::Error>> {
//! let mut output = Vec::new();
//! let report = stream_plain_paragraphs_to(
//!     &mut output, ["Hello", "  spaces\tand\nlines  "], context, StreamingLimits::default(),
//! )?;
//! assert_eq!(report.paragraphs(), 2);
//! # Ok(()) }
//! ```

use std::{
    error::Error as StdError,
    fmt,
    io::{self, Write},
    result::Result,
};

use litchi_core::{Error, ExecutionContext, ExecutionError, Resource};
use litchi_odf_common::core::{
    GeneratedXmlEnvelope, GeneratedXmlLimits, PackageWriter, PackageWriterError,
    PackageWriterLimitResource, PackageWriterLimits,
};

const ODT_MIME: &str = "application/vnd.oasis.opendocument.text";
const CONTENT_PATH: &str = "content.xml";
const CONTENT_MEDIA_TYPE: &str = "text/xml";
const MAX_TEXT_BLOCKS: usize = 1_000_000;
const MAX_TEXT_BYTES: usize = 64 * 1024 * 1024;
const MAX_CONTENT_BYTES: usize = 256 * 1024 * 1024;
const MAX_SPACE_COUNT: usize = 1_000_000;
const DEFAULT_MAX_PARAGRAPHS: usize = MAX_TEXT_BLOCKS;
const DEFAULT_MAX_PARAGRAPH_TEXT_BYTES: usize = 1 << 20;
// The common audit default is 16 MiB of lexical character data. Keep the
// default provider profile constructible through the checked constructor;
// callers may select the 64 MiB ordinary-reader ceiling with a wider audit
// profile.
const DEFAULT_MAX_TOTAL_TEXT_BYTES: usize = 16 << 20;
const DEFAULT_MAX_PARAGRAPH_XML_BYTES: usize = 4 << 20;
const DEFAULT_MAX_CONTENT_XML_BYTES: usize = 32 << 20;
const DEFAULT_MAX_OUTPUT_BYTES: u64 = 512 << 20;
const COMMON_METADATA_RESERVATION: usize = 64 * 1024;
const MAX_MEMBER_NAME_BYTES: u64 = 32;
// Poll cancellation for each scalar admitted to a span, and bound each
// already-safe text write and its Work charge to this many UTF-8 bytes.
const MAX_PLAIN_TEXT_SPAN_BYTES: usize = 256;

// The common generated-XML envelope accepts declaration/start events in its
// prefix and only matching end events in its suffix.  The Builder's optional
// empty scripts/font-face/automatic-style children therefore cannot be put in
// this prefix.  They are optional for this simple text document; the static
// styles and meta members below retain the Builder defaults exactly.
// The buffered and streaming roles therefore compare semantic paragraph
// projections for canonical corpora; they must not claim lexical content-XML
// identity.
const CONTENT_PREFIX: &[u8] = b"<?xml version=\"1.0\" encoding=\"UTF-8\"?><office:document-content xmlns:office=\"urn:oasis:names:tc:opendocument:xmlns:office:1.0\" xmlns:style=\"urn:oasis:names:tc:opendocument:xmlns:style:1.0\" xmlns:text=\"urn:oasis:names:tc:opendocument:xmlns:text:1.0\" xmlns:table=\"urn:oasis:names:tc:opendocument:xmlns:table:1.0\" xmlns:draw=\"urn:oasis:names:tc:opendocument:xmlns:drawing:1.0\" xmlns:fo=\"urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0\" xmlns:xlink=\"http://www.w3.org/1999/xlink\" xmlns:dc=\"http://purl.org/dc/elements/1.1/\" xmlns:meta=\"urn:oasis:names:tc:opendocument:xmlns:meta:1.0\" xmlns:number=\"urn:oasis:names:tc:opendocument:xmlns:datastyle:1.0\" xmlns:svg=\"urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0\" xmlns:chart=\"urn:oasis:names:tc:opendocument:xmlns:chart:1.0\" xmlns:dr3d=\"urn:oasis:names:tc:opendocument:xmlns:dr3d:1.0\" xmlns:math=\"http://www.w3.org/1998/Math/MathML\" xmlns:form=\"urn:oasis:names:tc:opendocument:xmlns:form:1.0\" xmlns:script=\"urn:oasis:names:tc:opendocument:xmlns:script:1.0\" xmlns:ooo=\"http://openoffice.org/2004/office\" xmlns:ooow=\"http://openoffice.org/2004/writer\" xmlns:oooc=\"http://openoffice.org/2004/calc\" xmlns:dom=\"http://www.w3.org/2001/xml-events\" xmlns:xforms=\"http://www.w3.org/2002/xforms\" xmlns:xsd=\"http://www.w3.org/2001/XMLSchema\" xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\" office:version=\"1.3\"><office:body><office:text>";
const CONTENT_SUFFIX: &[u8] = b"</office:text></office:body></office:document-content>";

const DEFAULT_META_XML: &[u8] = b"<?xml version=\"1.0\" encoding=\"UTF-8\"?><office:document-meta xmlns:office=\"urn:oasis:names:tc:opendocument:xmlns:office:1.0\" xmlns:xlink=\"http://www.w3.org/1999/xlink\" xmlns:dc=\"http://purl.org/dc/elements/1.1/\" xmlns:meta=\"urn:oasis:names:tc:opendocument:xmlns:meta:1.0\" office:version=\"1.3\"><office:meta><meta:generator>Litchi/0.0.1</meta:generator></office:meta></office:document-meta>";

const DEFAULT_STYLES_XML: &[u8] = b"<?xml version=\"1.0\" encoding=\"UTF-8\"?><office:document-styles xmlns:office=\"urn:oasis:names:tc:opendocument:xmlns:office:1.0\" xmlns:style=\"urn:oasis:names:tc:opendocument:xmlns:style:1.0\" xmlns:text=\"urn:oasis:names:tc:opendocument:xmlns:text:1.0\" xmlns:table=\"urn:oasis:names:tc:opendocument:xmlns:table:1.0\" xmlns:draw=\"urn:oasis:names:tc:opendocument:xmlns:drawing:1.0\" xmlns:fo=\"urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0\" xmlns:xlink=\"http://www.w3.org/1999/xlink\" xmlns:dc=\"http://purl.org/dc/elements/1.1/\" xmlns:meta=\"urn:oasis:names:tc:opendocument:xmlns:meta:1.0\" xmlns:number=\"urn:oasis:names:tc:opendocument:xmlns:datastyle:1.0\" xmlns:svg=\"urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0\" xmlns:chart=\"urn:oasis:names:tc:opendocument:xmlns:chart:1.0\" xmlns:dr3d=\"urn:oasis:names:tc:opendocument:xmlns:dr3d:1.0\" xmlns:math=\"http://www.w3.org/1998/Math/MathML\" xmlns:form=\"urn:oasis:names:tc:opendocument:xmlns:form:1.0\" xmlns:script=\"urn:oasis:names:tc:opendocument:xmlns:script:1.0\" office:version=\"1.3\"><office:font-face-decls/><office:styles><!-- Numbered list style --><text:list-style style:name=\"L1\"><text:list-level-style-number text:level=\"1\" text:style-name=\"Numbering_20_Symbols\" style:num-format=\"1\"><style:list-level-properties text:list-level-position-and-space-mode=\"label-alignment\"><style:list-level-label-alignment text:label-followed-by=\"listtab\" text:list-tab-stop-position=\"1.27cm\" fo:text-indent=\"-0.635cm\" fo:margin-left=\"1.27cm\"/></style:list-level-properties></text:list-level-style-number><text:list-level-style-number text:level=\"2\" text:style-name=\"Numbering_20_Symbols\" style:num-format=\"1\"><style:list-level-properties text:list-level-position-and-space-mode=\"label-alignment\"><style:list-level-label-alignment text:label-followed-by=\"listtab\" text:list-tab-stop-position=\"1.905cm\" fo:text-indent=\"-0.635cm\" fo:margin-left=\"1.905cm\"/></style:list-level-properties></text:list-level-style-number><text:list-level-style-number text:level=\"3\" text:style-name=\"Numbering_20_Symbols\" style:num-format=\"1\"><style:list-level-properties text:list-level-position-and-space-mode=\"label-alignment\"><style:list-level-label-alignment text:label-followed-by=\"listtab\" text:list-tab-stop-position=\"2.54cm\" fo:text-indent=\"-0.635cm\" fo:margin-left=\"2.54cm\"/></style:list-level-properties></text:list-level-style-number></text:list-style></office:styles><office:automatic-styles/><office:master-styles/></office:document-styles>";

/// Provider-owned finite limits.  The defaults are narrower than the reader's
/// hard ceilings; callers may widen them only within those ceilings.
#[derive(Clone, Copy, Debug, Eq, PartialEq)]
pub struct StreamingLimits {
    max_paragraphs: usize,
    max_paragraph_text_bytes: usize,
    max_total_text_bytes: usize,
    max_paragraph_xml_bytes: usize,
    max_content_xml_bytes: usize,
    max_output_bytes: u64,
    xml_audit: GeneratedXmlLimits,
}

impl StreamingLimits {
    #[allow(clippy::too_many_arguments)]
    pub fn new(
        max_paragraphs: usize,
        max_paragraph_text_bytes: usize,
        max_total_text_bytes: usize,
        max_paragraph_xml_bytes: usize,
        max_content_xml_bytes: usize,
        max_output_bytes: u64,
        xml_audit: GeneratedXmlLimits,
    ) -> Result<Self, StreamingError> {
        if max_paragraphs == 0 || max_paragraphs > MAX_TEXT_BLOCKS {
            return Err(invalid(format!(
                "max_paragraphs must be in 1..={MAX_TEXT_BLOCKS}"
            )));
        }
        if max_paragraph_text_bytes == 0 || max_paragraph_text_bytes > MAX_TEXT_BYTES {
            return Err(invalid(format!(
                "max_paragraph_text_bytes must be in 1..={MAX_TEXT_BYTES}"
            )));
        }
        if max_total_text_bytes == 0 || max_total_text_bytes > MAX_TEXT_BYTES {
            return Err(invalid(format!(
                "max_total_text_bytes must be in 1..={MAX_TEXT_BYTES}"
            )));
        }
        if max_paragraph_xml_bytes == 0 || max_paragraph_xml_bytes > MAX_CONTENT_BYTES {
            return Err(invalid(
                "max_paragraph_xml_bytes must be finite and positive",
            ));
        }
        if max_content_xml_bytes == 0 || max_content_xml_bytes > MAX_CONTENT_BYTES {
            return Err(invalid(format!(
                "max_content_xml_bytes must be in 1..={MAX_CONTENT_BYTES}"
            )));
        }
        let minimum_content_bytes = CONTENT_PREFIX
            .len()
            .checked_add(CONTENT_SUFFIX.len())
            .ok_or_else(|| invalid("fixed ODT content shell length overflows usize"))?;
        if max_content_xml_bytes < minimum_content_bytes {
            return Err(invalid(format!(
                "max_content_xml_bytes must fit the empty content shell ({minimum_content_bytes} bytes)"
            )));
        }
        if max_paragraph_xml_bytes > max_content_xml_bytes {
            return Err(invalid(
                "max_paragraph_xml_bytes must not exceed max_content_xml_bytes",
            ));
        }
        if max_output_bytes < 22 {
            return Err(invalid(
                "max_output_bytes must fit the minimum empty ZIP end record",
            ));
        }
        if max_content_xml_bytes > xml_audit.max_bytes() {
            return Err(invalid(
                "max_content_xml_bytes must not exceed XML byte audit limit",
            ));
        }
        if max_total_text_bytes > xml_audit.max_text_bytes() {
            return Err(invalid(
                "max_total_text_bytes must not exceed XML text audit limit",
            ));
        }
        Ok(Self {
            max_paragraphs,
            max_paragraph_text_bytes,
            max_total_text_bytes,
            max_paragraph_xml_bytes,
            max_content_xml_bytes,
            max_output_bytes,
            xml_audit,
        })
    }

    #[must_use]
    pub const fn max_paragraphs(self) -> usize {
        self.max_paragraphs
    }

    #[must_use]
    pub const fn max_paragraph_text_bytes(self) -> usize {
        self.max_paragraph_text_bytes
    }

    #[must_use]
    pub const fn max_total_text_bytes(self) -> usize {
        self.max_total_text_bytes
    }

    #[must_use]
    pub const fn max_paragraph_xml_bytes(self) -> usize {
        self.max_paragraph_xml_bytes
    }

    #[must_use]
    pub const fn max_content_xml_bytes(self) -> usize {
        self.max_content_xml_bytes
    }

    #[must_use]
    pub const fn max_output_bytes(self) -> u64 {
        self.max_output_bytes
    }

    /// Lexical XML-audit profile applied to the generated content member.
    #[must_use]
    pub const fn xml_audit(self) -> GeneratedXmlLimits {
        self.xml_audit
    }

    /// Modeled retained provider memory; ZIP/auditor allocator peaks are not
    /// included in this reservation claim.
    pub fn required_memory_bytes(self) -> Result<u64, StreamingError> {
        let shell = CONTENT_PREFIX
            .len()
            .checked_add(CONTENT_SUFFIX.len())
            .ok_or_else(|| invalid("fixed ODT shell length overflows usize"))?;
        let shell_memory = shell
            .checked_mul(2)
            .ok_or_else(|| invalid("ODT shell memory length overflows usize"))?;
        let bytes = self
            .max_paragraph_xml_bytes
            .checked_add(shell_memory)
            .and_then(|size| size.checked_add(COMMON_METADATA_RESERVATION))
            .ok_or_else(|| invalid("ODT retained memory length overflows usize"))?;
        u64::try_from(bytes).map_err(|_| invalid("ODT retained memory length exceeds u64"))
    }

    fn content_xml_audit(self) -> Result<GeneratedXmlLimits, StreamingError> {
        // The composed content member gets its own byte ceiling.  Keeping the
        // caller's lexical/text/event limits avoids applying the content cap
        // to the fixed styles and metadata members, which are audited through
        // their ordinary package path.
        GeneratedXmlLimits::new(
            self.max_content_xml_bytes,
            self.xml_audit.max_depth(),
            self.xml_audit.max_events(),
            self.xml_audit.max_attributes(),
            self.xml_audit.max_token_bytes(),
            self.xml_audit.max_text_bytes(),
        )
        .map_err(|error| invalid(format!("invalid content XML audit limits: {error}")))
    }
}

impl Default for StreamingLimits {
    fn default() -> Self {
        Self {
            max_paragraphs: DEFAULT_MAX_PARAGRAPHS,
            max_paragraph_text_bytes: DEFAULT_MAX_PARAGRAPH_TEXT_BYTES,
            max_total_text_bytes: DEFAULT_MAX_TOTAL_TEXT_BYTES,
            max_paragraph_xml_bytes: DEFAULT_MAX_PARAGRAPH_XML_BYTES,
            max_content_xml_bytes: DEFAULT_MAX_CONTENT_XML_BYTES,
            max_output_bytes: DEFAULT_MAX_OUTPUT_BYTES,
            xml_audit: GeneratedXmlLimits::default(),
        }
    }
}

/// Counts returned after all provider buffers and reservations have dropped.
#[derive(Clone, Copy, Debug, Eq, PartialEq)]
pub struct ParagraphStreamReport {
    paragraphs: usize,
    input_text_bytes: usize,
    content_xml_bytes: usize,
}

impl ParagraphStreamReport {
    #[must_use]
    pub const fn paragraphs(self) -> usize {
        self.paragraphs
    }

    #[must_use]
    pub const fn input_text_bytes(self) -> usize {
        self.input_text_bytes
    }

    #[must_use]
    pub const fn content_xml_bytes(self) -> usize {
        self.content_xml_bytes
    }
}

#[derive(Clone, Copy, Debug, Eq, PartialEq)]
/// Format-neutral source category for an output failure.
pub enum PublicationFailureKind {
    /// Metadata or envelope admission failed before payload publication.
    Preflight,
    /// The paragraph iterator or provider fragment producer failed.
    Producer,
    /// The caller-owned sink rejected output bytes.
    Sink,
    /// ZIP framing or a member publication failed.
    Archive,
    /// A finite provider, XML, or archive limit was reached.
    Limit,
    /// Manifest or final ZIP end-record publication failed.
    Finalization,
}

#[derive(Debug)]
/// A publication failure with the accepted caller-sink prefix length.
pub struct PublicationError {
    written: u64,
    kind: PublicationFailureKind,
    source: Box<dyn StdError + Send + Sync + 'static>,
}

impl PublicationError {
    #[must_use]
    pub const fn written(&self) -> u64 {
        self.written
    }

    #[must_use]
    pub const fn kind(&self) -> PublicationFailureKind {
        self.kind
    }
}

impl fmt::Display for PublicationError {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        write!(
            f,
            "ODT paragraph publication failed after {} byte(s): {}",
            self.written, self.source
        )
    }
}

impl StdError for PublicationError {
    fn source(&self) -> Option<&(dyn StdError + 'static)> {
        Some(self.source.as_ref())
    }
}

#[derive(Debug)]
/// Error returned by the bounded ODT paragraph stream.
pub enum StreamingError {
    Invalid {
        error: Error,
        written: u64,
    },
    Execution {
        written: u64,
        error: ExecutionError,
    },
    Producer {
        written: u64,
        source: Box<dyn StdError + Send + Sync + 'static>,
    },
    LimitExceeded {
        resource: &'static str,
        observed: u64,
        limit: u64,
        written: u64,
    },
    CounterOverflow {
        resource: &'static str,
        written: u64,
    },
    Publication(PublicationError),
}

impl fmt::Display for StreamingError {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Self::Invalid { error, written } => {
                write!(
                    f,
                    "ODT paragraph input was invalid after {written} byte(s): {error}"
                )
            },
            Self::Execution { written, error } => {
                write!(
                    f,
                    "ODT paragraph stream stopped after {written} byte(s): {error}"
                )
            },
            Self::Producer { written, source } => {
                write!(
                    f,
                    "ODT paragraph source failed after {written} byte(s): {source}"
                )
            },
            Self::LimitExceeded {
                resource,
                observed,
                limit,
                written,
            } => write!(
                f,
                "ODT paragraph limit for {resource} exceeded: {observed} > {limit} after {written} byte(s)"
            ),
            Self::CounterOverflow { resource, written } => write!(
                f,
                "ODT paragraph counter {resource} overflowed after {written} byte(s)"
            ),
            Self::Publication(error) => error.fmt(f),
        }
    }
}

impl StdError for StreamingError {
    fn source(&self) -> Option<&(dyn StdError + 'static)> {
        match self {
            Self::Invalid { error, .. } => Some(error),
            Self::Execution { error, .. } => Some(error),
            Self::Producer { source, .. } => Some(source.as_ref()),
            Self::Publication(error) => Some(error),
            Self::LimitExceeded { .. } | Self::CounterOverflow { .. } => None,
        }
    }
}

fn invalid(message: impl Into<String>) -> StreamingError {
    StreamingError::Invalid {
        error: Error::InvalidFormat(message.into()),
        written: 0,
    }
}

/// Infallible source adapter. The iterator is consumed exactly once.
pub fn stream_plain_paragraphs_to<W, I, S>(
    output: &mut W,
    paragraphs: I,
    context: &ExecutionContext,
    limits: StreamingLimits,
) -> Result<ParagraphStreamReport, StreamingError>
where
    W: Write + ?Sized,
    I: IntoIterator<Item = S>,
    S: AsRef<str>,
{
    try_stream_plain_paragraphs_to(
        output,
        paragraphs.into_iter().map(Ok::<_, Error>),
        context,
        limits,
    )
}

/// Fallible source variant. Each item may be a borrowed `&str`, owned
/// `String`, or `Cow<str>` through `AsRef<str>`. A source failure is retained
/// as a typed producer error and is never confused with common ZIP or
/// caller-sink failure.
pub fn try_stream_plain_paragraphs_to<W, I, S>(
    output: &mut W,
    paragraphs: I,
    context: &ExecutionContext,
    limits: StreamingLimits,
) -> Result<ParagraphStreamReport, StreamingError>
where
    W: Write + ?Sized,
    I: IntoIterator<Item = litchi_core::Result<S>>,
    S: AsRef<str>,
{
    context.check().map_err(|error| execution(error, 0))?;
    let shell_bytes = CONTENT_PREFIX.len() + CONTENT_SUFFIX.len();
    if shell_bytes > limits.max_content_xml_bytes {
        return Err(LocalLimit::new(
            "content XML bytes",
            shell_bytes,
            limits.max_content_xml_bytes,
        )
        .as_error(0));
    }
    let _memory = context
        .reserve(Resource::Memory, limits.required_memory_bytes()?)
        .map_err(|error| execution(error, 0))?;
    let envelope = GeneratedXmlEnvelope::try_new(CONTENT_PREFIX, CONTENT_SUFFIX)
        .map_err(|error| StreamingError::Invalid { error, written: 0 })?;
    let content_xml_audit = limits.content_xml_audit()?;
    let content_limit = u64::try_from(limits.max_content_xml_bytes)
        .map_err(|_| invalid("content XML limit exceeds u64"))?;
    let fixed_uncompressed = u64::try_from(
        ODT_MIME
            .len()
            .checked_add(DEFAULT_META_XML.len())
            .and_then(|size| size.checked_add(DEFAULT_STYLES_XML.len()))
            .ok_or_else(|| invalid("fixed ODT member size overflow"))?,
    )
    .map_err(|_| invalid("fixed ODT member size exceeds u64"))?;
    let total_limit = content_limit
        .checked_add(fixed_uncompressed)
        .and_then(|size| size.checked_add(256 * 1024))
        .ok_or_else(|| invalid("ODT package total size overflow"))?;
    let fixed_entry_limit = ODT_MIME
        .len()
        .max(DEFAULT_META_XML.len())
        .max(DEFAULT_STYLES_XML.len());
    let fixed_entry_limit = u64::try_from(fixed_entry_limit.max(8 * 1024))
        .map_err(|_| invalid("fixed ODT entry limit exceeds u64"))?;
    let max_entry_limit = content_limit.max(fixed_entry_limit);
    let archive_limits = PackageWriterLimits::new(
        5,
        MAX_MEMBER_NAME_BYTES,
        (64 * 1024).min(limits.max_output_bytes),
    )
    .with_byte_limits(max_entry_limit, total_limit, limits.max_output_bytes)
    .with_compressed_size_limit(limits.max_output_bytes);

    let shell_work = CONTENT_PREFIX
        .len()
        .checked_add(CONTENT_SUFFIX.len())
        .ok_or_else(|| invalid("ODT content shell work length overflows usize"))?;
    let shell_work = u64::try_from(shell_work)
        .map_err(|_| invalid("ODT content shell work length exceeds u64"))?;
    if let Err(error) = context.consume(Resource::Work, shell_work) {
        return Err(execution(error, 0));
    }

    let mut budgeted = BudgetedOutput::new(output, context, limits.max_output_bytes);
    let mut package = PackageWriter::with_writer_and_limits(&mut budgeted, archive_limits);
    if let Err(error) = package.set_mimetype_streaming(ODT_MIME) {
        drop(package);
        return Err(map_package_error(
            error,
            PublicationFailureKind::Preflight,
            &budgeted,
        ));
    }

    let mut producer = ProducerState::new(paragraphs.into_iter(), context, limits);
    let generated = package.add_generated_xml(
        CONTENT_PATH,
        CONTENT_MEDIA_TYPE,
        envelope,
        content_xml_audit,
        limits.max_paragraph_xml_bytes,
        |fragment| producer.next_fragment(fragment),
    );
    let generated = match generated {
        Ok(report) => report,
        Err(error) => {
            drop(package);
            return Err(producer.map_error(error, &budgeted));
        },
    };

    // These static XML members use the existing Builder grammar. The common
    // authored-XML seam audits comments and preserves typed archive progress;
    // BudgetedOutput still checks context and retains accepted sink bytes.
    for (path, bytes) in [
        ("styles.xml", DEFAULT_STYLES_XML),
        ("meta.xml", DEFAULT_META_XML),
    ] {
        if let Err(error) = context.check() {
            drop(package);
            return Err(execution(error, budgeted.accepted()));
        }
        let work = match u64::try_from(bytes.len()) {
            Ok(work) => work,
            Err(_) => {
                drop(package);
                return Err(StreamingError::Invalid {
                    error: Error::InvalidFormat("static ODT XML length exceeds u64".to_string()),
                    written: budgeted.accepted(),
                });
            },
        };
        if let Err(error) = context.consume(Resource::Work, work) {
            drop(package);
            return Err(execution(error, budgeted.accepted()));
        }
        if let Err(error) = package.add_authored_xml(path, bytes, CONTENT_MEDIA_TYPE) {
            drop(package);
            return Err(map_package_error(
                error,
                PublicationFailureKind::Archive,
                &budgeted,
            ));
        }
    }

    let paragraphs = producer.paragraphs;
    let input_text_bytes = producer.input_text_bytes;
    let finish = package.finish_to_writer();
    if let Err(error) = finish {
        return Err(producer.map_error_kind(
            error,
            PublicationFailureKind::Finalization,
            &budgeted,
        ));
    }
    if let Err(error) = budgeted.final_check() {
        return Err(execution(error, budgeted.accepted()));
    }
    Ok(ParagraphStreamReport {
        paragraphs,
        input_text_bytes,
        content_xml_bytes: generated.bytes(),
    })
}

struct ProducerState<'context, I, S> {
    source: I,
    context: &'context ExecutionContext,
    limits: StreamingLimits,
    paragraphs: usize,
    input_text_bytes: usize,
    content_xml_bytes: usize,
    execution_error: Option<ExecutionError>,
    producer_error: Option<Error>,
    invalid_error: Option<String>,
    limit_error: Option<LocalLimit>,
    counter_error: Option<&'static str>,
    producer_failed: bool,
    _item: std::marker::PhantomData<fn() -> S>,
}

impl<'context, I, S> ProducerState<'context, I, S> {
    fn new(source: I, context: &'context ExecutionContext, limits: StreamingLimits) -> Self {
        Self {
            source,
            context,
            limits,
            paragraphs: 0,
            input_text_bytes: 0,
            content_xml_bytes: CONTENT_PREFIX.len() + CONTENT_SUFFIX.len(),
            execution_error: None,
            producer_error: None,
            invalid_error: None,
            limit_error: None,
            counter_error: None,
            producer_failed: false,
            _item: std::marker::PhantomData,
        }
    }
}

impl<'context, I, S> ProducerState<'context, I, S>
where
    I: Iterator<Item = litchi_core::Result<S>>,
    S: AsRef<str>,
{
    fn next_fragment(&mut self, output: &mut dyn Write) -> litchi_core::Result<bool> {
        if let Err(error) = self.context.check() {
            self.execution_error = Some(error.clone());
            return Err(Error::Other(error.to_string()));
        }
        if self.paragraphs == self.limits.max_paragraphs {
            if let Some(item) = self.source.next() {
                if let Err(error) = item {
                    self.producer_failed = true;
                    self.producer_error = Some(error);
                    return Err(Error::Other(
                        "paragraph source returned an error".to_string(),
                    ));
                }
                self.producer_failed = true;
                self.limit_error = Some(LocalLimit::new(
                    "paragraphs",
                    self.paragraphs.saturating_add(1),
                    self.limits.max_paragraphs,
                ));
                return Err(Error::InvalidFormat(
                    "paragraph iterator exceeds max_paragraphs".to_string(),
                ));
            }
            return Ok(false);
        }
        let Some(item) = self.source.next() else {
            return Ok(false);
        };
        let value = match item {
            Ok(value) => value,
            Err(error) => {
                self.producer_failed = true;
                self.producer_error = Some(error);
                return Err(Error::Other(
                    "paragraph source returned an error".to_string(),
                ));
            },
        };
        let value = value.as_ref();
        if value.len() > self.limits.max_paragraph_text_bytes {
            self.limit_error = Some(LocalLimit::new(
                "paragraph text bytes",
                value.len(),
                self.limits.max_paragraph_text_bytes,
            ));
            self.producer_failed = true;
            return Err(Error::InvalidFormat(
                "paragraph text exceeds max_paragraph_text_bytes".to_string(),
            ));
        }
        let total = match self.input_text_bytes.checked_add(value.len()) {
            Some(total) => total,
            None => {
                self.counter_error = Some("aggregate text bytes");
                self.producer_failed = true;
                return Err(Error::InvalidFormat(
                    "aggregate text counter overflow".to_string(),
                ));
            },
        };
        if total > self.limits.max_total_text_bytes {
            self.limit_error = Some(LocalLimit::new(
                "aggregate text bytes",
                total,
                self.limits.max_total_text_bytes,
            ));
            self.producer_failed = true;
            return Err(Error::InvalidFormat(
                "paragraph source exceeds max_total_text_bytes".to_string(),
            ));
        }
        let input_bytes = match u64::try_from(value.len()) {
            Ok(bytes) => bytes,
            Err(_) => {
                self.counter_error = Some("input text bytes");
                self.producer_failed = true;
                return Err(Error::InvalidFormat(
                    "paragraph input length exceeds u64".to_string(),
                ));
            },
        };
        if let Err(error) = self.context.consume(Resource::InputBytes, input_bytes) {
            self.execution_error = Some(error.clone());
            return Err(Error::Other(error.to_string()));
        }
        // This charge accounts for the attempted source item. The report
        // counters below advance only after its complete fragment succeeds.
        match validate_plain_text(value, self.context) {
            Ok(()) => {},
            Err(PlainTextFailure::Invalid(error)) => {
                self.invalid_error = Some(error.to_string());
                self.producer_failed = true;
                return Err(error);
            },
            Err(PlainTextFailure::Execution(error)) => {
                self.execution_error = Some(error.clone());
                return Err(Error::Other(error.to_string()));
            },
        }
        if let Err(error) = self.context.consume(Resource::Objects, 1) {
            self.execution_error = Some(error.clone());
            return Err(Error::Other(error.to_string()));
        }
        // Objects covers the admitted paragraph stage, while the public
        // report remains a successful-fragment count.
        let mut paragraph = ParagraphWriter::new(
            output,
            self.context,
            self.limits.max_paragraph_xml_bytes,
            self.content_xml_bytes,
            self.limits.max_content_xml_bytes,
        );
        let result = emit_paragraph(&mut paragraph, value);
        if let Some(error) = paragraph.execution_error.take() {
            self.execution_error = Some(error);
        }
        if let Some(error) = paragraph.limit_error.take() {
            self.limit_error = Some(error);
        }
        if let Err(error) = result {
            self.producer_failed = true;
            return Err(error);
        }
        self.input_text_bytes = total;
        self.content_xml_bytes = self
            .content_xml_bytes
            .checked_add(paragraph.bytes)
            .ok_or_else(|| Error::InvalidFormat("content XML counter overflow".to_string()))?;
        self.paragraphs = match self.paragraphs.checked_add(1) {
            Some(value) => value,
            None => {
                self.counter_error = Some("paragraphs");
                self.producer_failed = true;
                return Err(Error::InvalidFormat(
                    "paragraph counter overflow".to_string(),
                ));
            },
        };
        Ok(true)
    }

    fn map_error<W: ?Sized>(
        &mut self,
        error: PackageWriterError,
        output: &BudgetedOutput<'_, W>,
    ) -> StreamingError {
        if let Some(source) = self.producer_error.take() {
            return StreamingError::Producer {
                written: output.accepted(),
                source: Box::new(source),
            };
        }
        if let Some(error) = self.invalid_error.take() {
            return StreamingError::Invalid {
                error: Error::InvalidFormat(error),
                written: output.accepted(),
            };
        }
        if let Some(limit) = self.limit_error.take() {
            return limit.as_error(output.accepted());
        }
        if let Some(resource) = self.counter_error.take() {
            return StreamingError::CounterOverflow {
                resource,
                written: output.accepted(),
            };
        }
        if let Some(error) = self.execution_error.take() {
            return execution(error, output.accepted());
        }
        map_package_error(
            error,
            if self.producer_failed {
                PublicationFailureKind::Producer
            } else {
                PublicationFailureKind::Archive
            },
            output,
        )
    }

    fn map_error_kind<W: ?Sized>(
        &mut self,
        error: PackageWriterError,
        kind: PublicationFailureKind,
        output: &BudgetedOutput<'_, W>,
    ) -> StreamingError {
        if let Some(source) = self.producer_error.take() {
            return StreamingError::Producer {
                written: output.accepted(),
                source: Box::new(source),
            };
        }
        if let Some(error) = self.invalid_error.take() {
            return StreamingError::Invalid {
                error: Error::InvalidFormat(error),
                written: output.accepted(),
            };
        }
        if let Some(limit) = self.limit_error.take() {
            return limit.as_error(output.accepted());
        }
        if let Some(resource) = self.counter_error.take() {
            return StreamingError::CounterOverflow {
                resource,
                written: output.accepted(),
            };
        }
        if let Some(error) = self.execution_error.take() {
            return execution(error, output.accepted());
        }
        map_package_error(error, kind, output)
    }
}

#[derive(Clone, Copy)]
struct LocalLimit {
    resource: &'static str,
    observed: usize,
    limit: usize,
}

impl LocalLimit {
    const fn new(resource: &'static str, observed: usize, limit: usize) -> Self {
        Self {
            resource,
            observed,
            limit,
        }
    }

    fn as_error(self, written: u64) -> StreamingError {
        StreamingError::LimitExceeded {
            resource: self.resource,
            observed: u64::try_from(self.observed).unwrap_or(u64::MAX),
            limit: u64::try_from(self.limit).unwrap_or(u64::MAX),
            written,
        }
    }
}

struct ParagraphWriter<'a> {
    output: &'a mut dyn Write,
    context: &'a ExecutionContext,
    maximum: usize,
    content_before: usize,
    content_maximum: usize,
    bytes: usize,
    execution_error: Option<ExecutionError>,
    limit_error: Option<LocalLimit>,
}

impl<'a> ParagraphWriter<'a> {
    fn new(
        output: &'a mut dyn Write,
        context: &'a ExecutionContext,
        maximum: usize,
        content_before: usize,
        content_maximum: usize,
    ) -> Self {
        Self {
            output,
            context,
            maximum,
            content_before,
            content_maximum,
            bytes: 0,
            execution_error: None,
            limit_error: None,
        }
    }

    fn check(&mut self) -> litchi_core::Result<()> {
        if let Err(error) = self.context.check() {
            self.execution_error = Some(error.clone());
            return Err(Error::Other(error.to_string()));
        }
        Ok(())
    }

    fn write_bytes(&mut self, bytes: &[u8]) -> litchi_core::Result<()> {
        self.check()?;
        let next = self.bytes.checked_add(bytes.len()).ok_or_else(|| {
            Error::InvalidFormat("paragraph XML byte counter overflow".to_string())
        })?;
        let content_next = self
            .content_before
            .checked_add(next)
            .ok_or_else(|| Error::InvalidFormat("content XML byte counter overflow".to_string()))?;
        if content_next > self.content_maximum {
            self.limit_error = Some(LocalLimit::new(
                "content XML bytes",
                content_next,
                self.content_maximum,
            ));
            return Err(Error::InvalidFormat(
                "content XML exceeds max_content_xml_bytes".to_string(),
            ));
        }
        if next > self.maximum {
            self.limit_error = Some(LocalLimit::new("paragraph XML bytes", next, self.maximum));
            return Err(Error::InvalidFormat(
                "paragraph XML exceeds max_paragraph_xml_bytes".to_string(),
            ));
        }
        let amount = u64::try_from(bytes.len())
            .map_err(|_| Error::InvalidFormat("paragraph XML length exceeds u64".to_string()))?;
        if let Err(error) = self.context.consume(Resource::Work, amount) {
            self.execution_error = Some(error.clone());
            return Err(Error::Other(error.to_string()));
        }
        self.output.write_all(bytes)?;
        self.bytes = next;
        Ok(())
    }

    /// Attempts the borrowed-span path for one already-safe UTF-8 span.
    ///
    /// `Fallback` deliberately leaves all failure state untouched. The caller
    /// then uses the scalar path so a paragraph/content limit or aggregate
    /// `Resource::Work` refusal reports the same first failing scalar write
    /// and cumulative usage as the original encoder. Cancellation is returned
    /// directly because it is an execution-policy failure rather than a
    /// batching boundary. `write_all` still copies into the common generated
    /// XML staging window; this is a borrowed-span path, not a no-copy promise.
    fn write_plain_span(&mut self, bytes: &[u8]) -> litchi_core::Result<PlainSpanResult> {
        self.check()?;
        let Some(next) = self.bytes.checked_add(bytes.len()) else {
            return Ok(PlainSpanResult::Fallback);
        };
        let Some(content_next) = self.content_before.checked_add(next) else {
            return Ok(PlainSpanResult::Fallback);
        };
        if content_next > self.content_maximum || next > self.maximum {
            return Ok(PlainSpanResult::Fallback);
        }
        let Ok(amount) = u64::try_from(bytes.len()) else {
            return Ok(PlainSpanResult::Fallback);
        };
        if let Err(error) = self.context.consume(Resource::Work, amount) {
            if matches!(&error, ExecutionError::ResourceLimit(_)) {
                return Ok(PlainSpanResult::Fallback);
            }
            self.execution_error = Some(error.clone());
            return Err(Error::Other(error.to_string()));
        }
        self.output.write_all(bytes)?;
        self.bytes = next;
        Ok(PlainSpanResult::Written)
    }
}

#[derive(Clone, Copy, Debug, Eq, PartialEq)]
enum PlainSpanResult {
    Written,
    Fallback,
}

fn emit_paragraph(output: &mut ParagraphWriter<'_>, value: &str) -> litchi_core::Result<()> {
    if value.is_empty() {
        return output.write_bytes(b"<text:p/>");
    }
    output.write_bytes(b"<text:p>")?;
    let mut chars = value.char_indices().peekable();
    loop {
        output.check()?;
        let Some((start, character)) = chars.next() else {
            break;
        };
        match character {
            ' ' => {
                let mut count = 1usize;
                let mut end = start + 1;
                while let Some(&(next, ' ')) = chars.peek() {
                    output.check()?;
                    chars.next();
                    count = count.saturating_add(1);
                    end = next + 1;
                }
                let at_start = start == 0;
                let at_end = end == value.len();
                let adjacent_control = value[..start]
                    .chars()
                    .next_back()
                    .is_some_and(|ch| matches!(ch, '\t' | '\n'))
                    || value[end..]
                        .chars()
                        .next()
                        .is_some_and(|ch| matches!(ch, '\t' | '\n'));
                if count > 1 || at_start || at_end || adjacent_control {
                    write_space_controls(output, count)?;
                } else {
                    output.write_bytes(b" ")?;
                }
            },
            '\t' => output.write_bytes(b"<text:tab/>")?,
            '\n' => output.write_bytes(b"<text:line-break/>")?,
            '&' => output.write_bytes(b"&amp;")?,
            '<' => output.write_bytes(b"&lt;")?,
            '>' => output.write_bytes(b"&gt;")?,
            '"' => output.write_bytes(b"&quot;")?,
            '\'' => output.write_bytes(b"&apos;")?,
            character => {
                if !is_plain_text_character(character) {
                    let mut encoded = [0_u8; 4];
                    output.write_bytes(character.encode_utf8(&mut encoded).as_bytes())?;
                    continue;
                }
                let mut end = start + character.len_utf8();
                while let Some(&(next, next_character)) = chars.peek() {
                    if !is_plain_text_character(next_character) {
                        break;
                    }
                    let next_length = next_character.len_utf8();
                    let span_length = next - start;
                    if span_length > MAX_PLAIN_TEXT_SPAN_BYTES.saturating_sub(next_length) {
                        break;
                    }
                    output.check()?;
                    chars.next();
                    end = next + next_length;
                    if end - start == MAX_PLAIN_TEXT_SPAN_BYTES {
                        break;
                    }
                }
                write_plain_span_or_scalars(output, &value[start..end])?;
            },
        }
    }
    output.write_bytes(b"</text:p>")
}

fn write_plain_span_or_scalars(
    output: &mut ParagraphWriter<'_>,
    bytes: &str,
) -> litchi_core::Result<()> {
    match output.write_plain_span(bytes.as_bytes())? {
        PlainSpanResult::Written => Ok(()),
        PlainSpanResult::Fallback => write_plain_scalars(output, bytes),
    }
}

fn write_plain_scalars(output: &mut ParagraphWriter<'_>, value: &str) -> litchi_core::Result<()> {
    for character in value.chars() {
        let mut encoded = [0_u8; 4];
        output.write_bytes(character.encode_utf8(&mut encoded).as_bytes())?;
    }
    Ok(())
}

const fn is_plain_text_character(character: char) -> bool {
    !matches!(
        character,
        ' ' | '\t' | '\n' | '\r' | '&' | '<' | '>' | '"' | '\''
    )
}

fn write_space_controls(
    output: &mut ParagraphWriter<'_>,
    mut count: usize,
) -> litchi_core::Result<()> {
    while count != 0 {
        let chunk = count.min(MAX_SPACE_COUNT);
        let mut lexical = [0_u8; 20];
        let lexical = write_usize(&mut lexical, chunk);
        output.write_bytes(b"<text:s text:c=\"")?;
        output.write_bytes(lexical)?;
        output.write_bytes(b"\"/>")?;
        count -= chunk;
    }
    Ok(())
}

fn write_usize(buffer: &mut [u8; 20], mut value: usize) -> &[u8] {
    // Callers split runs at MAX_SPACE_COUNT, so this fixed buffer always has
    // room for the decimal spelling. Keeping the helper allocation-free also
    // keeps the retained paragraph window independent of space-run length.
    let mut cursor = buffer.len();
    if value == 0 {
        cursor -= 1;
        buffer[cursor] = b'0';
    } else {
        while value != 0 {
            cursor -= 1;
            buffer[cursor] = b'0' + (value % 10) as u8;
            value /= 10;
        }
    }
    &buffer[cursor..]
}

enum PlainTextFailure {
    Invalid(Error),
    Execution(ExecutionError),
}

fn validate_plain_text(
    value: &str,
    context: &ExecutionContext,
) -> std::result::Result<(), PlainTextFailure> {
    for character in value.chars() {
        if let Err(error) = context.check() {
            return Err(PlainTextFailure::Execution(error));
        }
        if character == '\r' {
            return Err(PlainTextFailure::Invalid(Error::InvalidFormat(
                "ODT plaintext stream refuses carriage return; choose an explicit line-ending policy".to_string(),
            )));
        }
        if !is_xml10_character(character) {
            return Err(PlainTextFailure::Invalid(Error::InvalidFormat(
                "ODT paragraph contains an XML 1.0-incompatible character".to_string(),
            )));
        }
    }
    Ok(())
}

const fn is_xml10_character(character: char) -> bool {
    matches!(character as u32, 0x09 | 0x0A | 0x0D | 0x20..=0xD7FF | 0xE000..=0xFFFD | 0x10000..=0x10FFFF)
}

struct BudgetedOutput<'a, W: ?Sized> {
    sink: &'a mut W,
    context: &'a ExecutionContext,
    maximum: u64,
    accepted: u64,
    execution_error: Option<ExecutionError>,
    sink_failed: bool,
    limit_failed: bool,
    limit_observed: Option<u64>,
}

impl<'a, W: ?Sized> BudgetedOutput<'a, W> {
    fn new(sink: &'a mut W, context: &'a ExecutionContext, maximum: u64) -> Self {
        Self {
            sink,
            context,
            maximum,
            accepted: 0,
            execution_error: None,
            sink_failed: false,
            limit_failed: false,
            limit_observed: None,
        }
    }

    fn accepted(&self) -> u64 {
        self.accepted
    }

    fn final_check(&mut self) -> Result<(), ExecutionError> {
        self.context
            .check()
            .inspect_err(|error| self.execution_error = Some(error.clone()))
    }
}

#[derive(Debug)]
struct BudgetedIoError(ExecutionError);

impl fmt::Display for BudgetedIoError {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        self.0.fmt(f)
    }
}
impl StdError for BudgetedIoError {}

impl<W: Write + ?Sized> Write for BudgetedOutput<'_, W> {
    fn write(&mut self, bytes: &[u8]) -> io::Result<usize> {
        if bytes.is_empty() {
            return Ok(0);
        }
        self.context.check().map_err(|error| {
            self.execution_error = Some(error.clone());
            io::Error::other(BudgetedIoError(error))
        })?;
        let amount =
            u64::try_from(bytes.len()).map_err(|_| io::Error::other("output write exceeds u64"))?;
        let next = self.accepted.checked_add(amount).ok_or_else(|| {
            self.limit_failed = true;
            self.limit_observed = Some(u64::MAX);
            io::Error::other("accepted output counter overflow")
        })?;
        if next > self.maximum {
            self.limit_failed = true;
            self.limit_observed = Some(next);
            return Err(io::Error::other("ODT output byte limit exceeded"));
        }
        let reservation = self
            .context
            .reserve(Resource::OutputBytes, amount)
            .map_err(|error| {
                self.execution_error = Some(error.clone());
                io::Error::other(BudgetedIoError(error))
            })?;
        match self.sink.write(bytes) {
            Ok(written) => {
                if written == 0 {
                    self.sink_failed = true;
                }
                let written_u64 = u64::try_from(written)
                    .map_err(|_| io::Error::other("sink count exceeds u64"))?;
                if !reservation.commit(written_u64) {
                    self.sink_failed = true;
                    return Err(io::Error::other("output reservation commit failed"));
                }
                self.accepted = self.accepted.checked_add(written_u64).ok_or_else(|| {
                    self.limit_failed = true;
                    io::Error::other("accepted output counter overflow")
                })?;
                Ok(written)
            },
            Err(error) => {
                self.sink_failed = true;
                drop(reservation);
                Err(error)
            },
        }
    }

    fn flush(&mut self) -> io::Result<()> {
        self.context.check().map_err(|error| {
            self.execution_error = Some(error.clone());
            io::Error::other(BudgetedIoError(error))
        })?;
        self.sink.flush().inspect_err(|_| self.sink_failed = true)
    }
}

fn execution(error: ExecutionError, written: u64) -> StreamingError {
    StreamingError::Execution { written, error }
}

fn map_package_error<W: ?Sized>(
    error: PackageWriterError,
    kind: PublicationFailureKind,
    output: &BudgetedOutput<'_, W>,
) -> StreamingError {
    if let Some(error) = output.execution_error.clone() {
        return execution(error, output.accepted());
    }
    if output.limit_failed {
        return StreamingError::LimitExceeded {
            resource: "output bytes",
            observed: output.limit_observed.unwrap_or(u64::MAX),
            limit: output.maximum,
            written: output.accepted(),
        };
    }
    if let Some(limit) = error.limit() {
        return StreamingError::LimitExceeded {
            resource: match limit.resource() {
                PackageWriterLimitResource::FileCount => "file count",
                PackageWriterLimitResource::MemberNameBytes => "member name bytes",
                PackageWriterLimitResource::MetadataBytes => "metadata bytes",
                PackageWriterLimitResource::CompressedSize => "compressed member size",
                PackageWriterLimitResource::EntrySize => "member bytes",
                PackageWriterLimitResource::TotalSize => "total member bytes",
                PackageWriterLimitResource::OutputBytes => "output bytes",
                PackageWriterLimitResource::Other => "archive resource",
            },
            observed: limit.actual(),
            limit: limit.maximum(),
            // The provider sink adapter is authoritative; a transport error's
            // written field may be stale or duplicate that accepted count.
            written: output.accepted(),
        };
    }
    if let Some(limit) = error.xml_limit() {
        use litchi_odf_common::GeneratedXmlLimitResource as XmlResource;
        return StreamingError::LimitExceeded {
            resource: match limit.resource() {
                XmlResource::Attributes => "XML attributes",
                XmlResource::Bytes => "XML bytes",
                XmlResource::Depth => "XML depth",
                XmlResource::Events => "XML events",
                XmlResource::TextBytes => "XML text bytes",
                XmlResource::TokenBytes => "XML token bytes",
                _ => "XML resource",
            },
            observed: limit.actual() as u64,
            limit: limit.maximum() as u64,
            written: output.accepted(),
        };
    }
    StreamingError::Publication(PublicationError {
        written: output.accepted(),
        kind: if output.sink_failed {
            PublicationFailureKind::Sink
        } else {
            kind
        },
        source: Box::new(error),
    })
}

#[cfg(test)]
#[path = "streaming_span_tests.rs"]
mod span_tests;
