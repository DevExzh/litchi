//! Bounded fresh ODP publication from an ordered plain titled-slide source.
//!
//! This provider accepts only a title and body string for each slide. It emits
//! the same no-transition page/frame geometry as the established Builder while
//! consuming the source once and retaining one reusable `draw:page` fragment.
//! Rich slide fields belong to the model authoring APIs and are not flattened.
//!
//! A failure can leave a partial caller sink. The returned error preserves its
//! cause and the number of bytes acknowledged by that sink; discard that
//! output.
//!
//! Execution `InputBytes` counts raw UTF-8 title/body bytes and `Objects`
//! counts admitted slides. `OutputBytes` counts accepted sink bytes. `Work`
//! counts the fixed content shell, emitted slide XML, and fixed styles/meta
//! XML; it excludes ZIP framing, the manifest, compression, and audit work.
//! `Memory` reserves the modeled retained provider window, not every allocator
//! allocation. XML depth uses [`XmlAuditLimits`]; execution `Depth` is not
//! charged by this sequential provider. Cancellation is cooperative, so an
//! already admitted bounded text span may finish before cancellation is seen.
//!
//! Raw text limits and report counts differ from escaped lexical XML counts.
//! The shared authored-XML audit can reject ambiguous spacing between adjacent
//! text controls, as it does for the Builder's identical text grammar.
//! The source may be polled once beyond `max_slides` to distinguish exact
//! exhaustion from excess input; no further items are requested after refusal.

use std::{
    error::Error as StdError,
    fmt,
    io::{self, Write},
    marker::PhantomData,
    result::Result,
};

use litchi_core::{Error, ExecutionContext, ExecutionError, Resource};
use litchi_odf_common::core::{
    GeneratedXmlEnvelope, GeneratedXmlLimits, PackageWriter, PackageWriterError,
    PackageWriterLimitResource, PackageWriterLimits,
};

const ODP_MIME: &str = "application/vnd.oasis.opendocument.presentation";
const CONTENT_PATH: &str = "content.xml";
const CONTENT_MEDIA_TYPE: &str = "text/xml";
const MAX_SLIDES: usize = 65_536;
const MAX_TEXT_BYTES: usize = 64 * 1024 * 1024;
const MAX_CONTENT_BYTES: usize = 256 * 1024 * 1024;
const DEFAULT_MAX_SLIDES: usize = MAX_SLIDES;
const DEFAULT_MAX_TITLE_TEXT_BYTES: usize = 1 << 20;
const DEFAULT_MAX_BODY_TEXT_BYTES: usize = 1 << 20;
// The common audit default has a 16 MiB aggregate character-data budget.
// Keep the default provider profile constructible through the checked path;
// callers can widen the audit profile when they also widen this input limit.
const DEFAULT_MAX_TOTAL_TEXT_BYTES: usize = 16 << 20;
const DEFAULT_MAX_SLIDE_XML_BYTES: usize = 4 << 20;
// GeneratedXmlLimits::default() permits a 32 MiB document. Keep this default
// profile reconstructible through StreamingLimits::new instead of relying on
// the unchecked Default implementation.
const DEFAULT_MAX_CONTENT_XML_BYTES: usize = 32 << 20;
const DEFAULT_MAX_OUTPUT_BYTES: u64 = 512 << 20;
const COMMON_METADATA_RESERVATION: usize = 64 * 1024;
const MAX_MEMBER_NAME_BYTES: u64 = 32;
const MAX_PLAIN_TEXT_SPAN_BYTES: usize = 256;

// This is the exact no-animation root emitted by Builder::generate_content_xml
// for a plain titled slide sequence. The prelude includes the balanced dp1
// style and the final open body/presentation path. It is passed to the opt-in
// common prelude constructor; the old two-part envelope contract remains
// strict and is not widened.
const CONTENT_PREFIX: &[u8] = b"<?xml version=\"1.0\" encoding=\"UTF-8\"?><office:document-content xmlns:office=\"urn:oasis:names:tc:opendocument:xmlns:office:1.0\" xmlns:style=\"urn:oasis:names:tc:opendocument:xmlns:style:1.0\" xmlns:text=\"urn:oasis:names:tc:opendocument:xmlns:text:1.0\" xmlns:table=\"urn:oasis:names:tc:opendocument:xmlns:table:1.0\" xmlns:draw=\"urn:oasis:names:tc:opendocument:xmlns:drawing:1.0\" xmlns:fo=\"urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0\" xmlns:xlink=\"http://www.w3.org/1999/xlink\" xmlns:dc=\"http://purl.org/dc/elements/1.1/\" xmlns:meta=\"urn:oasis:names:tc:opendocument:xmlns:meta:1.0\" xmlns:number=\"urn:oasis:names:tc:opendocument:xmlns:datastyle:1.0\" xmlns:presentation=\"urn:oasis:names:tc:opendocument:xmlns:presentation:1.0\" xmlns:anim=\"urn:oasis:names:tc:opendocument:xmlns:animation:1.0\" xmlns:smil=\"urn:oasis:names:tc:opendocument:xmlns:smil-compatible:1.0\" xmlns:svg=\"urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0\" xmlns:chart=\"urn:oasis:names:tc:opendocument:xmlns:chart:1.0\" xmlns:dr3d=\"urn:oasis:names:tc:opendocument:xmlns:dr3d:1.0\" xmlns:math=\"http://www.w3.org/1998/Math/MathML\" xmlns:form=\"urn:oasis:names:tc:opendocument:xmlns:form:1.0\" xmlns:script=\"urn:oasis:names:tc:opendocument:xmlns:script:1.0\" xmlns:ooo=\"http://openoffice.org/2004/office\" office:version=\"1.3\"><office:scripts/><office:font-face-decls/><office:automatic-styles><style:style style:name=\"dp1\" style:family=\"drawing-page\"><style:drawing-page-properties/></style:style></office:automatic-styles><office:body><office:presentation>";
const CONTENT_SUFFIX: &[u8] = b"</office:presentation></office:body></office:document-content>";

// These are the pinned default auxiliary parts emitted by the common Builder
// grammar for a fresh ODP. They are borrowed slices so the streaming path does
// not allocate a whole replacement XML member.
const DEFAULT_META_XML: &[u8] = b"<?xml version=\"1.0\" encoding=\"UTF-8\"?><office:document-meta xmlns:office=\"urn:oasis:names:tc:opendocument:xmlns:office:1.0\" xmlns:xlink=\"http://www.w3.org/1999/xlink\" xmlns:dc=\"http://purl.org/dc/elements/1.1/\" xmlns:meta=\"urn:oasis:names:tc:opendocument:xmlns:meta:1.0\" office:version=\"1.3\"><office:meta><meta:generator>Litchi/0.0.1</meta:generator></office:meta></office:document-meta>";

const DEFAULT_STYLES_XML: &[u8] = b"<?xml version=\"1.0\" encoding=\"UTF-8\"?><office:document-styles xmlns:office=\"urn:oasis:names:tc:opendocument:xmlns:office:1.0\" xmlns:style=\"urn:oasis:names:tc:opendocument:xmlns:style:1.0\" xmlns:text=\"urn:oasis:names:tc:opendocument:xmlns:text:1.0\" xmlns:table=\"urn:oasis:names:tc:opendocument:xmlns:table:1.0\" xmlns:draw=\"urn:oasis:names:tc:opendocument:xmlns:drawing:1.0\" xmlns:fo=\"urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0\" xmlns:xlink=\"http://www.w3.org/1999/xlink\" xmlns:dc=\"http://purl.org/dc/elements/1.1/\" xmlns:meta=\"urn:oasis:names:tc:opendocument:xmlns:meta:1.0\" xmlns:number=\"urn:oasis:names:tc:opendocument:xmlns:datastyle:1.0\" xmlns:svg=\"urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0\" xmlns:chart=\"urn:oasis:names:tc:opendocument:xmlns:chart:1.0\" xmlns:dr3d=\"urn:oasis:names:tc:opendocument:xmlns:dr3d:1.0\" xmlns:math=\"http://www.w3.org/1998/Math/MathML\" xmlns:form=\"urn:oasis:names:tc:opendocument:xmlns:form:1.0\" xmlns:script=\"urn:oasis:names:tc:opendocument:xmlns:script:1.0\" xmlns:ooo=\"http://openoffice.org/2004/office\" xmlns:ooow=\"http://openoffice.org/2004/writer\" xmlns:oooc=\"http://openoffice.org/2004/calc\" xmlns:dom=\"http://www.w3.org/2001/xml-events\" xmlns:rpt=\"http://openoffice.org/2005/report\" xmlns:of=\"urn:oasis:names:tc:opendocument:xmlns:of:1.2\" xmlns:xhtml=\"http://www.w3.org/1999/xhtml\" xmlns:grddl=\"http://www.w3.org/2003/g/data-view#\" xmlns:tableooo=\"http://openoffice.org/2009/table\" xmlns:calcext=\"urn:org:documentfoundation:names:experimental:calc:xmlns:calcext:1.0\" xmlns:loext=\"urn:org:documentfoundation:names:experimental:office:xmlns:loext:1.0\" xmlns:field=\"urn:openoffice:names:experimental:ooo-ms-interop:xmlns:field:1.0\" xmlns:formx=\"urn:openoffice:names:experimental:ooxml-odf-interop:xmlns:form:1.0\" xmlns:css3t=\"http://www.w3.org/TR/css3-text/\" office:version=\"1.3\"><office:font-face-decls/><office:styles/><office:automatic-styles/><office:master-styles/></office:document-styles>";

/// One plain source slide. The same `S` may be borrowed, owned, or a `Cow`.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct PlainSlide<S> {
    /// Visible title. `Some("")` retains the Builder's empty-title frame.
    pub title: Option<S>,
    /// Visible body text. An empty body omits the body frame.
    pub body: S,
}

impl<S> PlainSlide<S> {
    /// Construct one plain slide from an optional title and body.
    #[must_use]
    pub const fn new(title: Option<S>, body: S) -> Self {
        Self { title, body }
    }
}

/// Public XML-audit limits without exposing xml-minifier's concrete type.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct XmlAuditLimits {
    max_bytes: usize,
    max_depth: usize,
    max_events: usize,
    max_attributes: usize,
    max_token_bytes: usize,
    max_text_bytes: usize,
}

impl XmlAuditLimits {
    /// Construct a checked finite XML-audit profile.
    pub fn new(
        max_bytes: usize,
        max_depth: usize,
        max_events: usize,
        max_attributes: usize,
        max_token_bytes: usize,
        max_text_bytes: usize,
    ) -> Result<Self, StreamingError> {
        let limits = GeneratedXmlLimits::new(
            max_bytes,
            max_depth,
            max_events,
            max_attributes,
            max_token_bytes,
            max_text_bytes,
        )
        .map_err(|error| invalid(format!("invalid XML audit limits: {error}")))?;
        Ok(Self::from_common(limits))
    }

    #[must_use]
    pub const fn max_bytes(self) -> usize {
        self.max_bytes
    }

    #[must_use]
    pub const fn max_depth(self) -> usize {
        self.max_depth
    }

    #[must_use]
    pub const fn max_events(self) -> usize {
        self.max_events
    }

    #[must_use]
    pub const fn max_attributes(self) -> usize {
        self.max_attributes
    }

    #[must_use]
    pub const fn max_token_bytes(self) -> usize {
        self.max_token_bytes
    }

    #[must_use]
    pub const fn max_text_bytes(self) -> usize {
        self.max_text_bytes
    }

    fn from_common(limits: GeneratedXmlLimits) -> Self {
        Self {
            max_bytes: limits.max_bytes(),
            max_depth: limits.max_depth(),
            max_events: limits.max_events(),
            max_attributes: limits.max_attributes(),
            max_token_bytes: limits.max_token_bytes(),
            max_text_bytes: limits.max_text_bytes(),
        }
    }
}

impl Default for XmlAuditLimits {
    fn default() -> Self {
        Self::from_common(GeneratedXmlLimits::default())
    }
}

/// Provider-owned finite limits for one fresh plain-slide package.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct StreamingLimits {
    max_slides: usize,
    max_title_text_bytes: usize,
    max_body_text_bytes: usize,
    max_total_text_bytes: usize,
    max_slide_xml_bytes: usize,
    max_content_xml_bytes: usize,
    max_output_bytes: u64,
    xml_audit: XmlAuditLimits,
}

impl StreamingLimits {
    /// Construct a checked finite provider profile.
    ///
    /// The minimum output ceiling is only a ZIP lower bound. Construction
    /// does not guarantee that the required ODP members fit that ceiling.
    #[allow(clippy::too_many_arguments)]
    pub fn new(
        max_slides: usize,
        max_title_text_bytes: usize,
        max_body_text_bytes: usize,
        max_total_text_bytes: usize,
        max_slide_xml_bytes: usize,
        max_content_xml_bytes: usize,
        max_output_bytes: u64,
        xml_audit: XmlAuditLimits,
    ) -> Result<Self, StreamingError> {
        if max_slides == 0 || max_slides > MAX_SLIDES {
            return Err(invalid(format!("max_slides must be in 1..={MAX_SLIDES}")));
        }
        for (name, value) in [
            ("max_title_text_bytes", max_title_text_bytes),
            ("max_body_text_bytes", max_body_text_bytes),
            ("max_total_text_bytes", max_total_text_bytes),
        ] {
            if value == 0 || value > MAX_TEXT_BYTES {
                return Err(invalid(format!("{name} must be in 1..={MAX_TEXT_BYTES}")));
            }
        }
        if max_slide_xml_bytes == 0 || max_slide_xml_bytes > MAX_CONTENT_BYTES {
            return Err(invalid(format!(
                "max_slide_xml_bytes must be in 1..={MAX_CONTENT_BYTES}"
            )));
        }
        if max_content_xml_bytes == 0 || max_content_xml_bytes > MAX_CONTENT_BYTES {
            return Err(invalid(format!(
                "max_content_xml_bytes must be in 1..={MAX_CONTENT_BYTES}"
            )));
        }
        let shell = CONTENT_PREFIX
            .len()
            .checked_add(CONTENT_SUFFIX.len())
            .ok_or_else(|| invalid("fixed ODP content shell length overflows usize"))?;
        if max_content_xml_bytes < shell {
            return Err(invalid(format!(
                "max_content_xml_bytes must fit the fixed ODP shell ({shell} bytes)"
            )));
        }
        if max_slide_xml_bytes > max_content_xml_bytes {
            return Err(invalid(
                "max_slide_xml_bytes must not exceed max_content_xml_bytes",
            ));
        }
        if max_output_bytes < 22 {
            return Err(invalid(
                "max_output_bytes must fit the minimum empty ZIP end record",
            ));
        }
        if max_content_xml_bytes > xml_audit.max_bytes() {
            return Err(invalid(
                "max_content_xml_bytes must not exceed the XML byte audit limit",
            ));
        }
        if max_total_text_bytes > xml_audit.max_text_bytes() {
            return Err(invalid(
                "max_total_text_bytes must not exceed the XML aggregate text limit",
            ));
        }
        Ok(Self {
            max_slides,
            max_title_text_bytes,
            max_body_text_bytes,
            max_total_text_bytes,
            max_slide_xml_bytes,
            max_content_xml_bytes,
            max_output_bytes,
            xml_audit,
        })
    }

    #[must_use]
    pub const fn max_slides(self) -> usize {
        self.max_slides
    }

    #[must_use]
    pub const fn max_title_text_bytes(self) -> usize {
        self.max_title_text_bytes
    }

    #[must_use]
    pub const fn max_body_text_bytes(self) -> usize {
        self.max_body_text_bytes
    }

    #[must_use]
    pub const fn max_total_text_bytes(self) -> usize {
        self.max_total_text_bytes
    }

    #[must_use]
    pub const fn max_slide_xml_bytes(self) -> usize {
        self.max_slide_xml_bytes
    }

    #[must_use]
    pub const fn max_content_xml_bytes(self) -> usize {
        self.max_content_xml_bytes
    }

    #[must_use]
    pub const fn max_output_bytes(self) -> u64 {
        self.max_output_bytes
    }

    #[must_use]
    pub const fn xml_audit(self) -> XmlAuditLimits {
        self.xml_audit
    }

    /// Narrow the XML audit profile while retaining all provider ceilings.
    pub fn with_xml_audit_limits(self, xml_audit: XmlAuditLimits) -> Result<Self, StreamingError> {
        Self::new(
            self.max_slides,
            self.max_title_text_bytes,
            self.max_body_text_bytes,
            self.max_total_text_bytes,
            self.max_slide_xml_bytes,
            self.max_content_xml_bytes,
            self.max_output_bytes,
            xml_audit,
        )
    }

    /// Modeled retained provider memory. ZIP/auditor allocator peak is not
    /// included in this reservation claim.
    pub fn required_memory_bytes(self) -> Result<u64, StreamingError> {
        let shell = CONTENT_PREFIX
            .len()
            .checked_add(CONTENT_SUFFIX.len())
            .ok_or_else(|| invalid("fixed ODP shell length overflows usize"))?;
        let shell_memory = shell
            .checked_mul(2)
            .ok_or_else(|| invalid("ODP shell memory length overflows usize"))?;
        let bytes = self
            .max_slide_xml_bytes
            .checked_add(shell_memory)
            .and_then(|value| value.checked_add(COMMON_METADATA_RESERVATION))
            .ok_or_else(|| invalid("ODP retained memory length overflows usize"))?;
        u64::try_from(bytes).map_err(|_| invalid("ODP retained memory exceeds u64"))
    }

    fn content_xml_audit(self) -> Result<GeneratedXmlLimits, StreamingError> {
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
            max_slides: DEFAULT_MAX_SLIDES,
            max_title_text_bytes: DEFAULT_MAX_TITLE_TEXT_BYTES,
            max_body_text_bytes: DEFAULT_MAX_BODY_TEXT_BYTES,
            max_total_text_bytes: DEFAULT_MAX_TOTAL_TEXT_BYTES,
            max_slide_xml_bytes: DEFAULT_MAX_SLIDE_XML_BYTES,
            max_content_xml_bytes: DEFAULT_MAX_CONTENT_XML_BYTES,
            max_output_bytes: DEFAULT_MAX_OUTPUT_BYTES,
            xml_audit: XmlAuditLimits::default(),
        }
    }
}

/// Counts returned after all provider buffers and reservations have dropped.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct SlideStreamReport {
    slides: usize,
    title_count: usize,
    body_count: usize,
    title_text_bytes: usize,
    body_text_bytes: usize,
    content_xml_bytes: usize,
}

impl SlideStreamReport {
    #[must_use]
    pub const fn slides(self) -> usize {
        self.slides
    }

    #[must_use]
    pub const fn title_count(self) -> usize {
        self.title_count
    }

    #[must_use]
    pub const fn body_count(self) -> usize {
        self.body_count
    }

    #[must_use]
    pub const fn title_text_bytes(self) -> usize {
        self.title_text_bytes
    }

    #[must_use]
    pub const fn body_text_bytes(self) -> usize {
        self.body_text_bytes
    }

    /// Authored `content.xml` bytes before ZIP compression.
    #[must_use]
    pub const fn content_xml_bytes(self) -> usize {
        self.content_xml_bytes
    }
}

/// Format-neutral source/publication failure category.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum PublicationFailureKind {
    /// Metadata, envelope, or other pre-publication refusal.
    Preflight,
    /// The slide iterator or typed XML producer failed.
    Producer,
    /// The caller sink rejected bytes.
    Sink,
    /// ZIP framing or member publication failed.
    Archive,
    /// A finite provider, XML, or archive limit was reached.
    Limit,
    /// Final manifest or ZIP end-record publication failed.
    Finalization,
}

/// Publication failure with the accepted caller-sink prefix length.
#[derive(Debug)]
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
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        write!(
            formatter,
            "ODP slide publication failed after {} byte(s): {}",
            self.written, self.source
        )
    }
}

impl StdError for PublicationError {
    fn source(&self) -> Option<&(dyn StdError + 'static)> {
        Some(self.source.as_ref())
    }
}

/// Error returned by the bounded ODP plain-slide stream.
#[derive(Debug)]
pub enum StreamingError {
    /// Typed provider/XML input refusal.
    Invalid { error: Error, written: u64 },
    /// Cancellation or a shared execution budget stopped the stream.
    Execution { written: u64, error: ExecutionError },
    /// A fallible source returned its typed core error.
    Producer {
        written: u64,
        source: Box<dyn StdError + Send + Sync + 'static>,
    },
    /// A finite provider, XML, or ZIP ceiling rejected the next value.
    LimitExceeded {
        resource: &'static str,
        observed: u64,
        limit: u64,
        written: u64,
    },
    /// A checked progress counter could not represent its next value.
    CounterOverflow {
        resource: &'static str,
        written: u64,
    },
    /// ZIP/publication failure with accepted sink progress.
    Publication(PublicationError),
}

impl fmt::Display for StreamingError {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Self::Invalid { error, written } => write!(
                formatter,
                "ODP slide input was invalid after {written} byte(s): {error}"
            ),
            Self::Execution { written, error } => {
                write!(
                    formatter,
                    "ODP slide stream stopped after {written} byte(s): {error}"
                )
            },
            Self::Producer { written, source } => write!(
                formatter,
                "ODP slide source failed after {written} byte(s): {source}"
            ),
            Self::LimitExceeded {
                resource,
                observed,
                limit,
                written,
            } => write!(
                formatter,
                "ODP slide limit for {resource} exceeded: {observed} > {limit} after {written} byte(s)"
            ),
            Self::CounterOverflow { resource, written } => write!(
                formatter,
                "ODP slide counter {resource} overflowed after {written} byte(s)"
            ),
            Self::Publication(error) => error.fmt(formatter),
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
pub fn stream_plain_slides_to<W, I, S>(
    output: &mut W,
    slides: I,
    context: &ExecutionContext,
    limits: StreamingLimits,
) -> Result<SlideStreamReport, StreamingError>
where
    W: Write + ?Sized,
    I: IntoIterator<Item = PlainSlide<S>>,
    S: AsRef<str>,
{
    try_stream_plain_slides_to(
        output,
        slides.into_iter().map(Ok::<_, Error>),
        context,
        limits,
    )
}

/// Fallible source variant. A source error remains a typed producer failure,
/// distinct from XML, ZIP, sink, and execution failures.
pub fn try_stream_plain_slides_to<W, I, S>(
    output: &mut W,
    slides: I,
    context: &ExecutionContext,
    limits: StreamingLimits,
) -> Result<SlideStreamReport, StreamingError>
where
    W: Write + ?Sized,
    I: IntoIterator<Item = litchi_core::Result<PlainSlide<S>>>,
    S: AsRef<str>,
{
    context.check().map_err(|error| execution(error, 0))?;
    let memory = limits.required_memory_bytes()?;
    let _memory = context
        .reserve(Resource::Memory, memory)
        .map_err(|error| execution(error, 0))?;

    let envelope = GeneratedXmlEnvelope::try_new_with_prelude(CONTENT_PREFIX, CONTENT_SUFFIX)
        .map_err(|error| StreamingError::Invalid { error, written: 0 })?;
    let content_xml_audit = limits.content_xml_audit()?;
    let content_limit = u64::try_from(limits.max_content_xml_bytes)
        .map_err(|_| invalid("content XML limit exceeds u64"))?;
    let fixed_uncompressed = ODP_MIME
        .len()
        .checked_add(DEFAULT_STYLES_XML.len())
        .and_then(|value| value.checked_add(DEFAULT_META_XML.len()))
        .ok_or_else(|| invalid("fixed ODP member size overflow"))?;
    let fixed_uncompressed = u64::try_from(fixed_uncompressed)
        .map_err(|_| invalid("fixed ODP member size exceeds u64"))?;
    let total_limit = content_limit
        .checked_add(fixed_uncompressed)
        .and_then(|value| value.checked_add(256 * 1024))
        .ok_or_else(|| invalid("ODP package total size overflow"))?;
    let fixed_entry_limit = DEFAULT_STYLES_XML
        .len()
        .max(DEFAULT_META_XML.len())
        .max(8 * 1024);
    let fixed_entry_limit = u64::try_from(fixed_entry_limit)
        .map_err(|_| invalid("fixed ODP entry limit exceeds u64"))?;
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
        .ok_or_else(|| invalid("ODP content shell work length overflows usize"))?;
    let shell_work = u64::try_from(shell_work)
        .map_err(|_| invalid("ODP content shell work length exceeds u64"))?;
    if let Err(error) = context.consume(Resource::Work, shell_work) {
        return Err(execution(error, 0));
    }

    let mut budgeted = BudgetedOutput::new(output, context, limits.max_output_bytes);
    let mut package = PackageWriter::with_writer_and_limits(&mut budgeted, archive_limits);
    if let Err(error) = package.set_mimetype_streaming(ODP_MIME) {
        drop(package);
        return Err(map_package_error(
            error,
            PublicationFailureKind::Preflight,
            &budgeted,
        ));
    }

    let mut producer = ProducerState::new(slides.into_iter(), context, limits);
    let generated = match package.add_generated_xml(
        CONTENT_PATH,
        CONTENT_MEDIA_TYPE,
        envelope,
        content_xml_audit,
        limits.max_slide_xml_bytes,
        |fragment| producer.next_fragment(fragment),
    ) {
        Ok(report) => report,
        Err(error) => {
            drop(package);
            return Err(producer.map_error(error, &budgeted));
        },
    };

    // Keep the existing Builder member order: content, styles, meta. These
    // fixed members use the typed authored-XML seam and borrowed constants.
    for (path, bytes) in [
        ("styles.xml", DEFAULT_STYLES_XML),
        ("meta.xml", DEFAULT_META_XML),
    ] {
        if let Err(error) = context.check() {
            drop(package);
            return Err(execution(error, budgeted.accepted()));
        }
        let work = match u64::try_from(bytes.len()) {
            Ok(value) => value,
            Err(_) => {
                drop(package);
                return Err(StreamingError::Invalid {
                    error: Error::InvalidFormat("static ODP XML length exceeds u64".to_string()),
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

    let report = SlideStreamReport {
        slides: producer.slides,
        title_count: producer.title_count,
        body_count: producer.body_count,
        title_text_bytes: producer.title_text_bytes,
        body_text_bytes: producer.body_text_bytes,
        content_xml_bytes: generated.bytes(),
    };
    if let Err(error) = package.finish_to_writer() {
        return Err(producer.map_error_kind(
            error,
            PublicationFailureKind::Finalization,
            &budgeted,
        ));
    }
    if let Err(error) = budgeted.final_check() {
        return Err(execution(error, budgeted.accepted()));
    }
    Ok(report)
}

struct ProducerState<'context, I, S> {
    source: I,
    context: &'context ExecutionContext,
    limits: StreamingLimits,
    slides: usize,
    title_count: usize,
    body_count: usize,
    title_text_bytes: usize,
    body_text_bytes: usize,
    total_text_bytes: usize,
    content_xml_bytes: usize,
    execution_error: Option<ExecutionError>,
    producer_error: Option<Error>,
    invalid_error: Option<String>,
    limit_error: Option<LocalLimit>,
    counter_error: Option<&'static str>,
    producer_failed: bool,
    _item: PhantomData<fn() -> S>,
}

impl<'context, I, S> ProducerState<'context, I, S> {
    fn new(source: I, context: &'context ExecutionContext, limits: StreamingLimits) -> Self {
        Self {
            source,
            context,
            limits,
            slides: 0,
            title_count: 0,
            body_count: 0,
            title_text_bytes: 0,
            body_text_bytes: 0,
            total_text_bytes: 0,
            content_xml_bytes: CONTENT_PREFIX.len() + CONTENT_SUFFIX.len(),
            execution_error: None,
            producer_error: None,
            invalid_error: None,
            limit_error: None,
            counter_error: None,
            producer_failed: false,
            _item: PhantomData,
        }
    }
}

impl<'context, I, S> ProducerState<'context, I, S>
where
    I: Iterator<Item = litchi_core::Result<PlainSlide<S>>>,
    S: AsRef<str>,
{
    fn next_fragment(&mut self, output: &mut dyn Write) -> litchi_core::Result<bool> {
        if let Err(error) = self.context.check() {
            self.execution_error = Some(error.clone());
            return Err(Error::Other(error.to_string()));
        }
        if self.slides == self.limits.max_slides {
            match self.source.next() {
                None => return Ok(false),
                Some(Err(error)) => {
                    self.producer_failed = true;
                    self.producer_error = Some(error);
                    return Err(Error::Other("slide source returned an error".to_string()));
                },
                Some(Ok(_)) => {
                    self.producer_failed = true;
                    self.limit_error = Some(LocalLimit::new(
                        "slides",
                        self.slides.saturating_add(1),
                        self.limits.max_slides,
                    ));
                    return Err(Error::InvalidFormat(
                        "slide iterator exceeds max_slides".to_string(),
                    ));
                },
            }
        }

        let Some(item) = self.source.next() else {
            return Ok(false);
        };
        let item = match item {
            Ok(item) => item,
            Err(error) => {
                self.producer_failed = true;
                self.producer_error = Some(error);
                return Err(Error::Other("slide source returned an error".to_string()));
            },
        };
        let title = item.title.as_ref().map(AsRef::as_ref);
        let body = item.body.as_ref();
        let title_bytes = title.map_or(0, str::len);
        let body_bytes = body.len();
        if title_bytes > self.limits.max_title_text_bytes {
            self.limit_error = Some(LocalLimit::new(
                "title text bytes",
                title_bytes,
                self.limits.max_title_text_bytes,
            ));
            self.producer_failed = true;
            return Err(Error::InvalidFormat(
                "title text exceeds max_title_text_bytes".to_string(),
            ));
        }
        if body_bytes > self.limits.max_body_text_bytes {
            self.limit_error = Some(LocalLimit::new(
                "body text bytes",
                body_bytes,
                self.limits.max_body_text_bytes,
            ));
            self.producer_failed = true;
            return Err(Error::InvalidFormat(
                "body text exceeds max_body_text_bytes".to_string(),
            ));
        }
        let item_text = match title_bytes.checked_add(body_bytes) {
            Some(value) => value,
            None => {
                self.counter_error = Some("slide text bytes");
                self.producer_failed = true;
                return Err(Error::InvalidFormat(
                    "slide text counter overflow".to_string(),
                ));
            },
        };
        let total = match self.total_text_bytes.checked_add(item_text) {
            Some(value) => value,
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
                "slide source exceeds max_total_text_bytes".to_string(),
            ));
        }
        let input_bytes = u64::try_from(item_text).map_err(|_| {
            self.counter_error = Some("input text bytes");
            self.producer_failed = true;
            Error::InvalidFormat("slide input length exceeds u64".to_string())
        })?;
        if let Err(error) = self.context.consume(Resource::InputBytes, input_bytes) {
            self.execution_error = Some(error.clone());
            return Err(Error::Other(error.to_string()));
        }
        if let Some(title) = title {
            match validate_plain_text(title, self.context) {
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
        }
        match validate_plain_text(body, self.context) {
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

        let mut slide = SlideWriter::new(
            output,
            self.context,
            self.limits.max_slide_xml_bytes,
            self.content_xml_bytes,
            self.limits.max_content_xml_bytes,
        );
        let result = emit_slide(&mut slide, self.slides, title, body);
        if let Some(error) = slide.execution_error.take() {
            self.execution_error = Some(error);
        }
        if let Some(error) = slide.limit_error.take() {
            self.limit_error = Some(error);
        }
        if let Err(error) = result {
            self.producer_failed = true;
            return Err(error);
        }

        self.total_text_bytes = total;
        self.title_text_bytes =
            self.title_text_bytes
                .checked_add(title_bytes)
                .ok_or_else(|| {
                    self.counter_error = Some("title text bytes");
                    self.producer_failed = true;
                    Error::InvalidFormat("title text counter overflow".to_string())
                })?;
        self.body_text_bytes = self
            .body_text_bytes
            .checked_add(body_bytes)
            .ok_or_else(|| {
                self.counter_error = Some("body text bytes");
                self.producer_failed = true;
                Error::InvalidFormat("body text counter overflow".to_string())
            })?;
        self.content_xml_bytes =
            self.content_xml_bytes
                .checked_add(slide.bytes)
                .ok_or_else(|| {
                    self.counter_error = Some("content XML bytes");
                    self.producer_failed = true;
                    Error::InvalidFormat("content XML counter overflow".to_string())
                })?;
        self.slides = match self.slides.checked_add(1) {
            Some(value) => value,
            None => {
                self.counter_error = Some("slides");
                self.producer_failed = true;
                return Err(Error::InvalidFormat("slide counter overflow".to_string()));
            },
        };
        if title.is_some() {
            self.title_count = self.title_count.checked_add(1).ok_or_else(|| {
                self.counter_error = Some("title count");
                self.producer_failed = true;
                Error::InvalidFormat("title count overflow".to_string())
            })?;
        }
        if !body.is_empty() {
            self.body_count = self.body_count.checked_add(1).ok_or_else(|| {
                self.counter_error = Some("body count");
                self.producer_failed = true;
                Error::InvalidFormat("body count overflow".to_string())
            })?;
        }
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

struct SlideWriter<'a> {
    output: &'a mut dyn Write,
    context: &'a ExecutionContext,
    maximum: usize,
    content_before: usize,
    content_maximum: usize,
    bytes: usize,
    execution_error: Option<ExecutionError>,
    limit_error: Option<LocalLimit>,
}

impl<'a> SlideWriter<'a> {
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
            self.limit_error = Some(LocalLimit::new("slide XML bytes", usize::MAX, self.maximum));
            Error::InvalidFormat("slide XML byte counter overflow".to_string())
        })?;
        let content_next = self.content_before.checked_add(next).ok_or_else(|| {
            self.limit_error = Some(LocalLimit::new(
                "content XML bytes",
                usize::MAX,
                self.content_maximum,
            ));
            Error::InvalidFormat("content XML byte counter overflow".to_string())
        })?;
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
            self.limit_error = Some(LocalLimit::new("slide XML bytes", next, self.maximum));
            return Err(Error::InvalidFormat(
                "slide XML exceeds max_slide_xml_bytes".to_string(),
            ));
        }
        let amount = u64::try_from(bytes.len())
            .map_err(|_| Error::InvalidFormat("slide XML length exceeds u64".to_string()))?;
        if let Err(error) = self.context.consume(Resource::Work, amount) {
            self.execution_error = Some(error.clone());
            return Err(Error::Other(error.to_string()));
        }
        self.output.write_all(bytes)?;
        self.bytes = next;
        Ok(())
    }

    /// Attempt one borrowed ordinary-text span. A resource-limit Work refusal
    /// falls back to scalar writes so the first failing scalar and accepted
    /// prefix remain observable; cancellation and sink errors remain failures.
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

fn emit_slide(
    output: &mut SlideWriter<'_>,
    index: usize,
    title: Option<&str>,
    body: &str,
) -> litchi_core::Result<()> {
    output.write_bytes(b"<draw:page draw:name=\"page")?;
    let mut ordinal = [0_u8; 20];
    output.write_bytes(write_usize(
        &mut ordinal,
        index
            .checked_add(1)
            .ok_or_else(|| Error::InvalidFormat("slide page name index overflow".to_string()))?,
    ))?;
    output.write_bytes(b"\" draw:style-name=\"dp1\" draw:master-page-name=\"Default\">")?;

    if let Some(title) = title {
        output.write_bytes(
            br#"<draw:frame draw:style-name="gr1" draw:text-style-name="P1" draw:layer="layout" presentation:class="title" svg:width="25.199cm" svg:height="3.506cm" svg:x="1.4cm" svg:y="0.962cm"><draw:text-box>"#,
        )?;
        emit_text_paragraphs(output, title, "P1")?;
        output.write_bytes(b"</draw:text-box></draw:frame>")?;
    }

    if !body.is_empty() {
        let y = if title.is_some() { "5.0cm" } else { "2.0cm" };
        output.write_bytes(
            br#"<draw:frame draw:style-name="gr2" draw:text-style-name="P2" draw:layer="layout" presentation:class="object" svg:width="25.199cm" svg:height="10cm" svg:x="1.4cm" svg:y=""#,
        )?;
        output.write_bytes(y.as_bytes())?;
        output.write_bytes(b"\"><draw:text-box>")?;
        emit_text_paragraphs(output, body, "P2")?;
        output.write_bytes(b"</draw:text-box></draw:frame>")?;
    }

    output.write_bytes(b"</draw:page>")
}

fn emit_text_paragraphs(
    output: &mut SlideWriter<'_>,
    value: &str,
    style: &str,
) -> litchi_core::Result<()> {
    for paragraph in value.split('\n') {
        output.write_bytes(b"<text:p text:style-name=\"")?;
        output.write_bytes(style.as_bytes())?;
        output.write_bytes(b"\">")?;
        emit_text_content(output, paragraph)?;
        output.write_bytes(b"</text:p>")?;
    }
    Ok(())
}

// This deliberately follows authoring/builder/xml.rs rather than the more
// permissive reader text projection. In particular, one interior space remains
// literal only after prior output exists and when another scalar follows;
// leading/trailing spaces and runs use text:s.
fn emit_text_content(output: &mut SlideWriter<'_>, value: &str) -> litchi_core::Result<()> {
    let mut chars = value.char_indices().peekable();
    let mut emitted = false;
    while let Some((start, character)) = chars.next() {
        match character {
            ' ' => {
                let mut count = 1usize;
                let mut end = start + 1;
                while let Some(&(next, ' ')) = chars.peek() {
                    chars.next();
                    count = count.checked_add(1).ok_or_else(|| {
                        Error::InvalidFormat("space-run counter overflow".to_string())
                    })?;
                    end = next + 1;
                    output.check()?;
                }
                if count == 1 && emitted && chars.peek().is_some() {
                    output.write_bytes(b" ")?;
                } else if count == 1 {
                    output.write_bytes(b"<text:s/>")?;
                } else {
                    output.write_bytes(b"<text:s text:c=\"")?;
                    let mut lexical = [0_u8; 20];
                    output.write_bytes(write_usize(&mut lexical, count))?;
                    output.write_bytes(b"\"/>")?;
                }
                emitted = true;
                let _ = end;
            },
            '\t' => {
                output.write_bytes(b"<text:tab/>")?;
                emitted = true;
            },
            '\r' => {
                output.write_bytes(b"<text:line-break/>")?;
                emitted = true;
            },
            '&' => {
                output.write_bytes(b"&amp;")?;
                emitted = true;
            },
            '<' => {
                output.write_bytes(b"&lt;")?;
                emitted = true;
            },
            '>' => {
                output.write_bytes(b"&gt;")?;
                emitted = true;
            },
            '"' => {
                output.write_bytes(b"&quot;")?;
                emitted = true;
            },
            '\'' => {
                output.write_bytes(b"&apos;")?;
                emitted = true;
            },
            character => {
                if !is_plain_text_character(character) {
                    let mut encoded = [0_u8; 4];
                    output.write_bytes(character.encode_utf8(&mut encoded).as_bytes())?;
                    emitted = true;
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
                emitted = true;
            },
        }
    }
    Ok(())
}

fn write_plain_span_or_scalars(
    output: &mut SlideWriter<'_>,
    bytes: &str,
) -> litchi_core::Result<()> {
    match output.write_plain_span(bytes.as_bytes())? {
        PlainSpanResult::Written => Ok(()),
        PlainSpanResult::Fallback => write_plain_scalars(output, bytes),
    }
}

fn write_plain_scalars(output: &mut SlideWriter<'_>, value: &str) -> litchi_core::Result<()> {
    for character in value.chars() {
        let mut encoded = [0_u8; 4];
        output.write_bytes(character.encode_utf8(&mut encoded).as_bytes())?;
    }
    Ok(())
}

#[derive(Clone, Copy, Debug, Eq, PartialEq)]
enum PlainSpanResult {
    Written,
    Fallback,
}

const fn is_plain_text_character(character: char) -> bool {
    !matches!(
        character,
        ' ' | '\t' | '\n' | '\r' | '&' | '<' | '>' | '"' | '\''
    )
}

fn write_usize(buffer: &mut [u8; 20], mut value: usize) -> &[u8] {
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

fn validate_plain_text(
    value: &str,
    context: &ExecutionContext,
) -> std::result::Result<(), PlainTextFailure> {
    for character in value.chars() {
        if let Err(error) = context.check() {
            return Err(PlainTextFailure::Execution(error));
        }
        if !is_xml10_character(character) {
            return Err(PlainTextFailure::Invalid(Error::InvalidFormat(
                "ODP plain slide contains an XML 1.0-incompatible character".to_string(),
            )));
        }
    }
    Ok(())
}

const fn is_xml10_character(character: char) -> bool {
    matches!(
        character as u32,
        0x09 | 0x0A | 0x0D | 0x20..=0xD7FF | 0xE000..=0xFFFD | 0x10000..=0x10FFFF
    )
}

enum PlainTextFailure {
    Invalid(Error),
    Execution(ExecutionError),
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
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        self.0.fmt(formatter)
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
        let amount = u64::try_from(bytes.len())
            .map_err(|_| io::Error::other("ODP output write exceeds u64"))?;
        let next = self.accepted.checked_add(amount).ok_or_else(|| {
            self.limit_failed = true;
            self.limit_observed = Some(u64::MAX);
            io::Error::other("ODP accepted output counter overflow")
        })?;
        if next > self.maximum {
            self.limit_failed = true;
            self.limit_observed = Some(next);
            return Err(io::Error::other("ODP output byte limit exceeded"));
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
                    .map_err(|_| io::Error::other("ODP sink count exceeds u64"))?;
                if !reservation.commit(written_u64) {
                    self.sink_failed = true;
                    return Err(io::Error::other("ODP output reservation commit failed"));
                }
                self.accepted = self.accepted.checked_add(written_u64).ok_or_else(|| {
                    self.limit_failed = true;
                    io::Error::other("ODP accepted output counter overflow")
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
            observed: u64::try_from(limit.actual()).unwrap_or(u64::MAX),
            limit: u64::try_from(limit.maximum()).unwrap_or(u64::MAX),
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
