//! Bounded, callback-scoped markup-compatibility event processing.
//!
//! This module deliberately lives beside the legacy byte-buffer processor.  It
//! does not create a normalized XML buffer: every event is consumed from the
//! caller's `BufRead`, and observer values are valid only for the duration of
//! their callback.

use quick_xml::{
    Reader, XmlVersion,
    encoding::Decoder,
    events::{BytesEnd, BytesStart, Event},
};
use std::{
    borrow::Cow,
    collections::HashSet,
    fmt,
    io::{self, BufRead, Read},
    str,
    sync::Arc,
};

use super::model::{Capabilities, Error, Limits, NAMESPACE, Name, Report, XML_NS};
use crate::xml_name::{self, QualifiedName};

/// XML namespace used by namespace declaration attributes.
pub const XMLNS_NAMESPACE: &str = "http://www.w3.org/2000/xmlns/";

const DEFAULT_MAX_EVENTS: usize = 1_000_000;
const DEFAULT_MAX_EVENT_BYTES: usize = 8 * 1024 * 1024;
const DEFAULT_MAX_ATTRIBUTES_PER_EVENT: usize = 4096;
const DEFAULT_MAX_ATTRIBUTE_BYTES_PER_EVENT: usize = 4 * 1024 * 1024;
const DEFAULT_MAX_CONTEXT_BYTES: usize = 16 * 1024 * 1024;
const DEFAULT_MAX_NAME_BYTES: usize = 1024 * 1024;

const HARD_MAX_EVENTS: usize = 16 * 1024 * 1024;
const HARD_MAX_EVENT_BYTES: usize = 256 * 1024 * 1024;
const HARD_MAX_ATTRIBUTES_PER_EVENT: usize = 1 << 20;
const HARD_MAX_ATTRIBUTE_BYTES_PER_EVENT: usize = 256 * 1024 * 1024;
const HARD_MAX_CONTEXT_BYTES: usize = 512 * 1024 * 1024;
const HARD_MAX_NAME_BYTES: usize = 16 * 1024 * 1024;
const HARD_MAX_DEPTH: usize = 65_534;
const HARD_MAX_NAMESPACE_BINDINGS: usize = 1 << 20;
const HARD_MAX_DIRECTIVE_TOKENS: usize = 1 << 20;
const HARD_MAX_CHOICES: usize = 1 << 20;
const HARD_MAX_INPUT_BYTES: usize = 2 * 1024 * 1024 * 1024;
const HARD_MAX_OUTPUT_BYTES: usize = 2 * 1024 * 1024 * 1024;
const MAX_INTERRUPTED_RETRIES: usize = 8;

/// Limits for a streaming MCE operation.
///
/// `processing` is the existing MCE policy.  The other fields bound the
/// event-facing state owned by this module.  Bounds are intentionally finite;
/// public callers may construct this type directly, so the same checks are
/// repeated when processing starts.
#[derive(Debug, Clone)]
pub struct StreamLimits {
    /// Existing input, depth, namespace, directive, and choice limits.
    pub processing: Limits,
    /// Maximum number of non-EOF XML events.
    pub max_events: usize,
    /// Maximum raw bytes consumed for one event, including delimiters.  An
    /// optional UTF-8 BOM is treated as a separate three-byte preamble on the
    /// first read.
    pub max_event_bytes: usize,
    /// Maximum attributes on one start or empty event.
    pub max_attributes_per_event: usize,
    /// Maximum combined raw attribute-name and value bytes on one event.
    pub max_attribute_bytes_per_event: usize,
    /// Maximum estimated retained namespace/directive context bytes.
    pub max_context_bytes: usize,
    /// Maximum bytes in one qualified or expanded name.
    pub max_name_bytes: usize,
}

impl Default for StreamLimits {
    fn default() -> Self {
        Self {
            processing: Limits::default(),
            max_events: DEFAULT_MAX_EVENTS,
            max_event_bytes: DEFAULT_MAX_EVENT_BYTES,
            max_attributes_per_event: DEFAULT_MAX_ATTRIBUTES_PER_EVENT,
            max_attribute_bytes_per_event: DEFAULT_MAX_ATTRIBUTE_BYTES_PER_EVENT,
            max_context_bytes: DEFAULT_MAX_CONTEXT_BYTES,
            max_name_bytes: DEFAULT_MAX_NAME_BYTES,
        }
    }
}

impl StreamLimits {
    /// Construct default streaming bounds around an existing MCE policy.
    #[must_use]
    pub fn new(processing: Limits) -> Self {
        Self {
            processing,
            ..Self::default()
        }
    }

    /// Construct streaming limits after checking all event-facing bounds.
    ///
    /// The existing MCE limits are cloned into the returned policy.  Use the
    /// struct fields when a caller needs to change one bound after creation.
    pub fn try_new(
        processing: Limits,
        max_events: usize,
        max_event_bytes: usize,
        max_attributes_per_event: usize,
        max_attribute_bytes_per_event: usize,
        max_context_bytes: usize,
        max_name_bytes: usize,
    ) -> Result<Self, Error> {
        let value = Self {
            processing,
            max_events,
            max_event_bytes,
            max_attributes_per_event,
            max_attribute_bytes_per_event,
            max_context_bytes,
            max_name_bytes,
        };
        value.validate()?;
        Ok(value)
    }

    /// Construct default event bounds around an existing MCE policy.
    #[must_use]
    pub fn from_limits(processing: &Limits) -> Self {
        Self {
            processing: processing.clone(),
            ..Self::default()
        }
    }

    /// Validate finite, non-zero bounds and the hard ceilings.
    pub fn validate(&self) -> Result<(), Error> {
        check_bound(self.max_events, HARD_MAX_EVENTS, "stream event count")?;
        check_bound(
            self.max_event_bytes,
            HARD_MAX_EVENT_BYTES,
            "stream event bytes",
        )?;
        check_bound(
            self.max_attributes_per_event,
            HARD_MAX_ATTRIBUTES_PER_EVENT,
            "stream attributes per event",
        )?;
        check_bound(
            self.max_attribute_bytes_per_event,
            HARD_MAX_ATTRIBUTE_BYTES_PER_EVENT,
            "stream attribute bytes per event",
        )?;
        check_bound(
            self.max_context_bytes,
            HARD_MAX_CONTEXT_BYTES,
            "stream context bytes",
        )?;
        check_bound(
            self.max_name_bytes,
            HARD_MAX_NAME_BYTES,
            "stream name bytes",
        )?;
        if self.processing.max_input_bytes == 0 || self.processing.max_output_bytes == 0 {
            return Err(limit("MCE processing bounds must be non-zero"));
        }
        if self.processing.max_input_bytes > HARD_MAX_INPUT_BYTES
            || self.processing.max_output_bytes > HARD_MAX_OUTPUT_BYTES
        {
            return Err(limit("MCE processing byte bounds"));
        }
        check_bound(
            self.processing.max_namespace_bindings,
            HARD_MAX_NAMESPACE_BINDINGS,
            "MCE namespace bindings",
        )?;
        check_bound(
            self.processing.max_directive_tokens,
            HARD_MAX_DIRECTIVE_TOKENS,
            "MCE directive tokens",
        )?;
        check_bound(
            self.processing.max_choices_per_alternate,
            HARD_MAX_CHOICES,
            "MCE AlternateContent choices",
        )?;
        check_bound(
            self.processing.max_depth,
            HARD_MAX_DEPTH,
            "MCE processing depth",
        )?;
        Ok(())
    }
}

fn check_bound(value: usize, hard: usize, resource: &'static str) -> Result<(), Error> {
    if value == 0 {
        return Err(limit(resource));
    }
    if value > hard {
        return Err(limit(resource));
    }
    Ok(())
}

/// A raw start or empty element kind.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[non_exhaustive]
pub enum RawElementKind {
    /// A non-empty start tag.
    Start,
    /// An empty element tag.
    Empty,
}

/// One source-order raw attribute observed by the raw callback.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct RawAttribute<'a> {
    /// Qualified lexical attribute name in the callback-scoped event snapshot.
    pub qualified_name: Cow<'a, [u8]>,
    /// Expanded attribute name. Namespace declarations use
    /// [`XMLNS_NAMESPACE`].
    pub expanded_name: Name,
    /// Attribute value bytes exactly as encoded in the source event.
    pub raw_value: Cow<'a, [u8]>,
    /// XML-decoded and line-normalized attribute value.
    pub decoded_value: Cow<'a, str>,
}

impl<'a> RawAttribute<'a> {
    /// Return the lexical qualified name.
    #[must_use]
    pub fn name(&self) -> &[u8] {
        self.qualified_name.as_ref()
    }

    /// Return the raw source value.
    #[must_use]
    pub fn value(&self) -> &[u8] {
        self.raw_value.as_ref()
    }

    /// Return the decoded value.
    #[must_use]
    pub fn decoded(&self) -> &str {
        self.decoded_value.as_ref()
    }
}

/// A raw element delivered before MCE selection.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct RawElement<'a> {
    /// Whether the source event was a start or empty element.
    pub kind: RawElementKind,
    /// Qualified lexical element name in the callback-scoped event snapshot.
    pub qualified_name: Cow<'a, [u8]>,
    /// Namespace-expanded element name.
    pub expanded_name: Name,
    /// Every source-order attribute, including `xmlns` and `mc:*` attrs.
    pub attributes: Vec<RawAttribute<'a>>,
}

impl<'a> RawElement<'a> {
    /// Return the lexical qualified element name.
    #[must_use]
    pub fn name(&self) -> &[u8] {
        self.qualified_name.as_ref()
    }

    /// Return source-order attributes.
    #[must_use]
    pub fn attrs(&self) -> &[RawAttribute<'a>] {
        &self.attributes
    }
}

/// One attribute visible in the selected semantic stream.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct SemanticAttribute<'a> {
    /// Qualified lexical attribute name in the callback-scoped event snapshot.
    pub qualified_name: Cow<'a, [u8]>,
    /// Expanded attribute name.
    pub expanded_name: Name,
    /// Source value bytes.
    pub raw_value: Cow<'a, [u8]>,
    /// Decoded attribute value.
    pub decoded_value: Cow<'a, str>,
}

impl<'a> SemanticAttribute<'a> {
    /// Return the lexical name.
    #[must_use]
    pub fn name(&self) -> &[u8] {
        self.qualified_name.as_ref()
    }

    /// Return the decoded value.
    #[must_use]
    pub fn value(&self) -> &str {
        self.decoded_value.as_ref()
    }
}

/// A selected semantic start or empty element.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct SemanticElement<'a> {
    /// Qualified lexical element name in the callback-scoped event snapshot.
    pub qualified_name: Cow<'a, [u8]>,
    /// Namespace-expanded element name.
    pub expanded_name: Name,
    /// Attributes after MCE directive and ignorable filtering.
    pub attributes: Vec<SemanticAttribute<'a>>,
}

impl<'a> SemanticElement<'a> {
    /// Return the lexical name.
    #[must_use]
    pub fn name(&self) -> &[u8] {
        self.qualified_name.as_ref()
    }

    /// Return semantic attributes.
    #[must_use]
    pub fn attrs(&self) -> &[SemanticAttribute<'a>] {
        &self.attributes
    }
}

/// A selected semantic end element.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct SemanticEnd<'a> {
    /// Qualified lexical end name in the callback-scoped event snapshot.
    pub qualified_name: Cow<'a, [u8]>,
    /// Namespace-expanded name inherited from the corresponding start tag.
    pub expanded_name: Name,
}

impl<'a> SemanticEnd<'a> {
    /// Return the lexical name.
    #[must_use]
    pub fn name(&self) -> &[u8] {
        self.qualified_name.as_ref()
    }
}

/// Text-like data delivered to the semantic callback.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct SemanticText<'a> {
    /// Source bytes for the event content, excluding XML delimiters.
    pub raw: Cow<'a, [u8]>,
    /// Decoded source content.
    pub decoded: Cow<'a, str>,
}

impl<'a> SemanticText<'a> {
    /// Return decoded content.
    #[must_use]
    pub fn text(&self) -> &str {
        self.decoded.as_ref()
    }
}

/// XML declaration delivered to the semantic callback.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct SemanticDecl<'a> {
    /// Declaration content without `<?` and `?>` delimiters.
    pub raw: Cow<'a, [u8]>,
}

/// General entity or character reference delivered to the semantic callback.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct SemanticGeneralRef<'a> {
    /// Reference name without `&` and `;` delimiters.
    pub name: Cow<'a, [u8]>,
}

/// Counters collected while a streaming MCE operation validates and selects
/// the input.  The fields have the same meanings as [`Report`]; `events` is
/// the additional count of counted non-EOF parser events.  Rejected DTD and
/// processing-instruction events fail before entering this counter.  Unlike
/// the legacy byte processor, the streaming API also rejects a missing root,
/// non-whitespace text outside the root, late/repeated XML declarations, and
/// custom general references in hidden branches as XML conformance errors.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct StreamReport {
    /// Number of `mc:AlternateContent` containers encountered.
    pub alternate_content_count: usize,
    /// Number of selected `mc:Choice` branches.
    pub selected_choices: usize,
    /// Number of selected `mc:Fallback` branches.
    pub selected_fallbacks: usize,
    /// Number of ignored elements.
    pub ignored_elements: usize,
    /// Number of ignored attributes.
    pub ignored_attributes: usize,
    /// Number of preserved elements.
    pub preserved_elements: usize,
    /// Number of preserved attributes.
    pub preserved_attributes: usize,
    /// Number of unwrapped elements.
    pub unwrapped_elements: usize,
    /// Number of counted non-EOF XML events consumed.  Rejected DTD and
    /// processing-instruction events fail before entering this counter.
    pub events: usize,
}

impl From<StreamReport> for Report {
    fn from(report: StreamReport) -> Self {
        Self {
            alternate_content_count: report.alternate_content_count,
            selected_choices: report.selected_choices,
            selected_fallbacks: report.selected_fallbacks,
            ignored_elements: report.ignored_elements,
            ignored_attributes: report.ignored_attributes,
            preserved_elements: report.preserved_elements,
            preserved_attributes: report.preserved_attributes,
            unwrapped_elements: report.unwrapped_elements,
        }
    }
}

/// Selected semantic stream event.
#[derive(Debug, Clone, PartialEq, Eq)]
#[non_exhaustive]
pub enum SemanticEvent<'a> {
    /// Selected non-empty start tag.
    Start(SemanticElement<'a>),
    /// Selected empty element tag.
    Empty(SemanticElement<'a>),
    /// Selected end tag.
    End(SemanticEnd<'a>),
    /// Escaped text.
    Text(SemanticText<'a>),
    /// CDATA content.
    CData(SemanticText<'a>),
    /// Comment content.
    Comment(SemanticText<'a>),
    /// XML declaration.
    Decl(SemanticDecl<'a>),
    /// Predefined, numeric, or otherwise source-visible general reference.
    GeneralRef(SemanticGeneralRef<'a>),
}

/// A raw input event exceeded the configured per-event byte bound.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct EventLimitExceeded {
    /// Configured maximum for the event.
    pub limit: usize,
}

impl fmt::Display for EventLimitExceeded {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        write!(f, "XML event exceeds {} bytes", self.limit)
    }
}

impl std::error::Error for EventLimitExceeded {}

/// A source-wide input byte bound was reached before EOF.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct InputLimitExceeded {
    /// Configured maximum input bytes.
    pub limit: usize,
}

impl fmt::Display for InputLimitExceeded {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        write!(f, "XML input exceeds {} bytes", self.limit)
    }
}

impl std::error::Error for InputLimitExceeded {}

/// Invalid use of the private event meter's `consume` contract.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
struct InvalidMeterConsume;

impl fmt::Display for InvalidMeterConsume {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        f.write_str("XML event meter received an invalid consume count")
    }
}

impl std::error::Error for InvalidMeterConsume {}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
struct InterruptedRetriesExceeded;

impl fmt::Display for InterruptedRetriesExceeded {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        f.write_str("XML input exceeded the Interrupted retry bound")
    }
}

impl std::error::Error for InterruptedRetriesExceeded {}

const INTERRUPTED_RETRY_BUFFER_SIZE: usize = 8 * 1024;

struct InterruptedRetryReader<'a> {
    input: &'a mut dyn BufRead,
    buffer: [u8; INTERRUPTED_RETRY_BUFFER_SIZE],
    position: usize,
    length: usize,
    invalid_consume: bool,
}

impl<'a> InterruptedRetryReader<'a> {
    fn new(input: &'a mut dyn BufRead) -> Self {
        Self {
            input,
            buffer: [0; INTERRUPTED_RETRY_BUFFER_SIZE],
            position: 0,
            length: 0,
            invalid_consume: false,
        }
    }
}

impl Read for InterruptedRetryReader<'_> {
    fn read(&mut self, output: &mut [u8]) -> io::Result<usize> {
        if output.is_empty() {
            return Ok(0);
        }
        let available = self.fill_buf()?;
        let amount = available.len().min(output.len());
        output[..amount].copy_from_slice(&available[..amount]);
        self.consume(amount);
        Ok(amount)
    }
}

impl BufRead for InterruptedRetryReader<'_> {
    fn fill_buf(&mut self) -> io::Result<&[u8]> {
        if self.invalid_consume {
            return Err(io::Error::new(
                io::ErrorKind::InvalidData,
                InvalidMeterConsume,
            ));
        }
        if self.position == self.length {
            self.position = 0;
            self.length = 0;
            let mut interrupted = 0;
            loop {
                match self.input.read(&mut self.buffer) {
                    Ok(length) => {
                        self.length = length;
                        break;
                    },
                    Err(error)
                        if error.kind() == io::ErrorKind::Interrupted
                            && interrupted < MAX_INTERRUPTED_RETRIES =>
                    {
                        interrupted += 1;
                    },
                    Err(error) if error.kind() == io::ErrorKind::Interrupted => {
                        return Err(io::Error::new(
                            io::ErrorKind::InvalidData,
                            InterruptedRetriesExceeded,
                        ));
                    },
                    Err(error) => return Err(error),
                }
            }
        }
        Ok(&self.buffer[self.position..self.length])
    }

    fn consume(&mut self, amount: usize) {
        if amount > self.length.saturating_sub(self.position) {
            self.invalid_consume = true;
            return;
        }
        self.position += amount;
    }
}

/// Failure of a streaming operation, retaining observer failures as
/// secondary diagnostics.
#[derive(Debug)]
#[non_exhaustive]
pub enum StreamError<RawE, ActiveE> {
    /// The input reader or event meter failed. Callback errors, if any, are
    /// retained in the same value. During raw-only recovery this is primary
    /// over the retained MCE error, which is available through
    /// [`Self::prior_mce_error`].
    Input {
        /// Primary input failure.
        error: io::Error,
        /// Earlier semantic MCE failure retained when recovery later found
        /// an input failure.
        prior_mce_error: Option<Error>,
        /// First raw observer failure.
        raw_error: Option<RawE>,
        /// First active observer failure.
        active_error: Option<ActiveE>,
    },
    /// XML or MCE validation failed. Callback errors, if any, are retained.
    Mce {
        /// Primary MCE/XML failure.
        error: Error,
        /// Earlier semantic MCE failure retained when a later raw structural
        /// failure was promoted to primary.
        prior_mce_error: Option<Error>,
        /// First raw observer failure.
        raw_error: Option<RawE>,
        /// First active observer failure.
        active_error: Option<ActiveE>,
    },
    /// Both observers failed, but parsing and validation reached EOF without a
    /// higher-priority input or MCE failure.
    Callback {
        /// First raw observer failure, if any.
        raw_error: Option<RawE>,
        /// First active observer failure, if any.
        active_error: Option<ActiveE>,
    },
}

impl<RawE, ActiveE> StreamError<RawE, ActiveE> {
    /// Return the retained raw observer error.
    #[must_use]
    pub fn raw_callback_error(&self) -> Option<&RawE> {
        match self {
            Self::Input { raw_error, .. }
            | Self::Mce { raw_error, .. }
            | Self::Callback { raw_error, .. } => raw_error.as_ref(),
        }
    }

    /// Return the retained active observer error.
    #[must_use]
    pub fn active_callback_error(&self) -> Option<&ActiveE> {
        match self {
            Self::Input { active_error, .. }
            | Self::Mce { active_error, .. }
            | Self::Callback { active_error, .. } => active_error.as_ref(),
        }
    }

    /// Return an earlier MCE failure retained during raw-only recovery.
    #[must_use]
    pub fn prior_mce_error(&self) -> Option<&Error> {
        match self {
            Self::Input {
                prior_mce_error, ..
            }
            | Self::Mce {
                prior_mce_error, ..
            } => prior_mce_error.as_ref(),
            Self::Callback { .. } => None,
        }
    }

    /// Alias for [`Self::raw_callback_error`].
    #[must_use]
    pub fn raw_error(&self) -> Option<&RawE> {
        self.raw_callback_error()
    }

    /// Alias for [`Self::active_callback_error`].
    #[must_use]
    pub fn active_error(&self) -> Option<&ActiveE> {
        self.active_callback_error()
    }

    /// Return the primary MCE error, if this is an MCE failure.
    #[must_use]
    pub fn mce_error(&self) -> Option<&Error> {
        match self {
            Self::Mce { error, .. } => Some(error),
            Self::Input { .. } | Self::Callback { .. } => None,
        }
    }

    /// Return the primary input error, if this is an input failure.
    #[must_use]
    pub fn input_error(&self) -> Option<&io::Error> {
        match self {
            Self::Input { error, .. } => Some(error),
            Self::Mce { .. } | Self::Callback { .. } => None,
        }
    }
}

impl<RawE: fmt::Display, ActiveE: fmt::Display> fmt::Display for StreamError<RawE, ActiveE> {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Self::Input {
                error,
                prior_mce_error,
                raw_error,
                active_error,
            } => write_primary_with_callbacks(
                f,
                "MCE stream input error",
                error,
                prior_mce_error.as_ref(),
                raw_error.as_ref(),
                active_error.as_ref(),
            ),
            Self::Mce {
                error,
                prior_mce_error,
                raw_error,
                active_error,
            } => write_primary_with_callbacks(
                f,
                "MCE stream validation error",
                error,
                prior_mce_error.as_ref(),
                raw_error.as_ref(),
                active_error.as_ref(),
            ),
            Self::Callback {
                raw_error,
                active_error,
            } => {
                f.write_str("MCE stream observer failed")?;
                if let Some(error) = raw_error {
                    write!(f, " (raw observer: {error})")?;
                }
                if let Some(error) = active_error {
                    write!(f, " (active observer: {error})")?;
                }
                Ok(())
            },
        }
    }
}

fn write_primary_with_callbacks<T: fmt::Display, RawE: fmt::Display, ActiveE: fmt::Display>(
    f: &mut fmt::Formatter<'_>,
    label: &str,
    primary: &T,
    prior_mce_error: Option<&Error>,
    raw_error: Option<&RawE>,
    active_error: Option<&ActiveE>,
) -> fmt::Result {
    write!(f, "{label}: {primary}")?;
    if let Some(error) = prior_mce_error {
        write!(f, " (prior MCE: {})", prior_mce_kind(error))?;
    }
    if let Some(error) = raw_error {
        write!(f, " (raw observer: {error})")?;
    }
    if let Some(error) = active_error {
        write!(f, " (active observer: {error})")?;
    }
    Ok(())
}

fn prior_mce_kind(error: &Error) -> &'static str {
    match error {
        Error::NonConformant(_) => "non-conformant",
        Error::MustUnderstand(_) => "must-understand",
        Error::LimitExceeded(_) => "limit-exceeded",
        Error::Allocation { .. } => "allocation",
        Error::Xml(_) => "xml",
    }
}

impl<RawE: std::error::Error + 'static, ActiveE: std::error::Error + 'static> std::error::Error
    for StreamError<RawE, ActiveE>
{
    fn source(&self) -> Option<&(dyn std::error::Error + 'static)> {
        match self {
            Self::Input { error, .. } => Some(error),
            Self::Mce { error, .. } => Some(error),
            Self::Callback { .. } => None,
        }
    }
}

/// Process MCE and invoke an active observer for selected semantic events.
///
/// The callback receives a borrowed event whose lifetime is tied to the call;
/// it cannot retain the event or any source slices without explicitly copying
/// them.  Raw source observations are disabled through an internal no-op
/// callback.
pub fn process_markup_compatibility_stream<ActiveE, Active>(
    input: &mut dyn BufRead,
    capabilities: &Capabilities,
    limits: &StreamLimits,
    active: Active,
) -> Result<StreamReport, StreamError<std::convert::Infallible, ActiveE>>
where
    Active: for<'a> FnMut(SemanticEvent<'a>) -> Result<(), ActiveE>,
{
    process_markup_compatibility_stream_with_observers(
        input,
        capabilities,
        limits,
        |_| Ok::<(), std::convert::Infallible>(()),
        active,
    )
}

/// Process MCE while observing both every raw element and the selected
/// semantic stream.
///
/// Raw observers run after namespace resolution and before MCE selection.
/// Active observers run only for visible selected events.  Each observer is
/// independently disabled after its first error; parsing and validation still
/// continue to EOF so a later XML/MCE error remains the primary failure.
/// After a recoverable semantic MCE error, active processing is disabled and
/// the same input is drained through the bounded raw validator so raw
/// observers still see subsequent elements.  If that drain encounters a
/// transport I/O error, [`StreamError::Input`] is primary and retains the
/// first semantic failure in [`StreamError::prior_mce_error`].
/// Recovery is limited to `NonConformant` and `MustUnderstand` failures after
/// the current raw event has been committed; XML/tokenization, allocation,
/// limit, and input failures terminate without draining.
/// The streaming conformance layer additionally requires one document root,
/// rejects non-whitespace text outside that root, and checks end-tag names;
/// callers must not assume byte-for-byte parity with the legacy normalizer for
/// malformed documents.  It also permits at most one XML declaration, only as
/// the first XML event after an optional BOM, and validates general references
/// in hidden MCE branches; hidden custom references are rejected rather than
/// skipped as they are by the legacy processor.
/// CDATA is rejected outside the document root even when it contains only
/// whitespace; ordinary whitespace text and comments remain allowed there.
/// The configured bounds cover this module's event and retained-context
/// allocations, not allocations internal to `quick_xml` or observer code.
pub fn process_markup_compatibility_stream_with_observers<RawE, ActiveE, Raw, Active>(
    input: &mut dyn BufRead,
    capabilities: &Capabilities,
    limits: &StreamLimits,
    mut raw: Raw,
    mut active: Active,
) -> Result<StreamReport, StreamError<RawE, ActiveE>>
where
    Raw: for<'a> FnMut(RawElement<'a>) -> Result<(), RawE>,
    Active: for<'a> FnMut(SemanticEvent<'a>) -> Result<(), ActiveE>,
{
    limits.validate().map_err(|error| StreamError::Mce {
        error,
        prior_mce_error: None,
        raw_error: None,
        active_error: None,
    })?;

    let mut retry_reader = InterruptedRetryReader::new(input);
    let mut prefix_reader = PrefixReader::new(&mut retry_reader);
    if let Err(error) = prefix_reader.initialize() {
        if error
            .get_ref()
            .and_then(|source| source.downcast_ref::<InterruptedRetriesExceeded>())
            .is_some()
        {
            return Err(StreamError::Mce {
                error: limit("input Interrupted retries"),
                prior_mce_error: None,
                raw_error: None,
                active_error: None,
            });
        }
        return Err(StreamError::Input {
            error,
            prior_mce_error: None,
            raw_error: None,
            active_error: None,
        });
    }
    let prefix_bytes = prefix_reader.consumed_bytes();
    if prefix_bytes > limits.processing.max_input_bytes {
        return Err(StreamError::Mce {
            error: limit("input bytes"),
            prior_mce_error: None,
            raw_error: None,
            active_error: None,
        });
    }
    let replay_bytes = prefix_reader.replay_bytes();
    let mut reader = Reader::from_reader(EventMeter::new(
        &mut prefix_reader,
        limits.max_event_bytes,
        limits.processing.max_input_bytes,
        prefix_bytes,
        replay_bytes,
    ));
    reader.config_mut().trim_text(false);
    reader.config_mut().check_end_names = true;
    reader.config_mut().check_comments = true;

    let mut event_buffer = Vec::new();
    event_buffer
        .try_reserve_exact(limits.max_event_bytes)
        .map_err(|source| StreamError::Mce {
            error: Error::Allocation {
                resource: "MCE stream event buffer",
                source,
            },
            prior_mce_error: None,
            raw_error: None,
            active_error: None,
        })?;

    let mut state = Processor::new(capabilities, limits);
    let mut raw_error = None;
    let mut active_error = None;
    let mut raw_enabled = true;
    let mut active_enabled = true;
    let mut recovering = false;
    let mut recovery_error = None;

    loop {
        state.raw_state_valid = false;
        reader.get_mut().begin_event();
        let decoder = reader.decoder();
        let event = match reader.read_event_into(&mut event_buffer) {
            Ok(event) => event,
            Err(error) => {
                let failure = stream_error_from_quick_xml(error, raw_error, active_error);
                if recovering {
                    return match failure {
                        StreamError::Input {
                            error,
                            raw_error,
                            active_error,
                            ..
                        } => Err(StreamError::Input {
                            error,
                            prior_mce_error: recovery_error.take(),
                            raw_error,
                            active_error,
                        }),
                        StreamError::Mce {
                            error,
                            raw_error,
                            active_error,
                            ..
                        } => Err(StreamError::Mce {
                            error,
                            prior_mce_error: recovery_error.take(),
                            raw_error,
                            active_error,
                        }),
                        StreamError::Callback {
                            raw_error,
                            active_error,
                        } => Err(StreamError::Mce {
                            error: recovery_error
                                .take()
                                .unwrap_or_else(|| bad("stream recovery failure")),
                            prior_mce_error: None,
                            raw_error,
                            active_error,
                        }),
                    };
                }
                return Err(failure);
            },
        };
        let is_eof = matches!(&event, Event::Eof);

        if !matches!(&event, &Event::Eof | &Event::DocType(_) | &Event::PI(_)) {
            if let Err(error) = state.count_event() {
                if recovering {
                    return Err(StreamError::Mce {
                        error,
                        prior_mce_error: recovery_error.take(),
                        raw_error,
                        active_error,
                    });
                }
                return Err(StreamError::Mce {
                    error,
                    prior_mce_error: None,
                    raw_error,
                    active_error,
                });
            }
        }

        let result = if recovering {
            state.raw_event(event, decoder, &mut raw, &mut raw_enabled, &mut raw_error)
        } else {
            match event {
                Event::Start(element) => state.start(
                    &element,
                    decoder,
                    false,
                    &mut raw,
                    &mut raw_enabled,
                    &mut raw_error,
                    &mut active,
                    &mut active_enabled,
                    &mut active_error,
                ),
                Event::Empty(element) => state.start(
                    &element,
                    decoder,
                    true,
                    &mut raw,
                    &mut raw_enabled,
                    &mut raw_error,
                    &mut active,
                    &mut active_enabled,
                    &mut active_error,
                ),
                Event::End(element) => state.end(
                    &element,
                    &mut active,
                    &mut active_enabled,
                    &mut active_error,
                ),
                Event::Text(text) => match text.decode().map_err(xml_error) {
                    Ok(decoded) => state.text(
                        text.as_ref(),
                        decoded,
                        TextKind::Text,
                        &mut active,
                        &mut active_enabled,
                        &mut active_error,
                    ),
                    Err(error) => Err(error),
                },
                Event::CData(text) => match text.decode().map_err(xml_error) {
                    Ok(decoded) => state.text(
                        text.as_ref(),
                        decoded,
                        TextKind::CData,
                        &mut active,
                        &mut active_enabled,
                        &mut active_error,
                    ),
                    Err(error) => Err(error),
                },
                Event::Comment(text) => match text.decode().map_err(xml_error) {
                    Ok(decoded) => state.text(
                        text.as_ref(),
                        decoded,
                        TextKind::Comment,
                        &mut active,
                        &mut active_enabled,
                        &mut active_error,
                    ),
                    Err(error) => Err(error),
                },
                Event::Decl(decl) => state.decl(
                    decl.as_ref(),
                    &mut active,
                    &mut active_enabled,
                    &mut active_error,
                ),
                Event::GeneralRef(reference) => state.reference(
                    &reference,
                    &mut active,
                    &mut active_enabled,
                    &mut active_error,
                ),
                Event::DocType(_) | Event::PI(_) => {
                    Err(bad("DTD and processing instructions are rejected"))
                },
                Event::Eof => state.eof(),
            }
        };

        event_buffer.clear();
        reader.get_mut().finish_event();

        if let Err(error) = result {
            if !recovering && state.can_recover(&error) && !is_eof {
                recovery_error = Some(error);
                recovering = true;
                active_enabled = false;
                state.stack.clear();
                state.stack.shrink_to_fit();
                continue;
            }
            if recovering {
                return Err(StreamError::Mce {
                    error,
                    prior_mce_error: recovery_error.take(),
                    raw_error,
                    active_error,
                });
            }
            return Err(StreamError::Mce {
                error,
                prior_mce_error: None,
                raw_error,
                active_error,
            });
        }
        if recovering {
            if is_eof {
                return Err(StreamError::Mce {
                    error: recovery_error
                        .take()
                        .unwrap_or_else(|| bad("stream recovery failure")),
                    prior_mce_error: None,
                    raw_error,
                    active_error,
                });
            }
            continue;
        }
        if state.finished {
            return finish_callbacks(raw_error, active_error, state.report.clone());
        }
    }
}

/// Alias retained for callers that prefer an explicit active-observer name.
pub fn process_markup_compatibility_stream_with_active_observer<ActiveE, Active>(
    input: &mut dyn BufRead,
    capabilities: &Capabilities,
    limits: &StreamLimits,
    active: Active,
) -> Result<StreamReport, StreamError<std::convert::Infallible, ActiveE>>
where
    Active: for<'a> FnMut(SemanticEvent<'a>) -> Result<(), ActiveE>,
{
    process_markup_compatibility_stream(input, capabilities, limits, active)
}

fn finish_callbacks<RawE, ActiveE>(
    raw_error: Option<RawE>,
    active_error: Option<ActiveE>,
    report: StreamReport,
) -> Result<StreamReport, StreamError<RawE, ActiveE>> {
    if raw_error.is_none() && active_error.is_none() {
        Ok(report)
    } else {
        Err(StreamError::Callback {
            raw_error,
            active_error,
        })
    }
}

fn stream_error_from_quick_xml<RawE, ActiveE>(
    error: quick_xml::Error,
    raw_error: Option<RawE>,
    active_error: Option<ActiveE>,
) -> StreamError<RawE, ActiveE> {
    match error {
        quick_xml::Error::Io(error) => {
            if error
                .get_ref()
                .and_then(|source| source.downcast_ref::<InterruptedRetriesExceeded>())
                .is_some()
            {
                return StreamError::Mce {
                    error: limit("input Interrupted retries"),
                    prior_mce_error: None,
                    raw_error,
                    active_error,
                };
            }
            if error
                .get_ref()
                .and_then(|source| source.downcast_ref::<EventLimitExceeded>())
                .is_some()
            {
                return StreamError::Mce {
                    error: limit("stream event bytes"),
                    prior_mce_error: None,
                    raw_error,
                    active_error,
                };
            }
            if error
                .get_ref()
                .and_then(|source| source.downcast_ref::<InputLimitExceeded>())
                .is_some()
            {
                return StreamError::Mce {
                    error: limit("input bytes"),
                    prior_mce_error: None,
                    raw_error,
                    active_error,
                };
            }
            if error
                .get_ref()
                .and_then(|source| source.downcast_ref::<InvalidMeterConsume>())
                .is_some()
            {
                return StreamError::Mce {
                    error: bad("invalid event meter consume"),
                    prior_mce_error: None,
                    raw_error,
                    active_error,
                };
            }
            StreamError::Input {
                error: Arc::try_unwrap(error)
                    .unwrap_or_else(|error| io::Error::new(error.kind(), error)),
                prior_mce_error: None,
                raw_error,
                active_error,
            }
        },
        error => StreamError::Mce {
            error: Error::Xml(error.to_string()),
            prior_mce_error: None,
            raw_error,
            active_error,
        },
    }
}

struct PrefixReader<'a> {
    input: &'a mut dyn BufRead,
    prefix: [u8; 3],
    prefix_len: usize,
    prefix_pos: usize,
    consumed: usize,
    initialized: bool,
    exposed_remaining: usize,
    invalid_consume: bool,
}

impl<'a> PrefixReader<'a> {
    fn new(input: &'a mut dyn BufRead) -> Self {
        Self {
            input,
            prefix: [0; 3],
            prefix_len: 0,
            prefix_pos: 0,
            consumed: 0,
            initialized: false,
            exposed_remaining: 0,
            invalid_consume: false,
        }
    }

    fn initialize(&mut self) -> io::Result<()> {
        if self.initialized {
            return Ok(());
        }
        while self.prefix_len < self.prefix.len() {
            let available = self.input.fill_buf()?;
            if available.is_empty() {
                break;
            }
            let amount = (self.prefix.len() - self.prefix_len).min(available.len());
            self.prefix[self.prefix_len..self.prefix_len + amount]
                .copy_from_slice(&available[..amount]);
            self.prefix_len += amount;
            self.consumed += amount;
            self.input.consume(amount);
        }
        self.initialized = true;
        if self.prefix_len == self.prefix.len() && self.prefix == [0xef, 0xbb, 0xbf] {
            self.prefix_len = 0;
        }
        self.prefix_pos = 0;
        self.exposed_remaining = 0;
        Ok(())
    }

    fn consumed_bytes(&self) -> usize {
        self.consumed
    }

    fn replay_bytes(&self) -> usize {
        self.prefix_len
    }
}

impl Read for PrefixReader<'_> {
    fn read(&mut self, output: &mut [u8]) -> io::Result<usize> {
        if output.is_empty() {
            return Ok(0);
        }
        let available = self.fill_buf()?;
        let count = available.len().min(output.len());
        output[..count].copy_from_slice(&available[..count]);
        self.consume(count);
        Ok(count)
    }
}

impl BufRead for PrefixReader<'_> {
    fn fill_buf(&mut self) -> io::Result<&[u8]> {
        if self.invalid_consume {
            return Err(io::Error::new(
                io::ErrorKind::InvalidData,
                InvalidMeterConsume,
            ));
        }
        if !self.initialized {
            self.initialize()?;
        }
        if self.prefix_pos < self.prefix_len {
            let available = &self.prefix[self.prefix_pos..self.prefix_len];
            self.exposed_remaining = available.len();
            return Ok(available);
        }
        let available = match self.input.fill_buf() {
            Ok(available) => available,
            Err(error) => {
                self.exposed_remaining = 0;
                return Err(error);
            },
        };
        self.exposed_remaining = available.len();
        Ok(available)
    }

    fn consume(&mut self, amount: usize) {
        if amount > self.exposed_remaining {
            self.invalid_consume = true;
            return;
        }
        if self.prefix_pos < self.prefix_len {
            self.prefix_pos += amount;
        } else {
            self.input.consume(amount);
        }
        self.exposed_remaining -= amount;
    }
}

struct EventMeter<'a> {
    input: &'a mut dyn BufRead,
    event_limit: usize,
    total_limit: usize,
    event_bytes: usize,
    total_bytes: usize,
    precharged_replay: usize,
    exposed_remaining: usize,
    invalid_consume: bool,
}

impl<'a> EventMeter<'a> {
    fn new(
        input: &'a mut dyn BufRead,
        event_limit: usize,
        total_limit: usize,
        initial_total_bytes: usize,
        precharged_replay: usize,
    ) -> Self {
        Self {
            input,
            event_limit,
            total_limit,
            event_bytes: 0,
            total_bytes: initial_total_bytes,
            precharged_replay,
            exposed_remaining: 0,
            invalid_consume: false,
        }
    }

    fn begin_event(&mut self) {
        self.event_bytes = 0;
        self.invalid_consume = false;
    }

    fn finish_event(&mut self) {
        self.event_bytes = 0;
        self.invalid_consume = false;
    }

    fn remaining(&self) -> usize {
        self.event_limit.saturating_sub(self.event_bytes)
    }
}

impl Read for EventMeter<'_> {
    fn read(&mut self, output: &mut [u8]) -> io::Result<usize> {
        if output.is_empty() {
            return Ok(0);
        }
        let available = self.fill_buf()?;
        let count = available.len().min(output.len());
        output[..count].copy_from_slice(&available[..count]);
        self.consume(count);
        Ok(count)
    }
}

impl BufRead for EventMeter<'_> {
    fn fill_buf(&mut self) -> io::Result<&[u8]> {
        if self.invalid_consume {
            return Err(io::Error::new(
                io::ErrorKind::InvalidData,
                InvalidMeterConsume,
            ));
        }
        let total_remaining = self
            .total_limit
            .saturating_sub(self.total_bytes)
            .saturating_add(self.precharged_replay);
        let available = self.input.fill_buf()?;
        if available.is_empty() {
            self.exposed_remaining = 0;
            return Ok(available);
        }
        let event_remaining = self.event_limit.saturating_sub(self.event_bytes);
        if total_remaining == 0 {
            self.exposed_remaining = 0;
            return Err(io::Error::new(
                io::ErrorKind::InvalidData,
                InputLimitExceeded {
                    limit: self.total_limit,
                },
            ));
        }
        if event_remaining == 0 {
            self.exposed_remaining = 0;
            return Err(io::Error::new(
                io::ErrorKind::InvalidData,
                EventLimitExceeded {
                    limit: self.event_limit,
                },
            ));
        }
        let count = available.len().min(event_remaining).min(total_remaining);
        self.exposed_remaining = count;
        Ok(&available[..count])
    }

    fn consume(&mut self, amount: usize) {
        let event_remaining = self.remaining();
        let total_remaining = self
            .total_limit
            .saturating_sub(self.total_bytes)
            .saturating_add(self.precharged_replay);
        let valid = amount <= self.exposed_remaining
            && amount <= event_remaining
            && amount <= total_remaining;
        if !valid {
            self.invalid_consume = true;
            return;
        }
        self.input.consume(amount);
        self.event_bytes = self
            .event_bytes
            .checked_add(amount)
            .unwrap_or(self.event_limit);
        let replayed = amount.min(self.precharged_replay);
        self.precharged_replay -= replayed;
        self.total_bytes = self
            .total_bytes
            .checked_add(amount - replayed)
            .unwrap_or(self.total_limit);
        self.exposed_remaining -= amount;
    }
}

#[derive(Clone)]
struct Namespaces {
    head: Option<Arc<NamespaceLayer>>,
    bindings: usize,
    bytes: usize,
}

struct NamespaceLayer {
    parent: Option<Arc<NamespaceLayer>>,
    local: Vec<(String, String)>,
}

impl Namespaces {
    fn root() -> Self {
        Self {
            head: None,
            bindings: 1,
            bytes: 0,
        }
    }

    fn get(&self, prefix: &str) -> Option<&str> {
        if prefix == "xml" {
            return Some(XML_NS);
        }
        let mut layer = self.head.as_deref();
        while let Some(current) = layer {
            if let Some((_, namespace)) = current
                .local
                .iter()
                .rev()
                .find(|(candidate, _)| candidate == prefix)
            {
                return Some(namespace);
            }
            layer = current.parent.as_deref();
        }
        None
    }

    fn with_local(
        &self,
        local: Vec<(String, String)>,
        limits: &StreamLimits,
    ) -> Result<Self, Error> {
        if local.is_empty() {
            return Ok(self.clone());
        }
        let mut bindings = self.bindings;
        let mut bytes = self.bytes;
        for (index, (prefix, namespace)) in local.iter().enumerate() {
            if local[..index]
                .iter()
                .any(|(candidate, _)| candidate == prefix)
            {
                return Err(bad("duplicate namespace declaration"));
            }
            if self.get(prefix).is_none() {
                bindings = bindings
                    .checked_add(1)
                    .ok_or_else(|| limit("namespace bindings"))?;
            }
            bytes = bytes
                .checked_add(prefix.len())
                .and_then(|bytes| bytes.checked_add(namespace.len()))
                .and_then(|bytes| bytes.checked_add(2))
                .ok_or_else(|| limit("stream context bytes"))?;
        }
        if bindings > limits.processing.max_namespace_bindings {
            return Err(limit("namespace bindings"));
        }
        if bytes > limits.max_context_bytes {
            return Err(limit("stream context bytes"));
        }
        Ok(Self {
            head: Some(Arc::new(NamespaceLayer {
                parent: self.head.clone(),
                local,
            })),
            bindings,
            bytes,
        })
    }
}

#[derive(Clone, PartialEq, Eq, Hash)]
enum NamePattern {
    Exact(Name),
    Namespace(String),
}

struct DirectiveLayer {
    parent: Option<Arc<DirectiveLayer>>,
    ignorable: HashSet<String>,
    process: HashSet<NamePattern>,
    preserve_elements: HashSet<NamePattern>,
    preserve_attributes: HashSet<NamePattern>,
}

#[derive(Clone)]
struct Context {
    ns: Namespaces,
    directives: Option<Arc<DirectiveLayer>>,
    opaque: bool,
    bytes: usize,
}

impl Context {
    fn root() -> Self {
        Self {
            ns: Namespaces::root(),
            directives: None,
            opaque: false,
            bytes: 0,
        }
    }

    fn is_ignorable(&self, namespace: &str) -> bool {
        let mut layer = self.directives.as_deref();
        while let Some(current) = layer {
            if current.ignorable.contains(namespace) {
                return true;
            }
            layer = current.parent.as_deref();
        }
        false
    }

    fn matches(
        &self,
        name: &Name,
        select: impl Fn(&DirectiveLayer) -> &HashSet<NamePattern>,
    ) -> bool {
        let mut layer = self.directives.as_deref();
        while let Some(current) = layer {
            if matches_pattern(select(current), name) {
                return true;
            }
            if current.ignorable.contains(&name.namespace) {
                return false;
            }
            layer = current.parent.as_deref();
        }
        false
    }

    fn processes(&self, name: &Name) -> bool {
        self.matches(name, |layer| &layer.process)
    }

    fn preserves_element(&self, name: &Name) -> bool {
        self.matches(name, |layer| &layer.preserve_elements)
    }

    fn preserves_attribute(&self, name: &Name) -> bool {
        self.matches(name, |layer| &layer.preserve_attributes)
    }
}

enum Mode {
    Emit,
    Unwrap,
    Skip,
    Alt {
        choices: usize,
        selected: bool,
        fallback: bool,
    },
    Branch,
}

#[derive(Clone, Copy)]
enum TextKind {
    Text,
    CData,
    Comment,
}

struct Frame {
    ctx: Context,
    mode: Mode,
    active: bool,
    qualified_name: String,
    expanded_name: Name,
}

struct RawFrame {
    ns: Namespaces,
    context_bytes: usize,
    qualified_name: String,
    opaque: bool,
    alternate_content: bool,
}

struct ElementData<'a> {
    qualified_name: Cow<'a, [u8]>,
    expanded_name: Name,
    namespace: Namespaces,
    attrs: Vec<RawAttribute<'a>>,
}

struct Processor<'a> {
    capabilities: &'a Capabilities,
    limits: &'a StreamLimits,
    stack: Vec<Frame>,
    raw_stack: Vec<RawFrame>,
    root_seen: bool,
    raw_root_seen: bool,
    prolog_started: bool,
    declaration_seen: bool,
    raw_prolog_started: bool,
    raw_declaration_seen: bool,
    raw_state_valid: bool,
    finished: bool,
    report: StreamReport,
}

impl<'a> Processor<'a> {
    fn new(capabilities: &'a Capabilities, limits: &'a StreamLimits) -> Self {
        Self {
            capabilities,
            limits,
            stack: Vec::new(),
            raw_stack: Vec::new(),
            root_seen: false,
            raw_root_seen: false,
            prolog_started: false,
            declaration_seen: false,
            raw_prolog_started: false,
            raw_declaration_seen: false,
            raw_state_valid: false,
            finished: false,
            report: StreamReport::default(),
        }
    }

    fn count_event(&mut self) -> Result<(), Error> {
        self.report.events = self
            .report
            .events
            .checked_add(1)
            .ok_or_else(|| limit("stream event count"))?;
        if self.report.events > self.limits.max_events {
            return Err(limit("stream event count"));
        }
        Ok(())
    }

    fn can_recover(&self, error: &Error) -> bool {
        self.raw_state_valid && matches!(error, Error::NonConformant(_) | Error::MustUnderstand(_))
    }

    #[allow(clippy::too_many_arguments)]
    fn start<RawE, ActiveE, Raw, Active>(
        &mut self,
        element: &BytesStart<'_>,
        decoder: Decoder,
        empty: bool,
        raw_callback: &mut Raw,
        raw_enabled: &mut bool,
        raw_error: &mut Option<RawE>,
        active_callback: &mut Active,
        active_enabled: &mut bool,
        active_error: &mut Option<ActiveE>,
    ) -> Result<(), Error>
    where
        Raw: for<'b> FnMut(RawElement<'b>) -> Result<(), RawE>,
        Active: for<'b> FnMut(SemanticEvent<'b>) -> Result<(), ActiveE>,
    {
        if self.raw_stack.len() >= self.limits.processing.max_depth {
            return Err(limit("depth"));
        }
        self.prolog_started = true;
        let raw_parent = self
            .raw_stack
            .last()
            .map_or_else(Namespaces::root, |frame| frame.ns.clone());
        let mut data = self.parse_element(element, decoder, &raw_parent)?;
        if *raw_enabled {
            let raw = RawElement {
                kind: if empty {
                    RawElementKind::Empty
                } else {
                    RawElementKind::Start
                },
                qualified_name: Cow::Borrowed(data.qualified_name.as_ref()),
                expanded_name: clone_bounded_name(
                    &data.expanded_name,
                    self.limits,
                    "MCE stream raw expanded element name",
                )?,
                attributes: clone_raw_attributes(&data.attrs, self.limits)?,
            };
            if let Err(error) = raw_callback(raw) {
                *raw_error = Some(error);
                *raw_enabled = false;
            }
        }

        validate_duplicate_attributes(&data.attrs, self.limits)?;
        self.raw_start_commit(&data, empty)?;
        self.raw_state_valid = true;

        let parent_active = self.stack.last().is_none_or(|frame| frame.active);
        let mut context = self
            .stack
            .last()
            .map_or_else(Context::root, |frame| frame.ctx.clone());

        let local_namespaces = local_namespaces(&data.attrs, self.limits)?;
        let inherited_namespace_bytes = context.ns.bytes;
        let inherited_context_bytes = context.bytes;
        context.ns = context.ns.with_local(local_namespaces, self.limits)?;
        let namespace_bytes = context
            .ns
            .bytes
            .checked_sub(inherited_namespace_bytes)
            .ok_or_else(|| limit("stream context bytes"))?;
        context.bytes = inherited_context_bytes
            .checked_add(namespace_bytes)
            .ok_or_else(|| limit("stream context bytes"))?;
        if context.bytes > self.limits.max_context_bytes {
            return Err(limit("stream context bytes"));
        }

        if context.opaque {
            let mode = Mode::Emit;
            if parent_active && *active_enabled {
                let event = SemanticEvent::from_element(
                    data.qualified_name.as_ref(),
                    &data.expanded_name,
                    empty,
                    semantic_attrs(
                        &mut data.attrs,
                        &context,
                        self.capabilities,
                        false,
                        self.limits,
                        &mut self.report,
                    )?,
                    self.limits,
                )?;
                invoke_active(active_callback, active_enabled, active_error, event);
            }
            return self.close_start(context, mode, parent_active, data, empty);
        }

        let directives = collect_directives(&data.attrs, self.limits)?;
        let directive_tokens = directives.tokens;
        apply_directives(&mut context, directives, self.capabilities, self.limits)?;
        let is_extension = self.capabilities.extensions.contains(&data.expanded_name);
        context.opaque = is_extension;

        validate_alternate_attributes_if_needed(&data, &context, self.capabilities)?;

        if let Some(parent) = self.stack.last_mut()
            && let Mode::Alt {
                choices,
                selected,
                fallback,
            } = &mut parent.mode
        {
            if data.expanded_name.namespace != NAMESPACE {
                if context.is_ignorable(&data.expanded_name.namespace)
                    && !self.capabilities.understands(&data.expanded_name.namespace)
                {
                    self.report.ignored_elements = self
                        .report
                        .ignored_elements
                        .checked_add(1)
                        .ok_or_else(|| limit("stream report counters"))?;
                    return self.close_start(context, Mode::Skip, false, data, empty);
                }
                return Err(bad("non-ignorable AlternateContent child"));
            }
            let (active, mode) = match data.expanded_name.local_name.as_str() {
                "Choice" => {
                    if *fallback {
                        return Err(bad("Choice after Fallback"));
                    }
                    *choices = choices.checked_add(1).ok_or_else(|| limit("choices"))?;
                    if *choices > self.limits.processing.max_choices_per_alternate {
                        return Err(limit("choices"));
                    }
                    let requires = attr_value(&data.attrs, "Requires")?
                        .ok_or_else(|| bad("Choice lacks Requires"))?;
                    let mut supported = true;
                    let mut count = 0usize;
                    for prefix in requires.split_whitespace() {
                        if !xml_name::is_ncname(prefix) {
                            return Err(bad("invalid Requires prefix"));
                        }
                        check_name_bytes(prefix.as_bytes(), self.limits)?;
                        count = count
                            .checked_add(1)
                            .ok_or_else(|| limit("directive tokens"))?;
                        if directive_tokens.checked_add(count).is_none_or(|tokens| {
                            tokens > self.limits.processing.max_directive_tokens
                        }) {
                            return Err(limit("directive tokens"));
                        }
                        let namespace = context
                            .ns
                            .get(prefix)
                            .ok_or_else(|| bad("unbound Requires prefix"))?;
                        supported &= self.capabilities.understands(namespace);
                    }
                    if count == 0 {
                        return Err(bad("empty Requires"));
                    }
                    let active = parent.active && !*selected && supported;
                    if active {
                        *selected = true;
                        self.report.selected_choices = self
                            .report
                            .selected_choices
                            .checked_add(1)
                            .ok_or_else(|| limit("stream report counters"))?;
                    }
                    (active, Mode::Branch)
                },
                "Fallback" => {
                    if *fallback {
                        return Err(bad("duplicate Fallback"));
                    }
                    *fallback = true;
                    let active = parent.active && !*selected;
                    if active {
                        *selected = true;
                        self.report.selected_fallbacks = self
                            .report
                            .selected_fallbacks
                            .checked_add(1)
                            .ok_or_else(|| limit("stream report counters"))?;
                    }
                    (active, Mode::Branch)
                },
                _ => return Err(bad("invalid AlternateContent child")),
            };
            return self.close_start(context, mode, active, data, empty);
        }

        let mut active = parent_active;
        let mode = if is_extension {
            Mode::Emit
        } else if data.expanded_name.namespace == NAMESPACE {
            match data.expanded_name.local_name.as_str() {
                "AlternateContent" => {
                    self.report.alternate_content_count = self
                        .report
                        .alternate_content_count
                        .checked_add(1)
                        .ok_or_else(|| limit("stream report counters"))?;
                    Mode::Alt {
                        choices: 0,
                        selected: false,
                        fallback: false,
                    }
                },
                _ => return Err(bad("Choice/Fallback outside AlternateContent")),
            }
        } else if context.is_ignorable(&data.expanded_name.namespace)
            && !self.capabilities.understands(&data.expanded_name.namespace)
        {
            if context.preserves_element(&data.expanded_name) {
                self.report.preserved_elements = self
                    .report
                    .preserved_elements
                    .checked_add(1)
                    .ok_or_else(|| limit("stream report counters"))?;
                Mode::Emit
            } else if context.processes(&data.expanded_name) {
                for attr in &data.attrs {
                    if attr.expanded_name.namespace == XML_NS
                        && matches!(
                            attr.expanded_name.local_name.as_str(),
                            "base" | "lang" | "space"
                        )
                    {
                        return Err(bad("xml context attribute on unwrapped element"));
                    }
                }
                self.report.unwrapped_elements = self
                    .report
                    .unwrapped_elements
                    .checked_add(1)
                    .ok_or_else(|| limit("stream report counters"))?;
                Mode::Unwrap
            } else {
                active = false;
                self.report.ignored_elements = self
                    .report
                    .ignored_elements
                    .checked_add(1)
                    .ok_or_else(|| limit("stream report counters"))?;
                Mode::Skip
            }
        } else {
            Mode::Emit
        };

        if parent_active && active && matches!(&mode, Mode::Emit) && *active_enabled {
            let event = SemanticEvent::from_element(
                data.qualified_name.as_ref(),
                &data.expanded_name,
                empty,
                semantic_attrs(
                    &mut data.attrs,
                    &context,
                    self.capabilities,
                    !is_extension,
                    self.limits,
                    &mut self.report,
                )?,
                self.limits,
            )?;
            invoke_active(active_callback, active_enabled, active_error, event);
        }
        self.close_start(context, mode, active, data, empty)
    }

    fn close_start(
        &mut self,
        mut context: Context,
        mode: Mode,
        active: bool,
        data: ElementData<'_>,
        empty: bool,
    ) -> Result<(), Error> {
        if self.stack.is_empty() {
            if self.root_seen {
                return Err(bad("multiple roots"));
            }
            self.root_seen = true;
        }
        if empty {
            if matches!(&mode, Mode::Alt { .. }) {
                return Err(bad("empty AlternateContent"));
            }
            return Ok(());
        }
        let retained_name_bytes = data
            .qualified_name
            .as_ref()
            .len()
            .checked_add(data.expanded_name.namespace.len())
            .and_then(|bytes| bytes.checked_add(data.expanded_name.local_name.len()))
            .ok_or_else(|| limit("stream context bytes"))?;
        context.bytes = context
            .bytes
            .checked_add(retained_name_bytes)
            .ok_or_else(|| limit("stream context bytes"))?;
        if context.bytes > self.limits.max_context_bytes {
            return Err(limit("stream context bytes"));
        }
        reserve_vec(&mut self.stack, 1, "MCE stream element stack")?;
        let qualified_name = {
            let name = str::from_utf8(data.qualified_name.as_ref()).map_err(xml_error)?;
            clone_bounded_string(name, self.limits, "MCE stream frame name")?
        };
        self.stack.push(Frame {
            ctx: context,
            mode,
            active,
            qualified_name,
            expanded_name: data.expanded_name,
        });
        Ok(())
    }

    fn end<ActiveE, Active>(
        &mut self,
        element: &BytesEnd<'_>,
        active_callback: &mut Active,
        active_enabled: &mut bool,
        active_error: &mut Option<ActiveE>,
    ) -> Result<(), Error>
    where
        Active: for<'b> FnMut(SemanticEvent<'b>) -> Result<(), ActiveE>,
    {
        self.raw_end(element)?;
        self.raw_state_valid = true;
        let frame = self.stack.last().ok_or_else(|| bad("unexpected end"))?;
        let end_name = element.name();
        let end_name = end_name.as_ref();
        check_name_bytes(end_name, self.limits)?;
        let found = str::from_utf8(end_name).map_err(xml_error)?;
        if found != frame.qualified_name {
            return Err(bad("mismatched end tag"));
        }
        let frame = self.stack.pop().ok_or_else(|| bad("unexpected end"))?;
        if frame.active && matches!(&frame.mode, Mode::Emit) && *active_enabled {
            let event = SemanticEvent::End(SemanticEnd {
                qualified_name: Cow::Borrowed(frame.qualified_name.as_bytes()),
                expanded_name: frame.expanded_name,
            });
            invoke_active(active_callback, active_enabled, active_error, event);
        }
        if let Mode::Alt {
            choices,
            selected: _,
            fallback: _,
        } = frame.mode
        {
            if choices == 0 {
                return Err(bad("AlternateContent requires Choice"));
            }
        }
        Ok(())
    }

    fn text<ActiveE, Active>(
        &mut self,
        raw: &[u8],
        decoded: Cow<'_, str>,
        kind: TextKind,
        active_callback: &mut Active,
        active_enabled: &mut bool,
        active_error: &mut Option<ActiveE>,
    ) -> Result<(), Error>
    where
        Active: for<'b> FnMut(SemanticEvent<'b>) -> Result<(), ActiveE>,
    {
        self.raw_text(raw, kind)?;
        self.raw_state_valid = true;
        self.prolog_started = true;
        if self.stack.is_empty() {
            if matches!(kind, TextKind::CData) {
                return Err(bad("CDATA outside document root"));
            }
            if matches!(kind, TextKind::Comment) {
                return Ok(());
            }
            if !raw.iter().all(|byte| byte.is_ascii_whitespace()) {
                return Err(bad("text outside document root"));
            }
            return Ok(());
        }
        if self.stack.last().is_some_and(|frame| {
            matches!(&frame.mode, Mode::Alt { .. })
                && (matches!(kind, TextKind::CData)
                    || (matches!(kind, TextKind::Text)
                        && !raw.iter().all(|byte| byte.is_ascii_whitespace())))
        }) {
            return Err(bad("invalid content directly inside AlternateContent"));
        }
        if self.stack.last().is_some_and(|frame| frame.active) && *active_enabled {
            let event = match kind {
                TextKind::CData => SemanticEvent::CData(SemanticText {
                    raw: Cow::Borrowed(raw),
                    decoded,
                }),
                TextKind::Text => {
                    if !self.stack.last().is_some_and(|frame| {
                        matches!(
                            &frame.mode,
                            Mode::Emit | Mode::Unwrap | Mode::Branch | Mode::Alt { .. }
                        )
                    }) {
                        return Ok(());
                    }
                    SemanticEvent::Text(SemanticText {
                        raw: Cow::Borrowed(raw),
                        decoded,
                    })
                },
                TextKind::Comment => SemanticEvent::Comment(SemanticText {
                    raw: Cow::Borrowed(raw),
                    decoded,
                }),
            };
            invoke_active(active_callback, active_enabled, active_error, event);
        }
        Ok(())
    }

    fn decl<ActiveE, Active>(
        &mut self,
        declaration: &[u8],
        active_callback: &mut Active,
        active_enabled: &mut bool,
        active_error: &mut Option<ActiveE>,
    ) -> Result<(), Error>
    where
        Active: for<'b> FnMut(SemanticEvent<'b>) -> Result<(), ActiveE>,
    {
        self.raw_decl(declaration)?;
        self.raw_state_valid = true;
        if self.declaration_seen || self.prolog_started || self.root_seen || !self.stack.is_empty()
        {
            return Err(bad("late XML declaration"));
        }
        self.declaration_seen = true;
        self.prolog_started = true;
        if *active_enabled {
            invoke_active(
                active_callback,
                active_enabled,
                active_error,
                SemanticEvent::Decl(SemanticDecl {
                    raw: Cow::Borrowed(declaration),
                }),
            );
        }
        Ok(())
    }

    fn reference<ActiveE, Active>(
        &mut self,
        reference: &quick_xml::events::BytesRef<'_>,
        active_callback: &mut Active,
        active_enabled: &mut bool,
        active_error: &mut Option<ActiveE>,
    ) -> Result<(), Error>
    where
        Active: for<'b> FnMut(SemanticEvent<'b>) -> Result<(), ActiveE>,
    {
        self.raw_reference(reference)?;
        self.raw_state_valid = true;
        if self.stack.is_empty() {
            return Err(bad("reference outside document root"));
        }
        if self
            .stack
            .last()
            .is_some_and(|frame| matches!(&frame.mode, Mode::Alt { .. }))
        {
            return Err(bad("general reference directly inside AlternateContent"));
        }
        self.prolog_started = true;
        check_name_bytes(reference.as_ref(), self.limits)?;
        let char_ref = reference.resolve_char_ref().map_err(xml_error)?;
        let name = reference.decode().map_err(xml_error)?;
        if !matches!(name.as_ref(), "amp" | "lt" | "gt" | "apos" | "quot") && char_ref.is_none() {
            return Err(bad("custom entity"));
        }
        if self.stack.last().is_some_and(|frame| frame.active) && *active_enabled {
            invoke_active(
                active_callback,
                active_enabled,
                active_error,
                SemanticEvent::GeneralRef(SemanticGeneralRef {
                    name: Cow::Borrowed(reference.as_ref()),
                }),
            );
        }
        Ok(())
    }

    fn eof(&mut self) -> Result<(), Error> {
        self.raw_eof()?;
        self.raw_state_valid = true;
        if !self.stack.is_empty() {
            return Err(bad("unterminated XML"));
        }
        if !self.root_seen {
            return Err(bad("missing document root"));
        }
        self.finished = true;
        Ok(())
    }

    #[allow(clippy::too_many_arguments)]
    fn raw_start<RawE, Raw>(
        &mut self,
        element: &BytesStart<'_>,
        decoder: Decoder,
        empty: bool,
        raw_callback: &mut Raw,
        raw_enabled: &mut bool,
        raw_error: &mut Option<RawE>,
    ) -> Result<(), Error>
    where
        Raw: for<'b> FnMut(RawElement<'b>) -> Result<(), RawE>,
    {
        if self.raw_stack.len() >= self.limits.processing.max_depth {
            return Err(limit("depth"));
        }
        let raw_parent = self
            .raw_stack
            .last()
            .map_or_else(Namespaces::root, |frame| frame.ns.clone());
        let data = self.parse_element(element, decoder, &raw_parent)?;
        if *raw_enabled {
            let raw = RawElement {
                kind: if empty {
                    RawElementKind::Empty
                } else {
                    RawElementKind::Start
                },
                qualified_name: Cow::Borrowed(data.qualified_name.as_ref()),
                expanded_name: clone_bounded_name(
                    &data.expanded_name,
                    self.limits,
                    "MCE stream raw expanded element name",
                )?,
                attributes: clone_raw_attributes(&data.attrs, self.limits)?,
            };
            if let Err(error) = raw_callback(raw) {
                *raw_error = Some(error);
                *raw_enabled = false;
            }
        }
        validate_duplicate_attributes(&data.attrs, self.limits)?;
        self.raw_start_commit(&data, empty)
    }

    fn raw_start_commit(&mut self, data: &ElementData<'_>, empty: bool) -> Result<(), Error> {
        self.raw_prolog_started = true;
        if self.raw_stack.is_empty() {
            if self.raw_root_seen {
                return Err(bad("multiple roots"));
            }
            self.raw_root_seen = true;
        }
        let parent_context_bytes = self.raw_stack.last().map_or(0, |frame| frame.context_bytes);
        let parent_namespace_bytes = self.raw_stack.last().map_or(0, |frame| frame.ns.bytes);
        let namespace_delta = data
            .namespace
            .bytes
            .checked_sub(parent_namespace_bytes)
            .ok_or_else(|| limit("stream context bytes"))?;
        let inherited_opaque = self.raw_stack.last().is_some_and(|frame| frame.opaque);
        let directive_bytes = if inherited_opaque {
            0
        } else {
            raw_directive_bytes(&data.attrs, &data.expanded_name, self.limits)?
        };
        let frame_name_bytes = data.qualified_name.as_ref().len();
        let context_bytes = parent_context_bytes
            .checked_add(namespace_delta)
            .and_then(|bytes| bytes.checked_add(directive_bytes))
            .and_then(|bytes| bytes.checked_add(frame_name_bytes))
            .ok_or_else(|| limit("stream context bytes"))?;
        if context_bytes > self.limits.max_context_bytes {
            return Err(limit("stream context bytes"));
        }
        if empty {
            return Ok(());
        }
        reserve_vec(&mut self.raw_stack, 1, "MCE stream raw element stack")?;
        let qualified_name = str::from_utf8(data.qualified_name.as_ref()).map_err(xml_error)?;
        let qualified_name =
            clone_bounded_string(qualified_name, self.limits, "MCE stream raw frame name")?;
        let current_opaque =
            inherited_opaque || self.capabilities.extensions.contains(&data.expanded_name);
        let alternate_content = !current_opaque
            && data.expanded_name.namespace == NAMESPACE
            && data.expanded_name.local_name == "AlternateContent";
        self.raw_stack.push(RawFrame {
            ns: data.namespace.clone(),
            context_bytes,
            qualified_name,
            opaque: current_opaque,
            alternate_content,
        });
        Ok(())
    }

    fn raw_end(&mut self, element: &BytesEnd<'_>) -> Result<(), Error> {
        let frame = self.raw_stack.last().ok_or_else(|| bad("unexpected end"))?;
        let end_name = element.name();
        let end_name = end_name.as_ref();
        check_name_bytes(end_name, self.limits)?;
        let found = str::from_utf8(end_name).map_err(xml_error)?;
        if found != frame.qualified_name {
            return Err(bad("mismatched end tag"));
        }
        self.raw_stack.pop();
        Ok(())
    }

    fn raw_text(&mut self, raw: &[u8], kind: TextKind) -> Result<(), Error> {
        self.raw_prolog_started = true;
        if self.raw_stack.is_empty() {
            if matches!(kind, TextKind::CData) {
                return Err(bad("CDATA outside document root"));
            }
            if matches!(kind, TextKind::Comment) {
                return Ok(());
            }
            if !raw.iter().all(|byte| byte.is_ascii_whitespace()) {
                return Err(bad("text outside document root"));
            }
            return Ok(());
        }
        if self.raw_stack.last().is_some_and(|frame| {
            frame.alternate_content
                && (matches!(kind, TextKind::CData)
                    || (matches!(kind, TextKind::Text)
                        && !raw.iter().all(|byte| byte.is_ascii_whitespace())))
        }) {
            return Err(bad("invalid content directly inside AlternateContent"));
        }
        Ok(())
    }

    fn raw_decl(&mut self, declaration: &[u8]) -> Result<(), Error> {
        let _ = declaration;
        if self.raw_declaration_seen
            || self.raw_prolog_started
            || self.raw_root_seen
            || !self.raw_stack.is_empty()
        {
            return Err(bad("late XML declaration"));
        }
        self.raw_declaration_seen = true;
        self.raw_prolog_started = true;
        Ok(())
    }

    fn raw_reference(&mut self, reference: &quick_xml::events::BytesRef<'_>) -> Result<(), Error> {
        if self.raw_stack.is_empty() {
            return Err(bad("reference outside document root"));
        }
        if self
            .raw_stack
            .last()
            .is_some_and(|frame| frame.alternate_content)
        {
            return Err(bad("general reference directly inside AlternateContent"));
        }
        self.raw_prolog_started = true;
        check_name_bytes(reference.as_ref(), self.limits)?;
        let char_ref = reference.resolve_char_ref().map_err(xml_error)?;
        let name = reference.decode().map_err(xml_error)?;
        if !matches!(name.as_ref(), "amp" | "lt" | "gt" | "apos" | "quot") && char_ref.is_none() {
            return Err(bad("custom entity"));
        }
        Ok(())
    }

    fn raw_eof(&mut self) -> Result<(), Error> {
        if !self.raw_stack.is_empty() {
            return Err(bad("unterminated XML"));
        }
        if !self.raw_root_seen {
            return Err(bad("missing document root"));
        }
        Ok(())
    }

    #[allow(clippy::too_many_arguments)]
    fn raw_event<RawE, Raw>(
        &mut self,
        event: Event<'_>,
        decoder: Decoder,
        raw_callback: &mut Raw,
        raw_enabled: &mut bool,
        raw_error: &mut Option<RawE>,
    ) -> Result<(), Error>
    where
        Raw: for<'b> FnMut(RawElement<'b>) -> Result<(), RawE>,
    {
        match event {
            Event::Start(element) => self.raw_start(
                &element,
                decoder,
                false,
                raw_callback,
                raw_enabled,
                raw_error,
            ),
            Event::Empty(element) => self.raw_start(
                &element,
                decoder,
                true,
                raw_callback,
                raw_enabled,
                raw_error,
            ),
            Event::End(element) => self.raw_end(&element),
            Event::Text(text) => self.raw_text(text.as_ref(), TextKind::Text),
            Event::CData(text) => self.raw_text(text.as_ref(), TextKind::CData),
            Event::Comment(text) => self.raw_text(text.as_ref(), TextKind::Comment),
            Event::Decl(declaration) => self.raw_decl(declaration.as_ref()),
            Event::GeneralRef(reference) => self.raw_reference(&reference),
            Event::DocType(_) | Event::PI(_) => {
                Err(bad("DTD and processing instructions are rejected"))
            },
            Event::Eof => self.raw_eof(),
        }
    }

    fn parse_element<'b>(
        &mut self,
        element: &BytesStart<'b>,
        decoder: Decoder,
        parent_ns: &Namespaces,
    ) -> Result<ElementData<'b>, Error> {
        let element_name = element.name();
        let qualified_name = element_name.as_ref();
        check_name_bytes(qualified_name, self.limits)?;
        let qualified = str::from_utf8(qualified_name).map_err(xml_error)?;
        let mut preliminary = Vec::new();
        let mut attr_bytes = 0usize;
        for attribute in element.attributes().with_checks(false) {
            let attribute = attribute.map_err(xml_error)?;
            check_name_bytes(attribute.key.as_ref(), self.limits)?;
            attr_bytes = attr_bytes
                .checked_add(attribute.key.as_ref().len())
                .and_then(|bytes| bytes.checked_add(attribute.value.as_ref().len()))
                .ok_or_else(|| limit("stream attribute bytes"))?;
            if attr_bytes > self.limits.max_attribute_bytes_per_event {
                return Err(limit("stream attribute bytes"));
            }
            if preliminary.len() >= self.limits.max_attributes_per_event {
                return Err(limit("stream attributes per event"));
            }
            let decoded = attribute
                .decoded_and_normalized_value(XmlVersion::Explicit1_0, decoder)
                .map_err(xml_error)?;
            reserve_vec(&mut preliminary, 1, "MCE stream attributes")?;
            preliminary.push(RawAttribute {
                qualified_name: clone_bounded_bytes(
                    attribute.key.as_ref(),
                    self.limits,
                    "MCE stream attribute name",
                )?,
                expanded_name: Name {
                    namespace: String::new(),
                    local_name: String::new(),
                },
                raw_value: clone_bounded_bytes(
                    attribute.value.as_ref(),
                    self.limits,
                    "MCE stream attribute value",
                )?,
                decoded_value: clone_bounded_text(
                    decoded.as_ref(),
                    self.limits,
                    "MCE stream decoded attribute value",
                )?,
            });
        }
        let local = local_namespaces(&preliminary, self.limits)?;
        let namespace = parent_ns.with_local(local, self.limits)?;
        let expanded_name = expand(qualified, &namespace, true, self.limits)?;
        check_expanded_name(&expanded_name, self.limits)?;
        for attribute in &mut preliminary {
            attribute.expanded_name =
                expand_attribute(attribute.qualified_name.as_ref(), &namespace, self.limits)?;
            check_expanded_name(&attribute.expanded_name, self.limits)?;
        }
        Ok(ElementData {
            qualified_name: clone_bounded_bytes(
                qualified_name,
                self.limits,
                "MCE stream element name",
            )?,
            expanded_name,
            namespace,
            attrs: preliminary,
        })
    }
}

fn raw_directive_bytes(
    attrs: &[RawAttribute<'_>],
    element_name: &Name,
    limits: &StreamLimits,
) -> Result<usize, Error> {
    let mut bytes = 0usize;
    let mut tokens = 0usize;
    for attribute in attrs {
        if is_namespace_attribute(attribute.qualified_name.as_ref()) {
            continue;
        }
        let is_requires = element_name.namespace == NAMESPACE
            && element_name.local_name == "Choice"
            && attribute.expanded_name.namespace.is_empty()
            && semantic_attr_name(attribute) == b"Requires";
        if attribute.expanded_name.namespace != NAMESPACE && !is_requires {
            continue;
        }
        let value = attribute.decoded_value.as_ref();
        bytes = bytes
            .checked_add(value.len())
            .and_then(|value_bytes| value_bytes.checked_add(1))
            .ok_or_else(|| limit("stream context bytes"))?;
        for token in value.split_whitespace() {
            tokens = tokens
                .checked_add(1)
                .ok_or_else(|| limit("directive tokens"))?;
            if tokens > limits.processing.max_directive_tokens {
                return Err(limit("directive tokens"));
            }
            check_name_bytes(token.as_bytes(), limits)?;
        }
    }
    Ok(bytes)
}

fn validate_duplicate_attributes(
    attrs: &[RawAttribute<'_>],
    limits: &StreamLimits,
) -> Result<(), Error> {
    let mut seen: HashSet<&Name> = HashSet::new();
    try_reserve_set(&mut seen, attrs.len(), "MCE stream attributes")?;
    for attribute in attrs {
        if is_namespace_attribute(attribute.qualified_name.as_ref()) {
            continue;
        }
        check_expanded_name(&attribute.expanded_name, limits)?;
        if !seen.insert(&attribute.expanded_name) {
            return Err(bad("duplicate attribute"));
        }
    }
    Ok(())
}

impl<'a> SemanticEvent<'a> {
    fn from_element(
        qualified_name: &'a [u8],
        expanded_name: &Name,
        empty: bool,
        attrs: Vec<SemanticAttribute<'a>>,
        limits: &StreamLimits,
    ) -> Result<Self, Error> {
        let element = SemanticElement {
            qualified_name: Cow::Borrowed(qualified_name),
            expanded_name: clone_bounded_name(
                expanded_name,
                limits,
                "MCE stream semantic expanded element name",
            )?,
            attributes: attrs,
        };
        Ok(if empty {
            Self::Empty(element)
        } else {
            Self::Start(element)
        })
    }
}

fn invoke_active<ActiveE, Active>(
    callback: &mut Active,
    enabled: &mut bool,
    error: &mut Option<ActiveE>,
    event: SemanticEvent<'_>,
) where
    Active: for<'a> FnMut(SemanticEvent<'a>) -> Result<(), ActiveE>,
{
    if let Err(callback_error) = callback(event) {
        *error = Some(callback_error);
        *enabled = false;
    }
}

fn borrow_raw_attribute<'a>(
    attr: &'a RawAttribute<'_>,
    limits: &StreamLimits,
) -> Result<RawAttribute<'a>, Error> {
    Ok(RawAttribute {
        qualified_name: Cow::Borrowed(attr.qualified_name.as_ref()),
        expanded_name: clone_bounded_name(
            &attr.expanded_name,
            limits,
            "MCE stream raw expanded attribute name",
        )?,
        raw_value: Cow::Borrowed(attr.raw_value.as_ref()),
        decoded_value: Cow::Borrowed(attr.decoded_value.as_ref()),
    })
}

fn clone_raw_attributes<'a>(
    attrs: &'a [RawAttribute<'_>],
    limits: &StreamLimits,
) -> Result<Vec<RawAttribute<'a>>, Error> {
    let mut cloned = Vec::new();
    reserve_vec(&mut cloned, attrs.len(), "MCE stream raw attributes")?;
    for attr in attrs {
        if cloned.len() >= limits.max_attributes_per_event {
            return Err(limit("stream attributes per event"));
        }
        cloned.push(borrow_raw_attribute(attr, limits)?);
    }
    Ok(cloned)
}

fn semantic_attrs<'a>(
    attrs: &'a mut [RawAttribute<'_>],
    context: &Context,
    capabilities: &Capabilities,
    filter: bool,
    limits: &StreamLimits,
    report: &mut StreamReport,
) -> Result<Vec<SemanticAttribute<'a>>, Error> {
    let mut output = Vec::new();
    reserve_vec(&mut output, attrs.len(), "MCE stream semantic attributes")?;
    for attr in attrs.iter_mut() {
        if is_namespace_attribute(attr.qualified_name.as_ref()) {
            continue;
        }
        if filter && attr.expanded_name.namespace == NAMESPACE {
            report.ignored_attributes = report
                .ignored_attributes
                .checked_add(1)
                .ok_or_else(|| limit("stream report counters"))?;
            continue;
        }
        if filter
            && !attr.expanded_name.namespace.is_empty()
            && context.is_ignorable(&attr.expanded_name.namespace)
            && !capabilities.understands(&attr.expanded_name.namespace)
            && !context.preserves_attribute(&attr.expanded_name)
        {
            report.ignored_attributes = report
                .ignored_attributes
                .checked_add(1)
                .ok_or_else(|| limit("stream report counters"))?;
            continue;
        }
        if filter
            && !attr.expanded_name.namespace.is_empty()
            && context.is_ignorable(&attr.expanded_name.namespace)
            && !capabilities.understands(&attr.expanded_name.namespace)
            && context.preserves_attribute(&attr.expanded_name)
        {
            report.preserved_attributes = report
                .preserved_attributes
                .checked_add(1)
                .ok_or_else(|| limit("stream report counters"))?;
        }
        check_expanded_name(&attr.expanded_name, limits)?;
        let expanded_name = std::mem::replace(
            &mut attr.expanded_name,
            Name {
                namespace: String::new(),
                local_name: String::new(),
            },
        );
        output.push(SemanticAttribute {
            qualified_name: Cow::Borrowed(attr.qualified_name.as_ref()),
            expanded_name,
            raw_value: Cow::Borrowed(attr.raw_value.as_ref()),
            decoded_value: Cow::Borrowed(attr.decoded_value.as_ref()),
        });
        if output.len() > limits.max_attributes_per_event {
            return Err(limit("stream attributes per event"));
        }
    }
    Ok(output)
}

fn local_namespaces(
    attrs: &[RawAttribute<'_>],
    limits: &StreamLimits,
) -> Result<Vec<(String, String)>, Error> {
    let mut local = Vec::new();
    for attr in attrs {
        let qualified = str::from_utf8(attr.qualified_name.as_ref()).map_err(xml_error)?;
        if qualified == "xmlns" {
            check_name_bytes(attr.raw_value.as_ref(), limits)?;
            check_name_bytes(attr.decoded_value.as_bytes(), limits)?;
            reserve_vec(&mut local, 1, "MCE stream namespace declarations")?;
            let value = clone_bounded_name_part(
                attr.decoded_value.as_ref(),
                limits,
                "stream namespace URI",
            )?;
            local.push((String::new(), value));
        } else if let Some(prefix) = qualified.strip_prefix("xmlns:") {
            if !xml_name::is_ncname(prefix) || prefix == "xmlns" {
                return Err(bad("invalid namespace"));
            }
            check_name_bytes(prefix.as_bytes(), limits)?;
            let value = attr.decoded_value.as_ref();
            if value.is_empty() || (prefix == "xml" && value != XML_NS) {
                return Err(bad("invalid namespace"));
            }
            check_name_bytes(value.as_bytes(), limits)?;
            reserve_vec(&mut local, 1, "MCE stream namespace declarations")?;
            let prefix = clone_bounded_name_part(prefix, limits, "stream namespace prefix")?;
            let value = clone_bounded_name_part(value, limits, "stream namespace URI")?;
            local.push((prefix, value));
        }
    }
    Ok(local)
}

#[derive(Clone, Copy)]
enum DirectiveKind {
    Ignorable,
    ProcessContent,
    PreserveElements,
    PreserveAttributes,
    MustUnderstand,
}

struct DirectiveValue {
    kind: DirectiveKind,
    value: String,
}

struct DirectiveValues {
    values: Vec<DirectiveValue>,
    tokens: usize,
}

fn collect_directives(
    attrs: &[RawAttribute<'_>],
    limits: &StreamLimits,
) -> Result<DirectiveValues, Error> {
    let mut values = Vec::new();
    let mut tokens = 0usize;
    for attr in attrs {
        if is_namespace_attribute(attr.qualified_name.as_ref())
            || attr.expanded_name.namespace != NAMESPACE
        {
            continue;
        }
        let kind = match attr.expanded_name.local_name.as_str() {
            "Ignorable" => DirectiveKind::Ignorable,
            "ProcessContent" => DirectiveKind::ProcessContent,
            "PreserveElements" => DirectiveKind::PreserveElements,
            "PreserveAttributes" => DirectiveKind::PreserveAttributes,
            "MustUnderstand" => DirectiveKind::MustUnderstand,
            _ => return Err(bad("unknown MCE attribute")),
        };
        tokens = tokens
            .checked_add(attr.decoded_value.split_whitespace().count())
            .ok_or_else(|| limit("directive tokens"))?;
        if tokens > limits.processing.max_directive_tokens {
            return Err(limit("directive tokens"));
        }
        reserve_vec(&mut values, 1, "MCE stream directives")?;
        values.push(DirectiveValue {
            kind,
            value: clone_bounded_string(
                attr.decoded_value.as_ref(),
                limits,
                "stream directive value",
            )?,
        });
    }
    Ok(DirectiveValues { values, tokens })
}

fn apply_directives(
    context: &mut Context,
    values: DirectiveValues,
    capabilities: &Capabilities,
    limits: &StreamLimits,
) -> Result<(), Error> {
    let additional_context_bytes = directive_bytes(&values)?;
    if context
        .bytes
        .checked_add(additional_context_bytes)
        .is_none_or(|bytes| bytes > limits.max_context_bytes)
    {
        return Err(limit("stream context bytes"));
    }
    let mut local_ignorable = HashSet::new();
    let mut seen = HashSet::new();
    for directive in &values.values {
        if !matches!(directive.kind, DirectiveKind::Ignorable) {
            continue;
        }
        for prefix in directive.value.split_whitespace() {
            if !xml_name::is_ncname(prefix) {
                return Err(bad("invalid or duplicate Ignorable prefix"));
            }
            try_reserve_set(&mut seen, 1, "MCE stream Ignorable directives")?;
            if !seen.insert(prefix) {
                return Err(bad("invalid or duplicate Ignorable prefix"));
            }
            check_name_bytes(prefix.as_bytes(), limits)?;
            let uri = context
                .ns
                .get(prefix)
                .ok_or_else(|| bad("unbound Ignorable prefix"))?;
            if uri == NAMESPACE {
                return Err(bad("MCE cannot be ignorable"));
            }
            try_reserve_set(&mut local_ignorable, 1, "MCE stream Ignorable directives")?;
            local_ignorable.insert(clone_bounded_name_part(
                uri,
                limits,
                "stream Ignorable namespace",
            )?);
        }
    }
    let mut new_ignorable = HashSet::new();
    try_reserve_set(
        &mut new_ignorable,
        local_ignorable.len(),
        "MCE stream effective Ignorable directives",
    )?;
    for namespace in &local_ignorable {
        if !context.is_ignorable(namespace) {
            new_ignorable.insert(clone_bounded_name_part(
                namespace,
                limits,
                "MCE stream Ignorable namespace",
            )?);
        }
    }

    let mut process = HashSet::new();
    let mut preserve_elements = HashSet::new();
    let mut preserve_attributes = HashSet::new();
    for directive in &values.values {
        match directive.kind {
            DirectiveKind::Ignorable => {},
            DirectiveKind::ProcessContent => {
                for token in directive.value.split_whitespace() {
                    let target = parse_target(token, &context.ns, true, limits)?;
                    if !local_ignorable.contains(pattern_namespace(&target))
                        && !context.is_ignorable(pattern_namespace(&target))
                    {
                        return Err(bad("ProcessContent target is not effectively ignorable"));
                    }
                    try_reserve_set(&mut process, 1, "MCE stream ProcessContent directives")?;
                    if !process.insert(target) {
                        return Err(bad("duplicate ProcessContent target"));
                    }
                }
            },
            DirectiveKind::PreserveElements => {
                for token in directive.value.split_whitespace() {
                    let target = parse_target(token, &context.ns, true, limits)?;
                    if !local_ignorable.contains(pattern_namespace(&target)) {
                        return Err(bad("PreserveElements target is not locally ignorable"));
                    }
                    try_reserve_set(
                        &mut preserve_elements,
                        1,
                        "MCE stream PreserveElements directives",
                    )?;
                    if !preserve_elements.insert(target) {
                        return Err(bad("duplicate PreserveElements target"));
                    }
                }
            },
            DirectiveKind::PreserveAttributes => {
                for token in directive.value.split_whitespace() {
                    let target = parse_target(token, &context.ns, true, limits)?;
                    if !local_ignorable.contains(pattern_namespace(&target)) {
                        return Err(bad("PreserveAttributes target is not locally ignorable"));
                    }
                    try_reserve_set(
                        &mut preserve_attributes,
                        1,
                        "MCE stream PreserveAttributes directives",
                    )?;
                    if !preserve_attributes.insert(target) {
                        return Err(bad("duplicate PreserveAttributes target"));
                    }
                }
            },
            DirectiveKind::MustUnderstand => {
                let mut seen = HashSet::new();
                for prefix in directive.value.split_whitespace() {
                    if !xml_name::is_ncname(prefix) {
                        return Err(bad("invalid or duplicate MustUnderstand prefix"));
                    }
                    try_reserve_set(&mut seen, 1, "MCE stream MustUnderstand directives")?;
                    if !seen.insert(prefix) {
                        return Err(bad("invalid or duplicate MustUnderstand prefix"));
                    }
                    check_name_bytes(prefix.as_bytes(), limits)?;
                    let namespace = context
                        .ns
                        .get(prefix)
                        .ok_or_else(|| bad("unbound MustUnderstand prefix"))?;
                    if !capabilities.understands(namespace) {
                        return Err(Error::MustUnderstand(clone_bounded_name_part(
                            namespace,
                            limits,
                            "MustUnderstand namespace",
                        )?));
                    }
                }
            },
        }
    }

    if !new_ignorable.is_empty()
        || !process.is_empty()
        || !preserve_elements.is_empty()
        || !preserve_attributes.is_empty()
    {
        context.directives = Some(Arc::new(DirectiveLayer {
            parent: context.directives.take(),
            ignorable: new_ignorable,
            process,
            preserve_elements,
            preserve_attributes,
        }));
    }
    context.bytes = context
        .bytes
        .checked_add(additional_context_bytes)
        .ok_or_else(|| limit("stream context bytes"))?;
    if context.bytes > limits.max_context_bytes {
        return Err(limit("stream context bytes"));
    }
    Ok(())
}

fn validate_alternate_attributes_if_needed(
    data: &ElementData<'_>,
    context: &Context,
    capabilities: &Capabilities,
) -> Result<(), Error> {
    if data.expanded_name.namespace != NAMESPACE {
        return Ok(());
    }
    let kind = match data.expanded_name.local_name.as_str() {
        "AlternateContent" => 0u8,
        "Choice" => 1u8,
        "Fallback" => 2u8,
        _ => return Ok(()),
    };
    for attr in &data.attrs {
        if is_namespace_attribute(attr.qualified_name.as_ref()) {
            continue;
        }
        let name = &attr.expanded_name;
        if name.namespace.is_empty() {
            if kind == 1 && name.local_name == "Requires" {
                continue;
            }
            return Err(bad("unexpected unprefixed AlternateContent attribute"));
        }
        if name.namespace == XML_NS && matches!(name.local_name.as_str(), "lang" | "space") {
            return Err(bad(
                "xml:lang and xml:space are forbidden on AlternateContent markup",
            ));
        }
        if name.namespace == NAMESPACE || name.namespace == XML_NS {
            continue;
        }
        if !capabilities.understands(&name.namespace) && !context.is_ignorable(&name.namespace) {
            return Err(bad(
                "AlternateContent attribute namespace is neither understood nor ignorable",
            ));
        }
    }
    Ok(())
}

fn semantic_attr_name<'a>(attr: &'a RawAttribute<'_>) -> &'a [u8] {
    attr.qualified_name.as_ref()
}

fn attr_value<'a>(attrs: &'a [RawAttribute<'_>], name: &str) -> Result<Option<&'a str>, Error> {
    let mut value = None;
    for attr in attrs {
        if attr.expanded_name.namespace.is_empty()
            && str::from_utf8(semantic_attr_name(attr)).map_err(xml_error)? == name
        {
            if value.is_some() {
                return Err(bad("duplicate attribute"));
            }
            value = Some(attr.decoded_value.as_ref());
        }
    }
    Ok(value)
}

fn is_namespace_attribute(name: &[u8]) -> bool {
    name == b"xmlns" || name.starts_with(b"xmlns:")
}

fn expand_attribute(q: &[u8], ns: &Namespaces, limits: &StreamLimits) -> Result<Name, Error> {
    let lexical = str::from_utf8(q).map_err(xml_error)?;
    if lexical == "xmlns" {
        return Ok(Name {
            namespace: clone_bounded_name_part(
                XMLNS_NAMESPACE,
                limits,
                "MCE stream namespace name",
            )?,
            local_name: clone_bounded_name_part("xmlns", limits, "MCE stream attribute name")?,
        });
    }
    if let Some(prefix) = lexical.strip_prefix("xmlns:") {
        return Ok(Name {
            namespace: clone_bounded_name_part(
                XMLNS_NAMESPACE,
                limits,
                "MCE stream namespace name",
            )?,
            local_name: clone_bounded_name_part(prefix, limits, "MCE stream namespace prefix")?,
        });
    }
    expand(lexical, ns, false, limits)
}

fn expand(
    q: &str,
    namespaces: &Namespaces,
    element: bool,
    limits: &StreamLimits,
) -> Result<Name, Error> {
    let qualified = QualifiedName::try_from(q).map_err(|_| bad("invalid QName"))?;
    let prefix = qualified.prefix().unwrap_or_default();
    let local = qualified.local();
    let namespace = if prefix.is_empty() {
        if element {
            clone_bounded_name_part(
                namespaces.get("").unwrap_or_default(),
                limits,
                "MCE stream namespace URI",
            )?
        } else {
            String::new()
        }
    } else {
        namespaces
            .get(prefix)
            .map(|value| clone_bounded_name_part(value, limits, "MCE stream namespace URI"))
            .transpose()?
            .ok_or_else(|| bad("unbound prefix"))?
    };
    Ok(Name {
        namespace,
        local_name: clone_bounded_name_part(local, limits, "MCE stream local name")?,
    })
}

fn parse_target(
    token: &str,
    namespaces: &Namespaces,
    wildcard: bool,
    limits: &StreamLimits,
) -> Result<NamePattern, Error> {
    check_name_bytes(token.as_bytes(), limits)?;
    let (prefix, local) = token
        .split_once(':')
        .ok_or_else(|| bad("preservation and processing targets must be prefixed QNames"))?;
    if token.matches(':').count() != 1 || !xml_name::is_ncname(prefix) || local.is_empty() {
        return Err(bad("invalid compatibility target QName"));
    }
    check_name_bytes(prefix.as_bytes(), limits)?;
    check_name_bytes(local.as_bytes(), limits)?;
    let namespace = namespaces
        .get(prefix)
        .map(|value| clone_bounded_name_part(value, limits, "MCE stream namespace URI"))
        .transpose()?
        .ok_or_else(|| bad("unbound compatibility target prefix"))?;
    if namespace == NAMESPACE {
        return Err(bad("compatibility target cannot use the MCE namespace"));
    }
    if local == "*" {
        return wildcard
            .then_some(NamePattern::Namespace(namespace))
            .ok_or_else(|| bad("wildcard is not allowed in this compatibility directive"));
    }
    if !xml_name::is_ncname(local) {
        return Err(bad("invalid compatibility target QName"));
    }
    Ok(NamePattern::Exact(Name {
        namespace,
        local_name: clone_bounded_name_part(local, limits, "MCE stream target local name")?,
    }))
}

fn pattern_namespace(pattern: &NamePattern) -> &str {
    match pattern {
        NamePattern::Exact(name) => &name.namespace,
        NamePattern::Namespace(namespace) => namespace,
    }
}

fn matches_pattern(patterns: &HashSet<NamePattern>, name: &Name) -> bool {
    patterns.iter().any(|pattern| match pattern {
        NamePattern::Exact(candidate) => candidate == name,
        NamePattern::Namespace(namespace) => namespace == &name.namespace,
    })
}

fn check_name_bytes(name: &[u8], limits: &StreamLimits) -> Result<(), Error> {
    if name.len() > limits.max_name_bytes {
        return Err(limit("stream name bytes"));
    }
    Ok(())
}

fn check_expanded_name(name: &Name, limits: &StreamLimits) -> Result<(), Error> {
    if name
        .namespace
        .len()
        .checked_add(name.local_name.len())
        .is_none_or(|bytes| bytes > limits.max_name_bytes)
    {
        return Err(limit("stream name bytes"));
    }
    Ok(())
}

fn clone_bounded_string(
    value: &str,
    limits: &StreamLimits,
    resource: &'static str,
) -> Result<String, Error> {
    if value.len() > limits.max_context_bytes {
        return Err(limit("stream context bytes"));
    }
    let mut cloned = String::new();
    cloned
        .try_reserve_exact(value.len())
        .map_err(|source| Error::Allocation { resource, source })?;
    cloned.push_str(value);
    Ok(cloned)
}

fn clone_bounded_bytes<'a>(
    value: &[u8],
    limits: &StreamLimits,
    resource: &'static str,
) -> Result<Cow<'a, [u8]>, Error> {
    if value.len() > limits.max_event_bytes {
        return Err(limit("stream event bytes"));
    }
    let mut cloned = Vec::new();
    cloned
        .try_reserve_exact(value.len())
        .map_err(|source| Error::Allocation { resource, source })?;
    cloned.extend_from_slice(value);
    Ok(Cow::Owned(cloned))
}

fn clone_bounded_text<'a>(
    value: &str,
    limits: &StreamLimits,
    resource: &'static str,
) -> Result<Cow<'a, str>, Error> {
    if value.len() > limits.max_event_bytes {
        return Err(limit("stream event bytes"));
    }
    let mut cloned = String::new();
    cloned
        .try_reserve_exact(value.len())
        .map_err(|source| Error::Allocation { resource, source })?;
    cloned.push_str(value);
    Ok(Cow::Owned(cloned))
}

fn clone_bounded_name_part(
    value: &str,
    limits: &StreamLimits,
    resource: &'static str,
) -> Result<String, Error> {
    if value.len() > limits.max_name_bytes {
        return Err(limit("stream name bytes"));
    }
    let mut cloned = String::new();
    cloned
        .try_reserve_exact(value.len())
        .map_err(|source| Error::Allocation { resource, source })?;
    cloned.push_str(value);
    Ok(cloned)
}

fn clone_bounded_name(
    value: &Name,
    limits: &StreamLimits,
    resource: &'static str,
) -> Result<Name, Error> {
    check_expanded_name(value, limits)?;
    Ok(Name {
        namespace: clone_bounded_name_part(&value.namespace, limits, resource)?,
        local_name: clone_bounded_name_part(&value.local_name, limits, resource)?,
    })
}

fn directive_bytes(values: &DirectiveValues) -> Result<usize, Error> {
    values.values.iter().try_fold(0usize, |sum, value| {
        sum.checked_add(value.value.len())
            .and_then(|sum| sum.checked_add(1))
            .ok_or_else(|| limit("stream context bytes"))
    })
}

fn reserve_vec<T>(
    values: &mut Vec<T>,
    additional: usize,
    resource: &'static str,
) -> Result<(), Error> {
    values
        .try_reserve_exact(additional)
        .map_err(|source| Error::Allocation { resource, source })
}

fn try_reserve_set<T: std::hash::Hash + Eq>(
    values: &mut HashSet<T>,
    additional: usize,
    resource: &'static str,
) -> Result<(), Error> {
    values
        .try_reserve(additional)
        .map_err(|source| Error::Allocation { resource, source })
}

fn bad(message: impl Into<String>) -> Error {
    Error::NonConformant(message.into())
}

fn limit(message: &str) -> Error {
    Error::LimitExceeded(message.to_owned())
}

fn xml_error(error: impl fmt::Display) -> Error {
    Error::Xml(error.to_string())
}
