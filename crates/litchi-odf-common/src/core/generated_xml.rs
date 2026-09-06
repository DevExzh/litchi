//! Bounded publication of generated XML members.
//!
//! This module owns only the format-neutral lexical seam. A format provider
//! supplies an immutable envelope and emits complete child elements through
//! the callback passed to GeneratedXmlReader. Namespace and schema decisions
//! stay with the provider. The common audit checks lexical XML syntax and
//! bounded accounting; it does not resolve namespace bindings. Providers must
//! declare every prefix/default namespace needed by their envelope and
//! fragments, keep those bindings stable across the insertion boundary, and
//! enforce their own namespace and schema contract for emitted elements.

use litchi_core::{Error, Result};
use quick_xml::{Reader, events::Event};
use std::fmt;
use std::io::{self, Read, Write};

/// Caller-selected XML resource limits for generated member publication.
///
/// This is an alias so format providers do not need a direct dependency on
/// xml-minifier merely to configure the common writer seam.
pub type GeneratedXmlLimits = xml_minifier::audit::Limits;

/// XML resource attributed by a generated-member limit failure.
pub type GeneratedXmlLimitResource = xml_minifier::audit::Resource;

/// Typed lexical or composed-document XML limit attribution.
#[derive(Clone, Copy, Debug, Eq, PartialEq)]
pub struct GeneratedXmlLimitExceeded {
    resource: GeneratedXmlLimitResource,
    actual: usize,
    maximum: usize,
}

impl GeneratedXmlLimitExceeded {
    /// The exhausted XML resource.
    #[must_use]
    pub const fn resource(self) -> GeneratedXmlLimitResource {
        self.resource
    }

    /// First observed value exceeding the configured inclusive ceiling.
    #[must_use]
    pub const fn actual(self) -> usize {
        self.actual
    }

    /// Configured inclusive ceiling.
    #[must_use]
    pub const fn maximum(self) -> usize {
        self.maximum
    }
}

impl fmt::Display for GeneratedXmlLimitExceeded {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        write!(
            formatter,
            "generated XML {:?} limit exceeded: {} > {}",
            self.resource, self.actual, self.maximum
        )
    }
}

impl std::error::Error for GeneratedXmlLimitExceeded {}

fn xml_limit(resource: GeneratedXmlLimitResource, actual: usize, maximum: usize) -> Error {
    // The reader callback uses the core Result contract. Preserve attribution
    // through that boundary and std::io::Read without parsing diagnostics.
    Error::Io(io::Error::other(GeneratedXmlLimitExceeded {
        resource,
        actual,
        maximum,
    }))
}

/// A checked XML shell with one insertion point between its prefix and suffix.
///
/// The prefix and suffix are owned after construction. [`Self::try_new`] checks
/// a prefix containing only an optional declaration followed by open elements.
/// [`Self::try_new_with_prelude`] also allows balanced fixed element subtrees
/// before the final open insertion path. Both require the suffix to close the
/// remaining open elements in reverse order.
#[derive(Debug)]
pub struct GeneratedXmlEnvelope {
    prefix: Vec<u8>,
    suffix: Vec<u8>,
    insertion_depth: usize,
}

impl GeneratedXmlEnvelope {
    /// Checks and owns a generated XML shell.
    ///
    /// The prefix plus suffix must already be a complete compact XML document.
    /// This is also the output used when the producer emits zero fragments.
    pub fn try_new(prefix: &[u8], suffix: &[u8]) -> Result<Self> {
        let shell_bytes = prefix
            .len()
            .checked_add(suffix.len())
            .ok_or_else(|| invalid_generated_xml("generated XML shell byte count overflow"))?;
        if shell_bytes > xml_minifier::audit::Limits::BYTE_CEILING {
            return Err(invalid_generated_xml(
                "generated XML envelope exceeds its immutable byte ceiling",
            ));
        }
        let combined = concatenate(prefix, suffix, "generated XML envelope shape")?;
        let hard_limits = immutable_hard_limits()?;
        let _ =
            xml_minifier::audit::verify_authored(&combined, hard_limits).map_err(audit_error)?;
        let insertion_depth = parse_envelope_shape(&combined, prefix.len())?;

        Ok(Self {
            prefix: owned_bytes(prefix, "generated XML envelope prefix")?,
            suffix: owned_bytes(suffix, "generated XML envelope suffix")?,
            insertion_depth,
        })
    }

    /// Checks and owns a generated XML shell with a fixed balanced prelude.
    ///
    /// The prefix may contain an optional declaration, one document root,
    /// balanced fixed element subtrees, and a final open insertion path. The
    /// suffix must contain only the matching end tags for the elements still
    /// open at the prefix boundary. Fixed text, CDATA, references, comments,
    /// processing instructions, doctypes, and `xml:space` are refused. Names
    /// are matched as raw qualified names; namespace URI and schema policy
    /// remain with the format provider.
    ///
    /// This is an opt-in extension for providers that need fixed child
    /// subtrees before the streamed fragments. [`Self::try_new`] retains its
    /// strict prefix-only-start contract.
    ///
    /// Fixed children may occur under any ancestor left open in the prefix.
    /// The final insertion path is the trailing sequence of start tags after
    /// the last fixed child closes; at least one such start tag is required.
    /// For example, `<root><slot><fixed/><inner>` inserts inside `inner` at
    /// depth three. Fixed character data is forbidden, so shell text-byte
    /// accounting is zero and all character data comes from fragments.
    pub fn try_new_with_prelude(prefix: &[u8], suffix: &[u8]) -> Result<Self> {
        let shell_bytes = prefix
            .len()
            .checked_add(suffix.len())
            .ok_or_else(|| invalid_generated_xml("generated XML shell byte count overflow"))?;
        if shell_bytes > xml_minifier::audit::Limits::BYTE_CEILING {
            return Err(invalid_generated_xml(
                "generated XML envelope exceeds its immutable byte ceiling",
            ));
        }
        let combined = concatenate(prefix, suffix, "generated XML envelope shape")?;
        let hard_limits = immutable_hard_limits()?;
        let _ =
            xml_minifier::audit::verify_authored(&combined, hard_limits).map_err(audit_error)?;
        let insertion_depth = parse_prelude_envelope_shape(&combined, prefix.len())?;

        Ok(Self {
            prefix: owned_bytes(prefix, "generated XML envelope prefix")?,
            suffix: owned_bytes(suffix, "generated XML envelope suffix")?,
            insertion_depth,
        })
    }

    pub(crate) fn prepare<F>(
        self,
        limits: GeneratedXmlLimits,
        max_fragment_bytes: usize,
        produce: F,
    ) -> Result<GeneratedXmlReader<F>>
    where
        F: FnMut(&mut dyn Write) -> Result<bool>,
    {
        GeneratedXmlReader::new(self, limits, max_fragment_bytes, produce)
    }
}

/// Accounting returned after a generated XML member has been published.
///
/// The event count has the same convention as xml-minifier's audit report:
/// it includes the final document EOF event.
#[derive(Clone, Copy, Debug, Eq, PartialEq)]
pub struct GeneratedXmlReport {
    attributes: usize,
    bytes: usize,
    events: usize,
    max_depth: usize,
    text_bytes: usize,
    fragments: usize,
}

impl GeneratedXmlReport {
    /// Total attributes in the composed XML document.
    #[must_use]
    pub const fn attributes(self) -> usize {
        self.attributes
    }

    /// Total XML bytes in the composed document.
    #[must_use]
    pub const fn bytes(self) -> usize {
        self.bytes
    }

    /// Parser events, including the final document EOF event.
    #[must_use]
    pub const fn events(self) -> usize {
        self.events
    }

    /// Greatest composed element nesting depth.
    #[must_use]
    pub const fn max_depth(self) -> usize {
        self.max_depth
    }

    /// Aggregate character-data bytes.
    #[must_use]
    pub const fn text_bytes(self) -> usize {
        self.text_bytes
    }

    /// Number of nonempty producer fragments.
    #[must_use]
    pub const fn fragments(self) -> usize {
        self.fragments
    }
}

/// A bounded reader that composes an envelope and independently audited
/// callback fragments without retaining the complete XML member.
pub(crate) struct GeneratedXmlReader<F> {
    envelope: GeneratedXmlEnvelope,
    limits: GeneratedXmlLimits,
    fragment: FragmentBuffer,
    produce: F,
    report: GeneratedXmlReport,
    state: ReaderState,
    offset: usize,
    fragment_ready: bool,
    first_prefetched: bool,
    failed: bool,
}

#[derive(Clone, Copy, Debug, Eq, PartialEq)]
enum ReaderState {
    Prefix,
    Fragment,
    NeedFragment,
    Suffix,
    Done,
}

impl<F> GeneratedXmlReader<F>
where
    F: FnMut(&mut dyn Write) -> Result<bool>,
{
    fn new(
        envelope: GeneratedXmlEnvelope,
        limits: GeneratedXmlLimits,
        max_fragment_bytes: usize,
        produce: F,
    ) -> Result<Self> {
        validate_limits(limits, max_fragment_bytes)?;
        let base = envelope.audit_shell(limits)?;
        let fragment = FragmentBuffer::new(max_fragment_bytes)?;
        let report = GeneratedXmlReport {
            attributes: base.attributes(),
            bytes: base.bytes(),
            events: base.events(),
            max_depth: base.max_depth(),
            text_bytes: base.text_bytes(),
            fragments: 0,
        };
        Ok(Self {
            envelope,
            limits,
            fragment,
            produce,
            report,
            state: ReaderState::Prefix,
            offset: 0,
            fragment_ready: false,
            first_prefetched: false,
            failed: false,
        })
    }

    /// Invokes the first producer callback when the reader is first consumed.
    /// An empty false result is valid and selects the direct shell. The
    /// PackageWriter lets the ZIP reader perform this call after entry
    /// admission; direct low-level callers may prefetch when they need a
    /// recoverable preflight boundary.
    pub(crate) fn prefetch_first(&mut self) -> Result<()> {
        if !self.first_prefetched {
            self.first_prefetched = true;
            if let Err(error) = self.fill_fragment() {
                self.failed = true;
                return Err(error);
            }
        }
        Ok(())
    }

    pub(crate) fn report(&self) -> GeneratedXmlReport {
        self.report
    }

    fn fill_fragment(&mut self) -> Result<()> {
        self.fragment.clear();
        let produced = (self.produce)(&mut self.fragment);

        // Preserve the producer's core error if it returned one, including
        // errors raised after a write attempt. A successful callback must
        // still observe any ignored scratch failure below.
        let produced = produced?;
        if let Some(error) = self.fragment.failure() {
            return Err(Error::Io(error));
        }

        if !produced {
            if !self.fragment.is_empty() {
                return Err(invalid_generated_xml(
                    "producer returned EOF after writing a nonempty fragment",
                ));
            }
            self.fragment_ready = false;
            return Ok(());
        }
        if self.fragment.is_empty() {
            return Err(invalid_generated_xml(
                "producer returned a non-EOF result without a fragment",
            ));
        }

        let fragment_report =
            xml_minifier::audit::verify_authored(self.fragment.as_bytes(), self.limits)
                .map_err(audit_error)?;
        validate_fragment_shape(self.fragment.as_bytes())?;
        self.add_fragment(fragment_report)?;
        self.fragment_ready = true;
        Ok(())
    }

    fn add_fragment(&mut self, fragment: xml_minifier::audit::Report) -> Result<()> {
        let fragment_events = fragment
            .events()
            .checked_sub(1)
            .ok_or_else(|| invalid_generated_xml("fragment audit did not report an EOF event"))?;
        let events = self
            .report
            .events
            .checked_add(fragment_events)
            .ok_or_else(|| invalid_generated_xml("generated XML event count overflow"))?;
        let bytes = self
            .report
            .bytes
            .checked_add(fragment.bytes())
            .ok_or_else(|| invalid_generated_xml("generated XML byte count overflow"))?;
        let attributes = self
            .report
            .attributes
            .checked_add(fragment.attributes())
            .ok_or_else(|| invalid_generated_xml("generated XML attribute count overflow"))?;
        let text_bytes = self
            .report
            .text_bytes
            .checked_add(fragment.text_bytes())
            .ok_or_else(|| invalid_generated_xml("generated XML text count overflow"))?;
        let depth = self
            .envelope
            .insertion_depth
            .checked_add(fragment.max_depth())
            .ok_or_else(|| invalid_generated_xml("generated XML depth overflow"))?;
        let next =
            GeneratedXmlReport {
                attributes,
                bytes,
                events,
                max_depth: self.report.max_depth.max(depth),
                text_bytes,
                fragments: self.report.fragments.checked_add(1).ok_or_else(|| {
                    invalid_generated_xml("generated XML fragment count overflow")
                })?,
            };
        check_report_limits(next, self.limits)?;
        self.report = next;
        Ok(())
    }

    fn fail_io(&mut self, error: Error) -> io::Error {
        self.failed = true;
        io::Error::other(GeneratedXmlIoError(error))
    }
}

impl<F> Read for GeneratedXmlReader<F>
where
    F: FnMut(&mut dyn Write) -> Result<bool>,
{
    fn read(&mut self, output: &mut [u8]) -> io::Result<usize> {
        if output.is_empty() {
            return Ok(0);
        }
        if self.failed {
            return Err(io::Error::other("generated XML reader is failed"));
        }
        if let Err(error) = self.prefetch_first() {
            return Err(self.fail_io(error));
        }

        loop {
            match self.state {
                ReaderState::Prefix => {
                    let bytes = &self.envelope.prefix;
                    let amount = copy_pending(&mut self.offset, output, bytes);
                    if amount != 0 {
                        return Ok(amount);
                    }
                    self.offset = 0;
                    self.state = if self.fragment_ready {
                        ReaderState::Fragment
                    } else {
                        ReaderState::Suffix
                    };
                },
                ReaderState::Fragment => {
                    let bytes = self.fragment.as_bytes();
                    let amount = copy_pending(&mut self.offset, output, bytes);
                    if amount != 0 {
                        return Ok(amount);
                    }
                    self.offset = 0;
                    self.fragment_ready = false;
                    self.state = ReaderState::NeedFragment;
                },
                ReaderState::NeedFragment => {
                    if let Err(error) = self.fill_fragment() {
                        return Err(self.fail_io(error));
                    }
                    self.offset = 0;
                    self.state = if self.fragment_ready {
                        ReaderState::Fragment
                    } else {
                        ReaderState::Suffix
                    };
                },
                ReaderState::Suffix => {
                    let bytes = &self.envelope.suffix;
                    let amount = copy_pending(&mut self.offset, output, bytes);
                    if amount != 0 {
                        return Ok(amount);
                    }
                    self.offset = 0;
                    self.state = ReaderState::Done;
                },
                ReaderState::Done => return Ok(0),
            }
        }
    }
}

fn copy_pending(offset: &mut usize, output: &mut [u8], bytes: &[u8]) -> usize {
    let remaining = bytes.len().saturating_sub(*offset);
    let amount = remaining.min(output.len());
    if amount != 0 {
        output[..amount].copy_from_slice(&bytes[*offset..*offset + amount]);
        *offset += amount;
    }
    amount
}

impl GeneratedXmlEnvelope {
    fn audit_shell(&self, limits: GeneratedXmlLimits) -> Result<xml_minifier::audit::Report> {
        let bytes = self
            .prefix
            .len()
            .checked_add(self.suffix.len())
            .ok_or_else(|| invalid_generated_xml("generated XML shell byte count overflow"))?;
        if bytes > limits.max_bytes() {
            return Err(xml_limit(
                GeneratedXmlLimitResource::Bytes,
                bytes,
                limits.max_bytes(),
            ));
        }
        let shell = concatenate(
            &self.prefix,
            &self.suffix,
            "generated XML envelope audit window",
        )?;
        xml_minifier::audit::verify_authored(&shell, limits).map_err(audit_error)
    }
}

#[derive(Debug)]
struct FragmentBuffer {
    bytes: Vec<u8>,
    maximum: usize,
    failure: Option<BufferFailure>,
}

#[derive(Clone, Copy, Debug, Eq, PartialEq)]
struct BufferFailure {
    attempted: usize,
    maximum: usize,
}

impl fmt::Display for BufferFailure {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        write!(
            formatter,
            "generated XML fragment exceeds {} bytes (attempted {})",
            self.maximum, self.attempted
        )
    }
}

impl std::error::Error for BufferFailure {}

impl FragmentBuffer {
    fn new(maximum: usize) -> Result<Self> {
        let mut bytes = Vec::new();
        bytes
            .try_reserve_exact(maximum)
            .map_err(|source| Error::Allocation {
                resource: "generated XML fragment buffer",
                source,
            })?;
        Ok(Self {
            bytes,
            maximum,
            failure: None,
        })
    }

    fn clear(&mut self) {
        self.bytes.clear();
        self.failure = None;
    }

    fn is_empty(&self) -> bool {
        self.bytes.is_empty()
    }

    fn as_bytes(&self) -> &[u8] {
        &self.bytes
    }

    fn failure(&self) -> Option<io::Error> {
        self.failure
            .map(|failure| io::Error::new(io::ErrorKind::WriteZero, failure))
    }
}

impl Write for FragmentBuffer {
    fn write(&mut self, input: &[u8]) -> io::Result<usize> {
        if let Some(failure) = self.failure {
            return Err(io::Error::new(io::ErrorKind::WriteZero, failure));
        }
        let attempted = self.bytes.len().checked_add(input.len()).ok_or_else(|| {
            let failure = BufferFailure {
                attempted: usize::MAX,
                maximum: self.maximum,
            };
            self.failure = Some(failure);
            io::Error::new(io::ErrorKind::InvalidData, failure)
        })?;
        let remaining = self.maximum.saturating_sub(self.bytes.len());
        let accepted = remaining.min(input.len());
        if accepted != 0 {
            self.bytes.extend_from_slice(&input[..accepted]);
        }
        if attempted > self.maximum {
            let failure = BufferFailure {
                attempted,
                maximum: self.maximum,
            };
            self.failure = Some(failure);
            if accepted == 0 {
                return Err(io::Error::new(io::ErrorKind::WriteZero, failure));
            }
        }
        Ok(accepted)
    }

    fn flush(&mut self) -> io::Result<()> {
        if let Some(failure) = self.failure {
            Err(io::Error::new(io::ErrorKind::WriteZero, failure))
        } else {
            Ok(())
        }
    }
}

#[derive(Debug)]
struct GeneratedXmlIoError(Error);

impl fmt::Display for GeneratedXmlIoError {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        self.0.fmt(formatter)
    }
}

impl std::error::Error for GeneratedXmlIoError {
    fn source(&self) -> Option<&(dyn std::error::Error + 'static)> {
        Some(&self.0)
    }
}

fn parse_envelope_shape(bytes: &[u8], boundary: usize) -> Result<usize> {
    if boundary > bytes.len() {
        return Err(invalid_generated_xml(
            "envelope prefix boundary exceeds the combined shell",
        ));
    }
    let text = std::str::from_utf8(bytes)
        .map_err(|error| invalid_generated_xml(format!("envelope is not UTF-8: {error}")))?;
    let mut reader = Reader::from_str(text);
    reader.config_mut().trim_text(false);
    reader.config_mut().check_end_names = false;
    let mut stack = Vec::new();
    let mut insertion_depth = 0usize;
    let mut seen = false;
    loop {
        let start_offset = usize::try_from(reader.buffer_position()).unwrap_or(usize::MAX);
        let event = reader
            .read_event()
            .map_err(|error| invalid_generated_xml(format!("invalid envelope: {error}")))?;
        let end_offset = usize::try_from(reader.buffer_position()).unwrap_or(usize::MAX);
        if start_offset < boundary && end_offset > boundary {
            return Err(invalid_generated_xml(
                "envelope boundary splits an XML event",
            ));
        }
        match event {
            Event::Decl(_) if start_offset == 0 && !seen && bytes.starts_with(b"<?xml") => {
                seen = true;
            },
            Event::Start(start) => {
                if start_offset >= boundary || end_offset > boundary {
                    return Err(invalid_generated_xml(
                        "envelope suffix may contain only end tags",
                    ));
                }
                seen = true;
                reject_xml_space(&start)?;
                let next_depth = insertion_depth
                    .checked_add(1)
                    .ok_or_else(|| invalid_generated_xml("envelope depth overflow"))?;
                if next_depth > xml_minifier::audit::Limits::DEPTH_CEILING {
                    return Err(invalid_generated_xml(
                        "envelope exceeds its immutable depth ceiling",
                    ));
                }
                let name = owned_name(start.name().as_ref())?;
                stack.try_reserve(1).map_err(|source| Error::Allocation {
                    resource: "generated XML envelope depth",
                    source,
                })?;
                stack.push(name);
                insertion_depth = next_depth;
            },
            Event::End(end) => {
                if start_offset < boundary || end_offset <= boundary {
                    return Err(invalid_generated_xml(
                        "envelope prefix may contain only start tags",
                    ));
                }
                let expected = stack
                    .pop()
                    .ok_or_else(|| invalid_generated_xml("envelope has an extra end tag"))?;
                if expected.as_slice() != end.name().as_ref() {
                    return Err(invalid_generated_xml(
                        "envelope end tag does not match its start tag",
                    ));
                }
            },
            Event::Eof => {
                if start_offset != bytes.len() || end_offset != bytes.len() {
                    return Err(invalid_generated_xml("envelope EOF position is invalid"));
                }
                break;
            },
            _ => {
                return Err(invalid_generated_xml(
                    "envelope must contain a declaration/start prefix and end-tag suffix",
                ));
            },
        }
    }
    if insertion_depth == 0 {
        return Err(invalid_generated_xml(
            "envelope must open at least one element",
        ));
    }
    if !stack.is_empty() {
        return Err(invalid_generated_xml(
            "envelope suffix leaves open elements",
        ));
    }
    Ok(insertion_depth)
}

fn parse_prelude_envelope_shape(bytes: &[u8], boundary: usize) -> Result<usize> {
    if boundary > bytes.len() {
        return Err(invalid_generated_xml(
            "envelope prefix boundary exceeds the combined shell",
        ));
    }
    let text = std::str::from_utf8(bytes)
        .map_err(|error| invalid_generated_xml(format!("envelope is not UTF-8: {error}")))?;
    let mut reader = Reader::from_str(text);
    reader.config_mut().trim_text(false);
    reader.config_mut().check_end_names = false;
    let mut stack = Vec::new();
    let mut insertion_depth = None;
    let mut saw_root = false;
    let mut saw_document_event = false;
    let mut trailing_starts = 0usize;

    loop {
        let start_offset = usize::try_from(reader.buffer_position()).unwrap_or(usize::MAX);
        let event = reader
            .read_event()
            .map_err(|error| invalid_generated_xml(format!("invalid envelope: {error}")))?;
        let end_offset = usize::try_from(reader.buffer_position()).unwrap_or(usize::MAX);
        if start_offset < boundary && end_offset > boundary {
            return Err(invalid_generated_xml(
                "envelope boundary splits an XML event",
            ));
        }
        let in_prefix = start_offset < boundary && end_offset <= boundary;
        if !in_prefix && insertion_depth.is_none() {
            insertion_depth = Some(stack.len());
        }

        match event {
            Event::Decl(_)
                if in_prefix
                    && start_offset == 0
                    && !saw_document_event
                    && bytes.starts_with(b"<?xml") =>
            {
                saw_document_event = true;
            },
            Event::Start(start) if in_prefix => {
                if stack.is_empty() {
                    if saw_root {
                        return Err(invalid_generated_xml(
                            "envelope prefix contains a second document root",
                        ));
                    }
                    saw_root = true;
                }
                reject_xml_space(&start)?;
                let name = owned_name(start.name().as_ref())?;
                stack.try_reserve(1).map_err(|source| Error::Allocation {
                    resource: "generated XML envelope depth",
                    source,
                })?;
                stack.push(name);
                trailing_starts = trailing_starts.checked_add(1).ok_or_else(|| {
                    invalid_generated_xml("generated XML envelope path depth overflow")
                })?;
                saw_document_event = true;
            },
            Event::Empty(empty) if in_prefix => {
                if stack.is_empty() {
                    return Err(invalid_generated_xml(
                        "envelope fixed element appears outside its document root",
                    ));
                }
                reject_xml_space(&empty)?;
                trailing_starts = 0;
                saw_document_event = true;
            },
            Event::End(end) if in_prefix => {
                if stack.len() <= 1 {
                    return Err(invalid_generated_xml(
                        "envelope prefix closes its document root",
                    ));
                }
                let expected = stack
                    .pop()
                    .ok_or_else(|| invalid_generated_xml("envelope has an extra end tag"))?;
                if expected.as_slice() != end.name().as_ref() {
                    return Err(invalid_generated_xml(
                        "envelope end tag does not match its start tag",
                    ));
                }
                trailing_starts = 0;
                saw_document_event = true;
            },
            Event::End(end) if !in_prefix => {
                let expected = stack
                    .pop()
                    .ok_or_else(|| invalid_generated_xml("envelope has an extra end tag"))?;
                if expected.as_slice() != end.name().as_ref() {
                    return Err(invalid_generated_xml(
                        "envelope end tag does not match its start tag",
                    ));
                }
            },
            Event::Eof => {
                if start_offset != bytes.len() || end_offset != bytes.len() {
                    return Err(invalid_generated_xml("envelope EOF position is invalid"));
                }
                break;
            },
            _ => {
                return Err(invalid_generated_xml(
                    "envelope prelude must contain only declaration, elements, and matching end tags",
                ));
            },
        }
    }

    let insertion_depth =
        insertion_depth.ok_or_else(|| invalid_generated_xml("envelope has no prefix boundary"))?;
    if !saw_root || !saw_document_event {
        return Err(invalid_generated_xml(
            "envelope must contain one document root",
        ));
    }
    if insertion_depth == 0 {
        return Err(invalid_generated_xml(
            "envelope must leave an open insertion path",
        ));
    }
    if trailing_starts == 0 {
        return Err(invalid_generated_xml(
            "envelope must end with open insertion-path start tags",
        ));
    }
    if !stack.is_empty() {
        return Err(invalid_generated_xml(
            "envelope suffix leaves open elements",
        ));
    }
    Ok(insertion_depth)
}

fn reject_xml_space(start: &quick_xml::events::BytesStart<'_>) -> Result<()> {
    for attribute in start.attributes() {
        let attribute = attribute.map_err(|error| {
            invalid_generated_xml(format!("invalid envelope attribute: {error}"))
        })?;
        if attribute.key.as_ref() == b"xml:space" {
            return Err(invalid_generated_xml(
                "envelope must not carry inherited xml:space state",
            ));
        }
    }
    Ok(())
}

fn validate_fragment_shape(bytes: &[u8]) -> Result<()> {
    let text = std::str::from_utf8(bytes)
        .map_err(|error| invalid_generated_xml(format!("fragment is not UTF-8: {error}")))?;
    let mut reader = Reader::from_str(text);
    reader.config_mut().trim_text(false);
    let mut depth = 0usize;
    let mut roots = 0usize;
    let mut closed_root = false;
    loop {
        let event = reader
            .read_event()
            .map_err(|error| invalid_generated_xml(format!("invalid fragment XML: {error}")))?;
        match event {
            Event::Start(_) => {
                if depth == 0 {
                    if closed_root || roots != 0 {
                        return Err(invalid_generated_xml(
                            "fragment must contain exactly one element root",
                        ));
                    }
                    roots = 1;
                }
                depth = depth
                    .checked_add(1)
                    .ok_or_else(|| invalid_generated_xml("fragment depth overflow"))?;
            },
            Event::Empty(_) => {
                if depth == 0 {
                    if closed_root || roots != 0 {
                        return Err(invalid_generated_xml(
                            "fragment must contain exactly one element root",
                        ));
                    }
                    roots = 1;
                    closed_root = true;
                }
            },
            Event::End(_) => {
                if depth == 0 {
                    return Err(invalid_generated_xml("fragment has an unexpected end tag"));
                }
                depth -= 1;
                if depth == 0 {
                    closed_root = true;
                }
            },
            Event::Text(_) | Event::CData(_) => {
                if depth == 0 {
                    return Err(invalid_generated_xml(
                        "fragment has character data outside its element root",
                    ));
                }
            },
            Event::GeneralRef(reference) => {
                if depth == 0 {
                    return Err(invalid_generated_xml(
                        "fragment has a reference outside its element root",
                    ));
                }
                if !is_predefined_or_numeric_reference(reference.as_ref()) {
                    return Err(invalid_generated_xml(
                        "fragment contains an undeclared named reference",
                    ));
                }
            },
            Event::Decl(_) | Event::DocType(_) | Event::Comment(_) | Event::PI(_) => {
                return Err(invalid_generated_xml(
                    "fragment contains a declaration, doctype, comment, or processing instruction",
                ));
            },
            Event::Eof => break,
        }
    }
    if roots != 1 || depth != 0 || !closed_root {
        return Err(invalid_generated_xml(
            "fragment must contain one complete element root",
        ));
    }
    Ok(())
}

fn is_predefined_or_numeric_reference(reference: &[u8]) -> bool {
    if matches!(reference, b"lt" | b"gt" | b"amp" | b"apos" | b"quot") {
        return true;
    }
    if let Some(hex) = reference.strip_prefix(b"#x") {
        return !hex.is_empty()
            && hex.iter().all(u8::is_ascii_hexdigit)
            && std::str::from_utf8(hex)
                .ok()
                .and_then(|value| u32::from_str_radix(value, 16).ok())
                .is_some_and(is_valid_xml_character);
    }
    if let Some(decimal) = reference.strip_prefix(b"#") {
        return !decimal.is_empty()
            && decimal.iter().all(u8::is_ascii_digit)
            && std::str::from_utf8(decimal)
                .ok()
                .and_then(|value| value.parse::<u32>().ok())
                .is_some_and(is_valid_xml_character);
    }
    false
}

fn immutable_hard_limits() -> Result<GeneratedXmlLimits> {
    xml_minifier::audit::Limits::new(
        xml_minifier::audit::Limits::BYTE_CEILING,
        xml_minifier::audit::Limits::DEPTH_CEILING,
        xml_minifier::audit::Limits::EVENT_CEILING,
        xml_minifier::audit::Limits::ATTRIBUTE_CEILING,
        xml_minifier::audit::Limits::TOKEN_BYTE_CEILING,
        xml_minifier::audit::Limits::TEXT_BYTE_CEILING,
    )
    .map_err(|error| invalid_generated_xml(format!("invalid immutable XML limits: {error}")))
}

const fn is_valid_xml_character(value: u32) -> bool {
    matches!(value, 0x9 | 0xA | 0xD | 0x20..=0xD7FF | 0xE000..=0xFFFD | 0x10000..=0x10FFFF)
}

fn owned_bytes(bytes: &[u8], resource: &'static str) -> Result<Vec<u8>> {
    let mut owned = Vec::new();
    owned
        .try_reserve_exact(bytes.len())
        .map_err(|source| Error::Allocation { resource, source })?;
    owned.extend_from_slice(bytes);
    Ok(owned)
}

fn concatenate(left: &[u8], right: &[u8], resource: &'static str) -> Result<Vec<u8>> {
    let length = left
        .len()
        .checked_add(right.len())
        .ok_or_else(|| invalid_generated_xml("generated XML shell byte count overflow"))?;
    let mut combined = Vec::new();
    combined
        .try_reserve_exact(length)
        .map_err(|source| Error::Allocation { resource, source })?;
    combined.extend_from_slice(left);
    combined.extend_from_slice(right);
    Ok(combined)
}

fn owned_name(bytes: &[u8]) -> Result<Vec<u8>> {
    owned_bytes(bytes, "generated XML element name")
}

fn validate_limits(limits: GeneratedXmlLimits, max_fragment_bytes: usize) -> Result<()> {
    let checks = [
        (
            "bytes",
            limits.max_bytes(),
            xml_minifier::audit::Limits::BYTE_CEILING,
        ),
        (
            "depth",
            limits.max_depth(),
            xml_minifier::audit::Limits::DEPTH_CEILING,
        ),
        (
            "events",
            limits.max_events(),
            xml_minifier::audit::Limits::EVENT_CEILING,
        ),
        (
            "attributes",
            limits.max_attributes(),
            xml_minifier::audit::Limits::ATTRIBUTE_CEILING,
        ),
        (
            "token bytes",
            limits.max_token_bytes(),
            xml_minifier::audit::Limits::TOKEN_BYTE_CEILING,
        ),
        (
            "text bytes",
            limits.max_text_bytes(),
            xml_minifier::audit::Limits::TEXT_BYTE_CEILING,
        ),
    ];
    for (resource, requested, ceiling) in checks {
        if requested > ceiling {
            return Err(invalid_generated_xml(format!(
                "XML {resource} limit exceeds its hard ceiling"
            )));
        }
    }
    if max_fragment_bytes == 0 {
        return Err(invalid_generated_xml(
            "generated XML fragment capacity must be nonzero",
        ));
    }
    if max_fragment_bytes > limits.max_bytes() {
        return Err(invalid_generated_xml(
            "generated XML fragment capacity exceeds the document byte limit",
        ));
    }
    Ok(())
}

fn check_report_limits(report: GeneratedXmlReport, limits: GeneratedXmlLimits) -> Result<()> {
    let checks = [
        (
            GeneratedXmlLimitResource::Bytes,
            report.bytes,
            limits.max_bytes(),
        ),
        (
            GeneratedXmlLimitResource::Depth,
            report.max_depth,
            limits.max_depth(),
        ),
        (
            GeneratedXmlLimitResource::Events,
            report.events,
            limits.max_events(),
        ),
        (
            GeneratedXmlLimitResource::Attributes,
            report.attributes,
            limits.max_attributes(),
        ),
        (
            GeneratedXmlLimitResource::TextBytes,
            report.text_bytes,
            limits.max_text_bytes(),
        ),
    ];
    for (resource, actual, maximum) in checks {
        if actual > maximum {
            return Err(xml_limit(resource, actual, maximum));
        }
    }
    Ok(())
}

fn audit_error(error: xml_minifier::audit::Error) -> Error {
    match error {
        xml_minifier::audit::Error::Limit {
            resource,
            limit,
            actual,
            ..
        } => xml_limit(resource, actual, limit),
        other => invalid_generated_xml(format!("XML audit failed: {other}")),
    }
}

fn invalid_generated_xml(message: impl Into<String>) -> Error {
    Error::InvalidFormat(format!("generated XML: {}", message.into()))
}
