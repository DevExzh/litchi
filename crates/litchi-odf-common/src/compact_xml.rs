//! Bounded compactness validation for freshly authored whole XML parts.

use quick_xml::{Reader, events::Event};
use std::fmt;

/// Default maximum size of one freshly authored XML part.
pub const DEFAULT_MAX_BYTES: usize = 64 * 1024 * 1024;
/// Default maximum element nesting depth.
pub const DEFAULT_MAX_DEPTH: usize = 256;
/// Absolute maximum accepted XML part size.
pub const HARD_MAX_BYTES: usize = 64 * 1024 * 1024;
/// Absolute maximum accepted element nesting depth.
pub const HARD_MAX_DEPTH: usize = 4_096;

/// Resource limits for compact XML validation.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub struct Limits {
    /// Maximum source length in bytes.
    max_bytes: usize,
    /// Maximum number of simultaneously open elements.
    max_depth: usize,
}

impl Default for Limits {
    fn default() -> Self {
        Self {
            max_bytes: DEFAULT_MAX_BYTES,
            max_depth: DEFAULT_MAX_DEPTH,
        }
    }
}

impl Limits {
    /// Construct caller-selected limits within the immutable safety ceilings.
    ///
    /// # Errors
    ///
    /// Returns a typed limit error when either value is zero or exceeds its
    /// public hard ceiling.
    pub fn new(max_bytes: usize, max_depth: usize) -> Result<Self, Error> {
        if max_bytes == 0 || max_bytes > HARD_MAX_BYTES {
            return Err(Error::new(ErrorKind::InputTooLarge, max_bytes));
        }
        if max_depth == 0 || max_depth > HARD_MAX_DEPTH {
            return Err(Error::new(ErrorKind::DepthLimit, max_depth));
        }
        Ok(Self {
            max_bytes,
            max_depth,
        })
    }

    /// Return the configured byte ceiling.
    #[must_use]
    pub const fn max_bytes(self) -> usize {
        self.max_bytes
    }

    /// Return the configured depth ceiling.
    #[must_use]
    pub const fn max_depth(self) -> usize {
        self.max_depth
    }

    /// Return a copy with a checked byte ceiling.
    ///
    /// # Errors
    ///
    /// Returns an error when `max_bytes` is zero or exceeds the hard ceiling.
    pub fn with_max_bytes(self, max_bytes: usize) -> Result<Self, Error> {
        Self::new(max_bytes, self.max_depth)
    }

    /// Return a copy with a checked depth ceiling.
    ///
    /// # Errors
    ///
    /// Returns an error when `max_depth` is zero or exceeds the hard ceiling.
    pub fn with_max_depth(self, max_depth: usize) -> Result<Self, Error> {
        Self::new(self.max_bytes, max_depth)
    }
}

/// The reason freshly authored XML is not a bounded compact whole part.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
#[non_exhaustive]
pub enum ErrorKind {
    /// The part exceeds its byte budget.
    InputTooLarge,
    /// Element nesting exceeds its depth budget.
    DepthLimit,
    /// Memory for bounded parser state could not be reserved.
    AllocationFailed,
    /// Formatting whitespace occurs between elements or inside markup.
    FormattingWhitespace,
    /// An empty element uses the non-compact ` />` terminator.
    SpacedEmptyElement,
    /// A document type declaration is forbidden in authored ODF XML.
    DocumentType,
    /// The input is not well-formed XML.
    MalformedXml,
}

/// Compact XML validation failure with a byte-oriented source position.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct Error {
    kind: ErrorKind,
    offset: usize,
}

impl Error {
    const fn new(kind: ErrorKind, offset: usize) -> Self {
        Self { kind, offset }
    }

    /// Return the stable failure category.
    #[must_use]
    pub const fn kind(&self) -> ErrorKind {
        self.kind
    }

    /// Return the byte offset at or immediately after the rejected construct.
    #[must_use]
    pub const fn offset(&self) -> usize {
        self.offset
    }
}

impl fmt::Display for Error {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        write!(
            formatter,
            "compact XML validation failed at byte {}: {:?}",
            self.offset, self.kind
        )
    }
}

impl std::error::Error for Error {}

impl From<Error> for litchi_core::Error {
    fn from(error: Error) -> Self {
        use litchi_core::xml::CompactnessKind;

        let kind = match error.kind {
            ErrorKind::InputTooLarge => CompactnessKind::InputTooLarge,
            ErrorKind::DepthLimit => CompactnessKind::DepthLimit,
            ErrorKind::AllocationFailed => CompactnessKind::AllocationFailed,
            ErrorKind::FormattingWhitespace => CompactnessKind::FormattingWhitespace,
            ErrorKind::SpacedEmptyElement => CompactnessKind::SpacedEmptyElement,
            ErrorKind::DocumentType => CompactnessKind::DocumentType,
            ErrorKind::MalformedXml => CompactnessKind::MalformedXml,
        };
        let offset = u64::try_from(error.offset).unwrap_or(u64::MAX);
        Self::XmlCompactness { kind, offset }
    }
}

/// Validate freshly authored XML using the default resource limits.
///
/// # Errors
///
/// Returns a typed [`Error`] for malformed, non-compact, or over-limit input.
pub fn validate(xml: &[u8]) -> Result<(), Error> {
    validate_with_limits(xml, Limits::default())
}

/// Validate freshly authored XML without normalizing any source bytes.
///
/// This rejects indentation text containing line breaks or tabs, line breaks
/// or tabs inside markup, spaced empty-element terminators, DTDs, malformed
/// XML, and configured resource-limit violations. Semantic character data,
/// CDATA, comments, processing instructions, and `xml:space="preserve"`
/// regions are inspected but never rewritten.
///
/// # Errors
///
/// Returns a typed [`Error`] for malformed, non-compact, or over-limit input.
pub fn validate_with_limits(xml: &[u8], requested_limits: Limits) -> Result<(), Error> {
    let limits = Limits::new(requested_limits.max_bytes, requested_limits.max_depth)?;
    if xml.len() > limits.max_bytes {
        return Err(Error::new(ErrorKind::InputTooLarge, xml.len()));
    }

    let mut reader = Reader::from_reader(xml);
    reader.config_mut().check_end_names = true;
    let mut preserve_stack = Vec::new();
    let mut root_seen = false;

    loop {
        let start = source_offset(reader.buffer_position());
        let event = reader
            .read_event()
            .map_err(|_read_error| Error::new(ErrorKind::MalformedXml, start))?;
        let end = source_offset(reader.buffer_position());

        match event {
            Event::Start(element) => {
                validate_markup(&xml[start..end], end)?;
                if preserve_stack.is_empty() {
                    if root_seen {
                        return Err(Error::new(ErrorKind::MalformedXml, start));
                    }
                    root_seen = true;
                }
                if preserve_stack.len() >= limits.max_depth {
                    return Err(Error::new(ErrorKind::DepthLimit, start));
                }
                if preserve_stack.len() == preserve_stack.capacity() {
                    preserve_stack.try_reserve(1).map_err(|_allocation_error| {
                        Error::new(ErrorKind::AllocationFailed, start)
                    })?;
                }
                let inherited = preserve_stack.last().copied().unwrap_or(false);
                preserve_stack.push(xml_space(&element, start)?.unwrap_or(inherited));
            },
            Event::Empty(element) => {
                validate_markup(&xml[start..end], end)?;
                if preserve_stack.is_empty() {
                    if root_seen {
                        return Err(Error::new(ErrorKind::MalformedXml, start));
                    }
                    root_seen = true;
                }
                if preserve_stack.len() >= limits.max_depth {
                    return Err(Error::new(ErrorKind::DepthLimit, start));
                }
                let _ = xml_space(&element, start)?;
            },
            Event::End(_) => {
                validate_markup(&xml[start..end], end)?;
                let _ = preserve_stack.pop();
            },
            Event::Text(text) => {
                let bytes: &[u8] = text.as_ref();
                if preserve_stack.is_empty() && !bytes.is_empty() {
                    return Err(Error::new(ErrorKind::FormattingWhitespace, start));
                }
                let preserve = preserve_stack.last().copied().unwrap_or(false);
                if !preserve
                    && bytes.iter().all(u8::is_ascii_whitespace)
                    && bytes
                        .iter()
                        .any(|byte| matches!(byte, b'\n' | b'\r' | b'\t'))
                {
                    return Err(Error::new(ErrorKind::FormattingWhitespace, start));
                }
            },
            Event::DocType(_) => return Err(Error::new(ErrorKind::DocumentType, start)),
            Event::Decl(_) => validate_declaration(&xml[start..end], end)?,
            Event::PI(_) => validate_processing_instruction(&xml[start..end], end)?,
            Event::Eof if preserve_stack.is_empty() && root_seen => return Ok(()),
            Event::Eof => return Err(Error::new(ErrorKind::MalformedXml, start)),
            Event::CData(_) | Event::GeneralRef(_) if preserve_stack.is_empty() => {
                return Err(Error::new(ErrorKind::MalformedXml, start));
            },
            Event::GeneralRef(reference) if !is_predefined_reference(reference.as_ref()) => {
                return Err(Error::new(ErrorKind::MalformedXml, start));
            },
            Event::Comment(_) | Event::CData(_) | Event::GeneralRef(_) => {},
        }
    }
}

fn is_predefined_reference(reference: &[u8]) -> bool {
    reference == b"amp"
        || reference == b"lt"
        || reference == b"gt"
        || reference == b"apos"
        || reference == b"quot"
}

fn validate_declaration(bytes: &[u8], offset: usize) -> Result<(), Error> {
    validate_markup(bytes, offset)?;
    if contains_unquoted(bytes, b" =")
        || contains_unquoted(bytes, b"= ")
        || contains_unquoted(bytes, b" ?>")
    {
        return Err(Error::new(ErrorKind::FormattingWhitespace, offset));
    }
    Ok(())
}

fn validate_processing_instruction(bytes: &[u8], offset: usize) -> Result<(), Error> {
    let body = bytes
        .strip_prefix(b"<?")
        .and_then(|body| body.strip_suffix(b"?>"))
        .ok_or_else(|| Error::new(ErrorKind::MalformedXml, offset))?;
    let separator = body
        .iter()
        .position(u8::is_ascii_whitespace)
        .unwrap_or(body.len());
    if separator == body.len() {
        return Ok(());
    }
    if body[separator] != b' ' || separator + 1 == body.len() {
        return Err(Error::new(ErrorKind::FormattingWhitespace, offset));
    }
    Ok(())
}

fn validate_markup(bytes: &[u8], offset: usize) -> Result<(), Error> {
    if bytes
        .iter()
        .any(|byte| matches!(byte, b'\n' | b'\r' | b'\t'))
    {
        return Err(Error::new(ErrorKind::FormattingWhitespace, offset));
    }
    let mut quote = None;
    let mut previous_space = false;
    for (index, byte) in bytes.iter().enumerate() {
        match (quote, byte) {
            (Some(delimiter), current) if current == &delimiter => quote = None,
            (None, b'\'' | b'"') => quote = Some(*byte),
            (None, b'/') if previous_space && bytes.get(index + 1).copied() == Some(b'>') => {
                return Err(Error::new(ErrorKind::SpacedEmptyElement, offset));
            },
            (None, b' ' | b'>') if previous_space => {
                return Err(Error::new(ErrorKind::FormattingWhitespace, offset));
            },
            _ => {},
        }
        previous_space = quote.is_none() && *byte == b' ';
    }
    Ok(())
}

fn contains_unquoted(bytes: &[u8], needle: &[u8]) -> bool {
    let mut quote = None;
    for (index, byte) in bytes.iter().enumerate() {
        match (quote, byte) {
            (Some(delimiter), current) if current == &delimiter => quote = None,
            (None, b'\'' | b'"') => quote = Some(*byte),
            (None, _) if bytes[index..].starts_with(needle) => return true,
            _ => {},
        }
    }
    false
}

fn xml_space(
    element: &quick_xml::events::BytesStart<'_>,
    offset: usize,
) -> Result<Option<bool>, Error> {
    let mut preserve = None;
    for raw_attribute in element.attributes() {
        let attribute = raw_attribute
            .map_err(|_attribute_error| Error::new(ErrorKind::MalformedXml, offset))?;
        if attribute.key.as_ref() == b"xml:space" {
            preserve = match attribute.value.as_ref() {
                b"preserve" => Some(true),
                b"default" => Some(false),
                _ => preserve,
            };
        }
    }
    Ok(preserve)
}

fn source_offset(position: u64) -> usize {
    usize::try_from(position).unwrap_or(usize::MAX)
}

#[cfg(test)]
mod tests {
    use super::{
        ErrorKind, HARD_MAX_BYTES, HARD_MAX_DEPTH, Limits, validate, validate_with_limits,
    };

    #[test]
    fn accepts_compact_xml_and_semantic_whitespace() -> Result<(), super::Error> {
        validate(b"<?xml version=\"1.0\"?><a value=\"quoted /> and > stay semantic\">word\nword<![CDATA[\n]]><b xml:space=\"preserve\">\n  </b><!--\n--><?p semantic  data\nkept?></a>")
    }

    #[test]
    fn accepts_only_predefined_general_references() -> Result<(), super::Error> {
        validate(b"<a>&amp;&lt;&gt;&apos;&quot;</a>")?;
        assert_eq!(
            validate(b"<a>&undeclared;</a>")
                .err()
                .map(|error| error.kind()),
            Some(ErrorKind::MalformedXml)
        );
        Ok(())
    }

    #[test]
    fn rejects_formatting_whitespace_and_spaced_empty_elements() {
        assert_eq!(
            validate(b"<a>\n<b/></a>").err().map(|error| error.kind()),
            Some(ErrorKind::FormattingWhitespace)
        );
        assert_eq!(
            validate(b"<a><b /></a>").err().map(|error| error.kind()),
            Some(ErrorKind::SpacedEmptyElement)
        );
        assert_eq!(
            validate(b"<a\nvalue=\"x\"/>")
                .err()
                .map(|error| error.kind()),
            Some(ErrorKind::FormattingWhitespace)
        );
        assert_eq!(
            validate(b"<a  value=\"x\"/>")
                .err()
                .map(|error| error.kind()),
            Some(ErrorKind::FormattingWhitespace)
        );
        assert_eq!(
            validate(b"<?xml  version=\"1.0\"?><a/>")
                .err()
                .map(|error| error.kind()),
            Some(ErrorKind::FormattingWhitespace)
        );
        assert_eq!(
            validate(b"<?xml version =\"1.0\"?><a/>")
                .err()
                .map(|error| error.kind()),
            Some(ErrorKind::FormattingWhitespace)
        );
        assert_eq!(
            validate(b"<a><?p\tdata?></a>")
                .err()
                .map(|error| error.kind()),
            Some(ErrorKind::FormattingWhitespace)
        );
        assert_eq!(
            validate(b"<a><?p ?></a>").err().map(|error| error.kind()),
            Some(ErrorKind::FormattingWhitespace)
        );
    }

    #[test]
    fn reports_security_and_resource_failures() -> Result<(), super::Error> {
        assert_eq!(
            validate(b"<!DOCTYPE a><a/>")
                .err()
                .map(|error| error.kind()),
            Some(ErrorKind::DocumentType)
        );
        assert_eq!(
            validate(b"<a>").err().map(|error| error.kind()),
            Some(ErrorKind::MalformedXml)
        );
        assert_eq!(
            validate(b"").err().map(|error| error.kind()),
            Some(ErrorKind::MalformedXml)
        );
        assert_eq!(
            validate(b"<a/><b/>").err().map(|error| error.kind()),
            Some(ErrorKind::MalformedXml)
        );
        assert_eq!(
            validate_with_limits(b"<a/>", Limits::new(3, 1)?)
                .err()
                .map(|error| error.kind()),
            Some(ErrorKind::InputTooLarge)
        );
        assert_eq!(
            validate_with_limits(b"<a><b/></a>", Limits::new(32, 1)?)
                .err()
                .map(|error| error.kind()),
            Some(ErrorKind::DepthLimit)
        );
        assert_eq!(
            validate(b"<![CDATA[outside]]><a/>")
                .err()
                .map(|error| error.kind()),
            Some(ErrorKind::MalformedXml)
        );
        assert_eq!(
            validate(b"&outside;<a/>").err().map(|error| error.kind()),
            Some(ErrorKind::MalformedXml)
        );
        assert!(Limits::new(HARD_MAX_BYTES, HARD_MAX_DEPTH).is_ok());
        assert_eq!(
            Limits::new(HARD_MAX_BYTES + 1, HARD_MAX_DEPTH)
                .err()
                .map(|error| error.kind()),
            Some(ErrorKind::InputTooLarge)
        );
        assert_eq!(
            Limits::new(HARD_MAX_BYTES, HARD_MAX_DEPTH + 1)
                .err()
                .map(|error| error.kind()),
            Some(ErrorKind::DepthLimit)
        );
        Ok(())
    }
}
