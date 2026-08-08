//! Bounded verification of the repository's compact XML output contract.
//!
//! Character data and CDATA are never normalized. Plain spaces remain content;
//! only whitespace-only nodes containing CR, LF, or tab are classified as
//! structural formatting outside an inherited `xml:space="preserve"` scope.

#![allow(
    clippy::arbitrary_source_item_ordering,
    reason = "semantic API types precede their streaming implementation and package submodule"
)]

use core::fmt;
use quick_xml::Reader;
use quick_xml::events::{BytesStart, Event};

/// Finite resource budgets for one XML document.
#[derive(Clone, Copy, Debug, Eq, PartialEq)]
pub struct Limits {
    attributes: usize,
    bytes: usize,
    depth: usize,
    events: usize,
    text_bytes: usize,
    token_bytes: usize,
}

impl Limits {
    /// Hard ceiling for aggregate attributes.
    pub const ATTRIBUTE_CEILING: usize = 1_000_000;
    /// Hard ceiling for one XML document in bytes.
    pub const BYTE_CEILING: usize = 256 * 1024 * 1024;
    /// Hard ceiling for element nesting.
    pub const DEPTH_CEILING: usize = 4_096;
    /// Hard ceiling for parser events.
    pub const EVENT_CEILING: usize = 4_000_000;
    /// Hard ceiling for aggregate character-data bytes.
    pub const TEXT_BYTE_CEILING: usize = 256 * 1024 * 1024;
    /// Hard ceiling for one lexical token in bytes.
    pub const TOKEN_BYTE_CEILING: usize = 64 * 1024 * 1024;

    /// Creates an explicit limit profile.
    ///
    /// # Errors
    ///
    /// Returns [`ConfigError`] when any requested value exceeds its immutable
    /// hard ceiling.
    pub fn new(
        max_bytes: usize,
        max_depth: usize,
        max_events: usize,
        max_attributes: usize,
        max_token_bytes: usize,
        max_text_bytes: usize,
    ) -> Result<Self, ConfigError> {
        let limits = Self {
            attributes: max_attributes,
            bytes: max_bytes,
            depth: max_depth,
            events: max_events,
            text_bytes: max_text_bytes,
            token_bytes: max_token_bytes,
        };
        limits.check()?;
        Ok(limits)
    }

    /// Starts a safe fallible limit builder from the default profile.
    #[must_use]
    pub fn builder() -> Builder {
        Builder::default()
    }

    /// Returns the immutable hard ceiling for `resource`.
    #[must_use]
    pub const fn ceiling(resource: Resource) -> usize {
        match resource {
            Resource::Attributes => Self::ATTRIBUTE_CEILING,
            Resource::Bytes => Self::BYTE_CEILING,
            Resource::Depth => Self::DEPTH_CEILING,
            Resource::Events => Self::EVENT_CEILING,
            Resource::TextBytes => Self::TEXT_BYTE_CEILING,
            Resource::TokenBytes => Self::TOKEN_BYTE_CEILING,
        }
    }

    /// Narrows one resource without permitting an increase.
    #[must_use]
    pub const fn narrow(mut self, resource: Resource, maximum: usize) -> Self {
        match resource {
            Resource::Attributes => self.attributes = minimum(self.attributes, maximum),
            Resource::Bytes => self.bytes = minimum(self.bytes, maximum),
            Resource::Depth => self.depth = minimum(self.depth, maximum),
            Resource::Events => self.events = minimum(self.events, maximum),
            Resource::TextBytes => self.text_bytes = minimum(self.text_bytes, maximum),
            Resource::TokenBytes => self.token_bytes = minimum(self.token_bytes, maximum),
        }
        self
    }

    const fn bounded(
        bytes: usize,
        depth: usize,
        events: usize,
        attributes: usize,
        token_bytes: usize,
        text_bytes: usize,
    ) -> Self {
        Self {
            attributes,
            bytes,
            depth,
            events,
            text_bytes,
            token_bytes,
        }
    }

    fn check(self) -> Result<(), ConfigError> {
        for resource in [
            Resource::Bytes,
            Resource::Depth,
            Resource::Events,
            Resource::Attributes,
            Resource::TokenBytes,
            Resource::TextBytes,
        ] {
            let requested = self.value(resource);
            let ceiling = Self::ceiling(resource);
            if requested > ceiling {
                return Err(ConfigError {
                    ceiling,
                    requested,
                    resource,
                });
            }
        }
        Ok(())
    }

    const fn value(self, resource: Resource) -> usize {
        match resource {
            Resource::Attributes => self.attributes,
            Resource::Bytes => self.bytes,
            Resource::Depth => self.depth,
            Resource::Events => self.events,
            Resource::TextBytes => self.text_bytes,
            Resource::TokenBytes => self.token_bytes,
        }
    }

    /// Maximum number of attributes in the document.
    #[must_use]
    pub const fn max_attributes(self) -> usize {
        self.attributes
    }

    /// Maximum input size in bytes.
    #[must_use]
    pub const fn max_bytes(self) -> usize {
        self.bytes
    }

    /// Maximum element nesting depth.
    #[must_use]
    pub const fn max_depth(self) -> usize {
        self.depth
    }

    /// Maximum number of parser events.
    #[must_use]
    pub const fn max_events(self) -> usize {
        self.events
    }

    /// Maximum aggregate character-data bytes.
    #[must_use]
    pub const fn max_text_bytes(self) -> usize {
        self.text_bytes
    }

    /// Maximum bytes in one lexical token.
    #[must_use]
    pub const fn max_token_bytes(self) -> usize {
        self.token_bytes
    }
}

impl Default for Limits {
    fn default() -> Self {
        Self::bounded(
            32 * 1024 * 1024,
            256,
            1_000_000,
            250_000,
            4 * 1024 * 1024,
            16 * 1024 * 1024,
        )
    }
}

/// Fallible construction of one bounded [`Limits`] profile.
#[derive(Clone, Copy, Debug, Default, Eq, PartialEq)]
pub struct Builder {
    limits: Limits,
}

impl Builder {
    /// Sets the aggregate attribute limit.
    ///
    /// # Errors
    ///
    /// Returns [`ConfigError`] when `maximum` exceeds the immutable ceiling.
    pub fn attributes(self, maximum: usize) -> Result<Self, ConfigError> {
        self.setting(Resource::Attributes, maximum)
    }

    /// Builds the checked profile without allocation.
    #[must_use]
    pub const fn build(self) -> Limits {
        self.limits
    }

    /// Sets the input-byte limit.
    ///
    /// # Errors
    ///
    /// Returns [`ConfigError`] when `maximum` exceeds the immutable ceiling.
    pub fn bytes(self, maximum: usize) -> Result<Self, ConfigError> {
        self.setting(Resource::Bytes, maximum)
    }

    /// Sets the element-depth limit.
    ///
    /// # Errors
    ///
    /// Returns [`ConfigError`] when `maximum` exceeds the immutable ceiling.
    pub fn depth(self, maximum: usize) -> Result<Self, ConfigError> {
        self.setting(Resource::Depth, maximum)
    }

    /// Sets the parser-event limit.
    ///
    /// # Errors
    ///
    /// Returns [`ConfigError`] when `maximum` exceeds the immutable ceiling.
    pub fn events(self, maximum: usize) -> Result<Self, ConfigError> {
        self.setting(Resource::Events, maximum)
    }

    /// Sets one typed resource limit.
    ///
    /// # Errors
    ///
    /// Returns [`ConfigError`] when `maximum` exceeds the resource's immutable
    /// ceiling.
    pub fn limit(self, resource: Resource, maximum: usize) -> Result<Self, ConfigError> {
        self.setting(resource, maximum)
    }

    /// Sets the aggregate character-data limit.
    ///
    /// # Errors
    ///
    /// Returns [`ConfigError`] when `maximum` exceeds the immutable ceiling.
    pub fn text_bytes(self, maximum: usize) -> Result<Self, ConfigError> {
        self.setting(Resource::TextBytes, maximum)
    }

    /// Sets the single-token byte limit.
    ///
    /// # Errors
    ///
    /// Returns [`ConfigError`] when `maximum` exceeds the immutable ceiling.
    pub fn token_bytes(self, maximum: usize) -> Result<Self, ConfigError> {
        self.setting(Resource::TokenBytes, maximum)
    }

    fn setting(mut self, resource: Resource, maximum: usize) -> Result<Self, ConfigError> {
        let ceiling = Limits::ceiling(resource);
        if maximum > ceiling {
            return Err(ConfigError {
                ceiling,
                requested: maximum,
                resource,
            });
        }
        match resource {
            Resource::Attributes => self.limits.attributes = maximum,
            Resource::Bytes => self.limits.bytes = maximum,
            Resource::Depth => self.limits.depth = maximum,
            Resource::Events => self.limits.events = maximum,
            Resource::TextBytes => self.limits.text_bytes = maximum,
            Resource::TokenBytes => self.limits.token_bytes = maximum,
        }
        Ok(self)
    }
}

/// Invalid limit configuration above an immutable hard ceiling.
#[derive(Clone, Copy, Debug, Eq, PartialEq)]
pub struct ConfigError {
    ceiling: usize,
    requested: usize,
    resource: Resource,
}

impl ConfigError {
    /// Immutable ceiling that was exceeded.
    #[must_use]
    pub const fn ceiling(self) -> usize {
        self.ceiling
    }

    /// Requested value.
    #[must_use]
    pub const fn requested(self) -> usize {
        self.requested
    }

    /// Resource whose ceiling was exceeded.
    #[must_use]
    pub const fn resource(self) -> Resource {
        self.resource
    }
}

impl fmt::Display for ConfigError {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        write!(
            formatter,
            "XML {:?} limit {} exceeds hard ceiling {}",
            self.resource, self.requested, self.ceiling
        )
    }
}

impl std::error::Error for ConfigError {}

/// A resource governed by [`Limits`].
#[derive(Clone, Copy, Debug, Eq, PartialEq)]
#[non_exhaustive]
pub enum Resource {
    /// Total attribute count.
    Attributes,
    /// Input byte length.
    Bytes,
    /// Element nesting depth.
    Depth,
    /// Parser event count.
    Events,
    /// Aggregate character-data bytes.
    TextBytes,
    /// One lexical token's byte length.
    TokenBytes,
}

/// A lexically provable compactness defect.
#[derive(Clone, Copy, Debug, Eq, PartialEq)]
#[non_exhaustive]
pub enum Kind {
    /// Whitespace-only character data outside `xml:space="preserve"`.
    FormattingWhitespace,
    /// Attribute boundaries do not use exactly one ASCII space.
    AttributeSeparation,
    /// Whitespace occurs immediately before `>`, `/>`, or an end-tag close.
    WhitespaceBeforeClose,
}

/// Location and category of a compactness defect.
#[derive(Clone, Copy, Debug, Eq, PartialEq)]
pub struct Violation {
    kind: Kind,
    offset: usize,
}

impl Violation {
    /// The stable defect category.
    #[must_use]
    pub const fn kind(self) -> Kind {
        self.kind
    }

    /// Zero-based byte offset in the original XML.
    #[must_use]
    pub const fn offset(self) -> usize {
        self.offset
    }
}

/// Failure to parse or verify a compact XML document.
#[derive(Debug, Eq, PartialEq)]
#[non_exhaustive]
pub enum Error {
    /// A finite audit budget was exceeded.
    Limit {
        /// Governed resource.
        resource: Resource,
        /// Configured inclusive limit.
        limit: usize,
        /// First observed value beyond the limit.
        actual: usize,
        /// Byte offset at which accounting failed.
        offset: usize,
    },
    /// The input was not UTF-8 XML.
    Encoding {
        /// First invalid UTF-8 byte.
        valid_up_to: usize,
    },
    /// XML parsing or document-structure validation failed.
    Malformed {
        /// Parser byte offset.
        offset: usize,
        /// Bounded-by-input parser diagnostic.
        detail: String,
    },
    /// XML is valid but violates the compact output contract.
    NotCompact(Violation),
    /// DTD and DOCTYPE declarations are ineligible for compact package XML.
    Doctype {
        /// Zero-based byte offset of the declaration.
        offset: usize,
    },
    /// The depth stack could not reserve one bounded entry.
    Allocation,
}

impl Error {
    fn malformed(offset: usize, detail: impl Into<String>) -> Self {
        Self::Malformed {
            offset,
            detail: detail.into(),
        }
    }
}

impl fmt::Display for Error {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Self::Limit {
                resource,
                limit,
                actual,
                offset,
            } => write!(
                formatter,
                "XML {resource:?} limit {limit} exceeded by {actual} at byte {offset}"
            ),
            Self::Encoding { valid_up_to } => {
                write!(formatter, "XML is not UTF-8 at byte {valid_up_to}")
            },
            Self::Malformed { offset, detail } => {
                write!(formatter, "malformed XML at byte {offset}: {detail}")
            },
            Self::NotCompact(violation) => write!(
                formatter,
                "noncompact XML {:?} at byte {}",
                violation.kind, violation.offset
            ),
            Self::Doctype { offset } => {
                write!(
                    formatter,
                    "DTD and DOCTYPE are not allowed at byte {offset}"
                )
            },
            Self::Allocation => {
                formatter.write_str("could not reserve the bounded XML depth stack")
            },
        }
    }
}

impl std::error::Error for Error {}

/// Accounting summary for a verified compact XML document.
#[derive(Clone, Copy, Debug, Eq, PartialEq)]
#[must_use]
pub struct Report {
    attributes: usize,
    bytes: usize,
    events: usize,
    max_depth: usize,
    text_bytes: usize,
}

impl Report {
    /// Total attributes parsed.
    #[must_use]
    pub const fn attributes(self) -> usize {
        self.attributes
    }

    /// Input bytes parsed.
    #[must_use]
    pub const fn bytes(self) -> usize {
        self.bytes
    }

    /// Parser events, including EOF.
    #[must_use]
    pub const fn events(self) -> usize {
        self.events
    }

    /// Greatest element nesting depth.
    #[must_use]
    pub const fn max_depth(self) -> usize {
        self.max_depth
    }

    /// Aggregate character-data bytes.
    #[must_use]
    pub const fn text_bytes(self) -> usize {
        self.text_bytes
    }
}

#[derive(Clone, Copy, Debug, Eq, PartialEq)]
enum Space {
    Default,
    Preserve,
}

struct State {
    attributes: usize,
    depth: usize,
    events: usize,
    max_depth: usize,
    roots: usize,
    spaces: Vec<Space>,
    text_bytes: usize,
}

impl State {
    fn new() -> Self {
        Self {
            attributes: 0,
            depth: 0,
            events: 0,
            max_depth: 0,
            roots: 0,
            spaces: Vec::new(),
            text_bytes: 0,
        }
    }

    fn current_space(&self) -> Space {
        self.spaces.last().copied().unwrap_or(Space::Default)
    }
}

/// Parses `input` and rejects the first compactness or resource defect.
///
/// This function is an auditor, not a postprocessor: it never changes input
/// and therefore cannot silently rewrite opaque or mixed-content XML.
///
/// # Errors
///
/// Returns [`Error`] for invalid UTF-8, malformed XML, a finite resource-limit
/// breach, or the first compactness violation.
pub fn verify(input: &[u8], limits: Limits) -> Result<Report, Error> {
    check_limit(Resource::Bytes, limits.bytes, input.len(), 0)?;
    let xml = std::str::from_utf8(input).map_err(|error| Error::Encoding {
        valid_up_to: error.valid_up_to(),
    })?;
    let mut reader = Reader::from_str(xml);
    reader.config_mut().trim_text(false);
    let mut state = State::new();

    loop {
        let start = position(&reader);
        let event = reader
            .read_event()
            .map_err(|error| Error::malformed(position(&reader), error.to_string()))?;
        let end = position(&reader);
        let raw = input
            .get(start..end)
            .ok_or_else(|| Error::malformed(start, "parser position escaped input"))?;

        state.events = checked_add(state.events, 1, Resource::Events, limits.events, start)?;
        check_limit(Resource::TokenBytes, limits.token_bytes, raw.len(), start)?;

        match event {
            Event::Start(tag) => {
                check_start(raw, false, start)?;
                let space =
                    inspect_attributes(&tag, state.current_space(), &mut state, limits, start)?;
                enter_element(&mut state, limits, start)?;
                state
                    .spaces
                    .try_reserve(1)
                    .map_err(|_allocation| Error::Allocation)?;
                state.spaces.push(space);
            },
            Event::Empty(tag) => {
                check_start(raw, true, start)?;
                inspect_attributes(&tag, state.current_space(), &mut state, limits, start)?;
                enter_empty(&mut state, limits, start)?;
            },
            Event::End(_) => {
                check_end(raw, start)?;
                if state.depth == 0 || state.spaces.pop().is_none() {
                    return Err(Error::malformed(start, "unexpected end element"));
                }
                state.depth -= 1;
            },
            Event::Text(text) => {
                let bytes = text.as_ref();
                check_character_context(state.depth, bytes, start)?;
                charge_text(&mut state, limits, bytes.len(), start)?;
                if is_structural_whitespace(bytes) && state.current_space() != Space::Preserve {
                    return Err(Error::NotCompact(Violation {
                        kind: Kind::FormattingWhitespace,
                        offset: start,
                    }));
                }
            },
            Event::CData(data) => {
                check_character_context(state.depth, data.as_ref(), start)?;
                charge_text(&mut state, limits, data.as_ref().len(), start)?;
            },
            Event::GeneralRef(reference) => {
                check_character_context(state.depth, reference.as_ref(), start)?;
                charge_text(&mut state, limits, raw.len(), start)?;
            },
            Event::Decl(_) => check_declaration(raw, start)?,
            Event::Comment(_) | Event::PI(_) => {},
            Event::DocType(_) => return Err(Error::Doctype { offset: start }),
            Event::Eof => break,
        }
    }

    if state.depth != 0 {
        return Err(Error::malformed(input.len(), "unclosed document element"));
    }
    if state.roots != 1 {
        return Err(Error::malformed(
            input.len(),
            "XML must contain exactly one document element",
        ));
    }

    Ok(Report {
        attributes: state.attributes,
        bytes: input.len(),
        events: state.events,
        max_depth: state.max_depth,
        text_bytes: state.text_bytes,
    })
}

fn charge_text(
    state: &mut State,
    limits: Limits,
    amount: usize,
    offset: usize,
) -> Result<(), Error> {
    state.text_bytes = checked_add(
        state.text_bytes,
        amount,
        Resource::TextBytes,
        limits.text_bytes,
        offset,
    )?;
    Ok(())
}

fn check_character_context(depth: usize, bytes: &[u8], offset: usize) -> Result<(), Error> {
    if depth == 0 && !is_xml_whitespace(bytes) {
        return Err(Error::malformed(
            offset,
            "character data outside the document element",
        ));
    }
    Ok(())
}

fn enter_element(state: &mut State, limits: Limits, offset: usize) -> Result<(), Error> {
    if state.depth == 0 {
        if state.roots != 0 {
            return Err(Error::malformed(offset, "multiple document elements"));
        }
        state.roots = 1;
    }
    state.depth = checked_add(state.depth, 1, Resource::Depth, limits.depth, offset)?;
    state.max_depth = state.max_depth.max(state.depth);
    Ok(())
}

fn enter_empty(state: &mut State, limits: Limits, offset: usize) -> Result<(), Error> {
    if state.depth == 0 {
        if state.roots != 0 {
            return Err(Error::malformed(offset, "multiple document elements"));
        }
        state.roots = 1;
    }
    let depth = checked_add(state.depth, 1, Resource::Depth, limits.depth, offset)?;
    state.max_depth = state.max_depth.max(depth);
    Ok(())
}

fn inspect_attributes(
    tag: &BytesStart<'_>,
    inherited: Space,
    state: &mut State,
    limits: Limits,
    offset: usize,
) -> Result<Space, Error> {
    let mut space = inherited;
    for attribute_result in tag.attributes() {
        let attribute =
            attribute_result.map_err(|error| Error::malformed(offset, error.to_string()))?;
        state.attributes = checked_add(
            state.attributes,
            1,
            Resource::Attributes,
            limits.attributes,
            offset,
        )?;
        if attribute.key.as_ref() == b"xml:space" {
            space = match attribute.value.as_ref() {
                b"default" => Space::Default,
                b"preserve" => Space::Preserve,
                _ => {
                    return Err(Error::malformed(
                        offset,
                        "xml:space must be 'default' or 'preserve'",
                    ));
                },
            };
        }
    }
    Ok(space)
}

fn check_declaration(raw: &[u8], offset: usize) -> Result<(), Error> {
    let Some(inner) = raw
        .strip_prefix(b"<?")
        .and_then(|value| value.strip_suffix(b"?>"))
    else {
        return Err(Error::malformed(offset, "invalid XML declaration boundary"));
    };
    check_attribute_layout(inner, offset + 2)
}

fn check_end(raw: &[u8], offset: usize) -> Result<(), Error> {
    let Some(inner) = raw
        .strip_prefix(b"</")
        .and_then(|value| value.strip_suffix(b">"))
    else {
        return Err(Error::malformed(offset, "invalid end-tag boundary"));
    };
    if let Some(index) = inner.iter().position(|byte| is_space(*byte)) {
        let kind = if inner[index..].iter().all(|byte| is_space(*byte)) {
            Kind::WhitespaceBeforeClose
        } else {
            Kind::AttributeSeparation
        };
        return Err(Error::NotCompact(Violation {
            kind,
            offset: offset + 2 + index,
        }));
    }
    Ok(())
}

fn check_start(raw: &[u8], empty: bool, offset: usize) -> Result<(), Error> {
    let Some(without_open) = raw.strip_prefix(b"<") else {
        return Err(Error::malformed(offset, "invalid start-tag boundary"));
    };
    let inner = if empty {
        without_open.strip_suffix(b"/>")
    } else {
        without_open.strip_suffix(b">")
    }
    .ok_or_else(|| Error::malformed(offset, "invalid start-tag close"))?;
    check_attribute_layout(inner, offset + 1)
}

fn check_attribute_layout(inner: &[u8], offset: usize) -> Result<(), Error> {
    let Some(mut cursor) = inner.iter().position(|byte| is_space(*byte)) else {
        return Ok(());
    };

    loop {
        let separator = cursor;
        while cursor < inner.len() && is_space(inner[cursor]) {
            cursor += 1;
        }
        if cursor == inner.len() {
            return Err(Error::NotCompact(Violation {
                kind: Kind::WhitespaceBeforeClose,
                offset: offset + separator,
            }));
        }
        if cursor != separator + 1 || inner[separator] != b' ' {
            return Err(Error::NotCompact(Violation {
                kind: Kind::AttributeSeparation,
                offset: offset + separator,
            }));
        }

        let name_start = cursor;
        while cursor < inner.len() && !is_space(inner[cursor]) && inner[cursor] != b'=' {
            cursor += 1;
        }
        if cursor == name_start {
            return Err(Error::malformed(offset + cursor, "missing attribute name"));
        }
        if cursor == inner.len() || is_space(inner[cursor]) {
            return Err(Error::NotCompact(Violation {
                kind: Kind::AttributeSeparation,
                offset: offset + cursor,
            }));
        }
        cursor += 1;
        if cursor == inner.len() || is_space(inner[cursor]) {
            return Err(Error::NotCompact(Violation {
                kind: Kind::AttributeSeparation,
                offset: offset + cursor,
            }));
        }

        let quote = inner[cursor];
        if quote != b'\'' && quote != b'"' {
            return Err(Error::malformed(
                offset + cursor,
                "attribute value must be quoted",
            ));
        }
        cursor += 1;
        while cursor < inner.len() && inner[cursor] != quote {
            cursor += 1;
        }
        if cursor == inner.len() {
            return Err(Error::malformed(
                offset + cursor,
                "unterminated attribute value",
            ));
        }
        cursor += 1;
        if cursor == inner.len() {
            return Ok(());
        }
        if !is_space(inner[cursor]) {
            return Err(Error::malformed(
                offset + cursor,
                "missing attribute separator",
            ));
        }
    }
}

fn check_limit(
    resource: Resource,
    limit: usize,
    actual: usize,
    offset: usize,
) -> Result<(), Error> {
    if actual > limit {
        return Err(Error::Limit {
            resource,
            limit,
            actual,
            offset,
        });
    }
    Ok(())
}

fn checked_add(
    current: usize,
    amount: usize,
    resource: Resource,
    limit: usize,
    offset: usize,
) -> Result<usize, Error> {
    let actual = current.saturating_add(amount);
    check_limit(resource, limit, actual, offset)?;
    Ok(actual)
}

fn is_xml_whitespace(bytes: &[u8]) -> bool {
    !bytes.is_empty() && bytes.iter().all(|byte| is_space(*byte))
}

fn is_structural_whitespace(bytes: &[u8]) -> bool {
    is_xml_whitespace(bytes)
        && bytes
            .iter()
            .any(|byte| matches!(byte, b'\t' | b'\n' | b'\r'))
}

const fn is_space(byte: u8) -> bool {
    matches!(byte, b' ' | b'\t' | b'\n' | b'\r')
}

const fn minimum(left: usize, right: usize) -> usize {
    if left < right { left } else { right }
}

fn position(reader: &Reader<&[u8]>) -> usize {
    match usize::try_from(reader.buffer_position()) {
        Ok(value) => value,
        Err(_) => usize::MAX,
    }
}

/// Bounded verification over named XML package parts.
pub mod package {
    use super::{Error as DocumentError, Limits as DocumentLimits, verify as verify_document};
    use core::fmt;

    /// A borrowed named XML package member.
    #[derive(Clone, Copy, Debug)]
    pub struct Part<'a> {
        bytes: &'a [u8],
        name: &'a str,
    }

    impl<'a> Part<'a> {
        /// Creates a borrowed part without copying its name or payload.
        #[must_use]
        pub const fn new(name: &'a str, bytes: &'a [u8]) -> Self {
            Self { bytes, name }
        }

        /// Part payload.
        #[must_use]
        pub const fn bytes(self) -> &'a [u8] {
            self.bytes
        }

        /// Archive-relative diagnostic name.
        #[must_use]
        pub const fn name(self) -> &'a str {
            self.name
        }
    }

    /// Finite aggregate and per-document package audit limits.
    #[derive(Clone, Copy, Debug, Eq, PartialEq)]
    pub struct Limits {
        document: DocumentLimits,
        bytes: usize,
        parts: usize,
    }

    impl Limits {
        /// Creates an explicit package profile.
        #[must_use]
        pub const fn new(document: DocumentLimits, max_parts: usize, max_bytes: usize) -> Self {
            Self {
                document,
                bytes: max_bytes,
                parts: max_parts,
            }
        }

        /// Per-document limits.
        #[must_use]
        pub const fn document(self) -> DocumentLimits {
            self.document
        }

        /// Maximum aggregate XML payload bytes.
        #[must_use]
        pub const fn max_bytes(self) -> usize {
            self.bytes
        }

        /// Maximum named XML parts.
        #[must_use]
        pub const fn max_parts(self) -> usize {
            self.parts
        }
    }

    impl Default for Limits {
        fn default() -> Self {
            Self::new(DocumentLimits::default(), 65_536, 256 * 1024 * 1024)
        }
    }

    /// Aggregate package audit accounting.
    #[derive(Clone, Copy, Debug, Eq, PartialEq)]
    #[must_use]
    pub struct Report {
        attributes: usize,
        bytes: usize,
        events: usize,
        max_depth: usize,
        parts: usize,
        text_bytes: usize,
    }

    impl Report {
        /// Aggregate attributes.
        #[must_use]
        pub const fn attributes(self) -> usize {
            self.attributes
        }

        /// Aggregate XML bytes.
        #[must_use]
        pub const fn bytes(self) -> usize {
            self.bytes
        }

        /// Aggregate parser events.
        #[must_use]
        pub const fn events(self) -> usize {
            self.events
        }

        /// Greatest per-document depth.
        #[must_use]
        pub const fn max_depth(self) -> usize {
            self.max_depth
        }

        /// Number of audited parts.
        #[must_use]
        pub const fn parts(self) -> usize {
            self.parts
        }

        /// Aggregate character-data bytes.
        #[must_use]
        pub const fn text_bytes(self) -> usize {
            self.text_bytes
        }
    }

    /// Package-level audit failure borrowing the failing part name.
    #[derive(Debug)]
    #[non_exhaustive]
    pub enum Error<'a> {
        /// Aggregate package budget exceeded.
        Limit {
            /// Resource name (`parts` or `bytes`).
            resource: &'static str,
            /// Inclusive limit.
            limit: usize,
            /// First value beyond the limit.
            actual: usize,
        },
        /// One named XML part failed verification.
        Part {
            /// Borrowed archive-relative name.
            name: &'a str,
            /// Typed document failure.
            source: DocumentError,
        },
    }

    impl fmt::Display for Error<'_> {
        fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
            match self {
                Self::Limit {
                    resource,
                    limit,
                    actual,
                } => write!(
                    formatter,
                    "XML package {resource} limit {limit} exceeded by {actual}"
                ),
                Self::Part { name, source } => write!(formatter, "XML part '{name}': {source}"),
            }
        }
    }

    impl std::error::Error for Error<'_> {
        fn source(&self) -> Option<&(dyn std::error::Error + 'static)> {
            match self {
                Self::Limit { .. } => None,
                Self::Part { source, .. } => Some(source),
            }
        }
    }

    /// Verifies a sequence of generated OOXML/ODF or referenced XML assets.
    ///
    /// Parts are consumed incrementally and payloads remain borrowed.
    ///
    /// # Errors
    ///
    /// Returns [`Error::Limit`] for an aggregate package-budget breach or
    /// [`Error::Part`] with the borrowed name and typed document failure.
    pub fn verify<'a, I>(parts: I, limits: Limits) -> Result<Report, Error<'a>>
    where
        I: IntoIterator<Item = Part<'a>>,
    {
        let mut report = Report {
            attributes: 0,
            bytes: 0,
            events: 0,
            max_depth: 0,
            parts: 0,
            text_bytes: 0,
        };

        for part in parts {
            report.parts = package_add(report.parts, 1, "parts", limits.parts)?;
            report.bytes = package_add(report.bytes, part.bytes.len(), "bytes", limits.bytes)?;
            let item =
                verify_document(part.bytes, limits.document).map_err(|source| Error::Part {
                    name: part.name,
                    source,
                })?;
            report.attributes = report.attributes.saturating_add(item.attributes());
            report.events = report.events.saturating_add(item.events());
            report.max_depth = report.max_depth.max(item.max_depth());
            report.text_bytes = report.text_bytes.saturating_add(item.text_bytes());
        }
        Ok(report)
    }

    fn package_add<'a>(
        current: usize,
        amount: usize,
        resource: &'static str,
        limit: usize,
    ) -> Result<usize, Error<'a>> {
        let actual = current.saturating_add(amount);
        if actual > limit {
            return Err(Error::Limit {
                resource,
                limit,
                actual,
            });
        }
        Ok(actual)
    }
}
