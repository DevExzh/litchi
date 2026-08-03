//! Typed anchors and opaque payloads for WordprocessingML alternative-format imports.
//!
//! Payloads are deliberately never parsed, executed, fetched, or opened as nested
//! packages. The supported authoring media types follow the Microsoft Word notes
//! in `[MS-OI29500]` §2.1.527 and `[MS-OE376]` §2.1.558.
//!
//! ```
//! use litchi_docx::alt::{Data, Import};
//!
//! let embedded = Import::data(Data::Html(b"<p>opaque</p>".to_vec()));
//! let linked = Import::link("https://example.invalid/import.html")?;
//! # let _ = (embedded, linked);
//! # Ok::<(), litchi_docx::Error>(())
//! ```

use crate::{Error, Result};
use litchi_opc::constants::relationship_type;
use litchi_opc::packuri::PackURI;
use litchi_opc::part::Part as OpcPart;
use quick_xml::XmlVersion;
use quick_xml::events::{BytesStart, Event};
use quick_xml::name::{Namespace, NamespaceResolver, ResolveResult};
use quick_xml::reader::NsReader;
use std::collections::BTreeMap;

const TRANSITIONAL_WORD_NAMESPACE: &[u8] =
    b"http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const STRICT_WORD_NAMESPACE: &[u8] = b"http://purl.oclc.org/ooxml/wordprocessingml/main";
const TRANSITIONAL_RELATIONSHIP_NAMESPACE: &[u8] =
    b"http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const STRICT_RELATIONSHIP_NAMESPACE: &[u8] =
    b"http://purl.oclc.org/ooxml/officeDocument/relationships";
const STRICT_RELATIONSHIP: &str = "http://purl.oclc.org/ooxml/officeDocument/relationships/afChunk";

/// Maximum payload accepted by the safe package-authoring facade.
pub const MAX_DATA_BYTES: usize = 128 * 1024 * 1024;
/// Maximum anchors accepted in one main-document part.
pub const MAX_CHUNKS: usize = 4096;
/// Maximum main-document XML accepted by the bounded scanner.
pub const MAX_XML_BYTES: usize = 32 * 1024 * 1024;
/// Maximum XML nesting accepted by the bounded scanner.
pub const MAX_XML_DEPTH: usize = 256;

const MAX_VISIBILITY_OFFSETS: usize = 1_000_000;
const MAX_MARKED_XML_BYTES: usize = 128 * 1024 * 1024;

/// A validated OPC relationship identifier used by a [`Chunk`].
///
/// This type is intentionally low-level. Package CRUD accepts semantic [`Import`]
/// values and allocates relationship identifiers itself.
#[derive(Debug, Clone, PartialEq, Eq, PartialOrd, Ord, Hash)]
#[repr(transparent)]
pub struct Rel(Box<str>);

impl Rel {
    /// Validate an identifier that can be emitted without XML escaping.
    pub fn new(value: impl Into<String>) -> Result<Self> {
        let value = value.into();
        if value.is_empty()
            || value.len() > 255
            || !value.bytes().all(|byte| {
                byte.is_ascii_alphanumeric() || matches!(byte, b'_' | b'-' | b'.' | b':')
            })
        {
            return Err(invalid(
                "altChunk relationship ID is not a safe XML attribute value",
            ));
        }
        Ok(Self(value.into_boxed_str()))
    }

    /// Borrow the underlying OPC identifier.
    #[inline]
    pub fn as_str(&self) -> &str {
        &self.0
    }
}

impl AsRef<str> for Rel {
    fn as_ref(&self) -> &str {
        self.as_str()
    }
}

/// A validated external target URI that is retained but never accessed.
#[derive(Debug, Clone, PartialEq, Eq, PartialOrd, Ord, Hash)]
#[repr(transparent)]
pub struct Uri(Box<str>);

impl Uri {
    /// Validate a bounded, inert external target.
    pub fn new(value: impl Into<String>) -> Result<Self> {
        let value = value.into();
        if value.is_empty() || value.len() > 32_768 || value.chars().any(char::is_control) {
            return Err(invalid("external altChunk target is empty or unsafe"));
        }
        Ok(Self(value.into_boxed_str()))
    }

    /// Borrow the preserved target URI.
    #[inline]
    pub fn as_str(&self) -> &str {
        &self.0
    }

    /// Move the target URI out without copying it.
    #[inline]
    pub fn into_string(self) -> String {
        self.0.into()
    }
}

impl AsRef<str> for Uri {
    fn as_ref(&self) -> &str {
        self.as_str()
    }
}

/// Namespace family used when emitting a new alternative-format anchor.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub enum Conformance {
    #[default]
    Transitional,
    Strict,
}

impl Conformance {
    /// Relationship type emitted for this conformance family.
    ///
    /// Word uses the case-sensitive `aFChunk` spelling in Transitional files.
    #[inline]
    pub const fn relationship(self) -> &'static str {
        match self {
            Self::Transitional => relationship_type::MS_ALTERNATIVE_FORMAT_IMPORT,
            Self::Strict => STRICT_RELATIONSHIP,
        }
    }

    const fn word_namespace(self) -> &'static str {
        match self {
            Self::Transitional => "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
            Self::Strict => "http://purl.oclc.org/ooxml/wordprocessingml/main",
        }
    }

    const fn relationship_namespace(self) -> &'static str {
        match self {
            Self::Transitional => {
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
            },
            Self::Strict => "http://purl.oclc.org/ooxml/officeDocument/relationships",
        }
    }
}

/// Owned bytes for every alternative-format media type supported by Word.
#[derive(Debug, PartialEq, Eq)]
pub enum Data {
    Docx(Vec<u8>),
    Docm(Vec<u8>),
    Dotx(Vec<u8>),
    Dotm(Vec<u8>),
    Mime(Vec<u8>),
    Html(Vec<u8>),
    Xhtml(Vec<u8>),
    Rtf(Vec<u8>),
    Text(Vec<u8>),
    Xml(Vec<u8>),
}

impl Data {
    /// Canonical media type for the payload.
    pub const fn media_type(&self) -> &'static str {
        match self {
            Self::Docx(_) => {
                "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"
            },
            Self::Docm(_) => "application/vnd.ms-word.document.macroEnabled.main+xml",
            Self::Dotx(_) => {
                "application/vnd.openxmlformats-officedocument.wordprocessingml.template.main+xml"
            },
            Self::Dotm(_) => "application/vnd.ms-word.template.macroEnabledTemplate.main+xml",
            Self::Mime(_) => "message/rfc822",
            Self::Html(_) => "text/html",
            Self::Xhtml(_) => "application/xhtml+xml",
            Self::Rtf(_) => "application/rtf",
            Self::Text(_) => "text/plain",
            Self::Xml(_) => "application/xml",
        }
    }

    /// Conventional package extension for the payload.
    pub const fn extension(&self) -> &'static str {
        match self {
            Self::Docx(_) => "docx",
            Self::Docm(_) => "docm",
            Self::Dotx(_) => "dotx",
            Self::Dotm(_) => "dotm",
            Self::Mime(_) => "eml",
            Self::Html(_) => "html",
            Self::Xhtml(_) => "xhtml",
            Self::Rtf(_) => "rtf",
            Self::Text(_) => "txt",
            Self::Xml(_) => "xml",
        }
    }

    /// Borrow the opaque payload bytes.
    pub fn bytes(&self) -> &[u8] {
        match self {
            Self::Docx(bytes)
            | Self::Docm(bytes)
            | Self::Dotx(bytes)
            | Self::Dotm(bytes)
            | Self::Mime(bytes)
            | Self::Html(bytes)
            | Self::Xhtml(bytes)
            | Self::Rtf(bytes)
            | Self::Text(bytes)
            | Self::Xml(bytes) => bytes,
        }
    }

    /// Payload size in bytes.
    #[inline]
    pub fn len(&self) -> usize {
        self.bytes().len()
    }

    /// Whether the payload is empty.
    #[inline]
    pub fn is_empty(&self) -> bool {
        self.bytes().is_empty()
    }

    /// Enforce the package-authoring resource limit.
    pub fn validate(&self) -> Result<()> {
        if self.len() > MAX_DATA_BYTES {
            return Err(invalid(
                "alternative-format part exceeds the 128 MiB authoring limit",
            ));
        }
        Ok(())
    }

    /// Move the opaque payload bytes out without copying them.
    pub fn into_bytes(self) -> Vec<u8> {
        match self {
            Self::Docx(bytes)
            | Self::Docm(bytes)
            | Self::Dotx(bytes)
            | Self::Dotm(bytes)
            | Self::Mime(bytes)
            | Self::Html(bytes)
            | Self::Xhtml(bytes)
            | Self::Rtf(bytes)
            | Self::Text(bytes)
            | Self::Xml(bytes) => bytes,
        }
    }
}

/// Package target to create for an alternative-format import.
#[derive(Debug, PartialEq, Eq)]
pub enum Import {
    /// An owned payload moved into the package.
    Data(Data),
    /// A validated URI that is preserved but never fetched or interpreted.
    Link(Uri),
}

impl Import {
    /// Wrap an owned internal payload.
    #[inline]
    pub const fn data(data: Data) -> Self {
        Self::Data(data)
    }

    /// Validate an inert external target.
    #[inline]
    pub fn link(uri: impl Into<String>) -> Result<Self> {
        Uri::new(uri).map(Self::Link)
    }
}

/// Resolved target of an alternative-format anchor.
pub enum Target<'a> {
    /// Borrowed internal package part.
    Part(Part<'a>),
    /// Borrowed external URI; it is never accessed.
    Link(&'a str),
}

/// A block-level `<w:altChunk>` import anchor.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Chunk {
    relationship: Rel,
    match_source: Option<bool>,
}

impl Chunk {
    /// Bind a low-level anchor to a validated OPC relationship.
    #[inline]
    pub const fn new(relationship: Rel, match_source: Option<bool>) -> Self {
        Self {
            relationship,
            match_source,
        }
    }

    /// Relationship identifying the alternative-format import target.
    #[inline]
    pub const fn relationship(&self) -> &Rel {
        &self.relationship
    }

    /// Whether imported formatting should match the source formatting.
    ///
    /// `None` means `<w:matchSrc>` was absent. `Some(true)` includes the
    /// empty-element form, while `Some(false)` represents an explicit false.
    #[inline]
    pub const fn match_source(&self) -> Option<bool> {
        self.match_source
    }

    /// Serialize this anchor using an isolated, namespace-complete element.
    pub fn xml(&self, conformance: Conformance) -> String {
        let word_ns = conformance.word_namespace();
        let relationship_ns = conformance.relationship_namespace();
        let opening = format!(
            r#"<w:altChunk xmlns:w="{word_ns}" xmlns:r="{relationship_ns}" r:id="{}""#,
            self.relationship.as_str()
        );
        match self.match_source {
            None => format!("{opening}/>"),
            Some(value) => format!(
                r#"{opening}><w:altChunkPr><w:matchSrc w:val="{}"/></w:altChunkPr></w:altChunk>"#,
                u8::from(value)
            ),
        }
    }
}

/// Recognized MIME family for an opaque alternative-format import part.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum Kind {
    Docx,
    Docm,
    Dotx,
    Dotm,
    Mime,
    Html,
    Xhtml,
    Rtf,
    Text,
    Xml,
    Unknown,
}

impl Kind {
    /// Classify a media type without inspecting its payload.
    pub fn from_media_type(value: &str) -> Self {
        let media_type = value.split(';').next().map_or(value, str::trim);
        const TYPES: &[(&str, Kind)] = &[
            (
                "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml",
                Kind::Docx,
            ),
            (
                "application/vnd.ms-word.document.macroEnabled.main+xml",
                Kind::Docm,
            ),
            (
                "application/vnd.openxmlformats-officedocument.wordprocessingml.template.main+xml",
                Kind::Dotx,
            ),
            (
                "application/vnd.ms-word.template.macroEnabledTemplate.main+xml",
                Kind::Dotm,
            ),
            ("message/rfc822", Kind::Mime),
            ("text/html", Kind::Html),
            ("application/xhtml+xml", Kind::Xhtml),
            ("application/rtf", Kind::Rtf),
            ("text/rtf", Kind::Rtf),
            ("text/plain", Kind::Text),
            ("application/xml", Kind::Xml),
            ("text/xml", Kind::Xml),
        ];
        match TYPES
            .iter()
            .find_map(|(known, kind)| media_type.eq_ignore_ascii_case(known).then_some(*kind))
        {
            Some(kind) => kind,
            None => Self::Unknown,
        }
    }
}

/// A borrowed, opaque alternative-format import payload.
///
/// Access never parses the foreign format, opens nested packages, fetches
/// resources, or performs filesystem or network I/O.
pub struct Part<'a> {
    part: &'a dyn OpcPart,
    kind: Kind,
}

impl<'a> Part<'a> {
    /// Borrow an opaque OPC part without copying its payload.
    pub fn new(part: &'a dyn OpcPart) -> Self {
        Self {
            kind: Kind::from_media_type(part.content_type()),
            part,
        }
    }

    /// OPC part name.
    #[inline]
    pub fn name(&self) -> &PackURI {
        self.part.partname()
    }

    /// Preserved OPC media type.
    #[inline]
    pub fn media_type(&self) -> &str {
        self.part.content_type()
    }

    /// Classified media family.
    #[inline]
    pub const fn kind(&self) -> Kind {
        self.kind
    }

    /// Return the raw OPC part bytes without interpreting them.
    #[inline]
    pub fn bytes(&self) -> &[u8] {
        self.part.blob()
    }
}

/// Whether `value` is a supported alternative-format relationship type.
pub fn is_relationship(value: &str) -> bool {
    matches!(
        value,
        relationship_type::ALTERNATIVE_FORMAT_IMPORT
            | relationship_type::MS_ALTERNATIVE_FORMAT_IMPORT
            | STRICT_RELATIONSHIP
    )
}

struct PendingChunk {
    root_depth: usize,
    start: u32,
    relationship: Rel,
    match_source: Option<bool>,
    saw_properties: bool,
    properties_depth: Option<usize>,
    opaque_depth: Option<usize>,
}

/// Retain offsets whose XML positions survive baseline markup-compatibility
/// processing.
///
/// The returned offsets always refer to `xml`, not to a rewritten MCE view.
/// Input order is preserved. This low-level helper lets package facades retain
/// exact source ranges while selecting only the active `mc:Choice` or
/// `mc:Fallback` branch.
pub fn active(xml: &[u8], offsets: &[u32]) -> Result<Vec<u32>> {
    validate_xml(xml)?;
    let limits = litchi_ooxml_common::mce::ActiveOffsetLimits {
        max_source_bytes: MAX_XML_BYTES,
        max_offsets: MAX_VISIBILITY_OFFSETS,
        max_marked_bytes: MAX_MARKED_XML_BYTES,
        mce: litchi_ooxml_common::mce::MceLimits {
            max_input_bytes: MAX_MARKED_XML_BYTES,
            max_output_bytes: MAX_MARKED_XML_BYTES,
            max_depth: MAX_XML_DEPTH,
            max_namespace_bindings: 4096,
            max_directive_tokens: 4096,
            max_choices_per_alternate: 1024,
        },
    };
    litchi_ooxml_common::mce::active_offsets(
        xml,
        offsets,
        &litchi_ooxml_common::mce::MceCapabilities::default(),
        &limits,
    )
    .map_err(Error::from)
}

/// Parse every altChunk anchor against the full namespace context.
pub fn scan(xml: &[u8]) -> Result<BTreeMap<u32, Chunk>> {
    validate_xml(xml)?;
    let mut reader = NsReader::from_reader(xml);
    let mut depth = 0usize;
    let mut pending: Option<PendingChunk> = None;
    let mut chunks = BTreeMap::new();

    loop {
        let event_start = u32::try_from(reader.buffer_position())
            .map_err(|_| Error::Invalid("altChunk XML offset does not fit u32".into()))?;
        let decoder = reader.decoder();
        let event = reader
            .read_event()
            .map_err(|error| Error::Xml(error.to_string()))?
            .into_owned();
        let resolver = reader.resolver().clone();
        let (namespace, event) = resolver.resolve_event(event);

        match event {
            Event::Start(element) => {
                let event_depth = next_depth(depth)?;
                if pending.is_none()
                    && is_word_namespace(&namespace)
                    && element.local_name().as_ref() == b"altChunk"
                {
                    pending = Some(PendingChunk {
                        root_depth: event_depth,
                        start: event_start,
                        relationship: relationship(&element, decoder, &resolver)?,
                        match_source: None,
                        saw_properties: false,
                        properties_depth: None,
                        opaque_depth: None,
                    });
                } else if let Some(chunk) = pending.as_mut() {
                    parse_child(
                        chunk,
                        event_depth,
                        &namespace,
                        &element,
                        decoder,
                        &resolver,
                        false,
                    )?;
                }
                depth = event_depth;
            },
            Event::Empty(element) => {
                let event_depth = next_depth(depth)?;
                if pending.is_none()
                    && is_word_namespace(&namespace)
                    && element.local_name().as_ref() == b"altChunk"
                {
                    let chunk = Chunk {
                        relationship: relationship(&element, decoder, &resolver)?,
                        match_source: None,
                    };
                    insert_chunk(&mut chunks, event_start, chunk)?;
                } else if let Some(chunk) = pending.as_mut() {
                    parse_child(
                        chunk,
                        event_depth,
                        &namespace,
                        &element,
                        decoder,
                        &resolver,
                        true,
                    )?;
                }
            },
            Event::End(_) => {
                if let Some(chunk) = pending.as_mut()
                    && chunk.opaque_depth == Some(depth)
                {
                    chunk.opaque_depth = None;
                }
                if let Some(chunk) = pending.as_mut()
                    && chunk.properties_depth == Some(depth)
                {
                    chunk.properties_depth = None;
                }
                if pending
                    .as_ref()
                    .is_some_and(|chunk| chunk.root_depth == depth)
                {
                    let chunk = pending
                        .take()
                        .ok_or_else(|| Error::Invalid("missing pending altChunk".into()))?;
                    insert_chunk(
                        &mut chunks,
                        chunk.start,
                        Chunk {
                            relationship: chunk.relationship,
                            match_source: chunk.match_source,
                        },
                    )?;
                }
                depth = depth
                    .checked_sub(1)
                    .ok_or_else(|| Error::Invalid("unexpected altChunk XML end element".into()))?;
            },
            Event::Text(text)
                if pending
                    .as_ref()
                    .is_some_and(|chunk| chunk.opaque_depth.is_none())
                    && text.as_ref().iter().any(|byte| !byte.is_ascii_whitespace()) =>
            {
                return Err(Error::Invalid("altChunk contains unexpected text".into()));
            },
            Event::CData(_) | Event::GeneralRef(_)
                if pending
                    .as_ref()
                    .is_some_and(|chunk| chunk.opaque_depth.is_none()) =>
            {
                return Err(Error::Invalid(
                    "altChunk contains unexpected character data".into(),
                ));
            },
            Event::Eof => {
                if pending.is_some() {
                    return Err(Error::Invalid("unterminated altChunk".into()));
                }
                break;
            },
            _ => {},
        }
    }

    let offsets = chunks.keys().copied().collect::<Vec<_>>();
    let active = active(xml, &offsets)?;
    let mut selected = active.into_iter();
    let mut next = selected.next();
    chunks.retain(|offset, _| {
        if next == Some(*offset) {
            next = selected.next();
            true
        } else {
            false
        }
    });
    Ok(chunks)
}

fn validate_xml(xml: &[u8]) -> Result<()> {
    if xml.len() > MAX_XML_BYTES {
        return Err(invalid(format!(
            "alternative-format scan input exceeds {MAX_XML_BYTES} bytes"
        )));
    }
    Ok(())
}

#[cfg(test)]
fn find(haystack: &[u8], needle: &[u8]) -> Option<usize> {
    if needle.is_empty() {
        return Some(0);
    }
    haystack
        .windows(needle.len())
        .position(|window| window == needle)
}

fn parse_child(
    chunk: &mut PendingChunk,
    depth: usize,
    namespace: &ResolveResult<'_>,
    element: &BytesStart<'_>,
    decoder: quick_xml::encoding::Decoder,
    resolver: &NamespaceResolver,
    empty: bool,
) -> Result<()> {
    if chunk.opaque_depth.is_some() {
        return Ok(());
    }
    let is_word = is_word_namespace(namespace);
    if !is_word {
        if !empty {
            chunk.opaque_depth = Some(depth);
        }
        return Ok(());
    }
    let properties_depth = chunk
        .root_depth
        .checked_add(1)
        .ok_or_else(|| invalid("altChunk XML nesting is too deep"))?;
    let value_depth = properties_depth
        .checked_add(1)
        .ok_or_else(|| invalid("altChunk XML nesting is too deep"))?;
    if depth == properties_depth
        && element.local_name().as_ref() == b"altChunkPr"
        && !chunk.saw_properties
    {
        chunk.saw_properties = true;
        if !empty {
            chunk.properties_depth = Some(depth);
        }
        return Ok(());
    }
    if depth == value_depth
        && chunk.properties_depth == Some(properties_depth)
        && element.local_name().as_ref() == b"matchSrc"
        && chunk.match_source.is_none()
    {
        chunk.match_source = Some(parse_on_off(
            element,
            decoder,
            resolver,
            is_transitional_word_namespace(namespace),
        )?);
        return Ok(());
    }
    Err(Error::Invalid("altChunk has invalid child content".into()))
}

fn relationship(
    element: &BytesStart<'_>,
    decoder: quick_xml::encoding::Decoder,
    resolver: &NamespaceResolver,
) -> Result<Rel> {
    let mut value = None;
    for attribute in element.attributes() {
        let attribute = attribute.map_err(|error| Error::Xml(error.to_string()))?;
        if attribute.key.local_name().as_ref() != b"id" {
            continue;
        }
        let (namespace, _) = resolver.resolve_attribute(attribute.key);
        let valid_namespace = matches!(
            namespace,
            ResolveResult::Bound(Namespace(uri))
                if uri == TRANSITIONAL_RELATIONSHIP_NAMESPACE
                    || uri == STRICT_RELATIONSHIP_NAMESPACE
        );
        if !valid_namespace {
            continue;
        }
        if value.is_some() {
            return Err(Error::Invalid(
                "altChunk has duplicate relationship IDs".into(),
            ));
        }
        value = Some(
            attribute
                .decoded_and_normalized_value(XmlVersion::Explicit1_0, decoder)
                .map_err(|error| Error::Xml(error.to_string()))?
                .into_owned(),
        );
    }
    let value = value
        .filter(|value| !value.is_empty())
        .ok_or_else(|| Error::Invalid("altChunk lacks a relationship ID".into()))?;
    Rel::new(value)
}

fn parse_on_off(
    element: &BytesStart<'_>,
    decoder: quick_xml::encoding::Decoder,
    resolver: &NamespaceResolver,
    allow_legacy_values: bool,
) -> Result<bool> {
    let mut value = None;
    for attribute in element.attributes() {
        let attribute = attribute.map_err(|error| Error::Xml(error.to_string()))?;
        if attribute.key.local_name().as_ref() != b"val" {
            continue;
        }
        let (namespace, _) = resolver.resolve_attribute(attribute.key);
        if !is_word_namespace(&namespace) && !matches!(namespace, ResolveResult::Unbound) {
            continue;
        }
        if value.is_some() {
            return Err(Error::Invalid("matchSrc has duplicate values".into()));
        }
        value = Some(
            attribute
                .decoded_and_normalized_value(XmlVersion::Explicit1_0, decoder)
                .map_err(|error| Error::Xml(error.to_string()))?
                .into_owned(),
        );
    }
    match value.as_deref() {
        None | Some("true" | "1") => Ok(true),
        Some("false" | "0") => Ok(false),
        Some("on") if allow_legacy_values => Ok(true),
        Some("off") if allow_legacy_values => Ok(false),
        Some(value) => Err(Error::Invalid(format!("invalid matchSrc value '{value}'"))),
    }
}

fn insert_chunk(chunks: &mut BTreeMap<u32, Chunk>, start: u32, chunk: Chunk) -> Result<()> {
    if chunks.len() >= MAX_CHUNKS {
        return Err(invalid("alternative-format anchor limit exceeded"));
    }
    if chunks.insert(start, chunk).is_some() {
        return Err(Error::Invalid("duplicate altChunk XML position".into()));
    }
    Ok(())
}

fn is_word_namespace(namespace: &ResolveResult<'_>) -> bool {
    matches!(
        namespace,
        ResolveResult::Bound(Namespace(uri))
            if *uri == TRANSITIONAL_WORD_NAMESPACE || *uri == STRICT_WORD_NAMESPACE
    )
}

fn is_transitional_word_namespace(namespace: &ResolveResult<'_>) -> bool {
    matches!(
        namespace,
        ResolveResult::Bound(Namespace(uri)) if *uri == TRANSITIONAL_WORD_NAMESPACE
    )
}

fn next_depth(depth: usize) -> Result<usize> {
    let next = depth
        .checked_add(1)
        .ok_or_else(|| invalid("alternative-format XML nesting overflowed"))?;
    if next > MAX_XML_DEPTH {
        return Err(invalid(format!(
            "alternative-format XML exceeds {MAX_XML_DEPTH} nesting levels"
        )));
    }
    Ok(next)
}

fn invalid(message: impl Into<String>) -> Error {
    Error::Invalid(message.into())
}

#[cfg(test)]
mod tests {
    use super::*;
    use litchi_opc::part::BlobPart;

    #[test]
    fn typed_payloads_cover_word_supported_media_types_and_move_bytes() {
        let cases = [
            (Data::Docx(vec![1]), Kind::Docx, "docx"),
            (Data::Docm(vec![2]), Kind::Docm, "docm"),
            (Data::Dotx(vec![3]), Kind::Dotx, "dotx"),
            (Data::Dotm(vec![4]), Kind::Dotm, "dotm"),
            (Data::Mime(vec![5]), Kind::Mime, "eml"),
            (Data::Html(vec![6]), Kind::Html, "html"),
            (Data::Xhtml(vec![7]), Kind::Xhtml, "xhtml"),
            (Data::Rtf(vec![8]), Kind::Rtf, "rtf"),
            (Data::Text(vec![9]), Kind::Text, "txt"),
            (Data::Xml(vec![10]), Kind::Xml, "xml"),
        ];
        for (data, kind, extension) in cases {
            assert_eq!(Kind::from_media_type(data.media_type()), kind);
            assert_eq!(data.extension(), extension);
        }

        let bytes = vec![11, 12, 13];
        let pointer = bytes.as_ptr();
        let moved = Data::Html(bytes).into_bytes();
        assert_eq!(moved.as_ptr(), pointer);
    }

    #[test]
    fn media_classification_is_parameter_tolerant_and_preserves_unknowns() {
        assert_eq!(
            Kind::from_media_type(" Text/HTML ; charset=utf-8"),
            Kind::Html
        );
        assert_eq!(Kind::from_media_type("text/rtf"), Kind::Rtf);
        assert_eq!(
            Kind::from_media_type("application/x-vendor-opaque"),
            Kind::Unknown
        );
    }

    #[test]
    fn identifiers_and_external_targets_are_validated_once() {
        assert!(Rel::new("rId42").is_ok());
        assert!(Rel::new("bad&value").is_err());
        assert!(Import::link("https://example.invalid/chunk.html").is_ok());
        assert!(Import::link("https://example.invalid/\nchunk").is_err());
    }

    #[test]
    fn scans_strict_and_transitional_anchors_in_source_order() {
        let xml = br#"<s:document xmlns:s="http://purl.oclc.org/ooxml/wordprocessingml/main" xmlns:q="http://purl.oclc.org/ooxml/officeDocument/relationships"><s:body><s:altChunk q:id="first"><s:altChunkPr><s:matchSrc s:val="0"/></s:altChunkPr></s:altChunk><s:altChunk q:id="second"/></s:body></s:document>"#;
        let chunks = scan(xml).unwrap().into_values().collect::<Vec<_>>();
        assert_eq!(chunks.len(), 2);
        assert_eq!(chunks[0].relationship().as_str(), "first");
        assert_eq!(chunks[0].match_source(), Some(false));
        assert_eq!(chunks[1].relationship().as_str(), "second");
        assert_eq!(chunks[1].match_source(), None);
    }

    #[test]
    fn enforces_conformance_specific_match_source_values() {
        const STRICT_WORD: &str = "http://purl.oclc.org/ooxml/wordprocessingml/main";
        const TRANSITIONAL_WORD: &str =
            "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
        const STRICT_RELATIONSHIPS: &str =
            "http://purl.oclc.org/ooxml/officeDocument/relationships";
        const TRANSITIONAL_RELATIONSHIPS: &str =
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
        let document = |word_namespace: &str, relationship_namespace: &str, value: &str| {
            format!(
                r#"<w:document xmlns:w="{word_namespace}" xmlns:r="{relationship_namespace}"><w:body><w:altChunk r:id="chunk"><w:altChunkPr><w:matchSrc w:val="{value}"/></w:altChunkPr></w:altChunk></w:body></w:document>"#
            )
        };

        for value in ["true", "1", "false", "0"] {
            assert!(
                scan(document(STRICT_WORD, STRICT_RELATIONSHIPS, value).as_bytes()).is_ok(),
                "{value}"
            );
        }
        for value in ["on", "off"] {
            assert!(
                scan(document(STRICT_WORD, STRICT_RELATIONSHIPS, value).as_bytes()).is_err(),
                "Strict unexpectedly accepted {value}"
            );
            assert!(
                scan(document(TRANSITIONAL_WORD, TRANSITIONAL_RELATIONSHIPS, value).as_bytes())
                    .is_ok(),
                "Transitional rejected {value}"
            );
        }
    }

    #[test]
    fn mce_selects_one_branch_without_rewriting_source_offsets() {
        let xml = br#"<q:document xmlns:q="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:rel="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:x="urn:unsupported"><q:body><mc:AlternateContent><mc:Choice Requires="x"><q:altChunk rel:id="inactive"/></mc:Choice><mc:Fallback><q:altChunk rel:id="fallback"/></mc:Fallback></mc:AlternateContent><mc:AlternateContent><mc:Choice Requires="q"><q:altChunk rel:id="choice"/></mc:Choice><mc:Fallback><q:altChunk rel:id="inactive-fallback"/></mc:Fallback></mc:AlternateContent></q:body></q:document>"#;
        let fallback = find(xml, br#"<q:altChunk rel:id="fallback"/>"#).unwrap();
        let choice = find(xml, br#"<q:altChunk rel:id="choice"/>"#).unwrap();

        let chunks = scan(xml).unwrap();
        assert_eq!(chunks.len(), 2);
        assert_eq!(
            chunks.keys().copied().collect::<Vec<_>>(),
            vec![
                u32::try_from(fallback).unwrap(),
                u32::try_from(choice).unwrap()
            ]
        );
        assert_eq!(
            chunks
                .values()
                .map(|chunk| chunk.relationship().as_str())
                .collect::<Vec<_>>(),
            vec!["fallback", "choice"]
        );
    }

    #[test]
    fn rejects_missing_duplicate_or_unsafe_relationships_and_invalid_children() {
        let wrapper = |anchor: &str| {
            format!(
                r#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:q="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:bad="urn:bad"><w:body>{anchor}</w:body></w:document>"#
            )
        };
        for anchor in [
            r#"<w:altChunk/>"#,
            r#"<w:altChunk bad:id="x"/>"#,
            r#"<w:altChunk r:id="x" q:id="y"/>"#,
            r#"<w:altChunk r:id="bad&amp;id"/>"#,
            r#"<w:altChunk r:id="x"><w:altChunkPr/><w:altChunkPr/></w:altChunk>"#,
            r#"<w:altChunk r:id="x"><w:altChunkPr><w:matchSrc w:val="maybe"/></w:altChunkPr></w:altChunk>"#,
        ] {
            assert!(scan(wrapper(anchor).as_bytes()).is_err(), "{anchor}");
        }
    }

    #[test]
    fn rejects_anchor_count_resource_exhaustion() {
        let mut xml = String::from(
            r#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body>"#,
        );
        for index in 0..=MAX_CHUNKS {
            xml.push_str(&format!(r#"<w:altChunk r:id="rId{index}"/>"#));
        }
        xml.push_str("</w:body></w:document>");
        assert!(scan(xml.as_bytes()).is_err());
    }

    #[test]
    fn rejects_xml_byte_and_depth_resource_exhaustion() {
        let oversized = vec![b' '; MAX_XML_BYTES + 1];
        assert!(scan(&oversized).is_err());

        let mut deep = "<x>".repeat(MAX_XML_DEPTH + 1);
        deep.push_str(&"</x>".repeat(MAX_XML_DEPTH + 1));
        assert!(scan(deep.as_bytes()).is_err());
    }

    #[test]
    fn emitted_anchors_round_trip_both_conformance_families() {
        let chunk = Chunk::new(Rel::new("rIdAlt1").unwrap(), Some(true));
        for conformance in [Conformance::Transitional, Conformance::Strict] {
            let xml = chunk.xml(conformance);
            let parsed = scan(xml.as_bytes()).unwrap().into_values().next().unwrap();
            assert_eq!(parsed, chunk);
            assert!(xml.contains(conformance.relationship_namespace()));
        }
    }

    #[test]
    fn part_lends_original_bytes_and_classifies_without_interpreting() {
        let bytes = b"opaque foreign payload".to_vec();
        let pointer = bytes.as_ptr();
        let raw = BlobPart::new(
            PackURI::new("/word/chunk.vendor").unwrap(),
            "application/x-vendor-opaque".into(),
            bytes,
        );
        let part = Part::new(&raw);
        assert_eq!(part.name().as_str(), "/word/chunk.vendor");
        assert_eq!(part.media_type(), "application/x-vendor-opaque");
        assert_eq!(part.kind(), Kind::Unknown);
        assert_eq!(part.bytes().as_ptr(), pointer);
    }

    #[test]
    fn recognizes_iso_word_and_strict_relationship_dialects() {
        assert!(is_relationship(
            relationship_type::ALTERNATIVE_FORMAT_IMPORT
        ));
        assert!(is_relationship(
            relationship_type::MS_ALTERNATIVE_FORMAT_IMPORT
        ));
        assert!(is_relationship(STRICT_RELATIONSHIP));
        assert!(!is_relationship(relationship_type::IMAGE));
        assert_eq!(
            Conformance::Transitional.relationship(),
            relationship_type::MS_ALTERNATIVE_FORMAT_IMPORT
        );
    }
}
