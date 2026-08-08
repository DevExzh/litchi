//! Bounded, inert inspection of ODF Dynamic Data Exchange metadata.
//!
//! This module models declarations and cached XML only. It contains no DDE
//! client, refresh path, resolver, process launcher, or ambient I/O.

use core::fmt;
use quick_xml::{
    XmlVersion,
    events::{BytesStart, Event},
    name::{Namespace, ResolveResult},
    reader::NsReader,
};
use std::{ops::Range, sync::Arc};

const OFFICE: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:office:1.0";
const TABLE: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:table:1.0";
const MAX_INPUT_BYTES: usize = 256 * 1024 * 1024;
const MAX_LINKS: usize = 65_536;
const MAX_SHEET_SOURCES: usize = 65_536;
const MAX_TEXT_BYTES: usize = 65_536;
const MAX_CACHED_TABLE_BYTES: usize = 64 * 1024 * 1024;
const MAX_DEPTH: usize = 1_024;

/// A DDE metadata inspection result.
pub type Result<T> = std::result::Result<T, Error>;

/// Errors produced while inspecting inert DDE metadata.
#[derive(Clone, Debug, PartialEq, Eq)]
#[non_exhaustive]
pub enum Error {
    /// A configured or hard resource limit was exceeded.
    ResourceLimit {
        /// The bounded resource.
        resource: &'static str,
        /// The observed or configured value.
        actual: usize,
        /// The maximum accepted value.
        maximum: usize,
    },
    /// The XML stream could not be decoded.
    InvalidXml(String),
    /// The document has invalid DDE structure or content.
    InvalidStructure(String),
    /// An XML byte position cannot be represented on this platform.
    PositionOverflow,
}

impl fmt::Display for Error {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Self::ResourceLimit {
                resource,
                actual,
                maximum,
            } => write!(
                formatter,
                "{resource} limit exceeded: observed {actual}, maximum {maximum}"
            ),
            Self::InvalidXml(message) => write!(formatter, "invalid XML: {message}"),
            Self::InvalidStructure(message) => formatter.write_str(message),
            Self::PositionOverflow => formatter.write_str("XML byte position exceeds usize"),
        }
    }
}

impl std::error::Error for Error {}

#[derive(Clone, Copy, PartialEq, Eq)]
enum NamespaceKind {
    Office,
    Table,
    Other,
}

/// Resource limits for inert DDE inspection.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub struct Limits {
    input_bytes: usize,
    links: usize,
    sheet_sources: usize,
    text_bytes: usize,
    cached_table_bytes: usize,
    depth: usize,
}

impl Default for Limits {
    fn default() -> Self {
        Self {
            input_bytes: MAX_INPUT_BYTES,
            links: MAX_LINKS,
            sheet_sources: MAX_SHEET_SOURCES,
            text_bytes: MAX_TEXT_BYTES,
            cached_table_bytes: MAX_CACHED_TABLE_BYTES,
            depth: MAX_DEPTH,
        }
    }
}

impl Limits {
    #[must_use]
    pub const fn with_input_bytes(mut self, value: usize) -> Self {
        self.input_bytes = value;
        self
    }

    #[must_use]
    pub const fn with_links(mut self, value: usize) -> Self {
        self.links = value;
        self
    }

    #[must_use]
    pub const fn with_sheet_sources(mut self, value: usize) -> Self {
        self.sheet_sources = value;
        self
    }

    #[must_use]
    pub const fn with_text_bytes(mut self, value: usize) -> Self {
        self.text_bytes = value;
        self
    }

    #[must_use]
    pub const fn with_cached_table_bytes(mut self, value: usize) -> Self {
        self.cached_table_bytes = value;
        self
    }

    #[must_use]
    pub const fn with_depth(mut self, value: usize) -> Self {
        self.depth = value;
        self
    }

    fn validate(self) -> Result<Self> {
        for (name, value, ceiling) in [
            ("input bytes", self.input_bytes, MAX_INPUT_BYTES),
            ("links", self.links, MAX_LINKS),
            ("sheet sources", self.sheet_sources, MAX_SHEET_SOURCES),
            ("text bytes", self.text_bytes, MAX_TEXT_BYTES),
            (
                "cached table bytes",
                self.cached_table_bytes,
                MAX_CACHED_TABLE_BYTES,
            ),
            ("XML depth", self.depth, MAX_DEPTH),
        ] {
            if value > ceiling {
                return Err(Error::ResourceLimit {
                    resource: name,
                    actual: value,
                    maximum: ceiling,
                });
            }
        }
        Ok(self)
    }
}

/// ODF conversion policy retained without applying it.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
#[non_exhaustive]
pub enum ConversionMode {
    /// No conversion mode was specified.
    Unspecified,
    IntoDefaultStyleDataStyle,
    IntoEnglishNumber,
    KeepText,
}

/// Whether a DDE source requests automatic updates.
#[derive(Clone, Copy, Debug, Default, PartialEq, Eq)]
#[non_exhaustive]
pub enum AutomaticUpdate {
    /// The source did not specify an update policy.
    #[default]
    Unspecified,
    /// Automatic updates were requested by the source document.
    Enabled,
    /// Automatic updates were explicitly disabled.
    Disabled,
}

impl From<Option<bool>> for AutomaticUpdate {
    fn from(value: Option<bool>) -> Self {
        match value {
            Some(true) => Self::Enabled,
            Some(false) => Self::Disabled,
            None => Self::Unspecified,
        }
    }
}

impl ConversionMode {
    fn parse(value: &str) -> Result<Self> {
        match value {
            "into-default-style-data-style" => Ok(Self::IntoDefaultStyleDataStyle),
            "into-english-number" => Ok(Self::IntoEnglishNumber),
            "keep-text" => Ok(Self::KeepText),
            _ => Err(invalid(format!("invalid office:conversion-mode '{value}'"))),
        }
    }
}

/// A non-executing `office:dde-source` declaration.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct Source {
    application: String,
    topic: String,
    item: String,
    name: Option<String>,
    conversion_mode: ConversionMode,
    automatic_update: AutomaticUpdate,
}

impl Source {
    /// Creates a detached, inert DDE source descriptor.
    ///
    /// # Errors
    ///
    /// Returns an error when a required identifier is empty, too large, or
    /// contains invalid XML character data.
    pub fn new(
        application: impl Into<String>,
        topic: impl Into<String>,
        item: impl Into<String>,
    ) -> Result<Self> {
        let source = Self {
            application: application.into(),
            topic: topic.into(),
            item: item.into(),
            name: None,
            conversion_mode: ConversionMode::Unspecified,
            automatic_update: AutomaticUpdate::Unspecified,
        };
        validate_required_source_value("office:dde-application", &source.application)?;
        validate_required_source_value("office:dde-topic", &source.topic)?;
        validate_required_source_value("office:dde-item", &source.item)?;
        Ok(source)
    }

    /// Returns a copy with a checked optional source name.
    ///
    /// # Errors
    ///
    /// Returns an error when `name` is empty, too large, or contains invalid
    /// XML character data.
    pub fn named(mut self, source_name: impl Into<String>) -> Result<Self> {
        let name = source_name.into();
        validate_required_source_value("office:name", &name)?;
        self.name = Some(name);
        Ok(self)
    }

    /// Returns a copy with the requested conversion policy.
    #[must_use]
    pub const fn with_conversion_mode(mut self, mode: ConversionMode) -> Self {
        self.conversion_mode = mode;
        self
    }

    /// Returns a copy with the requested automatic-update policy.
    #[must_use]
    pub const fn with_automatic_update(mut self, policy: AutomaticUpdate) -> Self {
        self.automatic_update = policy;
        self
    }

    #[must_use]
    pub fn application(&self) -> &str {
        &self.application
    }

    #[must_use]
    pub fn topic(&self) -> &str {
        &self.topic
    }

    #[must_use]
    pub fn item(&self) -> &str {
        &self.item
    }

    #[must_use]
    pub fn name(&self) -> Option<&str> {
        self.name.as_deref()
    }

    #[must_use]
    pub const fn conversion_mode(&self) -> ConversionMode {
        self.conversion_mode
    }

    #[must_use]
    pub const fn automatic_update(&self) -> AutomaticUpdate {
        self.automatic_update
    }
}

/// One formula DDE link and its exact cached `table:table` subtree.
#[derive(Clone, Debug)]
pub struct Link {
    source: Source,
    content: Arc<str>,
    cached_table: Range<usize>,
}

impl Link {
    #[must_use]
    pub fn source(&self) -> &Source {
        &self.source
    }

    /// Return cached table XML without copying or interpreting its values.
    #[must_use]
    pub fn cached_table_xml(&self) -> &str {
        &self.content[self.cached_table.clone()]
    }
}

/// A sheet-local DDE source declaration.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct SheetSource {
    sheet: String,
    source: Source,
}

impl SheetSource {
    #[must_use]
    pub fn sheet(&self) -> &str {
        &self.sheet
    }

    #[must_use]
    pub fn source(&self) -> &Source {
        &self.source
    }
}

/// Immutable source-bound inventory of all spreadsheet DDE declarations.
#[derive(Clone, Debug)]
pub struct Snapshot {
    content: Arc<str>,
    links: Vec<Link>,
    sheet_sources: Vec<SheetSource>,
}

impl Snapshot {
    /// Parse the default-bounded inert DDE inventory.
    ///
    /// # Errors
    ///
    /// Returns an error when XML is malformed, violates the ODF DDE grammar,
    /// or exceeds a default resource limit.
    pub fn parse(content_xml: &str) -> Result<Self> {
        Self::parse_with(content_xml, Limits::default())
    }

    /// Parse the inert DDE inventory under caller-provided resource limits.
    ///
    /// # Errors
    ///
    /// Returns an error when XML is malformed, violates the ODF DDE grammar,
    /// or exceeds `limits`.
    pub fn parse_with(content_xml: &str, requested_limits: Limits) -> Result<Self> {
        let limits = requested_limits.validate()?;
        if content_xml.len() > limits.input_bytes {
            return Err(invalid("content.xml exceeds the DDE input limit"));
        }
        let content: Arc<str> = Arc::from(content_xml);
        let mut reader = NsReader::from_str(content_xml);
        reader.config_mut().check_end_names = true;
        reader.config_mut().trim_text(false);
        let mut buffer = Vec::new();
        let mut depth = 0usize;
        let mut spreadsheet_depth = None;
        let mut links_depth = None;
        let mut links_seen = false;
        let mut link: Option<LinkBuilder> = None;
        let mut source_depth = None;
        let mut cached_depth = None;
        let mut sheet: Option<SheetBuilder> = None;
        let mut links = Vec::new();
        let mut sheet_sources = Vec::new();

        loop {
            let event_start = xml_position(&reader)?;
            let (resolved_namespace, raw_event) = reader
                .read_resolved_event_into(&mut buffer)
                .map_err(|error| Error::InvalidXml(error.to_string()))?;
            let namespace = namespace_kind(&resolved_namespace);
            let event = raw_event.into_owned();
            let event_end = xml_position(&reader)?;
            match event {
                Event::Start(element) => {
                    depth = depth
                        .checked_add(1)
                        .ok_or_else(|| invalid("XML depth overflow"))?;
                    if depth > limits.depth {
                        return Err(invalid("DDE XML exceeds the nesting limit"));
                    }
                    if is(
                        namespace,
                        element.local_name().as_ref(),
                        NamespaceKind::Office,
                        b"spreadsheet",
                    ) {
                        if spreadsheet_depth.replace(depth).is_some() {
                            return Err(invalid("duplicate or nested office:spreadsheet"));
                        }
                    } else if spreadsheet_depth.is_some_and(|value| depth == value + 1)
                        && is(
                            namespace,
                            element.local_name().as_ref(),
                            NamespaceKind::Table,
                            b"dde-links",
                        )
                    {
                        if links_seen || links_depth.replace(depth).is_some() {
                            return Err(invalid("duplicate table:dde-links owner"));
                        }
                        links_seen = true;
                    } else if links_depth.is_some_and(|value| depth == value + 1)
                        && is(
                            namespace,
                            element.local_name().as_ref(),
                            NamespaceKind::Table,
                            b"dde-link",
                        )
                    {
                        if link.replace(LinkBuilder::new(depth)).is_some() {
                            return Err(invalid("nested table:dde-link"));
                        }
                    } else if link.as_ref().is_some_and(|value| depth == value.depth + 1)
                        && is(
                            namespace,
                            element.local_name().as_ref(),
                            NamespaceKind::Office,
                            b"dde-source",
                        )
                    {
                        let value = parse_source(&element, &reader, limits.text_bytes)?;
                        let Some(builder) = link.as_mut() else {
                            return Err(invalid("DDE link parser state is missing"));
                        };
                        if builder.source.replace(value).is_some() || builder.cached.is_some() {
                            return Err(invalid("office:dde-source must be the first link child"));
                        }
                        source_depth = Some(depth);
                    } else if link.as_ref().is_some_and(|value| depth == value.depth + 1)
                        && is(
                            namespace,
                            element.local_name().as_ref(),
                            NamespaceKind::Table,
                            b"table",
                        )
                    {
                        let Some(builder) = link.as_mut() else {
                            return Err(invalid("DDE link parser state is missing"));
                        };
                        if builder.source.is_none() || builder.cached.is_some() {
                            return Err(invalid("cached table must follow exactly one DDE source"));
                        }
                        builder.cached = Some(event_start..event_start);
                        cached_depth = Some(depth);
                    } else if source_depth.is_some() {
                        return Err(invalid("office:dde-source must not contain child elements"));
                    } else if link.is_some() && cached_depth.is_none() && source_depth.is_none() {
                        return Err(invalid("unsupported child in table:dde-link"));
                    } else if spreadsheet_depth.is_some_and(|value| depth == value + 1)
                        && is(
                            namespace,
                            element.local_name().as_ref(),
                            NamespaceKind::Table,
                            b"table",
                        )
                    {
                        sheet = Some(SheetBuilder {
                            depth,
                            name: required_attr(
                                &element,
                                &reader,
                                TABLE,
                                b"name",
                                limits.text_bytes,
                            )?,
                            source_seen: false,
                        });
                    } else if sheet.as_ref().is_some_and(|value| depth == value.depth + 1)
                        && is(
                            namespace,
                            element.local_name().as_ref(),
                            NamespaceKind::Office,
                            b"dde-source",
                        )
                    {
                        let value = parse_source(&element, &reader, limits.text_bytes)?;
                        let Some(current) = sheet.as_mut() else {
                            return Err(invalid("DDE sheet parser state is missing"));
                        };
                        if current.source_seen {
                            return Err(invalid("duplicate sheet office:dde-source"));
                        }
                        current.source_seen = true;
                        sheet_sources.push(SheetSource {
                            sheet: current.name.clone(),
                            source: value,
                        });
                        if sheet_sources.len() > limits.sheet_sources {
                            return Err(invalid("sheet DDE source count exceeds its limit"));
                        }
                        source_depth = Some(depth);
                    }
                },
                Event::Empty(element) => {
                    let event_depth = depth + 1;
                    if is(
                        namespace,
                        element.local_name().as_ref(),
                        NamespaceKind::Table,
                        b"dde-links",
                    ) && spreadsheet_depth.is_some_and(|value| event_depth == value + 1)
                    {
                        return Err(invalid("table:dde-links must contain a link"));
                    } else if is(
                        namespace,
                        element.local_name().as_ref(),
                        NamespaceKind::Table,
                        b"dde-link",
                    ) && links_depth.is_some_and(|value| event_depth == value + 1)
                    {
                        return Err(invalid("table:dde-link requires a source and cached table"));
                    } else if source_depth.is_some() {
                        return Err(invalid("office:dde-source must not contain child elements"));
                    } else if link
                        .as_ref()
                        .is_some_and(|value| event_depth == value.depth + 1)
                        && is(
                            namespace,
                            element.local_name().as_ref(),
                            NamespaceKind::Office,
                            b"dde-source",
                        )
                    {
                        let value = parse_source(&element, &reader, limits.text_bytes)?;
                        let Some(builder) = link.as_mut() else {
                            return Err(invalid("DDE link parser state is missing"));
                        };
                        if builder.source.replace(value).is_some() || builder.cached.is_some() {
                            return Err(invalid("office:dde-source must be the first link child"));
                        }
                    } else if link
                        .as_ref()
                        .is_some_and(|value| event_depth == value.depth + 1)
                        && is(
                            namespace,
                            element.local_name().as_ref(),
                            NamespaceKind::Table,
                            b"table",
                        )
                    {
                        let Some(builder) = link.as_mut() else {
                            return Err(invalid("DDE link parser state is missing"));
                        };
                        if builder.source.is_none() || builder.cached.is_some() {
                            return Err(invalid("cached table must follow exactly one DDE source"));
                        }
                        if event_end - event_start > limits.cached_table_bytes {
                            return Err(invalid("DDE cached table exceeds its byte limit"));
                        }
                        builder.cached = Some(event_start..event_end);
                    } else if sheet
                        .as_ref()
                        .is_some_and(|value| event_depth == value.depth + 1)
                        && is(
                            namespace,
                            element.local_name().as_ref(),
                            NamespaceKind::Office,
                            b"dde-source",
                        )
                    {
                        let value = parse_source(&element, &reader, limits.text_bytes)?;
                        let Some(current) = sheet.as_mut() else {
                            return Err(invalid("DDE sheet parser state is missing"));
                        };
                        if current.source_seen {
                            return Err(invalid("duplicate sheet office:dde-source"));
                        }
                        current.source_seen = true;
                        sheet_sources.push(SheetSource {
                            sheet: current.name.clone(),
                            source: value,
                        });
                        if sheet_sources.len() > limits.sheet_sources {
                            return Err(invalid("sheet DDE source count exceeds its limit"));
                        }
                    }
                },
                Event::End(element) => {
                    if cached_depth == Some(depth)
                        && is(
                            namespace,
                            element.local_name().as_ref(),
                            NamespaceKind::Table,
                            b"table",
                        )
                    {
                        let Some(builder) = link.as_mut() else {
                            return Err(invalid("DDE cached-table parser state is missing"));
                        };
                        let Some(cached) = builder.cached.as_ref() else {
                            return Err(invalid("DDE cached-table start is missing"));
                        };
                        let cached_start = cached.start;
                        if event_end - cached_start > limits.cached_table_bytes {
                            return Err(invalid("DDE cached table exceeds its byte limit"));
                        }
                        builder.cached = Some(cached_start..event_end);
                        cached_depth = None;
                    } else if source_depth == Some(depth)
                        && is(
                            namespace,
                            element.local_name().as_ref(),
                            NamespaceKind::Office,
                            b"dde-source",
                        )
                    {
                        source_depth = None;
                    } else if link.as_ref().is_some_and(|value| depth == value.depth)
                        && is(
                            namespace,
                            element.local_name().as_ref(),
                            NamespaceKind::Table,
                            b"dde-link",
                        )
                    {
                        let Some(builder) = link.take() else {
                            return Err(invalid("DDE link parser state is missing"));
                        };
                        links.push(builder.finish(Arc::clone(&content))?);
                        if links.len() > limits.links {
                            return Err(invalid("DDE link count exceeds its limit"));
                        }
                    } else if links_depth == Some(depth)
                        && is(
                            namespace,
                            element.local_name().as_ref(),
                            NamespaceKind::Table,
                            b"dde-links",
                        )
                    {
                        if links.is_empty() {
                            return Err(invalid("table:dde-links must contain a link"));
                        }
                        links_depth = None;
                    } else if sheet.as_ref().is_some_and(|value| depth == value.depth)
                        && is(
                            namespace,
                            element.local_name().as_ref(),
                            NamespaceKind::Table,
                            b"table",
                        )
                    {
                        sheet = None;
                    } else if spreadsheet_depth == Some(depth)
                        && is(
                            namespace,
                            element.local_name().as_ref(),
                            NamespaceKind::Office,
                            b"spreadsheet",
                        )
                    {
                        spreadsheet_depth = None;
                    }
                    depth = depth.saturating_sub(1);
                },
                Event::Text(text) if source_depth.is_some() => {
                    let value = text
                        .xml_content(XmlVersion::Explicit1_0)
                        .map_err(|error| invalid(format!("invalid DDE source text: {error}")))?;
                    if !value.trim().is_empty() {
                        return Err(invalid("office:dde-source must be empty"));
                    }
                },
                Event::CData(text) if source_depth.is_some() => {
                    let value = text
                        .xml_content(XmlVersion::Explicit1_0)
                        .map_err(|error| invalid(format!("invalid DDE source CDATA: {error}")))?;
                    if !value.trim().is_empty() {
                        return Err(invalid("office:dde-source must be empty"));
                    }
                },
                Event::Text(text) if link.is_some() || links_depth.is_some() => {
                    let value = text
                        .xml_content(XmlVersion::Explicit1_0)
                        .map_err(|error| invalid(format!("invalid DDE container text: {error}")))?;
                    if !value.trim().is_empty() {
                        return Err(invalid("DDE containers must not contain text"));
                    }
                },
                Event::CData(text) if link.is_some() || links_depth.is_some() => {
                    let value = text.xml_content(XmlVersion::Explicit1_0).map_err(|error| {
                        invalid(format!("invalid DDE container CDATA: {error}"))
                    })?;
                    if !value.trim().is_empty() {
                        return Err(invalid("DDE containers must not contain CDATA"));
                    }
                },
                Event::GeneralRef(_) if link.is_some() || links_depth.is_some() => {
                    return Err(invalid("DDE containers must not contain entity references"));
                },
                Event::DocType(_) => return Err(invalid("DTD content is not accepted")),
                Event::Eof => break,
                Event::Text(_)
                | Event::CData(_)
                | Event::GeneralRef(_)
                | Event::Comment(_)
                | Event::Decl(_)
                | Event::PI(_) => {},
            }
            buffer.clear();
        }
        if depth != 0 || link.is_some() || cached_depth.is_some() || source_depth.is_some() {
            return Err(invalid("unfinished DDE XML structure"));
        }
        Ok(Self {
            content,
            links,
            sheet_sources,
        })
    }

    #[must_use]
    pub fn source_xml(&self) -> &str {
        &self.content
    }

    #[must_use]
    pub fn links(&self) -> &[Link] {
        &self.links
    }

    #[must_use]
    pub fn sheet_sources(&self) -> &[SheetSource] {
        &self.sheet_sources
    }
}

struct LinkBuilder {
    depth: usize,
    source: Option<Source>,
    cached: Option<Range<usize>>,
}

impl LinkBuilder {
    fn new(depth: usize) -> Self {
        Self {
            depth,
            source: None,
            cached: None,
        }
    }

    fn finish(self, content: Arc<str>) -> Result<Link> {
        Ok(Link {
            source: self
                .source
                .ok_or_else(|| invalid("DDE link has no source"))?,
            content,
            cached_table: self
                .cached
                .ok_or_else(|| invalid("DDE link has no cached table"))?,
        })
    }
}

struct SheetBuilder {
    depth: usize,
    name: String,
    source_seen: bool,
}

fn parse_source(
    element: &BytesStart<'_>,
    reader: &NsReader<&[u8]>,
    text_bytes: usize,
) -> Result<Source> {
    let application = required_attr(element, reader, OFFICE, b"dde-application", text_bytes)?;
    let topic = required_attr(element, reader, OFFICE, b"dde-topic", text_bytes)?;
    let item = required_attr(element, reader, OFFICE, b"dde-item", text_bytes)?;
    let source_name = optional_attr(element, reader, OFFICE, b"name", text_bytes)?;
    let conversion_mode = optional_attr(element, reader, OFFICE, b"conversion-mode", text_bytes)?
        .as_deref()
        .map(ConversionMode::parse)
        .transpose()?
        .unwrap_or(ConversionMode::Unspecified);
    let automatic_update = optional_attr(element, reader, OFFICE, b"automatic-update", text_bytes)?
        .as_deref()
        .map(parse_bool)
        .transpose()?
        .into();
    let mut source = Source::new(application, topic, item)?;
    if let Some(name) = source_name {
        source = source.named(name)?;
    }
    Ok(source
        .with_conversion_mode(conversion_mode)
        .with_automatic_update(automatic_update))
}

fn required_attr(
    element: &BytesStart<'_>,
    reader: &NsReader<&[u8]>,
    namespace: &[u8],
    local: &[u8],
    limit: usize,
) -> Result<String> {
    optional_attr(element, reader, namespace, local, limit)?.ok_or_else(|| {
        invalid(format!(
            "missing required attribute {}",
            String::from_utf8_lossy(local)
        ))
    })
}

fn optional_attr(
    element: &BytesStart<'_>,
    reader: &NsReader<&[u8]>,
    namespace: &[u8],
    local: &[u8],
    limit: usize,
) -> Result<Option<String>> {
    let mut value = None;
    for raw_attribute in element.attributes().with_checks(true) {
        let attribute =
            raw_attribute.map_err(|error| invalid(format!("invalid DDE attribute: {error}")))?;
        let (resolved, name) = reader.resolver().resolve_attribute(attribute.key);
        if matches!(resolved, ResolveResult::Bound(Namespace(uri)) if uri == namespace)
            && name.as_ref() == local
        {
            if value.is_some() {
                return Err(invalid("duplicate DDE attribute"));
            }
            let decoded = attribute
                .decoded_and_normalized_value(XmlVersion::Explicit1_0, reader.decoder())
                .map_err(|error| invalid(format!("invalid DDE attribute value: {error}")))?
                .into_owned();
            if decoded.len() > limit || !xml_text_is_valid(&decoded) {
                return Err(invalid("invalid or oversized DDE attribute value"));
            }
            value = Some(decoded);
        }
    }
    Ok(value)
}

fn parse_bool(value: &str) -> Result<bool> {
    match value {
        "true" | "1" => Ok(true),
        "false" | "0" => Ok(false),
        _ => Err(invalid(format!("invalid XML boolean '{value}'"))),
    }
}

fn xml_text_is_valid(value: &str) -> bool {
    !value.chars().any(|character| {
        matches!(
            character,
            '\u{0000}'..='\u{0008}' | '\u{000B}'..='\u{000C}' | '\u{000E}'..='\u{001F}'
        )
    })
}

fn namespace_kind(namespace: &ResolveResult<'_>) -> NamespaceKind {
    match namespace {
        ResolveResult::Bound(Namespace(uri)) if *uri == OFFICE => NamespaceKind::Office,
        ResolveResult::Bound(Namespace(uri)) if *uri == TABLE => NamespaceKind::Table,
        ResolveResult::Bound(_) | ResolveResult::Unbound | ResolveResult::Unknown(_) => {
            NamespaceKind::Other
        },
    }
}

fn is(namespace: NamespaceKind, local: &[u8], expected_ns: NamespaceKind, expected: &[u8]) -> bool {
    namespace == expected_ns && local == expected
}

fn invalid(message: impl Into<String>) -> Error {
    Error::InvalidStructure(message.into())
}

fn xml_position(reader: &NsReader<&[u8]>) -> Result<usize> {
    usize::try_from(reader.buffer_position()).map_err(|_position_error| Error::PositionOverflow)
}

fn validate_required_source_value(name: &'static str, value: &str) -> Result<()> {
    if value.is_empty() {
        return Err(invalid(format!("DDE attribute {name} must not be empty")));
    }
    if value.len() > MAX_TEXT_BYTES {
        return Err(Error::ResourceLimit {
            resource: name,
            actual: value.len(),
            maximum: MAX_TEXT_BYTES,
        });
    }
    if !xml_text_is_valid(value) {
        return Err(invalid(format!(
            "DDE attribute {name} contains invalid XML character data"
        )));
    }
    Ok(())
}
