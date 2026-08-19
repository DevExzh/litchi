//! Source-aware `content.xml` locations and protection replacements.

use super::model::{Document, Permissions, Sheet, Styles};
use crate::model::protection as wire;
use crate::model::style_protection::{self, PreservedXmlFragment};
use litchi_core::{Error, Result, xml::escape_xml};
use quick_xml::{
    XmlVersion,
    encoding::Decoder,
    events::{BytesStart, Event},
    name::{Namespace, NamespaceResolver, ResolveResult},
    reader::NsReader,
};
use std::ops::Range;

pub(crate) const MAX_CONTENT_BYTES: usize = 64 * 1024 * 1024;
const OFFICE_NAMESPACE: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:office:1.0";
const TABLE_NAMESPACE: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:table:1.0";
const LOEXT_NAMESPACE: &[u8] =
    b"urn:org:documentfoundation:names:experimental:office:xmlns:loext:1.0";
const OFFICE_EXT_NAMESPACE: &[u8] = b"http://openoffice.org/2009/office";

/// Enforce the pre-parse byte-size limit of the protection source.
pub(crate) fn validate_size(source: &str) -> Result<()> {
    if source.len() > MAX_CONTENT_BYTES {
        return Err(Error::InvalidFormat(
            "ODS protection content.xml exceeds the snapshot limit".to_string(),
        ));
    }
    Ok(())
}

#[derive(Clone, Debug)]
pub(crate) struct ElementLocation {
    pub(crate) start: Range<usize>,
    pub(crate) full: Range<usize>,
    pub(crate) close_offset: usize,
    pub(crate) name: String,
    attrs: Vec<AttributeLocation>,
}

#[derive(Clone, Debug)]
struct AttributeLocation {
    range: Range<usize>,
    namespace: Vec<u8>,
    local: Vec<u8>,
}

#[derive(Clone, Debug)]
pub(crate) struct SheetLocation {
    pub(crate) name: String,
    pub(crate) start: ElementLocation,
    pub(crate) end_start: Option<usize>,
    pub(crate) protection: Option<ElementLocation>,
}

/// The exact source context owned by one protection snapshot.
#[derive(Clone, Debug)]
pub(crate) struct Location {
    source_length: usize,
    fingerprint: u64,
    spreadsheet: ElementLocation,
    body_start: Option<usize>,
    sheets: Vec<SheetLocation>,
    automatic: Option<ElementLocation>,
    automatic_fragment: Option<PreservedXmlFragment>,
    automatic_validation_xml: Option<String>,
    styles_xml: Option<String>,
    table_prefix: Option<String>,
    loext_prefix: Option<String>,
}

impl Location {
    /// This standalone path intentionally keeps the historical inline event
    /// loop as the oracle for the fused protection parse ([`super::fused`]),
    /// which drives the equivalent [`LocationHandler`] over the same event
    /// stream.
    #[cfg_attr(not(test), allow(dead_code))]
    pub(crate) fn parse(source: &str, styles_xml: Option<&str>) -> Result<Self> {
        if source.len() > MAX_CONTENT_BYTES {
            return Err(Error::InvalidFormat(
                "ODS protection content.xml exceeds the snapshot limit".to_string(),
            ));
        }

        let mut reader = NsReader::from_str(source);
        let mut buffer = Vec::new();
        let mut stack = Vec::<OpenElement>::new();
        let mut spreadsheet = None;
        let mut body_start = None;
        let mut sheets = Vec::new();
        let mut automatic = None;
        let mut table_prefix = None;
        let mut loext_prefix = None;

        loop {
            let event_start = usize::try_from(reader.buffer_position()).map_err(|_error| {
                Error::InvalidFormat("protection XML position exceeds usize".to_string())
            })?;
            let (namespace, event) = reader
                .read_resolved_event_into(&mut buffer)
                .map_err(|error| Error::InvalidFormat(format!("XML parsing error: {error}")))?;
            let namespace = namespace_uri(&namespace);
            let event = event.into_owned();
            let event_end = usize::try_from(reader.buffer_position()).map_err(|_error| {
                Error::InvalidFormat("protection XML position exceeds usize".to_string())
            })?;
            match event {
                Event::Start(element) => {
                    let info = element_location(source, &reader, &element, event_start, event_end)?;
                    let kind = classify(
                        &namespace,
                        element.local_name().as_ref(),
                        stack.last().map(|open| open.kind),
                    );
                    if matches!(kind, ElementKind::DocumentContent) {
                        collect_prefixes(&reader, &element, &mut table_prefix, &mut loext_prefix)?;
                    }
                    let sheet_index = match kind {
                        ElementKind::Spreadsheet => {
                            if spreadsheet.is_some() {
                                return Err(Error::InvalidFormat(
                                    "duplicate office:spreadsheet element".to_string(),
                                ));
                            }
                            spreadsheet = Some(info.clone());
                            None
                        },
                        ElementKind::Body => {
                            body_start = Some(event_start);
                            None
                        },
                        ElementKind::Sheet => {
                            let name =
                                attribute_value(&reader, &element, TABLE_NAMESPACE, b"name")?
                                    .ok_or_else(|| {
                                        Error::InvalidFormat(
                                            "ODS protected table is missing table:name".to_string(),
                                        )
                                    })?;
                            let index = sheets.len();
                            sheets.push(SheetLocation {
                                name,
                                start: info.clone(),
                                end_start: None,
                                protection: None,
                            });
                            Some(index)
                        },
                        ElementKind::Protection => {
                            let index = stack.iter().rev().find_map(|open| open.sheet_index);
                            if let Some(index) = index {
                                if sheets[index].protection.is_some() {
                                    return Err(Error::InvalidFormat(
                                        "duplicate ODS table-protection element".to_string(),
                                    ));
                                }
                                sheets[index].protection = Some(info.clone());
                            }
                            index
                        },
                        ElementKind::Automatic => {
                            if automatic.is_some() {
                                return Err(Error::InvalidFormat(
                                    "duplicate office:automatic-styles element".to_string(),
                                ));
                            }
                            automatic = Some(info.clone());
                            None
                        },
                        ElementKind::Other | ElementKind::DocumentContent => None,
                    };
                    stack.push(OpenElement {
                        kind,
                        info,
                        sheet_index,
                    });
                },
                Event::Empty(element) => {
                    let info = element_location(source, &reader, &element, event_start, event_end)?;
                    let kind = classify(
                        &namespace,
                        element.local_name().as_ref(),
                        stack.last().map(|open| open.kind),
                    );
                    if matches!(kind, ElementKind::DocumentContent) {
                        collect_prefixes(&reader, &element, &mut table_prefix, &mut loext_prefix)?;
                    }
                    match kind {
                        ElementKind::Spreadsheet => {
                            if spreadsheet.replace(info).is_some() {
                                return Err(Error::InvalidFormat(
                                    "duplicate office:spreadsheet element".to_string(),
                                ));
                            }
                        },
                        ElementKind::Body => body_start = Some(event_start),
                        ElementKind::Sheet => {
                            let name =
                                attribute_value(&reader, &element, TABLE_NAMESPACE, b"name")?
                                    .ok_or_else(|| {
                                        Error::InvalidFormat(
                                            "ODS protected table is missing table:name".to_string(),
                                        )
                                    })?;
                            sheets.push(SheetLocation {
                                name,
                                start: info,
                                end_start: None,
                                protection: None,
                            });
                        },
                        ElementKind::Protection => {
                            if let Some(index) =
                                stack.iter().rev().find_map(|open| open.sheet_index)
                            {
                                if sheets[index].protection.is_some() {
                                    return Err(Error::InvalidFormat(
                                        "duplicate ODS table-protection element".to_string(),
                                    ));
                                }
                                sheets[index].protection = Some(info);
                            }
                        },
                        ElementKind::Automatic if automatic.replace(info).is_some() => {
                            return Err(Error::InvalidFormat(
                                "duplicate office:automatic-styles element".to_string(),
                            ));
                        },
                        ElementKind::Other
                        | ElementKind::DocumentContent
                        | ElementKind::Automatic => {},
                    }
                },
                Event::End(_) => {
                    let open = stack.pop().ok_or_else(|| {
                        Error::InvalidFormat("unexpected ODS XML closing tag".to_string())
                    })?;
                    let full_end = event_end;
                    match open.kind {
                        ElementKind::Sheet => {
                            let index = open.sheet_index.ok_or_else(|| {
                                Error::InvalidFormat(
                                    "ODS protection sheet index is missing".to_string(),
                                )
                            })?;
                            let sheet = sheets.get_mut(index).ok_or_else(|| {
                                Error::InvalidFormat(
                                    "ODS protection sheet index is out of bounds".to_string(),
                                )
                            })?;
                            sheet.start.full = open.info.start.start..full_end;
                            sheet.end_start = Some(event_start);
                        },
                        ElementKind::Protection => {
                            if let Some(index) = open.sheet_index
                                && let Some(protection) = sheets[index].protection.as_mut()
                            {
                                protection.full = protection.start.start..full_end;
                            }
                        },
                        ElementKind::Automatic => {
                            if let Some(value) = automatic.as_mut() {
                                value.full = value.start.start..full_end;
                            }
                        },
                        ElementKind::Other
                        | ElementKind::DocumentContent
                        | ElementKind::Body
                        | ElementKind::Spreadsheet => {},
                    }
                },
                Event::Eof => break,
                Event::Text(_)
                | Event::CData(_)
                | Event::Comment(_)
                | Event::Decl(_)
                | Event::PI(_)
                | Event::DocType(_)
                | Event::GeneralRef(_) => {},
            }
            buffer.clear();
        }

        if !stack.is_empty() {
            return Err(Error::InvalidFormat(
                "unterminated ODS protection XML element".to_string(),
            ));
        }
        let spreadsheet = spreadsheet.ok_or_else(|| {
            Error::InvalidFormat("missing office:spreadsheet element".to_string())
        })?;
        let automatic_fragment = automatic
            .as_ref()
            .map(|_| style_protection::extract_automatic_styles(source))
            .transpose()?
            .flatten();
        let automatic_fragment = automatic_fragment.map(|mut fragment| {
            fragment.xml = validation_fragment(&fragment);
            fragment
        });
        let automatic_validation_xml = automatic_fragment.as_ref().map(validation_fragment);

        Ok(Self {
            source_length: source.len(),
            fingerprint: fingerprint(source.as_bytes()),
            spreadsheet,
            body_start,
            sheets,
            automatic,
            automatic_fragment,
            automatic_validation_xml,
            styles_xml: styles_xml.map(str::to_owned),
            table_prefix,
            loext_prefix,
        })
    }

    pub(crate) fn check_source(&self, source: &str) -> Result<()> {
        if source.len() != self.source_length || fingerprint(source.as_bytes()) != self.fingerprint
        {
            return Err(Error::InvalidFormat(
                "ODS protection source changed before commit".to_string(),
            ));
        }
        Ok(())
    }

    pub(crate) fn sheets(&self) -> &[SheetLocation] {
        &self.sheets
    }

    pub(crate) fn automatic_xml(&self) -> Option<&str> {
        self.automatic_validation_xml.as_deref()
    }

    pub(crate) fn styles_xml(&self) -> Option<&str> {
        self.styles_xml.as_deref()
    }

    pub(crate) fn automatic_fragment(&self) -> Option<&PreservedXmlFragment> {
        self.automatic_fragment.as_ref()
    }

    pub(crate) fn automatic_range(&self) -> Option<Range<usize>> {
        self.automatic
            .as_ref()
            .map(|location| location.full.clone())
    }

    pub(crate) fn body_start(&self) -> Option<usize> {
        self.body_start
    }

    pub(crate) fn sheet_limit(&self) -> usize {
        65_536
    }

    pub(crate) fn has_automatic_owner(&self) -> bool {
        self.automatic.is_some()
    }

    pub(crate) fn table_prefix(&self) -> &str {
        self.table_prefix.as_deref().unwrap_or("table")
    }

    pub(crate) fn loext_prefix(&self) -> Option<&str> {
        self.loext_prefix.as_deref()
    }
}

/// The namespace classification the fused protection parse passes to
/// [`LocationHandler::on_event`].
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub(crate) enum LocationNamespace {
    Office,
    Table,
    Loext,
    OfficeExt,
    Other,
}

impl LocationNamespace {
    /// The URI bytes [`classify`] compares against; `Other` maps to the same
    /// empty slice the historical `namespace_uri` produced for unbound,
    /// unknown, or foreign resolutions.
    fn uri(self) -> &'static [u8] {
        match self {
            Self::Office => OFFICE_NAMESPACE,
            Self::Table => TABLE_NAMESPACE,
            Self::Loext => LOEXT_NAMESPACE,
            Self::OfficeExt => OFFICE_EXT_NAMESPACE,
            Self::Other => b"",
        }
    }
}

/// Classify the resolved event namespace for the source-locator pass.
///
/// The resolved value borrows the reader mutably, so callers classify it
/// immediately after the read exactly as the historical loop body did.
pub(crate) fn location_namespace(namespace: &ResolveResult<'_>) -> LocationNamespace {
    match namespace {
        ResolveResult::Bound(Namespace(uri)) if *uri == OFFICE_NAMESPACE => {
            LocationNamespace::Office
        },
        ResolveResult::Bound(Namespace(uri)) if *uri == TABLE_NAMESPACE => LocationNamespace::Table,
        ResolveResult::Bound(Namespace(uri)) if *uri == LOEXT_NAMESPACE => LocationNamespace::Loext,
        ResolveResult::Bound(Namespace(uri)) if *uri == OFFICE_EXT_NAMESPACE => {
            LocationNamespace::OfficeExt
        },
        ResolveResult::Unbound | ResolveResult::Bound(_) | ResolveResult::Unknown(_) => {
            LocationNamespace::Other
        },
    }
}

/// Streaming event handler holding the [`Location::parse`] scan state.
///
/// The fused protection parse ([`super::fused`]) drives one shared tokenizer
/// through this handler, while [`Location::parse`] keeps its historical
/// inline loop as the standalone oracle; both apply the same checks and
/// error messages at the same events.
#[derive(Debug)]
pub(crate) struct LocationHandler<'a> {
    source: &'a str,
    styles_xml: Option<&'a str>,
    stack: Vec<OpenElement>,
    spreadsheet: Option<ElementLocation>,
    body_start: Option<usize>,
    sheets: Vec<SheetLocation>,
    automatic: Option<ElementLocation>,
    table_prefix: Option<String>,
    loext_prefix: Option<String>,
}

impl<'a> LocationHandler<'a> {
    /// Start the locator scan over `source`, carrying `styles_xml` into the
    /// resulting [`Location`].
    pub(crate) fn new(source: &'a str, styles_xml: Option<&'a str>) -> Self {
        Self {
            source,
            styles_xml,
            stack: Vec::new(),
            spreadsheet: None,
            body_start: None,
            sheets: Vec::new(),
            automatic: None,
            table_prefix: None,
            loext_prefix: None,
        }
    }

    /// Process one resolved event at byte positions `pos_before`/`pos_after`.
    ///
    /// `namespace` is the caller-classified resolution of the event's
    /// namespace; the resolved value borrows the reader mutably, so callers
    /// classify it immediately after the read exactly as the historical loop
    /// body did.
    pub(crate) fn on_event(
        &mut self,
        namespace: LocationNamespace,
        event: &Event<'_>,
        resolver: &NamespaceResolver,
        decoder: Decoder,
        pos_before: u64,
        pos_after: u64,
    ) -> Result<()> {
        let event_start = usize::try_from(pos_before).map_err(|_error| {
            Error::InvalidFormat("protection XML position exceeds usize".to_string())
        })?;
        let event_end = usize::try_from(pos_after).map_err(|_error| {
            Error::InvalidFormat("protection XML position exceeds usize".to_string())
        })?;
        match event {
            Event::Start(element) => {
                let info = element_location_resolved(
                    self.source,
                    resolver,
                    element,
                    event_start,
                    event_end,
                )?;
                let kind = classify(
                    namespace.uri(),
                    element.local_name().as_ref(),
                    self.stack.last().map(|open| open.kind),
                );
                if matches!(kind, ElementKind::DocumentContent) {
                    collect_prefixes_resolved(
                        decoder,
                        element,
                        &mut self.table_prefix,
                        &mut self.loext_prefix,
                    )?;
                }
                let sheet_index = match kind {
                    ElementKind::Spreadsheet => {
                        if self.spreadsheet.is_some() {
                            return Err(Error::InvalidFormat(
                                "duplicate office:spreadsheet element".to_string(),
                            ));
                        }
                        self.spreadsheet = Some(info.clone());
                        None
                    },
                    ElementKind::Body => {
                        self.body_start = Some(event_start);
                        None
                    },
                    ElementKind::Sheet => {
                        let name = attribute_value_resolved(
                            resolver,
                            decoder,
                            element,
                            TABLE_NAMESPACE,
                            b"name",
                        )?
                        .ok_or_else(|| {
                            Error::InvalidFormat(
                                "ODS protected table is missing table:name".to_string(),
                            )
                        })?;
                        let index = self.sheets.len();
                        self.sheets.push(SheetLocation {
                            name,
                            start: info.clone(),
                            end_start: None,
                            protection: None,
                        });
                        Some(index)
                    },
                    ElementKind::Protection => {
                        let index = self.stack.iter().rev().find_map(|open| open.sheet_index);
                        if let Some(index) = index {
                            if self.sheets[index].protection.is_some() {
                                return Err(Error::InvalidFormat(
                                    "duplicate ODS table-protection element".to_string(),
                                ));
                            }
                            self.sheets[index].protection = Some(info.clone());
                        }
                        index
                    },
                    ElementKind::Automatic => {
                        if self.automatic.is_some() {
                            return Err(Error::InvalidFormat(
                                "duplicate office:automatic-styles element".to_string(),
                            ));
                        }
                        self.automatic = Some(info.clone());
                        None
                    },
                    ElementKind::Other | ElementKind::DocumentContent => None,
                };
                self.stack.push(OpenElement {
                    kind,
                    info,
                    sheet_index,
                });
            },
            Event::Empty(element) => {
                let info = element_location_resolved(
                    self.source,
                    resolver,
                    element,
                    event_start,
                    event_end,
                )?;
                let kind = classify(
                    namespace.uri(),
                    element.local_name().as_ref(),
                    self.stack.last().map(|open| open.kind),
                );
                if matches!(kind, ElementKind::DocumentContent) {
                    collect_prefixes_resolved(
                        decoder,
                        element,
                        &mut self.table_prefix,
                        &mut self.loext_prefix,
                    )?;
                }
                match kind {
                    ElementKind::Spreadsheet => {
                        if self.spreadsheet.replace(info).is_some() {
                            return Err(Error::InvalidFormat(
                                "duplicate office:spreadsheet element".to_string(),
                            ));
                        }
                    },
                    ElementKind::Body => self.body_start = Some(event_start),
                    ElementKind::Sheet => {
                        let name = attribute_value_resolved(
                            resolver,
                            decoder,
                            element,
                            TABLE_NAMESPACE,
                            b"name",
                        )?
                        .ok_or_else(|| {
                            Error::InvalidFormat(
                                "ODS protected table is missing table:name".to_string(),
                            )
                        })?;
                        self.sheets.push(SheetLocation {
                            name,
                            start: info,
                            end_start: None,
                            protection: None,
                        });
                    },
                    ElementKind::Protection => {
                        if let Some(index) =
                            self.stack.iter().rev().find_map(|open| open.sheet_index)
                        {
                            if self.sheets[index].protection.is_some() {
                                return Err(Error::InvalidFormat(
                                    "duplicate ODS table-protection element".to_string(),
                                ));
                            }
                            self.sheets[index].protection = Some(info);
                        }
                    },
                    ElementKind::Automatic if self.automatic.replace(info).is_some() => {
                        return Err(Error::InvalidFormat(
                            "duplicate office:automatic-styles element".to_string(),
                        ));
                    },
                    ElementKind::Other | ElementKind::DocumentContent | ElementKind::Automatic => {
                    },
                }
            },
            Event::End(_) => {
                let open = self.stack.pop().ok_or_else(|| {
                    Error::InvalidFormat("unexpected ODS XML closing tag".to_string())
                })?;
                let full_end = event_end;
                match open.kind {
                    ElementKind::Sheet => {
                        let index = open.sheet_index.ok_or_else(|| {
                            Error::InvalidFormat(
                                "ODS protection sheet index is missing".to_string(),
                            )
                        })?;
                        let sheet = self.sheets.get_mut(index).ok_or_else(|| {
                            Error::InvalidFormat(
                                "ODS protection sheet index is out of bounds".to_string(),
                            )
                        })?;
                        sheet.start.full = open.info.start.start..full_end;
                        sheet.end_start = Some(event_start);
                    },
                    ElementKind::Protection => {
                        if let Some(index) = open.sheet_index
                            && let Some(protection) = self.sheets[index].protection.as_mut()
                        {
                            protection.full = protection.start.start..full_end;
                        }
                    },
                    ElementKind::Automatic => {
                        if let Some(value) = self.automatic.as_mut() {
                            value.full = value.start.start..full_end;
                        }
                    },
                    ElementKind::Other
                    | ElementKind::DocumentContent
                    | ElementKind::Body
                    | ElementKind::Spreadsheet => {},
                }
            },
            Event::Eof => {},
            Event::Text(_)
            | Event::CData(_)
            | Event::Comment(_)
            | Event::Decl(_)
            | Event::PI(_)
            | Event::DocType(_)
            | Event::GeneralRef(_) => {},
        }
        Ok(())
    }

    /// Validate the end-of-document state and build the located source
    /// context.
    pub(crate) fn finish(self) -> Result<Location> {
        if !self.stack.is_empty() {
            return Err(Error::InvalidFormat(
                "unterminated ODS protection XML element".to_string(),
            ));
        }
        let spreadsheet = self.spreadsheet.ok_or_else(|| {
            Error::InvalidFormat("missing office:spreadsheet element".to_string())
        })?;
        let automatic_fragment = self
            .automatic
            .as_ref()
            .map(|_| style_protection::extract_automatic_styles(self.source))
            .transpose()?
            .flatten();
        let automatic_fragment = automatic_fragment.map(|mut fragment| {
            fragment.xml = validation_fragment(&fragment);
            fragment
        });
        let automatic_validation_xml = automatic_fragment.as_ref().map(validation_fragment);

        Ok(Location {
            source_length: self.source.len(),
            fingerprint: fingerprint(self.source.as_bytes()),
            spreadsheet,
            body_start: self.body_start,
            sheets: self.sheets,
            automatic: self.automatic,
            automatic_fragment,
            automatic_validation_xml,
            styles_xml: self.styles_xml.map(str::to_owned),
            table_prefix: self.table_prefix,
            loext_prefix: self.loext_prefix,
        })
    }
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
enum ElementKind {
    Other,
    DocumentContent,
    Body,
    Spreadsheet,
    Sheet,
    Protection,
    Automatic,
}

#[derive(Debug)]
struct OpenElement {
    kind: ElementKind,
    info: ElementLocation,
    sheet_index: Option<usize>,
}

fn classify(namespace: &[u8], local: &[u8], parent: Option<ElementKind>) -> ElementKind {
    let office = namespace == OFFICE_NAMESPACE;
    let table = namespace == TABLE_NAMESPACE;
    if office && local == b"document-content" {
        ElementKind::DocumentContent
    } else if office && local == b"body" {
        ElementKind::Body
    } else if office && local == b"spreadsheet" {
        ElementKind::Spreadsheet
    } else if office && local == b"automatic-styles" {
        ElementKind::Automatic
    } else if table && local == b"table" && parent == Some(ElementKind::Spreadsheet) {
        ElementKind::Sheet
    } else if local == b"table-protection"
        && (namespace == TABLE_NAMESPACE
            || namespace == LOEXT_NAMESPACE
            || namespace == OFFICE_EXT_NAMESPACE)
        && parent == Some(ElementKind::Sheet)
    {
        ElementKind::Protection
    } else {
        ElementKind::Other
    }
}

fn element_location(
    source: &str,
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
    start: usize,
    end: usize,
) -> Result<ElementLocation> {
    element_location_resolved(source, reader.resolver(), element, start, end)
}

fn element_location_resolved(
    source: &str,
    resolver: &NamespaceResolver,
    element: &BytesStart<'_>,
    start: usize,
    end: usize,
) -> Result<ElementLocation> {
    let raw = source.get(start..end).ok_or_else(|| {
        Error::InvalidFormat("ODS XML element range is outside content.xml".to_string())
    })?;
    let (ranges, close_offset) = scan_attributes(raw)?;
    let attributes = element
        .attributes()
        .collect::<std::result::Result<Vec<_>, _>>()
        .map_err(|error| Error::InvalidFormat(format!("invalid XML attribute: {error}")))?;
    if ranges.len() != attributes.len() {
        return Err(Error::InvalidFormat(
            "ODS XML attribute scanner disagrees with quick-xml".to_string(),
        ));
    }
    let attrs = attributes
        .into_iter()
        .zip(ranges)
        .map(|(attribute, range)| {
            let (namespace, local) = resolver.resolve_attribute(attribute.key);
            let namespace = match namespace {
                ResolveResult::Bound(Namespace(value)) => value.to_vec(),
                ResolveResult::Unbound | ResolveResult::Unknown(_) => Vec::new(),
            };
            AttributeLocation {
                range,
                namespace,
                local: local.as_ref().to_vec(),
            }
        })
        .collect();
    let name = raw
        .get(1..)
        .and_then(|value| value.split([' ', '\t', '\r', '\n', '>', '/']).next())
        .ok_or_else(|| Error::InvalidFormat("invalid ODS XML element name".to_string()))?
        .to_string();
    Ok(ElementLocation {
        start: start..end,
        full: start..end,
        close_offset,
        name,
        attrs,
    })
}

fn scan_attributes(raw: &str) -> Result<(Vec<Range<usize>>, usize)> {
    let bytes = raw.as_bytes();
    if bytes.first() != Some(&b'<') {
        return Err(Error::InvalidFormat(
            "ODS XML start tag is malformed".to_string(),
        ));
    }
    let mut cursor = 1usize;
    while cursor < bytes.len()
        && !bytes[cursor].is_ascii_whitespace()
        && bytes[cursor] != b'>'
        && bytes[cursor] != b'/'
    {
        cursor += 1;
    }
    let mut ranges = Vec::new();
    loop {
        while cursor < bytes.len() && bytes[cursor].is_ascii_whitespace() {
            cursor += 1;
        }
        if bytes.get(cursor) == Some(&b'>') {
            return Ok((ranges, cursor));
        }
        if bytes.get(cursor) == Some(&b'/') && bytes.get(cursor + 1) == Some(&b'>') {
            return Ok((ranges, cursor));
        }
        let start = cursor;
        while cursor < bytes.len()
            && !bytes[cursor].is_ascii_whitespace()
            && bytes[cursor] != b'='
            && bytes[cursor] != b'>'
        {
            cursor += 1;
        }
        if start == cursor {
            return Err(Error::InvalidFormat(
                "ODS XML attribute name is empty".to_string(),
            ));
        }
        while cursor < bytes.len() && bytes[cursor].is_ascii_whitespace() {
            cursor += 1;
        }
        if bytes.get(cursor) != Some(&b'=') {
            return Err(Error::InvalidFormat(
                "ODS XML attribute is missing '='".to_string(),
            ));
        }
        cursor += 1;
        while cursor < bytes.len() && bytes[cursor].is_ascii_whitespace() {
            cursor += 1;
        }
        let quote = *bytes.get(cursor).ok_or_else(|| {
            Error::InvalidFormat("ODS XML attribute value is truncated".to_string())
        })?;
        if quote != b'\'' && quote != b'"' {
            return Err(Error::InvalidFormat(
                "ODS XML attribute value is unquoted".to_string(),
            ));
        }
        cursor += 1;
        while cursor < bytes.len() && bytes[cursor] != quote {
            cursor += 1;
        }
        if bytes.get(cursor) != Some(&quote) {
            return Err(Error::InvalidFormat(
                "ODS XML attribute value is unterminated".to_string(),
            ));
        }
        cursor += 1;
        ranges.push(start..cursor);
    }
}

fn collect_prefixes(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
    table_prefix: &mut Option<String>,
    loext_prefix: &mut Option<String>,
) -> Result<()> {
    collect_prefixes_resolved(reader.decoder(), element, table_prefix, loext_prefix)
}

fn collect_prefixes_resolved(
    decoder: Decoder,
    element: &BytesStart<'_>,
    table_prefix: &mut Option<String>,
    loext_prefix: &mut Option<String>,
) -> Result<()> {
    for attribute in element.attributes() {
        let attribute = attribute
            .map_err(|error| Error::InvalidFormat(format!("invalid XML attribute: {error}")))?;
        let key = attribute.key.as_ref();
        let prefix = if key == b"xmlns" {
            String::new()
        } else if let Some(value) = key.strip_prefix(b"xmlns:") {
            String::from_utf8(value.to_vec()).map_err(|_error| {
                Error::InvalidFormat("invalid XML namespace prefix".to_string())
            })?
        } else {
            continue;
        };
        let value = attribute
            .decoded_and_normalized_value(XmlVersion::Implicit1_0, decoder)
            .map_err(|error| Error::InvalidFormat(format!("invalid XML namespace: {error}")))?;
        if value.as_bytes() == TABLE_NAMESPACE && table_prefix.is_none() {
            *table_prefix = Some(prefix.clone());
        }
        if value.as_bytes() == LOEXT_NAMESPACE && loext_prefix.is_none() {
            *loext_prefix = Some(prefix);
        }
    }
    Ok(())
}

fn attribute_value(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
    namespace: &[u8],
    local_name: &[u8],
) -> Result<Option<String>> {
    attribute_value_resolved(
        reader.resolver(),
        reader.decoder(),
        element,
        namespace,
        local_name,
    )
}

fn attribute_value_resolved(
    resolver: &NamespaceResolver,
    decoder: Decoder,
    element: &BytesStart<'_>,
    namespace: &[u8],
    local_name: &[u8],
) -> Result<Option<String>> {
    for attribute in element.attributes() {
        let attribute = attribute
            .map_err(|error| Error::InvalidFormat(format!("invalid XML attribute: {error}")))?;
        let (resolved, local) = resolver.resolve_attribute(attribute.key);
        if is_namespace(&resolved, namespace) && local.as_ref() == local_name {
            return attribute
                .decoded_and_normalized_value(XmlVersion::Implicit1_0, decoder)
                .map(|value| Some(value.into_owned()))
                .map_err(|error| Error::InvalidFormat(format!("invalid XML attribute: {error}")));
        }
    }
    Ok(None)
}

fn is_namespace(namespace: &ResolveResult<'_>, expected: &[u8]) -> bool {
    matches!(namespace, ResolveResult::Bound(Namespace(value)) if *value == expected)
}

fn namespace_uri(namespace: &ResolveResult<'_>) -> Vec<u8> {
    match namespace {
        ResolveResult::Bound(Namespace(value)) => value.to_vec(),
        ResolveResult::Unbound | ResolveResult::Unknown(_) => Vec::new(),
    }
}

fn validation_fragment(fragment: &PreservedXmlFragment) -> String {
    let mut xml = fragment.xml.clone();
    let root_end = xml.find('>').unwrap_or(xml.len());
    let declared = fragment
        .namespace_prefixes()
        .filter(|prefix| {
            let attribute = if prefix.is_empty() {
                "xmlns=".to_string()
            } else {
                format!("xmlns:{prefix}=")
            };
            xml[..root_end].contains(&attribute)
        })
        .collect::<Vec<_>>();
    let mut declarations = String::new();
    fragment.write_missing_namespaces(&mut declarations, declared);
    xml.insert_str(root_end, &declarations);
    xml
}

fn fingerprint(bytes: &[u8]) -> u64 {
    let mut value = 0xcbf29ce484222325u64;
    for byte in bytes {
        value ^= u64::from(*byte);
        value = value.wrapping_mul(0x100000001b3);
    }
    value
}

pub(crate) fn replace(
    source: &str,
    location: &Location,
    document: &Document,
    sheets: &[Sheet],
    styles: &Styles,
    rewrite_styles: bool,
) -> Result<String> {
    location.check_source(source)?;
    if sheets.len() != location.sheets.len() {
        return Err(Error::InvalidFormat(
            "ODS protection sheet count changed outside the worksheet owner".to_string(),
        ));
    }

    let mut replacements = Vec::<Replacement>::new();
    let document_additions = document_attributes(document, location);
    let document_namespace = namespace_declaration(location, &document_additions);
    let spreadsheet = rewrite_start_tag(
        source,
        &location.spreadsheet,
        &[
            (TABLE_NAMESPACE, b"structure-protected".as_slice()),
            (TABLE_NAMESPACE, b"protection-key".as_slice()),
            (
                TABLE_NAMESPACE,
                b"protection-key-digest-algorithm".as_slice(),
            ),
            (
                LOEXT_NAMESPACE,
                b"protection-key-digest-algorithm-2".as_slice(),
            ),
        ],
        document_additions,
        &document_namespace,
    )?;
    replacements.push(Replacement::range(
        location.spreadsheet.start.clone(),
        spreadsheet,
    ));

    for (location_sheet, sheet) in location.sheets.iter().zip(sheets) {
        let sheet_additions = sheet_attributes(sheet, location);
        let sheet_namespace = namespace_declaration(location, &sheet_additions);
        let replacement = rewrite_start_tag(
            source,
            &location_sheet.start,
            &[
                (TABLE_NAMESPACE, b"protected".as_slice()),
                (TABLE_NAMESPACE, b"protection-key".as_slice()),
                (
                    TABLE_NAMESPACE,
                    b"protection-key-digest-algorithm".as_slice(),
                ),
                (
                    LOEXT_NAMESPACE,
                    b"protection-key-digest-algorithm-2".as_slice(),
                ),
            ],
            sheet_additions,
            &sheet_namespace,
        )?;
        if replacement != source[location_sheet.start.start.clone()] {
            replacements.push(Replacement::range(
                location_sheet.start.start.clone(),
                replacement,
            ));
        }

        let protection_additions = options_attributes(&sheet.permissions, location);
        if let Some(protection) = &location_sheet.protection {
            let namespace_declaration =
                namespace_declaration(location, protection_additions.as_slice());
            let start = rewrite_start_tag(
                source,
                protection,
                &[
                    (TABLE_NAMESPACE, b"select-protected-cells".as_slice()),
                    (TABLE_NAMESPACE, b"select-unprotected-cells".as_slice()),
                    (TABLE_NAMESPACE, b"insert-columns".as_slice()),
                    (TABLE_NAMESPACE, b"insert-rows".as_slice()),
                    (TABLE_NAMESPACE, b"delete-columns".as_slice()),
                    (TABLE_NAMESPACE, b"delete-rows".as_slice()),
                    (TABLE_NAMESPACE, b"use-autofilter".as_slice()),
                    (TABLE_NAMESPACE, b"use-pivot".as_slice()),
                    (LOEXT_NAMESPACE, b"select-protected-cells".as_slice()),
                    (LOEXT_NAMESPACE, b"select-unprotected-cells".as_slice()),
                    (LOEXT_NAMESPACE, b"insert-columns".as_slice()),
                    (LOEXT_NAMESPACE, b"insert-rows".as_slice()),
                    (LOEXT_NAMESPACE, b"delete-columns".as_slice()),
                    (LOEXT_NAMESPACE, b"delete-rows".as_slice()),
                    (LOEXT_NAMESPACE, b"use-autofilter".as_slice()),
                    (LOEXT_NAMESPACE, b"use-pivot".as_slice()),
                    (OFFICE_EXT_NAMESPACE, b"select-protected-cells".as_slice()),
                    (OFFICE_EXT_NAMESPACE, b"select-unprotected-cells".as_slice()),
                    (OFFICE_EXT_NAMESPACE, b"insert-columns".as_slice()),
                    (OFFICE_EXT_NAMESPACE, b"insert-rows".as_slice()),
                    (OFFICE_EXT_NAMESPACE, b"delete-columns".as_slice()),
                    (OFFICE_EXT_NAMESPACE, b"delete-rows".as_slice()),
                    (OFFICE_EXT_NAMESPACE, b"use-autofilter".as_slice()),
                    (OFFICE_EXT_NAMESPACE, b"use-pivot".as_slice()),
                ],
                protection_additions,
                &namespace_declaration,
            )?;
            let original = &source[protection.full.clone()];
            let replacement = if protection.full == protection.start {
                start
            } else {
                let tail_start = protection.start.end - protection.start.start;
                format!("{}{}", start, &original[tail_start..])
            };
            if replacement != original {
                replacements.push(Replacement::range(protection.full.clone(), replacement));
            }
        } else if !protection_additions.is_empty() {
            let options = render_options(location, &protection_additions)?;
            if let Some(end_start) = location_sheet.end_start {
                replacements.push(Replacement::insertion(end_start, options));
            } else {
                let original = &source[location_sheet.start.full.clone()];
                let close = location_sheet.start.close_offset;
                let expanded = format!(
                    "{}>{}</{}>",
                    &original[..close],
                    options,
                    location_sheet
                        .start
                        .name
                        .trim_start_matches('<')
                        .trim_end_matches('>')
                );
                replacements.push(Replacement::range(
                    location_sheet.start.full.clone(),
                    expanded,
                ));
            }
        }
    }

    if rewrite_styles {
        let fragment = style_protection::rewrite_managed_cell_styles(
            location.automatic_fragment(),
            styles.conditional(),
            &styles_to_wire(styles),
        )?;
        if let Some(range) = location.automatic_range() {
            replacements.push(Replacement::range(range, fragment.xml));
        } else if !styles.is_empty() {
            let insertion = location.body_start().ok_or_else(|| {
                Error::InvalidFormat(
                    "ODS protection cannot insert automatic styles without office:body".to_string(),
                )
            })?;
            replacements.push(Replacement::insertion(insertion, fragment.xml));
        }
    }

    replacements.sort_by_key(|replacement| (replacement.start(), replacement.end()));
    for pair in replacements.windows(2) {
        if pair[0].end() > pair[1].start() {
            return Err(Error::InvalidFormat(
                "ODS protection replacements overlap".to_string(),
            ));
        }
    }
    let mut output = String::with_capacity(source.len());
    let mut cursor = 0usize;
    for replacement in replacements {
        output.push_str(&source[cursor..replacement.start()]);
        output.push_str(&replacement.value);
        cursor = replacement.end();
    }
    output.push_str(&source[cursor..]);
    Ok(output)
}

#[derive(Clone, Debug)]
struct Replacement {
    range: Range<usize>,
    value: String,
}

impl Replacement {
    fn range(range: Range<usize>, value: String) -> Self {
        Self { range, value }
    }

    fn insertion(offset: usize, value: String) -> Self {
        Self {
            range: offset..offset,
            value,
        }
    }

    fn start(&self) -> usize {
        self.range.start
    }

    fn end(&self) -> usize {
        self.range.end
    }
}

fn document_attributes(value: &Document, location: &Location) -> Vec<AttributeAddition> {
    let mut additions = Vec::new();
    let prefix = location.table_prefix().to_string();
    add_bool(
        &mut additions,
        prefix.clone(),
        "structure-protected",
        value.structure_protected,
    );
    add_key(&mut additions, prefix, &value.key, location);
    additions
}

fn sheet_attributes(value: &Sheet, location: &Location) -> Vec<AttributeAddition> {
    let mut additions = Vec::new();
    let prefix = location.table_prefix().to_string();
    add_bool(&mut additions, prefix.clone(), "protected", value.protected);
    add_key(&mut additions, prefix, &value.key, location);
    additions
}

fn add_key(
    additions: &mut Vec<AttributeAddition>,
    table_prefix: String,
    key: &wire::Key,
    location: &Location,
) {
    add_text(
        additions,
        table_prefix.clone(),
        "protection-key",
        key.value.as_deref(),
    );
    add_text(
        additions,
        table_prefix.clone(),
        "protection-key-digest-algorithm",
        key.digest_algorithm.as_deref(),
    );
    add_text(
        additions,
        location.loext_prefix().unwrap_or("loext").to_string(),
        "protection-key-digest-algorithm-2",
        key.secondary_digest_algorithm.as_deref(),
    );
}

fn add_bool(
    additions: &mut Vec<AttributeAddition>,
    prefix: String,
    name: &str,
    value: Option<bool>,
) {
    add_text(
        additions,
        prefix,
        name,
        value.map(|value| if value { "true" } else { "false" }),
    );
}

fn add_text(
    additions: &mut Vec<AttributeAddition>,
    prefix: String,
    name: &str,
    value: Option<&str>,
) {
    if let Some(value) = value {
        additions.push(AttributeAddition {
            prefix,
            name: name.to_string(),
            value: value.to_string(),
        });
    }
}

#[derive(Clone, Debug)]
struct AttributeAddition {
    prefix: String,
    name: String,
    value: String,
}

fn options_attributes(value: &Permissions, location: &Location) -> Vec<AttributeAddition> {
    let prefix = location.loext_prefix().unwrap_or("loext").to_string();
    let mut additions = Vec::new();
    for (name, option) in [
        ("select-protected-cells", value.select_protected_cells),
        ("select-unprotected-cells", value.select_unprotected_cells),
        ("insert-columns", value.insert_columns),
        ("insert-rows", value.insert_rows),
        ("delete-columns", value.delete_columns),
        ("delete-rows", value.delete_rows),
        ("use-autofilter", value.use_auto_filter),
        ("use-pivot", value.use_pivot),
    ] {
        add_bool(&mut additions, prefix.clone(), name, option);
    }
    additions
}

fn namespace_declaration(location: &Location, additions: &[AttributeAddition]) -> String {
    let Some(attribute) = additions
        .iter()
        .find(|attribute| attribute.prefix == "loext")
    else {
        return String::new();
    };
    if location.loext_prefix().is_some() {
        return String::new();
    }
    format!(
        " xmlns:{}=\"{}\"",
        attribute.prefix,
        escape_xml(&String::from_utf8_lossy(LOEXT_NAMESPACE))
    )
}

fn render_options(location: &Location, additions: &[AttributeAddition]) -> Result<String> {
    if additions.is_empty() {
        return Ok(String::new());
    }
    let prefix = additions[0].prefix.as_str();
    let mut out = format!("<{prefix}:table-protection");
    out.push_str(&namespace_declaration(location, additions));
    for addition in additions {
        out.push(' ');
        out.push_str(&addition.prefix);
        out.push(':');
        out.push_str(&addition.name);
        out.push_str("=\"");
        out.push_str(&escape_xml(&addition.value));
        out.push('"');
    }
    out.push_str("/>");
    Ok(out)
}

fn rewrite_start_tag(
    source: &str,
    element: &ElementLocation,
    removals: &[(&[u8], &[u8])],
    additions: Vec<AttributeAddition>,
    namespace_declaration: &str,
) -> Result<String> {
    let raw = source.get(element.start.clone()).ok_or_else(|| {
        Error::InvalidFormat("ODS protection element range is outside content.xml".to_string())
    })?;
    let mut out = String::with_capacity(raw.len() + additions.len() * 32);
    let mut cursor = 0usize;
    for attribute in &element.attrs {
        if removals.iter().any(|(namespace, local)| {
            attribute.namespace == *namespace && attribute.local == *local
        }) {
            out.push_str(&raw[cursor..attribute.range.start]);
            cursor = attribute.range.end;
        }
    }
    out.push_str(&raw[cursor..element.close_offset]);
    out.push_str(namespace_declaration);
    for addition in additions {
        out.push(' ');
        out.push_str(&addition.prefix);
        out.push(':');
        out.push_str(&addition.name);
        out.push_str("=\"");
        out.push_str(&escape_xml(&addition.value));
        out.push('"');
    }
    out.push_str(&raw[element.close_offset..]);
    Ok(out)
}

fn styles_to_wire(value: &Styles) -> Vec<style_protection::TableStyle> {
    value
        .automatic()
        .iter()
        .map(|style| {
            let mut wire = style_protection::TableStyle::new(style.name.clone(), style.protection);
            if let Some(parent) = &style.parent_name {
                wire = wire.with_parent_style_name(parent.clone());
            }
            wire
        })
        .collect()
}

pub(crate) fn parse(
    source: &str,
    styles_xml: Option<&str>,
) -> Result<(Location, Document, Vec<Sheet>, Styles)> {
    let (location, document, wire_sheets) = super::fused::parse(source, styles_xml)?;
    if wire_sheets.len() != location.sheets.len() {
        return Err(Error::InvalidFormat(
            "ODS protection sheet parser and source locator disagree".to_string(),
        ));
    }
    let sheets = location
        .sheets
        .iter()
        .zip(wire_sheets)
        .map(|(location, sheet)| super::model::sheet_from_wire(location.name.clone(), &sheet))
        .collect::<Vec<_>>();
    let automatic_xml = location
        .automatic_fragment
        .as_ref()
        .map_or("", |fragment| fragment.xml.as_str());
    let registry = style_protection::CellStyleRegistry::parse(styles_xml, automatic_xml)?;
    let styles = super::model::styles_from_wire(&registry);
    Ok((
        location,
        super::model::document_from_wire(&document),
        sheets,
        styles,
    ))
}
