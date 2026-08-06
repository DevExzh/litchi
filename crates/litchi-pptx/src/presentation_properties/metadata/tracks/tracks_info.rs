//! Lossless XML discovery for `[MS-PPTX]` media track metadata.
//!
//! This codec deliberately discovers only the typed `p173:tracksInfo` and
//! `p15:isNarration` islands.  Every other child of the media/shape extension
//! lists remains opaque and is retained by range-based edits in `transaction`.

use std::ops::Range;

use quick_xml::XmlVersion;
use quick_xml::encoding::Decoder;
use quick_xml::events::{BytesStart, Event};
use quick_xml::name::{Namespace, ResolveResult};
use quick_xml::reader::NsReader;

use super::model::MediaKey;
use crate::{Error, Result};

pub(crate) const PML: &[u8] = b"http://schemas.openxmlformats.org/presentationml/2006/main";
pub(crate) const PML_STRICT: &[u8] = b"http://purl.oclc.org/ooxml/presentationml/main";
pub(crate) const P14: &[u8] = b"http://schemas.microsoft.com/office/powerpoint/2010/main";
pub(crate) const P15: &[u8] = b"http://schemas.microsoft.com/office/powerpoint/2012/main";
pub(crate) const P173: &[u8] = b"http://schemas.microsoft.com/office/powerpoint/2017/3/main";
pub const TRACKS_INFO_URI: &str = "{3AFAAA56-56D3-431D-BCD4-E75A35582382}";
pub const NARRATION_URI: &str = "{42D2F446-02D8-4167-A562-619A0277C38B}";

const MAX_SLIDE_BYTES: usize = 64 * 1024 * 1024;
const MAX_TRACKS: usize = 4_096;
const MAX_STRING_BYTES: usize = 1024 * 1024;

#[derive(Debug, Clone)]
pub(crate) struct Attr {
    pub(crate) value: String,
    pub(crate) span: Range<usize>,
    pub(crate) full_span: Range<usize>,
}

#[derive(Debug, Clone)]
pub(crate) struct Track {
    pub(crate) id: Attr,
    pub(crate) label: Attr,
    pub(crate) language: Option<Attr>,
    pub(crate) embed: Option<Attr>,
    pub(crate) link: Option<Attr>,
    pub(crate) opening_insert: usize,
}

#[derive(Debug, Clone)]
pub(crate) struct TracksInfo {
    pub(crate) display_location: Attr,
    pub(crate) tracks: Vec<Track>,
}

#[derive(Debug, Clone)]
pub(crate) struct Narration {
    pub(crate) value: Option<Attr>,
    pub(crate) opening_insert: usize,
}

#[derive(Debug, Clone)]
pub(crate) struct Media {
    pub(crate) embed: Option<Attr>,
    pub(crate) link: Option<Attr>,
    pub(crate) tracks_info: Option<TracksInfo>,
}

#[derive(Debug, Clone)]
pub(crate) struct Found {
    pub(crate) key: MediaKey,
    pub(crate) media: Media,
    pub(crate) narration: Option<Narration>,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum NamespaceKind {
    Pml,
    P14,
    P15,
    P173,
    Other,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum Kind {
    Shape,
    Media,
    TracksInfo,
    Track,
    Narration,
    Other,
}

#[derive(Debug)]
struct Open {
    kind: Kind,
    track: Option<Track>,
}

#[derive(Debug)]
struct Shape {
    id: Option<u32>,
    media: Option<Media>,
    narration: Option<Narration>,
}

pub(crate) fn discover(xml: &[u8], key: &MediaKey) -> Result<Option<Found>> {
    if xml.is_empty() {
        return Err(invalid("slide XML is empty"));
    }
    if xml.len() > MAX_SLIDE_BYTES {
        return Err(Error::Limit {
            resource: "slide XML bytes for media tracks",
            limit: MAX_SLIDE_BYTES,
        });
    }

    let mut reader = NsReader::from_reader(xml);
    reader.config_mut().trim_text(false);
    let mut stack = Vec::new();
    let mut shape = None;
    let mut found = None;
    let mut root_seen = false;
    let mut root_closed = false;
    let mut nodes = 0usize;

    loop {
        let start = position(&reader)?;
        let decoder = reader.decoder();
        let (namespace, event) = reader.read_resolved_event().map_err(xml_error)?;
        let namespace = namespace_kind(namespace);
        let end = position(&reader)?;
        let event = event.into_owned();

        match event {
            Event::Start(element) => {
                nodes = nodes
                    .checked_add(1)
                    .ok_or_else(|| invalid("slide XML node count overflow"))?;
                if nodes > 1_000_000 {
                    return Err(Error::Limit {
                        resource: "slide XML nodes for media tracks",
                        limit: 1_000_000,
                    });
                }
                let kind = classify(
                    namespace,
                    element.local_name().as_ref(),
                    &stack,
                    shape.as_ref(),
                );
                if !root_seen {
                    root_seen = true;
                }
                match kind {
                    Kind::Shape => {
                        if shape.is_some() {
                            return Err(invalid(
                                "nested PresentationML media shapes are unsupported",
                            ));
                        }
                        shape = Some(Shape {
                            id: None,
                            media: None,
                            narration: None,
                        });
                    },
                    Kind::Media => {
                        let current = shape
                            .as_mut()
                            .ok_or_else(|| invalid("media extension is outside a shape"))?;
                        if current.media.is_some() {
                            return Err(invalid(
                                "media shape contains duplicate p14:media elements",
                            ));
                        }
                        current.media = Some(Media {
                            embed: attr(&element, b"embed", Some(b"r"), decoder, start)?,
                            link: attr(&element, b"link", Some(b"r"), decoder, start)?,
                            tracks_info: None,
                        });
                    },
                    Kind::TracksInfo => {
                        let current = shape
                            .as_mut()
                            .and_then(|value| value.media.as_mut())
                            .ok_or_else(|| invalid("tracksInfo is outside p14:media"))?;
                        if current.tracks_info.is_some() {
                            return Err(invalid("media contains duplicate tracksInfo elements"));
                        }
                        current.tracks_info = Some(TracksInfo {
                            display_location: required_attr(
                                &element,
                                b"displayLoc",
                                None,
                                decoder,
                                start,
                                "tracksInfo displayLoc",
                            )?,
                            tracks: Vec::new(),
                        });
                    },
                    Kind::Track => {
                        let current = shape
                            .as_mut()
                            .and_then(|value| value.media.as_mut())
                            .and_then(|value| value.tracks_info.as_mut())
                            .ok_or_else(|| invalid("track is outside tracksInfo"))?;
                        if current.tracks.len() >= MAX_TRACKS {
                            return Err(Error::Limit {
                                resource: "media caption track count",
                                limit: MAX_TRACKS,
                            });
                        }
                        let track = Track {
                            id: required_attr(&element, b"id", None, decoder, start, "track id")?,
                            label: required_attr(
                                &element,
                                b"label",
                                None,
                                decoder,
                                start,
                                "track label",
                            )?,
                            language: attr(&element, b"lang", None, decoder, start)?,
                            embed: attr(&element, b"embed", Some(b"r"), decoder, start)?,
                            link: attr(&element, b"link", Some(b"r"), decoder, start)?,
                            opening_insert: opening_insert(&element, start)?,
                        };
                        if track.embed.is_none() && track.link.is_none() {
                            return Err(invalid("track requires r:embed or r:link"));
                        }
                        let open = Open {
                            kind,
                            track: Some(track),
                        };
                        stack.push(open);
                        continue;
                    },
                    Kind::Narration => {
                        let current = shape
                            .as_mut()
                            .ok_or_else(|| invalid("isNarration is outside a shape"))?;
                        if current.narration.is_some() {
                            return Err(invalid("shape contains duplicate isNarration elements"));
                        }
                        current.narration = Some(Narration {
                            value: attr(&element, b"val", None, decoder, start)?,
                            opening_insert: opening_insert(&element, start)?,
                        });
                    },
                    Kind::Other => {},
                }
                if namespace == NamespaceKind::Pml
                    && element.local_name().as_ref() == b"cNvPr"
                    && shape.as_ref().is_some_and(|value| value.id.is_none())
                {
                    let id = parse_shape_id(&element, decoder, start)?;
                    if let Some(current) = shape.as_mut() {
                        current.id = id;
                    }
                }
                stack.push(Open { kind, track: None });
            },
            Event::Empty(element) => {
                nodes = nodes
                    .checked_add(1)
                    .ok_or_else(|| invalid("slide XML node count overflow"))?;
                if nodes > 1_000_000 {
                    return Err(Error::Limit {
                        resource: "slide XML nodes for media tracks",
                        limit: 1_000_000,
                    });
                }
                let kind = classify(
                    namespace,
                    element.local_name().as_ref(),
                    &stack,
                    shape.as_ref(),
                );
                root_seen = true;
                match kind {
                    Kind::Shape => {
                        let candidate = Shape {
                            id: parse_shape_id(&element, decoder, start)?,
                            media: None,
                            narration: None,
                        };
                        if candidate.id == Some(key.shape_id) {
                            return Err(invalid("selected media shape has no p14:media element"));
                        }
                    },
                    Kind::Media => {
                        let current = shape
                            .as_mut()
                            .ok_or_else(|| invalid("media extension is outside a shape"))?;
                        if current.media.is_some() {
                            return Err(invalid(
                                "media shape contains duplicate p14:media elements",
                            ));
                        }
                        current.media = Some(Media {
                            embed: attr(&element, b"embed", Some(b"r"), decoder, start)?,
                            link: attr(&element, b"link", Some(b"r"), decoder, start)?,
                            tracks_info: None,
                        });
                    },
                    Kind::TracksInfo => {
                        let current = shape
                            .as_mut()
                            .and_then(|value| value.media.as_mut())
                            .ok_or_else(|| invalid("tracksInfo is outside p14:media"))?;
                        if current.tracks_info.is_some() {
                            return Err(invalid("media contains duplicate tracksInfo elements"));
                        }
                        current.tracks_info = Some(TracksInfo {
                            display_location: required_attr(
                                &element,
                                b"displayLoc",
                                None,
                                decoder,
                                start,
                                "tracksInfo displayLoc",
                            )?,
                            tracks: Vec::new(),
                        });
                    },
                    Kind::Track => {
                        let current = shape
                            .as_mut()
                            .and_then(|value| value.media.as_mut())
                            .and_then(|value| value.tracks_info.as_mut())
                            .ok_or_else(|| invalid("track is outside tracksInfo"))?;
                        if current.tracks.len() >= MAX_TRACKS {
                            return Err(Error::Limit {
                                resource: "media caption track count",
                                limit: MAX_TRACKS,
                            });
                        }
                        let track = parse_track(&element, decoder, start)?;
                        current.tracks.push(track);
                    },
                    Kind::Narration => {
                        let current = shape
                            .as_mut()
                            .ok_or_else(|| invalid("isNarration is outside a shape"))?;
                        if current.narration.is_some() {
                            return Err(invalid("shape contains duplicate isNarration elements"));
                        }
                        current.narration = Some(Narration {
                            value: attr(&element, b"val", None, decoder, start)?,
                            opening_insert: opening_insert(&element, start)?,
                        });
                    },
                    Kind::Other => {},
                }
                if namespace == NamespaceKind::Pml
                    && element.local_name().as_ref() == b"cNvPr"
                    && shape.as_ref().is_some_and(|value| value.id.is_none())
                {
                    let id = parse_shape_id(&element, decoder, start)?;
                    if let Some(current) = shape.as_mut() {
                        current.id = id;
                    }
                }
            },
            Event::End(element) => {
                let open = stack
                    .pop()
                    .ok_or_else(|| invalid("slide XML element stack underflow"))?;
                if element.local_name().as_ref().is_empty() {
                    return Err(invalid("slide XML element has an empty name"));
                }
                match open.kind {
                    Kind::Track => {
                        let current = shape
                            .as_mut()
                            .and_then(|value| value.media.as_mut())
                            .and_then(|value| value.tracks_info.as_mut())
                            .ok_or_else(|| invalid("track closed outside tracksInfo"))?;
                        if let Some(track) = open.track {
                            current.tracks.push(track);
                        }
                    },
                    Kind::Shape => {
                        let candidate = shape
                            .take()
                            .ok_or_else(|| invalid("shape state disappeared before close"))?;
                        if candidate.id == Some(key.shape_id) {
                            if found.is_some() {
                                return Err(invalid(
                                    "multiple media shapes have the selected shape ID",
                                ));
                            }
                            let media = candidate.media.ok_or_else(|| {
                                invalid("selected shape has no p14:media element")
                            })?;
                            found = Some(Found {
                                key: key.clone(),
                                media,
                                narration: candidate.narration,
                            });
                        }
                    },
                    Kind::Media | Kind::TracksInfo | Kind::Narration | Kind::Other => {},
                }
                if stack.is_empty() {
                    root_closed = true;
                }
            },
            Event::Text(text) => {
                let value = text
                    .decode()
                    .map_err(|error| Error::Xml(error.to_string()))?;
                if !value.trim().is_empty()
                    && stack.iter().any(|open| {
                        matches!(open.kind, Kind::TracksInfo | Kind::Track | Kind::Narration)
                    })
                {
                    return Err(invalid("media track metadata cannot contain text"));
                }
            },
            Event::CData(text) => {
                if !text
                    .decode()
                    .map_err(|error| Error::Xml(error.to_string()))?
                    .trim()
                    .is_empty()
                    && stack.iter().any(|open| {
                        matches!(open.kind, Kind::TracksInfo | Kind::Track | Kind::Narration)
                    })
                {
                    return Err(invalid("media track metadata cannot contain CDATA"));
                }
            },
            Event::DocType(_) | Event::PI(_) | Event::GeneralRef(_) => {
                return Err(invalid("unsupported XML construct in slide media metadata"));
            },
            Event::Comment(_) | Event::Decl(_) => {},
            Event::Eof => break,
        }

        if end > xml.len() {
            return Err(invalid("slide XML event range is outside its source"));
        }
    }

    if !root_seen || !root_closed || !stack.is_empty() {
        return Err(invalid("slide XML is not one complete document"));
    }
    Ok(found)
}

fn classify(namespace: NamespaceKind, local: &[u8], stack: &[Open], shape: Option<&Shape>) -> Kind {
    if namespace == NamespaceKind::Pml && local == b"pic" && shape.is_none() {
        return Kind::Shape;
    }
    if namespace == NamespaceKind::P14 && local == b"media" && shape.is_some() {
        return Kind::Media;
    }
    if namespace == NamespaceKind::P173
        && local == b"tracksInfo"
        && stack.iter().any(|open| open.kind == Kind::Media)
    {
        return Kind::TracksInfo;
    }
    if namespace == NamespaceKind::P173
        && local == b"track"
        && stack.iter().any(|open| open.kind == Kind::TracksInfo)
    {
        return Kind::Track;
    }
    if namespace == NamespaceKind::P15 && local == b"isNarration" && shape.is_some() {
        return Kind::Narration;
    }
    Kind::Other
}

fn parse_track(element: &BytesStart<'_>, decoder: Decoder, start: usize) -> Result<Track> {
    let track = Track {
        id: required_attr(element, b"id", None, decoder, start, "track id")?,
        label: required_attr(element, b"label", None, decoder, start, "track label")?,
        language: attr(element, b"lang", None, decoder, start)?,
        embed: attr(element, b"embed", Some(b"r"), decoder, start)?,
        link: attr(element, b"link", Some(b"r"), decoder, start)?,
        opening_insert: opening_insert(element, start)?,
    };
    if track.embed.is_none() && track.link.is_none() {
        return Err(invalid("track requires r:embed or r:link"));
    }
    Ok(track)
}

fn parse_shape_id(element: &BytesStart<'_>, decoder: Decoder, start: usize) -> Result<Option<u32>> {
    Ok(attr(element, b"id", None, decoder, start)?
        .map(|value| {
            value
                .value
                .parse::<u32>()
                .map_err(|_| invalid("media shape cNvPr/@id is not an unsigned integer"))
        })
        .transpose()?)
}

fn required_attr(
    element: &BytesStart<'_>,
    local: &[u8],
    prefix: Option<&[u8]>,
    decoder: Decoder,
    start: usize,
    label: &str,
) -> Result<Attr> {
    attr(element, local, prefix, decoder, start)?
        .ok_or_else(|| invalid(format!("{label} is missing")))
}

fn attr(
    element: &BytesStart<'_>,
    local: &[u8],
    prefix: Option<&[u8]>,
    decoder: Decoder,
    start: usize,
) -> Result<Option<Attr>> {
    let mut result = None;
    for attribute in element.attributes() {
        let attribute = attribute.map_err(|error| Error::Xml(error.to_string()))?;
        let raw_key = attribute.key.as_ref();
        let colon = raw_key.iter().position(|byte| *byte == b':');
        let actual_prefix = colon.map(|offset| &raw_key[..offset]);
        let actual_local = colon.map_or(raw_key, |offset| &raw_key[offset + 1..]);
        if actual_local != local || actual_prefix != prefix {
            continue;
        }
        if result.is_some() {
            return Err(invalid(format!(
                "duplicate media track attribute '{}'",
                String::from_utf8_lossy(local)
            )));
        }
        let value = attribute
            .decoded_and_normalized_value(XmlVersion::Explicit1_0, decoder)
            .map_err(xml_error)?
            .into_owned();
        if value.len() > MAX_STRING_BYTES {
            return Err(Error::Limit {
                resource: "media track attribute",
                limit: MAX_STRING_BYTES,
            });
        }
        let local_span = attribute_span(element.as_ref(), attribute.key.as_ref())?;
        let value_start = start
            .checked_add(1)
            .and_then(|value| value.checked_add(local_span.value_start))
            .ok_or_else(|| invalid("media track attribute offset overflow"))?;
        let value_end = start
            .checked_add(1)
            .and_then(|value| value.checked_add(local_span.value_end))
            .ok_or_else(|| invalid("media track attribute offset overflow"))?;
        let full_start = start
            .checked_add(1)
            .and_then(|value| value.checked_add(local_span.name_start))
            .ok_or_else(|| invalid("media track attribute offset overflow"))?;
        let full_end = start
            .checked_add(1)
            .and_then(|value| value.checked_add(local_span.end))
            .ok_or_else(|| invalid("media track attribute offset overflow"))?;
        result = Some(Attr {
            value,
            span: value_start..value_end,
            full_span: full_start..full_end,
        });
    }
    Ok(result)
}

struct LocalAttrSpan {
    name_start: usize,
    value_start: usize,
    value_end: usize,
    end: usize,
}

fn attribute_span(raw: &[u8], key: &[u8]) -> Result<LocalAttrSpan> {
    let mut index = 0usize;
    while index < raw.len()
        && !raw[index].is_ascii_whitespace()
        && !matches!(raw[index], b'>' | b'/')
    {
        index += 1;
    }
    while index < raw.len() {
        while index < raw.len() && raw[index].is_ascii_whitespace() {
            index += 1;
        }
        if index >= raw.len() || raw[index] == b'>' || raw[index] == b'/' {
            break;
        }
        let name_start = index;
        while index < raw.len()
            && !raw[index].is_ascii_whitespace()
            && !matches!(raw[index], b'=' | b'>' | b'/')
        {
            index += 1;
        }
        let name = &raw[name_start..index];
        while index < raw.len() && raw[index].is_ascii_whitespace() {
            index += 1;
        }
        if index >= raw.len() || raw[index] != b'=' {
            return Err(invalid("media track attribute has no value"));
        }
        index += 1;
        while index < raw.len() && raw[index].is_ascii_whitespace() {
            index += 1;
        }
        let quote = *raw
            .get(index)
            .ok_or_else(|| invalid("media track attribute value is unterminated"))?;
        if quote != b'"' && quote != b'\'' {
            return Err(invalid("media track attribute value is not quoted"));
        }
        index += 1;
        let value_start = index;
        while index < raw.len() && raw[index] != quote {
            index += 1;
        }
        if index >= raw.len() {
            return Err(invalid("media track attribute value is unterminated"));
        }
        if name == key {
            return Ok(LocalAttrSpan {
                name_start,
                value_start,
                value_end: index,
                end: index + 1,
            });
        }
        index += 1;
    }
    Err(invalid(format!(
        "media track attribute '{}' has no source span",
        String::from_utf8_lossy(key)
    )))
}

fn opening_insert(element: &BytesStart<'_>, start: usize) -> Result<usize> {
    let raw = element.as_ref();
    let Some(close) = raw.iter().rposition(|byte| *byte == b'>') else {
        return start
            .checked_add(1)
            .and_then(|value| value.checked_add(raw.len()))
            .ok_or_else(|| invalid("media track opening offset overflow"));
    };
    let insert = if close > 0 && raw[close - 1] == b'/' {
        close - 1
    } else {
        close
    };
    start
        .checked_add(1)
        .and_then(|value| value.checked_add(insert))
        .ok_or_else(|| invalid("media track opening offset overflow"))
}

fn namespace_kind(value: ResolveResult<'_>) -> NamespaceKind {
    match value {
        ResolveResult::Bound(Namespace(value)) if value == PML || value == PML_STRICT => {
            NamespaceKind::Pml
        },
        ResolveResult::Bound(Namespace(value)) if value == P14 => NamespaceKind::P14,
        ResolveResult::Bound(Namespace(value)) if value == P15 => NamespaceKind::P15,
        ResolveResult::Bound(Namespace(value)) if value == P173 => NamespaceKind::P173,
        _ => NamespaceKind::Other,
    }
}

fn position(reader: &NsReader<&[u8]>) -> Result<usize> {
    usize::try_from(reader.buffer_position())
        .map_err(|_| invalid("media tracks XML offset does not fit usize"))
}

fn xml_error(error: quick_xml::Error) -> Error {
    Error::Xml(error.to_string())
}

fn invalid(message: impl Into<String>) -> Error {
    Error::Invalid(message.into())
}
