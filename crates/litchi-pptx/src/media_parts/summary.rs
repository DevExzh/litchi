//! What one media picture of a shape scene is (record 0321): its kind and
//! its playback trim, read from the picture's own fragment, so a scene's
//! `p:pic` can be classified without parsing the whole slide.

use super::{Event, Kind, Namespace, NsReader, ResolveResult, Result, Trim, parse_time, xml_error};
use litchi_ooxml_common::xml::attributes::BytesStartExt as _;

const DML: &[u8] = b"http://schemas.openxmlformats.org/drawingml/2006/main";
const STRICT_DML: &[u8] = b"http://purl.oclc.org/ooxml/drawingml/main";
const P14: &[u8] = b"http://schemas.microsoft.com/office/powerpoint/2010/main";
const MAX_DEPTH: usize = 256;

/// A media picture's kind and trim.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Summary {
    /// Audio or video.
    pub kind: Kind,
    /// The PowerPoint 2010 playback trim (`p14:media/p14:trim`), if any.
    pub trim: Option<Trim>,
}

/// The media summary of a picture shape: `Some` when its non-visual
/// properties name an audio or video file (`a:audioFile`, `a:wavAudioFile`,
/// `a:audioCd`, `a:videoFile`, `a:quickTimeFile`), else `None`.
///
/// # Errors
///
/// Returns an error for a malformed fragment or trim offset.
pub fn summary(shape: crate::shape::Shape<'_>) -> Result<Option<Summary>> {
    if !matches!(shape, crate::shape::Shape::Picture(_)) {
        return Ok(None);
    }
    let fragment = shape.self_contained_xml()?;
    let mut reader = NsReader::from_reader(fragment.as_ref());
    let mut depth = 0usize;
    let mut kind = None;
    let mut trim = None;
    loop {
        let (namespace, event) = reader.read_resolved_event().map_err(xml_error)?;
        let dml = matches!(namespace, ResolveResult::Bound(Namespace(value)) if value == DML || value == STRICT_DML);
        let p14 = matches!(namespace, ResolveResult::Bound(Namespace(value)) if value == P14);
        match event {
            Event::Start(ref element) | Event::Empty(ref element) => {
                let local = element.local_name();
                if dml && kind.is_none() {
                    kind = match local.as_ref() {
                        b"audioFile" | b"wavAudioFile" | b"audioCd" => Some(Kind::Audio),
                        b"videoFile" | b"quickTimeFile" => Some(Kind::Video),
                        _ => None,
                    };
                }
                if p14 && local.as_ref() == b"trim" {
                    let mut value = Trim::default();
                    for attribute in element.checked_attributes() {
                        let attribute =
                            attribute.map_err(|error| crate::Error::Xml(error.to_string()))?;
                        let text = attribute
                            .decoded_and_normalized_value(
                                quick_xml::XmlVersion::Implicit1_0,
                                reader.decoder(),
                            )
                            .map_err(|error| crate::Error::Xml(error.to_string()))?;
                        match attribute.key.as_ref() {
                            b"st" => value.start = Some(parse_time(&text)?),
                            b"end" => value.end = Some(parse_time(&text)?),
                            _ => {},
                        }
                    }
                    trim = Some(value);
                }
                if matches!(event, Event::Start(_)) {
                    depth = depth.checked_add(1).ok_or_else(|| {
                        super::limit(
                            crate::error::CodecLimitKind::XmlDepth,
                            usize::MAX,
                            MAX_DEPTH,
                        )
                    })?;
                    if depth > MAX_DEPTH {
                        return Err(super::limit(
                            crate::error::CodecLimitKind::XmlDepth,
                            depth,
                            MAX_DEPTH,
                        ));
                    }
                }
            },
            Event::End(_) => {
                depth = depth.saturating_sub(1);
                if depth == 0 {
                    break;
                }
            },
            Event::Eof => break,
            _ => {},
        }
    }
    Ok(kind.map(|kind| Summary { kind, trim }))
}
