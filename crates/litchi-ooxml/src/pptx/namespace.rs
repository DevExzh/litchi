use crate::error::{OoxmlError, Result};
use litchi_ooxml_common::xml::unqualified_attribute_value;
use quick_xml::XmlVersion;
use quick_xml::encoding::Decoder;
use quick_xml::events::{BytesStart, Event};
use quick_xml::name::{Namespace, NamespaceResolver, QName, ResolveResult};
use quick_xml::reader::NsReader;

pub(crate) const PRESENTATIONML_NAMESPACE: &[u8] =
    b"http://schemas.openxmlformats.org/presentationml/2006/main";
pub(crate) const STRICT_PRESENTATIONML_NAMESPACE: &[u8] =
    b"http://purl.oclc.org/ooxml/presentationml/main";
pub(crate) const RELATIONSHIPS_NAMESPACE: &[u8] =
    b"http://schemas.openxmlformats.org/officeDocument/2006/relationships";
pub(crate) const STRICT_RELATIONSHIPS_NAMESPACE: &[u8] =
    b"http://purl.oclc.org/ooxml/officeDocument/relationships";

pub(crate) fn is_presentationml_name(
    namespace: &ResolveResult<'_>,
    name: QName<'_>,
    local_name: &[u8],
) -> bool {
    if name.local_name().as_ref() != local_name {
        return false;
    }
    match namespace {
        ResolveResult::Bound(Namespace(value)) => {
            *value == PRESENTATIONML_NAMESPACE || *value == STRICT_PRESENTATIONML_NAMESPACE
        },
        // Shape fragments normally inherit the conventional prefix from a slide root.
        ResolveResult::Unknown(prefix) => prefix.as_slice() == b"p",
        ResolveResult::Unbound => false,
    }
}

pub(crate) fn relationship_attribute_value(
    element: &BytesStart<'_>,
    name: &[u8],
    decoder: Decoder,
    resolver: &NamespaceResolver,
) -> Result<Option<String>> {
    let mut value = None;
    for attribute in element.attributes() {
        let attribute = attribute.map_err(|error| OoxmlError::Xml(error.to_string()))?;
        if attribute.key.local_name().as_ref() != name {
            continue;
        }
        let (namespace, _) = resolver.resolve_attribute(attribute.key);
        let is_relationship = matches!(
            namespace,
            ResolveResult::Bound(Namespace(value))
                if value == RELATIONSHIPS_NAMESPACE || value == STRICT_RELATIONSHIPS_NAMESPACE
        ) || matches!(namespace, ResolveResult::Unknown(prefix) if prefix.as_slice() == b"r");
        if !is_relationship {
            continue;
        }
        if value.is_some() {
            return Err(OoxmlError::InvalidFormat(format!(
                "duplicate relationship attribute '{}'",
                String::from_utf8_lossy(name)
            )));
        }
        value = Some(
            attribute
                .decoded_and_normalized_value(XmlVersion::Explicit1_0, decoder)
                .map_err(|error| OoxmlError::Xml(error.to_string()))?
                .into_owned(),
        );
    }
    Ok(value)
}

pub(crate) fn presentation_name(xml_bytes: &[u8]) -> Result<String> {
    let mut reader = NsReader::from_reader(xml_bytes);
    loop {
        let decoder = reader.decoder();
        let (namespace, event) = reader
            .read_resolved_event()
            .map_err(|error| OoxmlError::Xml(error.to_string()))?;
        match event {
            Event::Start(element) | Event::Empty(element)
                if is_presentationml_name(&namespace, element.name(), b"cSld") =>
            {
                return Ok(
                    unqualified_attribute_value(&element, b"name", decoder)?.unwrap_or_default()
                );
            },
            Event::Eof => return Ok(String::new()),
            _ => {},
        }
    }
}

/// Maximum capture nesting depth accepted by the shared PresentationML
/// element scanner, matching the hardened presentation part parser.
const MAX_SCAN_DEPTH: usize = 128;
/// Maximum number of elements scanned in one slide part.
const MAX_SCAN_NODES: usize = 1_000_000;

pub(crate) fn scan_presentationml_element_ranges(
    xml_bytes: &[u8],
    targets: &[&[u8]],
    mut emit: impl FnMut(usize, u32, u32) -> Result<()>,
) -> Result<()> {
    enum ScanEvent {
        Start(usize),
        NestedStart,
        Empty(usize),
        End,
        Eof,
        Other,
    }

    let mut reader = NsReader::from_reader(xml_bytes);
    let mut capture: Option<(usize, usize, usize)> = None;
    let mut nodes = 0usize;
    let mut total_depth = 0usize;
    loop {
        let event_start = usize::try_from(reader.buffer_position()).map_err(|_| {
            OoxmlError::InvalidFormat("PresentationML offset does not fit usize".to_string())
        })?;
        let event = {
            let (namespace, event) = reader
                .read_resolved_event()
                .map_err(|error| OoxmlError::Xml(error.to_string()))?;
            if matches!(event, Event::Start(_) | Event::Empty(_)) {
                nodes = nodes.checked_add(1).ok_or_else(|| {
                    OoxmlError::InvalidFormat("PresentationML element counter overflow".to_string())
                })?;
                if nodes > MAX_SCAN_NODES {
                    return Err(OoxmlError::InvalidFormat(format!(
                        "PresentationML XML exceeds {MAX_SCAN_NODES} elements"
                    )));
                }
            }
            // Total nesting is tracked separately from capture depth so
            // deeply nested non-target content is rejected before
            // quick-xml's own namespace resolver overflows (u16).
            if matches!(event, Event::Start(_)) {
                total_depth = total_depth.checked_add(1).ok_or_else(|| {
                    OoxmlError::InvalidFormat("PresentationML nesting is too deep".to_string())
                })?;
                if total_depth > MAX_SCAN_DEPTH {
                    return Err(OoxmlError::InvalidFormat(format!(
                        "PresentationML nesting exceeds the {MAX_SCAN_DEPTH} depth limit"
                    )));
                }
            }
            if matches!(event, Event::End(_)) {
                total_depth = total_depth.checked_sub(1).ok_or_else(|| {
                    OoxmlError::InvalidFormat("invalid PresentationML nesting".to_string())
                })?;
            }
            match event {
                Event::Start(_) if capture.is_some() => ScanEvent::NestedStart,
                Event::Start(element) => targets
                    .iter()
                    .position(|target| is_presentationml_name(&namespace, element.name(), target))
                    .map_or(ScanEvent::Other, ScanEvent::Start),
                Event::Empty(element) if capture.is_none() => targets
                    .iter()
                    .position(|target| is_presentationml_name(&namespace, element.name(), target))
                    .map_or(ScanEvent::Other, ScanEvent::Empty),
                Event::End(_) if capture.is_some() => ScanEvent::End,
                Event::Eof => ScanEvent::Eof,
                _ => ScanEvent::Other,
            }
        };
        let event_end = usize::try_from(reader.buffer_position()).map_err(|_| {
            OoxmlError::InvalidFormat("PresentationML offset does not fit usize".to_string())
        })?;

        match event {
            ScanEvent::Start(target) => capture = Some((target, event_start, 1)),
            ScanEvent::NestedStart => {
                let Some((_, _, depth)) = capture.as_mut() else {
                    return Err(OoxmlError::InvalidFormat(
                        "missing captured PresentationML element".to_string(),
                    ));
                };
                *depth = depth.checked_add(1).ok_or_else(|| {
                    OoxmlError::InvalidFormat("PresentationML nesting is too deep".to_string())
                })?;
                if *depth > MAX_SCAN_DEPTH {
                    return Err(OoxmlError::InvalidFormat(format!(
                        "PresentationML nesting exceeds the {MAX_SCAN_DEPTH} depth limit"
                    )));
                }
            },
            ScanEvent::Empty(target) => {
                emit_range(target, event_start, event_end, &mut emit)?;
            },
            ScanEvent::End => {
                let Some((_, _, depth)) = capture.as_mut() else {
                    return Err(OoxmlError::InvalidFormat(
                        "missing captured PresentationML element".to_string(),
                    ));
                };
                *depth = depth.checked_sub(1).ok_or_else(|| {
                    OoxmlError::InvalidFormat("invalid PresentationML nesting".to_string())
                })?;
                if *depth == 0 {
                    let Some((target, start, _)) = capture.take() else {
                        return Err(OoxmlError::InvalidFormat(
                            "missing PresentationML element range".to_string(),
                        ));
                    };
                    emit_range(target, start, event_end, &mut emit)?;
                }
            },
            ScanEvent::Eof if capture.is_some() => {
                return Err(OoxmlError::InvalidFormat(
                    "unterminated PresentationML element".to_string(),
                ));
            },
            ScanEvent::Eof => break,
            ScanEvent::Other => {},
        }
    }
    Ok(())
}

fn emit_range(
    target: usize,
    start: usize,
    end: usize,
    emit: &mut impl FnMut(usize, u32, u32) -> Result<()>,
) -> Result<()> {
    let length = end.checked_sub(start).ok_or_else(|| {
        OoxmlError::InvalidFormat("invalid PresentationML element range".to_string())
    })?;
    let start = u32::try_from(start).map_err(|_| {
        OoxmlError::InvalidFormat("PresentationML element offset exceeds u32".to_string())
    })?;
    let length = u32::try_from(length).map_err(|_| {
        OoxmlError::InvalidFormat("PresentationML element length exceeds u32".to_string())
    })?;
    emit(target, start, length)
}
