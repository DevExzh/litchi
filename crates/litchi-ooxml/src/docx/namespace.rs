use crate::error::{OoxmlError, Result};
use quick_xml::XmlVersion;
use quick_xml::encoding::Decoder;
use quick_xml::events::{BytesStart, Event};
use quick_xml::name::{Namespace, NamespaceResolver, ResolveResult};
use quick_xml::reader::NsReader;

pub(crate) const WORDPROCESSINGML_NAMESPACE: &[u8] =
    b"http://schemas.openxmlformats.org/wordprocessingml/2006/main";
pub(crate) const STRICT_WORDPROCESSINGML_NAMESPACE: &[u8] =
    b"http://purl.oclc.org/ooxml/wordprocessingml/main";

pub(crate) fn is_wordprocessing_namespace(namespace: &ResolveResult<'_>) -> bool {
    matches!(
        namespace,
        ResolveResult::Bound(Namespace(value))
            if *value == WORDPROCESSINGML_NAMESPACE
                || *value == STRICT_WORDPROCESSINGML_NAMESPACE
    )
}

pub(crate) fn word_attribute_value(
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
        let is_word_attribute = is_wordprocessing_namespace(&namespace)
            || matches!(namespace, ResolveResult::Unbound)
            || matches!(namespace, ResolveResult::Unknown(prefix) if prefix.as_slice() == b"w");
        if !is_word_attribute {
            continue;
        }
        if value.is_some() {
            return Err(OoxmlError::InvalidFormat(format!(
                "duplicate Word attribute '{}'",
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

fn is_fragment_word_namespace(
    namespace: &ResolveResult<'_>,
    fragment_prefix: &Option<Option<Vec<u8>>>,
) -> bool {
    if is_wordprocessing_namespace(namespace) {
        return true;
    }
    match namespace {
        ResolveResult::Unknown(prefix) => {
            fragment_prefix
                .as_ref()
                .and_then(|prefix| prefix.as_deref())
                == Some(prefix.as_slice())
        },
        ResolveResult::Unbound => fragment_prefix == &Some(None),
        ResolveResult::Bound(_) => false,
    }
}

pub(crate) fn scan_word_element_ranges(
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
    let mut fragment_prefix: Option<Option<Vec<u8>>> = None;
    let mut capture: Option<(usize, usize, usize)> = None;

    loop {
        let event_start = usize::try_from(reader.buffer_position()).map_err(|_| {
            OoxmlError::InvalidFormat("Word XML offset does not fit usize".to_string())
        })?;
        let event = {
            let (namespace, event) = reader
                .read_resolved_event()
                .map_err(|error| OoxmlError::Xml(error.to_string()))?;

            if fragment_prefix.is_none()
                && let Event::Start(element) = &event
                && !matches!(namespace, ResolveResult::Bound(_))
            {
                fragment_prefix = Some(
                    element
                        .name()
                        .prefix()
                        .map(|prefix| prefix.into_inner().to_vec()),
                );
            }

            match event {
                Event::Start(_) if capture.is_some() => ScanEvent::NestedStart,
                Event::Start(element)
                    if is_fragment_word_namespace(&namespace, &fragment_prefix) =>
                {
                    targets
                        .iter()
                        .position(|target| element.local_name().as_ref() == *target)
                        .map_or(ScanEvent::Other, ScanEvent::Start)
                },
                Event::Empty(element)
                    if capture.is_none()
                        && is_fragment_word_namespace(&namespace, &fragment_prefix) =>
                {
                    targets
                        .iter()
                        .position(|target| element.local_name().as_ref() == *target)
                        .map_or(ScanEvent::Other, ScanEvent::Empty)
                },
                Event::End(_) if capture.is_some() => ScanEvent::End,
                Event::Eof => ScanEvent::Eof,
                _ => ScanEvent::Other,
            }
        };
        let event_end = usize::try_from(reader.buffer_position()).map_err(|_| {
            OoxmlError::InvalidFormat("Word XML offset does not fit usize".to_string())
        })?;

        match event {
            ScanEvent::Start(target) => capture = Some((target, event_start, 1)),
            ScanEvent::NestedStart => {
                let Some((_, _, depth)) = capture.as_mut() else {
                    return Err(OoxmlError::InvalidFormat(
                        "missing captured Word element".to_string(),
                    ));
                };
                *depth = depth.checked_add(1).ok_or_else(|| {
                    OoxmlError::InvalidFormat("Word element nesting is too deep".to_string())
                })?;
            },
            ScanEvent::Empty(target) => {
                emit_word_element_range(target, event_start, event_end, &mut emit)?;
            },
            ScanEvent::End => {
                let Some((_, _, depth)) = capture.as_mut() else {
                    return Err(OoxmlError::InvalidFormat(
                        "missing captured Word element".to_string(),
                    ));
                };
                *depth = depth.checked_sub(1).ok_or_else(|| {
                    OoxmlError::InvalidFormat("invalid Word element nesting".to_string())
                })?;
                if *depth == 0 {
                    let Some((target, start, _)) = capture.take() else {
                        return Err(OoxmlError::InvalidFormat(
                            "missing captured Word element range".to_string(),
                        ));
                    };
                    emit_word_element_range(target, start, event_end, &mut emit)?;
                }
            },
            ScanEvent::Eof if capture.is_some() => {
                return Err(OoxmlError::InvalidFormat(
                    "unterminated Word element".to_string(),
                ));
            },
            ScanEvent::Eof => break,
            ScanEvent::Other => {},
        }
    }

    Ok(())
}

fn emit_word_element_range(
    target: usize,
    start: usize,
    end: usize,
    emit: &mut impl FnMut(usize, u32, u32) -> Result<()>,
) -> Result<()> {
    let length = end
        .checked_sub(start)
        .ok_or_else(|| OoxmlError::InvalidFormat("invalid Word element byte range".to_string()))?;
    let start = u32::try_from(start)
        .map_err(|_| OoxmlError::InvalidFormat("Word element offset exceeds u32".to_string()))?;
    let length = u32::try_from(length)
        .map_err(|_| OoxmlError::InvalidFormat("Word element length exceeds u32".to_string()))?;
    emit(target, start, length)
}
