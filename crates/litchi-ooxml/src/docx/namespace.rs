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

pub(crate) fn direct_word_property_value(
    xml_bytes: &[u8],
    root_name: &[u8],
    properties_name: &[u8],
    property_name: &[u8],
) -> Result<Option<String>> {
    let mut reader = NsReader::from_reader(xml_bytes);
    let mut fragment_prefix: Option<Option<Vec<u8>>> = None;
    let mut depth = 0usize;
    let mut properties_depth = None;
    let mut saw_properties = false;
    let mut value = None;
    let mut saw_root = false;

    loop {
        let decoder = reader.decoder();
        let event = reader
            .read_event()
            .map_err(|error| OoxmlError::Xml(error.to_string()))?
            .into_owned();
        let resolver = reader.resolver().clone();
        let (namespace, event) = resolver.resolve_event(event);

        if fragment_prefix.is_none()
            && depth == 0
            && let Event::Start(element) | Event::Empty(element) = &event
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
            Event::Start(element) => {
                depth = depth.checked_add(1).ok_or_else(|| {
                    OoxmlError::InvalidFormat("Word property XML nesting is too deep".into())
                })?;
                let is_word = is_fragment_word_namespace(&namespace, &fragment_prefix);
                if depth == 1 {
                    if saw_root || !is_word || element.local_name().as_ref() != root_name {
                        return Err(OoxmlError::InvalidFormat(
                            "Word property XML has an invalid root".into(),
                        ));
                    }
                    saw_root = true;
                } else if depth == 2 && is_word && element.local_name().as_ref() == properties_name
                {
                    if saw_properties {
                        return Err(OoxmlError::InvalidFormat(
                            "duplicate Word property container".into(),
                        ));
                    }
                    saw_properties = true;
                    properties_depth = Some(depth);
                } else if depth == 3
                    && properties_depth == Some(2)
                    && is_word
                    && element.local_name().as_ref() == property_name
                {
                    set_direct_property_value(
                        &mut value,
                        &element,
                        decoder,
                        &resolver,
                        &fragment_prefix,
                        property_name,
                    )?;
                }
            },
            Event::Empty(element) => {
                let child_depth = depth.checked_add(1).ok_or_else(|| {
                    OoxmlError::InvalidFormat("Word property XML nesting is too deep".into())
                })?;
                let is_word = is_fragment_word_namespace(&namespace, &fragment_prefix);
                if child_depth == 1 {
                    if saw_root || !is_word || element.local_name().as_ref() != root_name {
                        return Err(OoxmlError::InvalidFormat(
                            "Word property XML has an invalid root".into(),
                        ));
                    }
                    saw_root = true;
                } else if child_depth == 2
                    && is_word
                    && element.local_name().as_ref() == properties_name
                {
                    if saw_properties {
                        return Err(OoxmlError::InvalidFormat(
                            "duplicate Word property container".into(),
                        ));
                    }
                    saw_properties = true;
                } else if child_depth == 3
                    && properties_depth == Some(2)
                    && is_word
                    && element.local_name().as_ref() == property_name
                {
                    set_direct_property_value(
                        &mut value,
                        &element,
                        decoder,
                        &resolver,
                        &fragment_prefix,
                        property_name,
                    )?;
                }
            },
            Event::End(_) => {
                if properties_depth == Some(depth) {
                    properties_depth = None;
                }
                depth = depth.checked_sub(1).ok_or_else(|| {
                    OoxmlError::InvalidFormat("invalid Word property XML nesting".into())
                })?;
            },
            Event::Eof if depth != 0 => {
                return Err(OoxmlError::InvalidFormat(
                    "unterminated Word property XML".into(),
                ));
            },
            Event::Eof => break,
            _ => {},
        }
    }

    if !saw_root {
        return Err(OoxmlError::InvalidFormat(
            "Word property XML has no root".into(),
        ));
    }
    Ok(value)
}

fn set_direct_property_value(
    slot: &mut Option<String>,
    element: &BytesStart<'_>,
    decoder: Decoder,
    resolver: &NamespaceResolver,
    fragment_prefix: &Option<Option<Vec<u8>>>,
    property_name: &[u8],
) -> Result<()> {
    if slot.is_some() {
        return Err(OoxmlError::InvalidFormat(format!(
            "duplicate Word property '{}'",
            String::from_utf8_lossy(property_name)
        )));
    }
    let mut value = None;
    for attribute in element.attributes() {
        let attribute = attribute.map_err(|error| OoxmlError::Xml(error.to_string()))?;
        if attribute.key.local_name().as_ref() != b"val" {
            continue;
        }
        let (namespace, _) = resolver.resolve_attribute(attribute.key);
        if !is_fragment_word_namespace(&namespace, fragment_prefix) {
            continue;
        }
        if value.is_some() {
            return Err(OoxmlError::InvalidFormat(
                "duplicate Word property value attribute".into(),
            ));
        }
        value = Some(
            attribute
                .decoded_and_normalized_value(XmlVersion::Explicit1_0, decoder)
                .map_err(|error| OoxmlError::Xml(error.to_string()))?
                .into_owned(),
        );
    }
    *slot = Some(value.ok_or_else(|| {
        OoxmlError::InvalidFormat(format!(
            "Word property '{}' requires a value",
            String::from_utf8_lossy(property_name)
        ))
    })?);
    Ok(())
}

pub(crate) fn normalize_xml_integer(value: String, description: &str) -> Result<String> {
    let value = value.trim();
    let digits = value
        .strip_prefix('+')
        .or_else(|| value.strip_prefix('-'))
        .unwrap_or(value);
    if digits.is_empty() || !digits.bytes().all(|byte| byte.is_ascii_digit()) {
        return Err(OoxmlError::InvalidFormat(format!(
            "invalid {description} value '{value}'"
        )));
    }
    Ok(value.to_owned())
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
