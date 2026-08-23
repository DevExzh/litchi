//! Bounded `DrawingML` text extraction and borrowed element-range scanning.
//!
//! The functions in this focused module operate on one shape owner at a time;
//! they preserve the source bytes and allocate only the requested text result
//! or the caller's selected ranges.

use crate::{Error, Result};
use litchi_ooxml_common::private::BindingTracker;
use litchi_ooxml_common::xml::{
    DRAWINGML_NAMESPACE, STRICT_DRAWINGML_NAMESPACE, decode_xml_reference,
};
use quick_xml::XmlVersion;
use quick_xml::events::Event;
use quick_xml::name::{Namespace, QName, ResolveResult};
use quick_xml::reader::Reader;

/// Maximum nesting depth accepted by the `DrawingML` text and range scanners,
/// matching the hardened slide element scanner.
const MAX_TEXT_SCAN_DEPTH: usize = 128;
/// Maximum number of elements visited by either `DrawingML` scanner.
const MAX_TEXT_SCAN_NODES: usize = 1_000_000;

fn is_drawingml_name(
    namespace: &ResolveResult<'_>,
    name: QName<'_>,
    local_name: &[u8],
    fragment_prefix: &Option<Option<Vec<u8>>>,
) -> bool {
    if name.local_name().as_ref() != local_name {
        return false;
    }
    match namespace {
        ResolveResult::Bound(Namespace(value)) => {
            *value == DRAWINGML_NAMESPACE || *value == STRICT_DRAWINGML_NAMESPACE
        },
        ResolveResult::Unknown(prefix) => {
            prefix.as_slice() == b"a"
                || fragment_prefix
                    .as_ref()
                    .and_then(|prefix| prefix.as_deref())
                    == Some(prefix.as_slice())
        },
        ResolveResult::Unbound => fragment_prefix == &Some(None),
    }
}

/// Extract visible `DrawingML` text from a shape or text-body fragment.
///
/// # Errors
///
/// Returns an error if the input cannot be read or is malformed.
pub fn extract(xml_bytes: &[u8], paragraph_separator: Option<char>) -> Result<String> {
    let mut reader = Reader::from_reader(xml_bytes);
    let mut tracker = BindingTracker::new();
    let mut pending_pop = false;
    let mut result = String::with_capacity(xml_bytes.len() / 8);
    let mut fragment_prefix: Option<Option<Vec<u8>>> = None;
    let mut depth = 0usize;
    let mut nodes = 0usize;
    let mut text_depth = None;
    let mut seen_paragraph = false;

    loop {
        if pending_pop {
            tracker.pop();
            pending_pop = false;
        }
        let event = reader
            .read_event()
            .map_err(|error| Error::Xml(error.to_string()))?;
        let namespace = match &event {
            Event::Start(element) => {
                tracker
                    .push(element)
                    .map_err(|error| Error::Xml(error.to_string()))?;
                tracker.resolve_element(element.name()).0
            },
            Event::Empty(element) => {
                tracker
                    .push(element)
                    .map_err(|error| Error::Xml(error.to_string()))?;
                pending_pop = true;
                tracker.resolve_element(element.name()).0
            },
            Event::End(element) => {
                pending_pop = true;
                tracker.resolve_element(element.name()).0
            },
            _ => ResolveResult::Unbound,
        };
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
            Event::Start(element) => {
                nodes = nodes.checked_add(1).ok_or_else(|| {
                    Error::Invalid("DrawingML element counter overflow".to_string())
                })?;
                if nodes > MAX_TEXT_SCAN_NODES {
                    return Err(Error::Invalid(format!(
                        "DrawingML XML exceeds {MAX_TEXT_SCAN_NODES} elements"
                    )));
                }
                depth = depth
                    .checked_add(1)
                    .ok_or_else(|| Error::Invalid("DrawingML nesting is too deep".to_string()))?;
                if depth > MAX_TEXT_SCAN_DEPTH {
                    return Err(Error::Invalid(format!(
                        "DrawingML nesting exceeds the {MAX_TEXT_SCAN_DEPTH} depth limit"
                    )));
                }
                if is_drawingml_name(&namespace, element.name(), b"p", &fragment_prefix) {
                    if seen_paragraph
                        && !result.is_empty()
                        && let Some(separator) = paragraph_separator
                        && !result.ends_with(separator)
                    {
                        result.push(separator);
                    }
                    seen_paragraph = true;
                } else if text_depth.is_none()
                    && is_drawingml_name(&namespace, element.name(), b"t", &fragment_prefix)
                {
                    text_depth = Some(depth);
                } else if is_drawingml_name(&namespace, element.name(), b"br", &fragment_prefix) {
                    result.push('\n');
                } else if is_drawingml_name(&namespace, element.name(), b"tab", &fragment_prefix) {
                    result.push('\t');
                }
            },
            Event::Empty(element) => {
                nodes = nodes.checked_add(1).ok_or_else(|| {
                    Error::Invalid("DrawingML element counter overflow".to_string())
                })?;
                if nodes > MAX_TEXT_SCAN_NODES {
                    return Err(Error::Invalid(format!(
                        "DrawingML XML exceeds {MAX_TEXT_SCAN_NODES} elements"
                    )));
                }
                if is_drawingml_name(&namespace, element.name(), b"p", &fragment_prefix) {
                    if seen_paragraph
                        && !result.is_empty()
                        && let Some(separator) = paragraph_separator
                        && !result.ends_with(separator)
                    {
                        result.push(separator);
                    }
                    seen_paragraph = true;
                } else if is_drawingml_name(&namespace, element.name(), b"br", &fragment_prefix) {
                    result.push('\n');
                } else if is_drawingml_name(&namespace, element.name(), b"tab", &fragment_prefix) {
                    result.push('\t');
                }
            },
            Event::Text(text) if text_depth.is_some() => {
                let decoded = text
                    .xml_content(XmlVersion::Explicit1_0)
                    .map_err(|error| Error::Xml(error.to_string()))?;
                result.push_str(
                    &quick_xml::escape::unescape(&decoded)
                        .map_err(|error| Error::Xml(error.to_string()))?,
                );
            },
            Event::CData(text) if text_depth.is_some() => {
                result.push_str(
                    &text
                        .xml_content(XmlVersion::Explicit1_0)
                        .map_err(|error| Error::Xml(error.to_string()))?,
                );
            },
            Event::GeneralRef(reference) if text_depth.is_some() => {
                result.push_str(&decode_xml_reference(&reference)?);
            },
            Event::End(element) => {
                if text_depth == Some(depth)
                    && is_drawingml_name(&namespace, element.name(), b"t", &fragment_prefix)
                {
                    text_depth = None;
                }
                depth = depth
                    .checked_sub(1)
                    .ok_or_else(|| Error::Invalid("invalid DrawingML nesting".to_string()))?;
            },
            Event::Eof if depth != 0 || text_depth.is_some() => {
                return Err(Error::Invalid("unterminated DrawingML XML".to_string()));
            },
            Event::Eof => break,
            _ => {},
        }
    }
    result.shrink_to_fit();
    Ok(result)
}

/// Scan checked byte ranges for `DrawingML` elements with the requested local name.
///
/// # Errors
///
/// Returns an error if the input cannot be read or is malformed.
pub fn scan_ranges(
    xml_bytes: &[u8],
    target: &[u8],
    mut emit: impl FnMut(u32, u32) -> Result<()>,
) -> Result<()> {
    enum ScanEvent {
        Start,
        NestedStart,
        Empty,
        End,
        Eof,
        Other,
    }

    let mut reader = Reader::from_reader(xml_bytes);
    let mut tracker = BindingTracker::new();
    let mut pending_pop = false;
    let mut fragment_prefix: Option<Option<Vec<u8>>> = None;
    let mut capture: Option<(usize, usize)> = None;
    let mut depth = 0usize;
    let mut nodes = 0usize;
    loop {
        if pending_pop {
            tracker.pop();
            pending_pop = false;
        }
        let event_start = usize::try_from(reader.buffer_position())
            .map_err(|_err| Error::Invalid("DrawingML offset does not fit usize".to_string()))?;
        let event = {
            let event = reader
                .read_event()
                .map_err(|error| Error::Xml(error.to_string()))?;
            let namespace = match &event {
                Event::Start(element) => {
                    tracker
                        .push(element)
                        .map_err(|error| Error::Xml(error.to_string()))?;
                    tracker.resolve_element(element.name()).0
                },
                Event::Empty(element) => {
                    tracker
                        .push(element)
                        .map_err(|error| Error::Xml(error.to_string()))?;
                    pending_pop = true;
                    tracker.resolve_element(element.name()).0
                },
                Event::End(element) => {
                    pending_pop = true;
                    tracker.resolve_element(element.name()).0
                },
                _ => ResolveResult::Unbound,
            };
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
            match &event {
                Event::Start(_) => {
                    nodes = nodes.checked_add(1).ok_or_else(|| {
                        Error::Invalid("DrawingML element counter overflow".to_string())
                    })?;
                    if nodes > MAX_TEXT_SCAN_NODES {
                        return Err(Error::Invalid(format!(
                            "DrawingML XML exceeds {MAX_TEXT_SCAN_NODES} elements"
                        )));
                    }
                    depth = depth.checked_add(1).ok_or_else(|| {
                        Error::Invalid("DrawingML nesting is too deep".to_string())
                    })?;
                    if depth > MAX_TEXT_SCAN_DEPTH {
                        return Err(Error::Invalid(format!(
                            "DrawingML nesting exceeds the {MAX_TEXT_SCAN_DEPTH} depth limit"
                        )));
                    }
                },
                Event::Empty(_) => {
                    nodes = nodes.checked_add(1).ok_or_else(|| {
                        Error::Invalid("DrawingML element counter overflow".to_string())
                    })?;
                    if nodes > MAX_TEXT_SCAN_NODES {
                        return Err(Error::Invalid(format!(
                            "DrawingML XML exceeds {MAX_TEXT_SCAN_NODES} elements"
                        )));
                    }
                },
                Event::End(_) => {
                    depth = depth
                        .checked_sub(1)
                        .ok_or_else(|| Error::Invalid("invalid DrawingML nesting".to_string()))?;
                },
                _ => {},
            }
            match event {
                Event::Start(_) if capture.is_some() => ScanEvent::NestedStart,
                Event::Start(element)
                    if is_drawingml_name(&namespace, element.name(), target, &fragment_prefix) =>
                {
                    ScanEvent::Start
                },
                Event::Empty(element)
                    if capture.is_none()
                        && is_drawingml_name(
                            &namespace,
                            element.name(),
                            target,
                            &fragment_prefix,
                        ) =>
                {
                    ScanEvent::Empty
                },
                Event::End(_) if capture.is_some() => ScanEvent::End,
                Event::Eof => ScanEvent::Eof,
                _ => ScanEvent::Other,
            }
        };
        let event_end = usize::try_from(reader.buffer_position())
            .map_err(|_err| Error::Invalid("DrawingML offset does not fit usize".to_string()))?;

        match event {
            ScanEvent::Start => capture = Some((event_start, 1)),
            ScanEvent::NestedStart => {
                let Some((_, depth)) = capture.as_mut() else {
                    return Err(Error::Invalid(
                        "missing captured DrawingML element".to_string(),
                    ));
                };
                *depth = depth
                    .checked_add(1)
                    .ok_or_else(|| Error::Invalid("DrawingML nesting is too deep".to_string()))?;
            },
            ScanEvent::Empty => emit_drawingml_range(event_start, event_end, &mut emit)?,
            ScanEvent::End => {
                let Some((_, depth)) = capture.as_mut() else {
                    return Err(Error::Invalid(
                        "missing captured DrawingML element".to_string(),
                    ));
                };
                *depth = depth
                    .checked_sub(1)
                    .ok_or_else(|| Error::Invalid("invalid DrawingML nesting".to_string()))?;
                if *depth == 0 {
                    let Some((start, _)) = capture.take() else {
                        return Err(Error::Invalid(
                            "missing DrawingML element range".to_string(),
                        ));
                    };
                    emit_drawingml_range(start, event_end, &mut emit)?;
                }
            },
            ScanEvent::Eof if capture.is_some() || depth != 0 => {
                return Err(Error::Invalid("unterminated DrawingML element".to_string()));
            },
            ScanEvent::Eof => break,
            ScanEvent::Other => {},
        }
    }
    Ok(())
}

fn emit_drawingml_range(
    start: usize,
    end: usize,
    emit: &mut impl FnMut(u32, u32) -> Result<()>,
) -> Result<()> {
    let length = end
        .checked_sub(start)
        .ok_or_else(|| Error::Invalid("invalid DrawingML element range".to_string()))?;
    emit(
        u32::try_from(start)
            .map_err(|_err| Error::Invalid("DrawingML offset exceeds u32".to_string()))?,
        u32::try_from(length)
            .map_err(|_err| Error::Invalid("DrawingML element length exceeds u32".to_string()))?,
    )
}

#[cfg(test)]
#[allow(
    clippy::unwrap_used,
    clippy::expect_used,
    reason = "test assertions panic on failure by design"
)]
mod tests {
    use super::*;
    use litchi_ooxml_common::xml::extract_omml_formulas;
    use quick_xml::reader::NsReader;

    fn extract_nsreader_oracle(
        xml_bytes: &[u8],
        paragraph_separator: Option<char>,
    ) -> Result<String> {
        let mut reader = NsReader::from_reader(xml_bytes);
        let mut result = String::with_capacity(xml_bytes.len() / 8);
        let mut fragment_prefix: Option<Option<Vec<u8>>> = None;
        let mut depth = 0usize;
        let mut nodes = 0usize;
        let mut text_depth = None;
        let mut seen_paragraph = false;

        loop {
            let (namespace, event) = reader
                .read_resolved_event()
                .map_err(|error| Error::Xml(error.to_string()))?;
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
                Event::Start(element) => {
                    nodes = nodes.checked_add(1).ok_or_else(|| {
                        Error::Invalid("DrawingML element counter overflow".to_string())
                    })?;
                    if nodes > MAX_TEXT_SCAN_NODES {
                        return Err(Error::Invalid(format!(
                            "DrawingML XML exceeds {MAX_TEXT_SCAN_NODES} elements"
                        )));
                    }
                    depth = depth.checked_add(1).ok_or_else(|| {
                        Error::Invalid("DrawingML nesting is too deep".to_string())
                    })?;
                    if depth > MAX_TEXT_SCAN_DEPTH {
                        return Err(Error::Invalid(format!(
                            "DrawingML nesting exceeds the {MAX_TEXT_SCAN_DEPTH} depth limit"
                        )));
                    }
                    if is_drawingml_name(&namespace, element.name(), b"p", &fragment_prefix) {
                        if seen_paragraph
                            && !result.is_empty()
                            && let Some(separator) = paragraph_separator
                            && !result.ends_with(separator)
                        {
                            result.push(separator);
                        }
                        seen_paragraph = true;
                    } else if text_depth.is_none()
                        && is_drawingml_name(&namespace, element.name(), b"t", &fragment_prefix)
                    {
                        text_depth = Some(depth);
                    } else if is_drawingml_name(&namespace, element.name(), b"br", &fragment_prefix)
                    {
                        result.push('\n');
                    } else if is_drawingml_name(
                        &namespace,
                        element.name(),
                        b"tab",
                        &fragment_prefix,
                    ) {
                        result.push('\t');
                    }
                },
                Event::Empty(element) => {
                    nodes = nodes.checked_add(1).ok_or_else(|| {
                        Error::Invalid("DrawingML element counter overflow".to_string())
                    })?;
                    if nodes > MAX_TEXT_SCAN_NODES {
                        return Err(Error::Invalid(format!(
                            "DrawingML XML exceeds {MAX_TEXT_SCAN_NODES} elements"
                        )));
                    }
                    if is_drawingml_name(&namespace, element.name(), b"p", &fragment_prefix) {
                        if seen_paragraph
                            && !result.is_empty()
                            && let Some(separator) = paragraph_separator
                            && !result.ends_with(separator)
                        {
                            result.push(separator);
                        }
                        seen_paragraph = true;
                    } else if is_drawingml_name(&namespace, element.name(), b"br", &fragment_prefix)
                    {
                        result.push('\n');
                    } else if is_drawingml_name(
                        &namespace,
                        element.name(),
                        b"tab",
                        &fragment_prefix,
                    ) {
                        result.push('\t');
                    }
                },
                Event::Text(text) if text_depth.is_some() => {
                    let decoded = text
                        .xml_content(XmlVersion::Explicit1_0)
                        .map_err(|error| Error::Xml(error.to_string()))?;
                    result.push_str(
                        &quick_xml::escape::unescape(&decoded)
                            .map_err(|error| Error::Xml(error.to_string()))?,
                    );
                },
                Event::CData(text) if text_depth.is_some() => {
                    result.push_str(
                        &text
                            .xml_content(XmlVersion::Explicit1_0)
                            .map_err(|error| Error::Xml(error.to_string()))?,
                    );
                },
                Event::GeneralRef(reference) if text_depth.is_some() => {
                    result.push_str(&decode_xml_reference(&reference)?);
                },
                Event::End(element) => {
                    if text_depth == Some(depth)
                        && is_drawingml_name(&namespace, element.name(), b"t", &fragment_prefix)
                    {
                        text_depth = None;
                    }
                    depth = depth
                        .checked_sub(1)
                        .ok_or_else(|| Error::Invalid("invalid DrawingML nesting".to_string()))?;
                },
                Event::Eof if depth != 0 || text_depth.is_some() => {
                    return Err(Error::Invalid("unterminated DrawingML XML".to_string()));
                },
                Event::Eof => break,
                _ => {},
            }
        }
        result.shrink_to_fit();
        Ok(result)
    }

    fn assert_extract_parity(xml: &[u8], paragraph_separator: Option<char>) {
        let tracked = extract(xml, paragraph_separator);
        let oracle = extract_nsreader_oracle(xml, paragraph_separator);
        match (tracked, oracle) {
            (Ok(tracked), Ok(oracle)) => assert_eq!(tracked, oracle),
            (Err(tracked), Err(oracle)) => assert_eq!(tracked.to_string(), oracle.to_string()),
            (tracked, oracle) => {
                panic!("tracker/oracle outcome mismatch: {tracked:?} vs {oracle:?}")
            },
        }
    }

    fn scan_ranges_nsreader_oracle(xml_bytes: &[u8], target: &[u8]) -> Result<Vec<(u32, u32)>> {
        enum ScanEvent {
            Start,
            NestedStart,
            Empty,
            End,
            Eof,
            Other,
        }

        let mut reader = NsReader::from_reader(xml_bytes);
        let mut fragment_prefix: Option<Option<Vec<u8>>> = None;
        let mut capture: Option<(usize, usize)> = None;
        let mut depth = 0usize;
        let mut nodes = 0usize;
        let mut ranges = Vec::new();
        loop {
            let event_start = usize::try_from(reader.buffer_position()).map_err(|_error| {
                Error::Invalid("DrawingML offset does not fit usize".to_string())
            })?;
            let event = {
                let (namespace, event) = reader
                    .read_resolved_event()
                    .map_err(|error| Error::Xml(error.to_string()))?;
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
                match &event {
                    Event::Start(_) => {
                        nodes = nodes.checked_add(1).ok_or_else(|| {
                            Error::Invalid("DrawingML element counter overflow".to_string())
                        })?;
                        if nodes > MAX_TEXT_SCAN_NODES {
                            return Err(Error::Invalid(format!(
                                "DrawingML XML exceeds {MAX_TEXT_SCAN_NODES} elements"
                            )));
                        }
                        depth = depth.checked_add(1).ok_or_else(|| {
                            Error::Invalid("DrawingML nesting is too deep".to_string())
                        })?;
                        if depth > MAX_TEXT_SCAN_DEPTH {
                            return Err(Error::Invalid(format!(
                                "DrawingML nesting exceeds the {MAX_TEXT_SCAN_DEPTH} depth limit"
                            )));
                        }
                    },
                    Event::Empty(_) => {
                        nodes = nodes.checked_add(1).ok_or_else(|| {
                            Error::Invalid("DrawingML element counter overflow".to_string())
                        })?;
                        if nodes > MAX_TEXT_SCAN_NODES {
                            return Err(Error::Invalid(format!(
                                "DrawingML XML exceeds {MAX_TEXT_SCAN_NODES} elements"
                            )));
                        }
                    },
                    Event::End(_) => {
                        depth = depth.checked_sub(1).ok_or_else(|| {
                            Error::Invalid("invalid DrawingML nesting".to_string())
                        })?;
                    },
                    _ => {},
                }
                match event {
                    Event::Start(_) if capture.is_some() => ScanEvent::NestedStart,
                    Event::Start(element)
                        if is_drawingml_name(
                            &namespace,
                            element.name(),
                            target,
                            &fragment_prefix,
                        ) =>
                    {
                        ScanEvent::Start
                    },
                    Event::Empty(element)
                        if capture.is_none()
                            && is_drawingml_name(
                                &namespace,
                                element.name(),
                                target,
                                &fragment_prefix,
                            ) =>
                    {
                        ScanEvent::Empty
                    },
                    Event::End(_) if capture.is_some() => ScanEvent::End,
                    Event::Eof => ScanEvent::Eof,
                    _ => ScanEvent::Other,
                }
            };
            let event_end = usize::try_from(reader.buffer_position()).map_err(|_error| {
                Error::Invalid("DrawingML offset does not fit usize".to_string())
            })?;
            match event {
                ScanEvent::Start => capture = Some((event_start, 1)),
                ScanEvent::NestedStart => {
                    let Some((_, captured_depth)) = capture.as_mut() else {
                        return Err(Error::Invalid(
                            "missing captured DrawingML element".to_string(),
                        ));
                    };
                    *captured_depth = captured_depth.checked_add(1).ok_or_else(|| {
                        Error::Invalid("DrawingML nesting is too deep".to_string())
                    })?;
                },
                ScanEvent::Empty => {
                    emit_drawingml_range(event_start, event_end, &mut |start, length| {
                        ranges.push((start, length));
                        Ok(())
                    })?;
                },
                ScanEvent::End => {
                    let Some((_, captured_depth)) = capture.as_mut() else {
                        return Err(Error::Invalid(
                            "missing captured DrawingML element".to_string(),
                        ));
                    };
                    *captured_depth = captured_depth
                        .checked_sub(1)
                        .ok_or_else(|| Error::Invalid("invalid DrawingML nesting".to_string()))?;
                    if *captured_depth == 0 {
                        let Some((start, _)) = capture.take() else {
                            return Err(Error::Invalid(
                                "missing DrawingML element range".to_string(),
                            ));
                        };
                        emit_drawingml_range(start, event_end, &mut |start, length| {
                            ranges.push((start, length));
                            Ok(())
                        })?;
                    }
                },
                ScanEvent::Eof if capture.is_some() || depth != 0 => {
                    return Err(Error::Invalid("unterminated DrawingML element".to_string()));
                },
                ScanEvent::Eof => break,
                ScanEvent::Other => {},
            }
        }
        Ok(ranges)
    }

    fn scan_ranges_tracked(xml: &[u8], target: &[u8]) -> Result<Vec<(u32, u32)>> {
        let mut ranges = Vec::new();
        scan_ranges(xml, target, |start, length| {
            ranges.push((start, length));
            Ok(())
        })?;
        Ok(ranges)
    }

    fn assert_scan_ranges_parity(xml: &[u8], target: &[u8]) {
        let tracked = scan_ranges_tracked(xml, target);
        let oracle = scan_ranges_nsreader_oracle(xml, target);
        match (tracked, oracle) {
            (Ok(tracked), Ok(oracle)) => assert_eq!(tracked, oracle),
            (Err(tracked), Err(oracle)) => assert_eq!(tracked.to_string(), oracle.to_string()),
            (tracked, oracle) => {
                panic!("tracker/oracle outcome mismatch: {tracked:?} vs {oracle:?}")
            },
        }
    }

    #[test]
    fn drawingml_text_preserves_runs_whitespace_and_paragraphs() {
        let xml = br#"<p:sp xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:d="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:false="urn:not-drawingml">
            <d:txBody>
                <d:p><d:r><d:t xml:space="preserve"> A &amp; </d:t></d:r><d:r><false:t>ignored</false:t><d:t><![CDATA[B < C]]></d:t><d:tab/><d:br/></d:r></d:p>
                <d:p><d:r><d:t>Second</d:t></d:r></d:p>
            </d:txBody>
        </p:sp>"#;
        assert_eq!(extract(xml, Some('\n')).unwrap(), " A & B < C\t\nSecond");
        let mut paragraphs = Vec::new();
        scan_ranges(xml, b"p", |start, length| {
            let start = start as usize;
            let end = start + length as usize;
            paragraphs.push(extract(&xml[start..end], None)?);
            Ok(())
        })
        .unwrap();
        assert_eq!(paragraphs.len(), 2);
        assert_eq!(paragraphs[0], " A & B < C\t\n");
        assert_eq!(paragraphs[1], "Second");
    }

    #[test]
    fn drawingml_paragraph_text_accepts_inherited_conventional_prefix() {
        let paragraph = br"<a:p><a:r><a:t>one</a:t></a:r><a:r><a:t>two</a:t></a:r></a:p>";
        assert_eq!(extract(paragraph, None).unwrap(), "onetwo");
    }

    #[test]
    fn drawingml_text_rejects_foreign_lookalikes_and_truncation() {
        let foreign = br#"<x:p xmlns:x="urn:not-drawingml"><x:r><x:t>ignored</x:t></x:r></x:p>"#;
        assert_eq!(extract(foreign, None).unwrap(), "");

        let truncated = br#"<a:p xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:r><a:t>bad</a:t>"#;
        assert!(extract(truncated, None).is_err());
    }

    #[test]
    fn drawingml_range_scan_rejects_structural_depth_over_bound() {
        let mut xml = String::from("<root>");
        xml.push_str(&"<node>".repeat(MAX_TEXT_SCAN_DEPTH));
        xml.push_str(&"</node>".repeat(MAX_TEXT_SCAN_DEPTH));
        xml.push_str("</root>");
        assert!(scan_ranges(xml.as_bytes(), b"p", |_, _| Ok(())).is_err());
    }

    #[test]
    fn drawingml_tracker_matches_nsreader_on_event_and_namespace_adversaries() {
        let drawingml = "http://schemas.openxmlformats.org/drawingml/2006/main";
        let xml_namespace = "http://www.w3.org/XML/1998/namespace";
        let xmlns_namespace = "http://www.w3.org/2000/xmlns/";
        let fixtures = [
            format!(
                r#"<?xml version="1.0"?><!-- before --><root xmlns:a="{drawingml}"><![CDATA[ignored]]><a:p><?inside?><a:r><a:t>x &amp; <![CDATA[y < z]]>&#x21;</a:t></a:r><a:br/><a:tab/></a:p></root>"#
            ),
            format!(
                r#"<root xmlns:a="{drawingml}" xmlns:q="urn:outer"><q:holder xmlns:q="urn:inner"><a:p><a:t>one</a:t></a:p></q:holder><a:p><a:t>two</a:t></a:p></root>"#
            ),
            // A malformed attribute after a valid declaration must retain
            // quick-xml's silent namespace-scan stop.
            format!(
                r#"<root xmlns:a="{drawingml}" a:broken="unterminated><a:p><a:t>x</a:t></a:p></root>"#
            ),
            format!(
                r#"<root xmlns:a="{drawingml}" xmlns:xml="{xml_namespace}"><a:p><a:t>x</a:t></a:p></root>"#
            ),
            format!(
                r#"<root xmlns:a="{drawingml}" xmlns:q="{xmlns_namespace}"><a:p><a:t>x</a:t></a:p></root>"#
            ),
        ];
        for fixture in fixtures {
            assert_extract_parity(fixture.as_bytes(), Some('\n'));
            assert_scan_ranges_parity(fixture.as_bytes(), b"p");
        }
    }

    #[test]
    fn drawingml_scan_ranges_tracker_matches_nsreader_on_namespace_adversaries() {
        let drawingml = "http://schemas.openxmlformats.org/drawingml/2006/main";
        let xml_namespace = "http://www.w3.org/XML/1998/namespace";
        let xmlns_namespace = "http://www.w3.org/2000/xmlns/";
        let declaration_attributes = |count: usize| {
            (0..count)
                .map(|index| format!(r#"xmlns:n{index}="urn:n{index}""#))
                .collect::<Vec<_>>()
                .join(" ")
        };
        let fixtures = [
            // Fragment scans preserve the conventional inherited prefix and
            // the unbound-prefix fallback used by shape snippets.
            String::from(r#"<a:p><a:r/></a:p>"#),
            String::from(r#"<p><r/></p>"#),
            format!(
                r#"<root xmlns:a="{drawingml}"><a:p/><scope xmlns:a="urn:foreign"><a:p/></scope><a:p/></root>"#
            ),
            format!(r#"<root xmlns:a="{drawingml}" xmlns:xml="urn:foreign"><a:p/></root>"#),
            format!(r#"<root xmlns:a="{drawingml}" xmlns:q="{xml_namespace}"><a:p/></root>"#),
            format!(r#"<root xmlns:a="{drawingml}" xmlns:q="{xmlns_namespace}"><a:p/></root>"#),
            // The valid declaration is before the malformed attribute, so the
            // unchecked namespace scan must stop at the same point as NsReader.
            format!(r#"<root xmlns:a="{drawingml}" a:broken="unterminated><a:p/></root>"#),
            // a plus 255 additional declarations is exactly the inclusive limit.
            format!(
                r#"<root xmlns:a="{drawingml}" {}><a:p/></root>"#,
                declaration_attributes(255)
            ),
            // One more declaration must fail before a range is emitted.
            format!(
                r#"<root xmlns:a="{drawingml}" {}><a:p/></root>"#,
                declaration_attributes(256)
            ),
        ];
        for fixture in fixtures {
            assert_scan_ranges_parity(fixture.as_bytes(), b"p");
        }

        let mut at_limit = format!(r#"<root xmlns:a="{drawingml}">"#);
        at_limit.push_str(&"<scope>".repeat(MAX_TEXT_SCAN_DEPTH - 1));
        at_limit.push_str("<a:p/>");
        at_limit.push_str(&"</scope>".repeat(MAX_TEXT_SCAN_DEPTH - 1));
        at_limit.push_str("</root>");
        assert_scan_ranges_parity(at_limit.as_bytes(), b"p");

        let mut too_deep = format!(r#"<root xmlns:a="{drawingml}">"#);
        too_deep.push_str(&"<scope>".repeat(MAX_TEXT_SCAN_DEPTH));
        too_deep.push_str("<a:p/>");
        too_deep.push_str(&"</scope>".repeat(MAX_TEXT_SCAN_DEPTH));
        too_deep.push_str("</root>");
        assert_scan_ranges_parity(too_deep.as_bytes(), b"p");
    }

    #[test]
    fn drawingml_tracker_preserves_depth_limit_and_malformed_errors() {
        let drawingml = "http://schemas.openxmlformats.org/drawingml/2006/main";
        let deep = format!(
            r#"<a:p xmlns:a="{drawingml}"><a:r><a:t>{}</a:t></a:r></a:p>"#,
            "x"
        );
        // The regular fixture exercises the no-limit path; an actually deep
        // fragment checks that the tracker does not bypass the caller's
        // existing structural bound.
        let mut nested = format!(r#"<a:p xmlns:a="{drawingml}">"#);
        nested.push_str(&"<a:r>".repeat(MAX_TEXT_SCAN_DEPTH));
        nested.push_str("<a:t>x</a:t>");
        nested.push_str(&"</a:r>".repeat(MAX_TEXT_SCAN_DEPTH));
        nested.push_str("</a:p>");
        assert_extract_parity(deep.as_bytes(), None);
        assert_extract_parity(nested.as_bytes(), None);

        let malformed = format!(r#"<a:p xmlns:a="{drawingml}"><a:t a="unterminated></a:p>"#);
        assert_extract_parity(malformed.as_bytes(), None);
    }

    #[test]
    fn omml_formulas_preserve_exact_namespaced_elements() {
        let xml = br#"<p:sp xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
            xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
            xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"
            xmlns:math="http://schemas.openxmlformats.org/officeDocument/2006/math"
            xmlns:false="urn:not-omml">
            <a:txBody><a:p><a:r><a:t>ordinary m:oMath text</a:t></a:r>
                <m:oMath data-id="1"><m:r><m:t><![CDATA[x < y]]></m:t></m:r></m:oMath>
                <false:oMath>ignored</false:oMath>
                <math:oMath math:id="2"/>
            </a:p></a:txBody>
        </p:sp>"#;
        assert_eq!(
            extract_omml_formulas(xml).unwrap(),
            vec![
                r#"<m:oMath data-id="1"><m:r><m:t><![CDATA[x < y]]></m:t></m:r></m:oMath>"#,
                r#"<math:oMath math:id="2"/>"#,
            ]
        );
    }

    #[test]
    fn omml_formulas_accept_strict_and_inherited_conventional_prefixes() {
        let strict = br#"<root xmlns:s="http://purl.oclc.org/ooxml/officeDocument/math"><s:oMath><s:r/></s:oMath></root>"#;
        assert_eq!(
            extract_omml_formulas(strict).unwrap(),
            vec!["<s:oMath><s:r/></s:oMath>"]
        );

        assert_eq!(
            extract_omml_formulas(br"<a:p><m:oMath><m:r/></m:oMath></a:p>").unwrap(),
            vec!["<m:oMath><m:r/></m:oMath>"]
        );
    }

    #[test]
    fn omml_formulas_reject_malformed_xml() {
        let xml = br#"<a:p xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"><m:oMath><m:r/></a:p>"#;
        assert!(extract_omml_formulas(xml).is_err());
    }
}
