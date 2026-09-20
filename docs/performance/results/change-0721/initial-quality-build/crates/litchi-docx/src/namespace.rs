#![expect(
    clippy::arbitrary_source_item_ordering,
    reason = "items remain grouped by OOXML schema family and package lifecycle"
)]
#![expect(
    clippy::needless_pass_by_value,
    reason = "the public API shape is retained for compatibility"
)]
#![expect(
    clippy::option_option,
    reason = "nested options distinguish omitted, present-empty, and present-valued XML"
)]
#![expect(
    clippy::ref_option,
    reason = "the public API shape is retained for compatibility"
)]
#![expect(
    clippy::shadow_reuse,
    reason = "parser bindings are intentionally refined after validation"
)]
#![expect(
    clippy::shadow_unrelated,
    reason = "local parser names mirror the OOXML role currently being decoded"
)]
use crate::error::{Error, Result};
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

/// One retained element span of `source`, with every namespace declaration it
/// inherits from `source` re-declared on its own root element.
///
/// Until change 0653 the shared markup-compatibility writer repeated every
/// in-scope declaration on every element it emitted, so any span cut out of a
/// processed part happened to resolve on its own. The writer now declares each
/// namespace once, where XML requires it, so a consumer that hands a span out
/// or parses it standalone asks for the inherited declarations here instead.
pub(crate) fn self_contained_element_xml(
    source: &[u8],
    start: u32,
    length: u32,
) -> Result<Vec<u8>> {
    let start = usize::try_from(start)
        .map_err(|_source_error| Error::InvalidFormat("Word XML offset exceeds usize".into()))?;
    let length = usize::try_from(length).map_err(|_source_error| {
        Error::InvalidFormat("Word XML range length exceeds usize".into())
    })?;
    Ok(litchi_ooxml_common::mce::self_contained_fragment(
        source,
        start,
        length,
        &litchi_ooxml_common::mce::Limits::default(),
    )?
    .into_owned())
}

pub(crate) fn word_attribute_value(
    element: &BytesStart<'_>,
    name: &[u8],
    decoder: Decoder,
    resolver: &NamespaceResolver,
) -> Result<Option<String>> {
    let mut value = None;
    for attribute in element.attributes() {
        let attribute = attribute.map_err(|error| Error::Xml(error.to_string()))?;
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
            return Err(Error::InvalidFormat(format!(
                "duplicate Word attribute '{}'",
                String::from_utf8_lossy(name)
            )));
        }
        value = Some(
            attribute
                .decoded_and_normalized_value(XmlVersion::Explicit1_0, decoder)
                .map_err(|error| Error::Xml(error.to_string()))?
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

/// Maximum capture nesting depth accepted by the shared `WordprocessingML`
/// element scanner, matching the hardened settings and mail-merge parsers.
const MAX_SCAN_DEPTH: usize = 128;
/// Maximum number of elements scanned in one document part.
const MAX_SCAN_NODES: usize = 1_000_000;

/// State for the bounded Word element span scanner.
///
/// The state is independent of the XML reader. The normal scanner owns its
/// reader, while the writer's alternative-format path can feed the same
/// range semantics from the reader it already uses for anchor validation.
pub(crate) struct WordElementRangeScanner<'a, 'b> {
    targets: &'a [&'b [u8]],
    fragment_prefix: Option<Option<Vec<u8>>>,
    capture: Option<(usize, usize, usize)>,
    nodes: usize,
    total_depth: usize,
}

#[derive(Clone, Copy)]
pub(crate) enum WordElementScanEvent {
    Start(usize),
    NestedStart,
    Empty(usize),
    End,
    Eof,
    Other,
}

impl<'a, 'b> WordElementRangeScanner<'a, 'b> {
    pub(crate) fn new(targets: &'a [&'b [u8]]) -> Self {
        Self {
            targets,
            fragment_prefix: None,
            capture: None,
            nodes: 0,
            total_depth: 0,
        }
    }

    /// Classify one already-resolved event and update range-side counters.
    ///
    /// The returned event is owned, so a caller using the legacy borrowed
    /// reader can release its event and namespace borrows before converting
    /// the event end offset.
    pub(crate) fn classify(
        &mut self,
        namespace: &ResolveResult<'_>,
        event: &Event<'_>,
    ) -> Result<WordElementScanEvent> {
        if matches!(event, Event::Start(_) | Event::Empty(_)) {
            self.nodes = self.nodes.checked_add(1).ok_or_else(|| {
                Error::InvalidFormat("Word XML element counter overflow".to_string())
            })?;
            if self.nodes > MAX_SCAN_NODES {
                return Err(Error::InvalidFormat(format!(
                    "Word XML exceeds {MAX_SCAN_NODES} elements"
                )));
            }
        }
        // Total nesting is tracked separately from capture depth so deeply
        // nested non-target content is rejected before quick-xml's own
        // namespace resolver overflows (u16).
        if matches!(event, Event::Start(_)) {
            self.total_depth = self
                .total_depth
                .checked_add(1)
                .ok_or_else(|| Error::InvalidFormat("Word XML nesting is too deep".to_string()))?;
            if self.total_depth > MAX_SCAN_DEPTH {
                return Err(Error::InvalidFormat(format!(
                    "Word XML nesting exceeds the {MAX_SCAN_DEPTH} depth limit"
                )));
            }
        }
        if matches!(event, Event::End(_)) {
            self.total_depth = self
                .total_depth
                .checked_sub(1)
                .ok_or_else(|| Error::InvalidFormat("invalid Word XML nesting".to_string()))?;
        }

        if self.fragment_prefix.is_none()
            && let Event::Start(element) = event
            && !matches!(namespace, ResolveResult::Bound(_))
        {
            self.fragment_prefix = Some(
                element
                    .name()
                    .prefix()
                    .map(|prefix| prefix.into_inner().to_vec()),
            );
        }

        Ok(match event {
            Event::Start(_) if self.capture.is_some() => WordElementScanEvent::NestedStart,
            Event::Start(element)
                if is_fragment_word_namespace(namespace, &self.fragment_prefix) =>
            {
                self.targets
                    .iter()
                    .position(|target| element.local_name().as_ref() == *target)
                    .map_or(WordElementScanEvent::Other, WordElementScanEvent::Start)
            },
            Event::Empty(element)
                if self.capture.is_none()
                    && is_fragment_word_namespace(namespace, &self.fragment_prefix) =>
            {
                self.targets
                    .iter()
                    .position(|target| element.local_name().as_ref() == *target)
                    .map_or(WordElementScanEvent::Other, WordElementScanEvent::Empty)
            },
            Event::End(_) if self.capture.is_some() => WordElementScanEvent::End,
            Event::Eof => WordElementScanEvent::Eof,
            Event::Start(_)
            | Event::End(_)
            | Event::Empty(_)
            | Event::Text(_)
            | Event::CData(_)
            | Event::Comment(_)
            | Event::Decl(_)
            | Event::PI(_)
            | Event::DocType(_)
            | Event::GeneralRef(_) => WordElementScanEvent::Other,
        })
    }

    /// Finish a previously classified event after its end offset is known.
    pub(crate) fn finish(
        &mut self,
        event_start: usize,
        event_end: usize,
        scan_event: WordElementScanEvent,
        emit: &mut impl FnMut(usize, u32, u32) -> Result<()>,
    ) -> Result<()> {
        match scan_event {
            WordElementScanEvent::Start(target) => {
                self.capture = Some((target, event_start, 1));
            },
            WordElementScanEvent::NestedStart => {
                let Some((_, _, depth)) = self.capture.as_mut() else {
                    return Err(Error::InvalidFormat(
                        "missing captured Word element".to_string(),
                    ));
                };
                *depth = depth.checked_add(1).ok_or_else(|| {
                    Error::InvalidFormat("Word element nesting is too deep".to_string())
                })?;
                if *depth > MAX_SCAN_DEPTH {
                    return Err(Error::InvalidFormat(format!(
                        "Word element nesting exceeds the {MAX_SCAN_DEPTH} depth limit"
                    )));
                }
            },
            WordElementScanEvent::Empty(target) => {
                emit_word_element_range(target, event_start, event_end, emit)?;
            },
            WordElementScanEvent::End => {
                let Some((_, _, depth)) = self.capture.as_mut() else {
                    return Err(Error::InvalidFormat(
                        "missing captured Word element".to_string(),
                    ));
                };
                *depth = depth.checked_sub(1).ok_or_else(|| {
                    Error::InvalidFormat("invalid Word element nesting".to_string())
                })?;
                if *depth == 0 {
                    let Some((target, start, _)) = self.capture.take() else {
                        return Err(Error::InvalidFormat(
                            "missing captured Word element range".to_string(),
                        ));
                    };
                    emit_word_element_range(target, start, event_end, emit)?;
                }
            },
            WordElementScanEvent::Eof if self.capture.is_some() => {
                return Err(Error::InvalidFormat(
                    "unterminated Word element".to_string(),
                ));
            },
            WordElementScanEvent::Eof | WordElementScanEvent::Other => {},
        }
        Ok(())
    }
}

pub(crate) fn scan_word_element_ranges(
    xml_bytes: &[u8],
    targets: &[&[u8]],
    mut emit: impl FnMut(usize, u32, u32) -> Result<()>,
) -> Result<()> {
    let mut reader = NsReader::from_reader(xml_bytes);
    let mut scanner = WordElementRangeScanner::new(targets);

    loop {
        let event_start = usize::try_from(reader.buffer_position()).map_err(|_source_error| {
            Error::InvalidFormat("Word XML offset does not fit usize".to_string())
        })?;
        let scan_event = {
            let (namespace, event) = reader
                .read_resolved_event()
                .map_err(|error| Error::Xml(error.to_string()))?;
            scanner.classify(&namespace, &event)?
        };
        let event_end = usize::try_from(reader.buffer_position()).map_err(|_source_error| {
            Error::InvalidFormat("Word XML offset does not fit usize".to_string())
        })?;
        let eof = matches!(scan_event, WordElementScanEvent::Eof);
        scanner.finish(event_start, event_end, scan_event, &mut emit)?;
        if eof {
            break;
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
            .map_err(|error| Error::Xml(error.to_string()))?
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
                    Error::InvalidFormat("Word property XML nesting is too deep".into())
                })?;
                let is_word = is_fragment_word_namespace(&namespace, &fragment_prefix);
                if depth == 1 {
                    if saw_root || !is_word || element.local_name().as_ref() != root_name {
                        return Err(Error::InvalidFormat(
                            "Word property XML has an invalid root".into(),
                        ));
                    }
                    saw_root = true;
                } else if depth == 2 && is_word && element.local_name().as_ref() == properties_name
                {
                    if saw_properties {
                        return Err(Error::InvalidFormat(
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
                    Error::InvalidFormat("Word property XML nesting is too deep".into())
                })?;
                let is_word = is_fragment_word_namespace(&namespace, &fragment_prefix);
                if child_depth == 1 {
                    if saw_root || !is_word || element.local_name().as_ref() != root_name {
                        return Err(Error::InvalidFormat(
                            "Word property XML has an invalid root".into(),
                        ));
                    }
                    saw_root = true;
                } else if child_depth == 2
                    && is_word
                    && element.local_name().as_ref() == properties_name
                {
                    if saw_properties {
                        return Err(Error::InvalidFormat(
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
                    Error::InvalidFormat("invalid Word property XML nesting".into())
                })?;
            },
            Event::Eof if depth != 0 => {
                return Err(Error::InvalidFormat(
                    "unterminated Word property XML".into(),
                ));
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
    }

    if !saw_root {
        return Err(Error::InvalidFormat("Word property XML has no root".into()));
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
        return Err(Error::InvalidFormat(format!(
            "duplicate Word property '{}'",
            String::from_utf8_lossy(property_name)
        )));
    }
    let mut value = None;
    for attribute in element.attributes() {
        let attribute = attribute.map_err(|error| Error::Xml(error.to_string()))?;
        if attribute.key.local_name().as_ref() != b"val" {
            continue;
        }
        let (namespace, _) = resolver.resolve_attribute(attribute.key);
        if !is_fragment_word_namespace(&namespace, fragment_prefix) {
            continue;
        }
        if value.is_some() {
            return Err(Error::InvalidFormat(
                "duplicate Word property value attribute".into(),
            ));
        }
        value = Some(
            attribute
                .decoded_and_normalized_value(XmlVersion::Explicit1_0, decoder)
                .map_err(|error| Error::Xml(error.to_string()))?
                .into_owned(),
        );
    }
    *slot = Some(value.ok_or_else(|| {
        Error::InvalidFormat(format!(
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
        return Err(Error::InvalidFormat(format!(
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
        .ok_or_else(|| Error::InvalidFormat("invalid Word element byte range".to_string()))?;
    let start = u32::try_from(start).map_err(|_source_error| {
        Error::InvalidFormat("Word element offset exceeds u32".to_string())
    })?;
    let length = u32::try_from(length).map_err(|_source_error| {
        Error::InvalidFormat("Word element length exceeds u32".to_string())
    })?;
    emit(target, start, length)
}

#[cfg(test)]
mod tests {
    use quick_xml::events::Event;
    use quick_xml::name::ResolveResult;
    use quick_xml::reader::NsReader;

    /// Whether the fragment's root element resolves into a namespace.
    fn root_is_bound(fragment: &[u8]) -> bool {
        let mut reader = NsReader::from_reader(fragment);
        let mut buffer = Vec::new();
        loop {
            match reader.read_event_into(&mut buffer) {
                Ok(Event::Start(ref element) | Event::Empty(ref element)) => {
                    return matches!(
                        reader.resolver().resolve_element(element.name()).0,
                        ResolveResult::Bound(_)
                    );
                },
                Ok(Event::Eof) | Err(_) => return false,
                Ok(_) => {},
            }
            buffer.clear();
        }
    }

    /// Change 0653: a `w:p` span cut out of a real, marker-bearing
    /// `word/document.xml` no longer carries `xmlns:w`, because the shared
    /// markup-compatibility writer stopped repeating every in-scope
    /// declaration on every element. `Paragraph::self_contained_xml` restores
    /// it at the slice boundary, which is what `Paragraph::extensions` and
    /// `Row::extension_ids` now parse.
    #[test]
    fn a_retained_span_resolves_only_after_the_inherited_declarations_return() {
        let package = crate::Package::open("../../test-data/ooxml/docx/table-alignment.docx")
            .expect("a Word-authored fixture opens");
        let document = package.document().expect("the document part is readable");
        let paragraph = document
            .paragraph(0)
            .expect("the first paragraph is readable")
            .expect("the fixture has a paragraph");

        assert!(
            !root_is_bound(paragraph.xml_bytes()),
            "the retained span still carries every in-scope declaration"
        );
        let repaired = paragraph
            .self_contained_xml()
            .expect("the span can be made self-contained");
        assert!(
            root_is_bound(&repaired),
            "the repaired span still does not resolve: {}",
            String::from_utf8_lossy(&repaired[..repaired.len().min(200)])
        );
    }
}
