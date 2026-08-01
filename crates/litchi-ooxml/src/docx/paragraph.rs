use crate::docx::drawing::{DrawingObject, parse_drawing_objects};
use crate::docx::hyperlink::Hyperlink;
use crate::docx::image::{InlineImage, parse_inline_images};
use crate::docx::math::OfficeMath;
use crate::docx::namespace::{
    direct_word_property_value, is_wordprocessing_namespace, normalize_xml_integer,
    scan_word_element_ranges,
};
use crate::docx::numbering::ParagraphNumbering;
use crate::docx::revision::{Revision, parse_revisions};
use crate::docx::smart_tag::SmartTag;
use crate::docx::{ThemeColor, UnderlineStyle};
use crate::error::{OoxmlError, Result};
/// Paragraph and Run structures for Word documents.
use litchi_core::VerticalPosition;
use litchi_core::XmlSlice;
use litchi_ooxml_common::xml::{
    decode_xml_reference, extract_omml_formulas, omml_formula_xml, scan_omml_formula_ranges,
};
use litchi_opc::rel::Relationships;
use quick_xml::events::{BytesStart, Event};
use quick_xml::name::{NamespaceResolver, QName, ResolveResult};
use quick_xml::reader::NsReader;
use quick_xml::{Reader, XmlVersion, encoding::Decoder};
use smallvec::SmallVec;
use std::borrow::Cow;
use std::sync::Arc;

/// Maximum nesting depth accepted when extracting paragraph text, matching
/// the hardened document element scanner.
const MAX_TEXT_SCAN_DEPTH: usize = 128;
/// Maximum number of elements scanned while extracting paragraph text.
const MAX_TEXT_SCAN_NODES: usize = 1_000_000;

fn is_fragment_word_name(
    namespace: &ResolveResult<'_>,
    name: QName<'_>,
    local_name: &[u8],
    fragment_prefix: &Option<Option<Vec<u8>>>,
) -> bool {
    if name.local_name().as_ref() != local_name {
        return false;
    }
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

impl Paragraph {
    /// Return direct paragraph numbering properties, including `numId=0` cancellation.
    pub fn numbering(&self) -> Result<Option<ParagraphNumbering>> {
        Ok(self.list_properties()?.0)
    }

    /// Return the paragraph style identifier from `<w:pPr>`.
    pub fn style_id(&self) -> Result<Option<String>> {
        Ok(self.list_properties()?.1)
    }

    fn list_properties(&self) -> Result<(Option<ParagraphNumbering>, Option<String>)> {
        let mut reader = NsReader::from_reader(self.xml_bytes());
        let mut depth = 0usize;
        let mut word_prefix: Option<Vec<u8>> = None;
        let mut ppr_depth = None;
        let mut numpr_depth = None;
        let mut saw_numpr = false;
        let mut num_id = None;
        let mut level = None;
        let mut style_id = None;

        loop {
            let decoder = reader.decoder();
            let event = reader
                .read_event()
                .map_err(|error| OoxmlError::Xml(error.to_string()))?
                .into_owned();
            match event {
                Event::Start(element) => {
                    depth = depth.checked_add(1).ok_or_else(|| {
                        OoxmlError::InvalidFormat("paragraph XML nesting is too deep".to_owned())
                    })?;
                    if depth == 1 && element.local_name().as_ref() == b"p" {
                        word_prefix = Some(element_prefix(&element));
                    }
                    if !same_word_prefix(&element, word_prefix.as_deref()) {
                        continue;
                    }
                    match element.local_name().as_ref() {
                        b"pPr" if depth == 2 => ppr_depth = Some(depth),
                        b"numPr" if ppr_depth.is_some_and(|value| depth == value + 1) => {
                            if saw_numpr {
                                return Err(OoxmlError::InvalidFormat(
                                    "paragraph has duplicate numPr".to_owned(),
                                ));
                            }
                            saw_numpr = true;
                            numpr_depth = Some(depth);
                        },
                        b"pStyle" if ppr_depth.is_some_and(|value| depth == value + 1) => {
                            set_paragraph_property(
                                &mut style_id,
                                paragraph_attribute(&element, b"val", decoder)?,
                                "pStyle",
                            )?;
                        },
                        b"numId" if numpr_depth.is_some_and(|value| depth == value + 1) => {
                            let raw = paragraph_attribute(&element, b"val", decoder)?;
                            let parsed = raw.parse::<u32>().map_err(|_| {
                                OoxmlError::InvalidFormat(format!(
                                    "invalid paragraph numId '{raw}'"
                                ))
                            })?;
                            set_paragraph_property(&mut num_id, parsed, "numId")?;
                        },
                        b"ilvl" if numpr_depth.is_some_and(|value| depth == value + 1) => {
                            let raw = paragraph_attribute(&element, b"val", decoder)?;
                            let parsed = raw
                                .parse::<u8>()
                                .ok()
                                .filter(|value| *value <= 8)
                                .ok_or_else(|| {
                                    OoxmlError::InvalidFormat(format!(
                                        "invalid paragraph ilvl '{raw}'"
                                    ))
                                })?;
                            set_paragraph_property(&mut level, parsed, "ilvl")?;
                        },
                        _ => {},
                    }
                },
                Event::Empty(element) => {
                    let child_depth = depth.checked_add(1).ok_or_else(|| {
                        OoxmlError::InvalidFormat("paragraph XML nesting is too deep".to_owned())
                    })?;
                    if !same_word_prefix(&element, word_prefix.as_deref()) {
                        continue;
                    }
                    match element.local_name().as_ref() {
                        b"pStyle" if ppr_depth.is_some_and(|value| child_depth == value + 1) => {
                            set_paragraph_property(
                                &mut style_id,
                                paragraph_attribute(&element, b"val", decoder)?,
                                "pStyle",
                            )?;
                        },
                        b"numPr" if ppr_depth.is_some_and(|value| child_depth == value + 1) => {
                            return Err(OoxmlError::InvalidFormat(
                                "paragraph numPr is missing numId".to_owned(),
                            ));
                        },
                        b"numId" if numpr_depth.is_some_and(|value| child_depth == value + 1) => {
                            let raw = paragraph_attribute(&element, b"val", decoder)?;
                            let parsed = raw.parse::<u32>().map_err(|_| {
                                OoxmlError::InvalidFormat(format!(
                                    "invalid paragraph numId '{raw}'"
                                ))
                            })?;
                            set_paragraph_property(&mut num_id, parsed, "numId")?;
                        },
                        b"ilvl" if numpr_depth.is_some_and(|value| child_depth == value + 1) => {
                            let raw = paragraph_attribute(&element, b"val", decoder)?;
                            let parsed = raw
                                .parse::<u8>()
                                .ok()
                                .filter(|value| *value <= 8)
                                .ok_or_else(|| {
                                    OoxmlError::InvalidFormat(format!(
                                        "invalid paragraph ilvl '{raw}'"
                                    ))
                                })?;
                            set_paragraph_property(&mut level, parsed, "ilvl")?;
                        },
                        _ => {},
                    }
                },
                Event::End(element) => {
                    if same_word_prefix_end(&element, word_prefix.as_deref()) {
                        match element.local_name().as_ref() {
                            b"numPr" if numpr_depth == Some(depth) => numpr_depth = None,
                            b"pPr" if ppr_depth == Some(depth) => ppr_depth = None,
                            _ => {},
                        }
                    }
                    depth = depth.checked_sub(1).ok_or_else(|| {
                        OoxmlError::InvalidFormat("invalid paragraph XML nesting".to_owned())
                    })?;
                },
                Event::Eof => break,
                _ => {},
            }
        }
        let numbering = if saw_numpr {
            Some(ParagraphNumbering {
                num_id: num_id.ok_or_else(|| {
                    OoxmlError::InvalidFormat("paragraph numPr is missing numId".to_owned())
                })?,
                level: level.unwrap_or(0),
            })
        } else {
            None
        };
        Ok((numbering, style_id))
    }
}

fn element_prefix(element: &BytesStart<'_>) -> Vec<u8> {
    let name = element.name();
    let raw = name.as_ref();
    raw.iter()
        .position(|byte| *byte == b':')
        .map_or_else(Vec::new, |index| raw[..index].to_vec())
}

fn same_word_prefix(element: &BytesStart<'_>, prefix: Option<&[u8]>) -> bool {
    prefix.is_some_and(|prefix| element_prefix(element) == prefix)
}

fn same_word_prefix_end(element: &quick_xml::events::BytesEnd<'_>, prefix: Option<&[u8]>) -> bool {
    let name = element.name();
    let raw = name.as_ref();
    let end = raw.iter().position(|byte| *byte == b':').unwrap_or(0);
    prefix.is_some_and(|prefix| {
        if end == 0 {
            prefix.is_empty()
        } else {
            &raw[..end] == prefix
        }
    })
}

fn paragraph_attribute(
    element: &BytesStart<'_>,
    name: &[u8],
    decoder: quick_xml::encoding::Decoder,
) -> Result<String> {
    for attribute in element.attributes().with_checks(false) {
        let attribute = attribute.map_err(|error| OoxmlError::Xml(error.to_string()))?;
        if attribute.key.local_name().as_ref() == name {
            return attribute
                .decoded_and_normalized_value(quick_xml::XmlVersion::Implicit1_0, decoder)
                .map(|value| value.into_owned())
                .map_err(|error| OoxmlError::Xml(error.to_string()));
        }
    }
    Err(OoxmlError::InvalidFormat(format!(
        "paragraph property is missing '{}'",
        String::from_utf8_lossy(name)
    )))
}

fn set_paragraph_property<T>(slot: &mut Option<T>, value: T, name: &str) -> Result<()> {
    if slot.is_some() {
        return Err(OoxmlError::InvalidFormat(format!(
            "paragraph has duplicate {name}"
        )));
    }
    *slot = Some(value);
    Ok(())
}

pub(crate) fn extract_word_text(xml_bytes: &[u8]) -> Result<String> {
    let mut reader = NsReader::from_reader(xml_bytes);
    let mut result = String::with_capacity(xml_bytes.len() / 8);
    let mut fragment_prefix: Option<Option<Vec<u8>>> = None;
    let mut depth = 0usize;
    let mut nodes = 0usize;
    let mut text_depth = None;

    loop {
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
            Event::Start(element) => {
                nodes = nodes.checked_add(1).ok_or_else(|| {
                    OoxmlError::InvalidFormat("Word XML element counter overflow".to_string())
                })?;
                if nodes > MAX_TEXT_SCAN_NODES {
                    return Err(OoxmlError::InvalidFormat(format!(
                        "Word XML exceeds {MAX_TEXT_SCAN_NODES} elements"
                    )));
                }
                depth = depth.checked_add(1).ok_or_else(|| {
                    OoxmlError::InvalidFormat("Word XML nesting is too deep".to_string())
                })?;
                if depth > MAX_TEXT_SCAN_DEPTH {
                    return Err(OoxmlError::InvalidFormat(format!(
                        "Word XML nesting exceeds the {MAX_TEXT_SCAN_DEPTH} depth limit"
                    )));
                }
                if text_depth.is_none()
                    && is_fragment_word_name(&namespace, element.name(), b"t", &fragment_prefix)
                {
                    text_depth = Some(depth);
                } else if let Some(character) =
                    word_special_character(&namespace, element.name(), &fragment_prefix)
                {
                    result.push(character);
                }
            },
            Event::Empty(element) => {
                nodes = nodes.checked_add(1).ok_or_else(|| {
                    OoxmlError::InvalidFormat("Word XML element counter overflow".to_string())
                })?;
                if nodes > MAX_TEXT_SCAN_NODES {
                    return Err(OoxmlError::InvalidFormat(format!(
                        "Word XML exceeds {MAX_TEXT_SCAN_NODES} elements"
                    )));
                }
                if let Some(character) =
                    word_special_character(&namespace, element.name(), &fragment_prefix)
                {
                    result.push(character);
                }
            },
            Event::Text(text) if text_depth.is_some() => {
                let decoded = text
                    .xml_content(XmlVersion::Explicit1_0)
                    .map_err(|error| OoxmlError::Xml(error.to_string()))?;
                let unescaped = quick_xml::escape::unescape(&decoded)
                    .map_err(|error| OoxmlError::Xml(error.to_string()))?;
                result.push_str(&unescaped);
            },
            Event::CData(text) if text_depth.is_some() => {
                let decoded = text
                    .xml_content(XmlVersion::Explicit1_0)
                    .map_err(|error| OoxmlError::Xml(error.to_string()))?;
                result.push_str(&decoded);
            },
            Event::GeneralRef(reference) if text_depth.is_some() => {
                result.push_str(&decode_xml_reference(&reference)?);
            },
            Event::End(element) => {
                if text_depth == Some(depth)
                    && is_fragment_word_name(&namespace, element.name(), b"t", &fragment_prefix)
                {
                    text_depth = None;
                }
                depth = depth.checked_sub(1).ok_or_else(|| {
                    OoxmlError::InvalidFormat("invalid Word XML nesting".to_string())
                })?;
            },
            Event::Eof if depth != 0 || text_depth.is_some() => {
                return Err(OoxmlError::InvalidFormat(
                    "unterminated Word XML".to_string(),
                ));
            },
            Event::Eof => break,
            _ => {},
        }
    }
    result.shrink_to_fit();
    Ok(result)
}

fn word_special_character(
    namespace: &ResolveResult<'_>,
    name: QName<'_>,
    fragment_prefix: &Option<Option<Vec<u8>>>,
) -> Option<char> {
    if is_fragment_word_name(namespace, name, b"tab", fragment_prefix) {
        Some('\t')
    } else if is_fragment_word_name(namespace, name, b"br", fragment_prefix)
        || is_fragment_word_name(namespace, name, b"cr", fragment_prefix)
    {
        Some('\n')
    } else if is_fragment_word_name(namespace, name, b"noBreakHyphen", fragment_prefix) {
        Some('\u{2011}')
    } else if is_fragment_word_name(namespace, name, b"softHyphen", fragment_prefix) {
        Some('\u{00ad}')
    } else {
        None
    }
}

/// A paragraph in a Word document.
///
/// Represents a `<w:p>` element. Paragraphs contain runs which in turn
/// contain the actual text and formatting.
///
/// # Example
///
/// ```rust,ignore
/// for para in document.paragraphs()? {
///     println!("Paragraph text: {}", para.text());
///     for run in para.runs()? {
///         println!("  Run: {} (bold: {:?})", run.text(), run.bold());
///     }
/// }
/// ```
/// Internal storage for paragraph XML data.
/// Supports both owned data (for standalone parsing) and shared slices (for arena-based parsing).
#[derive(Debug, Clone)]
enum XmlData {
    /// Owned data for standalone paragraphs
    Owned(Box<[u8]>),
    /// Shared slice into an arena for zero-copy batch parsing
    Shared(XmlSlice),
}

impl XmlData {
    #[inline]
    fn as_bytes(&self) -> &[u8] {
        match self {
            XmlData::Owned(b) => b,
            XmlData::Shared(s) => s.as_bytes(),
        }
    }

    /// Get or create an Arc for this data.
    /// If already shared, returns the existing Arc (cheap clone).
    /// If owned, creates a new Arc (allocates once).
    #[inline]
    fn get_or_create_arc(&self) -> (Arc<Vec<u8>>, u32) {
        match self {
            XmlData::Owned(b) => (Arc::new(b.to_vec()), 0),
            XmlData::Shared(s) => (s.arc(), s.start()),
        }
    }
}

#[derive(Debug, Clone)]
pub struct Paragraph {
    /// The raw XML bytes for this paragraph
    xml_data: XmlData,
}

impl Paragraph {
    /// Create a new Paragraph from XML bytes (owned).
    ///
    /// # Arguments
    ///
    /// * `xml_bytes` - The XML content of the `<w:p>` element
    #[inline]
    pub fn new(xml_bytes: Vec<u8>) -> Self {
        Self {
            xml_data: XmlData::Owned(xml_bytes.into_boxed_slice()),
        }
    }

    /// Create a new Paragraph from a shared XML slice (zero-copy).
    ///
    /// This is used for arena-based parsing where all element XMLs are stored
    /// in a single contiguous buffer.
    #[inline]
    pub fn from_slice(slice: XmlSlice) -> Self {
        Self {
            xml_data: XmlData::Shared(slice),
        }
    }

    /// Create a Paragraph from an `Arc<Vec<u8>>` and byte range.
    ///
    /// This is a convenience method for arena-based parsing.
    #[inline]
    pub fn from_arc_range(arena: Arc<Vec<u8>>, start: u32, len: u32) -> Self {
        Self::from_slice(XmlSlice::new(arena, start, len))
    }

    /// Get the raw XML bytes.
    #[inline]
    fn xml_bytes(&self) -> &[u8] {
        self.xml_data.as_bytes()
    }

    /// Get the text content of this paragraph.
    ///
    /// Concatenates all text from all runs in the paragraph.
    ///
    /// # Performance
    ///
    /// Uses streaming XML parsing with pre-allocated buffer to extract text efficiently.
    pub fn text(&self) -> Result<String> {
        extract_word_text(self.xml_bytes())
    }

    /// Return the HTML division ID referenced by this paragraph, if present.
    pub fn division_id(&self) -> Result<Option<String>> {
        direct_word_property_value(self.xml_bytes(), b"p", b"pPr", b"divId")?
            .map(|value| normalize_xml_integer(value, "Word paragraph division ID"))
            .transpose()
    }

    /// Get an iterator over the runs in this paragraph.
    ///
    /// Each run represents a `<w:r>` element and may have different formatting.
    ///
    /// # Performance
    ///
    /// Uses namespace-aware streaming boundary detection and shared XML slices.
    pub fn runs(&self) -> Result<SmallVec<[Run; 8]>> {
        enum RunEvent {
            Start,
            NestedStart,
            Empty,
            End,
            Eof,
            Other,
        }

        let xml_bytes = self.xml_bytes();
        let (source_arc, base_offset) = self.xml_data.get_or_create_arc();
        let mut reader = NsReader::from_reader(xml_bytes);
        let mut runs = SmallVec::new();
        let mut run_start = None;
        let mut run_depth = 0usize;
        let mut fragment_prefix: Option<Option<Vec<u8>>> = None;

        loop {
            let event_start = usize::try_from(reader.buffer_position()).map_err(|_| {
                OoxmlError::InvalidFormat("Word paragraph offset does not fit usize".to_string())
            })?;
            let event = {
                let (namespace, event) = reader
                    .read_resolved_event()
                    .map_err(|error| OoxmlError::Xml(error.to_string()))?;
                match event {
                    Event::Start(ref element)
                        if fragment_prefix.is_none()
                            && element.local_name().as_ref() == b"p"
                            && !matches!(namespace, ResolveResult::Bound(_)) =>
                    {
                        fragment_prefix = Some(
                            element
                                .name()
                                .prefix()
                                .map(|prefix| prefix.into_inner().to_vec()),
                        );
                        RunEvent::Other
                    },
                    Event::Start(_) if run_start.is_some() => RunEvent::NestedStart,
                    Event::Start(element)
                        if is_fragment_word_name(
                            &namespace,
                            element.name(),
                            b"r",
                            &fragment_prefix,
                        ) =>
                    {
                        RunEvent::Start
                    },
                    Event::Empty(element)
                        if run_start.is_none()
                            && is_fragment_word_name(
                                &namespace,
                                element.name(),
                                b"r",
                                &fragment_prefix,
                            ) =>
                    {
                        RunEvent::Empty
                    },
                    Event::End(_) if run_start.is_some() => RunEvent::End,
                    Event::Eof => RunEvent::Eof,
                    _ => RunEvent::Other,
                }
            };
            let event_end = usize::try_from(reader.buffer_position()).map_err(|_| {
                OoxmlError::InvalidFormat("Word paragraph offset does not fit usize".to_string())
            })?;

            match event {
                RunEvent::Start => {
                    run_start = Some(event_start);
                    run_depth = 1;
                },
                RunEvent::NestedStart => {
                    run_depth = run_depth.checked_add(1).ok_or_else(|| {
                        OoxmlError::InvalidFormat("Word run nesting is too deep".to_string())
                    })?;
                },
                RunEvent::Empty => {
                    Self::push_run_slice(
                        &mut runs,
                        &source_arc,
                        base_offset,
                        event_start,
                        event_end,
                    )?;
                },
                RunEvent::End => {
                    run_depth = run_depth.checked_sub(1).ok_or_else(|| {
                        OoxmlError::InvalidFormat("invalid Word run nesting".to_string())
                    })?;
                    if run_depth == 0 {
                        let Some(start) = run_start.take() else {
                            return Err(OoxmlError::InvalidFormat(
                                "missing Word run start offset".to_string(),
                            ));
                        };
                        Self::push_run_slice(
                            &mut runs,
                            &source_arc,
                            base_offset,
                            start,
                            event_end,
                        )?;
                    }
                },
                RunEvent::Eof if run_start.is_some() => {
                    return Err(OoxmlError::InvalidFormat(
                        "unterminated Word run".to_string(),
                    ));
                },
                RunEvent::Eof => break,
                _ => {},
            }
        }

        Ok(runs)
    }

    /// Return all run-level smart tags in document order, including nested tags.
    pub fn smart_tags(&self) -> Result<Vec<SmartTag>> {
        enum SmartTagEvent {
            Start(bool),
            Empty(bool),
            End(bool),
            Eof,
            Other,
        }

        let xml_bytes = self.xml_bytes();
        let (source_arc, base_offset) = self.xml_data.get_or_create_arc();
        let mut reader = NsReader::from_reader(xml_bytes);
        let mut fragment_prefix: Option<Option<Vec<u8>>> = None;
        let mut depth = 0usize;
        let mut open_tags = Vec::new();
        let mut ranges = Vec::new();

        loop {
            let event_start = usize::try_from(reader.buffer_position()).map_err(|_| {
                OoxmlError::InvalidFormat("Word smart-tag offset does not fit usize".into())
            })?;
            let event = {
                let (namespace, event) = reader
                    .read_resolved_event()
                    .map_err(|error| OoxmlError::Xml(error.to_string()))?;

                if fragment_prefix.is_none()
                    && let Event::Start(element) = &event
                    && element.local_name().as_ref() == b"p"
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
                    Event::Start(element) => SmartTagEvent::Start(is_fragment_word_name(
                        &namespace,
                        element.name(),
                        b"smartTag",
                        &fragment_prefix,
                    )),
                    Event::Empty(element) => SmartTagEvent::Empty(is_fragment_word_name(
                        &namespace,
                        element.name(),
                        b"smartTag",
                        &fragment_prefix,
                    )),
                    Event::End(element) => SmartTagEvent::End(is_fragment_word_name(
                        &namespace,
                        element.name(),
                        b"smartTag",
                        &fragment_prefix,
                    )),
                    Event::Eof => SmartTagEvent::Eof,
                    _ => SmartTagEvent::Other,
                }
            };
            let event_end = usize::try_from(reader.buffer_position()).map_err(|_| {
                OoxmlError::InvalidFormat("Word smart-tag offset does not fit usize".into())
            })?;

            match event {
                SmartTagEvent::Start(is_smart_tag) => {
                    depth = depth.checked_add(1).ok_or_else(|| {
                        OoxmlError::InvalidFormat("Word XML nesting is too deep".into())
                    })?;
                    if is_smart_tag {
                        open_tags.push((event_start, depth));
                    }
                },
                SmartTagEvent::Empty(true) => {
                    ranges.push((event_start, event_end));
                },
                SmartTagEvent::End(is_smart_tag) => {
                    if is_smart_tag {
                        let Some((start, tag_depth)) = open_tags.pop() else {
                            return Err(OoxmlError::InvalidFormat(
                                "Word smart tag has no opening element".into(),
                            ));
                        };
                        if tag_depth != depth {
                            return Err(OoxmlError::InvalidFormat(
                                "invalid nested Word smart tag".into(),
                            ));
                        }
                        ranges.push((start, event_end));
                    }
                    depth = depth.checked_sub(1).ok_or_else(|| {
                        OoxmlError::InvalidFormat("invalid Word XML nesting".into())
                    })?;
                },
                SmartTagEvent::Eof if !open_tags.is_empty() || depth != 0 => {
                    return Err(OoxmlError::InvalidFormat(
                        "unterminated Word smart-tag XML".into(),
                    ));
                },
                SmartTagEvent::Eof => break,
                _ => {},
            }
        }

        ranges.sort_unstable_by_key(|&(start, _)| start);
        ranges
            .into_iter()
            .map(|(start, end)| {
                let start = u32::try_from(start).map_err(|_| {
                    OoxmlError::InvalidFormat("Word smart-tag offset exceeds u32".into())
                })?;
                let length = u32::try_from(end.checked_sub(start as usize).ok_or_else(|| {
                    OoxmlError::InvalidFormat("invalid Word smart-tag byte range".into())
                })?)
                .map_err(|_| {
                    OoxmlError::InvalidFormat("Word smart-tag length exceeds u32".into())
                })?;
                let absolute_start = base_offset.checked_add(start).ok_or_else(|| {
                    OoxmlError::InvalidFormat("Word smart-tag absolute offset exceeds u32".into())
                })?;
                SmartTag::parse(XmlSlice::new(
                    Arc::clone(&source_arc),
                    absolute_start,
                    length,
                ))
            })
            .collect()
    }

    fn push_run_slice(
        runs: &mut SmallVec<[Run; 8]>,
        source: &Arc<Vec<u8>>,
        base_offset: u32,
        start: usize,
        end: usize,
    ) -> Result<()> {
        let start = u32::try_from(start)
            .map_err(|_| OoxmlError::InvalidFormat("Word run offset exceeds u32".to_string()))?;
        let length =
            u32::try_from(end.checked_sub(start as usize).ok_or_else(|| {
                OoxmlError::InvalidFormat("invalid Word run byte range".to_string())
            })?)
            .map_err(|_| OoxmlError::InvalidFormat("Word run length exceeds u32".to_string()))?;
        let absolute_start = base_offset.checked_add(start).ok_or_else(|| {
            OoxmlError::InvalidFormat("Word run absolute offset exceeds u32".to_string())
        })?;
        runs.push(Run::from_slice(XmlSlice::new(
            Arc::clone(source),
            absolute_start,
            length,
        )));
        Ok(())
    }

    /// Extract all OMML formulas from this paragraph.
    ///
    /// Returns a vector of OMML formula strings found in any run within this paragraph.
    /// This extracts inline formulas (formulas within runs).
    pub fn omml_formulas(&self) -> Result<Vec<String>> {
        let mut run_ranges = Vec::new();
        scan_word_element_ranges(self.xml_bytes(), &[b"r".as_slice()], |_, start, length| {
            let start = start as usize;
            let end = start
                .checked_add(length as usize)
                .ok_or_else(|| OoxmlError::InvalidFormat("Word run range overflows".to_string()))?;
            run_ranges.push((start, end));
            Ok(())
        })?;

        let mut formulas = Vec::new();
        scan_omml_formula_ranges(self.xml_bytes(), |start, length| {
            let formula_start = start as usize;
            let formula_end = formula_start.checked_add(length as usize).ok_or_else(|| {
                OoxmlError::InvalidFormat("OMML formula range overflows".to_string())
            })?;
            let is_inline = run_ranges
                .iter()
                .any(|&(run_start, run_end)| formula_start >= run_start && formula_end <= run_end);
            if is_inline {
                formulas.push(omml_formula_xml(self.xml_bytes(), start, length)?);
            }
            Ok::<(), OoxmlError>(())
        })?;
        Ok(formulas)
    }

    /// Extract inline Office Math equations as validated typed fragments.
    ///
    /// Inline equations are `<m:oMath>` elements nested in Word runs.  Their
    /// exact raw XML remains available through [`Self::omml_formulas`]; this
    /// method turns each fragment into a validated [`OfficeMath`] value
    /// suitable for reuse with the mutable writer.
    pub fn inline_office_math(&self) -> Result<Vec<OfficeMath>> {
        self.omml_formulas()?
            .into_iter()
            .map(OfficeMath::from_xml)
            .collect()
    }

    /// Extract all inline images from this paragraph.
    ///
    /// Returns a vector of `InlineImage` objects found in `<w:drawing>` elements
    /// within this paragraph.
    ///
    /// # Example
    ///
    /// ```rust,ignore
    /// for para in document.paragraphs()? {
    ///     for image in para.images()? {
    ///         println!("Image: {} ({}x{} pixels)",
    ///             image.name(),
    ///             image.width_px(),
    ///             image.height_px()
    ///         );
    ///     }
    /// }
    /// ```
    #[inline]
    pub fn images(&self) -> Result<SmallVec<[InlineImage; 4]>> {
        parse_inline_images(self.xml_bytes())
    }

    /// Extract all drawing objects (shapes, text boxes) from this paragraph.
    ///
    /// Returns a vector of `DrawingObject` objects found in `<w:drawing>` elements
    /// within this paragraph. This includes shapes, text boxes, and other DrawingML objects.
    ///
    /// # Example
    ///
    /// ```rust,ignore
    /// for para in document.paragraphs()? {
    ///     for drawing in para.drawing_objects()? {
    ///         println!("Shape: {} (type: {:?})",
    ///             drawing.name(),
    ///             drawing.shape_type()
    ///         );
    ///         if !drawing.text().is_empty() {
    ///             println!("  Text: {}", drawing.text());
    ///         }
    ///     }
    /// }
    /// ```
    #[inline]
    pub fn drawing_objects(&self) -> Result<SmallVec<[DrawingObject; 4]>> {
        parse_drawing_objects(self.xml_bytes())
    }

    /// Extract all tracked changes (revisions) from this paragraph.
    ///
    /// Returns a vector of `Revision` objects representing all tracked changes
    /// (insertions, deletions, moves) within this paragraph.
    ///
    /// # Example
    ///
    /// ```rust,ignore
    /// for para in document.paragraphs()? {
    ///     for revision in para.revisions()? {
    ///         println!("Revision by {}: {} - {}",
    ///             revision.author(),
    ///             revision.revision_type(),
    ///             revision.text()
    ///         );
    ///     }
    /// }
    /// ```
    #[inline]
    pub fn revisions(&self) -> Result<SmallVec<[Revision; 4]>> {
        parse_revisions(self.xml_bytes())
    }

    /// Extract paragraph-level OMML formulas.
    ///
    /// Returns a vector of OMML formula strings that are direct children of the paragraph
    /// (display math), not nested within runs. These are block-level formulas.
    ///
    /// # Example
    /// ```ignore
    /// let para = document.paragraphs()?[0];
    /// let display_formulas = para.paragraph_level_formulas()?;
    /// for formula in display_formulas {
    ///     println!("Display formula: {}", formula);
    /// }
    /// ```
    pub fn paragraph_level_formulas(&self) -> Result<Vec<String>> {
        let mut run_ranges = Vec::new();
        scan_word_element_ranges(self.xml_bytes(), &[b"r".as_slice()], |_, start, length| {
            let start = start as usize;
            let end = start
                .checked_add(length as usize)
                .ok_or_else(|| OoxmlError::InvalidFormat("Word run range overflows".to_string()))?;
            run_ranges.push((start, end));
            Ok(())
        })?;

        let mut formulas = Vec::new();
        scan_omml_formula_ranges(self.xml_bytes(), |start, length| {
            let formula_start = start as usize;
            let formula_end = formula_start.checked_add(length as usize).ok_or_else(|| {
                OoxmlError::InvalidFormat("OMML formula range overflows".to_string())
            })?;
            let is_inline = run_ranges
                .iter()
                .any(|&(run_start, run_end)| formula_start >= run_start && formula_end <= run_end);
            if !is_inline {
                formulas.push(omml_formula_xml(self.xml_bytes(), start, length)?);
            }
            Ok::<(), OoxmlError>(())
        })?;
        Ok(formulas)
    }

    /// Extract display Office Math equations as validated typed fragments.
    ///
    /// Display equations are `<m:oMath>` elements outside Word runs, normally
    /// enclosed by an `<m:oMathPara>` math paragraph.  The result is flattened
    /// into document order; use [`Self::paragraph_level_formulas`] when the
    /// original XML strings are required.
    pub fn display_office_math(&self) -> Result<Vec<OfficeMath>> {
        self.paragraph_level_formulas()?
            .into_iter()
            .map(OfficeMath::from_xml)
            .collect()
    }
    /// Get all hyperlinks in this paragraph.
    ///
    /// Returns a vector of `Hyperlink` objects representing all hyperlinks
    /// found in this paragraph. Requires relationships to resolve external URLs.
    ///
    /// # Arguments
    ///
    /// * `rels` - Relationships for resolving relationship IDs to URLs
    ///
    /// # Examples
    ///
    /// ```rust,ignore
    /// let para = doc.paragraph(0)?.unwrap();
    /// let hyperlinks = para.hyperlinks(&main_part.rels())?;
    /// for link in hyperlinks {
    ///     println!("Link: {} -> {:?}", link.text(), link.url());
    /// }
    /// ```
    pub fn hyperlinks(&self, rels: &Relationships) -> Result<Vec<Hyperlink>> {
        Hyperlink::extract_from_paragraph(self.xml_bytes(), rels)
    }
}

/// A run within a paragraph.
///
/// Represents a `<w:r>` element. A run is a region of text with a single
/// set of formatting properties.
///
/// # Example
///
/// ```rust,ignore
/// let run = runs[0];
/// println!("Text: {}", run.text()?);
/// println!("Bold: {:?}", run.bold()?);
/// println!("Italic: {:?}", run.italic()?);
///
/// // Check for embedded formulas
/// if let Some(omml) = run.omml_formula()? {
///     println!("OMML formula: {}", omml);
/// }
/// ```
/// Cached formatting properties for a Run.
///
/// This struct stores all commonly accessed formatting properties
/// to avoid repeated XML parsing.
#[derive(Debug, Clone, Copy, Default)]
pub struct RunProperties {
    /// Whether the run is bold
    pub bold: Option<bool>,
    /// Whether the run is italic
    pub italic: Option<bool>,
    /// Whether the run is strikethrough
    pub strikethrough: Option<bool>,
    /// Explicit underline pattern, including [`UnderlineStyle::None`]
    pub underline: Option<UnderlineStyle>,
    /// Vertical position (superscript/subscript)
    pub vertical_position: Option<VerticalPosition>,
}

fn update_run_properties(props: &mut RunProperties, element: &BytesStart<'_>) -> Result<()> {
    let property = element.local_name();
    if !matches!(
        property.as_ref(),
        b"b" | b"i" | b"strike" | b"u" | b"vertAlign"
    ) {
        return Ok(());
    }

    let mut value = None;
    for attribute in element.attributes() {
        let attribute = attribute.map_err(|error| OoxmlError::Xml(error.to_string()))?;
        if attribute.key.local_name().as_ref() == b"val" {
            value = Some(attribute.value);
            break;
        }
    }

    match property.as_ref() {
        b"b" => props.bold = Some(value.as_deref().is_none_or(is_on)),
        b"i" => props.italic = Some(value.as_deref().is_none_or(is_on)),
        b"strike" => props.strikethrough = Some(value.as_deref().is_none_or(is_on)),
        b"u" => {
            props.underline = Some(match value.as_deref() {
                None => UnderlineStyle::Single,
                Some(value) => UnderlineStyle::from_xml(
                    std::str::from_utf8(value)
                        .map_err(|error| OoxmlError::InvalidFormat(error.to_string()))?,
                )
                .ok_or_else(|| {
                    OoxmlError::InvalidFormat(format!(
                        "invalid Word underline style '{}'",
                        String::from_utf8_lossy(value)
                    ))
                })?,
            });
        },
        b"vertAlign" => {
            props.vertical_position = match value.as_deref() {
                Some(b"superscript") => Some(VerticalPosition::Superscript),
                Some(b"subscript") => Some(VerticalPosition::Subscript),
                _ => None,
            };
        },
        _ => {},
    }
    Ok(())
}

/// A direct color applied to a WordprocessingML underline.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum RunUnderlineColor {
    /// Automatic color selected by the consumer.
    Auto,
    /// Explicit red, green, and blue components.
    Rgb([u8; 3]),
}

/// Complete direct underline formatting for a run.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct RunUnderline {
    /// Underline pattern, including the explicit `none` value.
    pub style: UnderlineStyle,
    /// Direct automatic or RGB color.
    pub color: Option<RunUnderlineColor>,
    /// Theme color used instead of, or to transform, the direct color.
    pub theme_color: Option<ThemeColor>,
    /// Theme tint transform byte.
    pub theme_tint: Option<u8>,
    /// Theme shade transform byte.
    pub theme_shade: Option<u8>,
}

fn parse_run_underline(xml_bytes: &[u8]) -> Result<Option<RunUnderline>> {
    let mut reader = NsReader::from_reader(xml_bytes);
    let mut fragment_prefix: Option<Option<Vec<u8>>> = None;
    let mut depth = 0usize;
    let mut properties_depth = None;
    let mut saw_root = false;
    let mut saw_properties = false;
    let mut underline = None;

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
                    OoxmlError::InvalidFormat("Word run XML nesting is too deep".into())
                })?;
                let is_word = is_fragment_word_name(
                    &namespace,
                    element.name(),
                    element.local_name().as_ref(),
                    &fragment_prefix,
                );
                if depth == 1 {
                    if saw_root || !is_word || element.local_name().as_ref() != b"r" {
                        return Err(OoxmlError::InvalidFormat(
                            "Word underline XML has an invalid run root".into(),
                        ));
                    }
                    saw_root = true;
                } else if depth == 2 && is_word && element.local_name().as_ref() == b"rPr" {
                    if saw_properties {
                        return Err(OoxmlError::InvalidFormat(
                            "duplicate Word run property container".into(),
                        ));
                    }
                    saw_properties = true;
                    properties_depth = Some(depth);
                } else if depth == 3
                    && properties_depth == Some(2)
                    && is_word
                    && element.local_name().as_ref() == b"u"
                {
                    set_run_underline(
                        &mut underline,
                        &element,
                        decoder,
                        &resolver,
                        &fragment_prefix,
                    )?;
                }
            },
            Event::Empty(element) => {
                let child_depth = depth.checked_add(1).ok_or_else(|| {
                    OoxmlError::InvalidFormat("Word run XML nesting is too deep".into())
                })?;
                let is_word = is_fragment_word_name(
                    &namespace,
                    element.name(),
                    element.local_name().as_ref(),
                    &fragment_prefix,
                );
                if child_depth == 1 {
                    if saw_root || !is_word || element.local_name().as_ref() != b"r" {
                        return Err(OoxmlError::InvalidFormat(
                            "Word underline XML has an invalid run root".into(),
                        ));
                    }
                    saw_root = true;
                } else if child_depth == 2 && is_word && element.local_name().as_ref() == b"rPr" {
                    if saw_properties {
                        return Err(OoxmlError::InvalidFormat(
                            "duplicate Word run property container".into(),
                        ));
                    }
                    saw_properties = true;
                } else if child_depth == 3
                    && properties_depth == Some(2)
                    && is_word
                    && element.local_name().as_ref() == b"u"
                {
                    set_run_underline(
                        &mut underline,
                        &element,
                        decoder,
                        &resolver,
                        &fragment_prefix,
                    )?;
                }
            },
            Event::End(_) => {
                if properties_depth == Some(depth) {
                    properties_depth = None;
                }
                depth = depth.checked_sub(1).ok_or_else(|| {
                    OoxmlError::InvalidFormat("invalid Word run XML nesting".into())
                })?;
            },
            Event::Eof if depth != 0 => {
                return Err(OoxmlError::InvalidFormat(
                    "unterminated Word run XML".into(),
                ));
            },
            Event::Eof => break,
            _ => {},
        }
    }

    if !saw_root {
        return Err(OoxmlError::InvalidFormat(
            "Word underline XML has no run root".into(),
        ));
    }
    Ok(underline)
}

fn set_run_underline(
    slot: &mut Option<RunUnderline>,
    element: &BytesStart<'_>,
    decoder: Decoder,
    resolver: &NamespaceResolver,
    fragment_prefix: &Option<Option<Vec<u8>>>,
) -> Result<()> {
    if slot.is_some() {
        return Err(OoxmlError::InvalidFormat(
            "duplicate Word underline property".into(),
        ));
    }
    let style = run_underline_attribute(element, b"val", decoder, resolver, fragment_prefix)?
        .map(|value| {
            UnderlineStyle::from_xml(&value).ok_or_else(|| {
                OoxmlError::InvalidFormat(format!("invalid Word underline style '{value}'"))
            })
        })
        .transpose()?
        .unwrap_or(UnderlineStyle::Single);
    let color = run_underline_attribute(element, b"color", decoder, resolver, fragment_prefix)?
        .map(|value| parse_run_underline_color(&value))
        .transpose()?;
    let theme_color =
        run_underline_attribute(element, b"themeColor", decoder, resolver, fragment_prefix)?
            .map(|value| {
                ThemeColor::from_xml(&value).ok_or_else(|| {
                    OoxmlError::InvalidFormat(format!(
                        "invalid Word underline theme color '{value}'"
                    ))
                })
            })
            .transpose()?;
    let theme_tint =
        run_underline_attribute(element, b"themeTint", decoder, resolver, fragment_prefix)?
            .map(|value| parse_run_underline_hex_byte(&value, "theme tint"))
            .transpose()?;
    let theme_shade =
        run_underline_attribute(element, b"themeShade", decoder, resolver, fragment_prefix)?
            .map(|value| parse_run_underline_hex_byte(&value, "theme shade"))
            .transpose()?;

    *slot = Some(RunUnderline {
        style,
        color,
        theme_color,
        theme_tint,
        theme_shade,
    });
    Ok(())
}

fn run_underline_attribute(
    element: &BytesStart<'_>,
    name: &[u8],
    decoder: Decoder,
    resolver: &NamespaceResolver,
    fragment_prefix: &Option<Option<Vec<u8>>>,
) -> Result<Option<String>> {
    let mut value = None;
    for attribute in element.attributes() {
        let attribute = attribute.map_err(|error| OoxmlError::Xml(error.to_string()))?;
        let (namespace, _) = resolver.resolve_attribute(attribute.key);
        if !is_fragment_word_name(&namespace, attribute.key, name, fragment_prefix) {
            continue;
        }
        if value.is_some() {
            return Err(OoxmlError::InvalidFormat(format!(
                "duplicate Word underline attribute '{}'",
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

fn parse_run_underline_color(value: &str) -> Result<RunUnderlineColor> {
    if value == "auto" {
        return Ok(RunUnderlineColor::Auto);
    }
    if value.len() != 6 || !value.bytes().all(|byte| byte.is_ascii_hexdigit()) {
        return Err(OoxmlError::InvalidFormat(format!(
            "invalid Word underline color '{value}'"
        )));
    }
    let mut rgb = [0u8; 3];
    for (index, component) in rgb.iter_mut().enumerate() {
        *component = u8::from_str_radix(&value[index * 2..index * 2 + 2], 16).map_err(|_| {
            OoxmlError::InvalidFormat(format!("invalid Word underline color '{value}'"))
        })?;
    }
    Ok(RunUnderlineColor::Rgb(rgb))
}

fn parse_run_underline_hex_byte(value: &str, description: &str) -> Result<u8> {
    if value.len() != 2 || !value.bytes().all(|byte| byte.is_ascii_hexdigit()) {
        return Err(OoxmlError::InvalidFormat(format!(
            "invalid Word underline {description} '{value}'"
        )));
    }
    u8::from_str_radix(value, 16).map_err(|_| {
        OoxmlError::InvalidFormat(format!("invalid Word underline {description} '{value}'"))
    })
}

#[inline]
fn is_on(value: &[u8]) -> bool {
    matches!(value, b"true" | b"1" | b"on")
}

/// Internal storage for run XML data (same pattern as Paragraph).
#[derive(Debug, Clone)]
enum RunXmlData {
    Owned(Vec<u8>),
    Shared(XmlSlice),
}

impl RunXmlData {
    #[inline]
    fn as_bytes(&self) -> &[u8] {
        match self {
            RunXmlData::Owned(v) => v,
            RunXmlData::Shared(s) => s.as_bytes(),
        }
    }
}

#[derive(Debug, Clone)]
pub struct Run {
    /// The raw XML data for this run
    xml_data: RunXmlData,
}

/// The semantic type of an explicit WordprocessingML run break.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub enum RunBreakType {
    /// A normal line break within the current text flow.
    #[default]
    TextWrapping,
    /// A page break.
    Page,
    /// A column break.
    Column,
}

/// How text wrapping resumes after a line break around floating objects.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub enum RunBreakClear {
    /// Resume on the next line without clearing either side.
    #[default]
    None,
    /// Resume when the left side is clear.
    Left,
    /// Resume when the right side is clear.
    Right,
    /// Resume when both sides are clear.
    All,
}

/// A typed `<w:br>` element contained in a run.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub struct RunBreak {
    /// Break type; omitted `w:type` defaults to text wrapping.
    pub break_type: RunBreakType,
    /// Wrapping-clear behavior; omitted `w:clear` defaults to none.
    pub clear: RunBreakClear,
}

impl Run {
    /// Create a new Run from XML bytes (owned).
    pub fn new(xml_bytes: Vec<u8>) -> Self {
        Self {
            xml_data: RunXmlData::Owned(xml_bytes),
        }
    }

    /// Create a Run from a shared XML slice (zero-copy).
    #[inline]
    pub fn from_slice(slice: XmlSlice) -> Self {
        Self {
            xml_data: RunXmlData::Shared(slice),
        }
    }

    /// Get the raw XML bytes.
    #[inline]
    fn xml_bytes(&self) -> &[u8] {
        self.xml_data.as_bytes()
    }

    /// Get the text content of this run.
    ///
    /// Extracts text from `<w:t>` elements and converts special characters:
    /// - `<w:tab/>` → tab character
    /// - `<w:br/>` → newline character
    pub fn text(&self) -> Result<String> {
        extract_word_text(self.xml_bytes())
    }

    /// Parse all explicit break elements in this run, preserving type and clear behavior.
    pub fn breaks(&self) -> Result<SmallVec<[RunBreak; 2]>> {
        let mut reader = Reader::from_reader(self.xml_bytes());
        reader.config_mut().trim_text(true);
        let mut breaks = SmallVec::new();
        loop {
            match reader.read_event() {
                Ok(Event::Start(e)) | Ok(Event::Empty(e)) if e.local_name().as_ref() == b"br" => {
                    let mut run_break = RunBreak::default();
                    for attribute in e.attributes() {
                        let attribute =
                            attribute.map_err(|error| OoxmlError::Xml(error.to_string()))?;
                        let value = attribute
                            .decoded_and_normalized_value(XmlVersion::Explicit1_0, reader.decoder())
                            .map_err(|error| OoxmlError::Xml(error.to_string()))?;
                        match attribute.key.local_name().as_ref() {
                            b"type" => {
                                run_break.break_type = match value.as_ref() {
                                    "textWrapping" => RunBreakType::TextWrapping,
                                    "page" => RunBreakType::Page,
                                    "column" => RunBreakType::Column,
                                    _ => {
                                        return Err(OoxmlError::InvalidFormat(format!(
                                            "invalid Word run break type '{value}'"
                                        )));
                                    },
                                };
                            },
                            b"clear" => {
                                run_break.clear = match value.as_ref() {
                                    "none" => RunBreakClear::None,
                                    "left" => RunBreakClear::Left,
                                    "right" => RunBreakClear::Right,
                                    "all" => RunBreakClear::All,
                                    _ => {
                                        return Err(OoxmlError::InvalidFormat(format!(
                                            "invalid Word run break clear value '{value}'"
                                        )));
                                    },
                                };
                            },
                            _ => {},
                        }
                    }
                    breaks.push(run_break);
                },
                Ok(Event::Eof) => break,
                Err(error) => return Err(OoxmlError::Xml(error.to_string())),
                _ => {},
            }
        }
        Ok(breaks)
    }

    /// Count layout-engine page-break hints in this run.
    ///
    /// `<w:lastRenderedPageBreak>` is not an authored break; it records where Word last
    /// paginated content, so it is intentionally exposed separately from [`Self::breaks`].
    pub fn last_rendered_page_break_count(&self) -> Result<usize> {
        let mut reader = Reader::from_reader(self.xml_bytes());
        reader.config_mut().trim_text(true);
        let mut count = 0usize;
        loop {
            match reader.read_event() {
                Ok(Event::Start(e)) | Ok(Event::Empty(e))
                    if e.local_name().as_ref() == b"lastRenderedPageBreak" =>
                {
                    count = count.checked_add(1).ok_or_else(|| {
                        OoxmlError::InvalidFormat(
                            "too many rendered page break markers in one run".to_string(),
                        )
                    })?;
                },
                Ok(Event::Eof) => break,
                Err(error) => return Err(OoxmlError::Xml(error.to_string())),
                _ => {},
            }
        }
        Ok(count)
    }

    /// Check if this run is bold.
    ///
    /// Returns `Some(true)` if bold is explicitly enabled,
    /// `Some(false)` if explicitly disabled,
    /// `None` if not specified (inherits from style).
    pub fn bold(&self) -> Result<Option<bool>> {
        self.get_bool_property(b"b")
    }

    /// Check if this run is italic.
    ///
    /// Returns `Some(true)` if italic is explicitly enabled,
    /// `Some(false)` if explicitly disabled,
    /// `None` if not specified (inherits from style).
    pub fn italic(&self) -> Result<Option<bool>> {
        self.get_bool_property(b"i")
    }

    /// Check whether this run has a direct underline enabled.
    ///
    /// Returns `Some(false)` for an explicit `w:val="none"` and `None` when the
    /// run has no direct underline property and therefore inherits from styles.
    pub fn underline(&self) -> Result<Option<bool>> {
        Ok(self
            .underline_style()?
            .map(|style| style != UnderlineStyle::None))
    }

    /// Return the direct underline pattern, including an explicit
    /// [`UnderlineStyle::None`].
    pub fn underline_style(&self) -> Result<Option<UnderlineStyle>> {
        Ok(self
            .underline_formatting()?
            .map(|underline| underline.style))
    }

    /// Return the complete direct underline formatting for this run.
    ///
    /// This preserves all `CT_Underline` fields without resolving style or
    /// theme inheritance. A present `<w:u/>` is interpreted as a single
    /// underline for compatibility with documents emitted by Word processors.
    pub fn underline_formatting(&self) -> Result<Option<RunUnderline>> {
        parse_run_underline(self.xml_bytes())
    }

    /// Check if this run is strikethrough.
    ///
    /// Returns `Some(true)` if strikethrough is present,
    /// `None` if not specified.
    pub fn strikethrough(&self) -> Result<Option<bool>> {
        self.get_bool_property(b"strike")
    }

    /// Get text and properties in a single XML parse.
    ///
    /// This is **the fastest way** to extract both text content and formatting properties
    /// from a run, as it parses the XML only once instead of twice (text() + get_properties()).
    ///
    /// # Performance
    ///
    /// This provides 2x speedup over calling `text()` and `get_properties()` separately,
    /// and 4-6x speedup over individual property methods.
    ///
    /// # Returns
    ///
    /// A tuple of (text_content, properties)
    ///
    /// # Example
    ///
    /// ```rust,ignore
    /// // Fastest: Single XML parse for both text and properties
    /// let (text, props) = run.get_text_and_properties()?;
    /// if props.bold.unwrap_or(false) {
    ///     write_bold(&text);
    /// }
    /// ```
    pub fn get_text_and_properties(&self) -> Result<(String, RunProperties)> {
        let mut reader = Reader::from_reader(self.xml_bytes());
        reader.config_mut().trim_text(false);

        let mut props = RunProperties::default();
        let mut text = String::with_capacity(self.xml_bytes().len() / 8);
        let mut in_r_pr = false;
        let mut in_text_element = false;

        loop {
            match reader.read_event() {
                Ok(Event::Start(e)) => {
                    let name = e.local_name();

                    if name.as_ref() == b"t" {
                        in_text_element = true;
                    } else if name.as_ref() == b"rPr" {
                        in_r_pr = true;
                    } else {
                        match name.as_ref() {
                            b"tab" => text.push('\t'),
                            b"br" | b"cr" => text.push('\n'),
                            b"noBreakHyphen" => text.push('\u{2011}'),
                            b"softHyphen" => text.push('\u{00ad}'),
                            _ => {},
                        }
                        if in_r_pr {
                            update_run_properties(&mut props, &e)?;
                        }
                    }
                },
                Ok(Event::Empty(e)) => {
                    let name = e.local_name();
                    match name.as_ref() {
                        b"tab" => text.push('\t'),
                        b"br" | b"cr" => text.push('\n'),
                        b"noBreakHyphen" => text.push('\u{2011}'),
                        b"softHyphen" => text.push('\u{00ad}'),
                        _ => {},
                    }
                    if in_r_pr {
                        update_run_properties(&mut props, &e)?;
                    }
                },
                Ok(Event::Text(content)) if in_text_element => {
                    let decoded = content
                        .xml_content(XmlVersion::Explicit1_0)
                        .map_err(|error| OoxmlError::Xml(error.to_string()))?;
                    let unescaped = quick_xml::escape::unescape(&decoded)
                        .map_err(|error| OoxmlError::Xml(error.to_string()))?;
                    text.push_str(&unescaped);
                },
                Ok(Event::CData(content)) if in_text_element => {
                    let decoded = content
                        .xml_content(XmlVersion::Explicit1_0)
                        .map_err(|error| OoxmlError::Xml(error.to_string()))?;
                    text.push_str(&decoded);
                },
                Ok(Event::GeneralRef(reference)) if in_text_element => {
                    text.push_str(&decode_xml_reference(&reference)?);
                },
                Ok(Event::End(e)) => {
                    let name = e.local_name();
                    if name.as_ref() == b"t" {
                        in_text_element = false;
                    } else if name.as_ref() == b"rPr" {
                        in_r_pr = false;
                    }
                },
                Ok(Event::Eof) => break,
                Err(e) => return Err(OoxmlError::Xml(e.to_string())),
                _ => {},
            }
        }

        Ok((text, props))
    }

    /// Get all formatting properties in a single pass.
    ///
    /// This is **significantly faster** than calling individual property methods
    /// (bold(), italic(), strikethrough(), vertical_position()) because it parses
    /// the XML only once instead of multiple times.
    ///
    /// # Performance
    ///
    /// For documents with many runs, using this method can provide 3-4x speedup
    /// compared to individual property access.
    ///
    /// # Example
    ///
    /// ```rust,ignore
    /// // Fast: Single XML parse
    /// let props = run.get_properties()?;
    /// if props.bold.unwrap_or(false) {
    ///     // Handle bold text
    /// }
    ///
    /// // Slow: Multiple XML parses
    /// if run.bold()?.unwrap_or(false) {
    ///     // Handle bold text
    /// }
    /// ```
    pub fn get_properties(&self) -> Result<RunProperties> {
        let mut reader = Reader::from_reader(self.xml_bytes());
        reader.config_mut().trim_text(true);

        let mut props = RunProperties::default();
        let mut in_r_pr = false;

        loop {
            match reader.read_event() {
                Ok(Event::Start(e)) => {
                    let name = e.local_name();
                    if name.as_ref() == b"rPr" {
                        in_r_pr = true;
                    } else if in_r_pr {
                        update_run_properties(&mut props, &e)?;
                    }
                },
                Ok(Event::Empty(e)) if in_r_pr && e.local_name().as_ref() != b"rPr" => {
                    update_run_properties(&mut props, &e)?;
                },
                Ok(Event::End(e)) if e.local_name().as_ref() == b"rPr" => {
                    // Exit early once we've finished parsing rPr
                    return Ok(props);
                },
                Ok(Event::Eof) => break,
                Err(e) => return Err(OoxmlError::Xml(e.to_string())),
                _ => {},
            }
        }

        Ok(props)
    }

    /// Get the vertical position of this run (superscript/subscript).
    ///
    /// Returns the vertical positioning if specified, None if normal.
    pub fn vertical_position(&self) -> Result<Option<VerticalPosition>> {
        let mut reader = Reader::from_reader(self.xml_bytes());
        reader.config_mut().trim_text(true);

        let mut in_r_pr = false;

        loop {
            match reader.read_event() {
                Ok(Event::Start(e)) | Ok(Event::Empty(e)) => {
                    let name = e.local_name();
                    if name.as_ref() == b"rPr" {
                        in_r_pr = true;
                    } else if in_r_pr && name.as_ref() == b"vertAlign" {
                        for attr in e.attributes().flatten() {
                            if attr.key.local_name().as_ref() == b"val" {
                                let value = attr.value.as_ref();
                                match value {
                                    b"superscript" => {
                                        return Ok(Some(VerticalPosition::Superscript));
                                    },
                                    b"subscript" => return Ok(Some(VerticalPosition::Subscript)),
                                    _ => {},
                                }
                            }
                        }
                    }
                },
                Ok(Event::End(e)) if e.local_name().as_ref() == b"rPr" => {
                    in_r_pr = false;
                },
                Ok(Event::Eof) => break,
                Err(e) => return Err(OoxmlError::Xml(e.to_string())),
                _ => {},
            }
        }

        Ok(None)
    }

    /// Get the font name for this run.
    ///
    /// Returns the typeface name if specified, None if inherited.
    pub fn font_name(&self) -> Result<Option<String>> {
        let mut reader = Reader::from_reader(self.xml_bytes());
        reader.config_mut().trim_text(true);

        let mut in_r_pr = false;

        loop {
            match reader.read_event() {
                Ok(Event::Start(e)) | Ok(Event::Empty(e)) => {
                    let name = e.local_name();
                    if name.as_ref() == b"rPr" {
                        in_r_pr = true;
                    } else if in_r_pr && name.as_ref() == b"rFonts" {
                        for attr in e.attributes().flatten() {
                            if attr.key.local_name().as_ref() == b"ascii" {
                                let value = attr
                                    .decoded_and_normalized_value(
                                        XmlVersion::Implicit1_0,
                                        reader.decoder(),
                                    )
                                    .unwrap_or(Cow::Borrowed(""));
                                return Ok(Some(value.to_string()));
                            }
                        }
                    }
                },
                Ok(Event::End(e)) if e.local_name().as_ref() == b"rPr" => {
                    break;
                },
                Ok(Event::Eof) => break,
                Err(e) => return Err(OoxmlError::Xml(e.to_string())),
                _ => {},
            }
        }

        Ok(None)
    }

    /// Get the font size for this run in half-points.
    ///
    /// Returns the size if specified, None if inherited.
    /// Note: Word stores font size in half-points (e.g., 24 = 12pt).
    pub fn font_size(&self) -> Result<Option<u32>> {
        let mut reader = Reader::from_reader(self.xml_bytes());
        reader.config_mut().trim_text(true);

        let mut in_r_pr = false;

        loop {
            match reader.read_event() {
                Ok(Event::Start(e)) | Ok(Event::Empty(e)) => {
                    let name = e.local_name();
                    if name.as_ref() == b"rPr" {
                        in_r_pr = true;
                    } else if in_r_pr && name.as_ref() == b"sz" {
                        for attr in e.attributes().flatten() {
                            if attr.key.local_name().as_ref() == b"val"
                                && let Ok(value) = std::str::from_utf8(&attr.value)
                                && let Ok(size) = value.parse::<u32>()
                            {
                                return Ok(Some(size));
                            }
                        }
                    }
                },
                Ok(Event::End(e)) if e.local_name().as_ref() == b"rPr" => {
                    break;
                },
                Ok(Event::Eof) => break,
                Err(e) => return Err(OoxmlError::Xml(e.to_string())),
                _ => {},
            }
        }

        Ok(None)
    }

    /// Check if this run contains an OMML formula.
    ///
    /// Returns the OMML XML content if this run contains a mathematical formula,
    /// None otherwise. This method looks for `<m:oMath>` elements embedded in the run.
    ///
    /// The returned string preserves the exact source XML for the first formula in the run.
    pub fn omml_formula(&self) -> Result<Option<String>> {
        Ok(extract_omml_formulas(self.xml_bytes())?.into_iter().next())
    }
    /// Helper to extract boolean properties from run properties.
    ///
    /// Handles the tri-state logic where w:val can be "true", "false", "1", "0"
    /// or the element can be present without a val attribute (implies true).
    fn get_bool_property(&self, property_name: &[u8]) -> Result<Option<bool>> {
        let mut reader = Reader::from_reader(self.xml_bytes());
        reader.config_mut().trim_text(true);

        let mut in_r_pr = false;

        loop {
            match reader.read_event() {
                Ok(Event::Start(e)) | Ok(Event::Empty(e)) => {
                    let name = e.local_name();
                    if name.as_ref() == b"rPr" {
                        in_r_pr = true;
                    } else if in_r_pr && name.as_ref() == property_name {
                        // Check for w:val attribute
                        for attr in e.attributes().flatten() {
                            if attr.key.local_name().as_ref() == b"val" {
                                let value = attr.value.as_ref();
                                return Ok(Some(is_on(value)));
                            }
                        }
                        // Element present without val attribute means true
                        return Ok(Some(true));
                    }
                },
                Ok(Event::End(e)) if e.local_name().as_ref() == b"rPr" => {
                    in_r_pr = false;
                },
                Ok(Event::Eof) => break,
                Err(e) => return Err(OoxmlError::Xml(e.to_string())),
                _ => {},
            }
        }

        Ok(None)
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_run_text_extraction() {
        let xml = br#"<w:r xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
            <w:t>Hello, World!</w:t>
        </w:r>"#;

        let run = Run::new(xml.to_vec());
        let text = run.text().unwrap();
        assert_eq!(text, "Hello, World!");
    }

    #[test]
    fn extracts_decoded_word_text_and_special_characters() {
        let xml = br#"<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
            <w:r><w:t xml:space="preserve">  A &amp; B &lt; C &#x1F600;  </w:t></w:r>
            <w:r><w:tab/><w:t/><w:br/><w:cr/><w:noBreakHyphen/><w:softHyphen/><w:t>tail</w:t></w:r>
        </w:p>"#;
        let paragraph = Paragraph::new(xml.to_vec());
        assert_eq!(
            paragraph.text().unwrap(),
            "  A & B < C 😀  \t\n\n‑\u{00ad}tail"
        );
        let runs = paragraph.runs().unwrap();
        assert_eq!(runs[0].text().unwrap(), "  A & B < C 😀  ");
        assert_eq!(runs[1].text().unwrap(), "\t\n\n‑\u{00ad}tail");
    }

    #[test]
    fn runs_resolve_namespace_aliases_and_ignore_lookalikes() {
        let xml = br#"<wp:p xmlns:wp="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:false="urn:not-wordprocessingml">
            <false:r><false:t>ignored outer</false:t></false:r>
            <wp:r><false:t>ignored inner</false:t><false:tab/><false:br/><wp:t>kept</wp:t><wp:tab/></wp:r>
            <wp:r/>
        </wp:p>"#;

        let paragraph = Paragraph::new(xml.to_vec());
        assert_eq!(paragraph.text().unwrap(), "kept\t");
        let runs = paragraph.runs().unwrap();
        assert_eq!(runs.len(), 2);
        assert_eq!(runs[0].text().unwrap(), "kept\t");
        assert_eq!(runs[1].text().unwrap(), "");
    }

    #[test]
    fn runs_accept_the_strict_wordprocessingml_namespace() {
        let xml = br#"<s:p xmlns:s="http://purl.oclc.org/ooxml/wordprocessingml/main">
            <s:r><s:t>strict</s:t></s:r>
        </s:p>"#;

        let paragraph = Paragraph::new(xml.to_vec());
        assert_eq!(paragraph.text().unwrap(), "strict");
        let runs = paragraph.runs().unwrap();
        assert_eq!(runs.len(), 1);
        assert_eq!(runs[0].text().unwrap(), "strict");
    }

    #[test]
    fn runs_accept_fragments_with_an_inherited_namespace_binding() {
        let xml = br#"<wp:p><wp:r><wp:t>inherited</wp:t></wp:r></wp:p>"#;

        let paragraph = Paragraph::new(xml.to_vec());
        assert_eq!(paragraph.text().unwrap(), "inherited");
        let runs = paragraph.runs().unwrap();
        assert_eq!(runs.len(), 1);
        assert_eq!(runs[0].text().unwrap(), "inherited");
    }

    #[test]
    fn reads_nested_smart_tags_and_their_typed_metadata() {
        let xml = br#"<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:false="urn:not-wordprocessingml">
            <w:smartTag w:uri="urn:contacts" w:element="person">
                <w:smartTagPr>
                    <w:attr w:uri="urn:meta" w:name="kind" w:val="friend &amp; peer"/>
                </w:smartTagPr>
                <w:r><w:t>A &amp; </w:t></w:r>
                <w:smartTag w:element="givenName">
                    <w:smartTagPr><w:attr w:name="language" w:val="en"/></w:smartTagPr>
                    <w:r><w:t>Bob</w:t></w:r>
                </w:smartTag>
            </w:smartTag>
            <false:smartTag false:element="ignored"><w:r><w:t>not a tag</w:t></w:r></false:smartTag>
            <w:smartTag w:element="empty"/>
        </w:p>"#;

        let paragraph = Paragraph::new(xml.to_vec());
        let tags = paragraph.smart_tags().unwrap();
        assert_eq!(tags.len(), 3);

        assert_eq!(tags[0].uri.as_deref(), Some("urn:contacts"));
        assert_eq!(tags[0].element, "person");
        assert_eq!(tags[0].attributes.len(), 1);
        assert_eq!(tags[0].attributes[0].uri.as_deref(), Some("urn:meta"));
        assert_eq!(tags[0].attributes[0].name, "kind");
        assert_eq!(tags[0].attributes[0].value, "friend & peer");
        assert_eq!(tags[0].text().unwrap(), "A & Bob");

        assert_eq!(tags[1].element, "givenName");
        assert_eq!(tags[1].attributes[0].name, "language");
        assert_eq!(tags[1].text().unwrap(), "Bob");
        assert_eq!(tags[2].element, "empty");
        assert_eq!(tags[2].text().unwrap(), "");

        assert_eq!(paragraph.runs().unwrap().len(), 3);
    }

    #[test]
    fn smart_tags_require_schema_mandated_attributes() {
        let missing_element = Paragraph::new(
            br#"<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:smartTag><w:r><w:t>invalid</w:t></w:r></w:smartTag></w:p>"#
                .to_vec(),
        );
        assert!(missing_element.smart_tags().is_err());

        let missing_property_value = Paragraph::new(
            br#"<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:smartTag w:element="person"><w:smartTagPr><w:attr w:name="kind"/></w:smartTagPr></w:smartTag></w:p>"#
                .to_vec(),
        );
        assert!(missing_property_value.smart_tags().is_err());
    }

    #[test]
    fn runs_reject_unterminated_run_xml() {
        let xml = br#"<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:r><w:t>truncated</w:t>"#;
        let paragraph = Paragraph::new(xml.to_vec());
        assert!(paragraph.text().is_err());
        assert!(paragraph.runs().is_err());
    }

    #[test]
    fn optimized_run_extraction_matches_text_and_reads_qualified_properties() {
        let run = Run::new(
            br#"<w:r xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                <w:rPr><w:b w:val="0"/><w:i w:val="on"/><w:strike/><w:u w:val="dashLongHeavy"/><w:vertAlign w:val="superscript"/></w:rPr>
                <w:t xml:space="preserve"> A &amp; <![CDATA[B < C]]> &#x1F600; </w:t><w:t/>
                <w:tab/><w:br/><w:cr/><w:noBreakHyphen/><w:softHyphen/><w:t>tail</w:t>
            </w:r>"#
                .to_vec(),
        );

        let expected_text = " A & B < C 😀 \t\n\n‑\u{00ad}tail";
        let (text, properties) = run.get_text_and_properties().unwrap();
        assert_eq!(text, expected_text);
        assert_eq!(text, run.text().unwrap());
        assert_eq!(properties.bold, Some(false));
        assert_eq!(properties.italic, Some(true));
        assert_eq!(properties.strikethrough, Some(true));
        assert_eq!(properties.underline, Some(UnderlineStyle::DashLongHeavy));
        assert_eq!(
            properties.vertical_position,
            Some(VerticalPosition::Superscript)
        );

        let properties_only = run.get_properties().unwrap();
        assert_eq!(properties_only.bold, properties.bold);
        assert_eq!(properties_only.italic, properties.italic);
        assert_eq!(properties_only.strikethrough, properties.strikethrough);
        assert_eq!(properties_only.underline, properties.underline);
        assert_eq!(
            properties_only.vertical_position,
            properties.vertical_position
        );
        assert_eq!(run.bold().unwrap(), Some(false));
        assert_eq!(run.italic().unwrap(), Some(true));
        assert_eq!(
            run.underline_style().unwrap(),
            Some(UnderlineStyle::DashLongHeavy)
        );
        assert_eq!(
            run.vertical_position().unwrap(),
            properties.vertical_position
        );
    }

    #[test]
    fn reads_every_wordprocessingml_underline_pattern() {
        let patterns = [
            ("none", UnderlineStyle::None),
            ("single", UnderlineStyle::Single),
            ("words", UnderlineStyle::Words),
            ("double", UnderlineStyle::Double),
            ("thick", UnderlineStyle::Thick),
            ("dotted", UnderlineStyle::Dotted),
            ("dottedHeavy", UnderlineStyle::DottedHeavy),
            ("dash", UnderlineStyle::Dashed),
            ("dashedHeavy", UnderlineStyle::DashedHeavy),
            ("dashLong", UnderlineStyle::DashLong),
            ("dashLongHeavy", UnderlineStyle::DashLongHeavy),
            ("dotDash", UnderlineStyle::DotDash),
            ("dashDotHeavy", UnderlineStyle::DashDotHeavy),
            ("dotDotDash", UnderlineStyle::DotDotDash),
            ("dashDotDotHeavy", UnderlineStyle::DashDotDotHeavy),
            ("wave", UnderlineStyle::Wave),
            ("wavyHeavy", UnderlineStyle::WavyHeavy),
            ("wavyDouble", UnderlineStyle::WavyDouble),
        ];

        for (value, expected) in patterns {
            let run = Run::new(
                format!(
                    r#"<w:r xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:rPr><w:u w:val="{value}"/></w:rPr></w:r>"#
                )
                .into_bytes(),
            );
            assert_eq!(run.underline_style().unwrap(), Some(expected));
            assert_eq!(run.get_properties().unwrap().underline, Some(expected));
            assert_eq!(
                run.underline().unwrap(),
                Some(expected != UnderlineStyle::None)
            );
        }

        let implicit_single = Run::new(
            br#"<w:r xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:rPr><w:u/></w:rPr></w:r>"#
                .to_vec(),
        );
        assert_eq!(
            implicit_single.underline_style().unwrap(),
            Some(UnderlineStyle::Single)
        );
        let inherited = Run::new(
            br#"<w:r xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:rPr/></w:r>"#
                .to_vec(),
        );
        assert_eq!(inherited.underline().unwrap(), None);
    }

    #[test]
    fn reads_complete_underline_metadata_namespace_aware() {
        let strict = Run::new(
            br#"<s:r xmlns:s="http://purl.oclc.org/ooxml/wordprocessingml/main" xmlns:false="urn:not-wordprocessingml"><s:rPr><false:u false:val="double"/><s:u s:val="wavyDouble" s:color="A0b1C2" s:themeColor="accent4" s:themeTint="0a" s:themeShade="FF"/></s:rPr></s:r>"#
                .to_vec(),
        );
        assert_eq!(
            strict.underline_formatting().unwrap(),
            Some(RunUnderline {
                style: UnderlineStyle::WavyDouble,
                color: Some(RunUnderlineColor::Rgb([0xA0, 0xB1, 0xC2])),
                theme_color: Some(ThemeColor::Accent4),
                theme_tint: Some(0x0A),
                theme_shade: Some(0xFF),
            })
        );

        let inherited =
            Run::new(br#"<q:r><q:rPr><q:u q:val="words" q:color="auto"/></q:rPr></q:r>"#.to_vec());
        assert_eq!(
            inherited.underline_formatting().unwrap(),
            Some(RunUnderline {
                style: UnderlineStyle::Words,
                color: Some(RunUnderlineColor::Auto),
                theme_color: None,
                theme_tint: None,
                theme_shade: None,
            })
        );
    }

    #[test]
    fn rejects_invalid_or_duplicate_underline_properties() {
        for xml in [
            br#"<w:r xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:rPr><w:u w:val="triple"/></w:rPr></w:r>"#.as_slice(),
            br#"<w:r xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:rPr><w:u w:color="12345"/></w:rPr></w:r>"#.as_slice(),
            br#"<w:r xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:rPr><w:u w:themeColor="accent7"/></w:rPr></w:r>"#.as_slice(),
            br#"<w:r xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:rPr><w:u w:themeTint="000"/></w:rPr></w:r>"#.as_slice(),
            br#"<w:r xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:rPr><w:u/><w:u/></w:rPr></w:r>"#.as_slice(),
        ] {
            assert!(Run::new(xml.to_vec()).underline_formatting().is_err());
        }
    }

    #[test]
    fn rejects_unknown_entities_in_word_text() {
        let run = Run::new(
            br#"<w:r xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:t>&unknown;</w:t></w:r>"#
                .to_vec(),
        );
        assert!(run.text().is_err());
        assert!(run.get_text_and_properties().is_err());
    }

    #[test]
    fn omml_formulas_preserve_inline_and_display_xml_exactly() {
        let xml = br#"<wp:p xmlns:wp="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"
            xmlns:q="http://schemas.openxmlformats.org/officeDocument/2006/math"
            xmlns:false="urn:not-omml">
            <wp:r>
                <m:oMath data-id="1"><m:r><wp:rPr/><m:t><![CDATA[x < y]]></m:t></m:r></m:oMath>
                <q:oMath q:id="2"/>
            </wp:r>
            <m:oMathPara><q:oMath><q:r/></q:oMath></m:oMathPara>
            <false:oMath>ignored</false:oMath>
        </wp:p>"#;
        let paragraph = Paragraph::new(xml.to_vec());

        assert_eq!(
            paragraph.omml_formulas().unwrap(),
            vec![
                r#"<m:oMath data-id="1"><m:r><wp:rPr/><m:t><![CDATA[x < y]]></m:t></m:r></m:oMath>"#,
                r#"<q:oMath q:id="2"/>"#,
            ]
        );
        assert_eq!(
            paragraph.paragraph_level_formulas().unwrap(),
            vec!["<q:oMath><q:r/></q:oMath>"]
        );
        assert_eq!(
            paragraph.runs().unwrap()[0].omml_formula().unwrap(),
            Some(
                r#"<m:oMath data-id="1"><m:r><wp:rPr/><m:t><![CDATA[x < y]]></m:t></m:r></m:oMath>"#
                    .to_string()
            )
        );
        let inline = paragraph.inline_office_math().unwrap();
        assert_eq!(inline.len(), 2);
        assert_eq!(
            inline[1].xml(),
            r#"<q:oMath q:id="2" xmlns:q="http://schemas.openxmlformats.org/officeDocument/2006/math"/>"#
        );
        assert_eq!(
            paragraph.display_office_math().unwrap()[0].xml(),
            r#"<q:oMath xmlns:q="http://schemas.openxmlformats.org/officeDocument/2006/math"><q:r/></q:oMath>"#
        );
    }

    #[test]
    fn omml_formulas_accept_strict_and_inherited_prefixes() {
        let strict = Paragraph::new(
            br#"<s:p xmlns:s="http://purl.oclc.org/ooxml/wordprocessingml/main" xmlns:math="http://purl.oclc.org/ooxml/officeDocument/math"><s:r><math:oMath><math:r/></math:oMath></s:r></s:p>"#
                .to_vec(),
        );
        assert_eq!(
            strict.omml_formulas().unwrap(),
            vec!["<math:oMath><math:r/></math:oMath>"]
        );

        let inherited = Run::new(br#"<w:r><m:oMath><m:r/></m:oMath></w:r>"#.to_vec());
        assert_eq!(
            inherited.omml_formula().unwrap().as_deref(),
            Some("<m:oMath><m:r/></m:oMath>")
        );

        let inherited_default = Paragraph::new(
            br#"<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns="http://schemas.openxmlformats.org/officeDocument/2006/math"><w:r><oMath><r/></oMath></w:r></w:p>"#
                .to_vec(),
        );
        assert_eq!(
            inherited_default.inline_office_math().unwrap()[0].xml(),
            r#"<oMath xmlns="http://schemas.openxmlformats.org/officeDocument/2006/math"><r/></oMath>"#
        );

        let foreign =
            Run::new(br#"<w:r xmlns:m="urn:not-omml"><m:oMath><m:r/></m:oMath></w:r>"#.to_vec());
        assert_eq!(foreign.omml_formula().unwrap(), None);
    }

    #[test]
    fn omml_formulas_reject_malformed_xml() {
        let paragraph = Paragraph::new(
            br#"<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"><w:r><m:oMath><m:r/></w:r></w:p>"#
                .to_vec(),
        );
        assert!(paragraph.omml_formulas().is_err());
        assert!(paragraph.paragraph_level_formulas().is_err());
    }

    #[test]
    fn test_run_bold() {
        let xml = br#"<w:r xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
            <w:rPr><w:b/></w:rPr>
            <w:t>Bold text</w:t>
        </w:r>"#;

        let run = Run::new(xml.to_vec());
        assert!(run.bold().unwrap().unwrap_or(false));
    }

    #[test]
    fn test_run_italic() {
        let xml = br#"<w:r xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
            <w:rPr><w:i/></w:rPr>
            <w:t>Italic text</w:t>
        </w:r>"#;

        let run = Run::new(xml.to_vec());
        assert!(run.italic().unwrap().unwrap_or(false));
    }

    #[test]
    fn parses_typed_run_breaks_and_rendered_hints() {
        let xml = br#"<w:r xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
            <w:t>Before</w:t><w:br/><w:br w:type="page"/><w:br w:type="column" w:clear="all"/>
            <w:br w:type="textWrapping" w:clear="left"></w:br>
            <w:lastRenderedPageBreak/><w:lastRenderedPageBreak></w:lastRenderedPageBreak>
        </w:r>"#;
        let run = Run::new(xml.to_vec());
        assert_eq!(
            run.breaks().unwrap().as_slice(),
            [
                RunBreak::default(),
                RunBreak {
                    break_type: RunBreakType::Page,
                    clear: RunBreakClear::None,
                },
                RunBreak {
                    break_type: RunBreakType::Column,
                    clear: RunBreakClear::All,
                },
                RunBreak {
                    break_type: RunBreakType::TextWrapping,
                    clear: RunBreakClear::Left,
                },
            ]
        );
        assert_eq!(run.last_rendered_page_break_count().unwrap(), 2);
        assert_eq!(run.text().unwrap(), "Before\n\n\n\n");
    }

    #[test]
    fn rejects_invalid_run_break_enums() {
        let invalid_type = Run::new(
            br#"<w:r xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:br w:type="section"/></w:r>"#
                .to_vec(),
        );
        assert!(invalid_type.breaks().is_err());

        let invalid_clear = Run::new(
            br#"<w:r xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:br w:clear="center"/></w:r>"#
                .to_vec(),
        );
        assert!(invalid_clear.breaks().is_err());
    }

    #[test]
    fn reads_direct_paragraph_division_ids_namespace_aware() {
        let paragraph = Paragraph::new(
            br#"<q:p xmlns:q="http://purl.oclc.org/ooxml/wordprocessingml/main" xmlns:false="urn:not-wordprocessingml"><q:pPr><false:divId false:val="9"/><q:divId q:val=" +123456789012345678901234567890 "/></q:pPr><q:r><q:t>text</q:t></q:r></q:p>"#
                .to_vec(),
        );
        assert_eq!(
            paragraph.division_id().unwrap().as_deref(),
            Some("+123456789012345678901234567890")
        );

        let inherited =
            Paragraph::new(br#"<q:p><q:pPr><q:divId q:val="-7"/></q:pPr></q:p>"#.to_vec());
        assert_eq!(inherited.division_id().unwrap().as_deref(), Some("-7"));

        let invalid = Paragraph::new(
            br#"<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:pPr><w:divId w:val="1.5"/></w:pPr></w:p>"#
                .to_vec(),
        );
        assert!(invalid.division_id().is_err());
    }
}
