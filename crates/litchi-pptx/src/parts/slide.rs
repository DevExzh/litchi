//! Borrowed slide, layout, and master part views.

use std::borrow::Cow;

use litchi_ooxml_common::mce::{Capabilities, Limits as MceLimits, process_markup_compatibility};
use litchi_ooxml_common::xml::{DRAWINGML_NAMESPACE, STRICT_DRAWINGML_NAMESPACE};
use litchi_opc::constants::{content_type as ct, relationship_type as rt};
use litchi_opc::{OpcPackage, Part};
use quick_xml::events::Event;
use quick_xml::name::{Namespace, QName, ResolveResult};
use quick_xml::reader::{NsReader, Reader};

use super::{
    MceCapture, invalid, processed_xml, processed_xml_with_capture, processed_xml_with_source,
    related_part_by_type, validate_content_type,
};
use crate::notes::SlideRootProof;
use crate::shape::Scene;
use crate::{Error, Result};

// The semantic sink deliberately uses a smaller, stream-specific policy than
// the general PresentationML part reader. It retains at most one selected
// slide string and never retains the processed XML after that slide returns.
const MAX_SEMANTIC_TEXT_RAW_XML_BYTES: usize = 64 * 1024 * 1024;
const MAX_SEMANTIC_TEXT_PROCESSED_XML_BYTES: usize = 64 * 1024 * 1024;
const MAX_SEMANTIC_TEXT_EVENTS: usize = 1_000_000;
const MAX_SEMANTIC_TEXT_DEPTH: usize = 128;
const MAX_SEMANTIC_TEXT_RUNS: usize = 100_000;
const MAX_SEMANTIC_TEXT_OBJECTS: usize = 100_000;
const MAX_SEMANTIC_TEXT_EVENT_BYTES: usize = 1024 * 1024;
const MAX_SEMANTIC_TEXT_REFERENCE_BYTES: usize = 64 * 1024;
const MAX_SEMANTIC_TEXT_BYTES: usize = 16 * 1024 * 1024;

fn root_name_from_xml(xml: &[u8]) -> Result<String> {
    let mut reader = NsReader::from_reader(xml);
    loop {
        let (namespace, event) = reader.read_resolved_event()?;
        match event {
            Event::Start(element) | Event::Empty(element) => {
                if !crate::namespace::is_presentationml_name(
                    &namespace,
                    element.name(),
                    element.local_name().as_ref(),
                ) {
                    return Err(invalid("PresentationML part has an invalid root namespace"));
                }
                return String::from_utf8(element.local_name().as_ref().to_vec())
                    .map_err(|_err| invalid("PresentationML root name is not UTF-8"));
            },
            Event::Decl(_) | Event::Comment(_) => {},
            Event::Text(text) if text.as_ref().iter().all(u8::is_ascii_whitespace) => {},
            _ => return Err(invalid("PresentationML part lacks an element root")),
        }
    }
}

fn root_name(part: &dyn Part) -> Result<String> {
    let xml = processed_xml(part)?;
    root_name_from_xml(xml.as_ref())
}

fn c_sld_name_from_xml(xml: &[u8]) -> Result<Option<String>> {
    crate::namespace::presentation_name(xml)
}

fn c_sld_name(part: &dyn Part) -> Result<Option<String>> {
    let xml = processed_xml(part)?;
    c_sld_name_from_xml(xml.as_ref())
}

/// Read the ordered `p:sldLayoutIdLst` relationship references owned by a
/// slide master. The OPC relationship collection may contain stale or
/// producer-private edges; the XML list is the semantic owner of the layout
/// inventory.
fn layout_relationship_ids(part: &dyn Part) -> Result<Vec<String>> {
    let xml = processed_xml(part)?;
    let mut reader = NsReader::from_reader(xml.as_ref());
    let mut depth = 0usize;
    let mut in_list = false;
    let mut seen_list = false;
    let mut relationship_ids = Vec::new();

    loop {
        let (_namespace, event) = reader.read_resolved_event()?;
        match event {
            Event::Start(element) => {
                depth = depth
                    .checked_add(1)
                    .ok_or_else(|| invalid("slide-master XML nesting is too deep"))?;
                if depth == 2 && element.local_name().as_ref() == b"sldLayoutIdLst" {
                    if seen_list {
                        return Err(invalid("duplicate slide-layout ID list"));
                    }
                    seen_list = true;
                    in_list = true;
                } else if depth == 3 && in_list && element.local_name().as_ref() == b"sldLayoutId" {
                    relationship_ids.push(
                        crate::namespace::relationship_attribute_value(
                            &element,
                            b"id",
                            reader.decoder(),
                            reader.resolver(),
                        )?
                        .ok_or_else(|| {
                            invalid("slide-layout entry is missing its relationship ID")
                        })?,
                    );
                }
            },
            Event::Empty(element) => {
                let child_depth = depth
                    .checked_add(1)
                    .ok_or_else(|| invalid("slide-master XML nesting is too deep"))?;
                if child_depth == 2 && element.local_name().as_ref() == b"sldLayoutIdLst" {
                    if seen_list {
                        return Err(invalid("duplicate slide-layout ID list"));
                    }
                    seen_list = true;
                } else if child_depth == 3
                    && in_list
                    && element.local_name().as_ref() == b"sldLayoutId"
                {
                    relationship_ids.push(
                        crate::namespace::relationship_attribute_value(
                            &element,
                            b"id",
                            reader.decoder(),
                            reader.resolver(),
                        )?
                        .ok_or_else(|| {
                            invalid("slide-layout entry is missing its relationship ID")
                        })?,
                    );
                }
            },
            Event::End(element) => {
                if depth == 0 {
                    return Err(invalid("unexpected closing element in slide-master XML"));
                }
                if depth == 2 && element.local_name().as_ref() == b"sldLayoutIdLst" {
                    in_list = false;
                }
                depth -= 1;
            },
            Event::DocType(_) | Event::PI(_) => {
                return Err(invalid("DTDs and processing instructions are rejected"));
            },
            Event::Eof => break,
            _ => {},
        }
    }

    if depth != 0 {
        return Err(invalid("unterminated slide-master XML"));
    }
    Ok(relationship_ids)
}

fn root_bool(part: &dyn Part, attribute: &[u8], field: &str, default: bool) -> Result<bool> {
    let xml = processed_xml(part)?;
    let mut reader = Reader::from_reader(xml.as_ref());
    loop {
        match reader.read_event()? {
            Event::Start(element) | Event::Empty(element) => {
                let value = litchi_ooxml_common::xml::unqualified_attribute_value(
                    &element,
                    attribute,
                    reader.decoder(),
                )?;
                return value.map_or(Ok(default), |value| super::parse_bool(&value, field));
            },
            Event::Decl(_) | Event::Comment(_) => {},
            _ => return Err(invalid("PresentationML part lacks an element root")),
        }
    }
}

fn text_from_part(part: &dyn Part) -> Result<Option<String>> {
    let value = semantic_text_from_part(part, "\n")?;
    Ok((!value.is_empty()).then_some(value))
}

#[derive(Default)]
struct SemanticTextXmlBudget {
    events: usize,
    depth: usize,
}

impl SemanticTextXmlBudget {
    fn observe_event(&mut self, event_bytes: usize) -> Result<()> {
        self.events = self.events.checked_add(1).ok_or_else(|| {
            Error::Invalid("semantic slide XML event counter overflow".to_string())
        })?;
        if self.events > MAX_SEMANTIC_TEXT_EVENTS {
            return Err(Error::Limit {
                resource: "semantic slide XML events",
                limit: MAX_SEMANTIC_TEXT_EVENTS,
            });
        }
        if event_bytes > MAX_SEMANTIC_TEXT_EVENT_BYTES {
            return Err(Error::Limit {
                resource: "semantic slide XML event bytes",
                limit: MAX_SEMANTIC_TEXT_EVENT_BYTES,
            });
        }
        Ok(())
    }

    fn start(&mut self) -> Result<()> {
        self.depth = self.depth.checked_add(1).ok_or_else(|| {
            Error::Invalid("semantic slide XML depth counter overflow".to_string())
        })?;
        if self.depth > MAX_SEMANTIC_TEXT_DEPTH {
            return Err(Error::Limit {
                resource: "semantic slide XML depth",
                limit: MAX_SEMANTIC_TEXT_DEPTH,
            });
        }
        Ok(())
    }

    fn end(&mut self) -> Result<()> {
        if self.depth == 0 {
            return Err(invalid(
                "semantic slide XML has an unexpected closing element",
            ));
        }
        self.depth -= 1;
        Ok(())
    }
}

fn semantic_event_bytes(event: &Event<'_>) -> usize {
    match event {
        Event::Start(element) => element.as_ref().len(),
        Event::Empty(element) => element.as_ref().len(),
        Event::End(element) => element.as_ref().len(),
        Event::Text(text) => text.as_ref().len(),
        Event::CData(text) => text.as_ref().len(),
        Event::Comment(comment) => comment.as_ref().len(),
        Event::DocType(doctype) => doctype.as_ref().len(),
        Event::PI(pi) => pi.as_ref().len(),
        Event::Decl(decl) => decl.as_ref().len(),
        Event::GeneralRef(reference) => reference.as_ref().len(),
        Event::Eof => 0,
    }
}

fn validate_semantic_attributes(element: &quick_xml::events::BytesStart<'_>) -> Result<()> {
    let mut total = 0usize;
    for attribute in element.attributes().with_checks(true) {
        let attribute = attribute.map_err(|error| Error::Xml(error.to_string()))?;
        let attribute_bytes = attribute
            .key
            .as_ref()
            .len()
            .checked_add(attribute.value.len())
            .ok_or_else(|| Error::Invalid("semantic slide XML attribute length overflow".into()))?;
        total = total
            .checked_add(attribute_bytes)
            .ok_or_else(|| Error::Invalid("semantic slide XML attribute length overflow".into()))?;
        if total > MAX_SEMANTIC_TEXT_EVENT_BYTES {
            return Err(Error::Limit {
                resource: "semantic slide XML attribute bytes",
                limit: MAX_SEMANTIC_TEXT_EVENT_BYTES,
            });
        }
        let value = attribute
            .normalized_value(quick_xml::XmlVersion::Implicit1_0)
            .map_err(|error| Error::Xml(error.to_string()))?;
        validate_xml_characters(&value)?;
    }
    Ok(())
}

fn validate_semantic_attribute_names(
    reader: &NsReader<&[u8]>,
    element: &quick_xml::events::BytesStart<'_>,
) -> Result<()> {
    for attribute in element.attributes().with_checks(true) {
        let attribute = attribute.map_err(|error| Error::Xml(error.to_string()))?;
        let name = attribute.key.as_ref();
        if name == b"xmlns" || name.starts_with(b"xmlns:") {
            continue;
        }
        if let ResolveResult::Unknown(prefix) = reader.resolver().resolve_attribute(attribute.key).0
        {
            return Err(invalid(format!(
                "unresolved semantic slide attribute namespace prefix '{}'",
                String::from_utf8_lossy(prefix.as_ref())
            )));
        }
    }
    Ok(())
}

fn validate_semantic_element_namespace(namespace: &ResolveResult<'_>) -> Result<()> {
    if let ResolveResult::Unknown(prefix) = namespace {
        return Err(invalid(format!(
            "unresolved semantic slide element namespace prefix '{}'",
            String::from_utf8_lossy(prefix.as_ref())
        )));
    }
    Ok(())
}

fn validate_xml_characters(value: &str) -> Result<()> {
    if value.chars().all(|character| {
        matches!(
            character,
            '\u{9}'
                | '\u{a}'
                | '\u{d}'
                | '\u{20}'..='\u{d7ff}'
                | '\u{e000}'..='\u{fffd}'
                | '\u{10000}'..='\u{10ffff}'
        )
    }) {
        Ok(())
    } else {
        Err(invalid(
            "semantic slide XML contains an invalid XML character",
        ))
    }
}

fn validate_xml_comment(comment: &str) -> Result<()> {
    validate_xml_characters(comment)?;
    if comment.contains("--") || comment.ends_with('-') {
        return Err(invalid("semantic slide XML contains an invalid comment"));
    }
    Ok(())
}

fn semantic_mce_limits() -> MceLimits {
    MceLimits {
        max_input_bytes: MAX_SEMANTIC_TEXT_RAW_XML_BYTES,
        max_output_bytes: MAX_SEMANTIC_TEXT_PROCESSED_XML_BYTES,
        max_depth: MAX_SEMANTIC_TEXT_DEPTH,
        max_namespace_bindings: 4096,
        max_directive_tokens: 4096,
        max_choices_per_alternate: 1024,
    }
}

/// A reader configured exactly as both semantic-text passes configure theirs.
fn semantic_text_reader(xml: &[u8]) -> NsReader<&[u8]> {
    let mut reader = NsReader::from_reader(xml);
    reader.config_mut().trim_text(false);
    reader.config_mut().check_end_names = true;
    reader
}

/// Raw-byte validation of one slide for semantic text, one event at a time.
///
/// [`Self::observe`] applies to each event exactly the checks the raw scan
/// applies, in the same order, so the stand-alone scan and the single-pass
/// route refuse the same event with the same error.
#[derive(Default)]
struct RawTextScan {
    budget: SemanticTextXmlBudget,
    root_seen: bool,
    declaration_seen: bool,
    document_event_seen: bool,
}

impl RawTextScan {
    /// Validate one event; returns `true` once the document is complete.
    ///
    /// `namespace` is the element namespace `read_resolved_event` returned
    /// with `event`. For a start, empty or end element that is the value
    /// `resolve_element(name)` returns against the same bindings, which is
    /// what the scan validates.
    fn observe(
        &mut self,
        reader: &NsReader<&[u8]>,
        namespace: &ResolveResult<'_>,
        event: &Event<'_>,
    ) -> Result<bool> {
        let declaration_is_first = !self.document_event_seen;
        if !matches!(event, Event::Eof) {
            self.document_event_seen = true;
        }
        self.budget.observe_event(semantic_event_bytes(event))?;
        match event {
            Event::Start(element) => {
                validate_semantic_attributes(element)?;
                validate_semantic_attribute_names(reader, element)?;
                validate_semantic_element_namespace(namespace)?;
                if self.budget.depth == 0 {
                    if self.root_seen {
                        return Err(invalid("semantic slide XML has multiple roots"));
                    }
                    self.root_seen = true;
                }
                self.budget.start()?;
            },
            Event::Empty(element) => {
                validate_semantic_attributes(element)?;
                validate_semantic_attribute_names(reader, element)?;
                validate_semantic_element_namespace(namespace)?;
                if self.budget.depth == 0 {
                    if self.root_seen {
                        return Err(invalid("semantic slide XML has multiple roots"));
                    }
                    self.root_seen = true;
                }
            },
            Event::End(_) => {
                validate_semantic_element_namespace(namespace)?;
                self.budget.end()?;
            },
            Event::DocType(_) => {
                return Err(invalid("DTD declarations are not permitted in slide text"));
            },
            Event::PI(_) => {
                return Err(invalid(
                    "processing instructions are not permitted in slide text",
                ));
            },
            Event::Decl(_) => {
                if self.declaration_seen || !declaration_is_first || self.root_seen {
                    return Err(invalid("XML declarations must be the first document event"));
                }
                self.declaration_seen = true;
            },
            Event::Text(text) => {
                let decoded = text
                    .decode()
                    .map_err(|error| Error::Xml(error.to_string()))?;
                validate_xml_characters(&decoded)?;
                if decoded.len() > MAX_SEMANTIC_TEXT_EVENT_BYTES {
                    return Err(Error::Limit {
                        resource: "semantic slide decoded text event bytes",
                        limit: MAX_SEMANTIC_TEXT_EVENT_BYTES,
                    });
                }
                if self.budget.depth == 0 && !decoded.as_bytes().iter().all(u8::is_ascii_whitespace)
                {
                    return Err(invalid("semantic slide XML has text outside its root"));
                }
            },
            Event::CData(_) if self.budget.depth == 0 => {
                return Err(invalid("slide XML has CDATA outside its document root"));
            },
            Event::CData(text) => {
                let decoded = text
                    .decode()
                    .map_err(|error| Error::Xml(error.to_string()))?;
                validate_xml_characters(&decoded)?;
                if decoded.len() > MAX_SEMANTIC_TEXT_EVENT_BYTES {
                    return Err(Error::Limit {
                        resource: "semantic slide decoded text event bytes",
                        limit: MAX_SEMANTIC_TEXT_EVENT_BYTES,
                    });
                }
            },
            Event::Comment(comment) => {
                let decoded = comment
                    .decode()
                    .map_err(|error| Error::Xml(error.to_string()))?;
                validate_xml_comment(&decoded)?;
            },
            Event::GeneralRef(reference) => {
                if self.budget.depth == 0 {
                    return Err(invalid("XML entity reference is outside the document root"));
                }
                if reference.as_ref().len() > MAX_SEMANTIC_TEXT_REFERENCE_BYTES {
                    return Err(Error::Limit {
                        resource: "semantic slide XML reference bytes",
                        limit: MAX_SEMANTIC_TEXT_REFERENCE_BYTES,
                    });
                }
            },
            Event::Eof => {
                if !self.root_seen {
                    return Err(invalid("semantic slide XML lacks an element root"));
                }
                if self.budget.depth != 0 {
                    return Err(invalid("semantic slide XML has unbalanced elements"));
                }
                return Ok(true);
            },
        }
        Ok(false)
    }
}

fn check_semantic_text_raw_len(xml: &[u8]) -> Result<()> {
    if xml.len() > MAX_SEMANTIC_TEXT_RAW_XML_BYTES {
        return Err(Error::Limit {
            resource: "semantic slide raw XML bytes",
            limit: MAX_SEMANTIC_TEXT_RAW_XML_BYTES,
        });
    }
    Ok(())
}

fn scan_raw_semantic_text_xml(xml: &[u8]) -> Result<()> {
    check_semantic_text_raw_len(xml)?;
    let mut reader = semantic_text_reader(xml);
    let mut scan = RawTextScan::default();
    loop {
        let event = reader
            .read_event()
            .map_err(|error| Error::Xml(error.to_string()))?;
        // `read_resolved_event` is exactly these two steps; splitting them
        // lets the namespace borrow the reader shared, so the element's
        // attribute names can be resolved while it is held.
        let (namespace, event) = reader.resolver().resolve_event(event);
        if scan.observe(&reader, &namespace, &event)? {
            return Ok(());
        }
    }
}

fn is_drawingml_element(namespace: &ResolveResult<'_>, name: QName<'_>, local_name: &[u8]) -> bool {
    if name.local_name().as_ref() != local_name {
        return false;
    }
    matches!(
        namespace,
        ResolveResult::Bound(Namespace(value))
            if *value == DRAWINGML_NAMESPACE || *value == STRICT_DRAWINGML_NAMESPACE
    )
}

fn is_presentationml_slide(namespace: &ResolveResult<'_>, name: QName<'_>) -> bool {
    if name.local_name().as_ref() != b"sld" {
        return false;
    }
    matches!(
        namespace,
        ResolveResult::Bound(Namespace(value))
            if *value == crate::namespace::PRESENTATIONML_NAMESPACE
                || *value == crate::namespace::STRICT_PRESENTATIONML_NAMESPACE
    )
}

struct SemanticTextParser<'a> {
    budget: SemanticTextXmlBudget,
    output: String,
    root_seen: bool,
    declaration_seen: bool,
    document_event_seen: bool,
    active_text_depth: Option<usize>,
    text_has_payload: bool,
    runs: usize,
    objects: usize,
    paragraph_separator: &'a str,
    /// Whether the raw scan validates each event before this parser reads it.
    ///
    /// Only the single-pass route sets it. There the raw scan has already
    /// accepted the element's attribute values and attribute-name prefixes
    /// for this very event against these very bindings; both checks are pure
    /// functions of the two, so repeating them cannot fail and is skipped.
    /// Every check that is not a repetition of the raw scan still runs.
    raw_validated: bool,
}

impl<'a> SemanticTextParser<'a> {
    fn new(paragraph_separator: &'a str) -> Self {
        Self {
            budget: SemanticTextXmlBudget::default(),
            output: String::new(),
            root_seen: false,
            declaration_seen: false,
            document_event_seen: false,
            active_text_depth: None,
            text_has_payload: false,
            runs: 0,
            objects: 0,
            paragraph_separator,
            raw_validated: false,
        }
    }

    /// The parser for the single-pass route, whose raw scan validates every
    /// event first.
    fn after_raw_scan(paragraph_separator: &'a str) -> Self {
        Self {
            raw_validated: true,
            ..Self::new(paragraph_separator)
        }
    }

    /// The attribute-value check of the semantic pass, which on the
    /// single-pass route repeats the raw scan's check of the same element.
    fn validate_attributes(&self, element: &quick_xml::events::BytesStart<'_>) -> Result<()> {
        if self.raw_validated {
            debug_assert!(
                validate_semantic_attributes(element).is_ok(),
                "the raw scan accepted attributes the semantic pass refuses"
            );
            return Ok(());
        }
        validate_semantic_attributes(element)
    }

    /// The attribute-name check of the semantic pass, which on the
    /// single-pass route repeats the raw scan's check of the same element
    /// against the same bindings.
    fn validate_attribute_names(
        &self,
        reader: &NsReader<&[u8]>,
        element: &quick_xml::events::BytesStart<'_>,
    ) -> Result<()> {
        if self.raw_validated {
            debug_assert!(
                validate_semantic_attribute_names(reader, element).is_ok(),
                "the raw scan accepted attribute names the semantic pass refuses"
            );
            return Ok(());
        }
        validate_semantic_attribute_names(reader, element)
    }
}

impl<'a> SemanticTextParser<'a> {
    fn increment(value: &mut usize, limit: usize, resource: &'static str) -> Result<()> {
        *value = value
            .checked_add(1)
            .ok_or_else(|| Error::Invalid(format!("{resource} counter overflow")))?;
        if *value > limit {
            return Err(Error::Limit { resource, limit });
        }
        Ok(())
    }

    fn append_output(&mut self, value: &str) -> Result<()> {
        let observed = self.output.len().checked_add(value.len()).ok_or_else(|| {
            Error::Invalid("semantic slide decoded text length overflow".to_string())
        })?;
        if observed > MAX_SEMANTIC_TEXT_BYTES {
            return Err(Error::Limit {
                resource: "semantic slide decoded text bytes",
                limit: MAX_SEMANTIC_TEXT_BYTES,
            });
        }
        self.output
            .try_reserve(value.len())
            .map_err(|source| Error::Allocation {
                resource: "semantic slide decoded text",
                source,
            })?;
        self.output.push_str(value);
        Ok(())
    }

    fn append_text_fragment(&mut self, value: &str) -> Result<()> {
        if value.is_empty() {
            return Ok(());
        }
        if !self.text_has_payload {
            if !self.output.is_empty() {
                self.append_output(self.paragraph_separator)?;
            }
            self.text_has_payload = true;
        }
        self.append_output(value)
    }

    fn finish_text(&mut self) -> Result<()> {
        if !self.text_has_payload && !self.output.is_empty() {
            self.append_output(self.paragraph_separator)?;
        }
        self.active_text_depth = None;
        self.text_has_payload = false;
        Ok(())
    }

    fn start_element(
        &mut self,
        namespace: &ResolveResult<'_>,
        element: &quick_xml::events::BytesStart<'_>,
    ) -> Result<()> {
        self.validate_attributes(element)?;
        if element.name().local_name().as_ref() == b"t" {
            if !is_drawingml_element(namespace, element.name(), b"t") {
                return Err(invalid(
                    "foreign text element is not a DrawingML a:t element",
                ));
            }
            if self.active_text_depth.is_some() {
                return Err(invalid("nested DrawingML text elements are not permitted"));
            }
            Self::increment(
                &mut self.objects,
                MAX_SEMANTIC_TEXT_OBJECTS,
                "semantic slide text objects",
            )?;
            self.active_text_depth = Some(self.budget.depth);
            self.text_has_payload = false;
        } else if is_drawingml_element(namespace, element.name(), b"r") {
            Self::increment(
                &mut self.runs,
                MAX_SEMANTIC_TEXT_RUNS,
                "semantic slide text runs",
            )?;
        }
        Ok(())
    }

    fn empty_element(
        &mut self,
        namespace: &ResolveResult<'_>,
        element: &quick_xml::events::BytesStart<'_>,
    ) -> Result<()> {
        self.validate_attributes(element)?;
        if element.name().local_name().as_ref() == b"t" {
            if !is_drawingml_element(namespace, element.name(), b"t") {
                return Err(invalid(
                    "foreign text element is not a DrawingML a:t element",
                ));
            }
            if self.active_text_depth.is_some() {
                return Err(invalid("nested DrawingML text elements are not permitted"));
            }
            Self::increment(
                &mut self.objects,
                MAX_SEMANTIC_TEXT_OBJECTS,
                "semantic slide text objects",
            )?;
            if !self.output.is_empty() {
                self.append_output(self.paragraph_separator)?;
            }
        } else if is_drawingml_element(namespace, element.name(), b"r") {
            Self::increment(
                &mut self.runs,
                MAX_SEMANTIC_TEXT_RUNS,
                "semantic slide text runs",
            )?;
        }
        Ok(())
    }

    fn end_element(
        &mut self,
        namespace: &ResolveResult<'_>,
        element: &quick_xml::events::BytesEnd<'_>,
    ) -> Result<()> {
        if element.name().local_name().as_ref() == b"t" {
            if !is_drawingml_element(namespace, element.name(), b"t") {
                return Err(invalid(
                    "foreign text element is not a DrawingML a:t element",
                ));
            }
            if self.active_text_depth != Some(self.budget.depth) {
                return Err(invalid("unbalanced DrawingML text element"));
            }
            self.finish_text()?
        }
        Ok(())
    }

    fn text_event(&mut self, text: &quick_xml::events::BytesText<'_>) -> Result<()> {
        let decoded = text
            .decode()
            .map_err(|error| Error::Xml(error.to_string()))?;
        let decoded =
            quick_xml::escape::unescape(&decoded).map_err(|error| Error::Xml(error.to_string()))?;
        validate_xml_characters(&decoded)?;
        if decoded.len() > MAX_SEMANTIC_TEXT_EVENT_BYTES {
            return Err(Error::Limit {
                resource: "semantic slide decoded text event bytes",
                limit: MAX_SEMANTIC_TEXT_EVENT_BYTES,
            });
        }
        self.append_text_fragment(&decoded)
    }

    fn cdata_event(&mut self, text: &quick_xml::events::BytesCData<'_>) -> Result<()> {
        let decoded = text
            .decode()
            .map_err(|error| Error::Xml(error.to_string()))?;
        validate_xml_characters(&decoded)?;
        if decoded.len() > MAX_SEMANTIC_TEXT_EVENT_BYTES {
            return Err(Error::Limit {
                resource: "semantic slide decoded text event bytes",
                limit: MAX_SEMANTIC_TEXT_EVENT_BYTES,
            });
        }
        self.append_text_fragment(&decoded)
    }

    fn reference_event(&mut self, reference: &quick_xml::events::BytesRef<'_>) -> Result<()> {
        if reference.as_ref().len() > MAX_SEMANTIC_TEXT_REFERENCE_BYTES {
            return Err(Error::Limit {
                resource: "semantic slide XML reference bytes",
                limit: MAX_SEMANTIC_TEXT_REFERENCE_BYTES,
            });
        }
        let decoded = litchi_ooxml_common::xml::decode_xml_reference(reference)
            .map_err(|error| Error::Xml(error.to_string()))?;
        validate_xml_characters(&decoded)?;
        if decoded.len() > MAX_SEMANTIC_TEXT_EVENT_BYTES {
            return Err(Error::Limit {
                resource: "semantic slide decoded reference bytes",
                limit: MAX_SEMANTIC_TEXT_EVENT_BYTES,
            });
        }
        self.append_text_fragment(&decoded)
    }

    /// Consume one event read from `reader`, with the element-name check the
    /// semantic pass applies before the event itself.
    ///
    /// `namespace` is the element namespace `read_resolved_event` returned,
    /// the value `resolve_element(name)` returns against the same bindings.
    fn consume_event(
        &mut self,
        reader: &NsReader<&[u8]>,
        namespace: ResolveResult<'_>,
        event: Event<'_>,
    ) -> Result<bool> {
        match event {
            Event::Start(element) => {
                self.validate_attribute_names(reader, &element)?;
                self.consume(namespace, Event::Start(element))
            },
            Event::Empty(element) => {
                self.validate_attribute_names(reader, &element)?;
                self.consume(namespace, Event::Empty(element))
            },
            event => self.consume(namespace, event),
        }
    }

    fn consume(&mut self, namespace: ResolveResult<'_>, event: Event<'_>) -> Result<bool> {
        self.budget.observe_event(semantic_event_bytes(&event))?;
        let declaration_is_first = !self.document_event_seen;
        if !matches!(&event, Event::Eof) {
            self.document_event_seen = true;
        }
        match event {
            Event::Start(element) => {
                validate_semantic_element_namespace(&namespace)?;
                if self.budget.depth == 0 {
                    if self.root_seen || !is_presentationml_slide(&namespace, element.name()) {
                        return Err(invalid("semantic slide XML has an invalid root"));
                    }
                    self.root_seen = true;
                }
                self.budget.start()?;
                self.start_element(&namespace, &element)?;
            },
            Event::Empty(element) => {
                validate_semantic_element_namespace(&namespace)?;
                if self.budget.depth == 0 {
                    if self.root_seen || !is_presentationml_slide(&namespace, element.name()) {
                        return Err(invalid("semantic slide XML has an invalid root"));
                    }
                    self.root_seen = true;
                }
                self.empty_element(&namespace, &element)?;
            },
            Event::End(element) => {
                validate_semantic_element_namespace(&namespace)?;
                self.end_element(&namespace, &element)?;
                self.budget.end()?;
            },
            Event::Text(text) if self.active_text_depth.is_some() => {
                self.text_event(&text)?;
            },
            Event::CData(text) if self.active_text_depth.is_some() => {
                self.cdata_event(&text)?;
            },
            Event::GeneralRef(reference) if self.active_text_depth.is_some() => {
                self.reference_event(&reference)?;
            },
            Event::Decl(_) => {
                if self.declaration_seen || !declaration_is_first || self.root_seen {
                    return Err(invalid("XML declarations must be the first document event"));
                }
                self.declaration_seen = true;
            },
            Event::Text(text) if self.budget.depth == 0 => {
                let decoded = text
                    .decode()
                    .map_err(|error| Error::Xml(error.to_string()))?;
                validate_xml_characters(&decoded)?;
                if decoded.len() > MAX_SEMANTIC_TEXT_EVENT_BYTES {
                    return Err(Error::Limit {
                        resource: "semantic slide decoded text event bytes",
                        limit: MAX_SEMANTIC_TEXT_EVENT_BYTES,
                    });
                }
                if !decoded.as_bytes().iter().all(u8::is_ascii_whitespace) {
                    return Err(invalid("semantic slide XML has text outside its root"));
                }
            },
            Event::Text(text) => {
                let decoded = text
                    .decode()
                    .map_err(|error| Error::Xml(error.to_string()))?;
                validate_xml_characters(&decoded)?;
                if decoded.len() > MAX_SEMANTIC_TEXT_EVENT_BYTES {
                    return Err(Error::Limit {
                        resource: "semantic slide decoded text event bytes",
                        limit: MAX_SEMANTIC_TEXT_EVENT_BYTES,
                    });
                }
            },
            Event::CData(_) => {
                return Err(invalid("slide XML has CDATA outside DrawingML text"));
            },
            Event::GeneralRef(_) => {
                return Err(invalid("XML entity reference is outside DrawingML text"));
            },
            Event::DocType(_) => {
                return Err(invalid("DTD declarations are not permitted in slide text"));
            },
            Event::PI(_) => {
                return Err(invalid(
                    "processing instructions are not permitted in slide text",
                ));
            },
            Event::Comment(comment) => {
                let decoded = comment
                    .decode()
                    .map_err(|error| Error::Xml(error.to_string()))?;
                validate_xml_comment(&decoded)?;
            },
            Event::Eof => {
                if !self.root_seen || self.budget.depth != 0 || self.active_text_depth.is_some() {
                    return Err(invalid("semantic slide XML has unbalanced elements"));
                }
                return Ok(true);
            },
        }
        Ok(false)
    }
}

fn semantic_text_from_part(part: &dyn Part, paragraph_separator: &str) -> Result<String> {
    let raw = part.blob();
    check_semantic_text_raw_len(raw)?;
    // Markup-compatibility processing is a pure function of the raw bytes,
    // so running it before the raw scan changes no value; only the order in
    // which failures are reported could move, and both routes below restore
    // the scan-first order.
    let processed =
        process_markup_compatibility(raw, &Capabilities::ooxml_baseline(), &semantic_mce_limits());
    // The single pass drops the processed-size check below. A borrowed MCE
    // output is the raw slice, already within the raw ceiling, so that check
    // cannot fire while the processed ceiling is at least the raw one (both
    // are 64 MiB). Changing either limit must keep this ordering or restore
    // the check on the single-pass route.
    const _: () = assert!(
        MAX_SEMANTIC_TEXT_PROCESSED_XML_BYTES >= MAX_SEMANTIC_TEXT_RAW_XML_BYTES,
        "the single-pass semantic text route relies on the processed ceiling covering the raw one"
    );
    if let Ok(output) = &processed
        && let Cow::Borrowed(unchanged) = &output.xml
        && std::ptr::eq(*unchanged, raw)
    {
        // A marker-free slide: the processed bytes are the raw bytes, which
        // are within both ceilings, so the two passes would read one event
        // stream. Read it once.
        return semantic_text_single_pass(raw, paragraph_separator);
    }
    scan_raw_semantic_text_xml(raw)?;
    let processed = processed?;
    if processed.xml.len() > MAX_SEMANTIC_TEXT_PROCESSED_XML_BYTES {
        return Err(Error::Limit {
            resource: "semantic slide processed XML bytes",
            limit: MAX_SEMANTIC_TEXT_PROCESSED_XML_BYTES,
        });
    }
    parse_semantic_text(processed.xml.as_ref(), paragraph_separator)
}

/// The semantic pass over already-scanned, MCE-processed slide XML.
fn parse_semantic_text(xml: &[u8], paragraph_separator: &str) -> Result<String> {
    let mut reader = semantic_text_reader(xml);
    let mut parser = SemanticTextParser::new(paragraph_separator);
    loop {
        let event = reader
            .read_event()
            .map_err(|error| Error::Xml(error.to_string()))?;
        // `read_resolved_event` is exactly these two steps; splitting them
        // lets the namespace borrow the reader shared, so the element's
        // attribute names can be resolved while it is held.
        let (namespace, event) = reader.resolver().resolve_event(event);
        if parser.consume_event(&reader, namespace, event)? {
            return Ok(parser.output);
        }
    }
}

/// Both passes over one event stream, for a slide whose processed bytes are
/// its raw bytes.
///
/// The raw scan refuses before the semantic pass ever starts, so a raw
/// failure at any event outranks a semantic failure at any event, earlier or
/// later. Each event is therefore validated by the raw scan first, and a raw
/// failure is returned at once; a semantic failure is only recorded, the
/// semantic pass stops, and the raw scan continues to the end, where the
/// recorded failure is returned only if the raw scan passed. The semantic
/// pass consumes exactly the prefix of the stream it would have consumed
/// alone, so its value and its first failure are unchanged.
fn semantic_text_single_pass(xml: &[u8], paragraph_separator: &str) -> Result<String> {
    let mut reader = semantic_text_reader(xml);
    let mut scan = RawTextScan::default();
    let mut parser = SemanticTextParser::after_raw_scan(paragraph_separator);
    let mut deferred = None;
    loop {
        let event = reader
            .read_event()
            .map_err(|error| Error::Xml(error.to_string()))?;
        // `read_resolved_event` is exactly these two steps; splitting them
        // lets the namespace borrow the reader shared, so the element's
        // attribute names can be resolved while it is held.
        let (namespace, event) = reader.resolver().resolve_event(event);
        let complete = scan.observe(&reader, &namespace, &event)?;
        if deferred.is_none() {
            match parser.consume_event(&reader, namespace, event) {
                Ok(finished) => debug_assert_eq!(
                    finished, complete,
                    "the semantic pass finished on a different event from the raw scan"
                ),
                Err(error) => deferred = Some(error),
            }
        }
        if complete {
            return match deferred {
                Some(error) => Err(error),
                None => Ok(parser.output),
            };
        }
    }
}

#[cfg(test)]
thread_local! {
    /// Proved slide-root classifications reused on this thread, for tests
    /// that must show a capture skipped the scan rather than only that its
    /// result is right.
    static PROVED_ROOT_HITS: std::cell::Cell<usize> = const { std::cell::Cell::new(0) };
    /// Complete notes-root scans run on this thread, including the debug
    /// re-derivation of every reused classification.
    static ROOT_SCANS: std::cell::Cell<usize> = const { std::cell::Cell::new(0) };
}

/// Return and reset this thread's count of reused slide-root classifications.
#[cfg(test)]
pub(crate) fn take_proved_root_hits() -> usize {
    PROVED_ROOT_HITS.with(|hits| hits.replace(0))
}

/// Return and reset this thread's count of complete notes-root scans.
#[cfg(test)]
pub(crate) fn take_root_scans() -> usize {
    ROOT_SCANS.with(|scans| scans.replace(0))
}

fn text_and_name_from_part(part: &dyn Part) -> Result<(String, String)> {
    // Keep the established individual projections as the semantic source of
    // truth. Text uses the same bounded namespace-aware parser as the sink,
    // while `name` preserves its early-return namespace behavior. Source-
    // backed callers still materialize the selected Part payload only once;
    // only the processed XML projections are repeated.
    let text = text_from_part(part)?.unwrap_or_default();
    let name = c_sld_name(part)?.unwrap_or_else(|| part.partname().to_string());
    Ok((text, name))
}

/// Borrowed view of a `PresentationML` slide part.
#[derive(Clone, Copy)]
pub struct SlidePart<'a> {
    part: &'a dyn Part,
}

impl<'a> SlidePart<'a> {
    pub(crate) const fn semantic_text_raw_xml_limit() -> usize {
        MAX_SEMANTIC_TEXT_RAW_XML_BYTES
    }

    pub(crate) fn semantic_text_from_part(
        part: &'a dyn Part,
        paragraph_separator: &str,
    ) -> Result<String> {
        semantic_text_from_part(part, paragraph_separator)
    }

    /// Validate and wrap a slide part.
    ///
    /// # Errors
    ///
    /// Returns an error if the input cannot be read or is malformed.
    pub fn from_part(part: &'a dyn Part) -> Result<Self> {
        validate_content_type(part, ct::PML_SLIDE)?;
        if root_name(part)? != "sld" {
            return Err(invalid("slide part does not have a p:sld root"));
        }
        Ok(Self { part })
    }

    /// Validate a slide and project its producer name from one temporary MCE
    /// result. The optional notes proof is classified while that result is
    /// alive, then the processed bytes are dropped before a part-name
    /// fallback is allocated.
    pub(crate) fn from_part_with_name(
        part: &'a dyn Part,
        collect_notes_proof: bool,
    ) -> Result<(Self, Result<String>, Option<SlideRootProof<'a>>)> {
        validate_content_type(part, ct::PML_SLIDE)?;
        let xml = processed_xml_with_source(part)?;
        Self::finish_from_processed(part, collect_notes_proof, xml, None)
    }

    /// Capture variant that may reuse a default-profile transformed slide
    /// projection already retained by the opened snapshot, and a notes-root
    /// classification an earlier capture proved for the same payload
    /// allocation. Every source, root, name, and MCE limit check still runs
    /// for this call; only the complete notes-root scan of a payload whose
    /// classification is already proved is not repeated.
    pub(crate) fn from_part_with_name_with_capture<'parent>(
        part: &'a dyn Part,
        collect_notes_proof: bool,
        capture: &mut MceCapture<'a, 'parent>,
        proved_roots: Option<&crate::notes::SlideRootMemo>,
    ) -> Result<(Self, Result<String>, Option<SlideRootProof<'a>>)> {
        validate_content_type(part, ct::PML_SLIDE)?;
        let xml = processed_xml_with_capture(part, capture)?;
        Self::finish_from_processed(part, collect_notes_proof, xml, proved_roots)
    }

    fn finish_from_processed(
        part: &'a dyn Part,
        collect_notes_proof: bool,
        xml: super::ProcessedXml<'a>,
        proved_roots: Option<&crate::notes::SlideRootMemo>,
    ) -> Result<(Self, Result<String>, Option<SlideRootProof<'a>>)> {
        if root_name_from_xml(xml.processed.as_ref())? != "sld" {
            return Err(invalid("slide part does not have a p:sld root"));
        }
        let name = c_sld_name_from_xml(xml.processed.as_ref());
        let notes_proof = (collect_notes_proof && name.is_ok()).then(|| {
            let scan = || {
                #[cfg(test)]
                ROOT_SCANS.with(|scans| scans.set(scans.get() + 1));
                crate::notes::root_conformance_from_processed(
                    xml.processed.as_ref(),
                    xml.source.len(),
                    crate::notes::MAX_SLIDE_XML,
                    "sld",
                )
            };
            // The memo is keyed on the exact raw observation MCE just read, so
            // a hit names these very bytes; a miss is the ordinary scan.
            let conformance = match proved_roots.and_then(|memo| memo.lookup(xml.source)) {
                Some(proved) => {
                    debug_assert_eq!(
                        scan(),
                        Some(proved),
                        "a proved slide-root classification answered for different bytes"
                    );
                    #[cfg(test)]
                    PROVED_ROOT_HITS.with(|hits| hits.set(hits.get() + 1));
                    Some(proved)
                },
                None => scan(),
            };
            SlideRootProof::new(xml.source, conformance)
        });
        drop(xml);
        let name = name.map(|name| name.unwrap_or_else(|| part.partname().to_string()));
        Ok((Self { part }, name, notes_proof))
    }

    /// The underlying OPC part.
    #[inline]
    #[must_use]
    pub fn part(&self) -> &'a dyn Part {
        self.part
    }

    /// Producer-visible slide name, if present.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn name(&self) -> Result<String> {
        Ok(c_sld_name(self.part)?.unwrap_or_else(|| self.part.partname().to_string()))
    }

    /// Whether the slide is marked hidden by its root `show` attribute.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn is_hidden(&self) -> Result<bool> {
        Ok(!root_bool(self.part, b"show", "slide show", true)?)
    }

    /// Flatten `DrawingML` text runs in source order.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn text(&self) -> Result<String> {
        Ok(text_from_part(self.part)?.unwrap_or_default())
    }

    /// Read the producer-visible name and flattened text while preserving the
    /// exact semantics of [`Self::name`] and [`Self::text`].
    ///
    /// This combined projection is useful to source-backed callers that need
    /// both values. Source-backed callers materialize the selected Part only
    /// once; the two established processed-XML projections retain their
    /// independent reader behavior.
    pub fn text_and_name(&self) -> Result<(String, String)> {
        text_and_name_from_part(self.part)
    }

    /// Build the bounded borrowed shape scene for this slide.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn shapes(&self) -> Result<Scene<'a>> {
        Scene::read(self.part.blob())
    }

    /// Resolve ordinary `DrawingML` chart parts related to this slide.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn charts(&self, package: &'a OpcPackage) -> Result<Vec<crate::chart::Part<'a>>> {
        crate::chart::related(package, self.part)
    }

    /// Resolve Microsoft `ChartEx` parts related to this slide.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn chart_extensions(
        &self,
        package: &'a OpcPackage,
    ) -> Result<Vec<crate::chart::extension::Part<'a>>> {
        crate::chart::extension::related(package, self.part)
    }

    /// Resolve the optional legacy comments list attached to this slide.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn comments(
        &self,
        package: &'a OpcPackage,
    ) -> Result<Option<crate::comments::ListPart<'a>>> {
        let part = related_part_by_type(
            package,
            self.part,
            crate::comments::COMMENTS_REL,
            "comments",
            ct::PML_COMMENTS,
        )?;
        part.map(crate::comments::ListPart::from_part).transpose()
    }

    /// Resolve the slide's optional layout relationship.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn layout(&self, package: &'a OpcPackage) -> Result<Option<SlideLayoutPart<'a>>> {
        related_part_by_type(
            package,
            self.part,
            rt::SLIDE_LAYOUT,
            "slideLayout",
            ct::PML_SLIDE_LAYOUT,
        )?
        .map(SlideLayoutPart::from_part)
        .transpose()
    }
}

/// Borrowed view of a `PresentationML` slide-layout part.
#[derive(Clone, Copy)]
pub struct SlideLayoutPart<'a> {
    part: &'a dyn Part,
}

impl<'a> SlideLayoutPart<'a> {
    /// Validate and wrap a slide-layout part.
    ///
    /// # Errors
    ///
    /// Returns an error if the input cannot be read or is malformed.
    pub fn from_part(part: &'a dyn Part) -> Result<Self> {
        validate_content_type(part, ct::PML_SLIDE_LAYOUT)?;
        if root_name(part)? != "sldLayout" {
            return Err(invalid(
                "slide-layout part does not have a p:sldLayout root",
            ));
        }
        Ok(Self { part })
    }

    /// The underlying OPC part.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    #[inline]
    #[must_use]
    pub fn part(&self) -> &'a dyn Part {
        self.part
    }

    /// Producer-visible layout name, if present.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn name(&self) -> Result<String> {
        Ok(c_sld_name(self.part)?.unwrap_or_else(|| self.part.partname().to_string()))
    }

    /// Layout kind token from `p:sldLayout@type`, if present.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn kind(&self) -> Result<Option<String>> {
        let xml = processed_xml(self.part)?;
        let mut reader = Reader::from_reader(xml.as_ref());
        loop {
            match reader.read_event()? {
                Event::Start(element) | Event::Empty(element) => {
                    return Ok(litchi_ooxml_common::xml::unqualified_attribute_value(
                        &element,
                        b"type",
                        reader.decoder(),
                    )?);
                },
                Event::Decl(_) | Event::Comment(_) => {},
                _ => return Err(invalid("slide-layout part lacks an element root")),
            }
        }
    }

    /// Build the bounded borrowed shape scene for this layout.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn shapes(&self) -> Result<Scene<'a>> {
        Scene::read(self.part.blob())
    }

    /// Read the optional theme override attached to this layout.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn theme_override(
        &self,
        package: &'a OpcPackage,
    ) -> Result<Option<crate::shape::theme::Override>> {
        crate::shape::theme::package::load_override(package, self.part.partname().as_str())
    }

    /// Resolve the required slide-master relationship.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn master(&self, package: &'a OpcPackage) -> Result<SlideMasterPart<'a>> {
        let part = related_part_by_type(
            package,
            self.part,
            rt::SLIDE_MASTER,
            "slideMaster",
            ct::PML_SLIDE_MASTER,
        )?
        .ok_or_else(|| invalid("slide layout lacks its slide-master relationship"))?;
        SlideMasterPart::from_part(part)
    }
}

/// Borrowed view of a `PresentationML` slide-master part.
#[derive(Clone, Copy)]
pub struct SlideMasterPart<'a> {
    part: &'a dyn Part,
}

impl<'a> SlideMasterPart<'a> {
    /// Validate and wrap a slide-master part.
    ///
    /// # Errors
    ///
    /// Returns an error if the input cannot be read or is malformed.
    pub fn from_part(part: &'a dyn Part) -> Result<Self> {
        validate_content_type(part, ct::PML_SLIDE_MASTER)?;
        if root_name(part)? != "sldMaster" {
            return Err(invalid(
                "slide-master part does not have a p:sldMaster root",
            ));
        }
        Ok(Self { part })
    }

    /// The underlying OPC part.
    #[inline]
    #[must_use]
    pub fn part(&self) -> &'a dyn Part {
        self.part
    }

    /// Producer-visible master name, if present.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn name(&self) -> Result<String> {
        Ok(c_sld_name(self.part)?.unwrap_or_else(|| self.part.partname().to_string()))
    }

    /// Whether the master is marked preserved.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn is_preserved(&self) -> Result<bool> {
        root_bool(self.part, b"preserve", "slide-master preserve", false)
    }

    /// Build the bounded borrowed shape scene for this master.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn shapes(&self) -> Result<Scene<'a>> {
        Scene::read(self.part.blob())
    }

    /// Read the theme reached from this slide master.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn theme(
        &self,
        package: &'a OpcPackage,
    ) -> Result<Option<crate::shape::theme::ThemeSummary>> {
        let part = related_part_by_type(package, self.part, rt::THEME, "theme", ct::OFC_THEME)?;
        part.map(|part| crate::shape::theme::part::Part::from_part(part)?.read())
            .transpose()
    }

    /// Resolve the slide layouts listed by this master in XML order.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn layouts(&self, package: &'a OpcPackage) -> Result<Vec<SlideLayoutPart<'a>>> {
        let relationship_ids = layout_relationship_ids(self.part)?;
        let mut layouts = Vec::with_capacity(relationship_ids.len());
        for relationship_id in relationship_ids {
            let relationship = self.part.rels().get(&relationship_id).ok_or_else(|| {
                Error::Relationship(format!(
                    "slide master references missing slide-layout relationship '{relationship_id}'"
                ))
            })?;
            if relationship.is_external() {
                return Err(Error::Relationship(
                    "slide-layout relationship must be internal".into(),
                ));
            }
            if !super::is_relationship_type(relationship.reltype(), rt::SLIDE_LAYOUT, "slideLayout")
            {
                return Err(Error::Relationship(format!(
                    "relationship '{relationship_id}' is not a slide-layout relationship"
                )));
            }
            let target = relationship.target_partname()?;
            let part = package.get_part(&target)?;
            validate_content_type(part, ct::PML_SLIDE_LAYOUT)?;
            layouts.push(SlideLayoutPart::from_part(part)?);
        }
        Ok(layouts)
    }
}

#[cfg(test)]
mod tests {
    #![allow(
        clippy::expect_used,
        clippy::unwrap_used,
        reason = "focused low-level part tests use literal XML fixtures"
    )]

    use super::{SlidePart, semantic_text_from_part};
    use crate::Error;
    use litchi_opc::PackURI;
    use litchi_opc::constants::content_type as ct;
    use litchi_opc::part::BlobPart;

    // The two semantic-text passes as they stood before change 0743,
    // verbatim: a complete raw scan, then markup-compatibility processing,
    // then the semantic pass over the processed bytes. The single-pass route
    // must return exactly their value and exactly their first refusal.
    use super::{
        Capabilities, Event, MAX_SEMANTIC_TEXT_EVENT_BYTES, MAX_SEMANTIC_TEXT_PROCESSED_XML_BYTES,
        MAX_SEMANTIC_TEXT_RAW_XML_BYTES, MAX_SEMANTIC_TEXT_REFERENCE_BYTES, NsReader, Part, Result,
        SemanticTextParser, SemanticTextXmlBudget, invalid, process_markup_compatibility,
        semantic_event_bytes, semantic_mce_limits, validate_semantic_attribute_names,
        validate_semantic_attributes, validate_semantic_element_namespace, validate_xml_characters,
        validate_xml_comment,
    };

    fn oracle_scan_raw_semantic_text_xml(xml: &[u8]) -> Result<()> {
        if xml.len() > MAX_SEMANTIC_TEXT_RAW_XML_BYTES {
            return Err(Error::Limit {
                resource: "semantic slide raw XML bytes",
                limit: MAX_SEMANTIC_TEXT_RAW_XML_BYTES,
            });
        }

        let mut reader = NsReader::from_reader(xml);
        reader.config_mut().trim_text(false);
        reader.config_mut().check_end_names = true;
        let mut budget = SemanticTextXmlBudget::default();
        let mut root_seen = false;
        let mut declaration_seen = false;
        let mut document_event_seen = false;

        loop {
            let (namespace, event) = reader
                .read_resolved_event()
                .map_err(|error| Error::Xml(error.to_string()))?;
            let declaration_is_first = !document_event_seen;
            if !matches!(&event, Event::Eof) {
                document_event_seen = true;
            }
            budget.observe_event(semantic_event_bytes(&event))?;
            match event {
                Event::Start(element) => {
                    let _ = namespace;
                    validate_semantic_attributes(&element)?;
                    validate_semantic_attribute_names(&reader, &element)?;
                    let namespace = reader.resolver().resolve_element(element.name()).0;
                    validate_semantic_element_namespace(&namespace)?;
                    if budget.depth == 0 {
                        if root_seen {
                            return Err(invalid("semantic slide XML has multiple roots"));
                        }
                        root_seen = true;
                    }
                    budget.start()?;
                },
                Event::Empty(element) => {
                    let _ = namespace;
                    validate_semantic_attributes(&element)?;
                    validate_semantic_attribute_names(&reader, &element)?;
                    let namespace = reader.resolver().resolve_element(element.name()).0;
                    validate_semantic_element_namespace(&namespace)?;
                    if budget.depth == 0 {
                        if root_seen {
                            return Err(invalid("semantic slide XML has multiple roots"));
                        }
                        root_seen = true;
                    }
                },
                Event::End(_) => {
                    validate_semantic_element_namespace(&namespace)?;
                    budget.end()?;
                },
                Event::DocType(_) => {
                    return Err(invalid("DTD declarations are not permitted in slide text"));
                },
                Event::PI(_) => {
                    return Err(invalid(
                        "processing instructions are not permitted in slide text",
                    ));
                },
                Event::Decl(_) => {
                    if declaration_seen || !declaration_is_first || root_seen {
                        return Err(invalid("XML declarations must be the first document event"));
                    }
                    declaration_seen = true;
                },
                Event::Text(text) => {
                    let decoded = text
                        .decode()
                        .map_err(|error| Error::Xml(error.to_string()))?;
                    validate_xml_characters(&decoded)?;
                    if decoded.len() > MAX_SEMANTIC_TEXT_EVENT_BYTES {
                        return Err(Error::Limit {
                            resource: "semantic slide decoded text event bytes",
                            limit: MAX_SEMANTIC_TEXT_EVENT_BYTES,
                        });
                    }
                    if budget.depth == 0 && !decoded.as_bytes().iter().all(u8::is_ascii_whitespace)
                    {
                        return Err(invalid("semantic slide XML has text outside its root"));
                    }
                },
                Event::CData(_) if budget.depth == 0 => {
                    return Err(invalid("slide XML has CDATA outside its document root"));
                },
                Event::CData(text) => {
                    let decoded = text
                        .decode()
                        .map_err(|error| Error::Xml(error.to_string()))?;
                    validate_xml_characters(&decoded)?;
                    if decoded.len() > MAX_SEMANTIC_TEXT_EVENT_BYTES {
                        return Err(Error::Limit {
                            resource: "semantic slide decoded text event bytes",
                            limit: MAX_SEMANTIC_TEXT_EVENT_BYTES,
                        });
                    }
                },
                Event::Comment(comment) => {
                    let decoded = comment
                        .decode()
                        .map_err(|error| Error::Xml(error.to_string()))?;
                    validate_xml_comment(&decoded)?;
                },
                Event::GeneralRef(reference) => {
                    if budget.depth == 0 {
                        return Err(invalid("XML entity reference is outside the document root"));
                    }
                    if reference.as_ref().len() > MAX_SEMANTIC_TEXT_REFERENCE_BYTES {
                        return Err(Error::Limit {
                            resource: "semantic slide XML reference bytes",
                            limit: MAX_SEMANTIC_TEXT_REFERENCE_BYTES,
                        });
                    }
                },
                Event::Eof => {
                    if !root_seen {
                        return Err(invalid("semantic slide XML lacks an element root"));
                    }
                    if budget.depth != 0 {
                        return Err(invalid("semantic slide XML has unbalanced elements"));
                    }
                    return Ok(());
                },
            }
        }
    }

    fn oracle_semantic_text_from_part(
        part: &dyn Part,
        paragraph_separator: &str,
    ) -> Result<String> {
        let raw = part.blob();
        oracle_scan_raw_semantic_text_xml(raw)?;
        let processed = process_markup_compatibility(
            raw,
            &Capabilities::ooxml_baseline(),
            &semantic_mce_limits(),
        )?;
        if processed.xml.len() > MAX_SEMANTIC_TEXT_PROCESSED_XML_BYTES {
            return Err(Error::Limit {
                resource: "semantic slide processed XML bytes",
                limit: MAX_SEMANTIC_TEXT_PROCESSED_XML_BYTES,
            });
        }

        let mut reader = NsReader::from_reader(processed.xml.as_ref());
        reader.config_mut().trim_text(false);
        reader.config_mut().check_end_names = true;
        let mut parser = SemanticTextParser::new(paragraph_separator);
        loop {
            let (namespace, event) = reader
                .read_resolved_event()
                .map_err(|error| Error::Xml(error.to_string()))?;
            let finished = match event {
                Event::Start(element) => {
                    let _ = namespace;
                    validate_semantic_attribute_names(&reader, &element)?;
                    let namespace = reader.resolver().resolve_element(element.name()).0;
                    parser.consume(namespace, Event::Start(element))?
                },
                Event::Empty(element) => {
                    let _ = namespace;
                    validate_semantic_attribute_names(&reader, &element)?;
                    let namespace = reader.resolver().resolve_element(element.name()).0;
                    parser.consume(namespace, Event::Empty(element))?
                },
                event => parser.consume(namespace, event)?,
            };
            if finished {
                return Ok(parser.output);
            }
        }
    }

    fn slide_part(xml: &[u8]) -> BlobPart {
        BlobPart::new(
            PackURI::new("/ppt/slides/slide1.xml").unwrap(),
            ct::PML_SLIDE.to_owned(),
            xml.to_vec(),
        )
    }

    #[test]
    fn combined_text_and_name_matches_separate_reads_through_mce_and_unusual_text() {
        let xml = br#"<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:x="urn:producer-future" mc:Ignorable="x">
            <!-- producer formatting and an ignored extension are intentional -->
            <p:cSld name="  Producer &amp; Name  " x:future="retained"><p:spTree>
                <a:t> leading &amp; </a:t><x:future><a:t>ignored</a:t></x:future>
                <a:t><![CDATA[tail]]></a:t><a:t>two</a:t>
            </p:spTree></p:cSld>
        </p:sld>"#;
        let part = slide_part(xml);
        let slide = SlidePart::from_part(&part).unwrap();
        let separate = (slide.text().unwrap(), slide.name().unwrap());
        assert_eq!(slide.text_and_name().unwrap(), separate);
        assert_eq!(separate.0, " leading & \ntail\ntwo");
        assert_eq!(separate.1, "  Producer & Name  ");
    }

    #[test]
    fn late_reserved_prefix_rebinding_is_rejected_by_text_and_text_and_name() {
        let xml = br#"<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><p:cSld name="early"><p:spTree><a:t>text</a:t></p:spTree></p:cSld><p:extLst xmlns:xml="urn:invalid"/></p:sld>"#;
        let part = slide_part(xml);
        let slide = SlidePart::from_part(&part).unwrap();
        assert_eq!(slide.name().unwrap(), "early");

        let text_error = slide.text().unwrap_err();
        assert!(matches!(text_error, Error::Xml(_) | Error::Invalid(_)));

        let combined_error = slide.text_and_name().unwrap_err();
        assert!(matches!(combined_error, Error::Xml(_) | Error::Invalid(_)));
    }

    #[test]
    fn combined_text_and_name_preserves_missing_and_empty_name_semantics() {
        let fixtures = [
            (
                br#"<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:cSld><p:spTree/></p:cSld></p:sld>"#.as_slice(),
                "",
            ),
            (
                br#"<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:cSld name=""><p:spTree/></p:cSld></p:sld>"#.as_slice(),
                "",
            ),
            (
                br#"<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:spTree/></p:sld>"#.as_slice(),
                "/ppt/slides/slide1.xml",
            ),
        ];
        for (xml, expected_name) in fixtures {
            let part = slide_part(xml);
            let slide = SlidePart::from_part(&part).unwrap();
            let separate = (slide.text().unwrap(), slide.name().unwrap());
            assert_eq!(slide.text_and_name().unwrap(), separate);
            assert_eq!(separate, (String::new(), expected_name.to_owned()));
        }
    }

    #[test]
    fn combined_text_and_name_rejects_the_same_malformed_xml_as_separate_reads() {
        let part = slide_part(
            br#"<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:cSld><p:spTree><p:broken></p:cSld></p:sld>"#,
        );
        let slide = SlidePart::from_part(&part).unwrap();
        let name = slide.name();
        assert_eq!(name.unwrap(), "");
        let text_error = slide.text().unwrap_err().to_string();
        let combined_error = slide.text_and_name().unwrap_err().to_string();
        assert_eq!(combined_error, text_error);
    }

    #[test]
    fn semantic_text_enforces_the_cumulative_decoded_text_ceiling() {
        let chunk = "x".repeat(1024 * 1024);
        let mut xml = String::with_capacity(17 * (chunk.len() + 11) + 256);
        xml.push_str(
            r#"<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><p:cSld><p:spTree>"#,
        );
        for _ in 0..17 {
            xml.push_str("<a:t>");
            xml.push_str(&chunk);
            xml.push_str("</a:t>");
        }
        xml.push_str(r#"</p:spTree></p:cSld></p:sld>"#);

        let part = slide_part(xml.as_bytes());
        let error = semantic_text_from_part(&part, "\n").unwrap_err();
        assert!(matches!(
            error,
            Error::Limit {
                resource: "semantic slide decoded text bytes",
                ..
            }
        ));
    }

    const PML: &str = "http://schemas.openxmlformats.org/presentationml/2006/main";
    const DML: &str = "http://schemas.openxmlformats.org/drawingml/2006/main";
    const MCE: &str = "http://schemas.openxmlformats.org/markup-compatibility/2006";

    fn slide(body: &str) -> Vec<u8> {
        format!(r#"<p:sld xmlns:p="{PML}" xmlns:a="{DML}"><p:cSld><p:spTree>{body}</p:spTree></p:cSld></p:sld>"#)
            .into_bytes()
    }

    type TextOutcome = std::result::Result<String, String>;

    fn text_outcome(result: Result<String>) -> TextOutcome {
        result.map_err(|error| format!("{error:?}"))
    }

    /// Compare the routed semantic text with the three-pass oracle for one
    /// input and both separators. Returns (compared, accepted).
    fn assert_text_routes_agree(label: &str, xml: &[u8]) -> (usize, usize) {
        let part = slide_part(xml);
        let mut accepted = 0;
        for separator in ["\n", " | "] {
            let expected = text_outcome(oracle_semantic_text_from_part(&part, separator));
            let actual = text_outcome(semantic_text_from_part(&part, separator));
            assert_eq!(actual, expected, "{label}: separator {separator:?}");
            accepted += usize::from(expected.is_ok());
        }
        (2, accepted)
    }

    fn text_seeds() -> Vec<Vec<u8>> {
        let run = |text: &str| {
            format!("<p:sp><p:txBody><a:p><a:r><a:t>{text}</a:t></a:r></a:p></p:txBody></p:sp>")
        };
        let deep = format!("{}{}", "<a:g>".repeat(130), "</a:g>".repeat(130));
        let mut seeds: Vec<Vec<u8>> = vec![
            slide(&format!("{}{}", run("one &amp; two"), run("three"))),
            slide(&run("<![CDATA[raw <text>]]>")),
            slide(&run("x&#65;y&#x42;z")),
            slide(&format!("{}<a:br/><a:t/>{}", run("a"), run("b"))),
            slide(&run("<a:t>nested</a:t>")),
            slide(&format!("{}<!-- bad -- comment -->", run("<a:t>nested</a:t>"))),
            slide(&format!("<!-- bad -- comment -->{}", run("<a:t>nested</a:t>"))),
            slide(&format!("{}<p:x a=\"&#1;\"/>", run("<a:t>nested</a:t>"))),
            slide(&format!("<q:t>foreign</q:t>{}", run("x"))),
            slide(&format!("<p:t>foreign</p:t>{}", run("x"))),
            slide("<![CDATA[outside text]]>"),
            slide("&amp;"),
            slide(&run("&unknown;")),
            slide(&format!("{}<p:x q:a=\"1\"/>", run("x"))),
            slide(&format!("{}<p:x a=\"1\" a=\"2\"/>", run("x"))),
            slide(&deep),
            slide(&format!("<p:x>\u{1}</p:x>{}", run("y"))),
            format!(r#"<p:notSld xmlns:p="{PML}" xmlns:a="{DML}">{}</p:notSld>"#, run("x")).into_bytes(),
            format!(r#"<?xml version="1.0"?><p:sld xmlns:p="{PML}"/>"#).into_bytes(),
            format!(r#"<p:sld xmlns:p="{PML}"/><?xml version="1.0"?>"#).into_bytes(),
            format!(r#"<!DOCTYPE d><p:sld xmlns:p="{PML}"/>"#).into_bytes(),
            format!(r#"<p:sld xmlns:p="{PML}"><?pi x?></p:sld>"#).into_bytes(),
            format!(r#"<p:sld xmlns:p="{PML}"/>text"#).into_bytes(),
            format!(r#"<![CDATA[x]]><p:sld xmlns:p="{PML}"/>"#).into_bytes(),
            format!(r#"&amp;<p:sld xmlns:p="{PML}"/>"#).into_bytes(),
            format!(r#"<p:sld xmlns:p="{PML}"/><p:sld xmlns:p="{PML}"/>"#).into_bytes(),
            format!(r#"<p:sld xmlns:p="{PML}"><p:a></p:b></p:sld>"#).into_bytes(),
            format!(r#"<p:sld xmlns:p="{PML}">"#).into_bytes(),
            format!("\u{feff}<p:sld xmlns:p=\"{PML}\"/>").into_bytes(),
            format!(
                r#"<p:sld xmlns:p="{PML}" xmlns:a="{DML}" xmlns:mc="{MCE}" xmlns:x="urn:future" mc:Ignorable="x"><p:cSld><p:spTree>{}<x:ext><a:t>ignored</a:t></x:ext><mc:AlternateContent><mc:Choice Requires="x"><a:t>choice</a:t></mc:Choice><mc:Fallback><a:t>fallback</a:t></mc:Fallback></mc:AlternateContent></p:spTree></p:cSld></p:sld>"#,
                run("marked")
            )
            .into_bytes(),
            format!(
                r#"<p:sld xmlns:p="{PML}" xmlns:a="{DML}" xmlns:mc="{MCE}"><!-- bad -- --><mc:AlternateContent/></p:sld>"#
            )
            .into_bytes(),
            Vec::new(),
        ];
        let mut invalid_utf8 = slide(&run("ok"));
        invalid_utf8.splice(10..10, b"\xff\xfe".iter().copied());
        seeds.push(invalid_utf8);
        seeds
    }

    fn text_mutations(seed: &[u8], budget: usize) -> Vec<Vec<u8>> {
        const BYTES: &[u8] = b"<>&\"'/!?=:;\x00\x01\xff x";
        const SNIPPETS: &[&[u8]] = &[
            b"<![CDATA[c]]>",
            b"<?pi x?>",
            b"<!DOCTYPE d>",
            b"<!-- c -->",
            b"<!-- c -- d -->",
            b"<q:x/>",
            b"<a:t>t</a:t>",
            b"<a:t/>",
            b"&amp;",
            b"&bad;",
            b"&#1;",
            b"</a:t>",
            b" a=\"&#1;\"",
        ];
        let mut output = Vec::new();
        if seed.is_empty() {
            return output;
        }
        let step = (seed.len() / budget.max(1)).max(1);
        let mut state = 0x2545_f491_4f6c_dd1d_u64 ^ seed.len() as u64;
        let mut next = || {
            state ^= state << 13;
            state ^= state >> 7;
            state ^= state << 17;
            usize::try_from(state >> 8).unwrap_or(0)
        };
        for position in (0..seed.len()).step_by(step) {
            output.push(seed[..position].to_vec());
            let mut replaced = seed.to_vec();
            replaced[position] = BYTES[next() % BYTES.len()];
            output.push(replaced);
            let mut inserted = seed[..position].to_vec();
            inserted.extend_from_slice(SNIPPETS[next() % SNIPPETS.len()]);
            inserted.extend_from_slice(&seed[position..]);
            output.push(inserted);
        }
        output
    }

    #[test]
    fn single_pass_semantic_text_matches_the_three_pass_oracle_on_handcrafted_xml() {
        let mut compared = 0;
        let mut accepted = 0;
        for (index, seed) in text_seeds().iter().enumerate() {
            let (count, ok) = assert_text_routes_agree(&format!("seed {index}"), seed);
            compared += count;
            accepted += ok;
            for (variant, mutated) in text_mutations(seed, 60).iter().enumerate() {
                let (count, ok) =
                    assert_text_routes_agree(&format!("seed {index} variant {variant}"), mutated);
                compared += count;
                accepted += ok;
            }
        }
        println!("0743-text-oracle handcrafted compared={compared} accepted={accepted}");
        assert!(accepted > 0 && accepted < compared);
    }

    #[test]
    fn single_pass_semantic_text_matches_the_three_pass_oracle_on_the_pptx_corpus() {
        fn collect(directory: &std::path::Path, found: &mut Vec<std::path::PathBuf>) {
            let Ok(entries) = std::fs::read_dir(directory) else {
                return;
            };
            for entry in entries.flatten() {
                let path = entry.path();
                if path.is_dir() {
                    collect(&path, found);
                } else if path
                    .extension()
                    .is_some_and(|extension| extension == "pptx")
                {
                    found.push(path);
                }
            }
        }
        let root = std::path::Path::new(env!("CARGO_MANIFEST_DIR")).join("../../test-data");
        let mut fixtures = Vec::new();
        collect(&root, &mut fixtures);
        fixtures.sort();
        let mut parts = 0;
        let mut compared = 0;
        let mut accepted = 0;
        for fixture in fixtures {
            let Ok(bytes) = std::fs::read(&fixture) else {
                continue;
            };
            let Ok(package) = litchi_opc::OpcPackage::from_vec(bytes) else {
                continue;
            };
            // Package parts iterate in hash order; visit them by name so the
            // oracle compares, and mutates, the same parts on every run.
            let mut named: Vec<_> = package.try_iter_parts().flatten().collect();
            named.sort_by(|left, right| left.partname().as_str().cmp(right.partname().as_str()));
            for part in named {
                if !part.content_type().ends_with("+xml") {
                    continue;
                }
                let label = format!("{} {}", fixture.display(), part.partname());
                let (count, ok) = assert_text_routes_agree(&label, part.blob());
                compared += count;
                accepted += ok;
                if part.content_type() == ct::PML_SLIDE && parts % 4 == 0 {
                    for (variant, mutated) in text_mutations(part.blob(), 16).iter().enumerate() {
                        let (count, ok) = assert_text_routes_agree(
                            &format!("{label} variant {variant}"),
                            mutated,
                        );
                        compared += count;
                        accepted += ok;
                    }
                }
                parts += 1;
            }
        }
        println!("0743-text-oracle corpus parts={parts} compared={compared} accepted={accepted}");
        assert!(parts >= 300 && accepted > 0 && accepted < compared);
    }

    /// The precedence the single pass must keep, stated independently of the
    /// oracle: any raw-scan refusal outranks any semantic refusal, wherever
    /// either occurs, and a semantic refusal alone is the first one in order.
    #[test]
    fn single_pass_keeps_raw_refusals_ahead_of_earlier_semantic_refusals() {
        let run =
            "<p:sp><p:txBody><a:p><a:r><a:t>x<a:t>nested</a:t></a:t></a:r></a:p></p:txBody></p:sp>";
        let foreign = "<q:t>foreign</q:t>";
        let bad_comment = "<!-- bad -- comment -->";
        let cases: [(String, &str); 5] = [
            (format!("{run}{bad_comment}"), "invalid comment"),
            (format!("{bad_comment}{run}"), "invalid comment"),
            (
                run.to_owned(),
                "nested DrawingML text elements are not permitted",
            ),
            (
                format!("{run}<p:t>second semantic refusal</p:t>"),
                "nested DrawingML text elements are not permitted",
            ),
            (
                format!("{run}{foreign}"),
                "unresolved semantic slide element namespace prefix 'q'",
            ),
        ];
        for (body, expected) in cases {
            let xml = slide(&body);
            let part = slide_part(&xml);
            let error = semantic_text_from_part(&part, "\n")
                .expect_err("each case must refuse")
                .to_string();
            assert!(
                error.contains(expected),
                "{body}: expected {expected:?}, got {error:?}"
            );
        }
    }
}
