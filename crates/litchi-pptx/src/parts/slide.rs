//! Borrowed slide, layout, and master part views.

use litchi_ooxml_common::mce::{Capabilities, Limits as MceLimits, process_markup_compatibility};
use litchi_ooxml_common::xml::{DRAWINGML_NAMESPACE, STRICT_DRAWINGML_NAMESPACE};
use litchi_opc::constants::{content_type as ct, relationship_type as rt};
use litchi_opc::{OpcPackage, Part};
use quick_xml::events::Event;
use quick_xml::name::{Namespace, QName, ResolveResult};
use quick_xml::reader::{NsReader, Reader};

use super::{invalid, processed_xml, related_part_by_type, validate_content_type};
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

fn root_name(part: &dyn Part) -> Result<String> {
    let xml = processed_xml(part)?;
    let mut reader = NsReader::from_reader(xml.as_ref());
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

fn c_sld_name(part: &dyn Part) -> Result<Option<String>> {
    let xml = processed_xml(part)?;
    crate::namespace::presentation_name(xml.as_ref())
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

fn scan_raw_semantic_text_xml(xml: &[u8]) -> Result<()> {
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
                if budget.depth == 0 && !decoded.as_bytes().iter().all(u8::is_ascii_whitespace) {
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
        }
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
        validate_semantic_attributes(element)?;
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
        validate_semantic_attributes(element)?;
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
    scan_raw_semantic_text_xml(raw)?;
    let processed =
        process_markup_compatibility(raw, &Capabilities::ooxml_baseline(), &semantic_mce_limits())?;
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
}
