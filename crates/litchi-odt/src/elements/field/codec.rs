//! Namespace-aware ODF field XML codecs.

#[allow(
    clippy::wildcard_imports,
    reason = "the field codec consumes the complete typed field model"
)]
use super::model::*;
#[allow(
    clippy::wildcard_imports,
    reason = "the field codec shares the owner-level namespace and model vocabulary"
)]
use super::*;
use crate::elements::xml::decode_reference;
use litchi_core::{Error, Result};
use quick_xml::XmlVersion;
use quick_xml::events::Event;
use quick_xml::reader::NsReader;
use std::collections::{HashMap, HashSet};

pub(super) fn checked_field_depth(depth: usize) -> Result<usize> {
    let depth = depth
        .checked_add(1)
        .ok_or_else(|| Error::InvalidFormat("field nesting depth overflow".to_string()))?;
    if depth > MAX_FIELD_DEPTH {
        return Err(Error::InvalidFormat(format!(
            "field nesting exceeds {MAX_FIELD_DEPTH} levels"
        )));
    }
    Ok(depth)
}

impl DatabaseField {
    pub fn to_xml_fragment(&self) -> Result<String> {
        let field = validate_database_field(self.clone())?;
        validate_constructed_database_field(&field)?;
        let local = match field.kind {
            DatabaseFieldKind::Display => "database-display",
            DatabaseFieldKind::Next => "database-next",
            DatabaseFieldKind::RowSelect => "database-row-select",
            DatabaseFieldKind::RowNumber => "database-row-number",
            DatabaseFieldKind::Name => "database-name",
        };
        let mut xml = format!(
            "<text:{local} xmlns:text=\"{TEXT_DATABASE_NAMESPACE}\" xmlns:style=\"{STYLE_NAMESPACE}\" xmlns:form=\"{FORM_NAMESPACE}\" xmlns:xlink=\"{XLINK_NAMESPACE}\""
        );
        let mut attribute = |prefix: &str, name: &str, value: &str| {
            xml.push(' ');
            xml.push_str(prefix);
            xml.push(':');
            xml.push_str(name);
            xml.push_str("=\"");
            push_xml_attribute(&mut xml, value);
            xml.push('"');
        };
        if let Some(value) = field.source.database_name.as_deref() {
            attribute("text", "database-name", value);
        }
        attribute("text", "table-name", &field.source.table_name);
        if let Some(value) = field.source.table_type {
            attribute("text", "table-type", value.as_str());
        }
        match field.kind {
            DatabaseFieldKind::Display => {
                attribute(
                    "text",
                    "column-name",
                    field.column_name.as_deref().expect("validated"),
                );
                if let Some(value) = field.data_style_name.as_deref() {
                    attribute("style", "data-style-name", value);
                }
            },
            DatabaseFieldKind::Next => {
                if let Some(value) = field.condition.as_deref() {
                    attribute("text", "condition", value);
                }
            },
            DatabaseFieldKind::RowSelect => {
                if let Some(value) = field.condition.as_deref() {
                    attribute("text", "condition", value);
                }
                if let Some(value) = field.row_number {
                    attribute("text", "row-number", value.as_str());
                }
            },
            DatabaseFieldKind::RowNumber => {
                if let Some(value) = field.value {
                    attribute("text", "value", value.as_str());
                }
                if let Some(value) = field.number_format.as_deref() {
                    attribute("style", "num-format", value);
                }
                if let Some(value) = field.number_letter_sync {
                    attribute(
                        "style",
                        "num-letter-sync",
                        if value { "true" } else { "false" },
                    );
                }
            },
            DatabaseFieldKind::Name => {},
        }
        let _ = attribute;
        if field.source.connection_resource.is_none() && field.display_text.is_empty() {
            xml.push_str("/>");
            return Ok(xml);
        }
        xml.push('>');
        if let Some(resource) = &field.source.connection_resource {
            xml.push_str("<form:connection-resource xlink:href=\"");
            push_xml_attribute(&mut xml, &resource.href);
            xml.push_str("\"/>");
        }
        push_xml_text(&mut xml, &field.display_text);
        xml.push_str("</text:");
        xml.push_str(local);
        xml.push('>');
        Ok(xml)
    }
}

struct ActiveDatabaseField {
    depth: usize,
    field: DatabaseField,
    connection_depth: Option<usize>,
}

struct ActiveDropDownField {
    depth: usize,
    label_depth: Option<usize>,
    display_started: bool,
    aggregate: usize,
    name: String,
    labels: Vec<DropDownLabel>,
    display_text: String,
}

type DatabaseAttributes = HashMap<(String, String), String>;

#[derive(Debug)]
struct ActiveMetaField {
    depth: usize,
    order: usize,
    xml_id: String,
    data_style_name: Option<String>,
    builder: MetaContentBuilder,
}

#[derive(Debug)]
struct ActiveNoteBody {
    depth: usize,
    order: usize,
    builder: MetaContentBuilder,
}

#[derive(Debug)]
struct MetaContentBuilder {
    roots: Vec<MetaFieldNode>,
    stack: Vec<MetaFieldElement>,
    nodes: usize,
    aggregate: usize,
    root_grammar: MetaContentGrammar,
    root_name: &'static str,
}

impl Default for MetaContentBuilder {
    fn default() -> Self {
        Self::new(MetaContentGrammar::ParagraphOrHyperlink, "text:meta-field")
    }
}

impl MetaContentBuilder {
    fn new(root_grammar: MetaContentGrammar, root_name: &'static str) -> Self {
        Self {
            roots: Vec::new(),
            stack: Vec::new(),
            nodes: 0,
            aggregate: 0,
            root_grammar,
            root_name,
        }
    }

    fn note_body() -> Self {
        Self::new(MetaContentGrammar::NoteBody, "text:note-body")
    }

    fn push_text(&mut self, value: &str) -> Result<()> {
        if self.stack.is_empty() && self.root_grammar == MetaContentGrammar::NoteBody {
            if value.chars().all(char::is_whitespace) {
                return Ok(());
            }
            return Err(Error::InvalidFormat(
                "text:note-body cannot contain direct character data".to_string(),
            ));
        }
        add_meta_size(&mut self.aggregate, value.len())?;
        if let Some(MetaFieldNode::Text(text)) = self.current_nodes_mut().last_mut() {
            text.push_str(value);
        } else {
            self.add_node()?;
            self.current_nodes_mut()
                .push(MetaFieldNode::Text(value.to_string()));
        }
        Ok(())
    }

    fn start_element(
        &mut self,
        namespace_uri: String,
        local_name: String,
        attributes: Vec<MetaFieldAttribute>,
    ) -> Result<()> {
        if self.stack.is_empty()
            && meta_child_grammar(self.root_grammar, &namespace_uri, &local_name).is_err()
        {
            return Err(Error::InvalidFormat(format!(
                "{}:{local_name} is not permitted directly in {}",
                namespace_uri, self.root_name
            )));
        }
        validate_meta_element_parts(
            &namespace_uri,
            &local_name,
            &attributes,
            &mut self.aggregate,
        )?;
        if self.stack.len() >= MAX_META_FIELD_DEPTH {
            return Err(Error::InvalidFormat(format!(
                "text:meta-field content exceeds {MAX_META_FIELD_DEPTH} levels"
            )));
        }
        self.add_node()?;
        self.stack.push(MetaFieldElement {
            namespace_uri,
            local_name,
            attributes,
            children: Vec::new(),
        });
        Ok(())
    }

    fn empty_element(
        &mut self,
        namespace_uri: String,
        local_name: String,
        attributes: Vec<MetaFieldAttribute>,
    ) -> Result<()> {
        self.start_element(namespace_uri, local_name, attributes)?;
        self.end_element()
    }

    fn end_element(&mut self) -> Result<()> {
        let element = self.stack.pop().ok_or_else(|| {
            Error::InvalidFormat("text:meta-field content stack underflow".to_string())
        })?;
        self.current_nodes_mut()
            .push(MetaFieldNode::Element(element));
        Ok(())
    }

    fn finish_meta_field(self) -> Result<MetaFieldContent> {
        if !self.stack.is_empty() {
            return Err(Error::InvalidFormat(
                "incomplete text:meta-field content".to_string(),
            ));
        }
        MetaFieldContent::new(self.roots)
    }

    fn finish_note_body(self) -> Result<NoteBodyContent> {
        if !self.stack.is_empty() {
            return Err(Error::InvalidFormat(
                "incomplete text:note-body content".to_string(),
            ));
        }
        NoteBodyContent::new(self.roots)
    }

    fn current_nodes_mut(&mut self) -> &mut Vec<MetaFieldNode> {
        if let Some(element) = self.stack.last_mut() {
            &mut element.children
        } else {
            &mut self.roots
        }
    }

    fn add_node(&mut self) -> Result<()> {
        self.nodes = self.nodes.checked_add(1).ok_or_else(|| {
            Error::InvalidFormat("text:meta-field node count overflow".to_string())
        })?;
        if self.nodes > MAX_META_FIELD_NODES {
            return Err(Error::InvalidFormat(format!(
                "text:meta-field exceeds {MAX_META_FIELD_NODES} content nodes"
            )));
        }
        Ok(())
    }
}

pub(super) fn parse_meta_fields(xml: &str) -> Result<Vec<DynamicTextField>> {
    if xml.len() > MAX_META_FIELD_XML_BYTES {
        return Err(Error::InvalidFormat(format!(
            "field XML exceeds {MAX_META_FIELD_XML_BYTES} bytes"
        )));
    }
    let mut reader = NsReader::from_str(xml);
    reader.config_mut().check_end_names = true;
    reader.config_mut().trim_text(false);
    let mut buffer = Vec::new();
    let mut depth = 0usize;
    let mut stack: Vec<(Option<String>, String)> = Vec::new();
    let mut active: Vec<ActiveMetaField> = Vec::new();
    let mut completed = Vec::new();
    let mut next_order = 0usize;
    let mut document_xml_ids = HashSet::new();

    loop {
        let (namespace, event) = reader
            .read_resolved_event_into(&mut buffer)
            .map_err(|error| Error::InvalidFormat(format!("invalid meta-field XML: {error}")))?;
        match event {
            Event::Start(ref source) => {
                let namespace_uri = resolved_namespace(&namespace)?;
                collect_document_xml_id(&reader, source, &mut document_xml_ids)?;
                let local = utf8(source.local_name().as_ref(), "meta-field element name")?;
                if !active.is_empty() {
                    let attributes = parse_meta_node_attributes(&reader, source)?;
                    for field in &mut active {
                        field.builder.start_element(
                            namespace_uri.clone().ok_or_else(|| {
                                Error::InvalidFormat(
                                    "unqualified meta-field child element".to_string(),
                                )
                            })?,
                            local.clone(),
                            attributes.clone(),
                        )?;
                    }
                }
                depth = depth.checked_add(1).ok_or_else(|| {
                    Error::InvalidFormat("meta-field XML depth overflow".to_string())
                })?;
                if namespace_uri.as_deref() == Some(TEXT_DATABASE_NAMESPACE)
                    && local == "meta-field"
                {
                    validate_meta_field_parent(stack.last())?;
                    let (xml_id, data_style_name) = parse_meta_root_attributes(&reader, source)?;
                    if next_order >= MAX_FIELDS {
                        return Err(Error::InvalidFormat(
                            "too many text:meta-field elements".to_string(),
                        ));
                    }
                    active.push(ActiveMetaField {
                        depth,
                        order: next_order,
                        xml_id,
                        data_style_name,
                        builder: MetaContentBuilder::default(),
                    });
                    next_order += 1;
                }
                stack.push((namespace_uri, local));
            },
            Event::Empty(ref source) => {
                let namespace_uri = resolved_namespace(&namespace)?;
                collect_document_xml_id(&reader, source, &mut document_xml_ids)?;
                let local = utf8(source.local_name().as_ref(), "meta-field element name")?;
                if !active.is_empty() {
                    let attributes = parse_meta_node_attributes(&reader, source)?;
                    for field in &mut active {
                        field.builder.empty_element(
                            namespace_uri.clone().ok_or_else(|| {
                                Error::InvalidFormat(
                                    "unqualified meta-field child element".to_string(),
                                )
                            })?,
                            local.clone(),
                            attributes.clone(),
                        )?;
                    }
                }
                if namespace_uri.as_deref() == Some(TEXT_DATABASE_NAMESPACE)
                    && local == "meta-field"
                {
                    validate_meta_field_parent(stack.last())?;
                    let (xml_id, data_style_name) = parse_meta_root_attributes(&reader, source)?;
                    completed.push((
                        next_order,
                        DynamicTextField::MetaField {
                            xml_id,
                            data_style_name,
                            content: MetaFieldContent::new(Vec::new())?,
                        },
                    ));
                    next_order += 1;
                }
            },
            Event::Text(ref value) => {
                let value = value
                    .xml_content(XmlVersion::Explicit1_0)
                    .map_err(|error| {
                        Error::InvalidFormat(format!("invalid meta-field text: {error}"))
                    })?;
                for field in &mut active {
                    field.builder.push_text(&value)?;
                }
            },
            Event::CData(ref value) => {
                let value = value
                    .xml_content(XmlVersion::Explicit1_0)
                    .map_err(|error| {
                        Error::InvalidFormat(format!("invalid meta-field CDATA: {error}"))
                    })?;
                for field in &mut active {
                    field.builder.push_text(&value)?;
                }
            },
            Event::GeneralRef(ref reference) => {
                let value = decode_reference(reference, "meta-field")?;
                for field in &mut active {
                    field.builder.push_text(&value)?;
                }
            },
            Event::End(_) => {
                for field in &mut active {
                    if field.depth < depth {
                        field.builder.end_element()?;
                    }
                }
                if let Some(field) = active.pop_if(|field| field.depth == depth) {
                    completed.push((
                        field.order,
                        DynamicTextField::MetaField {
                            xml_id: field.xml_id,
                            data_style_name: field.data_style_name,
                            content: field.builder.finish_meta_field()?,
                        },
                    ));
                }
                stack.pop().ok_or_else(|| {
                    Error::InvalidFormat("meta-field XML stack underflow".to_string())
                })?;
                depth = depth.checked_sub(1).ok_or_else(|| {
                    Error::InvalidFormat("meta-field XML depth underflow".to_string())
                })?;
            },
            Event::DocType(_) => {
                return Err(Error::InvalidFormat(
                    "DOCTYPE is not permitted in ODF field XML".to_string(),
                ));
            },
            Event::PI(_) if !active.is_empty() => {
                return Err(Error::InvalidFormat(
                    "processing instructions are not permitted in text:meta-field".to_string(),
                ));
            },
            Event::Eof => break,
            _ => {},
        }
        buffer.clear();
    }
    if depth != 0 || !stack.is_empty() || !active.is_empty() {
        return Err(Error::InvalidFormat(
            "incomplete meta-field XML".to_string(),
        ));
    }
    completed.sort_by_key(|(order, _)| *order);
    Ok(completed.into_iter().map(|(_, field)| field).collect())
}

/// Parse every direct `text:note-body` child of an ODF `text:note` into the
/// shared inert mixed-content model. This does not evaluate fields, links,
/// event listeners, scripts, or macros represented by the nodes.
pub(super) fn parse_note_body_contents(xml: &str) -> Result<Vec<NoteBodyContent>> {
    let mut reader = NsReader::from_str(xml);
    reader.config_mut().check_end_names = true;
    reader.config_mut().trim_text(false);
    let mut buffer = Vec::new();
    let mut depth = 0usize;
    let mut stack: Vec<(Option<String>, String)> = Vec::new();
    let mut active: Vec<ActiveNoteBody> = Vec::new();
    let mut completed = Vec::new();
    let mut next_order = 0usize;

    loop {
        let (namespace, event) = reader
            .read_resolved_event_into(&mut buffer)
            .map_err(|error| Error::InvalidFormat(format!("invalid note-body XML: {error}")))?;
        match event {
            Event::Start(ref source) => {
                let namespace_uri = resolved_namespace(&namespace)?;
                let local = utf8(source.local_name().as_ref(), "note-body element name")?;
                let is_note_body = namespace_uri.as_deref() == Some(TEXT_DATABASE_NAMESPACE)
                    && local == "note-body"
                    && stack
                        .last()
                        .is_some_and(|(parent_namespace, parent_local)| {
                            parent_namespace.as_deref() == Some(TEXT_DATABASE_NAMESPACE)
                                && parent_local == "note"
                        });
                if is_note_body {
                    validate_note_body_attributes(source)?;
                }
                if !active.is_empty() {
                    let attributes = parse_meta_node_attributes(&reader, source)?;
                    let namespace_uri = namespace_uri.clone().ok_or_else(|| {
                        Error::InvalidFormat("unqualified note-body child element".to_string())
                    })?;
                    for body in &mut active {
                        body.builder.start_element(
                            namespace_uri.clone(),
                            local.clone(),
                            attributes.clone(),
                        )?;
                    }
                }
                depth = depth.checked_add(1).ok_or_else(|| {
                    Error::InvalidFormat("note-body XML depth overflow".to_string())
                })?;
                if depth > MAX_FIELD_DEPTH {
                    return Err(Error::InvalidFormat(format!(
                        "note-body XML exceeds {MAX_FIELD_DEPTH} levels"
                    )));
                }
                if is_note_body {
                    if next_order >= MAX_FIELDS {
                        return Err(Error::InvalidFormat(
                            "document exceeds note-body limit".to_string(),
                        ));
                    }
                    active.push(ActiveNoteBody {
                        depth,
                        order: next_order,
                        builder: MetaContentBuilder::note_body(),
                    });
                    next_order += 1;
                }
                stack.push((namespace_uri, local));
            },
            Event::Empty(ref source) => {
                let namespace_uri = resolved_namespace(&namespace)?;
                let local = utf8(source.local_name().as_ref(), "note-body element name")?;
                let is_note_body = namespace_uri.as_deref() == Some(TEXT_DATABASE_NAMESPACE)
                    && local == "note-body"
                    && stack
                        .last()
                        .is_some_and(|(parent_namespace, parent_local)| {
                            parent_namespace.as_deref() == Some(TEXT_DATABASE_NAMESPACE)
                                && parent_local == "note"
                        });
                if is_note_body {
                    validate_note_body_attributes(source)?;
                }
                if !active.is_empty() {
                    let attributes = parse_meta_node_attributes(&reader, source)?;
                    let namespace_uri = namespace_uri.clone().ok_or_else(|| {
                        Error::InvalidFormat("unqualified note-body child element".to_string())
                    })?;
                    for body in &mut active {
                        body.builder.empty_element(
                            namespace_uri.clone(),
                            local.clone(),
                            attributes.clone(),
                        )?;
                    }
                }
                if is_note_body {
                    if next_order >= MAX_FIELDS {
                        return Err(Error::InvalidFormat(
                            "document exceeds note-body limit".to_string(),
                        ));
                    }
                    completed.push((next_order, NoteBodyContent::new(Vec::new())?));
                    next_order += 1;
                }
            },
            Event::Text(ref value) if !active.is_empty() => {
                let value = value
                    .xml_content(XmlVersion::Explicit1_0)
                    .map_err(|error| {
                        Error::InvalidFormat(format!("invalid note-body text: {error}"))
                    })?;
                for body in &mut active {
                    body.builder.push_text(&value)?;
                }
            },
            Event::CData(ref value) if !active.is_empty() => {
                let value = value
                    .xml_content(XmlVersion::Explicit1_0)
                    .map_err(|error| {
                        Error::InvalidFormat(format!("invalid note-body CDATA: {error}"))
                    })?;
                for body in &mut active {
                    body.builder.push_text(&value)?;
                }
            },
            Event::GeneralRef(ref reference) if !active.is_empty() => {
                let value = decode_reference(reference, "note-body")?;
                for body in &mut active {
                    body.builder.push_text(&value)?;
                }
            },
            Event::End(_) => {
                for body in &mut active {
                    if body.depth < depth {
                        body.builder.end_element()?;
                    }
                }
                if let Some(body) = active.pop_if(|body| body.depth == depth) {
                    completed.push((body.order, body.builder.finish_note_body()?));
                }
                stack.pop().ok_or_else(|| {
                    Error::InvalidFormat("note-body XML stack underflow".to_string())
                })?;
                depth = depth.checked_sub(1).ok_or_else(|| {
                    Error::InvalidFormat("note-body XML depth underflow".to_string())
                })?;
            },
            Event::DocType(_) => {
                return Err(Error::InvalidFormat(
                    "DOCTYPE is not permitted in ODF note-body XML".to_string(),
                ));
            },
            Event::PI(_) if !active.is_empty() => {
                return Err(Error::InvalidFormat(
                    "processing instructions are not permitted in text:note-body".to_string(),
                ));
            },
            Event::Eof => break,
            _ => {},
        }
        buffer.clear();
    }
    if depth != 0 || !stack.is_empty() || !active.is_empty() {
        return Err(Error::InvalidFormat("incomplete note-body XML".to_string()));
    }
    completed.sort_by_key(|(order, _)| *order);
    Ok(completed.into_iter().map(|(_, body)| body).collect())
}

fn validate_note_body_attributes(source: &quick_xml::events::BytesStart<'_>) -> Result<()> {
    for attribute in source.attributes().with_checks(true) {
        let attribute = attribute.map_err(|error| {
            Error::InvalidFormat(format!("invalid text:note-body attribute: {error}"))
        })?;
        let raw = attribute.key.as_ref();
        if raw == b"xmlns" || raw.starts_with(b"xmlns:") {
            continue;
        }
        return Err(Error::InvalidFormat(
            "text:note-body does not permit attributes".to_string(),
        ));
    }
    Ok(())
}

fn collect_document_xml_id(
    reader: &NsReader<&[u8]>,
    source: &quick_xml::events::BytesStart<'_>,
    ids: &mut HashSet<String>,
) -> Result<()> {
    for attribute in source.attributes().with_checks(true) {
        let attribute = attribute.map_err(|error| {
            Error::InvalidFormat(format!(
                "invalid XML attribute while collecting xml:id: {error}"
            ))
        })?;
        let (namespace, local) = reader.resolver().resolve_attribute(attribute.key);
        if resolved_namespace(&namespace)?.as_deref() != Some(XML_NAMESPACE)
            || local.as_ref() != b"id"
        {
            continue;
        }
        let value = attribute
            .decoded_and_normalized_value(XmlVersion::Explicit1_0, reader.decoder())
            .map_err(|error| Error::InvalidFormat(format!("invalid xml:id value: {error}")))?
            .into_owned();
        validate_xml_id(&value)?;
        if !ids.insert(value.clone()) {
            return Err(Error::InvalidFormat(format!(
                "duplicate document xml:id '{value}'"
            )));
        }
    }
    Ok(())
}

fn parse_meta_root_attributes(
    reader: &NsReader<&[u8]>,
    source: &quick_xml::events::BytesStart<'_>,
) -> Result<(String, Option<String>)> {
    let mut xml_id = None;
    let mut data_style_name = None;
    for attribute in source.attributes() {
        let attribute = attribute.map_err(|error| {
            Error::InvalidFormat(format!("invalid meta-field attribute: {error}"))
        })?;
        let raw = attribute.key.as_ref();
        if raw == b"xmlns" || raw.starts_with(b"xmlns:") {
            continue;
        }
        let (namespace, local) = reader.resolver().resolve_attribute(attribute.key);
        let namespace = resolved_namespace(&namespace)?;
        let local = utf8(local.as_ref(), "meta-field attribute name")?;
        let value = attribute
            .decoded_and_normalized_value(XmlVersion::Explicit1_0, reader.decoder())
            .map_err(|error| {
                Error::InvalidFormat(format!("invalid meta-field attribute: {error}"))
            })?
            .into_owned();
        match (namespace.as_deref(), local.as_str()) {
            (Some(XML_NAMESPACE), "id") => xml_id = Some(value),
            (Some(STYLE_NAMESPACE), "data-style-name") => data_style_name = Some(value),
            _ => {
                return Err(Error::InvalidFormat(format!(
                    "unexpected text:meta-field attribute {}:{local}",
                    namespace.as_deref().unwrap_or("unqualified")
                )));
            },
        }
    }
    let xml_id = xml_id
        .ok_or_else(|| Error::InvalidFormat("text:meta-field requires xml:id".to_string()))?;
    validate_xml_id(&xml_id)?;
    Ok((xml_id, data_style_name))
}

fn parse_meta_node_attributes(
    reader: &NsReader<&[u8]>,
    source: &quick_xml::events::BytesStart<'_>,
) -> Result<Vec<MetaFieldAttribute>> {
    let mut attributes = Vec::new();
    for attribute in source.attributes() {
        let attribute = attribute.map_err(|error| {
            Error::InvalidFormat(format!("invalid meta-field child attribute: {error}"))
        })?;
        let raw = attribute.key.as_ref();
        if raw == b"xmlns" || raw.starts_with(b"xmlns:") {
            continue;
        }
        if attributes.len() >= MAX_META_FIELD_ATTRIBUTES {
            return Err(Error::InvalidFormat(format!(
                "meta-field child exceeds {MAX_META_FIELD_ATTRIBUTES} attributes"
            )));
        }
        let (namespace, local) = reader.resolver().resolve_attribute(attribute.key);
        let namespace_uri = resolved_namespace(&namespace)?.ok_or_else(|| {
            Error::InvalidFormat("unqualified meta-field child attribute".to_string())
        })?;
        if !is_allowed_meta_namespace(&namespace_uri) {
            return Err(Error::InvalidFormat(format!(
                "foreign meta-field attribute namespace '{namespace_uri}'"
            )));
        }
        attributes.push(MetaFieldAttribute {
            namespace_uri,
            local_name: utf8(local.as_ref(), "meta-field attribute name")?,
            value: attribute
                .decoded_and_normalized_value(XmlVersion::Explicit1_0, reader.decoder())
                .map_err(|error| {
                    Error::InvalidFormat(format!("invalid meta-field attribute: {error}"))
                })?
                .into_owned(),
        });
    }
    Ok(attributes)
}

fn validate_meta_field_parent(parent: Option<&(Option<String>, String)>) -> Result<()> {
    let valid = parent.is_some_and(|(namespace, local)| {
        namespace.as_deref() == Some(TEXT_DATABASE_NAMESPACE)
            && matches!(
                local.as_str(),
                "a" | "h" | "meta" | "meta-field" | "p" | "ruby-base" | "span"
            )
    });
    if valid {
        Ok(())
    } else {
        Err(Error::InvalidFormat(
            "text:meta-field occurs outside an ODF inline-text host".to_string(),
        ))
    }
}

pub(super) fn parse_database_fields(xml: &str) -> Result<Vec<DatabaseField>> {
    if xml.len() > MAX_META_FIELD_XML_BYTES {
        return Err(Error::InvalidFormat(
            "database field XML exceeds 64 MiB".to_string(),
        ));
    }
    let mut reader = NsReader::from_str(xml);
    reader.config_mut().check_end_names = true;
    let mut buffer = Vec::new();
    let mut depth = 0usize;
    let mut active: Option<ActiveDatabaseField> = None;
    let mut fields = Vec::new();
    let mut aggregate = 0usize;
    let mut stack: Vec<(Option<String>, String)> = Vec::new();

    loop {
        let (namespace, event) = reader
            .read_resolved_event_into(&mut buffer)
            .map_err(|error| {
                Error::InvalidFormat(format!("invalid database field XML: {error}"))
            })?;
        let namespace_uri = resolved_namespace(&namespace)?;
        match event {
            Event::Start(ref element) => {
                let local = utf8(element.local_name().as_ref(), "database field element")?;
                if let Some(field) = active.as_mut() {
                    if namespace_uri.as_deref() != Some(FORM_NAMESPACE)
                        || local != "connection-resource"
                        || depth != field.depth
                        || field.connection_depth.is_some()
                        || field.field.source.connection_resource.is_some()
                        || !field.field.display_text.is_empty()
                    {
                        return Err(Error::InvalidFormat(
                            "database fields may contain only one form:connection-resource"
                                .to_string(),
                        ));
                    }
                    field.field.source.connection_resource =
                        Some(parse_connection_resource(&reader, element, &mut aggregate)?);
                    field.connection_depth = Some(depth + 1);
                } else if namespace_uri.as_deref() == Some(TEXT_DATABASE_NAMESPACE)
                    && let Some(kind) = database_field_kind(&local)
                {
                    validate_database_parent(stack.last())?;
                    if fields.len() >= MAX_FIELDS {
                        return Err(Error::InvalidFormat(format!(
                            "document exceeds {MAX_FIELDS} database fields"
                        )));
                    }
                    active = Some(ActiveDatabaseField {
                        depth: depth + 1,
                        field: parse_database_field(&reader, element, kind, &mut aggregate)?,
                        connection_depth: None,
                    });
                }
                depth = checked_field_depth(depth)?;
                stack.push((namespace_uri, local));
            },
            Event::Empty(ref element) => {
                let local = utf8(element.local_name().as_ref(), "database field element")?;
                if let Some(field) = active.as_mut() {
                    if namespace_uri.as_deref() != Some(FORM_NAMESPACE)
                        || local != "connection-resource"
                        || depth != field.depth
                        || field.field.source.connection_resource.is_some()
                        || !field.field.display_text.is_empty()
                    {
                        return Err(Error::InvalidFormat(
                            "database fields may contain only one form:connection-resource"
                                .to_string(),
                        ));
                    }
                    field.field.source.connection_resource =
                        Some(parse_connection_resource(&reader, element, &mut aggregate)?);
                } else if namespace_uri.as_deref() == Some(TEXT_DATABASE_NAMESPACE)
                    && let Some(kind) = database_field_kind(&local)
                {
                    validate_database_parent(stack.last())?;
                    if fields.len() >= MAX_FIELDS {
                        return Err(Error::InvalidFormat(format!(
                            "document exceeds {MAX_FIELDS} database fields"
                        )));
                    }
                    let field = parse_database_field(&reader, element, kind, &mut aggregate)?;
                    fields.push(validate_database_field(field)?);
                }
            },
            Event::End(_) => {
                if let Some(field) = active.as_mut()
                    && field
                        .connection_depth
                        .is_some_and(|connection_depth| connection_depth == depth)
                {
                    field.connection_depth = None;
                }
                if active.as_ref().is_some_and(|field| field.depth == depth) {
                    let field = active.take().expect("checked database field").field;
                    fields.push(validate_database_field(field)?);
                }
                stack.pop().ok_or_else(|| {
                    Error::InvalidFormat("database field XML stack underflow".to_string())
                })?;
                depth = depth.checked_sub(1).ok_or_else(|| {
                    Error::InvalidFormat("database field XML depth underflow".to_string())
                })?;
            },
            Event::Text(ref text) if active.is_some() => {
                let value = text.decode().map_err(|error| {
                    Error::InvalidFormat(format!("invalid database field text: {error}"))
                })?;
                append_database_text(
                    active.as_mut().expect("checked field"),
                    &value,
                    &mut aggregate,
                )?;
            },
            Event::CData(ref text) if active.is_some() => {
                let value = text.decode().map_err(|error| {
                    Error::InvalidFormat(format!("invalid database field CDATA: {error}"))
                })?;
                append_database_text(
                    active.as_mut().expect("checked field"),
                    &value,
                    &mut aggregate,
                )?;
            },
            Event::GeneralRef(ref reference) if active.is_some() => {
                let name = std::str::from_utf8(reference.as_ref()).map_err(|_| {
                    Error::InvalidFormat("invalid database field entity reference".to_string())
                })?;
                let value = resolve_database_reference(name)?;
                append_database_text(
                    active.as_mut().expect("checked field"),
                    &value,
                    &mut aggregate,
                )?;
            },
            Event::DocType(_) | Event::PI(_) => {
                return Err(Error::InvalidFormat(
                    "DTDs and processing instructions are prohibited in database field XML"
                        .to_string(),
                ));
            },
            Event::Eof => break,
            _ => {},
        }
        buffer.clear();
    }
    if depth != 0 || active.is_some() || !stack.is_empty() {
        return Err(Error::InvalidFormat(
            "incomplete database field XML structure".to_string(),
        ));
    }
    Ok(fields)
}

pub(super) fn parse_drop_down_fields(xml: &str) -> Result<Vec<DynamicTextField>> {
    if xml.len() > MAX_META_FIELD_XML_BYTES {
        return Err(Error::InvalidFormat(
            "drop-down field XML exceeds 64 MiB".to_string(),
        ));
    }
    let mut reader = NsReader::from_str(xml);
    reader.config_mut().check_end_names = true;
    reader.config_mut().trim_text(false);
    let mut buffer = Vec::new();
    let mut depth = 0usize;
    let mut active: Option<ActiveDropDownField> = None;
    let mut fields = Vec::new();
    let mut stack: Vec<(Option<String>, String)> = Vec::new();

    loop {
        let (namespace, event) = reader
            .read_resolved_event_into(&mut buffer)
            .map_err(|error| {
                Error::InvalidFormat(format!("invalid drop-down field XML: {error}"))
            })?;
        let namespace_uri = resolved_namespace(&namespace)?;
        match event {
            Event::Start(ref element) => {
                let local = utf8(element.local_name().as_ref(), "drop-down field element")?;
                if let Some(field) = active.as_mut() {
                    begin_drop_down_label(
                        &reader,
                        element,
                        namespace_uri.as_deref(),
                        &local,
                        depth,
                        field,
                    )?;
                } else if namespace_uri.as_deref() == Some(TEXT_DATABASE_NAMESPACE)
                    && local == "drop-down"
                {
                    validate_drop_down_parent(stack.last())?;
                    if fields.len() >= MAX_FIELDS {
                        return Err(Error::InvalidFormat(format!(
                            "document exceeds {MAX_FIELDS} drop-down fields"
                        )));
                    }
                    let mut field = parse_drop_down_field(&reader, element)?;
                    field.depth = depth.checked_add(1).ok_or_else(|| {
                        Error::InvalidFormat("drop-down field depth overflow".to_string())
                    })?;
                    active = Some(field);
                }
                depth = checked_field_depth(depth)?;
                stack.push((namespace_uri, local));
            },
            Event::Empty(ref element) => {
                let local = utf8(element.local_name().as_ref(), "drop-down field element")?;
                if let Some(field) = active.as_mut() {
                    push_drop_down_label(
                        &reader,
                        element,
                        namespace_uri.as_deref(),
                        &local,
                        depth,
                        field,
                    )?;
                } else if namespace_uri.as_deref() == Some(TEXT_DATABASE_NAMESPACE)
                    && local == "drop-down"
                {
                    validate_drop_down_parent(stack.last())?;
                    if fields.len() >= MAX_FIELDS {
                        return Err(Error::InvalidFormat(format!(
                            "document exceeds {MAX_FIELDS} drop-down fields"
                        )));
                    }
                    fields.push(finish_drop_down_field(parse_drop_down_field(
                        &reader, element,
                    )?)?);
                }
            },
            Event::End(_) => {
                if let Some(field) = active.as_mut() {
                    if field.label_depth == Some(depth) {
                        field.label_depth = None;
                    } else if field.depth == depth {
                        let field = active.take().expect("checked drop-down field");
                        fields.push(finish_drop_down_field(field)?);
                    }
                }
                stack.pop().ok_or_else(|| {
                    Error::InvalidFormat("drop-down field XML stack underflow".to_string())
                })?;
                depth = depth.checked_sub(1).ok_or_else(|| {
                    Error::InvalidFormat("drop-down field XML depth underflow".to_string())
                })?;
            },
            Event::Text(ref text) if active.is_some() => {
                let value = text.xml_content(XmlVersion::Explicit1_0).map_err(|error| {
                    Error::InvalidFormat(format!("invalid drop-down field text: {error}"))
                })?;
                append_drop_down_text(
                    active.as_mut().expect("checked drop-down field"),
                    depth,
                    &value,
                )?;
            },
            Event::CData(ref text) if active.is_some() => {
                let value = text.xml_content(XmlVersion::Explicit1_0).map_err(|error| {
                    Error::InvalidFormat(format!("invalid drop-down field CDATA: {error}"))
                })?;
                append_drop_down_text(
                    active.as_mut().expect("checked drop-down field"),
                    depth,
                    &value,
                )?;
            },
            Event::GeneralRef(ref reference) if active.is_some() => {
                let value = decode_reference(reference, "drop-down field")?;
                append_drop_down_text(
                    active.as_mut().expect("checked drop-down field"),
                    depth,
                    &value,
                )?;
            },
            Event::DocType(_) => {
                return Err(Error::InvalidFormat(
                    "DOCTYPE is not permitted in ODF drop-down field XML".to_string(),
                ));
            },
            Event::PI(_) if active.is_some() => {
                return Err(Error::InvalidFormat(
                    "processing instructions are not permitted in text:drop-down".to_string(),
                ));
            },
            Event::Eof => break,
            _ => {},
        }
        buffer.clear();
    }
    if depth != 0 || active.is_some() || !stack.is_empty() {
        return Err(Error::InvalidFormat(
            "incomplete drop-down field XML structure".to_string(),
        ));
    }
    Ok(fields)
}

fn parse_drop_down_field(
    reader: &NsReader<&[u8]>,
    element: &quick_xml::events::BytesStart<'_>,
) -> Result<ActiveDropDownField> {
    let mut aggregate = 0usize;
    let attributes = drop_down_attributes(reader, element, &mut aggregate)?;
    reject_drop_down_attributes(&attributes, &[(TEXT_DATABASE_NAMESPACE, "name")])?;
    let name = required_drop_down_attribute(&attributes, TEXT_DATABASE_NAMESPACE, "name")?;
    Ok(ActiveDropDownField {
        depth: 0,
        label_depth: None,
        display_started: false,
        aggregate,
        name,
        labels: Vec::new(),
        display_text: String::new(),
    })
}

fn begin_drop_down_label(
    reader: &NsReader<&[u8]>,
    element: &quick_xml::events::BytesStart<'_>,
    namespace: Option<&str>,
    local: &str,
    depth: usize,
    field: &mut ActiveDropDownField,
) -> Result<()> {
    push_drop_down_label(reader, element, namespace, local, depth, field)?;
    field.label_depth = Some(
        depth
            .checked_add(1)
            .ok_or_else(|| Error::InvalidFormat("drop-down field depth overflow".to_string()))?,
    );
    Ok(())
}

fn push_drop_down_label(
    reader: &NsReader<&[u8]>,
    element: &quick_xml::events::BytesStart<'_>,
    namespace: Option<&str>,
    local: &str,
    depth: usize,
    field: &mut ActiveDropDownField,
) -> Result<()> {
    if namespace != Some(TEXT_DATABASE_NAMESPACE)
        || local != "label"
        || field.depth != depth
        || field.label_depth.is_some()
        || field.display_started
    {
        return Err(Error::InvalidFormat(
            "text:drop-down permits only leading empty text:label children".to_string(),
        ));
    }
    if field.labels.len() >= MAX_DROP_DOWN_LABELS {
        return Err(Error::InvalidFormat(format!(
            "text:drop-down exceeds {MAX_DROP_DOWN_LABELS} labels"
        )));
    }
    field.labels.push(parse_drop_down_label(
        reader,
        element,
        &mut field.aggregate,
    )?);
    Ok(())
}

fn parse_drop_down_label(
    reader: &NsReader<&[u8]>,
    element: &quick_xml::events::BytesStart<'_>,
    aggregate: &mut usize,
) -> Result<DropDownLabel> {
    let attributes = drop_down_attributes(reader, element, aggregate)?;
    reject_drop_down_attributes(
        &attributes,
        &[
            (TEXT_DATABASE_NAMESPACE, "value"),
            (TEXT_DATABASE_NAMESPACE, "current-selected"),
        ],
    )?;
    Ok(DropDownLabel {
        value: drop_down_attribute(&attributes, TEXT_DATABASE_NAMESPACE, "value")
            .map(str::to_string),
        current_selected: drop_down_attribute(
            &attributes,
            TEXT_DATABASE_NAMESPACE,
            "current-selected",
        )
        .map(parse_drop_down_boolean)
        .transpose()?,
    })
}

fn append_drop_down_text(field: &mut ActiveDropDownField, depth: usize, value: &str) -> Result<()> {
    if field.label_depth.is_some() || field.depth != depth {
        return Err(Error::InvalidFormat(
            "text:label must be empty in text:drop-down".to_string(),
        ));
    }
    field.display_started = true;
    validate_dynamic_value(
        "drop-down display text",
        Some(value),
        false,
        &mut field.aggregate,
    )?;
    field.display_text.push_str(value);
    Ok(())
}

fn finish_drop_down_field(field: ActiveDropDownField) -> Result<DynamicTextField> {
    if field.label_depth.is_some() {
        return Err(Error::InvalidFormat(
            "unterminated text:label in text:drop-down".to_string(),
        ));
    }
    let field = DynamicTextField::DropDown {
        name: field.name,
        labels: field.labels,
        display_text: field.display_text,
    };
    field.validate()?;
    Ok(field)
}

fn validate_drop_down_parent(parent: Option<&(Option<String>, String)>) -> Result<()> {
    if parent.is_some_and(|(namespace, local)| {
        namespace.as_deref() == Some(TEXT_DATABASE_NAMESPACE)
            && matches!(
                local.as_str(),
                "a" | "h" | "meta" | "meta-field" | "p" | "ruby-base" | "span"
            )
    }) {
        Ok(())
    } else {
        Err(Error::InvalidFormat(
            "text:drop-down occurs outside an ODF inline-text host".to_string(),
        ))
    }
}

fn parse_database_field(
    reader: &NsReader<&[u8]>,
    element: &quick_xml::events::BytesStart<'_>,
    kind: DatabaseFieldKind,
    aggregate: &mut usize,
) -> Result<DatabaseField> {
    let attributes = database_attributes(reader, element, aggregate)?;
    let allowed = match kind {
        DatabaseFieldKind::Display => &[
            (TEXT_DATABASE_NAMESPACE, "database-name"),
            (TEXT_DATABASE_NAMESPACE, "table-name"),
            (TEXT_DATABASE_NAMESPACE, "table-type"),
            (TEXT_DATABASE_NAMESPACE, "column-name"),
            (STYLE_NAMESPACE, "data-style-name"),
        ][..],
        DatabaseFieldKind::Next => &[
            (TEXT_DATABASE_NAMESPACE, "database-name"),
            (TEXT_DATABASE_NAMESPACE, "table-name"),
            (TEXT_DATABASE_NAMESPACE, "table-type"),
            (TEXT_DATABASE_NAMESPACE, "condition"),
        ][..],
        DatabaseFieldKind::RowSelect => &[
            (TEXT_DATABASE_NAMESPACE, "database-name"),
            (TEXT_DATABASE_NAMESPACE, "table-name"),
            (TEXT_DATABASE_NAMESPACE, "table-type"),
            (TEXT_DATABASE_NAMESPACE, "condition"),
            (TEXT_DATABASE_NAMESPACE, "row-number"),
        ][..],
        DatabaseFieldKind::RowNumber => &[
            (TEXT_DATABASE_NAMESPACE, "database-name"),
            (TEXT_DATABASE_NAMESPACE, "table-name"),
            (TEXT_DATABASE_NAMESPACE, "table-type"),
            (TEXT_DATABASE_NAMESPACE, "value"),
            (STYLE_NAMESPACE, "num-format"),
            (STYLE_NAMESPACE, "num-letter-sync"),
        ][..],
        DatabaseFieldKind::Name => &[
            (TEXT_DATABASE_NAMESPACE, "database-name"),
            (TEXT_DATABASE_NAMESPACE, "table-name"),
            (TEXT_DATABASE_NAMESPACE, "table-type"),
        ][..],
    };
    reject_database_attributes(&attributes, allowed)?;
    let table_name =
        required_database_attribute(&attributes, TEXT_DATABASE_NAMESPACE, "table-name")?;
    let table_type = database_attribute(&attributes, TEXT_DATABASE_NAMESPACE, "table-type")
        .map(DatabaseTableType::parse)
        .transpose()?;
    let row_number = database_attribute(&attributes, TEXT_DATABASE_NAMESPACE, "row-number")
        .map(NonNegativeInteger::new)
        .transpose()?;
    let value = database_attribute(&attributes, TEXT_DATABASE_NAMESPACE, "value")
        .map(NonNegativeInteger::new)
        .transpose()?;
    let number_letter_sync = database_attribute(&attributes, STYLE_NAMESPACE, "num-letter-sync")
        .map(parse_database_bool)
        .transpose()?;
    Ok(DatabaseField {
        kind,
        source: DatabaseSource {
            database_name: database_attribute(
                &attributes,
                TEXT_DATABASE_NAMESPACE,
                "database-name",
            )
            .map(str::to_string),
            table_name,
            table_type,
            connection_resource: None,
        },
        column_name: database_attribute(&attributes, TEXT_DATABASE_NAMESPACE, "column-name")
            .map(str::to_string),
        condition: database_attribute(&attributes, TEXT_DATABASE_NAMESPACE, "condition")
            .map(str::to_string),
        row_number,
        value,
        data_style_name: database_attribute(&attributes, STYLE_NAMESPACE, "data-style-name")
            .map(str::to_string),
        number_format: database_attribute(&attributes, STYLE_NAMESPACE, "num-format")
            .map(str::to_string),
        number_letter_sync,
        display_text: String::new(),
    })
}

fn validate_database_field(field: DatabaseField) -> Result<DatabaseField> {
    match field.kind {
        DatabaseFieldKind::Display if field.column_name.is_none() => {
            return Err(Error::InvalidFormat(
                "text:database-display requires text:column-name".to_string(),
            ));
        },
        DatabaseFieldKind::Next | DatabaseFieldKind::RowSelect
            if !field.display_text.is_empty() =>
        {
            return Err(Error::InvalidFormat(
                "database selection fields cannot contain character data".to_string(),
            ));
        },
        _ => {},
    }
    if field.number_letter_sync.is_some()
        && !matches!(field.number_format.as_deref(), Some("a" | "A"))
    {
        return Err(Error::InvalidFormat(
            "style:num-letter-sync requires style:num-format a or A".to_string(),
        ));
    }
    Ok(field)
}

fn validate_constructed_database_field(field: &DatabaseField) -> Result<()> {
    let mut aggregate = 0usize;
    for value in [
        field.source.database_name.as_deref(),
        Some(field.source.table_name.as_str()),
        field.column_name.as_deref(),
        field.condition.as_deref(),
        field.data_style_name.as_deref(),
        field.number_format.as_deref(),
        Some(field.display_text.as_str()),
        field
            .source
            .connection_resource
            .as_ref()
            .map(|resource| resource.href.as_str()),
    ]
    .into_iter()
    .flatten()
    {
        if !value.chars().all(is_xml_1_0_char) {
            return Err(Error::InvalidFormat(
                "database field contains forbidden XML characters".to_string(),
            ));
        }
        append_database_size(&mut aggregate, value.len())?;
    }
    if field
        .source
        .connection_resource
        .as_ref()
        .is_some_and(|resource| !resource.simple_link)
    {
        return Err(Error::InvalidFormat(
            "ODF form:connection-resource only supports xlink:href".to_string(),
        ));
    }
    let forbidden = match field.kind {
        DatabaseFieldKind::Display => {
            field.condition.is_some()
                || field.row_number.is_some()
                || field.value.is_some()
                || field.number_format.is_some()
                || field.number_letter_sync.is_some()
        },
        DatabaseFieldKind::Next => {
            field.column_name.is_some()
                || field.row_number.is_some()
                || field.value.is_some()
                || field.data_style_name.is_some()
                || field.number_format.is_some()
                || field.number_letter_sync.is_some()
        },
        DatabaseFieldKind::RowSelect => {
            field.column_name.is_some()
                || field.value.is_some()
                || field.data_style_name.is_some()
                || field.number_format.is_some()
                || field.number_letter_sync.is_some()
        },
        DatabaseFieldKind::RowNumber => {
            field.column_name.is_some()
                || field.condition.is_some()
                || field.row_number.is_some()
                || field.data_style_name.is_some()
        },
        DatabaseFieldKind::Name => {
            field.column_name.is_some()
                || field.condition.is_some()
                || field.row_number.is_some()
                || field.value.is_some()
                || field.data_style_name.is_some()
                || field.number_format.is_some()
                || field.number_letter_sync.is_some()
        },
    };
    if forbidden {
        return Err(Error::InvalidFormat(
            "database field contains attributes from another field kind".to_string(),
        ));
    }
    Ok(())
}

fn parse_connection_resource(
    reader: &NsReader<&[u8]>,
    element: &quick_xml::events::BytesStart<'_>,
    aggregate: &mut usize,
) -> Result<DatabaseConnectionResource> {
    let attributes = database_attributes(reader, element, aggregate)?;
    reject_database_attributes(&attributes, &[(XLINK_NAMESPACE, "href")])?;
    let href = required_database_attribute(&attributes, XLINK_NAMESPACE, "href")?;
    Ok(DatabaseConnectionResource {
        href,
        simple_link: true,
    })
}

fn drop_down_attributes(
    reader: &NsReader<&[u8]>,
    element: &quick_xml::events::BytesStart<'_>,
    aggregate: &mut usize,
) -> Result<DatabaseAttributes> {
    let mut attributes = HashMap::new();
    for attribute in element.attributes().with_checks(true) {
        let attribute = attribute.map_err(|error| {
            Error::InvalidFormat(format!("invalid drop-down field attribute: {error}"))
        })?;
        let raw = attribute.key.as_ref();
        if raw == b"xmlns" || raw.starts_with(b"xmlns:") {
            continue;
        }
        let (namespace, local) = reader.resolver().resolve_attribute(attribute.key);
        let namespace = resolved_namespace(&namespace)?.unwrap_or_default();
        let local = utf8(local.as_ref(), "drop-down field attribute")?;
        let value = attribute
            .decoded_and_normalized_value(XmlVersion::Explicit1_0, reader.decoder())
            .map_err(|error| {
                Error::InvalidFormat(format!("invalid drop-down field attribute value: {error}"))
            })?
            .into_owned();
        validate_dynamic_value("drop-down field attribute", Some(&value), false, aggregate)?;
        if attributes.insert((namespace, local), value).is_some() {
            return Err(Error::InvalidFormat(
                "duplicate expanded drop-down field attribute".to_string(),
            ));
        }
    }
    Ok(attributes)
}

fn reject_drop_down_attributes(
    attributes: &DatabaseAttributes,
    allowed: &[(&str, &str)],
) -> Result<()> {
    for (namespace, local) in attributes.keys() {
        if !allowed.iter().any(|(allowed_namespace, allowed_local)| {
            namespace == allowed_namespace && local == allowed_local
        }) {
            return Err(Error::InvalidFormat(format!(
                "unexpected drop-down field attribute {namespace}:{local}"
            )));
        }
    }
    Ok(())
}

fn drop_down_attribute<'a>(
    attributes: &'a DatabaseAttributes,
    namespace: &str,
    local: &str,
) -> Option<&'a str> {
    attributes
        .get(&(namespace.to_string(), local.to_string()))
        .map(String::as_str)
}

fn required_drop_down_attribute(
    attributes: &DatabaseAttributes,
    namespace: &str,
    local: &str,
) -> Result<String> {
    drop_down_attribute(attributes, namespace, local)
        .map(str::to_string)
        .ok_or_else(|| Error::InvalidFormat(format!("text:drop-down requires text:{local}")))
}

fn database_attributes(
    reader: &NsReader<&[u8]>,
    element: &quick_xml::events::BytesStart<'_>,
    aggregate: &mut usize,
) -> Result<DatabaseAttributes> {
    let mut attributes = HashMap::new();
    for attribute in element.attributes().with_checks(true) {
        let attribute = attribute.map_err(|error| {
            Error::InvalidFormat(format!("invalid database field attribute: {error}"))
        })?;
        let raw = attribute.key.as_ref();
        if raw == b"xmlns" || raw.starts_with(b"xmlns:") {
            continue;
        }
        let (namespace, local) = reader.resolver().resolve_attribute(attribute.key);
        let namespace = resolved_namespace(&namespace)?.unwrap_or_default();
        let local = utf8(local.as_ref(), "database field attribute")?;
        let value = attribute
            .decoded_and_normalized_value(XmlVersion::Explicit1_0, reader.decoder())
            .map_err(|error| {
                Error::InvalidFormat(format!("invalid database field attribute value: {error}"))
            })?
            .into_owned();
        append_database_size(aggregate, value.len())?;
        if attributes.insert((namespace, local), value).is_some() {
            return Err(Error::InvalidFormat(
                "duplicate expanded database field attribute".to_string(),
            ));
        }
    }
    Ok(attributes)
}

fn reject_database_attributes(
    attributes: &DatabaseAttributes,
    allowed: &[(&str, &str)],
) -> Result<()> {
    for (namespace, local) in attributes.keys() {
        if !allowed.iter().any(|(allowed_namespace, allowed_local)| {
            namespace == allowed_namespace && local == allowed_local
        }) {
            return Err(Error::InvalidFormat(format!(
                "unexpected database field attribute {namespace}:{local}"
            )));
        }
    }
    Ok(())
}

fn append_database_text(
    active: &mut ActiveDatabaseField,
    value: &str,
    aggregate: &mut usize,
) -> Result<()> {
    if active.connection_depth.is_some() {
        if value.is_empty() {
            return Ok(());
        }
        return Err(Error::InvalidFormat(
            "form:connection-resource must be empty".to_string(),
        ));
    }
    if active.field.display_text.len().saturating_add(value.len()) > MAX_DATABASE_VALUE {
        return Err(Error::InvalidFormat(
            "ODF database field display text exceeds the supported limit".to_string(),
        ));
    }
    append_database_size(aggregate, value.len())?;
    active.field.display_text.push_str(value);
    Ok(())
}

fn append_database_size(aggregate: &mut usize, amount: usize) -> Result<()> {
    if amount > MAX_DATABASE_VALUE {
        return Err(Error::InvalidFormat(
            "database field value exceeds 64 KiB".to_string(),
        ));
    }
    *aggregate = aggregate.checked_add(amount).ok_or_else(|| {
        Error::InvalidFormat("database field aggregate size overflow".to_string())
    })?;
    if *aggregate > MAX_DATABASE_AGGREGATE {
        return Err(Error::InvalidFormat(
            "database field metadata exceeds 16 MiB".to_string(),
        ));
    }
    Ok(())
}

fn database_field_kind(local: &str) -> Option<DatabaseFieldKind> {
    match local {
        "database-display" => Some(DatabaseFieldKind::Display),
        "database-next" => Some(DatabaseFieldKind::Next),
        "database-row-select" => Some(DatabaseFieldKind::RowSelect),
        "database-row-number" => Some(DatabaseFieldKind::RowNumber),
        "database-name" => Some(DatabaseFieldKind::Name),
        _ => None,
    }
}

fn database_attribute<'a>(
    attributes: &'a DatabaseAttributes,
    namespace: &str,
    local: &str,
) -> Option<&'a str> {
    attributes
        .get(&(namespace.to_string(), local.to_string()))
        .map(String::as_str)
}

fn required_database_attribute(
    attributes: &DatabaseAttributes,
    namespace: &str,
    local: &str,
) -> Result<String> {
    database_attribute(attributes, namespace, local)
        .map(str::to_string)
        .ok_or_else(|| Error::InvalidFormat(format!("database field requires {local}")))
}

fn parse_database_bool(value: &str) -> Result<bool> {
    match value {
        "true" | "1" => Ok(true),
        "false" | "0" => Ok(false),
        _ => Err(Error::InvalidFormat(format!(
            "invalid database field boolean '{value}'"
        ))),
    }
}

fn validate_database_parent(parent: Option<&(Option<String>, String)>) -> Result<()> {
    if parent.is_some_and(|(namespace, local)| {
        namespace.as_deref() == Some(TEXT_DATABASE_NAMESPACE)
            && matches!(
                local.as_str(),
                "a" | "h" | "meta" | "meta-field" | "p" | "ruby-base" | "span"
            )
    }) {
        Ok(())
    } else {
        Err(Error::InvalidFormat(
            "database field occurs outside an ODF inline-text host".to_string(),
        ))
    }
}

fn resolved_namespace(namespace: &quick_xml::name::ResolveResult<'_>) -> Result<Option<String>> {
    match namespace {
        quick_xml::name::ResolveResult::Bound(quick_xml::name::Namespace(value)) => {
            Ok(Some(utf8(value, "namespace URI")?))
        },
        quick_xml::name::ResolveResult::Unbound => Ok(None),
        quick_xml::name::ResolveResult::Unknown(prefix) => Err(Error::InvalidFormat(format!(
            "unbound namespace prefix '{}'",
            String::from_utf8_lossy(prefix)
        ))),
    }
}

fn utf8(value: &[u8], description: &str) -> Result<String> {
    std::str::from_utf8(value)
        .map(str::to_string)
        .map_err(|_| Error::InvalidFormat(format!("invalid UTF-8 {description}")))
}

fn resolve_database_reference(name: &str) -> Result<String> {
    if let Some(value) = quick_xml::escape::resolve_xml_entity(name) {
        return Ok(value.to_string());
    }
    let codepoint =
        if let Some(value) = name.strip_prefix("#x").or_else(|| name.strip_prefix("#X")) {
            u32::from_str_radix(value, 16)
        } else if let Some(value) = name.strip_prefix('#') {
            value.parse::<u32>()
        } else {
            return Err(Error::InvalidFormat(
                "undeclared entity in database field".to_string(),
            ));
        }
        .map_err(|_| {
            Error::InvalidFormat("invalid database field character reference".to_string())
        })?;
    char::from_u32(codepoint)
        .filter(|value| {
            matches!(*value, '\u{9}' | '\u{a}' | '\u{d}')
                || matches!(*value as u32, 0x20..=0xd7ff | 0xe000..=0xfffd | 0x10000..=0x10ffff)
        })
        .map(|value| value.to_string())
        .ok_or_else(|| Error::InvalidFormat("invalid XML character reference".to_string()))
}
