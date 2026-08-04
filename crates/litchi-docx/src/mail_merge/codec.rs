//! Bounded WordprocessingML mail-merge XML codec.

use super::model::{
    Conformance, DataSourceObject, DataType, Destination, FieldMap, FieldMappingType,
    MAX_ATTRIBUTES_PER_NODE, MAX_DEPTH, MAX_FIELD_MAPS, MAX_NODES, MAX_RECIPIENTS,
    MAX_RELATIONSHIP_ID_BYTES, MAX_STRING_BYTES, MAX_UNIQUE_TAG_BYTES, MainDocumentType, R,
    Recipient, Recipients, STRICT_R, STRICT_W, Settings, W, invalid,
};
use crate::{Error, Result};
use base64::Engine;
use base64::engine::general_purpose::STANDARD as BASE64;
use quick_xml::XmlVersion;
use quick_xml::events::{BytesStart, Event};
use quick_xml::name::{Namespace, NamespaceResolver, ResolveResult};
use quick_xml::reader::NsReader;

impl Settings {
    /// Serialize a standalone `w:mailMerge` fragment in schema order.
    pub fn to_xml(&self, conformance: Conformance) -> Result<String> {
        validate_model(self)?;
        let mut xml = format!(
            r#"<w:mailMerge xmlns:w="{}" xmlns:r="{}">"#,
            conformance.word(),
            conformance.relationships()
        );
        if self.main_document_type != MainDocumentType::FormLetters {
            value_leaf(
                &mut xml,
                "mainDocumentType",
                self.main_document_type.as_str(),
            );
        }
        on_off_leaf(&mut xml, "linkToQuery", self.link_to_query);
        if let Some(value) = self.data_type {
            value_leaf(&mut xml, "dataType", value.as_str());
        }
        optional_string_leaf(&mut xml, "connectString", self.connect_string.as_deref());
        optional_string_leaf(&mut xml, "query", self.query.as_deref());
        relationship_leaf(
            &mut xml,
            "dataSource",
            self.data_source_relationship_id.as_deref(),
        );
        relationship_leaf(
            &mut xml,
            "headerSource",
            self.header_source_relationship_id.as_deref(),
        );
        on_off_leaf(
            &mut xml,
            "doNotSuppressBlankLines",
            self.do_not_suppress_blank_lines,
        );
        if self.destination != Destination::NewDocument {
            value_leaf(&mut xml, "destination", self.destination.as_str());
        }
        optional_string_leaf(
            &mut xml,
            "addressFieldName",
            self.address_field_name.as_deref(),
        );
        optional_string_leaf(&mut xml, "mailSubject", self.mail_subject.as_deref());
        on_off_leaf(&mut xml, "mailAsAttachment", self.mail_as_attachment);
        on_off_leaf(&mut xml, "viewMergedData", self.view_merged_data);
        if self.active_record != 1 {
            value_leaf(&mut xml, "activeRecord", &self.active_record.to_string());
        }
        if self.check_errors != 2 {
            value_leaf(&mut xml, "checkErrors", &self.check_errors.to_string());
        }
        if let Some(odso) = &self.odso {
            write_odso(&mut xml, odso);
        }
        xml.push_str("</w:mailMerge>");
        Ok(xml)
    }
}

impl Recipients {
    pub fn parse_xml(xml: &[u8]) -> Result<Self> {
        let root = parse_tree(xml)?;
        require_word_element(&root, "recipients")?;
        ensure_no_schema_attrs(&root)?;
        let mut recipients = Vec::new();
        for child in &root.children {
            if !child.is_word() {
                if child.local == "recipientData" {
                    return Err(invalid("spoofed recipientData element namespace"));
                }
                continue;
            }
            if child.local != "recipientData" {
                return Err(invalid(format!(
                    "unexpected recipients child '{}'",
                    child.local
                )));
            }
            if recipients.len() >= MAX_RECIPIENTS {
                return Err(invalid("too many mail-merge recipients"));
            }
            recipients.push(parse_recipient(child)?);
        }
        Ok(Self { recipients })
    }

    pub fn to_xml(&self, conformance: Conformance) -> Result<String> {
        if self.recipients.len() > MAX_RECIPIENTS {
            return Err(invalid("too many mail-merge recipients"));
        }
        let mut xml = format!(r#"<w:recipients xmlns:w="{}">"#, conformance.word());
        for recipient in &self.recipients {
            xml.push_str("<w:recipientData>");
            if !recipient.active {
                xml.push_str(r#"<w:active w:val="0"/>"#);
            }
            if let Some(column) = recipient.column {
                if column < 0 {
                    return Err(invalid("mail-merge recipient column cannot be negative"));
                }
                value_leaf(&mut xml, "column", &column.to_string());
            }
            if let Some(tag) = &recipient.unique_tag {
                if tag.is_empty() || tag.len() > MAX_UNIQUE_TAG_BYTES {
                    return Err(invalid(
                        "mail-merge recipient unique tag has invalid length",
                    ));
                }
                value_leaf(&mut xml, "uniqueTag", &BASE64.encode(tag));
            }
            xml.push_str("</w:recipientData>");
        }
        xml.push_str("</w:recipients>");
        Ok(xml)
    }
}

pub fn parse_settings_mail_merge(xml: &[u8]) -> Result<Option<Settings>> {
    let root = parse_tree(xml)?;
    require_word_element(&root, "settings")?;
    let mut found = None;
    let mut mail_index = None;
    for (index, child) in root.children.iter().enumerate() {
        if child.local == "mailMerge" {
            if !child.is_word() {
                return Err(invalid("spoofed mailMerge element namespace"));
            }
            if found.is_some() {
                return Err(invalid("duplicate mailMerge setting"));
            }
            found = Some(parse_mail_merge(child)?);
            mail_index = Some(index);
        }
        reject_nested_mail_merge(child, child.local == "mailMerge")?;
    }
    if let Some(index) = mail_index {
        validate_settings_order(&root.children, index)?;
    }
    Ok(found)
}

fn reject_nested_mail_merge(node: &Node, is_direct: bool) -> Result<()> {
    for child in &node.children {
        if child.local == "mailMerge" && child.is_word() && !is_direct {
            return Err(invalid("mailMerge must be a direct settings child"));
        }
        reject_nested_mail_merge(child, false)?;
    }
    Ok(())
}

fn validate_settings_order(children: &[Node], mail_index: usize) -> Result<()> {
    const BEFORE: &[&str] = &[
        "writeProtection",
        "view",
        "zoom",
        "linkStyles",
        "removePersonalInformation",
        "removeDateAndTime",
        "doNotDisplayPageBoundaries",
        "displayBackgroundShape",
        "printPostScriptOverText",
        "printFractionalCharacterWidth",
        "printFormsData",
        "embedTrueTypeFonts",
        "embedSystemFonts",
        "saveSubsetFonts",
        "saveFormsData",
        "mirrorMargins",
        "alignBordersAndEdges",
        "bordersDoNotSurroundHeader",
        "bordersDoNotSurroundFooter",
        "gutterAtTop",
        "hideSpellingErrors",
        "hideGrammaticalErrors",
        "activeWritingStyle",
        "proofState",
        "formsDesign",
        "attachedTemplate",
        "stylePaneFormatFilter",
        "stylePaneSortMethod",
        "documentType",
    ];
    const AFTER: &[&str] = &[
        "revisionView",
        "trackRevisions",
        "doNotTrackMoves",
        "doNotTrackFormatting",
        "documentProtection",
        "autoFormatOverride",
        "styleLockTheme",
        "styleLockQFSet",
        "defaultTabStop",
    ];
    for child in &children[..mail_index] {
        if child.is_word() && AFTER.contains(&child.local.as_str()) {
            return Err(invalid(format!("{} must follow mailMerge", child.local)));
        }
    }
    for child in &children[mail_index + 1..] {
        if child.is_word() && BEFORE.contains(&child.local.as_str()) {
            return Err(invalid(format!("{} must precede mailMerge", child.local)));
        }
    }
    Ok(())
}

fn parse_mail_merge(node: &Node) -> Result<Settings> {
    ensure_no_schema_attrs(node)?;
    let names = [
        "mainDocumentType",
        "linkToQuery",
        "dataType",
        "connectString",
        "query",
        "dataSource",
        "headerSource",
        "doNotSuppressBlankLines",
        "destination",
        "addressFieldName",
        "mailSubject",
        "mailAsAttachment",
        "viewMergedData",
        "activeRecord",
        "checkErrors",
        "odso",
    ];
    let mut seen = [false; 16];
    let mut last = 0usize;
    let mut first = true;
    let mut value = Settings::default();
    for child in &node.children {
        if !child.is_word() {
            if names.contains(&child.local.as_str()) {
                return Err(invalid(format!(
                    "spoofed {} element namespace",
                    child.local
                )));
            }
            continue;
        }
        let Some(index) = names.iter().position(|name| *name == child.local) else {
            return Err(invalid(format!(
                "unexpected mailMerge child '{}'",
                child.local
            )));
        };
        if seen[index] {
            return Err(invalid(format!(
                "duplicate mailMerge child '{}'",
                child.local
            )));
        }
        if !first && index < last {
            return Err(invalid(format!(
                "mailMerge child '{}' is out of order",
                child.local
            )));
        }
        first = false;
        last = index;
        seen[index] = true;
        match index {
            0 => value.main_document_type = MainDocumentType::parse(&required_val(child)?)?,
            1 => value.link_to_query = on_off(child)?,
            2 => value.data_type = Some(DataType::parse(&required_val(child)?)?),
            3 => {
                value.connect_string = Some(bounded_string(required_val(child)?, "connectString")?)
            },
            4 => value.query = Some(bounded_string(required_val(child)?, "query")?),
            5 => value.data_source_relationship_id = Some(relationship_id(child)?),
            6 => value.header_source_relationship_id = Some(relationship_id(child)?),
            7 => value.do_not_suppress_blank_lines = on_off(child)?,
            8 => value.destination = Destination::parse(&required_val(child)?)?,
            9 => {
                value.address_field_name =
                    Some(bounded_string(required_val(child)?, "addressFieldName")?)
            },
            10 => value.mail_subject = Some(bounded_string(required_val(child)?, "mailSubject")?),
            11 => value.mail_as_attachment = on_off(child)?,
            12 => value.view_merged_data = on_off(child)?,
            13 => {
                value.active_record = decimal(child, "activeRecord")?;
                if value.active_record < 1 {
                    return Err(invalid("activeRecord must be at least 1"));
                }
            },
            14 => value.check_errors = decimal(child, "checkErrors")?,
            15 => value.odso = Some(parse_odso(child)?),
            _ => return Err(invalid("mailMerge child index is out of range")),
        }
    }
    validate_model(&value)?;
    Ok(value)
}

fn parse_odso(node: &Node) -> Result<DataSourceObject> {
    ensure_no_schema_attrs(node)?;
    let names = [
        "udl",
        "table",
        "src",
        "colDelim",
        "type",
        "fHdr",
        "fieldMapData",
        "recipientData",
    ];
    let mut seen = [false; 8];
    let mut last = 0usize;
    let mut first = true;
    let mut value = DataSourceObject::default();
    for child in &node.children {
        if !child.is_word() {
            if names.contains(&child.local.as_str()) {
                return Err(invalid(format!(
                    "spoofed ODSO {} element namespace",
                    child.local
                )));
            }
            continue;
        }
        let Some(index) = names.iter().position(|name| *name == child.local) else {
            return Err(invalid(format!("unexpected odso child '{}'", child.local)));
        };
        if index != 6 && seen[index] {
            return Err(invalid(format!("duplicate odso child '{}'", child.local)));
        }
        if !first && index < last {
            return Err(invalid(format!(
                "odso child '{}' is out of order",
                child.local
            )));
        }
        first = false;
        last = index;
        seen[index] = true;
        match index {
            0 => value.udl = Some(bounded_string(required_val(child)?, "odso udl")?),
            1 => value.table = Some(bounded_string(required_val(child)?, "odso table")?),
            2 => value.source_relationship_id = Some(relationship_id(child)?),
            3 => {
                let number = decimal(child, "colDelim")?;
                if number < 0 {
                    return Err(invalid("colDelim cannot be negative"));
                }
                value.column_delimiter = Some(number);
            },
            4 => {
                value.source_type = Some(bounded_string(required_val(child)?, "odso source type")?)
            },
            5 => value.first_row_header = on_off(child)?,
            6 => {
                if value.field_maps.len() >= MAX_FIELD_MAPS {
                    return Err(invalid("too many odso field maps"));
                }
                value.field_maps.push(parse_field_map(child)?);
            },
            7 => value.recipient_data_relationship_id = Some(relationship_id(child)?),
            _ => return Err(invalid("odso child index is out of range")),
        }
    }
    Ok(value)
}

fn parse_field_map(node: &Node) -> Result<FieldMap> {
    ensure_no_schema_attrs(node)?;
    let names = [
        "type",
        "name",
        "mappedName",
        "column",
        "lid",
        "dynamicAddress",
    ];
    let mut seen = [false; 6];
    let mut last = 0usize;
    let mut first = true;
    let mut value = FieldMap::default();
    for child in &node.children {
        if !child.is_word() {
            if names.contains(&child.local.as_str()) {
                return Err(invalid(format!(
                    "spoofed fieldMapData {} element namespace",
                    child.local
                )));
            }
            continue;
        }
        let Some(index) = names.iter().position(|name| *name == child.local) else {
            return Err(invalid(format!(
                "unexpected fieldMapData child '{}'",
                child.local
            )));
        };
        if seen[index] {
            return Err(invalid(format!(
                "duplicate fieldMapData child '{}'",
                child.local
            )));
        }
        if !first && index < last {
            return Err(invalid(format!(
                "fieldMapData child '{}' is out of order",
                child.local
            )));
        }
        first = false;
        last = index;
        seen[index] = true;
        match index {
            0 => value.mapping_type = Some(FieldMappingType::parse(&required_val(child)?)?),
            1 => value.name = Some(bounded_string(required_val(child)?, "field-map name")?),
            2 => {
                value.mapped_name = Some(bounded_string(required_val(child)?, "mapped field name")?)
            },
            3 => {
                let number = decimal(child, "field-map column")?;
                if number < 0 {
                    return Err(invalid("field-map column cannot be negative"));
                }
                value.column = Some(number);
            },
            4 => {
                value.language_id =
                    Some(bounded_string(required_val(child)?, "field-map language")?)
            },
            5 => value.dynamic_address = on_off(child)?,
            _ => return Err(invalid("fieldMapData child index is out of range")),
        }
    }
    Ok(value)
}

fn parse_recipient(node: &Node) -> Result<Recipient> {
    ensure_no_schema_attrs(node)?;
    let names = ["active", "column", "uniqueTag"];
    let mut seen = [false; 3];
    let mut last = 0usize;
    let mut first = true;
    let mut recipient = Recipient {
        active: true,
        column: None,
        unique_tag: None,
    };
    for child in &node.children {
        if !child.is_word() {
            if names.contains(&child.local.as_str()) {
                return Err(invalid(format!(
                    "spoofed recipient {} element namespace",
                    child.local
                )));
            }
            continue;
        }
        let Some(index) = names.iter().position(|name| *name == child.local) else {
            return Err(invalid(format!(
                "unexpected recipientData child '{}'",
                child.local
            )));
        };
        if seen[index] {
            return Err(invalid(format!(
                "duplicate recipientData child '{}'",
                child.local
            )));
        }
        if !first && index < last {
            return Err(invalid(format!(
                "recipientData child '{}' is out of order",
                child.local
            )));
        }
        first = false;
        last = index;
        seen[index] = true;
        match index {
            0 => recipient.active = on_off(child)?,
            1 => {
                let number = decimal(child, "recipient column")?;
                if number < 0 {
                    return Err(invalid("recipient column cannot be negative"));
                }
                recipient.column = Some(number);
            },
            2 => recipient.unique_tag = Some(strict_base64(&required_val(child)?)?),
            _ => return Err(invalid("recipientData child index is out of range")),
        }
    }
    Ok(recipient)
}

fn validate_model(value: &Settings) -> Result<()> {
    for (description, string) in [
        ("connectString", value.connect_string.as_deref()),
        ("query", value.query.as_deref()),
        ("addressFieldName", value.address_field_name.as_deref()),
        ("mailSubject", value.mail_subject.as_deref()),
    ] {
        if string.is_some_and(|text| text.len() > MAX_STRING_BYTES) {
            return Err(invalid(format!("{description} is too large")));
        }
    }
    if value.active_record < 1 {
        return Err(invalid("activeRecord must be at least 1"));
    }
    if let Some(odso) = &value.odso
        && odso.field_maps.len() > MAX_FIELD_MAPS
    {
        return Err(invalid("too many odso field maps"));
    }
    Ok(())
}

fn write_odso(xml: &mut String, odso: &DataSourceObject) {
    xml.push_str("<w:odso>");
    optional_string_leaf(xml, "udl", odso.udl.as_deref());
    optional_string_leaf(xml, "table", odso.table.as_deref());
    relationship_leaf(xml, "src", odso.source_relationship_id.as_deref());
    if let Some(value) = odso.column_delimiter {
        value_leaf(xml, "colDelim", &value.to_string());
    }
    optional_string_leaf(xml, "type", odso.source_type.as_deref());
    on_off_leaf(xml, "fHdr", odso.first_row_header);
    for field in &odso.field_maps {
        xml.push_str("<w:fieldMapData>");
        if let Some(value) = field.mapping_type {
            value_leaf(xml, "type", value.as_str());
        }
        optional_string_leaf(xml, "name", field.name.as_deref());
        optional_string_leaf(xml, "mappedName", field.mapped_name.as_deref());
        if let Some(value) = field.column {
            value_leaf(xml, "column", &value.to_string());
        }
        optional_string_leaf(xml, "lid", field.language_id.as_deref());
        on_off_leaf(xml, "dynamicAddress", field.dynamic_address);
        xml.push_str("</w:fieldMapData>");
    }
    relationship_leaf(
        xml,
        "recipientData",
        odso.recipient_data_relationship_id.as_deref(),
    );
    xml.push_str("</w:odso>");
}

fn value_leaf(xml: &mut String, name: &str, value: &str) {
    xml.push_str("<w:");
    xml.push_str(name);
    xml.push_str(" w:val=\"");
    escape(xml, value);
    xml.push_str("\"/>");
}
fn optional_string_leaf(xml: &mut String, name: &str, value: Option<&str>) {
    if let Some(value) = value {
        value_leaf(xml, name, value);
    }
}
fn relationship_leaf(xml: &mut String, name: &str, value: Option<&str>) {
    if let Some(value) = value {
        xml.push_str("<w:");
        xml.push_str(name);
        xml.push_str(" r:id=\"");
        escape(xml, value);
        xml.push_str("\"/>");
    }
}
fn on_off_leaf(xml: &mut String, name: &str, value: bool) {
    if value {
        xml.push_str("<w:");
        xml.push_str(name);
        xml.push_str("/>");
    }
}
fn escape(xml: &mut String, value: &str) {
    for character in value.chars() {
        match character {
            '&' => xml.push_str("&amp;"),
            '<' => xml.push_str("&lt;"),
            '>' => xml.push_str("&gt;"),
            '"' => xml.push_str("&quot;"),
            '\'' => xml.push_str("&apos;"),
            _ => xml.push(character),
        }
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
enum OwnedNamespace {
    Bound(String),
    Unbound,
    Unknown(String),
}

#[derive(Debug, Clone)]
struct Attribute {
    namespace: OwnedNamespace,
    local: String,
    value: String,
}

#[derive(Debug, Clone)]
struct Node {
    namespace: OwnedNamespace,
    local: String,
    attributes: Vec<Attribute>,
    children: Vec<Node>,
    has_text: bool,
}

impl Node {
    fn is_word(&self) -> bool {
        matches!(&self.namespace, OwnedNamespace::Bound(value) if value == W || value == STRICT_W)
    }
}

// The Text/CData arms keep their `?`-bearing whitespace checks out of the
// match guards on purpose; guards cannot use `?`.
#[allow(clippy::collapsible_match)]
fn parse_tree(xml: &[u8]) -> Result<Node> {
    let processed = litchi_ooxml_common::mce::process_ooxml(xml)
        .map_err(|error| invalid(format!("mail-merge MCE error: {error}")))?;
    let mut reader = NsReader::from_reader(processed.as_ref());
    reader.config_mut().trim_text(false);
    reader.config_mut().check_end_names = true;
    let mut stack: Vec<Node> = Vec::new();
    let mut root = None;
    let mut nodes = 0usize;
    loop {
        let decoder = reader.decoder();
        let event = reader
            .read_event()
            .map_err(|error| Error::Xml(error.to_string()))?
            .into_owned();
        let resolver = reader.resolver().clone();
        let (namespace, event) = resolver.resolve_event(event);
        match event {
            Event::Start(element) => {
                if stack.len() >= MAX_DEPTH {
                    return Err(invalid("mail-merge XML nesting is too deep"));
                }
                nodes = nodes
                    .checked_add(1)
                    .ok_or_else(|| invalid("mail-merge XML node count overflow"))?;
                if nodes > MAX_NODES {
                    return Err(invalid("mail-merge XML has too many nodes"));
                }
                stack.push(make_node(namespace, &element, decoder, &resolver)?);
            },
            Event::Empty(element) => {
                nodes = nodes
                    .checked_add(1)
                    .ok_or_else(|| invalid("mail-merge XML node count overflow"))?;
                if nodes > MAX_NODES {
                    return Err(invalid("mail-merge XML has too many nodes"));
                }
                append_node(
                    make_node(namespace, &element, decoder, &resolver)?,
                    &mut stack,
                    &mut root,
                )?;
            },
            Event::End(_) => {
                let node = stack
                    .pop()
                    .ok_or_else(|| invalid("unexpected mail-merge XML end element"))?;
                append_node(node, &mut stack, &mut root)?;
            },
            Event::Text(text) => {
                if !text
                    .decode()
                    .map_err(|error| Error::Xml(error.to_string()))?
                    .trim()
                    .is_empty()
                {
                    if let Some(node) = stack.last_mut() {
                        node.has_text = true;
                    } else {
                        return Err(invalid("text outside mail-merge XML root"));
                    }
                }
            },
            Event::CData(text) => {
                if !text
                    .decode()
                    .map_err(|error| Error::Xml(error.to_string()))?
                    .trim()
                    .is_empty()
                {
                    if let Some(node) = stack.last_mut() {
                        node.has_text = true;
                    } else {
                        return Err(invalid("CDATA outside mail-merge XML root"));
                    }
                }
            },
            Event::DocType(_) | Event::PI(_) => {
                return Err(invalid(
                    "DTD and processing instructions are rejected in mail-merge XML",
                ));
            },
            Event::Eof if !stack.is_empty() => return Err(invalid("unterminated mail-merge XML")),
            Event::Eof => break,
            _ => {},
        }
    }
    root.ok_or_else(|| invalid("mail-merge XML has no root element"))
}

fn make_node(
    namespace: ResolveResult<'_>,
    element: &BytesStart<'_>,
    decoder: quick_xml::encoding::Decoder,
    resolver: &NamespaceResolver,
) -> Result<Node> {
    let mut attributes = Vec::new();
    for attribute in element.attributes() {
        let attribute = attribute.map_err(|error| Error::Xml(error.to_string()))?;
        let raw = attribute.key.as_ref();
        if raw.len() > MAX_STRING_BYTES || attribute.value.len() > MAX_STRING_BYTES {
            return Err(invalid("mail-merge XML attribute is too large"));
        }
        if raw == b"xmlns" || raw.starts_with(b"xmlns:") {
            continue;
        }
        if attributes.len() >= MAX_ATTRIBUTES_PER_NODE {
            return Err(invalid(
                "mail-merge XML has too many attributes on one node",
            ));
        }
        let (namespace, _) = resolver.resolve_attribute(attribute.key);
        attributes.push(Attribute {
            namespace: own_namespace(namespace)?,
            local: owned_name(attribute.key.local_name().as_ref(), "attribute name")?,
            value: attribute
                .decoded_and_normalized_value(XmlVersion::Explicit1_0, decoder)
                .map_err(|error| Error::Xml(error.to_string()))?
                .into_owned(),
        });
    }
    Ok(Node {
        namespace: own_namespace(namespace)?,
        local: owned_name(element.local_name().as_ref(), "element name")?,
        attributes,
        children: Vec::new(),
        has_text: false,
    })
}

fn own_namespace(namespace: ResolveResult<'_>) -> Result<OwnedNamespace> {
    match namespace {
        ResolveResult::Bound(Namespace(value)) => {
            Ok(OwnedNamespace::Bound(owned_name(value, "namespace name")?))
        },
        ResolveResult::Unbound => Ok(OwnedNamespace::Unbound),
        ResolveResult::Unknown(prefix) => Ok(OwnedNamespace::Unknown(owned_name(
            &prefix,
            "namespace prefix",
        )?)),
    }
}

fn owned_name(value: &[u8], description: &str) -> Result<String> {
    if value.len() > MAX_STRING_BYTES {
        return Err(invalid(format!("mail-merge {description} is too large")));
    }
    Ok(String::from_utf8_lossy(value).into_owned())
}

fn append_node(node: Node, stack: &mut [Node], root: &mut Option<Node>) -> Result<()> {
    if let Some(parent) = stack.last_mut() {
        parent.children.push(node);
    } else if root.replace(node).is_some() {
        return Err(invalid("mail-merge XML has multiple root elements"));
    }
    Ok(())
}

fn require_word_element(node: &Node, local: &str) -> Result<()> {
    if !node.is_word() || node.local != local {
        return Err(invalid(format!("expected WordprocessingML {local} root")));
    }
    Ok(())
}

fn ensure_leaf(node: &Node) -> Result<()> {
    if !node.children.is_empty() || node.has_text {
        return Err(invalid(format!(
            "mail-merge leaf '{}' cannot contain content",
            node.local
        )));
    }
    Ok(())
}

fn ensure_no_schema_attrs(node: &Node) -> Result<()> {
    for attribute in &node.attributes {
        if matches!(&attribute.namespace, OwnedNamespace::Bound(value) if value == W || value == STRICT_W)
            || matches!(
                attribute.namespace,
                OwnedNamespace::Unbound | OwnedNamespace::Unknown(_)
            )
        {
            return Err(invalid(format!(
                "unexpected attribute '{}' on {}",
                attribute.local, node.local
            )));
        }
    }
    if node.has_text {
        return Err(invalid(format!("{} cannot contain text", node.local)));
    }
    Ok(())
}

fn required_val(node: &Node) -> Result<String> {
    ensure_leaf(node)?;
    schema_attribute(node, "val", false)?
        .ok_or_else(|| invalid(format!("{} requires w:val", node.local)))
}

fn relationship_id(node: &Node) -> Result<String> {
    ensure_leaf(node)?;
    let value = schema_attribute(node, "id", true)?
        .ok_or_else(|| invalid(format!("{} requires r:id", node.local)))?;
    if value.is_empty() || value.len() > MAX_RELATIONSHIP_ID_BYTES {
        return Err(invalid(format!(
            "{} has invalid relationship id length",
            node.local
        )));
    }
    Ok(value)
}

fn schema_attribute(node: &Node, local: &str, relationship: bool) -> Result<Option<String>> {
    let mut result = None;
    for attribute in &node.attributes {
        let expected_namespace = if relationship {
            matches!(&attribute.namespace, OwnedNamespace::Bound(value) if value == R || value == STRICT_R)
        } else {
            matches!(&attribute.namespace, OwnedNamespace::Bound(value) if value == W || value == STRICT_W)
                || matches!(attribute.namespace, OwnedNamespace::Unbound)
        };
        if attribute.local == local {
            if !expected_namespace {
                return Err(invalid(format!(
                    "{} has spoofed {} attribute namespace",
                    node.local, local
                )));
            }
            if result.replace(attribute.value.clone()).is_some() {
                return Err(invalid(format!(
                    "{} has duplicate {} attribute",
                    node.local, local
                )));
            }
        } else if expected_namespace
            || matches!(
                attribute.namespace,
                OwnedNamespace::Unbound | OwnedNamespace::Unknown(_)
            )
        {
            return Err(invalid(format!(
                "unexpected attribute '{}' on {}",
                attribute.local, node.local
            )));
        }
    }
    Ok(result)
}

fn on_off(node: &Node) -> Result<bool> {
    ensure_leaf(node)?;
    match schema_attribute(node, "val", false)?.as_deref() {
        None | Some("true" | "1" | "on") => Ok(true),
        Some("false" | "0" | "off") => Ok(false),
        Some(value) => Err(invalid(format!("invalid on/off value '{value}'"))),
    }
}

fn decimal(node: &Node, description: &str) -> Result<i32> {
    required_val(node)?.parse::<i32>().map_err(|_| {
        invalid(format!(
            "{description} is outside the supported 32-bit bound"
        ))
    })
}

fn bounded_string(value: String, description: &str) -> Result<String> {
    if value.len() > MAX_STRING_BYTES {
        return Err(invalid(format!("{description} is too large")));
    }
    Ok(value)
}

fn strict_base64(value: &str) -> Result<Vec<u8>> {
    let compact: String = value
        .chars()
        .filter(|character| !character.is_ascii_whitespace())
        .collect();
    let decoded = BASE64
        .decode(compact.as_bytes())
        .map_err(|_| invalid("invalid recipient uniqueTag base64"))?;
    if decoded.is_empty()
        || decoded.len() > MAX_UNIQUE_TAG_BYTES
        || BASE64.encode(&decoded) != compact
    {
        return Err(invalid(
            "recipient uniqueTag base64 is empty, non-canonical, or too large",
        ));
    }
    Ok(decoded)
}
