//! Strict, inert support for form-owned `form:connection-resource` elements.

use litchi_core::{Error, Result};
use quick_xml::XmlVersion;
use quick_xml::events::{BytesStart, Event};
use quick_xml::name::ResolveResult;
use quick_xml::reader::NsReader;

const FORM_NAMESPACE: &str = "urn:oasis:names:tc:opendocument:xmlns:form:1.0";
const OFFICE_NAMESPACE: &str = "urn:oasis:names:tc:opendocument:xmlns:office:1.0";
const SCRIPT_NAMESPACE: &str = "urn:oasis:names:tc:opendocument:xmlns:script:1.0";
const TEXT_NAMESPACE: &str = "urn:oasis:names:tc:opendocument:xmlns:text:1.0";
const XLINK_NAMESPACE: &str = "http://www.w3.org/1999/xlink";
const XML_NAMESPACE: &str = "http://www.w3.org/XML/1998/namespace";

const MAX_XML_SIZE: usize = 64 * 1024 * 1024;
const MAX_DEPTH: usize = 128;
const MAX_FORMS: usize = 4096;
const MAX_RESOURCES: usize = 4096;
const MAX_REFERENCE_SIZE: usize = 64 * 1024;
const MAX_AGGREGATE_SIZE: usize = 16 * 1024 * 1024;

/// An inert form connection URI. The value is never opened or resolved.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct OdfFormConnectionResource {
    pub href: String,
}

impl OdfFormConnectionResource {
    pub fn new(href: impl Into<String>) -> Result<Self> {
        let resource = Self { href: href.into() };
        validate_resource(&resource)?;
        Ok(resource)
    }

    pub fn to_xml_fragment(&self) -> Result<String> {
        validate_resource(self)?;
        let mut xml = String::from("<form:connection-resource xlink:href=\"");
        push_xml_attribute(&mut xml, &self.href);
        xml.push_str("\"/>");
        Ok(xml)
    }
}

/// A form builder whose connection resource is its final child.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct OdfConnectionResourceForm {
    pub name: String,
    pub xml_id: Option<String>,
    pub resource: OdfFormConnectionResource,
}

impl OdfConnectionResourceForm {
    pub fn new(name: impl Into<String>, resource: OdfFormConnectionResource) -> Self {
        Self {
            name: name.into(),
            xml_id: None,
            resource,
        }
    }

    pub fn to_xml_fragment(&self) -> Result<String> {
        validate_form(self)?;
        let mut xml = String::from("<form:form form:name=\"");
        push_xml_attribute(&mut xml, &self.name);
        xml.push('"');
        if let Some(xml_id) = self.xml_id.as_deref() {
            xml.push_str(" xml:id=\"");
            push_xml_attribute(&mut xml, xml_id);
            xml.push('"');
        }
        xml.push('>');
        xml.push_str(&self.resource.to_xml_fragment()?);
        xml.push_str("</form:form>");
        Ok(xml)
    }
}

/// A resource together with the direct `form:form` that owns it.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct OdfOwnedFormConnectionResource {
    pub form_index: usize,
    pub form_name: Option<String>,
    pub form_xml_id: Option<String>,
    pub resource: OdfFormConnectionResource,
}

/// Parses direct, form-owned connection resources in document order.
pub fn form_connection_resources(xml: &str) -> Result<Vec<OdfOwnedFormConnectionResource>> {
    Ok(scan_document(xml)?
        .resources
        .into_iter()
        .map(|entry| entry.value)
        .collect())
}

/// Inserts a resource as the final child of the indexed `form:form`.
pub fn insert_form_connection_resource_xml(
    xml: &str,
    form_index: usize,
    resource: &OdfFormConnectionResource,
) -> Result<String> {
    let scan = scan_document(xml)?;
    let form = scan
        .forms
        .get(form_index)
        .ok_or_else(|| Error::InvalidFormat(format!("form index {form_index} is out of bounds")))?;
    if form.has_resource {
        return Err(Error::InvalidFormat(
            "form already has a form:connection-resource".to_string(),
        ));
    }
    let fragment = resource_mutation_fragment(resource)?;
    if form.empty {
        let source = &xml[form.start..form.end];
        let close = source
            .rfind("/>")
            .ok_or_else(|| Error::InvalidFormat("invalid empty form serialization".to_string()))?;
        let mut replacement = String::with_capacity(source.len() + fragment.len() + 16);
        replacement.push_str(&source[..close]);
        replacement.push('>');
        replacement.push_str(&fragment);
        replacement.push_str("</");
        replacement.push_str(&form.qname);
        replacement.push('>');
        return replace_span(xml, form.start, form.end, &replacement);
    }
    insert_at(xml, form.end_start, &fragment)
}

/// Replaces the indexed form-owned connection resource without rebuilding its form.
pub fn replace_form_connection_resource_xml(
    xml: &str,
    resource_index: usize,
    resource: &OdfFormConnectionResource,
) -> Result<String> {
    let scan = scan_document(xml)?;
    let old = scan.resources.get(resource_index).ok_or_else(|| {
        Error::InvalidFormat(format!(
            "connection-resource index {resource_index} is out of bounds"
        ))
    })?;
    replace_span(
        xml,
        old.start,
        old.end,
        &resource_mutation_fragment(resource)?,
    )
}

/// Removes the indexed form-owned connection resource.
pub fn remove_form_connection_resource_xml(xml: &str, resource_index: usize) -> Result<String> {
    let scan = scan_document(xml)?;
    let old = scan.resources.get(resource_index).ok_or_else(|| {
        Error::InvalidFormat(format!(
            "connection-resource index {resource_index} is out of bounds"
        ))
    })?;
    replace_span(xml, old.start, old.end, "")
}

#[derive(Debug, Clone)]
struct Node {
    namespace: Option<String>,
    local: String,
}

#[derive(Debug)]
struct FormSpan {
    name: Option<String>,
    xml_id: Option<String>,
    depth: usize,
    start: usize,
    end_start: usize,
    end: usize,
    empty: bool,
    qname: String,
    has_resource: bool,
}

#[derive(Debug)]
struct ResourceSpan {
    value: OdfOwnedFormConnectionResource,
    start: usize,
    end: usize,
}

#[derive(Debug, Default)]
struct ScanResult {
    forms: Vec<FormSpan>,
    resources: Vec<ResourceSpan>,
}

#[derive(Debug)]
struct ActiveResource {
    depth: usize,
    result_index: Option<usize>,
}

fn scan_document(xml: &str) -> Result<ScanResult> {
    if xml.len() > MAX_XML_SIZE {
        return Err(Error::InvalidFormat(format!(
            "ODF XML exceeds the supported {MAX_XML_SIZE}-byte limit"
        )));
    }
    let mut reader = NsReader::from_reader(xml.as_bytes());
    reader.config_mut().trim_text(false);
    let mut buffer = Vec::new();
    let mut stack: Vec<Node> = Vec::new();
    let mut active_forms: Vec<usize> = Vec::new();
    let mut active_resource: Option<ActiveResource> = None;
    let mut result = ScanResult::default();
    let mut aggregate = 0usize;

    loop {
        let start = reader.buffer_position() as usize;
        let (resolution, event) =
            reader
                .read_resolved_event_into(&mut buffer)
                .map_err(|error| {
                    Error::InvalidFormat(format!("invalid form connection-resource XML: {error}"))
                })?;
        let namespace = resolved_namespace(resolution)?;
        let end = reader.buffer_position() as usize;
        match event {
            Event::Start(ref element) | Event::Empty(ref element) => {
                if stack.len() >= MAX_DEPTH {
                    return Err(Error::InvalidFormat(format!(
                        "ODF XML exceeds the supported depth of {MAX_DEPTH}"
                    )));
                }
                let empty = matches!(event, Event::Empty(_));
                let local = utf8(element.local_name().as_ref(), "element name")?.to_string();
                let node = Node {
                    namespace: namespace.clone(),
                    local: local.clone(),
                };

                if is_active_element(namespace.as_deref(), &local) {
                    return Err(Error::InvalidFormat(
                        "active event behavior is not accepted in inert form resources".to_string(),
                    ));
                }

                if namespace.as_deref() == Some(FORM_NAMESPACE) && local == "form" {
                    validate_form_parent(stack.last())?;
                    if result.forms.len() >= MAX_FORMS {
                        return Err(Error::InvalidFormat(format!(
                            "document exceeds {MAX_FORMS} forms"
                        )));
                    }
                    if let Some(parent_index) =
                        direct_form_parent(&active_forms, &result.forms, stack.len())
                    {
                        ensure_no_child_after_resource(&result.forms[parent_index])?;
                    }
                    let (name, xml_id) = parse_form_identity(&reader, element, &mut aggregate)?;
                    let qname = utf8(element.name().as_ref(), "form qualified name")?.to_string();
                    let index = result.forms.len();
                    result.forms.push(FormSpan {
                        name,
                        xml_id,
                        depth: stack.len(),
                        start,
                        end_start: end,
                        end,
                        empty,
                        qname,
                        has_resource: false,
                    });
                    if !empty {
                        active_forms.push(index);
                    }
                } else if namespace.as_deref() == Some(FORM_NAMESPACE)
                    && local == "connection-resource"
                {
                    if active_resource.is_some() {
                        return Err(Error::InvalidFormat(
                            "form:connection-resource must be empty".to_string(),
                        ));
                    }
                    let resource = parse_resource(&reader, element, &mut aggregate)?;
                    let owner = direct_form_parent(&active_forms, &result.forms, stack.len());
                    let database_owner = stack.last().is_some_and(is_database_field);
                    if owner.is_none() && !database_owner {
                        return Err(Error::InvalidFormat(
                            "form:connection-resource must be owned by form:form or an ODF database field"
                                .to_string(),
                        ));
                    }
                    let result_index = if let Some(form_index) = owner {
                        if result.forms[form_index].has_resource {
                            return Err(Error::InvalidFormat(
                                "form:form may contain at most one form:connection-resource"
                                    .to_string(),
                            ));
                        }
                        result.forms[form_index].has_resource = true;
                        if result.resources.len() >= MAX_RESOURCES {
                            return Err(Error::InvalidFormat(format!(
                                "document exceeds {MAX_RESOURCES} form connection resources"
                            )));
                        }
                        let entry_index = result.resources.len();
                        result.resources.push(ResourceSpan {
                            value: OdfOwnedFormConnectionResource {
                                form_index,
                                form_name: result.forms[form_index].name.clone(),
                                form_xml_id: result.forms[form_index].xml_id.clone(),
                                resource,
                            },
                            start,
                            end,
                        });
                        Some(entry_index)
                    } else {
                        None
                    };
                    if !empty {
                        active_resource = Some(ActiveResource {
                            depth: stack.len(),
                            result_index,
                        });
                    }
                } else if let Some(form_index) =
                    direct_form_parent(&active_forms, &result.forms, stack.len())
                {
                    ensure_no_child_after_resource(&result.forms[form_index])?;
                }

                if !empty {
                    stack.push(node);
                }
            },
            Event::End(_) => {
                let depth = stack.len().checked_sub(1).ok_or_else(|| {
                    Error::InvalidFormat("unexpected XML end element".to_string())
                })?;
                if active_resource
                    .as_ref()
                    .is_some_and(|active| active.depth == depth)
                    && let Some(index) = active_resource
                        .take()
                        .and_then(|active| active.result_index)
                {
                    result.resources[index].end = end;
                }
                if let Some(index) =
                    active_forms.pop_if(|index| result.forms[*index].depth == depth)
                {
                    result.forms[index].end_start = start;
                    result.forms[index].end = end;
                }
                stack.pop();
            },
            Event::Text(ref text) => {
                let value = text.decode().map_err(|error| {
                    Error::InvalidFormat(format!("invalid form resource text: {error}"))
                })?;
                if active_resource.is_some() && !value.is_empty() {
                    return Err(Error::InvalidFormat(
                        "form:connection-resource must be empty".to_string(),
                    ));
                }
                if let Some(form_index) =
                    direct_form_parent(&active_forms, &result.forms, stack.len())
                    && result.forms[form_index].has_resource
                    && !value.trim().is_empty()
                {
                    return Err(Error::InvalidFormat(
                        "form:connection-resource must be the final form child".to_string(),
                    ));
                }
            },
            Event::CData(ref data) => {
                if active_resource.is_some() || !data.is_empty() {
                    return Err(Error::InvalidFormat(
                        "CDATA is not accepted in inert form resources".to_string(),
                    ));
                }
            },
            Event::DocType(_) => {
                return Err(Error::InvalidFormat(
                    "DOCTYPE is not accepted in ODF form XML".to_string(),
                ));
            },
            Event::PI(_) | Event::GeneralRef(_) => {
                return Err(Error::InvalidFormat(
                    "processing instructions and entity references are not accepted in inert form resources"
                        .to_string(),
                ));
            },
            Event::Eof => break,
            Event::Decl(_) | Event::Comment(_) => {},
        }
        buffer.clear();
    }
    if !stack.is_empty() || active_resource.is_some() || !active_forms.is_empty() {
        return Err(Error::InvalidFormat(
            "unclosed ODF form element".to_string(),
        ));
    }
    Ok(result)
}

fn direct_form_parent(
    active_forms: &[usize],
    forms: &[FormSpan],
    child_depth: usize,
) -> Option<usize> {
    active_forms
        .last()
        .copied()
        .filter(|index| forms[*index].depth + 1 == child_depth)
}

fn ensure_no_child_after_resource(form: &FormSpan) -> Result<()> {
    if form.has_resource {
        return Err(Error::InvalidFormat(
            "form:connection-resource must be the final form child".to_string(),
        ));
    }
    Ok(())
}

fn validate_form_parent(parent: Option<&Node>) -> Result<()> {
    let valid = parent.is_some_and(|parent| {
        (parent.namespace.as_deref() == Some(OFFICE_NAMESPACE) && parent.local == "forms")
            || (parent.namespace.as_deref() == Some(FORM_NAMESPACE) && parent.local == "form")
    });
    if !valid {
        return Err(Error::InvalidFormat(
            "form:form must be owned by office:forms or another form:form".to_string(),
        ));
    }
    Ok(())
}

fn is_database_field(node: &Node) -> bool {
    node.namespace.as_deref() == Some(TEXT_NAMESPACE)
        && matches!(
            node.local.as_str(),
            "database-display"
                | "database-name"
                | "database-next"
                | "database-row-number"
                | "database-row-select"
        )
}

fn is_active_element(namespace: Option<&str>, local: &str) -> bool {
    (namespace == Some(OFFICE_NAMESPACE) && local == "event-listeners")
        || (namespace == Some(SCRIPT_NAMESPACE) && matches!(local, "event-listener" | "listener"))
}

fn parse_form_identity(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
    aggregate: &mut usize,
) -> Result<(Option<String>, Option<String>)> {
    let mut name = None;
    let mut xml_id = None;
    for attribute in element.attributes() {
        let attribute = attribute
            .map_err(|error| Error::InvalidFormat(format!("invalid form attribute: {error}")))?;
        if is_namespace_declaration(attribute.key.as_ref()) {
            continue;
        }
        let (resolution, local) = reader.resolver().resolve_attribute(attribute.key);
        let namespace = resolved_namespace(resolution)?;
        let local = utf8(local.as_ref(), "form attribute name")?;
        let target = if namespace.as_deref() == Some(FORM_NAMESPACE) && local == "name" {
            Some(&mut name)
        } else if namespace.as_deref() == Some(XML_NAMESPACE) && local == "id" {
            Some(&mut xml_id)
        } else {
            None
        };
        if let Some(target) = target {
            if target.is_some() {
                return Err(Error::InvalidFormat(format!(
                    "duplicate form attribute {local}"
                )));
            }
            let value = attribute
                .decoded_and_normalized_value(XmlVersion::Implicit1_0, reader.decoder())
                .map_err(|error| Error::InvalidFormat(format!("invalid form attribute: {error}")))?
                .into_owned();
            append_aggregate(aggregate, value.len())?;
            *target = Some(value);
        }
    }
    Ok((name, xml_id))
}

fn parse_resource(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
    aggregate: &mut usize,
) -> Result<OdfFormConnectionResource> {
    let mut href = None;
    for attribute in element.attributes() {
        let attribute = attribute.map_err(|error| {
            Error::InvalidFormat(format!("invalid connection-resource attribute: {error}"))
        })?;
        if is_namespace_declaration(attribute.key.as_ref()) {
            continue;
        }
        let (resolution, local) = reader.resolver().resolve_attribute(attribute.key);
        let namespace = resolved_namespace(resolution)?;
        let local = utf8(local.as_ref(), "connection-resource attribute name")?;
        if namespace.as_deref() != Some(XLINK_NAMESPACE) || local != "href" {
            return Err(Error::InvalidFormat(format!(
                "unsupported form:connection-resource attribute {local}"
            )));
        }
        if href.is_some() {
            return Err(Error::InvalidFormat(
                "duplicate xlink:href on form:connection-resource".to_string(),
            ));
        }
        let value = attribute
            .decoded_and_normalized_value(XmlVersion::Implicit1_0, reader.decoder())
            .map_err(|error| {
                Error::InvalidFormat(format!("invalid connection-resource href: {error}"))
            })?
            .into_owned();
        append_aggregate(aggregate, value.len())?;
        href = Some(value);
    }
    OdfFormConnectionResource::new(href.ok_or_else(|| {
        Error::InvalidFormat("form:connection-resource requires xlink:href".to_string())
    })?)
}

fn resolved_namespace(result: ResolveResult<'_>) -> Result<Option<String>> {
    match result {
        ResolveResult::Bound(namespace) => {
            Ok(Some(utf8(namespace.as_ref(), "namespace URI")?.to_string()))
        },
        ResolveResult::Unbound => Ok(None),
        ResolveResult::Unknown(prefix) => Err(Error::InvalidFormat(format!(
            "unbound XML namespace prefix {}",
            utf8(prefix.as_ref(), "namespace prefix")?
        ))),
    }
}

fn validate_resource(resource: &OdfFormConnectionResource) -> Result<()> {
    if resource.href.len() > MAX_REFERENCE_SIZE {
        return Err(Error::InvalidFormat(format!(
            "form connection URI exceeds {MAX_REFERENCE_SIZE} bytes"
        )));
    }
    validate_xml_chars(&resource.href, "form connection URI")
}

fn validate_form(form: &OdfConnectionResourceForm) -> Result<()> {
    validate_xml_chars(&form.name, "form name")?;
    if form.name.len() > MAX_REFERENCE_SIZE {
        return Err(Error::InvalidFormat(
            "form name exceeds the supported limit".to_string(),
        ));
    }
    if let Some(xml_id) = form.xml_id.as_deref() {
        validate_xml_id(xml_id)?;
    }
    validate_resource(&form.resource)
}

fn validate_xml_id(value: &str) -> Result<()> {
    validate_xml_chars(value, "xml:id")?;
    let mut chars = value.chars();
    let Some(first) = chars.next() else {
        return Err(Error::InvalidFormat("xml:id must not be empty".to_string()));
    };
    if !(first == '_' || first.is_alphabetic())
        || chars.any(|ch| !(ch == '_' || ch == '-' || ch == '.' || ch.is_alphanumeric()))
    {
        return Err(Error::InvalidFormat(
            "xml:id must be an XML NCName".to_string(),
        ));
    }
    Ok(())
}

fn validate_xml_chars(value: &str, context: &str) -> Result<()> {
    if value.chars().all(is_xml_1_0_char) {
        Ok(())
    } else {
        Err(Error::InvalidFormat(format!(
            "{context} contains forbidden XML characters"
        )))
    }
}

fn is_xml_1_0_char(ch: char) -> bool {
    matches!(ch, '\u{9}' | '\u{A}' | '\u{D}')
        || ('\u{20}'..='\u{D7FF}').contains(&ch)
        || ('\u{E000}'..='\u{FFFD}').contains(&ch)
        || ('\u{10000}'..='\u{10FFFF}').contains(&ch)
}

fn append_aggregate(aggregate: &mut usize, size: usize) -> Result<()> {
    *aggregate = aggregate
        .checked_add(size)
        .ok_or_else(|| Error::InvalidFormat("form resource aggregate size overflow".to_string()))?;
    if *aggregate > MAX_AGGREGATE_SIZE {
        return Err(Error::InvalidFormat(format!(
            "form resources exceed {MAX_AGGREGATE_SIZE} aggregate bytes"
        )));
    }
    Ok(())
}

fn is_namespace_declaration(name: &[u8]) -> bool {
    name == b"xmlns" || name.starts_with(b"xmlns:")
}

fn utf8<'a>(value: &'a [u8], context: &str) -> Result<&'a str> {
    std::str::from_utf8(value)
        .map_err(|error| Error::InvalidFormat(format!("invalid UTF-8 in {context}: {error}")))
}

fn push_xml_attribute(xml: &mut String, value: &str) {
    for ch in value.chars() {
        match ch {
            '&' => xml.push_str("&amp;"),
            '<' => xml.push_str("&lt;"),
            '"' => xml.push_str("&quot;"),
            '\r' => xml.push_str("&#13;"),
            '\n' => xml.push_str("&#10;"),
            '\t' => xml.push_str("&#9;"),
            _ => xml.push(ch),
        }
    }
}

fn resource_mutation_fragment(resource: &OdfFormConnectionResource) -> Result<String> {
    validate_resource(resource)?;
    let mut xml = format!(
        "<form:connection-resource xmlns:form=\"{FORM_NAMESPACE}\" xmlns:xlink=\"{XLINK_NAMESPACE}\" xlink:href=\""
    );
    push_xml_attribute(&mut xml, &resource.href);
    xml.push_str("\"/>");
    Ok(xml)
}

fn insert_at(xml: &str, offset: usize, fragment: &str) -> Result<String> {
    if !xml.is_char_boundary(offset) {
        return Err(Error::InvalidFormat(
            "XML mutation offset is not UTF-8 aligned".to_string(),
        ));
    }
    let mut output = String::with_capacity(xml.len() + fragment.len());
    output.push_str(&xml[..offset]);
    output.push_str(fragment);
    output.push_str(&xml[offset..]);
    Ok(output)
}

fn replace_span(xml: &str, start: usize, end: usize, replacement: &str) -> Result<String> {
    if start > end || end > xml.len() || !xml.is_char_boundary(start) || !xml.is_char_boundary(end)
    {
        return Err(Error::InvalidFormat(
            "invalid XML mutation span".to_string(),
        ));
    }
    let mut output = String::with_capacity(xml.len() - (end - start) + replacement.len());
    output.push_str(&xml[..start]);
    output.push_str(replacement);
    output.push_str(&xml[end..]);
    Ok(output)
}

#[cfg(test)]
mod tests {
    use super::*;

    const PREFIX: &str = r#"<o:document-content
        xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
        xmlns:f="urn:oasis:names:tc:opendocument:xmlns:form:1.0"
        xmlns:x="http://www.w3.org/1999/xlink"
        xmlns:t="urn:oasis:names:tc:opendocument:xmlns:text:1.0"><o:body><o:text><o:forms>"#;
    const SUFFIX: &str = "</o:forms><t:p/></o:text></o:body></o:document-content>";

    #[test]
    fn parses_nested_owned_resources_and_keeps_database_resources_separate() {
        let xml = format!(
            r#"{PREFIX}<f:form f:name="outer"><f:form f:name="inner" xml:id="i1"><f:connection-resource x:href="sdbc:inner&amp;x"/></f:form><f:connection-resource x:href="sdbc:outer"/></f:form>{SUFFIX}"#
        );
        let resources = form_connection_resources(&xml).unwrap();
        assert_eq!(resources.len(), 2);
        assert_eq!(resources[0].form_name.as_deref(), Some("inner"));
        assert_eq!(resources[0].form_xml_id.as_deref(), Some("i1"));
        assert_eq!(resources[0].resource.href, "sdbc:inner&x");
        assert_eq!(resources[1].form_name.as_deref(), Some("outer"));

        let database = format!(
            r#"<o:document-content xmlns:o="{OFFICE_NAMESPACE}" xmlns:f="{FORM_NAMESPACE}" xmlns:x="{XLINK_NAMESPACE}" xmlns:t="{TEXT_NAMESPACE}"><o:body><o:text><t:p><t:database-name t:table-name="t"><f:connection-resource x:href="db"/></t:database-name></t:p></o:text></o:body></o:document-content>"#
        );
        assert!(form_connection_resources(&database).unwrap().is_empty());
    }

    #[test]
    fn rejects_invalid_ownership_cardinality_order_attributes_and_active_content() {
        let bodies = [
            r#"<f:form f:name="a"><f:connection-resource/></f:form>"#,
            r#"<f:form f:name="a"><f:connection-resource x:href="db" x:type="simple"/></f:form>"#,
            r#"<f:form f:name="a"><f:connection-resource x:href="db">x</f:connection-resource></f:form>"#,
            r#"<f:form f:name="a"><f:connection-resource x:href="a"/><f:connection-resource x:href="b"/></f:form>"#,
            r#"<f:form f:name="a"><f:connection-resource x:href="a"/><f:form f:name="late"/></f:form>"#,
            r#"<f:form f:name="a"><o:event-listeners/><f:connection-resource x:href="db"/></f:form>"#,
            r#"<f:connection-resource x:href="db"/>"#,
        ];
        for body in bodies {
            let xml = format!("{PREFIX}{body}{SUFFIX}");
            assert!(form_connection_resources(&xml).is_err(), "accepted {body}");
        }
        let doctype =
            format!("<!DOCTYPE x [<!ENTITY e 'x'>]>{PREFIX}<f:form f:name=\"a\"/>{SUFFIX}");
        assert!(form_connection_resources(&doctype).is_err());
    }

    #[test]
    fn canonical_writer_and_lossless_mutations_cover_empty_and_nested_forms() {
        let resource = OdfFormConnectionResource::new("sdbc:a&b").unwrap();
        assert_eq!(
            resource.to_xml_fragment().unwrap(),
            r#"<form:connection-resource xlink:href="sdbc:a&amp;b"/>"#
        );
        let xml = format!(
            r#"{PREFIX}<f:form f:name="keep" data-v="1"/><f:form f:name="other"></f:form>{SUFFIX}"#
        );
        let inserted = insert_form_connection_resource_xml(&xml, 0, &resource).unwrap();
        assert!(inserted.contains(r#"data-v="1"><form:connection-resource"#));
        assert_eq!(form_connection_resources(&inserted).unwrap().len(), 1);
        let replacement = OdfFormConnectionResource::new("sdbc:new").unwrap();
        let replaced = replace_form_connection_resource_xml(&inserted, 0, &replacement).unwrap();
        assert!(replaced.contains("sdbc:new"));
        assert!(replaced.contains(r#"<f:form f:name="other"></f:form>"#));
        let removed = remove_form_connection_resource_xml(&replaced, 0).unwrap();
        assert!(form_connection_resources(&removed).unwrap().is_empty());

        let mut form = OdfConnectionResourceForm::new("source", replacement);
        form.xml_id = Some("source_1".into());
        assert!(
            form.to_xml_fragment()
                .unwrap()
                .ends_with(r#"<form:connection-resource xlink:href="sdbc:new"/></form:form>"#)
        );
    }
}
