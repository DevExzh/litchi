use super::{
    FORM, MAX_RAW, OFFICE, OdfForm, OdfFormControl, OdfFormNode, OdfFormPart, OdfFormProperty,
    OdfFormPropertyValue, OdfFormScalarValue, parse_form_parts,
};
use crate::odt::expanded_attributes;
use litchi_core::{Error, Result};
use quick_xml::events::{BytesStart, Event};
use quick_xml::reader::NsReader;
use std::ops::Range;

const MAX_PROPERTY_BYTES: usize = 64 * 1024;

/// A minimal standard form containing only typed custom properties.
#[derive(Debug, Clone, PartialEq)]
pub struct OdfPropertyForm {
    pub name: String,
    pub properties: Vec<OdfFormProperty>,
}

impl OdfPropertyForm {
    pub fn new(name: impl Into<String>) -> Self {
        Self {
            name: name.into(),
            properties: Vec::new(),
        }
    }

    pub fn add_property(&mut self, property: OdfFormProperty) -> Result<&mut Self> {
        property.to_xml_fragment()?;
        if self
            .properties
            .iter()
            .any(|existing| existing.name == property.name)
        {
            return invalid(format!("duplicate form property '{}'", property.name));
        }
        self.properties.push(property);
        Ok(self)
    }

    pub fn to_xml_fragment(&self) -> Result<String> {
        validate_string("form name", &self.name)?;
        let properties = properties_container_xml(&self.properties)?;
        Ok(format!(
            r#"<form:form form:name="{}">{properties}</form:form>"#,
            escape(&self.name)
        ))
    }
}

pub fn form_properties(xml: &str) -> Result<Vec<OdfFormProperty>> {
    let forms = parse_form_parts(&[(xml, OdfFormPart::Content)])?;
    let mut result = Vec::new();
    for group in forms.groups {
        for form in group.forms {
            collect_form_properties(form, &mut result);
        }
    }
    Ok(result)
}

fn collect_form_properties(form: OdfForm, result: &mut Vec<OdfFormProperty>) {
    result.extend(form.properties);
    for node in form.children {
        match node {
            OdfFormNode::Form(form) => collect_form_properties(form, result),
            OdfFormNode::Control(control) => collect_control_properties(control, result),
        }
    }
}

fn collect_control_properties(control: OdfFormControl, result: &mut Vec<OdfFormProperty>) {
    result.extend(control.properties);
    for node in control.children {
        if let OdfFormNode::Control(control) = node {
            collect_control_properties(control, result);
        }
    }
}

pub(crate) fn property_xml(property: &OdfFormProperty) -> Result<String> {
    validate_string("form property name", &property.name)?;
    let name = escape(&property.name);
    match &property.value {
        OdfFormPropertyValue::Scalar(value) => {
            let attributes = scalar_attributes(value, true)?;
            Ok(format!(
                r#"<form:property form:property-name="{name}"{attributes}/>"#
            ))
        },
        OdfFormPropertyValue::List { value_type, values } => {
            let value_type = value_type.as_deref().ok_or_else(|| {
                Error::InvalidFormat("list form property requires office:value-type".to_string())
            })?;
            validate_value_type(value_type)?;
            if values.is_empty() {
                return invalid("list form property requires at least one value");
            }
            let mut xml = format!(
                r#"<form:list-property form:property-name="{name}" office:value-type="{}">"#,
                escape(value_type)
            );
            for value in values {
                if scalar_type(value)? != value_type {
                    return invalid("form list property contains mixed value types");
                }
                xml.push_str("<form:list-value");
                xml.push_str(&scalar_attributes(value, false)?);
                xml.push_str("/>");
            }
            xml.push_str("</form:list-property>");
            Ok(xml)
        },
    }
}

fn properties_container_xml(properties: &[OdfFormProperty]) -> Result<String> {
    if properties.is_empty() {
        return Ok(String::new());
    }
    let mut names = std::collections::HashSet::new();
    let mut xml = String::from("<form:properties>");
    for property in properties {
        if !names.insert(property.name.as_str()) {
            return invalid(format!("duplicate form property '{}'", property.name));
        }
        xml.push_str(&property_xml(property)?);
    }
    xml.push_str("</form:properties>");
    Ok(xml)
}

fn scalar_type(value: &OdfFormScalarValue) -> Result<&str> {
    Ok(match value {
        OdfFormScalarValue::Boolean(_) => "boolean",
        OdfFormScalarValue::Number { value_type, .. } => {
            validate_value_type(value_type)?;
            value_type
        },
        OdfFormScalarValue::Text(_) => "string",
        OdfFormScalarValue::Date(_) => "date",
        OdfFormScalarValue::Time(_) => "time",
        OdfFormScalarValue::Void => "void",
        OdfFormScalarValue::Other { .. } => {
            return invalid("unsupported custom form property value type");
        },
    })
}

fn scalar_attributes(value: &OdfFormScalarValue, include_type: bool) -> Result<String> {
    let value_type = scalar_type(value)?;
    let mut result = if include_type {
        format!(r#" office:value-type="{value_type}""#)
    } else {
        String::new()
    };
    match value {
        OdfFormScalarValue::Boolean(value) => result.push_str(if *value {
            r#" office:boolean-value="true""#
        } else {
            r#" office:boolean-value="false""#
        }),
        OdfFormScalarValue::Number {
            value_type,
            lexical,
            currency,
        } => {
            let number = lexical
                .parse::<f64>()
                .map_err(|_| Error::InvalidFormat("invalid numeric form property".to_string()))?;
            if !number.is_finite() {
                return invalid("non-finite numeric form property");
            }
            result.push_str(&format!(r#" office:value="{}""#, escape(lexical)));
            if value_type == "currency" {
                let currency = currency.as_deref().ok_or_else(|| {
                    Error::InvalidFormat(
                        "currency form property requires office:currency".to_string(),
                    )
                })?;
                validate_string("form property currency", currency)?;
                result.push_str(&format!(r#" office:currency="{}""#, escape(currency)));
            } else if currency.is_some() {
                return invalid("office:currency is valid only for currency form properties");
            }
        },
        OdfFormScalarValue::Text(value) => {
            validate_string("form string property", value)?;
            result.push_str(&format!(r#" office:string-value="{}""#, escape(value)));
        },
        OdfFormScalarValue::Date(value) => {
            validate_temporal("date", value)?;
            result.push_str(&format!(r#" office:date-value="{}""#, escape(value)));
        },
        OdfFormScalarValue::Time(value) => {
            validate_temporal("time", value)?;
            result.push_str(&format!(r#" office:time-value="{}""#, escape(value)));
        },
        OdfFormScalarValue::Void => {},
        OdfFormScalarValue::Other { .. } => unreachable!(),
    }
    Ok(result)
}

fn validate_value_type(value: &str) -> Result<()> {
    if matches!(
        value,
        "boolean" | "float" | "percentage" | "currency" | "string" | "date" | "time" | "void"
    ) {
        Ok(())
    } else {
        invalid(format!("unsupported form property value type '{value}'"))
    }
}

fn validate_temporal(kind: &str, value: &str) -> Result<()> {
    validate_string("form temporal property", value)?;
    let plausible = if kind == "date" {
        value.len() >= 10
            && value.as_bytes().get(4) == Some(&b'-')
            && value.as_bytes().get(7) == Some(&b'-')
    } else {
        value.starts_with('P') || value.contains(':')
    };
    if plausible {
        Ok(())
    } else {
        invalid(format!("invalid {kind} form property"))
    }
}

pub fn insert_form_property_xml(
    xml: &str,
    owner_index: usize,
    property: &OdfFormProperty,
) -> Result<String> {
    let _ = form_properties(xml)?;
    let scan = scan(xml)?;
    let owner = scan.owners.get(owner_index).ok_or_else(|| {
        Error::InvalidFormat(format!(
            "form property owner {owner_index} is out of bounds"
        ))
    })?;
    if owner
        .property_names
        .iter()
        .any(|name| name == &property.name)
    {
        return invalid(format!("duplicate form property '{}'", property.name));
    }
    let fragment = bind_if_needed(xml, property_xml(property)?);
    if let Some(container_index) = owner.container {
        let container = &scan.containers[container_index];
        match &container.site {
            Site::Paired { close_start, .. } => {
                apply_edits(xml, vec![(*close_start..*close_start, fragment)])
            },
            Site::Empty { start, end, qname } => expand_empty(xml, *start, *end, qname, &fragment),
        }
    } else {
        let content = format!("<form:properties>{fragment}</form:properties>");
        match &owner.site {
            Site::Paired { open_end, .. } => apply_edits(
                xml,
                vec![(*open_end..*open_end, bind_container_if_needed(xml, content))],
            ),
            Site::Empty { start, end, qname } => expand_empty(
                xml,
                *start,
                *end,
                qname,
                &bind_container_if_needed(xml, content),
            ),
        }
    }
}

pub fn replace_form_property_xml(
    xml: &str,
    property_index: usize,
    replacement: &OdfFormProperty,
) -> Result<String> {
    let _ = form_properties(xml)?;
    let scan = scan(xml)?;
    let property = scan.properties.get(property_index).ok_or_else(|| {
        Error::InvalidFormat(format!("form property {property_index} is out of bounds"))
    })?;
    let owner = &scan.owners[property.owner];
    if owner
        .property_names
        .iter()
        .any(|name| name == &replacement.name && name != &property.name)
    {
        return invalid(format!("duplicate form property '{}'", replacement.name));
    }
    apply_edits(
        xml,
        vec![(
            property.span.clone(),
            bind_if_needed(xml, property_xml(replacement)?),
        )],
    )
}

pub fn remove_form_property_xml(xml: &str, property_index: usize) -> Result<String> {
    let _ = form_properties(xml)?;
    let scan = scan(xml)?;
    let property = scan.properties.get(property_index).ok_or_else(|| {
        Error::InvalidFormat(format!("form property {property_index} is out of bounds"))
    })?;
    let container = &scan.containers[property.container];
    let span = if container.count == 1 {
        container.full_span.clone()
    } else {
        property.span.clone()
    };
    apply_edits(xml, vec![(span, String::new())])
}

#[derive(Clone)]
enum Site {
    Paired {
        open_end: usize,
        close_start: usize,
    },
    Empty {
        start: usize,
        end: usize,
        qname: String,
    },
}
struct Owner {
    site: Site,
    container: Option<usize>,
    property_names: Vec<String>,
}
struct Container {
    site: Site,
    full_span: Range<usize>,
    count: usize,
}
struct PropertyLocation {
    span: Range<usize>,
    owner: usize,
    container: usize,
    name: String,
}
struct Scan {
    owners: Vec<Owner>,
    containers: Vec<Container>,
    properties: Vec<PropertyLocation>,
}
struct Open {
    local: Vec<u8>,
    start: usize,
    open_end: usize,
    owner: Option<usize>,
    container: Option<usize>,
    property: Option<usize>,
}

fn scan(xml: &str) -> Result<Scan> {
    if xml.len() > MAX_RAW {
        return invalid("form XML exceeds 64 MiB");
    }
    let mut reader = NsReader::from_str(xml);
    let mut buffer = Vec::new();
    let mut previous = 0usize;
    let mut stack = Vec::<Open>::new();
    let mut owner_stack = Vec::<usize>::new();
    let mut container_stack = Vec::<usize>::new();
    let mut owners = Vec::new();
    let mut containers = Vec::new();
    let mut properties = Vec::new();
    loop {
        let (namespace, event) = reader
            .read_resolved_event_into(&mut buffer)
            .map_err(|error| Error::InvalidFormat(format!("invalid form property XML: {error}")))?;
        let form = matches!(namespace, quick_xml::name::ResolveResult::Bound(ref uri) if uri.as_ref() == FORM.as_bytes());
        drop(namespace);
        let end = reader.buffer_position() as usize;
        match event {
            Event::Start(ref element) => {
                if stack.len() >= 128 {
                    return invalid("form property XML nesting exceeds 128 levels");
                }
                let local_name = element.local_name();
                let local = local_name.as_ref();
                let mut owner = None;
                let mut container = None;
                let mut property = None;
                if form
                    && (local == b"form"
                        || super::OdfFormControlKind::parse(std::str::from_utf8(local).map_err(
                            |_| Error::InvalidFormat("invalid form element name".to_string()),
                        )?)
                        .is_some())
                {
                    owner = Some(owners.len());
                    owners.push(Owner {
                        site: Site::Paired {
                            open_end: end,
                            close_start: 0,
                        },
                        container: None,
                        property_names: Vec::new(),
                    });
                    owner_stack.push(owner.unwrap());
                } else if form && local == b"properties" {
                    let owner_index = *owner_stack.last().ok_or_else(|| {
                        Error::InvalidFormat(
                            "form:properties has no form/control owner".to_string(),
                        )
                    })?;
                    if owners[owner_index].container.is_some() {
                        return invalid("duplicate form:properties container");
                    }
                    container = Some(containers.len());
                    owners[owner_index].container = container;
                    containers.push(Container {
                        site: Site::Paired {
                            open_end: end,
                            close_start: 0,
                        },
                        full_span: previous..0,
                        count: 0,
                    });
                    container_stack.push(container.unwrap());
                } else if form && matches!(local, b"property" | b"list-property") {
                    let owner_index = *owner_stack.last().ok_or_else(|| {
                        Error::InvalidFormat("form property has no owner".to_string())
                    })?;
                    let container_index = *container_stack.last().ok_or_else(|| {
                        Error::InvalidFormat("form property outside form:properties".to_string())
                    })?;
                    let name = validate_property_element(&reader, element, local)?;
                    if owners[owner_index]
                        .property_names
                        .iter()
                        .any(|existing| existing == &name)
                    {
                        return invalid(format!("duplicate form property '{name}'"));
                    }
                    owners[owner_index].property_names.push(name.clone());
                    containers[container_index].count += 1;
                    property = Some(properties.len());
                    properties.push(PropertyLocation {
                        span: previous..0,
                        owner: owner_index,
                        container: container_index,
                        name,
                    });
                } else if form && local == b"list-value" {
                    validate_list_value(&reader, element)?;
                }
                stack.push(Open {
                    local: local.to_vec(),
                    start: previous,
                    open_end: end,
                    owner,
                    container,
                    property,
                });
            },
            Event::Empty(ref element) => {
                let local_name = element.local_name();
                let local = local_name.as_ref();
                if form
                    && (local == b"form"
                        || super::OdfFormControlKind::parse(std::str::from_utf8(local).map_err(
                            |_| Error::InvalidFormat("invalid form element name".to_string()),
                        )?)
                        .is_some())
                {
                    owners.push(Owner {
                        site: Site::Empty {
                            start: previous,
                            end,
                            qname: qname(element)?,
                        },
                        container: None,
                        property_names: Vec::new(),
                    });
                } else if form && local == b"properties" {
                    return invalid("form:properties requires at least one property");
                } else if form && matches!(local, b"property" | b"list-property") {
                    let owner_index = *owner_stack.last().ok_or_else(|| {
                        Error::InvalidFormat("form property has no owner".to_string())
                    })?;
                    let container_index = *container_stack.last().ok_or_else(|| {
                        Error::InvalidFormat("form property outside form:properties".to_string())
                    })?;
                    let name = validate_property_element(&reader, element, local)?;
                    if local == b"list-property" {
                        return invalid("form:list-property requires list values");
                    }
                    if owners[owner_index]
                        .property_names
                        .iter()
                        .any(|existing| existing == &name)
                    {
                        return invalid(format!("duplicate form property '{name}'"));
                    }
                    owners[owner_index].property_names.push(name.clone());
                    containers[container_index].count += 1;
                    properties.push(PropertyLocation {
                        span: previous..end,
                        owner: owner_index,
                        container: container_index,
                        name,
                    });
                } else if form && local == b"list-value" {
                    validate_list_value(&reader, element)?;
                }
            },
            Event::End(ref element) => {
                let open = stack.pop().ok_or_else(|| {
                    Error::InvalidFormat("form property XML stack underflow".to_string())
                })?;
                if open.local.as_slice() != element.local_name().as_ref() {
                    return invalid("mismatched form property XML elements");
                }
                if let Some(index) = open.property {
                    properties[index].span.end = end;
                }
                if let Some(index) = open.container {
                    containers[index].site = Site::Paired {
                        open_end: open.open_end,
                        close_start: previous,
                    };
                    containers[index].full_span = open.start..end;
                    container_stack.pop();
                    if containers[index].count == 0 {
                        return invalid("form:properties requires at least one property");
                    }
                }
                if let Some(index) = open.owner {
                    owners[index].site = Site::Paired {
                        open_end: open.open_end,
                        close_start: previous,
                    };
                    owner_stack.pop();
                }
            },
            Event::DocType(_) => return invalid("DOCTYPE is not allowed in form property XML"),
            Event::Eof => break,
            _ => {},
        }
        previous = end;
        buffer.clear();
    }
    if !stack.is_empty() || !owner_stack.is_empty() || !container_stack.is_empty() {
        return invalid("incomplete form property XML");
    }
    Ok(Scan {
        owners,
        containers,
        properties,
    })
}

fn validate_property_element(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
    local: &[u8],
) -> Result<String> {
    let attributes = expanded_attributes(reader, element, "form property")?;
    let name = attributes
        .iter()
        .find(|attribute| {
            attribute.namespace_uri.as_deref() == Some(FORM)
                && attribute.local_name == "property-name"
        })
        .map(|attribute| attribute.value.clone())
        .ok_or_else(|| {
            Error::InvalidFormat("form property requires form:property-name".to_string())
        })?;
    validate_string("form property name", &name)?;
    let allowed = if local == b"list-property" {
        &[(FORM, "property-name"), (OFFICE, "value-type")][..]
    } else {
        &[
            (FORM, "property-name"),
            (OFFICE, "value-type"),
            (OFFICE, "boolean-value"),
            (OFFICE, "value"),
            (OFFICE, "currency"),
            (OFFICE, "string-value"),
            (OFFICE, "date-value"),
            (OFFICE, "time-value"),
        ][..]
    };
    if attributes.iter().any(|attribute| {
        !allowed.iter().any(|(namespace, name)| {
            attribute.namespace_uri.as_deref() == Some(*namespace) && attribute.local_name == *name
        })
    }) {
        return invalid("unexpected form property attribute");
    }
    Ok(name)
}

fn validate_list_value(reader: &NsReader<&[u8]>, element: &BytesStart<'_>) -> Result<()> {
    let attributes = expanded_attributes(reader, element, "form list value")?;
    if attributes.len() != 1
        || attributes[0].namespace_uri.as_deref() != Some(OFFICE)
        || !matches!(
            attributes[0].local_name.as_str(),
            "boolean-value" | "value" | "string-value" | "date-value" | "time-value"
        )
    {
        return invalid("invalid form:list-value attributes");
    }
    Ok(())
}

fn qname(element: &BytesStart<'_>) -> Result<String> {
    std::str::from_utf8(element.name().as_ref())
        .map(str::to_owned)
        .map_err(|_| Error::InvalidFormat("invalid form owner QName".to_string()))
}
fn bind_if_needed(xml: &str, fragment: String) -> String {
    let mut fragment = fragment;
    if !xml.contains(&format!(r#"xmlns:form="{FORM}""#)) {
        fragment = fragment.replacen(' ', &format!(r#" xmlns:form="{FORM}" "#), 1);
    }
    if !xml.contains(&format!(r#"xmlns:office="{OFFICE}""#)) {
        fragment = fragment.replacen(' ', &format!(r#" xmlns:office="{OFFICE}" "#), 1);
    }
    fragment
}
fn bind_container_if_needed(xml: &str, fragment: String) -> String {
    bind_if_needed(xml, fragment)
}
fn expand_empty(xml: &str, start: usize, end: usize, qname: &str, content: &str) -> Result<String> {
    let source = xml
        .get(start..end)
        .ok_or_else(|| Error::InvalidFormat("invalid empty form owner span".to_string()))?;
    let open = source
        .strip_suffix("/>")
        .ok_or_else(|| Error::InvalidFormat("empty form owner does not end with />".to_string()))?;
    apply_edits(
        xml,
        vec![(start..end, format!("{open}>{content}</{qname}>"))],
    )
}
fn apply_edits(xml: &str, mut edits: Vec<(Range<usize>, String)>) -> Result<String> {
    edits.sort_by(|a, b| b.0.start.cmp(&a.0.start));
    let mut output = xml.to_string();
    let mut previous = xml.len();
    for (span, replacement) in edits {
        if span.start > span.end || span.end > previous || span.end > output.len() {
            return invalid("invalid overlapping form property mutation spans");
        }
        output.replace_range(span.clone(), &replacement);
        previous = span.start;
    }
    Ok(output)
}
fn validate_string(label: &str, value: &str) -> Result<()> {
    if value.len() > MAX_PROPERTY_BYTES {
        return invalid(format!("{label} exceeds 64 KiB"));
    }
    if value.chars().any(|c| !matches!(c, '\u{9}' | '\u{A}' | '\u{D}' | '\u{20}'..='\u{D7FF}' | '\u{E000}'..='\u{FFFD}' | '\u{10000}'..='\u{10FFFF}')) { return invalid(format!("{label} contains a forbidden XML character")); }
    Ok(())
}
fn escape(value: &str) -> String {
    let mut output = String::new();
    for c in value.chars() {
        match c {
            '&' => output.push_str("&amp;"),
            '<' => output.push_str("&lt;"),
            '"' => output.push_str("&quot;"),
            '\r' => output.push_str("&#13;"),
            _ => output.push(c),
        }
    }
    output
}
fn invalid<T>(message: impl Into<String>) -> Result<T> {
    Err(Error::InvalidFormat(message.into()))
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn canonical_scalar_and_list_properties() {
        assert_eq!(
            OdfFormProperty::boolean("Enabled", true)
                .to_xml_fragment()
                .unwrap(),
            r#"<form:property form:property-name="Enabled" office:value-type="boolean" office:boolean-value="true"/>"#
        );
        assert_eq!(
            OdfFormProperty::text("Label", "A<&\"")
                .to_xml_fragment()
                .unwrap(),
            r#"<form:property form:property-name="Label" office:value-type="string" office:string-value="A&lt;&amp;&quot;"/>"#
        );
        let list = OdfFormProperty::list(
            "Choices",
            "string",
            vec![
                OdfFormScalarValue::Text("A".into()),
                OdfFormScalarValue::Text("B".into()),
            ],
        );
        assert_eq!(
            list.to_xml_fragment().unwrap(),
            r#"<form:list-property form:property-name="Choices" office:value-type="string"><form:list-value office:string-value="A"/><form:list-value office:string-value="B"/></form:list-property>"#
        );
    }

    #[test]
    fn lossless_insert_replace_remove_and_empty_owner_expansion() {
        let source = format!(
            r#"<o:document-content xmlns:o="{OFFICE}" xmlns:f="{FORM}" xmlns:office="{OFFICE}" xmlns:form="{FORM}"><o:body><o:text><o:forms><f:form f:name="F"><!--keep--><f:properties><f:property f:property-name="Old" o:value-type="string" o:string-value="x"/></f:properties><f:text f:name="C"/></f:form></o:forms></o:text></o:body></o:document-content>"#
        );
        let inserted =
            insert_form_property_xml(&source, 1, &OdfFormProperty::boolean("Enabled", true))
                .unwrap();
        assert!(inserted.contains(r#"<f:text f:name="C"><form:properties><form:property form:property-name="Enabled" office:value-type="boolean" office:boolean-value="true"/></form:properties></f:text>"#));
        let replaced =
            replace_form_property_xml(&inserted, 0, &OdfFormProperty::text("New", "y")).unwrap();
        assert!(replaced.contains("<!--keep--><f:properties>"));
        let removed = remove_form_property_xml(&replaced, 0).unwrap();
        assert!(!removed.contains("<f:properties>"));
        assert!(removed.contains("<!--keep-->"));
    }

    #[test]
    fn hostile_properties_are_rejected() {
        for property in [
            OdfFormProperty::number("N", "float", "NaN", None),
            OdfFormProperty::number("N", "currency", "1", None),
            OdfFormProperty {
                name: "X".into(),
                value: OdfFormPropertyValue::Scalar(OdfFormScalarValue::Other {
                    value_type: "object".into(),
                    lexical: None,
                }),
            },
            OdfFormProperty::list("L", "string", Vec::new()),
        ] {
            assert!(property.to_xml_fragment().is_err());
        }
        let hostile = format!(
            r#"<o:forms xmlns:o="{OFFICE}" xmlns:f="{FORM}" xmlns:u="urn:hostile"><f:form><f:properties><f:property f:property-name="X" o:value-type="string" o:string-value="x" u:extra="1"/></f:properties></f:form></o:forms>"#
        );
        assert!(insert_form_property_xml(&hostile, 0, &OdfFormProperty::text("Y", "y")).is_err());
    }

    #[test]
    fn libreoffice_odfpy_and_odfdo_properties_round_trip() {
        let libreoffice = include_str!(
            "../../../../test-data/libreoffice-core/vcl/qa/cppunit/pdfexport/data/PDF_export_with_formcontrol.fodt"
        );
        let parsed = form_properties(libreoffice).unwrap();
        assert!(parsed.len() >= 6);
        let updated = replace_form_property_xml(
            libreoffice,
            0,
            &OdfFormProperty::boolean("PropertyChangeNotificationEnabled", false),
        )
        .unwrap();
        assert!(updated.contains(r#"office:boolean-value="false""#));
        assert_eq!(form_properties(&updated).unwrap().len(), parsed.len());
        let producer = format!(
            r#"<o:document-content xmlns:o="{OFFICE}" xmlns:f="{FORM}"><o:body><o:text><o:forms><f:form f:name="P"><f:properties><f:list-property f:property-name="Values" o:value-type="string"><f:list-value o:string-value="one"/><f:list-value o:string-value="two"/></f:list-property></f:properties></f:form></o:forms></o:text></o:body></o:document-content>"#
        );
        assert_eq!(form_properties(&producer).unwrap().len(), 1);
    }
}
