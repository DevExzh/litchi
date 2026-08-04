//! Typed writing and lossless mutation for `form:value-range` controls.

use std::ops::Range;

use litchi_core::{Error, Result};
use quick_xml::NsReader;
use quick_xml::events::{BytesStart, Event};
use quick_xml::name::ResolveResult;

use super::generic_writing::GenericControlMetadata;

const FORM: &str = "urn:oasis:names:tc:opendocument:xmlns:form:1.0";
const OFFICE: &str = "urn:oasis:names:tc:opendocument:xmlns:office:1.0";
const SCRIPT: &str = "urn:oasis:names:tc:opendocument:xmlns:script:1.0";
const XFORMS: &str = "http://www.w3.org/2002/xforms";
const XLINK: &str = "http://www.w3.org/1999/xlink";
const XML: &str = "http://www.w3.org/XML/1998/namespace";
const MAX_XML: usize = 64 * 1024 * 1024;
const MAX_STRING: usize = 1024 * 1024;
const MAX_REFERENCE: usize = 64 * 1024;
const MAX_AGGREGATE: usize = 16 * 1024 * 1024;
const MAX_FORMS: usize = 4096;
const MAX_CONTROLS: usize = 65_536;

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ValueRangeControl {
    pub name: String,
    pub xml_id: String,
    pub metadata: GenericControlMetadata,
    pub input_required: Option<bool>,
    pub disabled: Option<bool>,
    pub printable: Option<bool>,
    pub tab_index: Option<ValueRangeNonNegativeInteger>,
    pub tab_stop: Option<bool>,
    pub title: Option<String>,
    pub value: Option<String>,
    pub linked_cell: Option<String>,
    pub repeat: Option<bool>,
    pub delay_for_repeat: Option<ValueRangeDuration>,
    pub max_value: Option<ValueRangeInteger>,
    pub min_value: Option<ValueRangeInteger>,
    pub step_size: Option<ValueRangePositiveInteger>,
    pub page_step_size: Option<ValueRangePositiveInteger>,
    pub orientation: Option<ValueRangeOrientation>,
}

impl ValueRangeControl {
    pub fn new(name: impl Into<String>, xml_id: impl Into<String>) -> Self {
        Self {
            name: name.into(),
            xml_id: xml_id.into(),
            metadata: GenericControlMetadata::default(),
            input_required: None,
            disabled: None,
            printable: None,
            tab_index: None,
            tab_stop: None,
            title: None,
            value: None,
            linked_cell: None,
            repeat: None,
            delay_for_repeat: None,
            max_value: None,
            min_value: None,
            step_size: None,
            page_step_size: None,
            orientation: None,
        }
    }

    pub fn to_xml_fragment(&self) -> Result<String> {
        value_range_xml(self)
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ValueRangeForm {
    pub name: String,
    pub controls: Vec<ValueRangeControl>,
    pub apply_filter: Option<bool>,
}

impl ValueRangeForm {
    pub fn new(name: impl Into<String>) -> Self {
        Self {
            name: name.into(),
            controls: Vec::new(),
            apply_filter: None,
        }
    }

    pub fn add_control(&mut self, control: ValueRangeControl) -> Result<()> {
        validate_control(&control)?;
        if self
            .controls
            .iter()
            .any(|existing| existing.name == control.name)
        {
            return invalid(format!(
                "duplicate value-range control name '{}'",
                control.name
            ));
        }
        if self
            .controls
            .iter()
            .any(|existing| existing.xml_id == control.xml_id)
        {
            return invalid(format!(
                "duplicate value-range control xml:id '{}'",
                control.xml_id
            ));
        }
        if self.controls.len() >= MAX_CONTROLS {
            return invalid("too many value-range controls");
        }
        self.controls.push(control);
        Ok(())
    }

    pub fn to_xml_fragment(&self) -> Result<String> {
        validate_name("form name", &self.name)?;
        validate_controls(&self.controls)?;
        let mut out = format!(
            r#"<form:form xmlns:form="{FORM}" xmlns:xforms="{XFORMS}" form:name="{}""#,
            escape(&self.name)
        );
        push_bool(&mut out, "form:apply-filter", self.apply_filter);
        if self.controls.is_empty() {
            out.push_str("/>");
        } else {
            out.push('>');
            for control in &self.controls {
                out.push_str(&control.to_xml_fragment()?);
            }
            out.push_str("</form:form>");
        }
        Ok(out)
    }
}

pub fn value_range_controls(xml: &str) -> Result<Vec<ValueRangeControl>> {
    Ok(scan(xml)?
        .controls
        .into_iter()
        .map(|item| item.control)
        .collect())
}

pub fn insert_value_range_control_xml(
    xml: &str,
    form_index: usize,
    control: &ValueRangeControl,
) -> Result<String> {
    validate_control(control)?;
    let scan = scan(xml)?;
    let form = scan
        .forms
        .get(form_index)
        .ok_or_else(|| Error::InvalidFormat(format!("form {form_index} is out of bounds")))?;
    reject_duplicate(form, control, None)?;
    if scan.ids.iter().any(|id| id == &control.xml_id) {
        return invalid(format!("duplicate xml:id '{}'", control.xml_id));
    }
    let fragment = bind_fragment(control.to_xml_fragment()?);
    match form.site.clone() {
        Site::Paired { close_start } => apply(xml, close_start..close_start, &fragment),
        Site::Empty { start, end, qname } => expand_empty(xml, start, end, &qname, &fragment),
    }
}

pub fn replace_value_range_control_xml(
    xml: &str,
    index: usize,
    replacement: &ValueRangeControl,
) -> Result<String> {
    validate_control(replacement)?;
    let scan = scan(xml)?;
    let old = scan.controls.get(index).ok_or_else(|| {
        Error::InvalidFormat(format!("value-range control {index} is out of bounds"))
    })?;
    reject_duplicate(&scan.forms[old.form], replacement, Some(&old.control))?;
    if replacement.xml_id != old.control.xml_id
        && scan.ids.iter().any(|id| id == &replacement.xml_id)
    {
        return invalid(format!("duplicate xml:id '{}'", replacement.xml_id));
    }
    apply(
        xml,
        old.span.clone(),
        &bind_fragment(replacement.to_xml_fragment()?),
    )
}

pub fn remove_value_range_control_xml(xml: &str, index: usize) -> Result<String> {
    let scan = scan(xml)?;
    let old = scan.controls.get(index).ok_or_else(|| {
        Error::InvalidFormat(format!("value-range control {index} is out of bounds"))
    })?;
    apply(xml, old.span.clone(), "")
}

fn value_range_xml(value: &ValueRangeControl) -> Result<String> {
    validate_control(value)?;
    let mut out = format!(
        r#"<form:value-range form:name="{}" xml:id="{}""#,
        escape(&value.name),
        escape(&value.xml_id)
    );
    push_string(&mut out, "form:id", value.metadata.form_id.as_deref());
    push_string(
        &mut out,
        "form:control-implementation",
        value.metadata.control_implementation.as_deref(),
    );
    push_string(
        &mut out,
        "xforms:bind",
        value.metadata.xforms_bind.as_deref(),
    );
    push_bool(&mut out, "form:input-required", value.input_required);
    push_bool(&mut out, "form:disabled", value.disabled);
    push_bool(&mut out, "form:printable", value.printable);
    push_string(
        &mut out,
        "form:tab-index",
        value
            .tab_index
            .as_ref()
            .map(ValueRangeNonNegativeInteger::as_str),
    );
    push_bool(&mut out, "form:tab-stop", value.tab_stop);
    push_string(&mut out, "form:title", value.title.as_deref());
    push_string(&mut out, "form:value", value.value.as_deref());
    push_string(&mut out, "form:linked-cell", value.linked_cell.as_deref());
    push_bool(&mut out, "form:repeat", value.repeat);
    push_string(
        &mut out,
        "form:delay-for-repeat",
        value
            .delay_for_repeat
            .as_ref()
            .map(ValueRangeDuration::as_str),
    );
    push_string(
        &mut out,
        "form:max-value",
        value.max_value.as_ref().map(ValueRangeInteger::as_str),
    );
    push_string(
        &mut out,
        "form:min-value",
        value.min_value.as_ref().map(ValueRangeInteger::as_str),
    );
    push_string(
        &mut out,
        "form:step-size",
        value
            .step_size
            .as_ref()
            .map(ValueRangePositiveInteger::as_str),
    );
    push_string(
        &mut out,
        "form:page-step-size",
        value
            .page_step_size
            .as_ref()
            .map(ValueRangePositiveInteger::as_str),
    );
    push_string(
        &mut out,
        "form:orientation",
        value.orientation.map(ValueRangeOrientation::as_str),
    );
    out.push_str("/>");
    Ok(out)
}

fn validate_control(value: &ValueRangeControl) -> Result<()> {
    validate_name("value-range control name", &value.name)?;
    validate_xml_id(&value.xml_id)?;
    if let Some(form_id) = value.metadata.form_id.as_deref() {
        validate_xml_id(form_id)?;
    }
    validate_optional(
        "control implementation",
        value.metadata.control_implementation.as_deref(),
        MAX_REFERENCE,
    )?;
    validate_optional(
        "XForms bind",
        value.metadata.xforms_bind.as_deref(),
        MAX_REFERENCE,
    )?;
    validate_optional("value-range title", value.title.as_deref(), MAX_STRING)?;
    validate_optional("value-range value", value.value.as_deref(), MAX_STRING)?;
    validate_optional(
        "value-range linked cell",
        value.linked_cell.as_deref(),
        MAX_REFERENCE,
    )?;
    if matches!((&value.min_value, &value.max_value), (Some(min), Some(max)) if min > max) {
        return invalid("form:min-value cannot exceed form:max-value");
    }
    Ok(())
}

fn validate_controls(controls: &[ValueRangeControl]) -> Result<()> {
    if controls.len() > MAX_CONTROLS {
        return invalid("too many value-range controls");
    }
    let mut names = Vec::<&str>::new();
    let mut ids = Vec::<&str>::new();
    let mut aggregate = 0usize;
    for control in controls {
        validate_control(control)?;
        if names.contains(&control.name.as_str()) {
            return invalid(format!(
                "duplicate value-range control name '{}'",
                control.name
            ));
        }
        if ids.contains(&control.xml_id.as_str()) {
            return invalid(format!(
                "duplicate value-range control xml:id '{}'",
                control.xml_id
            ));
        }
        names.push(&control.name);
        ids.push(&control.xml_id);
        aggregate = aggregate.saturating_add(control_size(control));
        if aggregate > MAX_AGGREGATE {
            return invalid("value-range control strings exceed 16 MiB");
        }
    }
    Ok(())
}

fn control_size(value: &ValueRangeControl) -> usize {
    [&value.name, &value.xml_id]
        .iter()
        .fold(0usize, |sum, value| sum.saturating_add(value.len()))
        .saturating_add(value.metadata.form_id.as_ref().map_or(0, String::len))
        .saturating_add(
            value
                .metadata
                .control_implementation
                .as_ref()
                .map_or(0, String::len),
        )
        .saturating_add(value.metadata.xforms_bind.as_ref().map_or(0, String::len))
        .saturating_add(value.title.as_ref().map_or(0, String::len))
        .saturating_add(value.value.as_ref().map_or(0, String::len))
        .saturating_add(value.linked_cell.as_ref().map_or(0, String::len))
        .saturating_add(
            value
                .delay_for_repeat
                .as_ref()
                .map_or(0, |value| value.as_str().len()),
        )
        .saturating_add(
            value
                .max_value
                .as_ref()
                .map_or(0, |value| value.as_str().len()),
        )
        .saturating_add(
            value
                .min_value
                .as_ref()
                .map_or(0, |value| value.as_str().len()),
        )
        .saturating_add(
            value
                .step_size
                .as_ref()
                .map_or(0, |value| value.as_str().len()),
        )
        .saturating_add(
            value
                .page_step_size
                .as_ref()
                .map_or(0, |value| value.as_str().len()),
        )
}

#[derive(Clone)]
struct ControlLocation {
    span: Range<usize>,
    form: usize,
    control: ValueRangeControl,
}
#[derive(Clone)]
struct FormLocation {
    site: Site,
    controls: Vec<ControlLocation>,
}
#[derive(Clone)]
enum Site {
    Paired {
        close_start: usize,
    },
    Empty {
        start: usize,
        end: usize,
        qname: String,
    },
}
struct Scan {
    forms: Vec<FormLocation>,
    controls: Vec<ControlLocation>,
    ids: Vec<String>,
}
struct Open {
    local: Vec<u8>,
    form: Option<usize>,
    control: Option<usize>,
}
#[derive(Clone)]
struct Attr {
    namespace: Option<String>,
    local: String,
    value: String,
}

fn scan(xml: &str) -> Result<Scan> {
    if xml.len() > MAX_XML {
        return invalid("value-range form XML exceeds 64 MiB");
    }
    let mut reader = NsReader::from_str(xml);
    let mut buffer = Vec::new();
    let mut previous = 0usize;
    let mut stack = Vec::<Open>::new();
    let mut form_stack = Vec::<usize>::new();
    let mut forms = Vec::<FormLocation>::new();
    let mut controls = Vec::<ControlLocation>::new();
    let mut ids = Vec::<String>::new();
    let mut aggregate = 0usize;
    loop {
        let (namespace, event) = reader
            .read_resolved_event_into(&mut buffer)
            .map_err(|error| {
                Error::InvalidFormat(format!("invalid value-range form XML: {error}"))
            })?;
        let namespace = match namespace {
            ResolveResult::Bound(uri) => Some(String::from_utf8_lossy(uri.as_ref()).into_owned()),
            _ => None,
        };
        let end = reader.buffer_position() as usize;
        match event {
            Event::Start(ref element) => {
                if stack.len() >= 128 {
                    return invalid("value-range form XML nesting exceeds 128 levels");
                }
                track_xml_id(&reader, element, &mut ids)?;
                let local = element.local_name().as_ref().to_vec();
                let mut form = None;
                let mut control = None;
                if namespace.as_deref() == Some(FORM) && local == b"form" {
                    validate_form(&reader, element)?;
                    if forms.len() >= MAX_FORMS {
                        return invalid("too many forms");
                    }
                    form = Some(forms.len());
                    forms.push(FormLocation {
                        site: Site::Paired { close_start: 0 },
                        controls: Vec::new(),
                    });
                    form_stack.push(form.unwrap());
                } else if namespace.as_deref() == Some(FORM) && local == b"value-range" {
                    if stack.iter().any(|open| open.control.is_some()) {
                        return invalid("value-range controls cannot be nested");
                    }
                    let owner = *form_stack.last().ok_or_else(|| {
                        Error::InvalidFormat("value-range control has no form owner".to_string())
                    })?;
                    let parsed = parse_control(&reader, element)?;
                    aggregate = aggregate.saturating_add(control_size(&parsed));
                    if aggregate > MAX_AGGREGATE {
                        return invalid("value-range control strings exceed 16 MiB");
                    }
                    if controls.len() >= MAX_CONTROLS {
                        return invalid("too many value-range controls");
                    }
                    control = Some(controls.len());
                    controls.push(ControlLocation {
                        span: previous..0,
                        form: owner,
                        control: parsed,
                    });
                } else if is_active(namespace.as_deref(), local.as_slice())
                    && !form_stack.is_empty()
                {
                    return invalid(
                        "event and macro content is outside the value-range mutation API",
                    );
                } else if stack.iter().any(|open| open.control.is_some())
                    && !is_property(namespace.as_deref(), local.as_slice())
                {
                    return invalid("unexpected child element in value-range control");
                }
                stack.push(Open {
                    local,
                    form,
                    control,
                });
            },
            Event::Empty(ref element) => {
                track_xml_id(&reader, element, &mut ids)?;
                let local = element.local_name().as_ref().to_vec();
                if namespace.as_deref() == Some(FORM) && local == b"form" {
                    validate_form(&reader, element)?;
                    if forms.len() >= MAX_FORMS {
                        return invalid("too many forms");
                    }
                    forms.push(FormLocation {
                        site: Site::Empty {
                            start: previous,
                            end,
                            qname: qname(element)?,
                        },
                        controls: Vec::new(),
                    });
                } else if namespace.as_deref() == Some(FORM) && local == b"value-range" {
                    if stack.iter().any(|open| open.control.is_some()) {
                        return invalid("value-range controls cannot be nested");
                    }
                    let owner = *form_stack.last().ok_or_else(|| {
                        Error::InvalidFormat("value-range control has no form owner".to_string())
                    })?;
                    let parsed = parse_control(&reader, element)?;
                    aggregate = aggregate.saturating_add(control_size(&parsed));
                    if aggregate > MAX_AGGREGATE {
                        return invalid("value-range control strings exceed 16 MiB");
                    }
                    if controls.len() >= MAX_CONTROLS {
                        return invalid("too many value-range controls");
                    }
                    controls.push(ControlLocation {
                        span: previous..end,
                        form: owner,
                        control: parsed,
                    });
                } else if is_active(namespace.as_deref(), local.as_slice())
                    && !form_stack.is_empty()
                {
                    return invalid(
                        "event and macro content is outside the value-range mutation API",
                    );
                } else if stack.iter().any(|open| open.control.is_some())
                    && !is_property(namespace.as_deref(), local.as_slice())
                {
                    return invalid("unexpected child element in value-range control");
                }
            },
            Event::Text(text) if stack.iter().any(|open| open.control.is_some()) => {
                let decoded = text.decode().map_err(|error| {
                    Error::InvalidFormat(format!("invalid value-range control text: {error}"))
                })?;
                if !decoded.trim().is_empty() {
                    return invalid("value-range controls cannot contain character data");
                }
            },
            Event::CData(text) if stack.iter().any(|open| open.control.is_some()) => {
                let decoded = text.decode().map_err(|error| {
                    Error::InvalidFormat(format!("invalid value-range control CDATA: {error}"))
                })?;
                if !decoded.trim().is_empty() {
                    return invalid("value-range controls cannot contain CDATA");
                }
            },
            Event::GeneralRef(_) if stack.iter().any(|open| open.control.is_some()) => {
                return invalid("value-range controls cannot contain entity references");
            },
            Event::End(ref element) => {
                let open = stack.pop().ok_or_else(|| {
                    Error::InvalidFormat("value-range form XML stack underflow".to_string())
                })?;
                if open.local.as_slice() != element.local_name().as_ref() {
                    return invalid("mismatched value-range form XML elements");
                }
                if let Some(index) = open.control {
                    controls[index].span.end = end;
                }
                if let Some(index) = open.form {
                    forms[index].site = Site::Paired {
                        close_start: previous,
                    };
                    form_stack.pop();
                }
            },
            Event::DocType(_) => return invalid("DOCTYPE is not allowed in value-range form XML"),
            Event::Eof => break,
            _ => {},
        }
        previous = end;
        buffer.clear();
    }
    if !stack.is_empty() {
        return invalid("unclosed value-range form XML elements");
    }
    for control in &controls {
        forms[control.form].controls.push(control.clone());
    }
    for form in &forms {
        validate_controls(
            &form
                .controls
                .iter()
                .map(|item| item.control.clone())
                .collect::<Vec<_>>(),
        )?;
    }
    Ok(Scan {
        forms,
        controls,
        ids,
    })
}

fn parse_control(reader: &NsReader<&[u8]>, element: &BytesStart<'_>) -> Result<ValueRangeControl> {
    let attrs = attributes(reader, element)?;
    validate_allowed(&attrs, VALUE_RANGE_ATTRS)?;
    let mut value = ValueRangeControl::new(
        required(&attrs, FORM, "name")?,
        required(&attrs, XML, "id")?,
    );
    value.metadata = GenericControlMetadata {
        form_id: optional(&attrs, FORM, "id"),
        control_implementation: optional(&attrs, FORM, "control-implementation"),
        xforms_bind: optional(&attrs, XFORMS, "bind"),
    };
    value.input_required = optional_bool(&attrs, FORM, "input-required")?;
    value.disabled = optional_bool(&attrs, FORM, "disabled")?;
    value.printable = optional_bool(&attrs, FORM, "printable")?;
    value.tab_index = optional(&attrs, FORM, "tab-index")
        .map(ValueRangeNonNegativeInteger::new)
        .transpose()
        .map_err(Error::InvalidFormat)?;
    value.tab_stop = optional_bool(&attrs, FORM, "tab-stop")?;
    value.title = optional(&attrs, FORM, "title");
    value.value = optional(&attrs, FORM, "value");
    value.linked_cell = optional(&attrs, FORM, "linked-cell");
    value.repeat = optional_bool(&attrs, FORM, "repeat")?;
    value.delay_for_repeat = optional(&attrs, FORM, "delay-for-repeat")
        .map(ValueRangeDuration::new)
        .transpose()
        .map_err(Error::InvalidFormat)?;
    value.max_value = optional(&attrs, FORM, "max-value")
        .map(ValueRangeInteger::new)
        .transpose()
        .map_err(Error::InvalidFormat)?;
    value.min_value = optional(&attrs, FORM, "min-value")
        .map(ValueRangeInteger::new)
        .transpose()
        .map_err(Error::InvalidFormat)?;
    value.step_size = optional(&attrs, FORM, "step-size")
        .map(ValueRangePositiveInteger::new)
        .transpose()
        .map_err(Error::InvalidFormat)?;
    value.page_step_size = optional(&attrs, FORM, "page-step-size")
        .map(ValueRangePositiveInteger::new)
        .transpose()
        .map_err(Error::InvalidFormat)?;
    value.orientation = optional(&attrs, FORM, "orientation")
        .map(|value| value.parse())
        .transpose()
        .map_err(Error::InvalidFormat)?;
    validate_control(&value)?;
    Ok(value)
}

const VALUE_RANGE_ATTRS: &[(&str, &str)] = &[
    (FORM, "name"),
    (XML, "id"),
    (FORM, "id"),
    (FORM, "control-implementation"),
    (FORM, "input-required"),
    (XFORMS, "bind"),
    (FORM, "disabled"),
    (FORM, "printable"),
    (FORM, "tab-index"),
    (FORM, "tab-stop"),
    (FORM, "title"),
    (FORM, "value"),
    (FORM, "linked-cell"),
    (FORM, "repeat"),
    (FORM, "delay-for-repeat"),
    (FORM, "max-value"),
    (FORM, "min-value"),
    (FORM, "step-size"),
    (FORM, "page-step-size"),
    (FORM, "orientation"),
];
const FORM_ATTRS: &[(&str, &str)] = &[
    (FORM, "name"),
    (FORM, "control-implementation"),
    (FORM, "service-name"),
    (FORM, "allow-deletes"),
    (FORM, "allow-inserts"),
    (FORM, "allow-updates"),
    (FORM, "apply-filter"),
    (FORM, "command"),
    (FORM, "command-type"),
    (FORM, "datasource"),
    (FORM, "detail-fields"),
    (FORM, "enctype"),
    (FORM, "escape-processing"),
    (FORM, "filter"),
    (FORM, "ignore-result"),
    (FORM, "master-fields"),
    (FORM, "method"),
    (FORM, "navigation-mode"),
    (FORM, "order"),
    (FORM, "tab-cycle"),
    (OFFICE, "target-frame"),
    (XLINK, "type"),
    (XLINK, "href"),
    (XLINK, "actuate"),
];

fn is_property(namespace: Option<&str>, local: &[u8]) -> bool {
    namespace == Some(FORM)
        && matches!(
            local,
            b"properties" | b"property" | b"list-property" | b"list-value"
        )
}
fn is_active(namespace: Option<&str>, local: &[u8]) -> bool {
    (namespace == Some(OFFICE) && matches!(local, b"event-listeners" | b"events" | b"scripts"))
        || namespace == Some(SCRIPT)
}

fn validate_form(reader: &NsReader<&[u8]>, element: &BytesStart<'_>) -> Result<()> {
    let attrs = attributes(reader, element)?;
    validate_allowed(&attrs, FORM_ATTRS)?;
    validate_name("form name", &required(&attrs, FORM, "name")?)?;
    for attr in &attrs {
        if attr.namespace.as_deref() == Some(FORM)
            && matches!(
                attr.local.as_str(),
                "allow-deletes"
                    | "allow-inserts"
                    | "allow-updates"
                    | "apply-filter"
                    | "escape-processing"
                    | "ignore-result"
            )
        {
            parse_bool(&attr.value, &attr.local)?;
        }
    }
    Ok(())
}

fn attributes(reader: &NsReader<&[u8]>, element: &BytesStart<'_>) -> Result<Vec<Attr>> {
    let mut out = Vec::new();
    for attr in element.attributes().with_checks(true) {
        let attr = attr.map_err(|error| {
            Error::InvalidFormat(format!("invalid value-range control attribute: {error}"))
        })?;
        let (namespace, local) = reader.resolver().resolve_attribute(attr.key);
        let namespace = match namespace {
            ResolveResult::Bound(uri) => Some(String::from_utf8_lossy(uri.as_ref()).into_owned()),
            _ => None,
        };
        if namespace.as_deref() == Some("http://www.w3.org/2000/xmlns/")
            || attr.key.as_ref() == b"xmlns"
        {
            continue;
        }
        let value = attr
            .decoded_and_normalized_value(quick_xml::XmlVersion::Implicit1_0, reader.decoder())
            .map_err(|error| {
                Error::InvalidFormat(format!(
                    "invalid value-range control attribute value: {error}"
                ))
            })?
            .into_owned();
        out.push(Attr {
            namespace,
            local: String::from_utf8_lossy(local.as_ref()).into_owned(),
            value,
        });
    }
    Ok(out)
}

fn track_xml_id(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
    ids: &mut Vec<String>,
) -> Result<()> {
    if let Some(id) = optional(&attributes(reader, element)?, XML, "id") {
        validate_xml_id(&id)?;
        if ids.iter().any(|existing| existing == &id) {
            return invalid(format!("duplicate xml:id '{id}'"));
        }
        ids.push(id);
    }
    Ok(())
}
fn validate_allowed(attrs: &[Attr], allowed: &[(&str, &str)]) -> Result<()> {
    for attr in attrs {
        if !allowed.iter().any(|(namespace, local)| {
            attr.namespace.as_deref() == Some(*namespace) && attr.local == *local
        }) {
            return invalid(format!(
                "unexpected value-range control attribute '{}'",
                attr.local
            ));
        }
    }
    Ok(())
}
fn optional(attrs: &[Attr], namespace: &str, local: &str) -> Option<String> {
    attrs
        .iter()
        .find(|attr| attr.namespace.as_deref() == Some(namespace) && attr.local == local)
        .map(|attr| attr.value.clone())
}
fn required(attrs: &[Attr], namespace: &str, local: &str) -> Result<String> {
    optional(attrs, namespace, local)
        .ok_or_else(|| Error::InvalidFormat(format!("missing required attribute {local}")))
}
fn optional_bool(attrs: &[Attr], namespace: &str, local: &str) -> Result<Option<bool>> {
    optional(attrs, namespace, local)
        .map(|value| parse_bool(&value, local))
        .transpose()
}
fn parse_bool(value: &str, local: &str) -> Result<bool> {
    match value {
        "true" | "1" => Ok(true),
        "false" | "0" => Ok(false),
        _ => invalid(format!("invalid boolean '{value}' for {local}")),
    }
}
fn validate_name(label: &str, value: &str) -> Result<()> {
    validate_string(label, value, MAX_STRING)?;
    if value.is_empty() {
        invalid(format!("{label} cannot be empty"))
    } else {
        Ok(())
    }
}
fn validate_xml_id(value: &str) -> Result<()> {
    validate_name("value-range control xml:id", value)?;
    let mut chars = value.chars();
    let first = chars.next().unwrap();
    if !(first == '_' || first.is_ascii_alphabetic())
        || !chars.all(|ch| ch == '_' || ch == '-' || ch == '.' || ch.is_ascii_alphanumeric())
    {
        return invalid(format!("invalid value-range control xml:id '{value}'"));
    }
    Ok(())
}
fn validate_optional(label: &str, value: Option<&str>, limit: usize) -> Result<()> {
    if let Some(value) = value {
        validate_string(label, value, limit)?;
    }
    Ok(())
}
fn validate_string(label: &str, value: &str, limit: usize) -> Result<()> {
    if value.len() > limit {
        return invalid(format!("{label} exceeds {limit} bytes"));
    }
    if value.chars().any(
        |ch| !matches!(ch as u32, 9 | 10 | 13 | 32..=0xD7FF | 0xE000..=0xFFFD | 0x10000..=0x10FFFF),
    ) {
        return invalid(format!("{label} contains invalid XML characters"));
    }
    Ok(())
}
fn reject_duplicate(
    form: &FormLocation,
    replacement: &ValueRangeControl,
    current: Option<&ValueRangeControl>,
) -> Result<()> {
    for item in &form.controls {
        if current.is_some_and(|value| value.xml_id == item.control.xml_id) {
            continue;
        }
        if item.control.name == replacement.name {
            return invalid(format!(
                "duplicate value-range control name '{}'",
                replacement.name
            ));
        }
        if item.control.xml_id == replacement.xml_id {
            return invalid(format!(
                "duplicate value-range control xml:id '{}'",
                replacement.xml_id
            ));
        }
    }
    Ok(())
}
fn push_string(out: &mut String, name: &str, value: Option<&str>) {
    if let Some(value) = value {
        out.push_str(&format!(r#" {name}="{}""#, escape(value)));
    }
}
fn push_bool(out: &mut String, name: &str, value: Option<bool>) {
    if let Some(value) = value {
        out.push_str(&format!(
            r#" {name}="{}""#,
            if value { "true" } else { "false" }
        ));
    }
}
fn escape(value: &str) -> String {
    value
        .replace('&', "&amp;")
        .replace('<', "&lt;")
        .replace('"', "&quot;")
        .replace('\'', "&apos;")
}
fn bind_fragment(mut value: String) -> String {
    if value.contains("form:") && !value.contains("xmlns:form=") {
        value = value.replacen(' ', &format!(" xmlns:form=\"{FORM}\" "), 1);
    }
    if value.contains("xforms:") && !value.contains("xmlns:xforms=") {
        value = value.replacen(' ', &format!(" xmlns:xforms=\"{XFORMS}\" "), 1);
    }
    value
}
fn qname(element: &BytesStart<'_>) -> Result<String> {
    String::from_utf8(element.name().as_ref().to_vec())
        .map_err(|_| Error::InvalidFormat("invalid value-range form element name".to_string()))
}
fn apply(xml: &str, span: Range<usize>, replacement: &str) -> Result<String> {
    let mut out = String::with_capacity(xml.len() - span.len() + replacement.len());
    out.push_str(&xml[..span.start]);
    out.push_str(replacement);
    out.push_str(&xml[span.end..]);
    Ok(out)
}
fn expand_empty(xml: &str, start: usize, end: usize, qname: &str, content: &str) -> Result<String> {
    let raw = &xml[start..end];
    let slash = raw
        .rfind("/>")
        .ok_or_else(|| Error::InvalidFormat("invalid empty form element".to_string()))?;
    apply(
        xml,
        start..end,
        &format!("{}>{content}</{qname}>", &raw[..slash]),
    )
}
fn invalid<T>(message: impl Into<String>) -> Result<T> {
    Err(Error::InvalidFormat(message.into()))
}

#[cfg(test)]
mod tests {
    use super::*;

    fn document(control: &str) -> String {
        format!(
            r#"<office:document-content xmlns:office="{OFFICE}" xmlns:form="{FORM}" xmlns:xforms="{XFORMS}"><office:body><office:text><form:form form:name="Ranges">{control}</form:form><office:tail/></office:text></office:body></office:document-content>"#
        )
    }

    #[test]
    fn normative_attributes_are_typed_and_canonical() {
        let xml = document(
            r#"<form:value-range form:name="Spin" xml:id="spin1" form:id="legacy" form:control-implementation="ooo:com.sun.star.form.component.SpinButton" xforms:bind="bind-1" form:input-required="true" form:disabled="false" form:printable="true" form:tab-index="+0004" form:tab-stop="false" form:title="Amount" form:value="50" form:linked-cell="Sheet1.A1" form:repeat="true" form:delay-for-repeat="PT0.5S" form:max-value="+00010000000000000000000000000000000000000" form:min-value="-0002" form:step-size="+0001" form:page-step-size="00010" form:orientation="horizontal"/>"#,
        );
        let value = value_range_controls(&xml).unwrap().remove(0);
        assert_eq!(value.tab_index.as_ref().unwrap().as_str(), "4");
        assert_eq!(
            value.max_value.as_ref().unwrap().as_str(),
            "10000000000000000000000000000000000000"
        );
        assert_eq!(value.min_value.as_ref().unwrap().as_str(), "-2");
        assert_eq!(value.step_size.as_ref().unwrap().as_str(), "1");
        assert_eq!(value.page_step_size.as_ref().unwrap().as_str(), "10");
        assert_eq!(value.delay_for_repeat.as_ref().unwrap().as_str(), "PT0.5S");
        assert_eq!(value.orientation, Some(ValueRangeOrientation::Horizontal));
        let canonical = value.to_xml_fragment().unwrap();
        assert!(canonical.contains(r#"form:tab-index="4""#));
        assert!(canonical.contains(r#"form:min-value="-2""#));
    }

    #[test]
    fn lexical_domains_bounds_and_hostile_xml_are_rejected() {
        assert_eq!(ValueRangeInteger::new("-000").unwrap().as_str(), "0");
        assert_eq!(
            ValueRangeNonNegativeInteger::new("-0").unwrap().as_str(),
            "0"
        );
        assert!(ValueRangePositiveInteger::new("0").is_err());
        assert!(ValueRangeDuration::new("P1Y2M3DT4H5M6.75S").is_ok());
        for invalid in ["P", "PT", "+P1D", "P1S", "PT.5S", "P1DT", "PT1.0M"] {
            assert!(ValueRangeDuration::new(invalid).is_err());
        }
        let mut inverted = ValueRangeControl::new("Spin", "spin1");
        inverted.min_value = Some(ValueRangeInteger::new("2").unwrap());
        inverted.max_value = Some(ValueRangeInteger::new("1").unwrap());
        assert!(inverted.to_xml_fragment().is_err());
        assert!(
            value_range_controls(&document(
                r#"<form:value-range form:name="Spin" xml:id="spin1" form:step-size="0"/>"#
            ))
            .is_err()
        );
        assert!(value_range_controls(&document(r#"<form:value-range xmlns:e="urn:evil" form:name="Spin" xml:id="spin1" e:value="4"/>"#)).is_err());
        assert!(value_range_controls(&document(r#"<form:value-range form:name="Spin" xml:id="spin1"><office:event-listeners/></form:value-range>"#)).is_err());
    }

    #[test]
    fn lossless_mutation_and_package_facades_round_trip() {
        let xml = document(r#"<form:value-range form:name="Old" xml:id="old" form:value="7"/>"#);
        let inserted = ValueRangeControl::new("New", "new");
        let updated = insert_value_range_control_xml(&xml, 0, &inserted).unwrap();
        assert!(updated.contains("<office:tail/>"));
        let replacement = ValueRangeControl::new("Replacement", "replacement");
        let updated = replace_value_range_control_xml(&updated, 0, &replacement).unwrap();
        assert_eq!(
            value_range_controls(&remove_value_range_control_xml(&updated, 1).unwrap()).unwrap(),
            std::slice::from_ref(&replacement)
        );

        let mut form = ValueRangeForm::new("Ranges");
        form.add_control(ValueRangeControl::new("Initial", "initial"))
            .unwrap();
        let mut builder = crate::Builder::new();
        builder.add_value_range_form(&form).unwrap();
        let document = crate::Document::from_bytes(builder.build().unwrap()).unwrap();
        let mut mutable = crate::MutableDocument::from_document(document).unwrap();
        assert_eq!(mutable.value_range_controls().unwrap().len(), 1);
        mutable.insert_value_range_control(0, &inserted).unwrap();
        assert_eq!(
            mutable
                .replace_value_range_control(0, &replacement)
                .unwrap()
                .name,
            "Initial"
        );
        assert_eq!(mutable.remove_value_range_control(1).unwrap().name, "New");
    }
}

use std::cmp::Ordering;
use std::fmt;
use std::str::FromStr;

const MAX_VALUE_RANGE_INTEGER_DIGITS: usize = 4096;
const MAX_VALUE_RANGE_DURATION_LEN: usize = 256;

/// The direction in which a value-range control changes its value.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum ValueRangeOrientation {
    Horizontal,
    Vertical,
}

impl ValueRangeOrientation {
    pub const fn as_str(self) -> &'static str {
        match self {
            Self::Horizontal => "horizontal",
            Self::Vertical => "vertical",
        }
    }
}

impl fmt::Display for ValueRangeOrientation {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        f.write_str(self.as_str())
    }
}

impl FromStr for ValueRangeOrientation {
    type Err = String;

    fn from_str(value: &str) -> std::result::Result<Self, Self::Err> {
        match value {
            "horizontal" => Ok(Self::Horizontal),
            "vertical" => Ok(Self::Vertical),
            _ => Err(format!("invalid form:orientation value {value:?}")),
        }
    }
}

fn canonical_integer(value: &str) -> std::result::Result<String, String> {
    if value.is_empty() || value.len() > MAX_VALUE_RANGE_INTEGER_DIGITS + 1 {
        return Err("integer lexical form is empty or exceeds the safety limit".into());
    }
    let (negative, digits) = match value.as_bytes()[0] {
        b'+' => (false, &value[1..]),
        b'-' => (true, &value[1..]),
        _ => (false, value),
    };
    if digits.is_empty() || digits.len() > MAX_VALUE_RANGE_INTEGER_DIGITS {
        return Err("integer lexical form is empty or exceeds the safety limit".into());
    }
    if !digits.bytes().all(|byte| byte.is_ascii_digit()) {
        return Err(format!("invalid XML Schema integer lexical form {value:?}"));
    }
    let digits = digits.trim_start_matches('0');
    if digits.is_empty() {
        return Ok("0".into());
    }
    Ok(if negative {
        format!("-{digits}")
    } else {
        digits.into()
    })
}

/// An arbitrary-precision XML Schema `integer` used by value-range bounds.
#[derive(Debug, Clone, PartialEq, Eq, Hash)]
pub struct ValueRangeInteger(String);

impl ValueRangeInteger {
    pub fn new(value: impl AsRef<str>) -> std::result::Result<Self, String> {
        canonical_integer(value.as_ref()).map(Self)
    }

    pub fn as_str(&self) -> &str {
        &self.0
    }
}

impl fmt::Display for ValueRangeInteger {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        f.write_str(&self.0)
    }
}

impl FromStr for ValueRangeInteger {
    type Err = String;

    fn from_str(value: &str) -> std::result::Result<Self, Self::Err> {
        Self::new(value)
    }
}

impl Ord for ValueRangeInteger {
    fn cmp(&self, other: &Self) -> Ordering {
        let lhs_negative = self.0.starts_with('-');
        let rhs_negative = other.0.starts_with('-');
        match (lhs_negative, rhs_negative) {
            (true, false) => Ordering::Less,
            (false, true) => Ordering::Greater,
            (false, false) => self
                .0
                .len()
                .cmp(&other.0.len())
                .then_with(|| self.0.cmp(&other.0)),
            (true, true) => {
                let lhs = &self.0[1..];
                let rhs = &other.0[1..];
                rhs.len().cmp(&lhs.len()).then_with(|| rhs.cmp(lhs))
            },
        }
    }
}

impl PartialOrd for ValueRangeInteger {
    fn partial_cmp(&self, other: &Self) -> Option<Ordering> {
        Some(self.cmp(other))
    }
}

/// An arbitrary-precision XML Schema `nonNegativeInteger` used by `form:tab-index`.
#[derive(Debug, Clone, PartialEq, Eq, Hash)]
pub struct ValueRangeNonNegativeInteger(String);

impl ValueRangeNonNegativeInteger {
    pub fn new(value: impl AsRef<str>) -> std::result::Result<Self, String> {
        let value = canonical_integer(value.as_ref())?;
        if value.starts_with('-') {
            return Err("nonNegativeInteger cannot be negative".into());
        }
        Ok(Self(value))
    }

    pub fn as_str(&self) -> &str {
        &self.0
    }
}

impl fmt::Display for ValueRangeNonNegativeInteger {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        f.write_str(&self.0)
    }
}

impl FromStr for ValueRangeNonNegativeInteger {
    type Err = String;

    fn from_str(value: &str) -> std::result::Result<Self, Self::Err> {
        Self::new(value)
    }
}

/// An arbitrary-precision XML Schema `positiveInteger` used by step sizes.
#[derive(Debug, Clone, PartialEq, Eq, Hash)]
pub struct ValueRangePositiveInteger(String);

impl ValueRangePositiveInteger {
    pub fn new(value: impl AsRef<str>) -> std::result::Result<Self, String> {
        let value = canonical_integer(value.as_ref())?;
        if value == "0" || value.starts_with('-') {
            return Err("positiveInteger must be greater than zero".into());
        }
        Ok(Self(value))
    }

    pub fn as_str(&self) -> &str {
        &self.0
    }
}

impl fmt::Display for ValueRangePositiveInteger {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        f.write_str(&self.0)
    }
}

impl FromStr for ValueRangePositiveInteger {
    type Err = String;

    fn from_str(value: &str) -> std::result::Result<Self, Self::Err> {
        Self::new(value)
    }
}

/// A validated, exact XML Schema `duration` lexical form.
///
/// It is deliberately not converted to a clock duration: year/month components
/// have no context-independent length and form bindings must remain inert.
#[derive(Debug, Clone, PartialEq, Eq, Hash)]
pub struct ValueRangeDuration(String);

impl ValueRangeDuration {
    pub fn new(value: impl AsRef<str>) -> std::result::Result<Self, String> {
        let value = value.as_ref();
        if !valid_xsd_duration(value) {
            return Err(format!(
                "invalid XML Schema duration lexical form {value:?}"
            ));
        }
        Ok(Self(value.into()))
    }

    pub fn as_str(&self) -> &str {
        &self.0
    }
}

impl fmt::Display for ValueRangeDuration {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        f.write_str(&self.0)
    }
}

impl FromStr for ValueRangeDuration {
    type Err = String;

    fn from_str(value: &str) -> std::result::Result<Self, Self::Err> {
        Self::new(value)
    }
}

fn valid_xsd_duration(value: &str) -> bool {
    if value.is_empty() || value.len() > MAX_VALUE_RANGE_DURATION_LEN || !value.is_ascii() {
        return false;
    }
    let bytes = value.as_bytes();
    let mut offset = usize::from(bytes[0] == b'-');
    if bytes.get(offset) != Some(&b'P') {
        return false;
    }
    offset += 1;
    let mut in_time = false;
    let mut saw_component = false;
    let mut saw_time_component = false;
    let mut last_order = 0_u8;
    while offset < bytes.len() {
        if bytes[offset] == b'T' {
            if in_time {
                return false;
            }
            in_time = true;
            offset += 1;
            continue;
        }
        let digits_start = offset;
        while offset < bytes.len() && bytes[offset].is_ascii_digit() {
            offset += 1;
        }
        if offset == digits_start {
            return false;
        }
        let mut fractional = false;
        if bytes.get(offset) == Some(&b'.') {
            fractional = true;
            offset += 1;
            let fraction_start = offset;
            while offset < bytes.len() && bytes[offset].is_ascii_digit() {
                offset += 1;
            }
            if offset == fraction_start {
                return false;
            }
        }
        let Some(&designator) = bytes.get(offset) else {
            return false;
        };
        offset += 1;
        let order = match (in_time, designator) {
            (false, b'Y') => 1,
            (false, b'M') => 2,
            (false, b'D') => 3,
            (true, b'H') => 4,
            (true, b'M') => 5,
            (true, b'S') => 6,
            _ => return false,
        };
        if order <= last_order || (fractional && designator != b'S') {
            return false;
        }
        last_order = order;
        saw_component = true;
        saw_time_component |= in_time;
    }
    saw_component && (!in_time || saw_time_component)
}
