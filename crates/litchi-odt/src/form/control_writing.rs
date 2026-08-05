use std::ops::Range;

use quick_xml::NsReader;
use quick_xml::events::{BytesStart, Event};
use quick_xml::name::ResolveResult;

use litchi_core::{Error, Result};

const FORM: &str = "urn:oasis:names:tc:opendocument:xmlns:form:1.0";
const OFFICE: &str = "urn:oasis:names:tc:opendocument:xmlns:office:1.0";
const TEXT: &str = "urn:oasis:names:tc:opendocument:xmlns:text:1.0";
const XLINK: &str = "http://www.w3.org/1999/xlink";
const XML: &str = "http://www.w3.org/XML/1998/namespace";
const MAX_XML: usize = 64 * 1024 * 1024;
const MAX_FORMS: usize = 16_384;
const MAX_CONTROLS: usize = 65_536;
const MAX_STRING: usize = 1_048_576;
const MAX_RESOURCE: usize = 8_192;

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum TextControlKind {
    Text,
    Textarea,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct TextControl {
    pub kind: TextControlKind,
    pub name: String,
    pub xml_id: String,
    pub value: Option<String>,
    pub current_value: Option<String>,
    pub disabled: Option<bool>,
    pub readonly: Option<bool>,
    pub printable: Option<bool>,
    pub max_length: Option<u64>,
    pub tab_index: Option<u64>,
    pub tab_stop: Option<bool>,
    pub title: Option<String>,
    pub convert_empty_to_null: Option<bool>,
    pub data_field: Option<String>,
    pub linked_cell: Option<String>,
    pub paragraphs: Vec<String>,
}

impl TextControl {
    pub fn text(name: impl Into<String>, xml_id: impl Into<String>) -> Self {
        Self::new(TextControlKind::Text, name, xml_id)
    }

    pub fn textarea(name: impl Into<String>, xml_id: impl Into<String>) -> Self {
        Self::new(TextControlKind::Textarea, name, xml_id)
    }

    fn new(kind: TextControlKind, name: impl Into<String>, xml_id: impl Into<String>) -> Self {
        Self {
            kind,
            name: name.into(),
            xml_id: xml_id.into(),
            value: None,
            current_value: None,
            disabled: None,
            readonly: None,
            printable: None,
            max_length: None,
            tab_index: None,
            tab_stop: None,
            title: None,
            convert_empty_to_null: None,
            data_field: None,
            linked_cell: None,
            paragraphs: Vec::new(),
        }
    }

    pub fn to_xml_fragment(&self) -> Result<String> {
        control_xml(self)
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ControlForm {
    pub name: String,
    pub controls: Vec<TextControl>,
    pub apply_filter: Option<bool>,
    pub command_type: Option<String>,
    pub command: Option<String>,
    pub datasource: Option<String>,
    pub target_frame: Option<String>,
    pub href: Option<String>,
}

impl ControlForm {
    pub fn new(name: impl Into<String>) -> Self {
        Self {
            name: name.into(),
            controls: Vec::new(),
            apply_filter: None,
            command_type: None,
            command: None,
            datasource: None,
            target_frame: None,
            href: None,
        }
    }

    pub fn add_control(&mut self, control: TextControl) -> Result<()> {
        validate_control(&control)?;
        if self
            .controls
            .iter()
            .any(|existing| existing.name == control.name)
        {
            return invalid(format!("duplicate form control name '{}'", control.name));
        }
        if self
            .controls
            .iter()
            .any(|existing| existing.xml_id == control.xml_id)
        {
            return invalid(format!(
                "duplicate form control xml:id '{}'",
                control.xml_id
            ));
        }
        self.controls.push(control);
        Ok(())
    }

    pub fn to_xml_fragment(&self) -> Result<String> {
        validate_name("form name", &self.name)?;
        validate_optional_string("form command", self.command.as_deref())?;
        validate_optional_string("form datasource", self.datasource.as_deref())?;
        validate_optional_string("form target frame", self.target_frame.as_deref())?;
        validate_resource(self.href.as_deref())?;
        if let Some(command_type) = &self.command_type
            && !matches!(command_type.as_str(), "table" | "query" | "command")
        {
            return invalid(format!("invalid form command type '{command_type}'"));
        }
        let mut result = format!(
            r#"<form:form xmlns:form="{FORM}" xmlns:office="{OFFICE}" xmlns:text="{TEXT}" xmlns:xlink="{XLINK}" form:name="{}""#,
            escape(&self.name)
        );
        push_bool(&mut result, "form:apply-filter", self.apply_filter);
        push_string(
            &mut result,
            "form:command-type",
            self.command_type.as_deref(),
        );
        push_string(&mut result, "form:command", self.command.as_deref());
        push_string(&mut result, "form:datasource", self.datasource.as_deref());
        push_string(
            &mut result,
            "office:target-frame",
            self.target_frame.as_deref(),
        );
        if let Some(href) = &self.href {
            result.push_str(r#" xlink:type="simple" xlink:actuate="onRequest""#);
            push_string(&mut result, "xlink:href", Some(href));
        }
        if self.controls.is_empty() {
            result.push_str("/>");
            return Ok(result);
        }
        result.push('>');
        for control in &self.controls {
            result.push_str(&control_xml(control)?);
        }
        result.push_str("</form:form>");
        Ok(result)
    }
}

pub fn text_controls(xml: &str) -> Result<Vec<TextControl>> {
    Ok(scan(xml)?
        .controls
        .into_iter()
        .map(|entry| entry.control)
        .collect())
}

pub fn insert_text_control_xml(
    xml: &str,
    form_index: usize,
    control: &TextControl,
) -> Result<String> {
    validate_control(control)?;
    let scan = scan(xml)?;
    let form = scan
        .forms
        .get(form_index)
        .ok_or_else(|| Error::InvalidFormat(format!("form {form_index} is out of bounds")))?;
    reject_duplicate(form, control, None)?;
    let fragment = bind_fragment(xml, control_xml(control)?);
    match &form.site {
        Site::Paired { close_start, .. } => apply(xml, (*close_start)..(*close_start), &fragment),
        Site::Empty { start, end, qname } => expand_empty(xml, *start, *end, qname, &fragment),
    }
}

pub fn replace_text_control_xml(
    xml: &str,
    control_index: usize,
    replacement: &TextControl,
) -> Result<String> {
    validate_control(replacement)?;
    let scan = scan(xml)?;
    let current = scan.controls.get(control_index).ok_or_else(|| {
        Error::InvalidFormat(format!("text control {control_index} is out of bounds"))
    })?;
    reject_duplicate(
        &scan.forms[current.form],
        replacement,
        Some(&current.control),
    )?;
    apply(
        xml,
        current.span.clone(),
        &bind_fragment(xml, control_xml(replacement)?),
    )
}

pub fn remove_text_control_xml(xml: &str, control_index: usize) -> Result<String> {
    let scan = scan(xml)?;
    let current = scan.controls.get(control_index).ok_or_else(|| {
        Error::InvalidFormat(format!("text control {control_index} is out of bounds"))
    })?;
    apply(xml, current.span.clone(), "")
}

fn control_xml(control: &TextControl) -> Result<String> {
    validate_control(control)?;
    let tag = match control.kind {
        TextControlKind::Text => "form:text",
        TextControlKind::Textarea => "form:textarea",
    };
    let mut result = format!(
        r#"<{tag} form:name="{}" xml:id="{}""#,
        escape(&control.name),
        escape(&control.xml_id)
    );
    push_string(&mut result, "form:value", control.value.as_deref());
    push_string(
        &mut result,
        "form:current-value",
        control.current_value.as_deref(),
    );
    push_bool(&mut result, "form:disabled", control.disabled);
    push_bool(&mut result, "form:readonly", control.readonly);
    push_bool(&mut result, "form:printable", control.printable);
    push_u64(&mut result, "form:max-length", control.max_length);
    push_u64(&mut result, "form:tab-index", control.tab_index);
    push_bool(&mut result, "form:tab-stop", control.tab_stop);
    push_string(&mut result, "form:title", control.title.as_deref());
    push_bool(
        &mut result,
        "form:convert-empty-to-null",
        control.convert_empty_to_null,
    );
    push_string(
        &mut result,
        "form:data-field",
        control.data_field.as_deref(),
    );
    push_string(
        &mut result,
        "form:linked-cell",
        control.linked_cell.as_deref(),
    );
    if control.paragraphs.is_empty() {
        result.push_str("/>");
    } else {
        result.push('>');
        for paragraph in &control.paragraphs {
            result.push_str(&format!("<text:p>{}</text:p>", escape(paragraph)));
        }
        result.push_str(&format!("</{tag}>"));
    }
    Ok(result)
}

fn validate_control(control: &TextControl) -> Result<()> {
    validate_name("form control name", &control.name)?;
    validate_xml_id(&control.xml_id)?;
    for (label, value) in [
        ("form value", control.value.as_deref()),
        ("form current value", control.current_value.as_deref()),
        ("form title", control.title.as_deref()),
        ("form data field", control.data_field.as_deref()),
        ("form linked cell", control.linked_cell.as_deref()),
    ] {
        validate_optional_string(label, value)?;
    }
    if control.kind == TextControlKind::Text && !control.paragraphs.is_empty() {
        return invalid("form:text cannot contain text:p paragraphs");
    }
    let mut aggregate = 0usize;
    for paragraph in &control.paragraphs {
        validate_string("textarea paragraph", paragraph)?;
        aggregate = aggregate
            .checked_add(paragraph.len())
            .ok_or_else(|| Error::InvalidFormat("textarea content is too large".to_string()))?;
        if aggregate > MAX_STRING {
            return invalid("textarea content exceeds 1 MiB");
        }
    }
    Ok(())
}

fn reject_duplicate(
    form: &FormLocation,
    replacement: &TextControl,
    current: Option<&TextControl>,
) -> Result<()> {
    for index in &form.controls {
        let existing = &index.control;
        if current.is_some_and(|value| value.xml_id == existing.xml_id) {
            continue;
        }
        if existing.name == replacement.name {
            return invalid(format!(
                "duplicate form control name '{}'",
                replacement.name
            ));
        }
        if existing.xml_id == replacement.xml_id {
            return invalid(format!(
                "duplicate form control xml:id '{}'",
                replacement.xml_id
            ));
        }
    }
    Ok(())
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
struct FormLocation {
    site: Site,
    controls: Vec<ControlLocation>,
}
#[derive(Clone)]
struct ControlLocation {
    span: Range<usize>,
    form: usize,
    control: TextControl,
}
struct Scan {
    forms: Vec<FormLocation>,
    controls: Vec<ControlLocation>,
}
struct Open {
    local: Vec<u8>,
    form: Option<usize>,
    control: Option<usize>,
    paragraph: bool,
}
struct Attr {
    namespace: Option<String>,
    local: String,
    value: String,
}

fn scan(xml: &str) -> Result<Scan> {
    if xml.len() > MAX_XML {
        return invalid("form control XML exceeds 64 MiB");
    }
    let mut reader = NsReader::from_str(xml);
    let mut buffer = Vec::new();
    let mut previous = 0usize;
    let mut stack = Vec::<Open>::new();
    let mut form_stack = Vec::<usize>::new();
    let mut forms = Vec::<FormLocation>::new();
    let mut controls = Vec::<ControlLocation>::new();
    let mut xml_ids = Vec::<String>::new();
    let mut paragraph_text = String::new();
    loop {
        let (namespace, event) = reader
            .read_resolved_event_into(&mut buffer)
            .map_err(|error| Error::InvalidFormat(format!("invalid form control XML: {error}")))?;
        let namespace_uri = match namespace {
            ResolveResult::Bound(uri) => Some(String::from_utf8_lossy(uri.as_ref()).into_owned()),
            _ => None,
        };
        let end = reader.buffer_position() as usize;
        match event {
            Event::Start(ref element) => {
                if stack.len() >= 128 {
                    return invalid("form control XML nesting exceeds 128 levels");
                }
                let local = element.local_name().as_ref().to_vec();
                track_xml_id(&reader, element, &mut xml_ids)?;
                let mut form = None;
                let mut control = None;
                let mut paragraph = false;
                if namespace_uri.as_deref() == Some(FORM) && local == b"form" {
                    validate_form_element(&reader, element)?;
                    if forms.len() >= MAX_FORMS {
                        return invalid("too many form elements");
                    }
                    form = Some(forms.len());
                    forms.push(FormLocation {
                        site: Site::Paired { close_start: 0 },
                        controls: Vec::new(),
                    });
                    form_stack.push(form.unwrap());
                } else if namespace_uri.as_deref() == Some(FORM)
                    && matches!(local.as_slice(), b"text" | b"textarea")
                {
                    let form_index = *form_stack.last().ok_or_else(|| {
                        Error::InvalidFormat("text control has no form owner".to_string())
                    })?;
                    let parsed = parse_control(&reader, element, local.as_slice())?;
                    if controls.len() >= MAX_CONTROLS {
                        return invalid("too many text controls");
                    }
                    control = Some(controls.len());
                    controls.push(ControlLocation {
                        span: previous..0,
                        form: form_index,
                        control: parsed,
                    });
                } else if !form_stack.is_empty()
                    && ((namespace_uri.as_deref() == Some(OFFICE) && local == b"event-listeners")
                        || (namespace_uri.as_deref()
                            == Some("urn:oasis:names:tc:opendocument:xmlns:script:1.0")
                            && local == b"event-listener"))
                {
                    return invalid(
                        "event and macro content is outside the text-control mutation API",
                    );
                } else if namespace_uri.as_deref() == Some(TEXT)
                    && local == b"p"
                    && let Some(control_index) = stack.iter().rev().find_map(|open| open.control)
                {
                    if controls[control_index].control.kind != TextControlKind::Textarea {
                        return invalid("text:p is only valid in form:textarea");
                    }
                    paragraph = true;
                    paragraph_text.clear();
                }
                stack.push(Open {
                    local,
                    form,
                    control,
                    paragraph,
                });
            },
            Event::Empty(ref element) => {
                let local = element.local_name().as_ref().to_vec();
                track_xml_id(&reader, element, &mut xml_ids)?;
                if namespace_uri.as_deref() == Some(FORM) && local == b"form" {
                    validate_form_element(&reader, element)?;
                    if forms.len() >= MAX_FORMS {
                        return invalid("too many form elements");
                    }
                    forms.push(FormLocation {
                        site: Site::Empty {
                            start: previous,
                            end,
                            qname: qname(element)?,
                        },
                        controls: Vec::new(),
                    });
                } else if namespace_uri.as_deref() == Some(FORM)
                    && matches!(local.as_slice(), b"text" | b"textarea")
                {
                    let form_index = *form_stack.last().ok_or_else(|| {
                        Error::InvalidFormat("text control has no form owner".to_string())
                    })?;
                    let parsed = parse_control(&reader, element, local.as_slice())?;
                    if controls.len() >= MAX_CONTROLS {
                        return invalid("too many text controls");
                    }
                    controls.push(ControlLocation {
                        span: previous..end,
                        form: form_index,
                        control: parsed,
                    });
                } else if !form_stack.is_empty()
                    && ((namespace_uri.as_deref() == Some(OFFICE) && local == b"event-listeners")
                        || (namespace_uri.as_deref()
                            == Some("urn:oasis:names:tc:opendocument:xmlns:script:1.0")
                            && local == b"event-listener"))
                {
                    return invalid(
                        "event and macro content is outside the text-control mutation API",
                    );
                } else if namespace_uri.as_deref() == Some(TEXT)
                    && local == b"p"
                    && let Some(control_index) = stack.iter().rev().find_map(|open| open.control)
                {
                    if controls[control_index].control.kind != TextControlKind::Textarea {
                        return invalid("text:p is only valid in form:textarea");
                    }
                    controls[control_index]
                        .control
                        .paragraphs
                        .push(String::new());
                }
            },
            Event::Text(text) if stack.last().is_some_and(|open| open.paragraph) => {
                paragraph_text.push_str(&text.decode().map_err(|error| {
                    Error::InvalidFormat(format!("invalid textarea text: {error}"))
                })?);
                if paragraph_text.len() > MAX_STRING {
                    return invalid("textarea paragraph exceeds 1 MiB");
                }
            },
            Event::End(ref element) => {
                let open = stack.pop().ok_or_else(|| {
                    Error::InvalidFormat("form control XML stack underflow".to_string())
                })?;
                if open.local.as_slice() != element.local_name().as_ref() {
                    return invalid("mismatched form control XML elements");
                }
                if open.paragraph {
                    let control_index = stack
                        .iter()
                        .rev()
                        .find_map(|parent| parent.control)
                        .ok_or_else(|| {
                            Error::InvalidFormat("textarea paragraph has no owner".to_string())
                        })?;
                    controls[control_index]
                        .control
                        .paragraphs
                        .push(std::mem::take(&mut paragraph_text));
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
            Event::DocType(_) => return invalid("DOCTYPE is not allowed in form control XML"),
            Event::Eof => break,
            _ => {},
        }
        previous = end;
        buffer.clear();
    }
    if !stack.is_empty() {
        return invalid("unclosed form control XML elements");
    }
    for (index, control) in controls.iter().enumerate() {
        forms[control.form].controls.push(control.clone());
        let form = &forms[control.form];
        if form.controls[..form.controls.len() - 1]
            .iter()
            .any(|existing| existing.control.name == control.control.name)
        {
            return invalid(format!(
                "duplicate form control name '{}'",
                control.control.name
            ));
        }
        let _ = index;
    }
    Ok(Scan { forms, controls })
}

fn parse_control(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
    local: &[u8],
) -> Result<TextControl> {
    let attrs = attributes(reader, element)?;
    validate_allowed(&attrs, CONTROL_ATTRS)?;
    let name = required(&attrs, FORM, "name")?;
    let xml_id = required(&attrs, XML, "id")?;
    validate_name("form control name", &name)?;
    validate_xml_id(&xml_id)?;
    let mut control = TextControl::new(
        if local == b"text" {
            TextControlKind::Text
        } else {
            TextControlKind::Textarea
        },
        name,
        xml_id,
    );
    control.value = optional(&attrs, FORM, "value");
    control.current_value = optional(&attrs, FORM, "current-value");
    control.disabled = optional_bool(&attrs, FORM, "disabled")?;
    control.readonly = optional_bool(&attrs, FORM, "readonly")?;
    control.printable = optional_bool(&attrs, FORM, "printable")?;
    control.max_length = optional_u64(&attrs, FORM, "max-length")?;
    control.tab_index = optional_u64(&attrs, FORM, "tab-index")?;
    control.tab_stop = optional_bool(&attrs, FORM, "tab-stop")?;
    control.title = optional(&attrs, FORM, "title");
    control.convert_empty_to_null = optional_bool(&attrs, FORM, "convert-empty-to-null")?;
    control.data_field = optional(&attrs, FORM, "data-field");
    control.linked_cell = optional(&attrs, FORM, "linked-cell");
    validate_control(&control)?;
    Ok(control)
}

fn track_xml_id(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
    xml_ids: &mut Vec<String>,
) -> Result<()> {
    if let Some(xml_id) = optional(&attributes(reader, element)?, XML, "id") {
        validate_xml_id(&xml_id)?;
        if xml_ids.iter().any(|existing| existing == &xml_id) {
            return invalid(format!("duplicate xml:id '{xml_id}'"));
        }
        xml_ids.push(xml_id);
    }
    Ok(())
}

fn validate_form_element(reader: &NsReader<&[u8]>, element: &BytesStart<'_>) -> Result<()> {
    let attrs = attributes(reader, element)?;
    validate_allowed(&attrs, FORM_ATTRS)?;
    validate_name("form name", &required(&attrs, FORM, "name")?)?;
    if let Some(value) = optional(&attrs, FORM, "command-type")
        && !matches!(value.as_str(), "table" | "query" | "command")
    {
        return invalid(format!("invalid form command type '{value}'"));
    }
    for name in [
        "allow-deletes",
        "allow-inserts",
        "allow-updates",
        "apply-filter",
        "escape-processing",
        "ignore-result",
    ] {
        let _ = optional_bool(&attrs, FORM, name)?;
    }
    if let Some(value) = optional(&attrs, XLINK, "type")
        && value != "simple"
    {
        return invalid("form xlink:type must be simple");
    }
    if let Some(value) = optional(&attrs, XLINK, "actuate")
        && value != "onRequest"
    {
        return invalid("form xlink:actuate must be onRequest");
    }
    validate_resource(optional(&attrs, XLINK, "href").as_deref())
}

const CONTROL_ATTRS: &[(&str, &str)] = &[
    (FORM, "name"),
    (XML, "id"),
    (FORM, "id"),
    (FORM, "control-implementation"),
    (FORM, "value"),
    (FORM, "current-value"),
    (FORM, "disabled"),
    (FORM, "max-length"),
    (FORM, "printable"),
    (FORM, "readonly"),
    (FORM, "tab-index"),
    (FORM, "tab-stop"),
    (FORM, "title"),
    (FORM, "convert-empty-to-null"),
    (FORM, "data-field"),
    (FORM, "linked-cell"),
    (FORM, "input-required"),
    (FORM, "text-style-name"),
    ("http://www.w3.org/2002/xforms", "bind"),
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

fn attributes(reader: &NsReader<&[u8]>, element: &BytesStart<'_>) -> Result<Vec<Attr>> {
    let mut result = Vec::new();
    for attribute in element.attributes().with_checks(true) {
        let attribute = attribute
            .map_err(|error| Error::InvalidFormat(format!("invalid form attribute: {error}")))?;
        let (namespace, local) = reader.resolver().resolve_attribute(attribute.key);
        let namespace = match namespace {
            ResolveResult::Bound(uri) => Some(String::from_utf8_lossy(uri.as_ref()).into_owned()),
            _ => None,
        };
        if namespace.as_deref() == Some("http://www.w3.org/2000/xmlns/")
            || attribute.key.as_ref() == b"xmlns"
        {
            continue;
        }
        let value = attribute
            .decoded_and_normalized_value(quick_xml::XmlVersion::Implicit1_0, reader.decoder())
            .map_err(|error| {
                Error::InvalidFormat(format!("invalid form attribute value: {error}"))
            })?
            .into_owned();
        result.push(Attr {
            namespace,
            local: String::from_utf8_lossy(local.as_ref()).into_owned(),
            value,
        });
    }
    Ok(result)
}

fn validate_allowed(attrs: &[Attr], allowed: &[(&str, &str)]) -> Result<()> {
    for attr in attrs {
        if !allowed.iter().any(|(namespace, local)| {
            attr.namespace.as_deref() == Some(*namespace) && attr.local == *local
        }) {
            return invalid(format!(
                "unsupported form attribute '{}:{}'",
                attr.namespace.as_deref().unwrap_or(""),
                attr.local
            ));
        }
    }
    Ok(())
}

fn required(attrs: &[Attr], namespace: &str, local: &str) -> Result<String> {
    optional(attrs, namespace, local)
        .ok_or_else(|| Error::InvalidFormat(format!("missing required attribute {local}")))
}
fn optional(attrs: &[Attr], namespace: &str, local: &str) -> Option<String> {
    attrs
        .iter()
        .find(|attr| attr.namespace.as_deref() == Some(namespace) && attr.local == local)
        .map(|attr| attr.value.clone())
}
fn optional_bool(attrs: &[Attr], namespace: &str, local: &str) -> Result<Option<bool>> {
    optional(attrs, namespace, local)
        .map(|value| match value.as_str() {
            "true" | "1" => Ok(true),
            "false" | "0" => Ok(false),
            _ => invalid(format!("invalid boolean value '{value}' for {local}")),
        })
        .transpose()
}
fn optional_u64(attrs: &[Attr], namespace: &str, local: &str) -> Result<Option<u64>> {
    optional(attrs, namespace, local)
        .map(|value| {
            value.parse::<u64>().map_err(|_| {
                Error::InvalidFormat(format!(
                    "invalid non-negative integer '{value}' for {local}"
                ))
            })
        })
        .transpose()
}

fn validate_name(label: &str, value: &str) -> Result<()> {
    validate_string(label, value)?;
    if value.is_empty() {
        invalid(format!("{label} cannot be empty"))
    } else {
        Ok(())
    }
}
fn validate_xml_id(value: &str) -> Result<()> {
    validate_name("form control xml:id", value)?;
    let mut chars = value.chars();
    let first = chars.next().unwrap();
    if !(first == '_' || first.is_ascii_alphabetic())
        || !chars.all(|ch| ch == '_' || ch == '-' || ch == '.' || ch.is_ascii_alphanumeric())
    {
        return invalid(format!("invalid form control xml:id '{value}'"));
    }
    Ok(())
}
fn validate_optional_string(label: &str, value: Option<&str>) -> Result<()> {
    if let Some(value) = value {
        validate_string(label, value)?;
    }
    Ok(())
}
fn validate_string(label: &str, value: &str) -> Result<()> {
    if value.len() > MAX_STRING {
        return invalid(format!("{label} exceeds 1 MiB"));
    }
    if value
        .chars()
        .any(|ch| matches!(ch as u32, 0..=8 | 11 | 12 | 14..=31))
    {
        return invalid(format!("{label} contains invalid XML characters"));
    }
    Ok(())
}
fn validate_resource(value: Option<&str>) -> Result<()> {
    if let Some(value) = value {
        if value.len() > MAX_RESOURCE {
            return invalid("form resource reference exceeds 8 KiB");
        }
        validate_string("form resource reference", value)?;
        let lower = value.trim().to_ascii_lowercase();
        if lower.starts_with("javascript:") || lower.starts_with("data:") {
            return invalid("active form resource scheme is not allowed");
        }
    }
    Ok(())
}

fn push_string(output: &mut String, name: &str, value: Option<&str>) {
    if let Some(value) = value {
        output.push_str(&format!(r#" {name}="{}""#, escape(value)));
    }
}
fn push_bool(output: &mut String, name: &str, value: Option<bool>) {
    if let Some(value) = value {
        output.push_str(&format!(
            r#" {name}="{}""#,
            if value { "true" } else { "false" }
        ));
    }
}
fn push_u64(output: &mut String, name: &str, value: Option<u64>) {
    if let Some(value) = value {
        output.push_str(&format!(r#" {name}="{value}""#));
    }
}
fn escape(value: &str) -> String {
    value
        .replace('&', "&amp;")
        .replace('<', "&lt;")
        .replace('"', "&quot;")
        .replace('\'', "&apos;")
}
fn invalid<T>(message: impl Into<String>) -> Result<T> {
    Err(Error::InvalidFormat(message.into()))
}
fn qname(element: &BytesStart<'_>) -> Result<String> {
    String::from_utf8(element.name().as_ref().to_vec())
        .map_err(|_| Error::InvalidFormat("invalid form element name".to_string()))
}
fn apply(xml: &str, span: Range<usize>, replacement: &str) -> Result<String> {
    let mut result = String::with_capacity(xml.len() - span.len() + replacement.len());
    result.push_str(&xml[..span.start]);
    result.push_str(replacement);
    result.push_str(&xml[span.end..]);
    Ok(result)
}
fn expand_empty(xml: &str, start: usize, end: usize, qname: &str, content: &str) -> Result<String> {
    let raw = &xml[start..end];
    let slash = raw
        .rfind("/>")
        .ok_or_else(|| Error::InvalidFormat("invalid empty form element".to_string()))?;
    let replacement = format!("{}>{content}</{qname}>", &raw[..slash]);
    apply(xml, start..end, &replacement)
}
fn bind_fragment(xml: &str, mut fragment: String) -> String {
    for (prefix, namespace) in [("form", FORM), ("text", TEXT)] {
        if fragment.contains(&format!("{prefix}:"))
            && !fragment.contains(&format!("xmlns:{prefix}="))
        {
            fragment = fragment.replacen(' ', &format!(" xmlns:{prefix}=\"{namespace}\" "), 1);
        }
    }
    let _ = xml;
    fragment
}

#[cfg(test)]
mod tests {
    use super::*;

    const ROOT: &str = r#"<o:document-content xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:f="urn:oasis:names:tc:opendocument:xmlns:form:1.0" xmlns:t="urn:oasis:names:tc:opendocument:xmlns:text:1.0" xmlns:x="http://www.w3.org/1999/xlink"><o:body><o:text><o:forms>"#;
    const END: &str = "</o:forms></o:text></o:body></o:document-content>";

    #[test]
    fn canonical_text_and_textarea_are_strict_and_escaped() {
        let mut text = TextControl::text("Search & replace", "search_1");
        text.value = Some("<query>".to_string());
        text.readonly = Some(false);
        text.max_length = Some(40);
        assert_eq!(
            text.to_xml_fragment().unwrap(),
            r#"<form:text form:name="Search &amp; replace" xml:id="search_1" form:value="&lt;query>" form:readonly="false" form:max-length="40"/>"#
        );
        let mut area = TextControl::textarea("Notes", "notes_1");
        area.paragraphs = vec!["one".into(), "two & three".into()];
        let mut form = ControlForm::new("Main");
        form.apply_filter = Some(true);
        form.add_control(text).unwrap();
        form.add_control(area).unwrap();
        let xml = form.to_xml_fragment().unwrap();
        assert!(xml.contains("<text:p>two &amp; three</text:p>"));
        assert_eq!(
            text_controls(&format!("{ROOT}{xml}{END}")).unwrap().len(),
            2
        );
    }

    #[test]
    fn aliases_and_lossless_insert_replace_remove_preserve_surroundings() {
        let xml = format!(
            r#"{ROOT}<f:form f:name="Main"><f:text f:name="A" xml:id="a" f:value="old"><f:properties><f:property f:property-name="Vendor" o:value-type="string" o:string-value="keep"/></f:properties></f:text><!--keep--></f:form>{END}"#
        );
        let mut area = TextControl::textarea("B", "b");
        area.current_value = Some("new".into());
        let inserted = insert_text_control_xml(&xml, 0, &area).unwrap();
        assert!(inserted.contains("<!--keep-->") && inserted.contains("form:textarea"));
        let mut replacement = TextControl::text("A2", "a2");
        replacement.value = Some("replaced".into());
        let replaced = replace_text_control_xml(&inserted, 0, &replacement).unwrap();
        assert!(replaced.contains("<!--keep-->") && !replaced.contains("Vendor"));
        let removed = remove_text_control_xml(&replaced, 1).unwrap();
        assert_eq!(text_controls(&removed).unwrap().len(), 1);
        let empty = format!(r#"{ROOT}<f:form f:name="Empty"/>{END}"#);
        assert!(
            insert_text_control_xml(&empty, 0, &TextControl::text("T", "t"))
                .unwrap()
                .contains("</f:form>")
        );
    }

    #[test]
    fn rejects_wrong_namespaces_hostile_attributes_resources_and_limits() {
        assert!(TextControl::text("T", "1bad").to_xml_fragment().is_err());
        let wrong =
            format!(r#"{ROOT}<f:form f:name="Main"><x:text f:name="T" xml:id="t"/></f:form>{END}"#);
        assert!(text_controls(&wrong).unwrap().is_empty());
        let hostile = format!(
            r#"{ROOT}<f:form f:name="Main"><f:text f:name="T" xml:id="t" x:href="https://evil.invalid"/></f:form>{END}"#
        );
        assert!(text_controls(&hostile).is_err());
        let duplicate_id = format!(
            r#"{ROOT}<t:p xml:id="t"/><f:form f:name="Main"><f:text f:name="T" xml:id="t"/></f:form>{END}"#
        );
        assert!(text_controls(&duplicate_id).is_err());
        let events = format!(r#"{ROOT}<f:form f:name="Main"><o:event-listeners/></f:form>{END}"#);
        assert!(text_controls(&events).is_err());
        let mut form = ControlForm::new("Main");
        form.href = Some("javascript:alert(1)".into());
        assert!(form.to_xml_fragment().is_err());
        let mut huge = TextControl::textarea("T", "t");
        huge.paragraphs.push("x".repeat(MAX_STRING + 1));
        assert!(huge.to_xml_fragment().is_err());
    }

    #[test]
    fn libreoffice_odfpy_and_odfdo_shapes_parse_without_resource_resolution() {
        let lo = include_str!(
            "../../../../test-data/libreoffice-core/vcl/qa/cppunit/pdfexport/data/PDF_export_with_formcontrol.fodt"
        );
        let controls = text_controls(lo).unwrap();
        assert!(
            controls
                .iter()
                .any(|control| control.kind == TextControlKind::Textarea)
        );
        let producer = format!(
            r#"{ROOT}<f:form f:name="odfpy"><f:text f:name="Text" xml:id="odfpy_text" f:value="value"/><f:textarea f:name="Textarea" xml:id="odfdo_textarea" f:current-value="current"><t:p>body</t:p></f:textarea></f:form>{END}"#
        );
        let parsed = text_controls(&producer).unwrap();
        assert_eq!(parsed.len(), 2);
        assert_eq!(parsed[1].paragraphs, ["body"]);
    }

    #[test]
    fn builder_and_mutable_document_round_trip_controls() {
        use crate::{Builder, Document, mutable::MutableDocument};

        let mut form = ControlForm::new("Main");
        form.add_control(TextControl::text("Query", "query_1"))
            .unwrap();
        let mut builder = Builder::new();
        builder.add_control_form(&form).unwrap();
        builder.add_paragraph("body").unwrap();
        let document = Document::from_bytes(builder.build().unwrap()).unwrap();
        let mut mutable = MutableDocument::from_document(document).unwrap();
        assert_eq!(mutable.text_controls().unwrap().len(), 1);

        let mut area = TextControl::textarea("Notes", "notes_1");
        area.paragraphs.push("inert text".into());
        mutable.insert_text_control(0, &area).unwrap();
        let mut replacement = TextControl::text("Search", "search_1");
        replacement.current_value = Some("term".into());
        let old = mutable.replace_text_control(0, &replacement).unwrap();
        assert_eq!(old.name, "Query");
        assert_eq!(mutable.remove_text_control(1).unwrap().name, "Notes");
        assert_eq!(mutable.text_controls().unwrap(), [replacement]);
    }
}
