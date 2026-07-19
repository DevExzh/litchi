use std::ops::Range;

use litchi_core::{Error, Result};
use quick_xml::NsReader;
use quick_xml::events::{BytesStart, Event};
use quick_xml::name::ResolveResult;

const FORM: &str = "urn:oasis:names:tc:opendocument:xmlns:form:1.0";
const OFFICE: &str = "urn:oasis:names:tc:opendocument:xmlns:office:1.0";
const XLINK: &str = "http://www.w3.org/1999/xlink";
const XML: &str = "http://www.w3.org/XML/1998/namespace";
const SCRIPT: &str = "urn:oasis:names:tc:opendocument:xmlns:script:1.0";
const XFORMS: &str = "http://www.w3.org/2002/xforms";
const MAX_XML: usize = 64 * 1024 * 1024;
const MAX_STRING: usize = 1_048_576;
const MAX_RESOURCE: usize = 8_192;
const MAX_FORMS: usize = 16_384;
const MAX_CONTROLS: usize = 65_536;

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum OdfButtonType {
    Submit,
    Reset,
    Push,
    Url,
}
impl OdfButtonType {
    fn token(self) -> &'static str {
        match self {
            Self::Submit => "submit",
            Self::Reset => "reset",
            Self::Push => "push",
            Self::Url => "url",
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum OdfCheckboxState {
    Unchecked,
    Checked,
    Unknown,
}
impl OdfCheckboxState {
    fn token(self) -> &'static str {
        match self {
            Self::Unchecked => "unchecked",
            Self::Checked => "checked",
            Self::Unknown => "unknown",
        }
    }
    fn parse(value: &str) -> Result<Self> {
        match value {
            "unchecked" => Ok(Self::Unchecked),
            "checked" => Ok(Self::Checked),
            "unknown" => Ok(Self::Unknown),
            _ => invalid(format!("invalid checkbox state '{value}'")),
        }
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct OdfButtonControl {
    pub name: String,
    pub xml_id: String,
    pub label: Option<String>,
    pub value: Option<String>,
    pub button_type: Option<OdfButtonType>,
    pub disabled: Option<bool>,
    pub printable: Option<bool>,
    pub tab_index: Option<u64>,
    pub tab_stop: Option<bool>,
    pub title: Option<String>,
    pub image_data: Option<String>,
    pub target_frame: Option<String>,
    pub href: Option<String>,
    pub image_align: Option<String>,
    pub image_position: Option<String>,
    pub repeat: Option<bool>,
    pub delay_for_repeat: Option<String>,
    pub default_button: Option<bool>,
    pub toggle: Option<bool>,
    pub focus_on_click: Option<bool>,
}

impl OdfButtonControl {
    pub fn new(name: impl Into<String>, xml_id: impl Into<String>) -> Self {
        Self {
            name: name.into(),
            xml_id: xml_id.into(),
            label: None,
            value: None,
            button_type: None,
            disabled: None,
            printable: None,
            tab_index: None,
            tab_stop: None,
            title: None,
            image_data: None,
            target_frame: None,
            href: None,
            image_align: None,
            image_position: None,
            repeat: None,
            delay_for_repeat: None,
            default_button: None,
            toggle: None,
            focus_on_click: None,
        }
    }
    pub fn to_xml_fragment(&self) -> Result<String> {
        button_xml(self)
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct OdfCheckboxControl {
    pub name: String,
    pub xml_id: String,
    pub label: Option<String>,
    pub value: Option<String>,
    pub disabled: Option<bool>,
    pub printable: Option<bool>,
    pub tab_index: Option<u64>,
    pub tab_stop: Option<bool>,
    pub title: Option<String>,
    pub data_field: Option<String>,
    pub linked_cell: Option<String>,
    pub state: Option<OdfCheckboxState>,
    pub current_state: Option<OdfCheckboxState>,
    pub is_tristate: Option<bool>,
    pub visual_effect: Option<String>,
    pub image_align: Option<String>,
    pub image_position: Option<String>,
}

impl OdfCheckboxControl {
    pub fn new(name: impl Into<String>, xml_id: impl Into<String>) -> Self {
        Self {
            name: name.into(),
            xml_id: xml_id.into(),
            label: None,
            value: None,
            disabled: None,
            printable: None,
            tab_index: None,
            tab_stop: None,
            title: None,
            data_field: None,
            linked_cell: None,
            state: None,
            current_state: None,
            is_tristate: None,
            visual_effect: None,
            image_align: None,
            image_position: None,
        }
    }
    pub fn to_xml_fragment(&self) -> Result<String> {
        checkbox_xml(self)
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub enum OdfInteractiveControl {
    Button(OdfButtonControl),
    Checkbox(OdfCheckboxControl),
}
impl From<OdfButtonControl> for OdfInteractiveControl {
    fn from(value: OdfButtonControl) -> Self {
        Self::Button(value)
    }
}
impl From<OdfCheckboxControl> for OdfInteractiveControl {
    fn from(value: OdfCheckboxControl) -> Self {
        Self::Checkbox(value)
    }
}
impl OdfInteractiveControl {
    pub fn name(&self) -> &str {
        match self {
            Self::Button(value) => &value.name,
            Self::Checkbox(value) => &value.name,
        }
    }
    pub fn xml_id(&self) -> &str {
        match self {
            Self::Button(value) => &value.xml_id,
            Self::Checkbox(value) => &value.xml_id,
        }
    }
    pub fn to_xml_fragment(&self) -> Result<String> {
        match self {
            Self::Button(value) => button_xml(value),
            Self::Checkbox(value) => checkbox_xml(value),
        }
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct OdfInteractiveForm {
    pub name: String,
    pub controls: Vec<OdfInteractiveControl>,
    pub apply_filter: Option<bool>,
}
impl OdfInteractiveForm {
    pub fn new(name: impl Into<String>) -> Self {
        Self {
            name: name.into(),
            controls: Vec::new(),
            apply_filter: None,
        }
    }
    pub fn add_control(&mut self, control: impl Into<OdfInteractiveControl>) -> Result<()> {
        let control = control.into();
        validate_control(&control)?;
        if self
            .controls
            .iter()
            .any(|existing| existing.name() == control.name())
        {
            return invalid(format!("duplicate form control name '{}'", control.name()));
        }
        if self
            .controls
            .iter()
            .any(|existing| existing.xml_id() == control.xml_id())
        {
            return invalid(format!(
                "duplicate form control xml:id '{}'",
                control.xml_id()
            ));
        }
        self.controls.push(control);
        Ok(())
    }
    pub fn to_xml_fragment(&self) -> Result<String> {
        validate_name("form name", &self.name)?;
        let mut result = format!(
            r#"<form:form xmlns:form="{FORM}" xmlns:office="{OFFICE}" xmlns:xlink="{XLINK}" form:name="{}""#,
            escape(&self.name)
        );
        push_bool(&mut result, "form:apply-filter", self.apply_filter);
        if self.controls.is_empty() {
            result.push_str("/>");
            return Ok(result);
        }
        result.push('>');
        for control in &self.controls {
            result.push_str(&control.to_xml_fragment()?);
        }
        result.push_str("</form:form>");
        Ok(result)
    }
}

pub fn interactive_controls(xml: &str) -> Result<Vec<OdfInteractiveControl>> {
    Ok(scan(xml)?
        .controls
        .into_iter()
        .map(|item| item.control)
        .collect())
}
pub fn insert_interactive_control_xml(
    xml: &str,
    form_index: usize,
    control: &OdfInteractiveControl,
) -> Result<String> {
    validate_control(control)?;
    let scan = scan(xml)?;
    let form = scan
        .forms
        .get(form_index)
        .ok_or_else(|| Error::InvalidFormat(format!("form {form_index} is out of bounds")))?;
    reject_duplicate(form, control, None)?;
    let fragment = bind_fragment(control.to_xml_fragment()?);
    match &form.site {
        Site::Paired { close_start } => apply(xml, (*close_start)..(*close_start), &fragment),
        Site::Empty { start, end, qname } => expand_empty(xml, *start, *end, qname, &fragment),
    }
}
pub fn replace_interactive_control_xml(
    xml: &str,
    control_index: usize,
    replacement: &OdfInteractiveControl,
) -> Result<String> {
    validate_control(replacement)?;
    let scan = scan(xml)?;
    let current = scan.controls.get(control_index).ok_or_else(|| {
        Error::InvalidFormat(format!(
            "interactive control {control_index} is out of bounds"
        ))
    })?;
    reject_duplicate(
        &scan.forms[current.form],
        replacement,
        Some(&current.control),
    )?;
    apply(
        xml,
        current.span.clone(),
        &bind_fragment(replacement.to_xml_fragment()?),
    )
}
pub fn remove_interactive_control_xml(xml: &str, control_index: usize) -> Result<String> {
    let scan = scan(xml)?;
    let current = scan.controls.get(control_index).ok_or_else(|| {
        Error::InvalidFormat(format!(
            "interactive control {control_index} is out of bounds"
        ))
    })?;
    apply(xml, current.span.clone(), "")
}

fn button_xml(value: &OdfButtonControl) -> Result<String> {
    validate_button(value)?;
    let mut output = format!(
        r#"<form:button form:name="{}" xml:id="{}""#,
        escape(&value.name),
        escape(&value.xml_id)
    );
    push_string(&mut output, "form:label", value.label.as_deref());
    push_string(&mut output, "form:value", value.value.as_deref());
    if let Some(kind) = value.button_type {
        push_string(&mut output, "form:button-type", Some(kind.token()));
    }
    push_bool(&mut output, "form:disabled", value.disabled);
    push_bool(&mut output, "form:printable", value.printable);
    push_u64(&mut output, "form:tab-index", value.tab_index);
    push_bool(&mut output, "form:tab-stop", value.tab_stop);
    push_string(&mut output, "form:title", value.title.as_deref());
    push_string(&mut output, "form:image-data", value.image_data.as_deref());
    push_string(
        &mut output,
        "office:target-frame",
        value.target_frame.as_deref(),
    );
    if let Some(href) = &value.href {
        push_string(&mut output, "xlink:type", Some("simple"));
        push_string(&mut output, "xlink:href", Some(href));
    }
    push_string(
        &mut output,
        "form:image-align",
        value.image_align.as_deref(),
    );
    push_string(
        &mut output,
        "form:image-position",
        value.image_position.as_deref(),
    );
    push_bool(&mut output, "form:repeat", value.repeat);
    push_string(
        &mut output,
        "form:delay-for-repeat",
        value.delay_for_repeat.as_deref(),
    );
    push_bool(&mut output, "form:default-button", value.default_button);
    push_bool(&mut output, "form:toggle", value.toggle);
    push_bool(&mut output, "form:focus-on-click", value.focus_on_click);
    output.push_str("/>");
    Ok(output)
}
fn checkbox_xml(value: &OdfCheckboxControl) -> Result<String> {
    validate_checkbox(value)?;
    let mut output = format!(
        r#"<form:checkbox form:name="{}" xml:id="{}""#,
        escape(&value.name),
        escape(&value.xml_id)
    );
    push_string(&mut output, "form:label", value.label.as_deref());
    push_string(&mut output, "form:value", value.value.as_deref());
    push_bool(&mut output, "form:disabled", value.disabled);
    push_bool(&mut output, "form:printable", value.printable);
    push_u64(&mut output, "form:tab-index", value.tab_index);
    push_bool(&mut output, "form:tab-stop", value.tab_stop);
    push_string(&mut output, "form:title", value.title.as_deref());
    push_string(&mut output, "form:data-field", value.data_field.as_deref());
    push_string(
        &mut output,
        "form:linked-cell",
        value.linked_cell.as_deref(),
    );
    if let Some(state) = value.state {
        push_string(&mut output, "form:state", Some(state.token()));
    }
    if let Some(state) = value.current_state {
        push_string(&mut output, "form:current-state", Some(state.token()));
    }
    push_bool(&mut output, "form:is-tristate", value.is_tristate);
    push_string(
        &mut output,
        "form:visual-effect",
        value.visual_effect.as_deref(),
    );
    push_string(
        &mut output,
        "form:image-align",
        value.image_align.as_deref(),
    );
    push_string(
        &mut output,
        "form:image-position",
        value.image_position.as_deref(),
    );
    output.push_str("/>");
    Ok(output)
}
fn validate_control(value: &OdfInteractiveControl) -> Result<()> {
    match value {
        OdfInteractiveControl::Button(value) => validate_button(value),
        OdfInteractiveControl::Checkbox(value) => validate_checkbox(value),
    }
}
fn validate_button(value: &OdfButtonControl) -> Result<()> {
    validate_identity(&value.name, &value.xml_id)?;
    for (label, item) in [
        ("button label", value.label.as_deref()),
        ("button value", value.value.as_deref()),
        ("button title", value.title.as_deref()),
        ("button target frame", value.target_frame.as_deref()),
    ] {
        validate_optional(label, item)?;
    }
    validate_resource("button image", value.image_data.as_deref(), true)?;
    validate_resource("button href", value.href.as_deref(), false)?;
    validate_token(
        "button image alignment",
        value.image_align.as_deref(),
        &["start", "center", "end"],
    )?;
    validate_token(
        "button image position",
        value.image_position.as_deref(),
        &[
            "center",
            "left",
            "right",
            "top",
            "bottom",
            "top-left",
            "top-right",
            "bottom-left",
            "bottom-right",
        ],
    )?;
    if let Some(duration) = &value.delay_for_repeat {
        validate_duration(duration)?;
    }
    Ok(())
}
fn validate_checkbox(value: &OdfCheckboxControl) -> Result<()> {
    validate_identity(&value.name, &value.xml_id)?;
    for (label, item) in [
        ("checkbox label", value.label.as_deref()),
        ("checkbox value", value.value.as_deref()),
        ("checkbox title", value.title.as_deref()),
        ("checkbox data field", value.data_field.as_deref()),
        ("checkbox linked cell", value.linked_cell.as_deref()),
    ] {
        validate_optional(label, item)?;
    }
    validate_token(
        "checkbox visual effect",
        value.visual_effect.as_deref(),
        &["none", "3d", "flat"],
    )?;
    validate_token(
        "checkbox image alignment",
        value.image_align.as_deref(),
        &["start", "center", "end"],
    )?;
    validate_token(
        "checkbox image position",
        value.image_position.as_deref(),
        &[
            "center",
            "left",
            "right",
            "top",
            "bottom",
            "top-left",
            "top-right",
            "bottom-left",
            "bottom-right",
        ],
    )?;
    if value.is_tristate == Some(false)
        && (matches!(value.state, Some(OdfCheckboxState::Unknown))
            || matches!(value.current_state, Some(OdfCheckboxState::Unknown)))
    {
        return invalid("non-tristate checkbox cannot have unknown state");
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
    names: Vec<String>,
    controls: Vec<ControlLocation>,
}
#[derive(Clone)]
struct ControlLocation {
    span: Range<usize>,
    form: usize,
    control: OdfInteractiveControl,
}
struct Scan {
    forms: Vec<FormLocation>,
    controls: Vec<ControlLocation>,
}
struct Open {
    local: Vec<u8>,
    form: Option<usize>,
    control: Option<usize>,
}
struct Attr {
    namespace: Option<String>,
    local: String,
    value: String,
}

fn scan(xml: &str) -> Result<Scan> {
    if xml.len() > MAX_XML {
        return invalid("interactive form XML exceeds 64 MiB");
    }
    let mut reader = NsReader::from_str(xml);
    let mut buffer = Vec::new();
    let mut previous = 0usize;
    let mut stack = Vec::<Open>::new();
    let mut form_stack = Vec::<usize>::new();
    let mut forms = Vec::<FormLocation>::new();
    let mut controls = Vec::<ControlLocation>::new();
    let mut xml_ids = Vec::<String>::new();
    loop {
        let (namespace, event) = reader
            .read_resolved_event_into(&mut buffer)
            .map_err(|error| {
                Error::InvalidFormat(format!("invalid interactive form XML: {error}"))
            })?;
        let namespace = match namespace {
            ResolveResult::Bound(uri) => Some(String::from_utf8_lossy(uri.as_ref()).into_owned()),
            _ => None,
        };
        let end = reader.buffer_position() as usize;
        match event {
            Event::Start(ref element) => {
                if stack.len() >= 128 {
                    return invalid("interactive form XML nesting exceeds 128 levels");
                }
                track_xml_id(&reader, element, &mut xml_ids)?;
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
                        names: Vec::new(),
                        controls: Vec::new(),
                    });
                    form_stack.push(form.unwrap());
                } else if namespace.as_deref() == Some(FORM)
                    && matches!(local.as_slice(), b"button" | b"checkbox")
                {
                    let form_index = *form_stack.last().ok_or_else(|| {
                        Error::InvalidFormat("interactive control has no form owner".to_string())
                    })?;
                    let parsed = parse_control(&reader, element, local.as_slice())?;
                    register_name(&mut forms[form_index], parsed.name())?;
                    if controls.len() >= MAX_CONTROLS {
                        return invalid("too many interactive controls");
                    }
                    control = Some(controls.len());
                    controls.push(ControlLocation {
                        span: previous..0,
                        form: form_index,
                        control: parsed,
                    });
                } else if namespace.as_deref() == Some(FORM) && is_other_control(&local) {
                    if let Some(form_index) = form_stack.last().copied() {
                        if let Some(name) = optional(&attributes(&reader, element)?, FORM, "name") {
                            register_name(&mut forms[form_index], &name)?;
                        }
                    }
                } else if !form_stack.is_empty()
                    && ((namespace.as_deref() == Some(OFFICE)
                        && matches!(local.as_slice(), b"event-listeners" | b"events"))
                        || (namespace.as_deref() == Some(SCRIPT)
                            && matches!(local.as_slice(), b"event-listener" | b"event")))
                {
                    return invalid(
                        "event and macro content is outside the interactive-control mutation API",
                    );
                }
                stack.push(Open {
                    local,
                    form,
                    control,
                });
            },
            Event::Empty(ref element) => {
                track_xml_id(&reader, element, &mut xml_ids)?;
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
                        names: Vec::new(),
                        controls: Vec::new(),
                    });
                } else if namespace.as_deref() == Some(FORM)
                    && matches!(local.as_slice(), b"button" | b"checkbox")
                {
                    let form_index = *form_stack.last().ok_or_else(|| {
                        Error::InvalidFormat("interactive control has no form owner".to_string())
                    })?;
                    let parsed = parse_control(&reader, element, local.as_slice())?;
                    register_name(&mut forms[form_index], parsed.name())?;
                    if controls.len() >= MAX_CONTROLS {
                        return invalid("too many interactive controls");
                    }
                    controls.push(ControlLocation {
                        span: previous..end,
                        form: form_index,
                        control: parsed,
                    });
                } else if namespace.as_deref() == Some(FORM) && is_other_control(&local) {
                    if let Some(form_index) = form_stack.last().copied() {
                        if let Some(name) = optional(&attributes(&reader, element)?, FORM, "name") {
                            register_name(&mut forms[form_index], &name)?;
                        }
                    }
                } else if !form_stack.is_empty()
                    && ((namespace.as_deref() == Some(OFFICE)
                        && matches!(local.as_slice(), b"event-listeners" | b"events"))
                        || (namespace.as_deref() == Some(SCRIPT)
                            && matches!(local.as_slice(), b"event-listener" | b"event")))
                {
                    return invalid(
                        "event and macro content is outside the interactive-control mutation API",
                    );
                }
            },
            Event::End(ref element) => {
                let open = stack.pop().ok_or_else(|| {
                    Error::InvalidFormat("interactive form XML stack underflow".to_string())
                })?;
                if open.local.as_slice() != element.local_name().as_ref() {
                    return invalid("mismatched interactive form XML elements");
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
            Event::DocType(_) => return invalid("DOCTYPE is not allowed in interactive form XML"),
            Event::Eof => break,
            _ => {},
        }
        previous = end;
        buffer.clear();
    }
    if !stack.is_empty() {
        return invalid("unclosed interactive form XML elements");
    }
    for item in &controls {
        forms[item.form].controls.push(item.clone());
    }
    Ok(Scan { forms, controls })
}

fn parse_control(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
    local: &[u8],
) -> Result<OdfInteractiveControl> {
    let attrs = attributes(reader, element)?;
    validate_allowed(
        &attrs,
        if local == b"button" {
            BUTTON_ATTRS
        } else {
            CHECKBOX_ATTRS
        },
    )?;
    let name = required(&attrs, FORM, "name")?;
    let xml_id = required(&attrs, XML, "id")?;
    if local == b"button" {
        let mut value = OdfButtonControl::new(name, xml_id);
        value.label = optional(&attrs, FORM, "label");
        value.value = optional(&attrs, FORM, "value");
        value.button_type = optional(&attrs, FORM, "button-type")
            .map(|item| match item.as_str() {
                "submit" => Ok(OdfButtonType::Submit),
                "reset" => Ok(OdfButtonType::Reset),
                "push" => Ok(OdfButtonType::Push),
                "url" => Ok(OdfButtonType::Url),
                _ => invalid(format!("invalid button type '{item}'")),
            })
            .transpose()?;
        value.disabled = optional_bool(&attrs, FORM, "disabled")?;
        value.printable = optional_bool(&attrs, FORM, "printable")?;
        value.tab_index = optional_u64(&attrs, FORM, "tab-index")?;
        value.tab_stop = optional_bool(&attrs, FORM, "tab-stop")?;
        value.title = optional(&attrs, FORM, "title");
        value.image_data = optional(&attrs, FORM, "image-data");
        value.target_frame = optional(&attrs, OFFICE, "target-frame");
        value.href = optional(&attrs, XLINK, "href");
        value.image_align = optional(&attrs, FORM, "image-align");
        value.image_position = optional(&attrs, FORM, "image-position");
        value.repeat = optional_bool(&attrs, FORM, "repeat")?;
        value.delay_for_repeat = optional(&attrs, FORM, "delay-for-repeat");
        value.default_button = optional_bool(&attrs, FORM, "default-button")?;
        value.toggle = optional_bool(&attrs, FORM, "toggle")?;
        value.focus_on_click = optional_bool(&attrs, FORM, "focus-on-click")?;
        if let Some(kind) = optional(&attrs, XLINK, "type") {
            if kind != "simple" {
                return invalid("button xlink:type must be simple");
            }
        }
        validate_button(&value)?;
        Ok(value.into())
    } else {
        let mut value = OdfCheckboxControl::new(name, xml_id);
        value.label = optional(&attrs, FORM, "label");
        value.value = optional(&attrs, FORM, "value");
        value.disabled = optional_bool(&attrs, FORM, "disabled")?;
        value.printable = optional_bool(&attrs, FORM, "printable")?;
        value.tab_index = optional_u64(&attrs, FORM, "tab-index")?;
        value.tab_stop = optional_bool(&attrs, FORM, "tab-stop")?;
        value.title = optional(&attrs, FORM, "title");
        value.data_field = optional(&attrs, FORM, "data-field");
        value.linked_cell = optional(&attrs, FORM, "linked-cell");
        value.state = optional(&attrs, FORM, "state")
            .map(|item| OdfCheckboxState::parse(&item))
            .transpose()?;
        value.current_state = optional(&attrs, FORM, "current-state")
            .map(|item| OdfCheckboxState::parse(&item))
            .transpose()?;
        value.is_tristate = optional_bool(&attrs, FORM, "is-tristate")?;
        value.visual_effect = optional(&attrs, FORM, "visual-effect");
        value.image_align = optional(&attrs, FORM, "image-align");
        value.image_position = optional(&attrs, FORM, "image-position");
        validate_checkbox(&value)?;
        Ok(value.into())
    }
}

const BUTTON_ATTRS: &[(&str, &str)] = &[
    (FORM, "name"),
    (XML, "id"),
    (FORM, "id"),
    (FORM, "control-implementation"),
    (FORM, "disabled"),
    (FORM, "printable"),
    (FORM, "tab-index"),
    (FORM, "tab-stop"),
    (FORM, "title"),
    (XFORMS, "bind"),
    (FORM, "label"),
    (FORM, "value"),
    (FORM, "button-type"),
    (FORM, "image-data"),
    (OFFICE, "target-frame"),
    (XLINK, "type"),
    (XLINK, "href"),
    (FORM, "image-align"),
    (FORM, "image-position"),
    (FORM, "repeat"),
    (FORM, "delay-for-repeat"),
    (FORM, "default-button"),
    (FORM, "toggle"),
    (FORM, "focus-on-click"),
    (FORM, "xforms-submission"),
];
const CHECKBOX_ATTRS: &[(&str, &str)] = &[
    (FORM, "name"),
    (XML, "id"),
    (FORM, "id"),
    (FORM, "control-implementation"),
    (FORM, "disabled"),
    (FORM, "printable"),
    (FORM, "tab-index"),
    (FORM, "tab-stop"),
    (FORM, "title"),
    (FORM, "input-required"),
    (XFORMS, "bind"),
    (FORM, "label"),
    (FORM, "value"),
    (FORM, "data-field"),
    (FORM, "linked-cell"),
    (FORM, "state"),
    (FORM, "current-state"),
    (FORM, "is-tristate"),
    (FORM, "visual-effect"),
    (FORM, "image-align"),
    (FORM, "image-position"),
];

fn validate_form(reader: &NsReader<&[u8]>, element: &BytesStart<'_>) -> Result<()> {
    let attrs = attributes(reader, element)?;
    validate_allowed(&attrs, FORM_ATTRS)?;
    let name = required(&attrs, FORM, "name")?;
    validate_name("form name", &name)?;
    for attr in &attrs {
        if attr.namespace.as_deref() == Some(FORM)
            && matches!(
                attr.local.as_str(),
                "apply-filter"
                    | "allow-deletes"
                    | "allow-inserts"
                    | "allow-updates"
                    | "escape-processing"
                    | "ignore-result"
            )
        {
            let _ = parse_bool(&attr.value, &attr.local)?;
        }
    }
    if let Some(kind) = optional(&attrs, XLINK, "type") {
        if kind != "simple" {
            return invalid("form xlink:type must be simple");
        }
    }
    validate_resource(
        "form href",
        optional(&attrs, XLINK, "href").as_deref(),
        false,
    )
}
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
fn is_other_control(local: &[u8]) -> bool {
    matches!(
        local,
        b"text"
            | b"textarea"
            | b"password"
            | b"file"
            | b"formatted-text"
            | b"number"
            | b"date"
            | b"time"
            | b"fixed-text"
            | b"combobox"
            | b"listbox"
            | b"radio"
            | b"frame"
            | b"image-frame"
            | b"hidden"
            | b"grid"
            | b"value-range"
            | b"generic-control"
            | b"image"
    )
}
fn register_name(form: &mut FormLocation, name: &str) -> Result<()> {
    validate_name("form control name", name)?;
    if form.names.iter().any(|existing| existing == name) {
        return invalid(format!("duplicate form control name '{name}'"));
    }
    form.names.push(name.to_string());
    Ok(())
}
fn reject_duplicate(
    form: &FormLocation,
    replacement: &OdfInteractiveControl,
    current: Option<&OdfInteractiveControl>,
) -> Result<()> {
    for name in &form.names {
        if name == replacement.name() && !current.is_some_and(|value| value.name() == name) {
            return invalid(format!(
                "duplicate form control name '{}'",
                replacement.name()
            ));
        }
    }
    for item in &form.controls {
        if item.control.xml_id() == replacement.xml_id()
            && !current.is_some_and(|value| value.xml_id() == item.control.xml_id())
        {
            return invalid(format!(
                "duplicate form control xml:id '{}'",
                replacement.xml_id()
            ));
        }
    }
    Ok(())
}

fn attributes(reader: &NsReader<&[u8]>, element: &BytesStart<'_>) -> Result<Vec<Attr>> {
    let mut output = Vec::new();
    for attribute in element.attributes().with_checks(true) {
        let attribute = attribute.map_err(|error| {
            Error::InvalidFormat(format!("invalid interactive control attribute: {error}"))
        })?;
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
                Error::InvalidFormat(format!(
                    "invalid interactive control attribute value: {error}"
                ))
            })?
            .into_owned();
        output.push(Attr {
            namespace,
            local: String::from_utf8_lossy(local.as_ref()).into_owned(),
            value,
        });
    }
    Ok(output)
}
fn track_xml_id(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
    ids: &mut Vec<String>,
) -> Result<()> {
    if let Some(value) = optional(&attributes(reader, element)?, XML, "id") {
        validate_xml_id(&value)?;
        if ids.iter().any(|existing| existing == &value) {
            return invalid(format!("duplicate xml:id '{value}'"));
        }
        ids.push(value);
    }
    Ok(())
}
fn validate_allowed(attrs: &[Attr], allowed: &[(&str, &str)]) -> Result<()> {
    for attr in attrs {
        if !allowed.iter().any(|(namespace, local)| {
            attr.namespace.as_deref() == Some(*namespace) && attr.local == *local
        }) {
            return invalid(format!(
                "unsupported interactive control attribute '{}:{}'",
                attr.namespace.as_deref().unwrap_or(""),
                attr.local
            ));
        }
    }
    Ok(())
}
fn optional(attrs: &[Attr], namespace: &str, local: &str) -> Option<String> {
    attrs
        .iter()
        .find(|item| item.namespace.as_deref() == Some(namespace) && item.local == local)
        .map(|item| item.value.clone())
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
fn validate_identity(name: &str, id: &str) -> Result<()> {
    validate_name("form control name", name)?;
    validate_xml_id(id)
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
fn validate_optional(label: &str, value: Option<&str>) -> Result<()> {
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
        .any(|ch| matches!(ch as u32,0..=8|11|12|14..=31))
    {
        return invalid(format!("{label} contains invalid XML characters"));
    }
    Ok(())
}
fn validate_token(label: &str, value: Option<&str>, allowed: &[&str]) -> Result<()> {
    if let Some(value) = value {
        if !allowed.contains(&value) {
            return invalid(format!("invalid {label} '{value}'"));
        }
    }
    Ok(())
}
fn validate_duration(value: &str) -> Result<()> {
    if value.len() > 128 || !value.starts_with('P') {
        return invalid(format!("invalid repeat duration '{value}'"));
    }
    let body = &value[1..];
    let mut pieces = body.split('T');
    let date = pieces.next().unwrap_or("");
    let time = pieces.next();
    if pieces.next().is_some() {
        return invalid(format!("invalid repeat duration '{value}'"));
    }
    let (date, years) = duration_component(date, 'Y', false)?;
    let (date, months) = duration_component(date, 'M', false)?;
    let (date, days) = duration_component(date, 'D', false)?;
    if !date.is_empty() {
        return invalid(format!("invalid repeat duration '{value}'"));
    }
    let mut any = years || months || days;
    if let Some(time) = time {
        if time.is_empty() {
            return invalid(format!("invalid repeat duration '{value}'"));
        }
        let (time, hours) = duration_component(time, 'H', false)?;
        let (time, minutes) = duration_component(time, 'M', false)?;
        let (time, seconds) = duration_component(time, 'S', true)?;
        if !time.is_empty() {
            return invalid(format!("invalid repeat duration '{value}'"));
        }
        any |= hours || minutes || seconds;
    }
    if !any {
        return invalid(format!("invalid repeat duration '{value}'"));
    }
    Ok(())
}
fn duration_component(value: &str, suffix: char, decimal: bool) -> Result<(&str, bool)> {
    let Some(position) = value.find(suffix) else {
        return Ok((value, false));
    };
    let number = &value[..position];
    if number.is_empty() {
        return invalid("empty repeat duration component");
    }
    let valid = if decimal {
        let mut parts = number.split('.');
        let whole = parts.next().unwrap_or("");
        let fraction = parts.next();
        !whole.is_empty()
            && whole.bytes().all(|byte| byte.is_ascii_digit())
            && fraction.is_none_or(|part| {
                !part.is_empty() && part.bytes().all(|byte| byte.is_ascii_digit())
            })
            && parts.next().is_none()
    } else {
        number.bytes().all(|byte| byte.is_ascii_digit())
    };
    if !valid {
        return invalid("invalid repeat duration component");
    }
    Ok((&value[position + 1..], true))
}
fn validate_resource(label: &str, value: Option<&str>, package: bool) -> Result<()> {
    if let Some(value) = value {
        if value.len() > MAX_RESOURCE {
            return invalid(format!("{label} reference exceeds 8 KiB"));
        }
        validate_string(label, value)?;
        let lower = value.trim().to_ascii_lowercase();
        if lower.starts_with("javascript:") || lower.starts_with("data:") {
            return invalid(format!("active {label} scheme is not allowed"));
        }
        if package && (value.starts_with('/') || value.contains("..") || value.contains('\\')) {
            return invalid("unsafe button image package path");
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
fn bind_fragment(mut value: String) -> String {
    for (prefix, namespace) in [("form", FORM), ("office", OFFICE), ("xlink", XLINK)] {
        if value.contains(&format!("{prefix}:")) && !value.contains(&format!("xmlns:{prefix}=")) {
            value = value.replacen(' ', &format!(" xmlns:{prefix}=\"{namespace}\" "), 1);
        }
    }
    value
}
fn qname(element: &BytesStart<'_>) -> Result<String> {
    String::from_utf8(element.name().as_ref().to_vec())
        .map_err(|_| Error::InvalidFormat("invalid interactive form element name".to_string()))
}
fn apply(xml: &str, span: Range<usize>, replacement: &str) -> Result<String> {
    let mut output = String::with_capacity(xml.len() - span.len() + replacement.len());
    output.push_str(&xml[..span.start]);
    output.push_str(replacement);
    output.push_str(&xml[span.end..]);
    Ok(output)
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
    const ROOT: &str = r#"<o:document-content xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:f="urn:oasis:names:tc:opendocument:xmlns:form:1.0" xmlns:x="http://www.w3.org/1999/xlink"><o:body><o:text><o:forms>"#;
    const END: &str = "</o:forms></o:text></o:body></o:document-content>";

    #[test]
    fn canonical_button_and_checkbox_are_typed_and_escaped() {
        let mut button = OdfButtonControl::new("Run & go", "run_1");
        button.label = Some("Run <now>".into());
        button.button_type = Some(OdfButtonType::Push);
        button.href = Some("https://example.invalid/action?a=1&b=2".into());
        button.image_data = Some("Pictures/run.png".into());
        button.repeat = Some(true);
        button.delay_for_repeat = Some("PT0.05S".into());
        let xml = button.to_xml_fragment().unwrap();
        assert!(xml.contains("form:button-type=\"push\"") && xml.contains("a=1&amp;b=2"));
        let mut check = OdfCheckboxControl::new("Enabled", "enabled_1");
        check.label = Some("Enabled".into());
        check.is_tristate = Some(true);
        check.state = Some(OdfCheckboxState::Unknown);
        check.current_state = Some(OdfCheckboxState::Checked);
        let mut form = OdfInteractiveForm::new("Main");
        form.add_control(button).unwrap();
        form.add_control(check).unwrap();
        let document = format!("{ROOT}{}{END}", form.to_xml_fragment().unwrap());
        assert_eq!(interactive_controls(&document).unwrap().len(), 2);
    }

    #[test]
    fn alias_namespace_lossless_mutation_preserves_unrelated_bytes() {
        let xml = format!(
            r#"{ROOT}<f:form f:name="Main"><f:button f:name="A" xml:id="a" f:label="old"><f:properties><f:property f:property-name="Keep" o:value-type="void"/></f:properties></f:button><!--keep--><f:text f:name="Text" xml:id="text_1"/></f:form>{END}"#
        );
        let check: OdfInteractiveControl = OdfCheckboxControl::new("B", "b").into();
        let inserted = insert_interactive_control_xml(&xml, 0, &check).unwrap();
        assert!(inserted.contains("<!--keep-->") && inserted.contains("form:checkbox"));
        let mut replacement = OdfButtonControl::new("A2", "a2");
        replacement.label = Some("new".into());
        let replaced = replace_interactive_control_xml(&inserted, 0, &replacement.into()).unwrap();
        assert!(replaced.contains("<!--keep-->") && replaced.contains("f:text"));
        let removed = remove_interactive_control_xml(&replaced, 1).unwrap();
        assert_eq!(interactive_controls(&removed).unwrap().len(), 1);
        let empty = format!(r#"{ROOT}<f:form f:name="Empty"/>{END}"#);
        assert!(
            insert_interactive_control_xml(&empty, 0, &OdfCheckboxControl::new("C", "c").into())
                .unwrap()
                .contains("</f:form>")
        );
    }

    #[test]
    fn hostile_namespaces_attributes_events_resources_and_limits_are_rejected() {
        assert!(
            OdfButtonControl::new("B", "1bad")
                .to_xml_fragment()
                .is_err()
        );
        let wrong = format!(
            r#"{ROOT}<f:form f:name="Main"><x:button f:name="B" xml:id="b"/></f:form>{END}"#
        );
        assert!(interactive_controls(&wrong).unwrap().is_empty());
        let attr = format!(
            r#"{ROOT}<f:form f:name="Main"><f:button f:name="B" xml:id="b" x:label="spoof"/></f:form>{END}"#
        );
        assert!(interactive_controls(&attr).is_err());
        let event = format!(r#"{ROOT}<f:form f:name="Main"><o:event-listeners/></f:form>{END}"#);
        assert!(interactive_controls(&event).is_err());
        let duplicate = format!(
            r#"{ROOT}<o:p xml:id="same"/><f:form f:name="Main"><f:checkbox f:name="C" xml:id="same"/></f:form>{END}"#
        );
        assert!(interactive_controls(&duplicate).is_err());
        let mut button = OdfButtonControl::new("B", "b");
        button.href = Some("javascript:alert(1)".into());
        assert!(button.to_xml_fragment().is_err());
        button.href = None;
        button.image_data = Some("../secret.png".into());
        assert!(button.to_xml_fragment().is_err());
        button.image_data = Some("x".repeat(MAX_RESOURCE + 1));
        assert!(button.to_xml_fragment().is_err());
        button.image_data = None;
        button.delay_for_repeat = Some("P+".into());
        assert!(button.to_xml_fragment().is_err());
    }

    #[test]
    fn libreoffice_odfpy_and_odfdo_shapes_parse_inertly() {
        let lo = include_str!(
            "../../../../3rdparty/libreoffice-core/vcl/qa/cppunit/pdfexport/data/formcontrol.fodt"
        );
        let parsed = interactive_controls(lo).unwrap();
        assert!(
            parsed
                .iter()
                .any(|item| matches!(item, OdfInteractiveControl::Checkbox(_)))
        );
        let producer = format!(
            r#"{ROOT}<f:form f:name="Producers"><f:button f:name="odfpy" xml:id="button_1" f:label="Run" f:button-type="push" o:target-frame="" x:href="" f:image-data="" f:delay-for-repeat="PT0.050000000S" f:image-position="center"/><f:checkbox f:name="odfdo" xml:id="check_1" f:label="Check" f:state="checked" f:current-state="unknown" f:is-tristate="true"/></f:form>{END}"#
        );
        assert_eq!(interactive_controls(&producer).unwrap().len(), 2);
    }

    #[test]
    fn builder_and_mutable_package_round_trip() {
        use crate::odt::{Document, DocumentBuilder, MutableDocument};
        let mut form = OdfInteractiveForm::new("Main");
        form.add_control(OdfButtonControl::new("Run", "run_1"))
            .unwrap();
        let mut builder = DocumentBuilder::new();
        builder.add_interactive_form(&form).unwrap();
        builder.add_paragraph("body").unwrap();
        let document = Document::from_bytes(builder.build().unwrap()).unwrap();
        let mut mutable = MutableDocument::from_document(document).unwrap();
        assert_eq!(mutable.interactive_controls().unwrap().len(), 1);
        let checkbox: OdfInteractiveControl = OdfCheckboxControl::new("Check", "check_1").into();
        mutable.insert_interactive_control(0, &checkbox).unwrap();
        let replacement: OdfInteractiveControl = OdfButtonControl::new("Go", "go_1").into();
        assert_eq!(
            mutable
                .replace_interactive_control(0, &replacement)
                .unwrap()
                .name(),
            "Run"
        );
        assert_eq!(
            mutable.remove_interactive_control(1).unwrap().name(),
            "Check"
        );
        assert_eq!(mutable.interactive_controls().unwrap(), [replacement]);
    }
}
