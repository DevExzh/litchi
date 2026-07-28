//! Typed writing and lossless mutation for radio, frame, and image-button controls.

use litchi_core::{Error, Result};
use quick_xml::events::{BytesStart, Event};
use quick_xml::name::ResolveResult;
use quick_xml::reader::NsReader;
use std::ops::Range;

const FORM: &str = "urn:oasis:names:tc:opendocument:xmlns:form:1.0";
const OFFICE: &str = "urn:oasis:names:tc:opendocument:xmlns:office:1.0";
const SCRIPT: &str = "urn:oasis:names:tc:opendocument:xmlns:script:1.0";
const XLINK: &str = "http://www.w3.org/1999/xlink";
const XFORMS: &str = "http://www.w3.org/2002/xforms";
const XML: &str = "http://www.w3.org/XML/1998/namespace";
const MAX_XML: usize = 64 * 1024 * 1024;
const MAX_STRING: usize = 1024 * 1024;
const MAX_URI: usize = 8192;
const MAX_AGGREGATE: usize = 16 * 1024 * 1024;
const MAX_FORMS: usize = 4096;
const MAX_CONTROLS: usize = 65_536;

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum OdfImageButtonType {
    Submit,
    Reset,
    Push,
    Url,
}

impl OdfImageButtonType {
    fn token(self) -> &'static str {
        match self {
            Self::Submit => "submit",
            Self::Reset => "reset",
            Self::Push => "push",
            Self::Url => "url",
        }
    }

    fn parse(value: &str) -> Result<Self> {
        match value {
            "submit" => Ok(Self::Submit),
            "reset" => Ok(Self::Reset),
            "push" => Ok(Self::Push),
            "url" => Ok(Self::Url),
            _ => invalid(format!("invalid image button type '{value}'")),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum OdfRadioVisualEffect {
    Flat,
    ThreeD,
}

impl OdfRadioVisualEffect {
    fn token(self) -> &'static str {
        match self {
            Self::Flat => "flat",
            Self::ThreeD => "3d",
        }
    }

    fn parse(value: &str) -> Result<Self> {
        match value {
            "flat" => Ok(Self::Flat),
            "3d" => Ok(Self::ThreeD),
            _ => invalid(format!("invalid radio visual effect '{value}'")),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum OdfRelativeImagePosition {
    Center,
    Start,
    End,
    Top,
    Bottom,
}

impl OdfRelativeImagePosition {
    fn token(self) -> &'static str {
        match self {
            Self::Center => "center",
            Self::Start => "start",
            Self::End => "end",
            Self::Top => "top",
            Self::Bottom => "bottom",
        }
    }

    fn parse(value: &str) -> Result<Self> {
        match value {
            "center" => Ok(Self::Center),
            "start" => Ok(Self::Start),
            "end" => Ok(Self::End),
            "top" => Ok(Self::Top),
            "bottom" => Ok(Self::Bottom),
            _ => invalid(format!("invalid relative image position '{value}'")),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum OdfRelativeImageAlign {
    Start,
    Center,
    End,
}

impl OdfRelativeImageAlign {
    fn token(self) -> &'static str {
        match self {
            Self::Start => "start",
            Self::Center => "center",
            Self::End => "end",
        }
    }

    fn parse(value: &str) -> Result<Self> {
        match value {
            "start" => Ok(Self::Start),
            "center" => Ok(Self::Center),
            "end" => Ok(Self::End),
            _ => invalid(format!("invalid relative image alignment '{value}'")),
        }
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct OdfRadioControl {
    pub name: String,
    pub xml_id: String,
    pub current_selected: Option<bool>,
    pub disabled: Option<bool>,
    pub label: Option<String>,
    pub printable: Option<bool>,
    pub selected: Option<bool>,
    pub tab_index: Option<u64>,
    pub tab_stop: Option<bool>,
    pub title: Option<String>,
    pub value: Option<String>,
    pub data_field: Option<String>,
    pub visual_effect: Option<OdfRadioVisualEffect>,
    pub image_position: Option<OdfRelativeImagePosition>,
    pub image_align: Option<OdfRelativeImageAlign>,
    pub linked_cell: Option<String>,
}

impl OdfRadioControl {
    pub fn new(name: impl Into<String>, xml_id: impl Into<String>) -> Self {
        Self {
            name: name.into(),
            xml_id: xml_id.into(),
            current_selected: None,
            disabled: None,
            label: None,
            printable: None,
            selected: None,
            tab_index: None,
            tab_stop: None,
            title: None,
            value: None,
            data_field: None,
            visual_effect: None,
            image_position: None,
            image_align: None,
            linked_cell: None,
        }
    }

    pub fn to_xml_fragment(&self) -> Result<String> {
        radio_xml(self)
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct OdfFrameControl {
    pub name: String,
    pub xml_id: String,
    pub disabled: Option<bool>,
    pub form_for: Option<String>,
    pub label: Option<String>,
    pub printable: Option<bool>,
    pub title: Option<String>,
}

impl OdfFrameControl {
    pub fn new(name: impl Into<String>, xml_id: impl Into<String>) -> Self {
        Self {
            name: name.into(),
            xml_id: xml_id.into(),
            disabled: None,
            form_for: None,
            label: None,
            printable: None,
            title: None,
        }
    }

    pub fn to_xml_fragment(&self) -> Result<String> {
        frame_xml(self)
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct OdfImageControl {
    pub name: String,
    pub xml_id: String,
    pub button_type: Option<OdfImageButtonType>,
    pub disabled: Option<bool>,
    pub image_data: Option<String>,
    pub printable: Option<bool>,
    pub tab_index: Option<u64>,
    pub tab_stop: Option<bool>,
    pub target_frame: Option<String>,
    pub href: Option<String>,
    pub title: Option<String>,
    pub value: Option<String>,
}

impl OdfImageControl {
    pub fn new(name: impl Into<String>, xml_id: impl Into<String>) -> Self {
        Self {
            name: name.into(),
            xml_id: xml_id.into(),
            button_type: None,
            disabled: None,
            image_data: None,
            printable: None,
            tab_index: None,
            tab_stop: None,
            target_frame: None,
            href: None,
            title: None,
            value: None,
        }
    }

    pub fn to_xml_fragment(&self) -> Result<String> {
        image_xml(self)
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub enum OdfVisualControl {
    Radio(OdfRadioControl),
    Frame(OdfFrameControl),
    Image(OdfImageControl),
}

impl From<OdfRadioControl> for OdfVisualControl {
    fn from(value: OdfRadioControl) -> Self {
        Self::Radio(value)
    }
}

impl From<OdfFrameControl> for OdfVisualControl {
    fn from(value: OdfFrameControl) -> Self {
        Self::Frame(value)
    }
}

impl From<OdfImageControl> for OdfVisualControl {
    fn from(value: OdfImageControl) -> Self {
        Self::Image(value)
    }
}

impl OdfVisualControl {
    pub fn name(&self) -> &str {
        match self {
            Self::Radio(v) => &v.name,
            Self::Frame(v) => &v.name,
            Self::Image(v) => &v.name,
        }
    }

    pub fn xml_id(&self) -> &str {
        match self {
            Self::Radio(v) => &v.xml_id,
            Self::Frame(v) => &v.xml_id,
            Self::Image(v) => &v.xml_id,
        }
    }

    pub fn to_xml_fragment(&self) -> Result<String> {
        match self {
            Self::Radio(v) => radio_xml(v),
            Self::Frame(v) => frame_xml(v),
            Self::Image(v) => image_xml(v),
        }
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct OdfVisualForm {
    pub name: String,
    pub controls: Vec<OdfVisualControl>,
    pub apply_filter: Option<bool>,
}

impl OdfVisualForm {
    pub fn new(name: impl Into<String>) -> Self {
        Self {
            name: name.into(),
            controls: Vec::new(),
            apply_filter: None,
        }
    }

    pub fn add_control(&mut self, control: impl Into<OdfVisualControl>) -> Result<()> {
        let control = control.into();
        validate_control(&control)?;
        if self
            .controls
            .iter()
            .any(|item| item.xml_id() == control.xml_id())
        {
            return invalid(format!(
                "duplicate visual control xml:id '{}'",
                control.xml_id()
            ));
        }
        self.controls.push(control);
        if let Err(error) = validate_radio_groups(&self.controls) {
            self.controls.pop();
            return Err(error);
        }
        Ok(())
    }

    pub fn to_xml_fragment(&self) -> Result<String> {
        validate_name("form name", &self.name)?;
        validate_controls(&self.controls)?;
        let mut out = format!(
            r#"<form:form xmlns:form="{FORM}" xmlns:office="{OFFICE}" xmlns:xlink="{XLINK}" form:name="{}""#,
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

pub fn visual_controls(xml: &str) -> Result<Vec<OdfVisualControl>> {
    Ok(scan(xml)?
        .controls
        .into_iter()
        .map(|item| item.control)
        .collect())
}

pub fn insert_visual_control_xml(
    xml: &str,
    form_index: usize,
    control: &OdfVisualControl,
) -> Result<String> {
    validate_control(control)?;
    let scan = scan(xml)?;
    let form = scan
        .forms
        .get(form_index)
        .ok_or_else(|| Error::InvalidFormat(format!("form {form_index} is out of bounds")))?;
    if scan.ids.iter().any(|id| id == control.xml_id()) {
        return invalid(format!("duplicate xml:id '{}'", control.xml_id()));
    }
    reject_group_conflict(&form.controls, control, None)?;
    let fragment = bind_fragment(control.to_xml_fragment()?);
    match form.site.clone() {
        Site::Paired { close_start } => apply(xml, close_start..close_start, &fragment),
        Site::Empty { start, end, qname } => expand_empty(xml, start, end, &qname, &fragment),
    }
}

pub fn replace_visual_control_xml(
    xml: &str,
    index: usize,
    replacement: &OdfVisualControl,
) -> Result<String> {
    validate_control(replacement)?;
    let scan = scan(xml)?;
    let old = scan
        .controls
        .get(index)
        .ok_or_else(|| Error::InvalidFormat(format!("visual control {index} is out of bounds")))?;
    if replacement.xml_id() != old.control.xml_id()
        && scan.ids.iter().any(|id| id == replacement.xml_id())
    {
        return invalid(format!("duplicate xml:id '{}'", replacement.xml_id()));
    }
    reject_group_conflict(&scan.forms[old.form].controls, replacement, Some(index))?;
    apply(
        xml,
        old.span.clone(),
        &bind_fragment(replacement.to_xml_fragment()?),
    )
}

pub fn remove_visual_control_xml(xml: &str, index: usize) -> Result<String> {
    let scan = scan(xml)?;
    let old = scan
        .controls
        .get(index)
        .ok_or_else(|| Error::InvalidFormat(format!("visual control {index} is out of bounds")))?;
    apply(xml, old.span.clone(), "")
}

fn radio_xml(value: &OdfRadioControl) -> Result<String> {
    validate_radio(value)?;
    let mut out = control_start("radio", &value.name, &value.xml_id);
    push_bool(&mut out, "form:current-selected", value.current_selected);
    push_bool(&mut out, "form:disabled", value.disabled);
    push_string(&mut out, "form:label", value.label.as_deref());
    push_bool(&mut out, "form:printable", value.printable);
    push_bool(&mut out, "form:selected", value.selected);
    push_u64(&mut out, "form:tab-index", value.tab_index);
    push_bool(&mut out, "form:tab-stop", value.tab_stop);
    push_string(&mut out, "form:title", value.title.as_deref());
    push_string(&mut out, "form:value", value.value.as_deref());
    push_string(&mut out, "form:data-field", value.data_field.as_deref());
    if let Some(v) = value.visual_effect {
        push_string(&mut out, "form:visual-effect", Some(v.token()));
    }
    if let Some(v) = value.image_position {
        push_string(&mut out, "form:image-position", Some(v.token()));
    }
    if let Some(v) = value.image_align {
        push_string(&mut out, "form:image-align", Some(v.token()));
    }
    push_string(&mut out, "form:linked-cell", value.linked_cell.as_deref());
    out.push_str("/>");
    Ok(out)
}

fn frame_xml(value: &OdfFrameControl) -> Result<String> {
    validate_frame(value)?;
    let mut out = control_start("frame", &value.name, &value.xml_id);
    push_bool(&mut out, "form:disabled", value.disabled);
    push_string(&mut out, "form:for", value.form_for.as_deref());
    push_string(&mut out, "form:label", value.label.as_deref());
    push_bool(&mut out, "form:printable", value.printable);
    push_string(&mut out, "form:title", value.title.as_deref());
    out.push_str("/>");
    Ok(out)
}

fn image_xml(value: &OdfImageControl) -> Result<String> {
    validate_image(value)?;
    let mut out = control_start("image", &value.name, &value.xml_id);
    if let Some(v) = value.button_type {
        push_string(&mut out, "form:button-type", Some(v.token()));
    }
    push_bool(&mut out, "form:disabled", value.disabled);
    push_string(&mut out, "form:image-data", value.image_data.as_deref());
    push_bool(&mut out, "form:printable", value.printable);
    push_u64(&mut out, "form:tab-index", value.tab_index);
    push_bool(&mut out, "form:tab-stop", value.tab_stop);
    push_string(
        &mut out,
        "office:target-frame",
        value.target_frame.as_deref(),
    );
    push_string(&mut out, "xlink:href", value.href.as_deref());
    push_string(&mut out, "form:title", value.title.as_deref());
    push_string(&mut out, "form:value", value.value.as_deref());
    out.push_str("/>");
    Ok(out)
}

fn control_start(kind: &str, name: &str, xml_id: &str) -> String {
    format!(
        r#"<form:{kind} form:name="{}" xml:id="{}""#,
        escape(name),
        escape(xml_id)
    )
}

fn validate_control(value: &OdfVisualControl) -> Result<()> {
    match value {
        OdfVisualControl::Radio(v) => validate_radio(v),
        OdfVisualControl::Frame(v) => validate_frame(v),
        OdfVisualControl::Image(v) => validate_image(v),
    }
}

fn validate_radio(value: &OdfRadioControl) -> Result<()> {
    validate_identity(&value.name, &value.xml_id)?;
    validate_optionals(&[
        ("radio label", value.label.as_deref()),
        ("radio title", value.title.as_deref()),
        ("radio value", value.value.as_deref()),
        ("radio data field", value.data_field.as_deref()),
        ("radio linked cell", value.linked_cell.as_deref()),
    ])?;
    if value.image_align.is_some()
        && !matches!(
            value.image_position,
            Some(
                OdfRelativeImagePosition::Start
                    | OdfRelativeImagePosition::End
                    | OdfRelativeImagePosition::Top
                    | OdfRelativeImagePosition::Bottom
            )
        )
    {
        return invalid("radio image-align requires a non-center image-position");
    }
    Ok(())
}

fn validate_frame(value: &OdfFrameControl) -> Result<()> {
    validate_identity(&value.name, &value.xml_id)?;
    validate_optionals(&[
        ("frame for", value.form_for.as_deref()),
        ("frame label", value.label.as_deref()),
        ("frame title", value.title.as_deref()),
    ])
}

fn validate_image(value: &OdfImageControl) -> Result<()> {
    validate_identity(&value.name, &value.xml_id)?;
    validate_optionals(&[
        ("image target frame", value.target_frame.as_deref()),
        ("image title", value.title.as_deref()),
        ("image value", value.value.as_deref()),
    ])?;
    validate_uri("image data", value.image_data.as_deref())?;
    validate_uri("image href", value.href.as_deref())
}

fn validate_controls(controls: &[OdfVisualControl]) -> Result<()> {
    if controls.len() > MAX_CONTROLS {
        return invalid("too many visual controls");
    }
    let mut ids: Vec<&str> = Vec::new();
    let mut aggregate = 0usize;
    for control in controls {
        validate_control(control)?;
        if ids.contains(&control.xml_id()) {
            return invalid(format!(
                "duplicate visual control xml:id '{}'",
                control.xml_id()
            ));
        }
        ids.push(control.xml_id());
        aggregate = aggregate.saturating_add(control_size(control));
        if aggregate > MAX_AGGREGATE {
            return invalid("visual control strings exceed 16 MiB");
        }
    }
    validate_radio_groups(controls)
}

fn validate_radio_groups(controls: &[OdfVisualControl]) -> Result<()> {
    for (index, control) in controls.iter().enumerate() {
        let OdfVisualControl::Radio(radio) = control else {
            continue;
        };
        for other in &controls[index + 1..] {
            let OdfVisualControl::Radio(other) = other else {
                continue;
            };
            if radio.name == other.name
                && radio.selected == Some(true)
                && other.selected == Some(true)
            {
                return invalid(format!(
                    "radio group '{}' has multiple default selections",
                    radio.name
                ));
            }
            if radio.name == other.name
                && radio.current_selected == Some(true)
                && other.current_selected == Some(true)
            {
                return invalid(format!(
                    "radio group '{}' has multiple current selections",
                    radio.name
                ));
            }
        }
    }
    Ok(())
}

fn reject_group_conflict(
    controls: &[ControlLocation],
    replacement: &OdfVisualControl,
    replaced_index: Option<usize>,
) -> Result<()> {
    let mut values = Vec::with_capacity(controls.len() + 1);
    for control in controls {
        if replaced_index.is_some_and(|index| control.global_index == index) {
            continue;
        }
        values.push(control.control.clone());
    }
    values.push(replacement.clone());
    validate_radio_groups(&values)
}

fn control_size(control: &OdfVisualControl) -> usize {
    match control {
        OdfVisualControl::Radio(v) => string_size(&[
            Some(&v.name),
            Some(&v.xml_id),
            v.label.as_ref(),
            v.title.as_ref(),
            v.value.as_ref(),
            v.data_field.as_ref(),
            v.linked_cell.as_ref(),
        ]),
        OdfVisualControl::Frame(v) => string_size(&[
            Some(&v.name),
            Some(&v.xml_id),
            v.form_for.as_ref(),
            v.label.as_ref(),
            v.title.as_ref(),
        ]),
        OdfVisualControl::Image(v) => string_size(&[
            Some(&v.name),
            Some(&v.xml_id),
            v.image_data.as_ref(),
            v.target_frame.as_ref(),
            v.href.as_ref(),
            v.title.as_ref(),
            v.value.as_ref(),
        ]),
    }
}

fn string_size(values: &[Option<&String>]) -> usize {
    values
        .iter()
        .flatten()
        .fold(0usize, |sum, value| sum.saturating_add(value.len()))
}

#[derive(Clone)]
struct ControlLocation {
    span: Range<usize>,
    form: usize,
    global_index: usize,
    control: OdfVisualControl,
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
        return invalid("visual form XML exceeds 64 MiB");
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
            .map_err(|error| Error::InvalidFormat(format!("invalid visual form XML: {error}")))?;
        let namespace = match namespace {
            ResolveResult::Bound(uri) => Some(String::from_utf8_lossy(uri.as_ref()).into_owned()),
            _ => None,
        };
        let end = reader.buffer_position() as usize;
        match event {
            Event::Start(ref element) => {
                if stack.len() >= 128 {
                    return invalid("visual form XML nesting exceeds 128 levels");
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
                } else if namespace.as_deref() == Some(FORM)
                    && matches!(local.as_slice(), b"radio" | b"frame" | b"image")
                {
                    if stack.iter().any(|open| open.control.is_some()) {
                        return invalid("visual controls cannot be nested");
                    }
                    let owner = *form_stack.last().ok_or_else(|| {
                        Error::InvalidFormat("visual control has no form owner".to_string())
                    })?;
                    let parsed = parse_control(&reader, element, local.as_slice())?;
                    aggregate = aggregate.saturating_add(control_size(&parsed));
                    if aggregate > MAX_AGGREGATE {
                        return invalid("visual control strings exceed 16 MiB");
                    }
                    if controls.len() >= MAX_CONTROLS {
                        return invalid("too many visual controls");
                    }
                    control = Some(controls.len());
                    controls.push(ControlLocation {
                        span: previous..0,
                        form: owner,
                        global_index: controls.len(),
                        control: parsed,
                    });
                } else if is_active_form_content(namespace.as_deref(), local.as_slice())
                    && !form_stack.is_empty()
                {
                    return invalid(
                        "event and macro content is outside the visual-control mutation API",
                    );
                } else if stack.iter().any(|open| open.control.is_some())
                    && !(namespace.as_deref() == Some(FORM)
                        && matches!(
                            local.as_slice(),
                            b"properties" | b"property" | b"list-property" | b"list-value"
                        ))
                {
                    return invalid("unexpected child element in visual control");
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
                } else if namespace.as_deref() == Some(FORM)
                    && matches!(local.as_slice(), b"radio" | b"frame" | b"image")
                {
                    if stack.iter().any(|open| open.control.is_some()) {
                        return invalid("visual controls cannot be nested");
                    }
                    let owner = *form_stack.last().ok_or_else(|| {
                        Error::InvalidFormat("visual control has no form owner".to_string())
                    })?;
                    let parsed = parse_control(&reader, element, local.as_slice())?;
                    aggregate = aggregate.saturating_add(control_size(&parsed));
                    if aggregate > MAX_AGGREGATE {
                        return invalid("visual control strings exceed 16 MiB");
                    }
                    if controls.len() >= MAX_CONTROLS {
                        return invalid("too many visual controls");
                    }
                    controls.push(ControlLocation {
                        span: previous..end,
                        form: owner,
                        global_index: controls.len(),
                        control: parsed,
                    });
                } else if is_active_form_content(namespace.as_deref(), local.as_slice())
                    && !form_stack.is_empty()
                {
                    return invalid(
                        "event and macro content is outside the visual-control mutation API",
                    );
                } else if stack.iter().any(|open| open.control.is_some())
                    && !(namespace.as_deref() == Some(FORM)
                        && matches!(
                            local.as_slice(),
                            b"properties" | b"property" | b"list-property" | b"list-value"
                        ))
                {
                    return invalid("unexpected child element in visual control");
                }
            },
            Event::Text(text)
                if stack.iter().any(|open| open.control.is_some()) => {
                    let decoded = text.decode().map_err(|error| {
                        Error::InvalidFormat(format!("invalid visual control text: {error}"))
                    })?;
                    if !decoded.trim().is_empty() {
                        return invalid("visual controls cannot contain character data");
                    }
                },
            Event::CData(text)
                if stack.iter().any(|open| open.control.is_some()) => {
                    let decoded = text.decode().map_err(|error| {
                        Error::InvalidFormat(format!("invalid visual control CDATA: {error}"))
                    })?;
                    if !decoded.trim().is_empty() {
                        return invalid("visual controls cannot contain CDATA");
                    }
                },
            Event::GeneralRef(_) if stack.iter().any(|open| open.control.is_some()) => {
                return invalid("visual controls cannot contain entity references");
            },
            Event::End(ref element) => {
                let open = stack.pop().ok_or_else(|| {
                    Error::InvalidFormat("visual form XML stack underflow".to_string())
                })?;
                if open.local.as_slice() != element.local_name().as_ref() {
                    return invalid("mismatched visual form XML elements");
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
            Event::DocType(_) => return invalid("DOCTYPE is not allowed in visual form XML"),
            Event::Eof => break,
            _ => {},
        }
        previous = end;
        buffer.clear();
    }
    if !stack.is_empty() {
        return invalid("unclosed visual form XML elements");
    }
    for control in &controls {
        forms[control.form].controls.push(control.clone());
    }
    for form in &forms {
        let values: Vec<_> = form
            .controls
            .iter()
            .map(|item| item.control.clone())
            .collect();
        validate_radio_groups(&values)?;
    }
    Ok(Scan {
        forms,
        controls,
        ids,
    })
}

fn parse_control(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
    local: &[u8],
) -> Result<OdfVisualControl> {
    let attrs = attributes(reader, element)?;
    let allowed = match local {
        b"radio" => RADIO_ATTRS,
        b"frame" => FRAME_ATTRS,
        _ => IMAGE_ATTRS,
    };
    validate_allowed(&attrs, allowed)?;
    let name = required(&attrs, FORM, "name")?;
    let xml_id = required(&attrs, XML, "id")?;
    match local {
        b"radio" => {
            let mut value = OdfRadioControl::new(name, xml_id);
            value.current_selected = optional_bool(&attrs, FORM, "current-selected")?;
            value.disabled = optional_bool(&attrs, FORM, "disabled")?;
            value.label = optional(&attrs, FORM, "label");
            value.printable = optional_bool(&attrs, FORM, "printable")?;
            value.selected = optional_bool(&attrs, FORM, "selected")?;
            value.tab_index = optional_u64(&attrs, FORM, "tab-index")?;
            value.tab_stop = optional_bool(&attrs, FORM, "tab-stop")?;
            value.title = optional(&attrs, FORM, "title");
            value.value = optional(&attrs, FORM, "value");
            value.data_field = optional(&attrs, FORM, "data-field");
            value.visual_effect = optional(&attrs, FORM, "visual-effect")
                .map(|v| OdfRadioVisualEffect::parse(&v))
                .transpose()?;
            value.image_position = optional(&attrs, FORM, "image-position")
                .map(|v| OdfRelativeImagePosition::parse(&v))
                .transpose()?;
            value.image_align = optional(&attrs, FORM, "image-align")
                .map(|v| OdfRelativeImageAlign::parse(&v))
                .transpose()?;
            value.linked_cell = optional(&attrs, FORM, "linked-cell");
            validate_radio(&value)?;
            Ok(value.into())
        },
        b"frame" => {
            let mut value = OdfFrameControl::new(name, xml_id);
            value.disabled = optional_bool(&attrs, FORM, "disabled")?;
            value.form_for = optional(&attrs, FORM, "for");
            value.label = optional(&attrs, FORM, "label");
            value.printable = optional_bool(&attrs, FORM, "printable")?;
            value.title = optional(&attrs, FORM, "title");
            validate_frame(&value)?;
            Ok(value.into())
        },
        _ => {
            let mut value = OdfImageControl::new(name, xml_id);
            value.button_type = optional(&attrs, FORM, "button-type")
                .map(|v| OdfImageButtonType::parse(&v))
                .transpose()?;
            value.disabled = optional_bool(&attrs, FORM, "disabled")?;
            value.image_data = optional(&attrs, FORM, "image-data");
            value.printable = optional_bool(&attrs, FORM, "printable")?;
            value.tab_index = optional_u64(&attrs, FORM, "tab-index")?;
            value.tab_stop = optional_bool(&attrs, FORM, "tab-stop")?;
            value.target_frame = optional(&attrs, OFFICE, "target-frame");
            value.href = optional(&attrs, XLINK, "href");
            value.title = optional(&attrs, FORM, "title");
            value.value = optional(&attrs, FORM, "value");
            validate_image(&value)?;
            Ok(value.into())
        },
    }
}

const RADIO_ATTRS: &[(&str, &str)] = &[
    (FORM, "name"),
    (XML, "id"),
    (FORM, "id"),
    (FORM, "control-implementation"),
    (FORM, "input-required"),
    (XFORMS, "bind"),
    (FORM, "current-selected"),
    (FORM, "disabled"),
    (FORM, "label"),
    (FORM, "printable"),
    (FORM, "selected"),
    (FORM, "tab-index"),
    (FORM, "tab-stop"),
    (FORM, "title"),
    (FORM, "value"),
    (FORM, "data-field"),
    (FORM, "visual-effect"),
    (FORM, "image-position"),
    (FORM, "image-align"),
    (FORM, "linked-cell"),
];
const FRAME_ATTRS: &[(&str, &str)] = &[
    (FORM, "name"),
    (XML, "id"),
    (FORM, "id"),
    (FORM, "control-implementation"),
    (FORM, "input-required"),
    (XFORMS, "bind"),
    (FORM, "disabled"),
    (FORM, "for"),
    (FORM, "label"),
    (FORM, "printable"),
    (FORM, "title"),
];
const IMAGE_ATTRS: &[(&str, &str)] = &[
    (FORM, "name"),
    (XML, "id"),
    (FORM, "id"),
    (FORM, "control-implementation"),
    (FORM, "input-required"),
    (XFORMS, "bind"),
    (FORM, "button-type"),
    (FORM, "disabled"),
    (FORM, "image-data"),
    (FORM, "printable"),
    (FORM, "tab-index"),
    (FORM, "tab-stop"),
    (OFFICE, "target-frame"),
    (XLINK, "href"),
    (XLINK, "type"),
    (XLINK, "actuate"),
    (FORM, "title"),
    (FORM, "value"),
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
            Error::InvalidFormat(format!("invalid visual control attribute: {error}"))
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
                Error::InvalidFormat(format!("invalid visual control attribute value: {error}"))
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
        if ids.iter().any(|value| value == &id) {
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
                "unexpected visual control attribute '{}'",
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

fn validate_identity(name: &str, xml_id: &str) -> Result<()> {
    validate_name("visual control name", name)?;
    validate_xml_id(xml_id)
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
    validate_name("visual control xml:id", value)?;
    let mut chars = value.chars();
    let first = chars.next().unwrap();
    if !(first == '_' || first.is_ascii_alphabetic())
        || !chars.all(|ch| ch == '_' || ch == '-' || ch == '.' || ch.is_ascii_alphanumeric())
    {
        return invalid(format!("invalid visual control xml:id '{value}'"));
    }
    Ok(())
}

fn validate_optionals(values: &[(&str, Option<&str>)]) -> Result<()> {
    for (label, value) in values {
        if let Some(value) = value {
            validate_string(label, value)?;
        }
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

fn validate_uri(label: &str, value: Option<&str>) -> Result<()> {
    let Some(value) = value else { return Ok(()) };
    if value.len() > MAX_URI {
        return invalid(format!("{label} exceeds 8 KiB"));
    }
    validate_string(label, value)?;
    if value.chars().any(char::is_whitespace) || value.contains('\\') {
        return invalid(format!("{label} contains invalid URI characters"));
    }
    let lower = value.to_ascii_lowercase();
    if ["javascript:", "macro:", "vnd.sun.star.script:", "data:"]
        .iter()
        .any(|scheme| lower.starts_with(scheme))
    {
        return invalid(format!("{label} uses an active or inline-data URI scheme"));
    }
    if !value.contains("://")
        && !value.starts_with('#')
        && value.split('/').any(|part| part == "." || part == "..")
    {
        return invalid(format!("{label} contains package path traversal"));
    }
    Ok(())
}

fn is_active_form_content(namespace: Option<&str>, local: &[u8]) -> bool {
    (namespace == Some(OFFICE) && matches!(local, b"event-listeners" | b"events" | b"scripts"))
        || namespace == Some(SCRIPT)
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

fn push_u64(out: &mut String, name: &str, value: Option<u64>) {
    if let Some(value) = value {
        out.push_str(&format!(r#" {name}="{value}""#));
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
    if value.contains("office:") && !value.contains("xmlns:office=") {
        value = value.replacen(' ', &format!(" xmlns:office=\"{OFFICE}\" "), 1);
    }
    if value.contains("xlink:") && !value.contains("xmlns:xlink=") {
        value = value.replacen(' ', &format!(" xmlns:xlink=\"{XLINK}\" "), 1);
    }
    value
}

fn qname(element: &BytesStart<'_>) -> Result<String> {
    String::from_utf8(element.name().as_ref().to_vec())
        .map_err(|_| Error::InvalidFormat("invalid visual form element name".to_string()))
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

    const ROOT: &str = r#"<o:document-content xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:f="urn:oasis:names:tc:opendocument:xmlns:form:1.0" xmlns:l="http://www.w3.org/1999/xlink"><o:body><o:text><o:forms>"#;
    const END: &str = "</o:forms></o:text></o:body></o:document-content>";

    #[test]
    fn canonical_family_round_trips() {
        let mut radio = OdfRadioControl::new("Group", "radio_1");
        radio.label = Some("A & B".into());
        radio.selected = Some(true);
        radio.image_position = Some(OdfRelativeImagePosition::Start);
        radio.image_align = Some(OdfRelativeImageAlign::Center);
        let mut frame = OdfFrameControl::new("Frame", "frame_1");
        frame.form_for = Some("radio_1".into());
        let mut image = OdfImageControl::new("Image", "image_1");
        image.button_type = Some(OdfImageButtonType::Submit);
        image.image_data = Some("Pictures/form.png".into());
        image.href = Some("https://example.invalid/submit".into());
        let mut form = OdfVisualForm::new("Main");
        form.add_control(radio).unwrap();
        form.add_control(frame).unwrap();
        form.add_control(image).unwrap();
        let xml = format!("{ROOT}{}{END}", form.to_xml_fragment().unwrap());
        let parsed = visual_controls(&xml).unwrap();
        assert_eq!(parsed.len(), 3);
        assert!(
            matches!(&parsed[0], OdfVisualControl::Radio(v) if v.label.as_deref() == Some("A & B"))
        );
        assert!(
            matches!(&parsed[2], OdfVisualControl::Image(v) if v.button_type == Some(OdfImageButtonType::Submit))
        );
    }

    #[test]
    fn libreoffice_odfpy_and_odfdo_shapes_parse() {
        let lo = include_str!(
            "../../../../test-data/libreoffice-core/vcl/qa/cppunit/pdfexport/data/tdf159817.fodt"
        );
        let parsed = visual_controls(lo).unwrap();
        assert_eq!(
            parsed
                .iter()
                .filter(|v| matches!(v, OdfVisualControl::Radio(_)))
                .count(),
            3
        );
        let producer = format!(
            r#"{ROOT}<f:form f:name="Producer"><f:frame f:name="odfpy" xml:id="frame" f:label="Group"/><f:image f:name="odfdo" xml:id="image" f:button-type="url" f:image-data="Pictures/form.png" o:target-frame="_blank" l:href="https://example.invalid/"/></f:form>{END}"#
        );
        assert_eq!(visual_controls(&producer).unwrap().len(), 2);
    }

    #[test]
    fn lossless_mutation_and_empty_form_expansion() {
        let xml = format!(
            r#"{ROOT}<f:form f:name="Main"><f:radio f:name="Old" xml:id="old"><f:properties><f:property f:property-name="keep" o:value-type="void"/></f:properties></f:radio><!--keep--><f:text f:name="Text" xml:id="text"/></f:form>{END}"#
        );
        let image: OdfVisualControl = OdfImageControl::new("Image", "image").into();
        let inserted = insert_visual_control_xml(&xml, 0, &image).unwrap();
        assert!(inserted.contains("<!--keep-->") && inserted.contains("f:text"));
        let frame: OdfVisualControl = OdfFrameControl::new("Frame", "frame").into();
        let replaced = replace_visual_control_xml(&inserted, 0, &frame).unwrap();
        let removed = remove_visual_control_xml(&replaced, 1).unwrap();
        assert_eq!(visual_controls(&removed).unwrap(), [frame]);
        let empty = format!(r#"{ROOT}<f:form f:name="Empty"/>{END}"#);
        assert!(
            insert_visual_control_xml(&empty, 0, &image)
                .unwrap()
                .contains("</f:form>")
        );
    }

    #[test]
    fn hostile_values_children_resources_and_active_content_are_rejected() {
        assert!(OdfRadioControl::new("R", "1bad").to_xml_fragment().is_err());
        let bad_token = format!(
            r#"{ROOT}<f:form f:name="Main"><f:radio f:name="R" xml:id="r" f:visual-effect="raised"/></f:form>{END}"#
        );
        assert!(visual_controls(&bad_token).is_err());
        let bad_child = format!(
            r#"{ROOT}<f:form f:name="Main"><f:frame f:name="F" xml:id="f"><o:p/></f:frame></f:form>{END}"#
        );
        assert!(visual_controls(&bad_child).is_err());
        let event = format!(r#"{ROOT}<f:form f:name="Main"><o:event-listeners/></f:form>{END}"#);
        assert!(visual_controls(&event).is_err());
        let mut image = OdfImageControl::new("I", "i");
        image.href = Some("javascript:alert(1)".into());
        assert!(image.to_xml_fragment().is_err());
        image.href = None;
        image.image_data = Some("../Pictures/x.png".into());
        assert!(image.to_xml_fragment().is_err());
    }

    #[test]
    fn radio_group_cardinality_is_atomic() {
        let mut first = OdfRadioControl::new("Group", "r1");
        first.selected = Some(true);
        let mut second = OdfRadioControl::new("Group", "r2");
        second.selected = Some(true);
        let mut form = OdfVisualForm::new("Main");
        form.add_control(first).unwrap();
        assert!(form.add_control(second).is_err());
        assert_eq!(form.controls.len(), 1);
    }

    #[test]
    fn builder_and_mutable_document_round_trip() {
        use crate::odt::{Document, DocumentBuilder, MutableDocument};
        let mut form = OdfVisualForm::new("Main");
        form.add_control(OdfRadioControl::new("Group", "radio"))
            .unwrap();
        let mut builder = DocumentBuilder::new();
        builder.add_visual_form(&form).unwrap();
        builder.add_paragraph("body").unwrap();
        let document = Document::from_bytes(builder.build().unwrap()).unwrap();
        let mut mutable = MutableDocument::from_document(document).unwrap();
        assert_eq!(mutable.visual_controls().unwrap().len(), 1);
        let image: OdfVisualControl = OdfImageControl::new("Image", "image").into();
        mutable.insert_visual_control(0, &image).unwrap();
        let frame: OdfVisualControl = OdfFrameControl::new("Frame", "frame").into();
        assert!(matches!(
            mutable.replace_visual_control(0, &frame).unwrap(),
            OdfVisualControl::Radio(_)
        ));
        assert!(matches!(
            mutable.remove_visual_control(1).unwrap(),
            OdfVisualControl::Image(_)
        ));
        assert_eq!(mutable.visual_controls().unwrap(), [frame]);
    }
}
