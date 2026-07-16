//! Inert, bounded semantic inventory of classic OpenDocument forms.

use litchi_core::{Error, Result};
use quick_xml::XmlVersion;
use quick_xml::events::{BytesRef, BytesStart, Event};
use quick_xml::name::{Namespace, ResolveResult};
use quick_xml::reader::NsReader;
use std::collections::HashMap;

const OFFICE: &str = "urn:oasis:names:tc:opendocument:xmlns:office:1.0";
const FORM: &str = "urn:oasis:names:tc:opendocument:xmlns:form:1.0";
const DRAW: &str = "urn:oasis:names:tc:opendocument:xmlns:drawing:1.0";
const TABLE: &str = "urn:oasis:names:tc:opendocument:xmlns:table:1.0";
const PRESENTATION: &str = "urn:oasis:names:tc:opendocument:xmlns:presentation:1.0";
const STYLE: &str = "urn:oasis:names:tc:opendocument:xmlns:style:1.0";
const SVG: &str = "urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0";
const XML: &str = "http://www.w3.org/XML/1998/namespace";
const XFORMS: &str = "http://www.w3.org/2002/xforms";
const MAX_RAW: usize = 64 * 1024 * 1024;
const MAX_DECODED: usize = 16 * 1024 * 1024;
const MAX_SCALAR: usize = 64 * 1024;
const MAX_TEXT: usize = 4 * 1024 * 1024;
const MAX_NODES: usize = 65_536;
const MAX_SHAPES: usize = 65_536;
const MAX_DEPTH: usize = 128;
const MAX_ATTRIBUTES: usize = 256;

#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum OdfFormPart {
    Content,
    Styles,
    Flat,
}

#[derive(Debug, Clone, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum OdfFormScope {
    Document,
    Text,
    Sheet { index: usize, name: Option<String> },
    DrawPage { index: usize, name: Option<String> },
    Notes { index: usize },
    MasterPage { index: usize, name: Option<String> },
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct OdfFormAttribute {
    pub namespace_uri: Option<String>,
    pub local_name: String,
    pub value: String,
}

#[derive(Debug, Clone, Default, PartialEq)]
pub struct OdfForms {
    pub groups: Vec<OdfFormGroup>,
    pub control_shapes: Vec<OdfControlShape>,
    pub has_xforms: bool,
    pub has_event_listeners: bool,
}

#[derive(Debug, Clone, PartialEq)]
pub struct OdfFormGroup {
    pub part: OdfFormPart,
    pub scope: OdfFormScope,
    pub automatic_focus: Option<bool>,
    pub apply_design_mode: Option<bool>,
    pub forms: Vec<OdfForm>,
    pub has_xforms: bool,
    pub has_event_listeners: bool,
    pub attributes: Vec<OdfFormAttribute>,
}

#[derive(Debug, Clone, PartialEq)]
pub struct OdfForm {
    pub xml_id: Option<String>,
    pub form_id: Option<String>,
    pub name: Option<String>,
    pub control_implementation: Option<String>,
    pub properties: Vec<OdfFormProperty>,
    pub children: Vec<OdfFormNode>,
    pub attributes: Vec<OdfFormAttribute>,
}

#[derive(Debug, Clone, PartialEq)]
pub enum OdfFormNode {
    Form(OdfForm),
    Control(OdfFormControl),
}

#[derive(Debug, Clone, PartialEq, Eq)]
#[non_exhaustive]
pub enum OdfFormControlKind {
    Text,
    TextArea,
    Password,
    File,
    FormattedText,
    Number,
    Date,
    Time,
    FixedText,
    ComboBox,
    Item,
    ListBox,
    Option,
    Button,
    Image,
    CheckBox,
    Radio,
    Frame,
    ImageFrame,
    Hidden,
    Grid,
    Column,
    ValueRange,
    GenericControl,
    Other(String),
}

impl OdfFormControlKind {
    fn parse(name: &str) -> Option<Self> {
        Some(match name {
            "text" => Self::Text,
            "textarea" => Self::TextArea,
            "password" => Self::Password,
            "file" => Self::File,
            "formatted-text" => Self::FormattedText,
            "number" => Self::Number,
            "date" => Self::Date,
            "time" => Self::Time,
            "fixed-text" => Self::FixedText,
            "combobox" => Self::ComboBox,
            "item" => Self::Item,
            "listbox" => Self::ListBox,
            "option" => Self::Option,
            "button" => Self::Button,
            "image" => Self::Image,
            "checkbox" => Self::CheckBox,
            "radio" => Self::Radio,
            "frame" => Self::Frame,
            "image-frame" => Self::ImageFrame,
            "hidden" => Self::Hidden,
            "grid" => Self::Grid,
            "column" => Self::Column,
            "value-range" => Self::ValueRange,
            "generic-control" => Self::GenericControl,
            "form" | "properties" | "property" | "list-property" | "list-value"
            | "connection-resource" => return None,
            other => Self::Other(other.to_string()),
        })
    }
}

#[derive(Debug, Clone, PartialEq)]
pub struct OdfFormControl {
    pub kind: OdfFormControlKind,
    pub xml_id: Option<String>,
    pub form_id: Option<String>,
    pub name: Option<String>,
    pub label: Option<String>,
    pub title: Option<String>,
    pub value: Option<String>,
    pub current_value: Option<String>,
    pub state: Option<String>,
    pub current_state: Option<String>,
    pub linked_cell: Option<String>,
    pub source_cell_range: Option<String>,
    pub image_data: Option<String>,
    pub disabled: Option<bool>,
    pub read_only: Option<bool>,
    pub selected: Option<bool>,
    pub current_selected: Option<bool>,
    pub text: String,
    pub properties: Vec<OdfFormProperty>,
    pub children: Vec<OdfFormNode>,
    pub attributes: Vec<OdfFormAttribute>,
}

#[derive(Debug, Clone, PartialEq)]
pub struct OdfFormProperty {
    pub name: String,
    pub value: OdfFormPropertyValue,
}

#[derive(Debug, Clone, PartialEq)]
pub enum OdfFormPropertyValue {
    Scalar(OdfFormScalarValue),
    List {
        value_type: Option<String>,
        values: Vec<OdfFormScalarValue>,
    },
}

#[derive(Debug, Clone, PartialEq)]
#[non_exhaustive]
pub enum OdfFormScalarValue {
    Boolean(bool),
    Number {
        value_type: String,
        lexical: String,
        currency: Option<String>,
    },
    Text(String),
    Date(String),
    Time(String),
    Void,
    Other {
        value_type: String,
        lexical: Option<String>,
    },
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct OdfControlRef {
    pub group_index: usize,
    pub form_index: usize,
    pub node_path: Vec<usize>,
}

#[derive(Debug, Clone, PartialEq)]
pub struct OdfControlShape {
    pub part: OdfFormPart,
    pub scope: OdfFormScope,
    pub control_id: String,
    pub resolved_control: Option<OdfControlRef>,
    pub draw_name: Option<String>,
    pub style_name: Option<String>,
    pub text_style_name: Option<String>,
    pub z_index: Option<u32>,
    pub x: Option<String>,
    pub y: Option<String>,
    pub width: Option<String>,
    pub height: Option<String>,
    pub attributes: Vec<OdfFormAttribute>,
}

enum Builder {
    Form(usize, OdfForm),
    Control(usize, OdfFormControl),
    List(usize, String, Option<String>, Vec<OdfFormScalarValue>),
}

impl Builder {
    fn depth(&self) -> usize {
        match self {
            Self::Form(depth, _) | Self::Control(depth, _) | Self::List(depth, ..) => *depth,
        }
    }
}

struct ScopeFrame(usize, OdfFormScope);

#[derive(Default)]
struct Limits {
    nodes: usize,
    shapes: usize,
    decoded: usize,
}

pub(crate) fn parse_form_parts(parts: &[(&str, OdfFormPart)]) -> Result<OdfForms> {
    let raw = parts.iter().try_fold(0usize, |sum, (xml, _)| {
        sum.checked_add(xml.len()).ok_or_else(|| err("form XML size overflow"))
    })?;
    if raw > MAX_RAW {
        return Err(err("form XML exceeds 64 MiB"));
    }
    let mut result = OdfForms::default();
    let mut limits = Limits::default();
    for &(xml, part) in parts {
        parse_part(xml, part, &mut result, &mut limits)?;
    }
    resolve_links(&mut result)?;
    Ok(result)
}

fn parse_part(
    xml: &str,
    part: OdfFormPart,
    result: &mut OdfForms,
    limits: &mut Limits,
) -> Result<()> {
    let mut reader = NsReader::from_str(xml);
    let mut buffer = Vec::new();
    let mut depth = 0usize;
    let mut skip = 0usize;
    let mut scopes = Vec::<ScopeFrame>::new();
    let mut builders = Vec::<Builder>::new();
    let mut group: Option<(usize, OdfFormGroup)> = None;
    let (mut sheet, mut page, mut notes, mut master) = (0, 0, 0, 0);
    loop {
        let (resolved, event) = reader
            .read_resolved_event_into(&mut buffer)
            .map_err(|e| err(format!("invalid form XML: {e}")))?;
        match event {
            Event::Start(ref element) => {
                depth = inc_depth(depth)?;
                if skip != 0 {
                    skip = inc_depth(skip)?;
                    buffer.clear();
                    continue;
                }
                let namespace = ns(&resolved)?;
                let local = name(element.local_name().as_ref())?;
                if namespace.as_deref() == Some(XFORMS) && local == "model" {
                    mark_xforms(result, group.as_mut());
                    skip = 1;
                    buffer.clear();
                    continue;
                }
                if namespace.as_deref() == Some(OFFICE)
                    && local == "event-listeners"
                    && group.is_some()
                {
                    mark_events(result, group.as_mut());
                    skip = 1;
                    buffer.clear();
                    continue;
                }
                if let Some(scope) = scope_start(
                    &reader,
                    element,
                    namespace.as_deref(),
                    &local,
                    &mut sheet,
                    &mut page,
                    &mut notes,
                    &mut master,
                    limits,
                )? {
                    scopes.push(ScopeFrame(depth, scope));
                }
                if namespace.as_deref() == Some(OFFICE) && local == "forms" {
                    if group.is_some() {
                        return Err(err("nested office:forms"));
                    }
                    let attrs = attrs(&reader, element, limits)?;
                    group = Some((depth, new_group(part, current_scope(&scopes), attrs)?));
                } else if group.is_some() {
                    form_start(
                        &reader,
                        element,
                        namespace.as_deref(),
                        &local,
                        depth,
                        &mut builders,
                        limits,
                    )?;
                }
                if namespace.as_deref() == Some(DRAW) && local == "control" {
                    let scope = current_scope(&scopes);
                    let shape = new_shape(part, scope, attrs(&reader, element, limits)?)?;
                    push_shape(result, shape, limits)?;
                }
            },
            Event::Empty(ref element) => {
                if skip != 0 {
                    buffer.clear();
                    continue;
                }
                let namespace = ns(&resolved)?;
                let local = name(element.local_name().as_ref())?;
                if namespace.as_deref() == Some(XFORMS) && local == "model" {
                    mark_xforms(result, group.as_mut());
                } else if namespace.as_deref() == Some(OFFICE)
                    && local == "event-listeners"
                    && group.is_some()
                {
                    mark_events(result, group.as_mut());
                } else if namespace.as_deref() == Some(OFFICE) && local == "forms" {
                    let attrs = attrs(&reader, element, limits)?;
                    result.groups.push(new_group(part, current_scope(&scopes), attrs)?);
                } else if group.is_some() {
                    form_empty(
                        &reader,
                        element,
                        namespace.as_deref(),
                        &local,
                        &mut group,
                        &mut builders,
                        limits,
                    )?;
                }
                if namespace.as_deref() == Some(DRAW) && local == "control" {
                    let shape = new_shape(
                        part,
                        current_scope(&scopes),
                        attrs(&reader, element, limits)?,
                    )?;
                    push_shape(result, shape, limits)?;
                }
            },
            Event::End(_) => {
                if skip != 0 {
                    skip -= 1;
                    depth = depth.checked_sub(1).ok_or_else(|| err("XML depth underflow"))?;
                    buffer.clear();
                    continue;
                }
                while builders.last().is_some_and(|builder| builder.depth() == depth) {
                    finish_builder(&mut group, &mut builders)?;
                }
                if group.as_ref().is_some_and(|(at, _)| *at == depth) {
                    if !builders.is_empty() {
                        return Err(err("unclosed form declaration"));
                    }
                    result.groups.push(group.take().expect("active group").1);
                }
                if scopes.last().is_some_and(|scope| scope.0 == depth) {
                    scopes.pop();
                }
                depth = depth.checked_sub(1).ok_or_else(|| err("XML depth underflow"))?;
            },
            Event::Text(ref text) if skip == 0 && group.is_some() => {
                let value = text
                    .xml_content(XmlVersion::Explicit1_0)
                    .map_err(|e| err(format!("invalid form text: {e}")))?;
                append_text(&mut builders, &value, limits)?;
            },
            Event::CData(ref text) if skip == 0 && group.is_some() => {
                let value = text
                    .xml_content(XmlVersion::Explicit1_0)
                    .map_err(|e| err(format!("invalid form CDATA: {e}")))?;
                append_text(&mut builders, &value, limits)?;
            },
            Event::GeneralRef(ref value) if skip == 0 && group.is_some() => {
                append_text(&mut builders, &reference(value)?, limits)?;
            },
            Event::DocType(_) => return Err(err("DOCTYPE is not allowed in form XML")),
            Event::Eof => break,
            _ => {},
        }
        buffer.clear();
    }
    if depth != 0 || skip != 0 || group.is_some() || !builders.is_empty() || !scopes.is_empty() {
        return Err(err("incomplete form XML structure"));
    }
    Ok(())
}

#[allow(clippy::too_many_arguments)]
fn scope_start(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
    namespace: Option<&str>,
    local: &str,
    sheet: &mut usize,
    page: &mut usize,
    notes: &mut usize,
    master: &mut usize,
    limits: &mut Limits,
) -> Result<Option<OdfFormScope>> {
    if namespace == Some(OFFICE) && local == "text" {
        return Ok(Some(OdfFormScope::Text));
    }
    if namespace == Some(TABLE) && local == "table" {
        let index = next_index(sheet)?;
        let attrs = attrs(reader, element, limits)?;
        return Ok(Some(OdfFormScope::Sheet {
            index,
            name: owned(&attrs, TABLE, "name"),
        }));
    }
    if namespace == Some(DRAW) && local == "page" {
        let index = next_index(page)?;
        let attrs = attrs(reader, element, limits)?;
        return Ok(Some(OdfFormScope::DrawPage {
            index,
            name: owned(&attrs, DRAW, "name"),
        }));
    }
    if namespace == Some(PRESENTATION) && local == "notes" {
        return Ok(Some(OdfFormScope::Notes {
            index: next_index(notes)?,
        }));
    }
    if namespace == Some(STYLE) && local == "master-page" {
        let index = next_index(master)?;
        let attrs = attrs(reader, element, limits)?;
        return Ok(Some(OdfFormScope::MasterPage {
            index,
            name: owned(&attrs, STYLE, "name"),
        }));
    }
    Ok(None)
}

fn form_start(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
    namespace: Option<&str>,
    local: &str,
    depth: usize,
    builders: &mut Vec<Builder>,
    limits: &mut Limits,
) -> Result<()> {
    if namespace != Some(FORM) {
        return Ok(());
    }
    let attrs = attrs(reader, element, limits)?;
    match local {
        "form" => {
            node(limits)?;
            builders.push(Builder::Form(depth, new_form(attrs)));
        },
        "property" => {
            node(limits)?;
            attach_property(builders, scalar_property(&attrs)?)?;
        },
        "list-property" => {
            node(limits)?;
            builders.push(list_builder(depth, &attrs)?);
        },
        "list-value" => {
            node(limits)?;
            list_value(builders, &attrs)?;
        },
        _ => {
            if let Some(kind) = OdfFormControlKind::parse(local) {
                node(limits)?;
                builders.push(Builder::Control(depth, new_control(kind, attrs)?));
            }
        },
    }
    Ok(())
}

fn form_empty(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
    namespace: Option<&str>,
    local: &str,
    group: &mut Option<(usize, OdfFormGroup)>,
    builders: &mut Vec<Builder>,
    limits: &mut Limits,
) -> Result<()> {
    if namespace != Some(FORM) {
        return Ok(());
    }
    let attrs = attrs(reader, element, limits)?;
    match local {
        "form" => {
            node(limits)?;
            attach_form(group, builders, new_form(attrs))?;
        },
        "property" => {
            node(limits)?;
            attach_property(builders, scalar_property(&attrs)?)?;
        },
        "list-property" => {
            node(limits)?;
            let Builder::List(_, name, value_type, values) = list_builder(0, &attrs)? else {
                unreachable!()
            };
            attach_property(
                builders,
                OdfFormProperty {
                    name,
                    value: OdfFormPropertyValue::List { value_type, values },
                },
            )?;
        },
        "list-value" => {
            node(limits)?;
            list_value(builders, &attrs)?;
        },
        _ => {
            if let Some(kind) = OdfFormControlKind::parse(local) {
                node(limits)?;
                attach_control(builders, new_control(kind, attrs)?)?;
            }
        },
    }
    Ok(())
}

fn finish_builder(
    group: &mut Option<(usize, OdfFormGroup)>,
    builders: &mut Vec<Builder>,
) -> Result<()> {
    match builders.pop().ok_or_else(|| err("form stack underflow"))? {
        Builder::Form(_, value) => attach_form(group, builders, value),
        Builder::Control(_, value) => attach_control(builders, value),
        Builder::List(_, name, value_type, values) => attach_property(
            builders,
            OdfFormProperty {
                name,
                value: OdfFormPropertyValue::List { value_type, values },
            },
        ),
    }
}

fn attach_form(
    group: &mut Option<(usize, OdfFormGroup)>,
    builders: &mut [Builder],
    form: OdfForm,
) -> Result<()> {
    match builders.last_mut() {
        Some(Builder::Form(_, parent)) => parent.children.push(OdfFormNode::Form(form)),
        Some(_) => return Err(err("form:form has an invalid parent")),
        None => group
            .as_mut()
            .ok_or_else(|| err("form:form outside office:forms"))?
            .1
            .forms
            .push(form),
    }
    Ok(())
}

fn attach_control(builders: &mut [Builder], control: OdfFormControl) -> Result<()> {
    match builders.last_mut() {
        Some(Builder::Form(_, parent)) => parent.children.push(OdfFormNode::Control(control)),
        Some(Builder::Control(_, parent)) => parent.children.push(OdfFormNode::Control(control)),
        _ => return Err(err("form control has an invalid parent")),
    }
    Ok(())
}

fn attach_property(builders: &mut [Builder], property: OdfFormProperty) -> Result<()> {
    for builder in builders.iter_mut().rev() {
        match builder {
            Builder::Form(_, value) => {
                value.properties.push(property);
                return Ok(());
            },
            Builder::Control(_, value) => {
                value.properties.push(property);
                return Ok(());
            },
            Builder::List(..) => {},
        }
    }
    Err(err("form property has no parent"))
}

fn append_text(builders: &mut [Builder], text: &str, limits: &mut Limits) -> Result<()> {
    for builder in builders.iter_mut().rev() {
        if let Builder::Control(_, control) = builder {
            if control.text.len().checked_add(text.len()).is_none_or(|size| size > MAX_TEXT) {
                return Err(err("form control text exceeds 4 MiB"));
            }
            decoded(limits, text.len())?;
            control.text.push_str(text);
            break;
        }
    }
    Ok(())
}

fn new_group(
    part: OdfFormPart,
    scope: OdfFormScope,
    attributes: Vec<OdfFormAttribute>,
) -> Result<OdfFormGroup> {
    Ok(OdfFormGroup {
        part,
        scope,
        automatic_focus: bool_attr(&attributes, FORM, "automatic-focus")?,
        apply_design_mode: bool_attr(&attributes, FORM, "apply-design-mode")?,
        forms: Vec::new(),
        has_xforms: false,
        has_event_listeners: false,
        attributes,
    })
}

fn new_form(attributes: Vec<OdfFormAttribute>) -> OdfForm {
    OdfForm {
        xml_id: owned(&attributes, XML, "id"),
        form_id: owned(&attributes, FORM, "id"),
        name: owned(&attributes, FORM, "name"),
        control_implementation: owned(&attributes, FORM, "control-implementation"),
        properties: Vec::new(),
        children: Vec::new(),
        attributes,
    }
}

fn new_control(
    kind: OdfFormControlKind,
    attributes: Vec<OdfFormAttribute>,
) -> Result<OdfFormControl> {
    Ok(OdfFormControl {
        kind,
        xml_id: owned(&attributes, XML, "id"),
        form_id: owned(&attributes, FORM, "id"),
        name: owned(&attributes, FORM, "name"),
        label: owned(&attributes, FORM, "label"),
        title: owned(&attributes, FORM, "title"),
        value: owned(&attributes, FORM, "value"),
        current_value: owned(&attributes, FORM, "current-value"),
        state: owned(&attributes, FORM, "state"),
        current_state: owned(&attributes, FORM, "current-state"),
        linked_cell: owned(&attributes, FORM, "linked-cell"),
        source_cell_range: owned(&attributes, FORM, "source-cell-range"),
        image_data: owned(&attributes, FORM, "image-data"),
        disabled: bool_attr(&attributes, FORM, "disabled")?,
        read_only: bool_attr(&attributes, FORM, "readonly")?,
        selected: bool_attr(&attributes, FORM, "selected")?,
        current_selected: bool_attr(&attributes, FORM, "current-selected")?,
        text: String::new(),
        properties: Vec::new(),
        children: Vec::new(),
        attributes,
    })
}

fn list_builder(depth: usize, attributes: &[OdfFormAttribute]) -> Result<Builder> {
    Ok(Builder::List(
        depth,
        required(attributes, FORM, "property-name")?.to_string(),
        attr(attributes, OFFICE, "value-type").map(str::to_owned),
        Vec::new(),
    ))
}

fn list_value(builders: &mut [Builder], attributes: &[OdfFormAttribute]) -> Result<()> {
    let Some(Builder::List(_, _, inherited, values)) = builders.last_mut() else {
        return Err(err("form:list-value outside form:list-property"));
    };
    values.push(scalar(attributes, inherited.as_deref())?);
    Ok(())
}

fn scalar_property(attributes: &[OdfFormAttribute]) -> Result<OdfFormProperty> {
    Ok(OdfFormProperty {
        name: required(attributes, FORM, "property-name")?.to_string(),
        value: OdfFormPropertyValue::Scalar(scalar(attributes, None)?),
    })
}

fn scalar(attributes: &[OdfFormAttribute], inherited: Option<&str>) -> Result<OdfFormScalarValue> {
    let kind = attr(attributes, OFFICE, "value-type")
        .or(inherited)
        .ok_or_else(|| err("form property requires office:value-type"))?;
    Ok(match kind {
        "boolean" => OdfFormScalarValue::Boolean(
            required(attributes, OFFICE, "boolean-value")?
                .parse::<bool>()
                .map_err(|_| err("invalid form boolean property"))?,
        ),
        "float" | "percentage" | "currency" => {
            let lexical = required(attributes, OFFICE, "value")?;
            let parsed = lexical
                .parse::<f64>()
                .map_err(|_| err("invalid numeric form property"))?;
            if !parsed.is_finite() {
                return Err(err("non-finite numeric form property"));
            }
            OdfFormScalarValue::Number {
                value_type: kind.to_string(),
                lexical: lexical.to_string(),
                currency: owned(attributes, OFFICE, "currency"),
            }
        },
        "string" => OdfFormScalarValue::Text(
            attr(attributes, OFFICE, "string-value")
                .unwrap_or_default()
                .to_string(),
        ),
        "date" => OdfFormScalarValue::Date(
            required(attributes, OFFICE, "date-value")?.to_string(),
        ),
        "time" => OdfFormScalarValue::Time(
            required(attributes, OFFICE, "time-value")?.to_string(),
        ),
        "void" => OdfFormScalarValue::Void,
        other => OdfFormScalarValue::Other {
            value_type: other.to_string(),
            lexical: attr(attributes, OFFICE, "string-value")
                .or_else(|| attr(attributes, OFFICE, "value"))
                .map(str::to_owned),
        },
    })
}

fn new_shape(
    part: OdfFormPart,
    scope: OdfFormScope,
    attributes: Vec<OdfFormAttribute>,
) -> Result<OdfControlShape> {
    let control_id = required(&attributes, DRAW, "control")?.to_string();
    if control_id.is_empty() {
        return Err(err("empty draw:control reference"));
    }
    Ok(OdfControlShape {
        part,
        scope,
        control_id,
        resolved_control: None,
        draw_name: owned(&attributes, DRAW, "name"),
        style_name: owned(&attributes, DRAW, "style-name"),
        text_style_name: owned(&attributes, DRAW, "text-style-name"),
        z_index: attr(&attributes, DRAW, "z-index")
            .map(|value| value.parse::<u32>().map_err(|_| err("invalid draw:z-index")))
            .transpose()?,
        x: owned(&attributes, SVG, "x"),
        y: owned(&attributes, SVG, "y"),
        width: owned(&attributes, SVG, "width"),
        height: owned(&attributes, SVG, "height"),
        attributes,
    })
}

fn push_shape(result: &mut OdfForms, shape: OdfControlShape, limits: &mut Limits) -> Result<()> {
    limits.shapes = limits.shapes.checked_add(1).ok_or_else(|| err("shape overflow"))?;
    if limits.shapes > MAX_SHAPES {
        return Err(err("document exceeds 65536 form control shapes"));
    }
    result.control_shapes.push(shape);
    Ok(())
}

fn resolve_links(result: &mut OdfForms) -> Result<()> {
    type Key = (OdfFormPart, OdfFormScope, String);
    let mut index = HashMap::<Key, OdfControlRef>::new();
    for (group_index, group) in result.groups.iter().enumerate() {
        for (form_index, form) in group.forms.iter().enumerate() {
            collect_form(
                form,
                group_index,
                form_index,
                group.part,
                &group.scope,
                &mut Vec::new(),
                &mut index,
            )?;
        }
    }
    for shape in &mut result.control_shapes {
        shape.resolved_control = Some(
            index
                .get(&(shape.part, shape.scope.clone(), shape.control_id.clone()))
                .cloned()
                .ok_or_else(|| err(format!("unresolved draw:control '{}'", shape.control_id)))?,
        );
    }
    Ok(())
}

#[allow(clippy::too_many_arguments)]
fn collect_form(
    form: &OdfForm,
    group: usize,
    form_index: usize,
    part: OdfFormPart,
    scope: &OdfFormScope,
    path: &mut Vec<usize>,
    index: &mut HashMap<(OdfFormPart, OdfFormScope, String), OdfControlRef>,
) -> Result<()> {
    for (position, node) in form.children.iter().enumerate() {
        path.push(position);
        match node {
            OdfFormNode::Form(value) => {
                collect_form(value, group, form_index, part, scope, path, index)?
            },
            OdfFormNode::Control(value) => {
                collect_control(value, group, form_index, part, scope, path, index)?
            },
        }
        path.pop();
    }
    Ok(())
}

#[allow(clippy::too_many_arguments)]
fn collect_control(
    control: &OdfFormControl,
    group: usize,
    form_index: usize,
    part: OdfFormPart,
    scope: &OdfFormScope,
    path: &mut Vec<usize>,
    index: &mut HashMap<(OdfFormPart, OdfFormScope, String), OdfControlRef>,
) -> Result<()> {
    let reference = OdfControlRef {
        group_index: group,
        form_index,
        node_path: path.clone(),
    };
    for id in [control.xml_id.as_deref(), control.form_id.as_deref()]
        .into_iter()
        .flatten()
    {
        if id.is_empty() {
            return Err(err("empty form control ID"));
        }
        let key = (part, scope.clone(), id.to_string());
        if let Some(previous) = index.insert(key, reference.clone()) {
            if previous != reference {
                return Err(err(format!("duplicate form control ID '{id}'")));
            }
        }
    }
    for (position, node) in control.children.iter().enumerate() {
        path.push(position);
        match node {
            OdfFormNode::Control(value) => {
                collect_control(value, group, form_index, part, scope, path, index)?
            },
            OdfFormNode::Form(_) => return Err(err("form nested inside control")),
        }
        path.pop();
    }
    Ok(())
}

fn attrs(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
    limits: &mut Limits,
) -> Result<Vec<OdfFormAttribute>> {
    let mut result = Vec::new();
    for raw in element.attributes().with_checks(true) {
        let raw = raw.map_err(|e| err(format!("invalid form attribute: {e}")))?;
        if result.len() >= MAX_ATTRIBUTES {
            return Err(err("form element exceeds 256 attributes"));
        }
        let (resolved, local) = reader.resolver().resolve_attribute(raw.key);
        let value = raw
            .decoded_and_normalized_value(XmlVersion::Explicit1_0, reader.decoder())
            .map_err(|e| err(format!("invalid form attribute value: {e}")))?
            .into_owned();
        if value.len() > MAX_SCALAR {
            return Err(err("form attribute exceeds 64 KiB"));
        }
        decoded(limits, value.len())?;
        result.push(OdfFormAttribute {
            namespace_uri: ns(&resolved)?,
            local_name: name(local.as_ref())?,
            value,
        });
    }
    Ok(result)
}

fn attr<'a>(attributes: &'a [OdfFormAttribute], namespace: &str, local: &str) -> Option<&'a str> {
    attributes
        .iter()
        .find(|item| item.namespace_uri.as_deref() == Some(namespace) && item.local_name == local)
        .map(|item| item.value.as_str())
}

fn required<'a>(
    attributes: &'a [OdfFormAttribute],
    namespace: &str,
    local: &str,
) -> Result<&'a str> {
    attr(attributes, namespace, local)
        .ok_or_else(|| err(format!("missing required form attribute '{local}'")))
}

fn owned(attributes: &[OdfFormAttribute], namespace: &str, local: &str) -> Option<String> {
    attr(attributes, namespace, local).map(str::to_owned)
}

fn bool_attr(
    attributes: &[OdfFormAttribute],
    namespace: &str,
    local: &str,
) -> Result<Option<bool>> {
    attr(attributes, namespace, local)
        .map(|value| {
            value
                .parse::<bool>()
                .map_err(|_| err(format!("invalid boolean form attribute '{local}'")))
        })
        .transpose()
}

fn current_scope(scopes: &[ScopeFrame]) -> OdfFormScope {
    scopes
        .last()
        .map(|scope| scope.1.clone())
        .unwrap_or(OdfFormScope::Document)
}

fn mark_xforms(result: &mut OdfForms, group: Option<&mut (usize, OdfFormGroup)>) {
    result.has_xforms = true;
    if let Some(group) = group {
        group.1.has_xforms = true;
    }
}

fn mark_events(result: &mut OdfForms, group: Option<&mut (usize, OdfFormGroup)>) {
    result.has_event_listeners = true;
    if let Some(group) = group {
        group.1.has_event_listeners = true;
    }
}

fn next_index(value: &mut usize) -> Result<usize> {
    let current = *value;
    *value = value.checked_add(1).ok_or_else(|| err("form scope count overflow"))?;
    Ok(current)
}

fn inc_depth(value: usize) -> Result<usize> {
    let value = value.checked_add(1).ok_or_else(|| err("form XML depth overflow"))?;
    if value > MAX_DEPTH {
        return Err(err("form XML nesting exceeds 128 levels"));
    }
    Ok(value)
}

fn node(limits: &mut Limits) -> Result<()> {
    limits.nodes = limits.nodes.checked_add(1).ok_or_else(|| err("form node overflow"))?;
    if limits.nodes > MAX_NODES {
        return Err(err("document exceeds 65536 form nodes"));
    }
    Ok(())
}

fn decoded(limits: &mut Limits, amount: usize) -> Result<()> {
    limits.decoded = limits
        .decoded
        .checked_add(amount)
        .ok_or_else(|| err("form metadata size overflow"))?;
    if limits.decoded > MAX_DECODED {
        return Err(err("decoded form metadata exceeds 16 MiB"));
    }
    Ok(())
}

fn ns(value: &ResolveResult<'_>) -> Result<Option<String>> {
    match value {
        ResolveResult::Bound(Namespace(value)) => Ok(Some(
            std::str::from_utf8(value)
                .map_err(|_| err("invalid form namespace URI"))?
                .to_string(),
        )),
        ResolveResult::Unbound => Ok(None),
        ResolveResult::Unknown(_) => Err(err("unknown form XML namespace prefix")),
    }
}

fn name(value: &[u8]) -> Result<String> {
    std::str::from_utf8(value)
        .map(str::to_owned)
        .map_err(|_| err("invalid UTF-8 in form XML name"))
}

fn reference(value: &BytesRef<'_>) -> Result<String> {
    if let Some(character) = value
        .resolve_char_ref()
        .map_err(|e| err(format!("invalid form character reference: {e}")))?
    {
        return Ok(character.to_string());
    }
    let entity_name: &[u8] = value.as_ref();
    Ok(match entity_name {
        b"amp" => '&',
        b"lt" => '<',
        b"gt" => '>',
        b"apos" => '\'',
        b"quot" => '"',
        _ => return Err(err("unsupported entity in form content")),
    }
    .to_string())
}

fn err(message: impl Into<String>) -> Error {
    Error::InvalidFormat(message.into())
}
