//! Bounded XML codec for classic OpenDocument forms.

use super::{
    Actuate, Attribute, Control, ControlKind, ControlRef, DRAW, FORM, Form, Forms, Group, Listener,
    MAX_ATTRIBUTES, MAX_DECODED, MAX_DEPTH, MAX_NODES, MAX_RAW, MAX_SCALAR, MAX_SHAPES, MAX_TEXT,
    Node, OFFICE, PRESENTATION, Part, Property, PropertyValue, SCRIPT, STYLE, SVG, ScalarValue,
    Scope, Shape, TABLE, Target, XFORMS, XLINK, XML,
};
use litchi_core::{Error, Result};
use quick_xml::XmlVersion;
use quick_xml::events::{BytesRef, BytesStart, Event};
use quick_xml::name::{Namespace, ResolveResult};
use quick_xml::reader::NsReader;
use std::collections::HashMap;

#[allow(clippy::large_enum_variant)] // parse-stack builder; boxing would churn many match sites
enum Builder {
    Form(usize, Form),
    Control(usize, Control),
    List(usize, String, Option<String>, Vec<ScalarValue>),
}

impl Builder {
    fn depth(&self) -> usize {
        match self {
            Self::Form(depth, _) | Self::Control(depth, _) | Self::List(depth, ..) => *depth,
        }
    }
}

struct ScopeFrame(usize, Scope);

struct EventListenersBuilder {
    depth: usize,
    listener_depth: Option<usize>,
    target: Target,
    listeners: Vec<Listener>,
}

#[derive(Default)]
struct Limits {
    nodes: usize,
    shapes: usize,
    decoded: usize,
}

pub(crate) fn parse_form_parts(parts: &[(&str, Part)]) -> Result<Forms> {
    let raw = parts.iter().try_fold(0usize, |sum, (xml, _)| {
        sum.checked_add(xml.len())
            .ok_or_else(|| err("form XML size overflow"))
    })?;
    if raw > MAX_RAW {
        return Err(err("form XML exceeds 64 MiB"));
    }
    let mut result = Forms::default();
    let mut limits = Limits::default();
    for &(xml, part) in parts {
        parse_part(xml, part, &mut result, &mut limits)?;
    }
    resolve_links(&mut result)?;
    Ok(result)
}

fn parse_part(xml: &str, part: Part, result: &mut Forms, limits: &mut Limits) -> Result<()> {
    let mut reader = NsReader::from_str(xml);
    let mut buffer = Vec::new();
    let mut depth = 0usize;
    let mut skip = 0usize;
    let mut scopes = Vec::<ScopeFrame>::new();
    let mut builders = Vec::<Builder>::new();
    let mut event_listeners: Option<EventListenersBuilder> = None;
    let mut group: Option<(usize, Group)> = None;
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
                if let Some(events) = event_listeners.as_mut() {
                    if events.listener_depth.is_some()
                        || depth != events.depth + 1
                        || namespace.as_deref() != Some(SCRIPT)
                        || local != "event-listener"
                    {
                        return Err(err("invalid child in office:event-listeners"));
                    }
                    node(limits)?;
                    events.listeners.push(new_event_listener(
                        events.target.clone(),
                        attrs(&reader, element, limits)?,
                    )?);
                    events.listener_depth = Some(depth);
                    buffer.clear();
                    continue;
                }
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
                    event_listeners = Some(EventListenersBuilder {
                        depth,
                        listener_depth: None,
                        target: event_target(part, current_scope(&scopes), &builders)?,
                        listeners: Vec::new(),
                    });
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
                if let Some(events) = event_listeners.as_mut() {
                    if events.listener_depth.is_some()
                        || depth != events.depth
                        || namespace.as_deref() != Some(SCRIPT)
                        || local != "event-listener"
                    {
                        return Err(err("invalid child in office:event-listeners"));
                    }
                    node(limits)?;
                    events.listeners.push(new_event_listener(
                        events.target.clone(),
                        attrs(&reader, element, limits)?,
                    )?);
                    buffer.clear();
                    continue;
                }
                if namespace.as_deref() == Some(XFORMS) && local == "model" {
                    mark_xforms(result, group.as_mut());
                } else if namespace.as_deref() == Some(OFFICE)
                    && local == "event-listeners"
                    && group.is_some()
                {
                    mark_events(result, group.as_mut());
                } else if namespace.as_deref() == Some(OFFICE) && local == "forms" {
                    let attrs = attrs(&reader, element, limits)?;
                    result
                        .groups
                        .push(new_group(part, current_scope(&scopes), attrs)?);
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
            Event::End(ref element) => {
                if let Some(events) = event_listeners.as_mut() {
                    let namespace = ns(&resolved)?;
                    let local = name(element.local_name().as_ref())?;
                    if events.listener_depth == Some(depth) {
                        if namespace.as_deref() != Some(SCRIPT) || local != "event-listener" {
                            return Err(err("invalid script:event-listener end element"));
                        }
                        events.listener_depth = None;
                        depth = depth
                            .checked_sub(1)
                            .ok_or_else(|| err("XML depth underflow"))?;
                        buffer.clear();
                        continue;
                    }
                    if events.depth == depth {
                        if namespace.as_deref() != Some(OFFICE) || local != "event-listeners" {
                            return Err(err("invalid office:event-listeners end element"));
                        }
                        let completed = event_listeners.take().expect("active event listeners");
                        result.event_listeners.extend(completed.listeners);
                        depth = depth
                            .checked_sub(1)
                            .ok_or_else(|| err("XML depth underflow"))?;
                        buffer.clear();
                        continue;
                    }
                    return Err(err("invalid nesting in office:event-listeners"));
                }
                if skip != 0 {
                    skip -= 1;
                    depth = depth
                        .checked_sub(1)
                        .ok_or_else(|| err("XML depth underflow"))?;
                    buffer.clear();
                    continue;
                }
                while builders
                    .last()
                    .is_some_and(|builder| builder.depth() == depth)
                {
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
                depth = depth
                    .checked_sub(1)
                    .ok_or_else(|| err("XML depth underflow"))?;
            },
            Event::Text(ref text) if skip == 0 && group.is_some() => {
                if event_listeners.is_some() {
                    let bytes: &[u8] = text.as_ref();
                    if bytes.iter().any(|byte| !byte.is_ascii_whitespace()) {
                        return Err(err("office:event-listeners cannot contain text"));
                    }
                    buffer.clear();
                    continue;
                }
                let value = text
                    .xml_content(XmlVersion::Explicit1_0)
                    .map_err(|e| err(format!("invalid form text: {e}")))?;
                append_text(&mut builders, &value, limits)?;
            },
            Event::CData(ref text) if skip == 0 && group.is_some() => {
                if event_listeners.is_some() {
                    let bytes: &[u8] = text.as_ref();
                    if bytes.iter().any(|byte| !byte.is_ascii_whitespace()) {
                        return Err(err("office:event-listeners cannot contain CDATA"));
                    }
                    buffer.clear();
                    continue;
                }
                let value = text
                    .xml_content(XmlVersion::Explicit1_0)
                    .map_err(|e| err(format!("invalid form CDATA: {e}")))?;
                append_text(&mut builders, &value, limits)?;
            },
            Event::GeneralRef(ref value) if skip == 0 && group.is_some() => {
                if event_listeners.is_some() {
                    return Err(err(
                        "office:event-listeners cannot contain entity references",
                    ));
                }
                append_text(&mut builders, &reference(value)?, limits)?;
            },
            Event::PI(_) if event_listeners.is_some() => {
                return Err(err(
                    "processing instructions are not allowed in event listeners",
                ));
            },
            Event::DocType(_) => return Err(err("DOCTYPE is not allowed in form XML")),
            Event::Eof => break,
            _ => {},
        }
        buffer.clear();
    }
    if depth != 0
        || skip != 0
        || event_listeners.is_some()
        || group.is_some()
        || !builders.is_empty()
        || !scopes.is_empty()
    {
        return Err(err("incomplete form XML structure"));
    }
    Ok(())
}

fn event_target(part: Part, scope: Scope, builders: &[Builder]) -> Result<Target> {
    Ok(match builders.last() {
        Some(Builder::Form(_, form)) => Target::Form {
            xml_id: form.xml_id.clone(),
            form_id: form.form_id.clone(),
            name: form.name.clone(),
        },
        Some(Builder::Control(_, control)) => Target::Control {
            kind: control.kind.clone(),
            xml_id: control.xml_id.clone(),
            form_id: control.form_id.clone(),
            name: control.name.clone(),
        },
        Some(Builder::List(..)) => {
            return Err(err(
                "office:event-listeners cannot be nested in a form property",
            ));
        },
        None => Target::Forms { part, scope },
    })
}

fn new_event_listener(target: Target, attributes: Vec<Attribute>) -> Result<Listener> {
    let event_name = owned(&attributes, SCRIPT, "event-name");
    let language = owned(&attributes, SCRIPT, "language");
    if event_name.as_deref().is_some_and(str::is_empty)
        || language.as_deref().is_some_and(str::is_empty)
    {
        return Err(err("event listener name and language must not be empty"));
    }
    let macro_name = owned(&attributes, SCRIPT, "macro-name");
    if macro_name.as_deref().is_some_and(str::is_empty) {
        return Err(err("script:macro-name must not be empty"));
    }
    let href = owned(&attributes, XLINK, "href");
    if href.as_deref().is_some_and(str::is_empty) {
        return Err(err("xlink:href must not be empty"));
    }
    let simple_link = match owned(&attributes, XLINK, "type") {
        Some(value) if value == "simple" => true,
        Some(_) => return Err(err("script:event-listener xlink:type must be 'simple'")),
        None => false,
    };
    let actuate = owned(&attributes, XLINK, "actuate")
        .map(|value| match value.as_str() {
            "onLoad" => Ok(Actuate::OnLoad),
            "onRequest" => Ok(Actuate::OnRequest),
            "other" => Ok(Actuate::Other),
            "none" => Ok(Actuate::None),
            _ => Err(err("invalid xlink:actuate on script:event-listener")),
        })
        .transpose()?;
    Ok(Listener {
        target,
        event_name,
        language,
        macro_name,
        href,
        actuate,
        simple_link,
        attributes,
    })
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
) -> Result<Option<Scope>> {
    if namespace == Some(OFFICE) && local == "text" {
        return Ok(Some(Scope::Text));
    }
    if namespace == Some(TABLE) && local == "table" {
        let index = next_index(sheet)?;
        let attrs = attrs(reader, element, limits)?;
        return Ok(Some(Scope::Sheet {
            index,
            name: owned(&attrs, TABLE, "name"),
        }));
    }
    if namespace == Some(DRAW) && local == "page" {
        let index = next_index(page)?;
        let attrs = attrs(reader, element, limits)?;
        return Ok(Some(Scope::DrawPage {
            index,
            name: owned(&attrs, DRAW, "name"),
        }));
    }
    if namespace == Some(PRESENTATION) && local == "notes" {
        return Ok(Some(Scope::Notes {
            index: next_index(notes)?,
        }));
    }
    if namespace == Some(STYLE) && local == "master-page" {
        let index = next_index(master)?;
        let attrs = attrs(reader, element, limits)?;
        return Ok(Some(Scope::MasterPage {
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
            if let Some(kind) = ControlKind::parse(local) {
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
    group: &mut Option<(usize, Group)>,
    builders: &mut [Builder],
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
                Property {
                    name,
                    value: PropertyValue::List { value_type, values },
                },
            )?;
        },
        "list-value" => {
            node(limits)?;
            list_value(builders, &attrs)?;
        },
        _ => {
            if let Some(kind) = ControlKind::parse(local) {
                node(limits)?;
                attach_control(builders, new_control(kind, attrs)?)?;
            }
        },
    }
    Ok(())
}

fn finish_builder(group: &mut Option<(usize, Group)>, builders: &mut Vec<Builder>) -> Result<()> {
    match builders.pop().ok_or_else(|| err("form stack underflow"))? {
        Builder::Form(_, value) => attach_form(group, builders, value),
        Builder::Control(_, value) => attach_control(builders, value),
        Builder::List(_, name, value_type, values) => attach_property(
            builders,
            Property {
                name,
                value: PropertyValue::List { value_type, values },
            },
        ),
    }
}

fn attach_form(
    group: &mut Option<(usize, Group)>,
    builders: &mut [Builder],
    form: Form,
) -> Result<()> {
    match builders.last_mut() {
        Some(Builder::Form(_, parent)) => parent.children.push(Node::Form(form)),
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

fn attach_control(builders: &mut [Builder], control: Control) -> Result<()> {
    match builders.last_mut() {
        Some(Builder::Form(_, parent)) => parent.children.push(Node::Control(control)),
        Some(Builder::Control(_, parent)) => parent.children.push(Node::Control(control)),
        _ => return Err(err("form control has an invalid parent")),
    }
    Ok(())
}

fn attach_property(builders: &mut [Builder], property: Property) -> Result<()> {
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
            if control
                .text
                .len()
                .checked_add(text.len())
                .is_none_or(|size| size > MAX_TEXT)
            {
                return Err(err("form control text exceeds 4 MiB"));
            }
            decoded(limits, text.len())?;
            control.text.push_str(text);
            break;
        }
    }
    Ok(())
}

fn new_group(part: Part, scope: Scope, attributes: Vec<Attribute>) -> Result<Group> {
    Ok(Group {
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

fn new_form(attributes: Vec<Attribute>) -> Form {
    Form {
        xml_id: owned(&attributes, XML, "id"),
        form_id: owned(&attributes, FORM, "id"),
        name: owned(&attributes, FORM, "name"),
        control_implementation: owned(&attributes, FORM, "control-implementation"),
        properties: Vec::new(),
        children: Vec::new(),
        attributes,
    }
}

fn new_control(kind: ControlKind, attributes: Vec<Attribute>) -> Result<Control> {
    Ok(Control {
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

fn list_builder(depth: usize, attributes: &[Attribute]) -> Result<Builder> {
    Ok(Builder::List(
        depth,
        required(attributes, FORM, "property-name")?.to_string(),
        attr(attributes, OFFICE, "value-type").map(str::to_owned),
        Vec::new(),
    ))
}

fn list_value(builders: &mut [Builder], attributes: &[Attribute]) -> Result<()> {
    let Some(Builder::List(_, _, inherited, values)) = builders.last_mut() else {
        return Err(err("form:list-value outside form:list-property"));
    };
    values.push(scalar(attributes, inherited.as_deref())?);
    Ok(())
}

fn scalar_property(attributes: &[Attribute]) -> Result<Property> {
    Ok(Property {
        name: required(attributes, FORM, "property-name")?.to_string(),
        value: PropertyValue::Scalar(scalar(attributes, None)?),
    })
}

fn scalar(attributes: &[Attribute], inherited: Option<&str>) -> Result<ScalarValue> {
    let kind = attr(attributes, OFFICE, "value-type")
        .or(inherited)
        .ok_or_else(|| err("form property requires office:value-type"))?;
    Ok(match kind {
        "boolean" => ScalarValue::Boolean(
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
            ScalarValue::Number {
                value_type: kind.to_string(),
                lexical: lexical.to_string(),
                currency: owned(attributes, OFFICE, "currency"),
            }
        },
        "string" => ScalarValue::Text(
            attr(attributes, OFFICE, "string-value")
                .unwrap_or_default()
                .to_string(),
        ),
        "date" => ScalarValue::Date(required(attributes, OFFICE, "date-value")?.to_string()),
        "time" => ScalarValue::Time(required(attributes, OFFICE, "time-value")?.to_string()),
        "void" => ScalarValue::Void,
        other => ScalarValue::Other {
            value_type: other.to_string(),
            lexical: attr(attributes, OFFICE, "string-value")
                .or_else(|| attr(attributes, OFFICE, "value"))
                .map(str::to_owned),
        },
    })
}

fn new_shape(part: Part, scope: Scope, attributes: Vec<Attribute>) -> Result<Shape> {
    let control_id = required(&attributes, DRAW, "control")?.to_string();
    if control_id.is_empty() {
        return Err(err("empty draw:control reference"));
    }
    Ok(Shape {
        part,
        scope,
        control_id,
        resolved_control: None,
        draw_name: owned(&attributes, DRAW, "name"),
        style_name: owned(&attributes, DRAW, "style-name"),
        text_style_name: owned(&attributes, DRAW, "text-style-name"),
        z_index: attr(&attributes, DRAW, "z-index")
            .map(|value| {
                value
                    .parse::<u32>()
                    .map_err(|_| err("invalid draw:z-index"))
            })
            .transpose()?,
        x: owned(&attributes, SVG, "x"),
        y: owned(&attributes, SVG, "y"),
        width: owned(&attributes, SVG, "width"),
        height: owned(&attributes, SVG, "height"),
        attributes,
    })
}

fn push_shape(result: &mut Forms, shape: Shape, limits: &mut Limits) -> Result<()> {
    limits.shapes = limits
        .shapes
        .checked_add(1)
        .ok_or_else(|| err("shape overflow"))?;
    if limits.shapes > MAX_SHAPES {
        return Err(err("document exceeds 65536 form control shapes"));
    }
    result.control_shapes.push(shape);
    Ok(())
}

fn resolve_links(result: &mut Forms) -> Result<()> {
    type Key = (Part, Scope, String);
    let mut index = HashMap::<Key, ControlRef>::new();
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
    form: &Form,
    group: usize,
    form_index: usize,
    part: Part,
    scope: &Scope,
    path: &mut Vec<usize>,
    index: &mut HashMap<(Part, Scope, String), ControlRef>,
) -> Result<()> {
    for (position, node) in form.children.iter().enumerate() {
        path.push(position);
        match node {
            Node::Form(value) => collect_form(value, group, form_index, part, scope, path, index)?,
            Node::Control(value) => {
                collect_control(value, group, form_index, part, scope, path, index)?
            },
        }
        path.pop();
    }
    Ok(())
}

#[allow(clippy::too_many_arguments)]
fn collect_control(
    control: &Control,
    group: usize,
    form_index: usize,
    part: Part,
    scope: &Scope,
    path: &mut Vec<usize>,
    index: &mut HashMap<(Part, Scope, String), ControlRef>,
) -> Result<()> {
    let reference = ControlRef {
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
        if let Some(previous) = index.insert(key, reference.clone())
            && previous != reference
        {
            return Err(err(format!("duplicate form control ID '{id}'")));
        }
    }
    for (position, node) in control.children.iter().enumerate() {
        path.push(position);
        match node {
            Node::Control(value) => {
                collect_control(value, group, form_index, part, scope, path, index)?
            },
            Node::Form(_) => return Err(err("form nested inside control")),
        }
        path.pop();
    }
    Ok(())
}

fn attrs(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
    limits: &mut Limits,
) -> Result<Vec<Attribute>> {
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
        result.push(Attribute {
            namespace_uri: ns(&resolved)?,
            local_name: name(local.as_ref())?,
            value,
        });
    }
    Ok(result)
}

fn attr<'a>(attributes: &'a [Attribute], namespace: &str, local: &str) -> Option<&'a str> {
    attributes
        .iter()
        .find(|item| item.namespace_uri.as_deref() == Some(namespace) && item.local_name == local)
        .map(|item| item.value.as_str())
}

fn required<'a>(attributes: &'a [Attribute], namespace: &str, local: &str) -> Result<&'a str> {
    attr(attributes, namespace, local)
        .ok_or_else(|| err(format!("missing required form attribute '{local}'")))
}

fn owned(attributes: &[Attribute], namespace: &str, local: &str) -> Option<String> {
    attr(attributes, namespace, local).map(str::to_owned)
}

fn bool_attr(attributes: &[Attribute], namespace: &str, local: &str) -> Result<Option<bool>> {
    attr(attributes, namespace, local)
        .map(|value| {
            value
                .parse::<bool>()
                .map_err(|_| err(format!("invalid boolean form attribute '{local}'")))
        })
        .transpose()
}

fn current_scope(scopes: &[ScopeFrame]) -> Scope {
    scopes
        .last()
        .map(|scope| scope.1.clone())
        .unwrap_or(Scope::Document)
}

fn mark_xforms(result: &mut Forms, group: Option<&mut (usize, Group)>) {
    result.has_xforms = true;
    if let Some(group) = group {
        group.1.has_xforms = true;
    }
}

fn mark_events(result: &mut Forms, group: Option<&mut (usize, Group)>) {
    result.has_event_listeners = true;
    if let Some(group) = group {
        group.1.has_event_listeners = true;
    }
}

fn next_index(value: &mut usize) -> Result<usize> {
    let current = *value;
    *value = value
        .checked_add(1)
        .ok_or_else(|| err("form scope count overflow"))?;
    Ok(current)
}

fn inc_depth(value: usize) -> Result<usize> {
    let value = value
        .checked_add(1)
        .ok_or_else(|| err("form XML depth overflow"))?;
    if value > MAX_DEPTH {
        return Err(err("form XML nesting exceeds 128 levels"));
    }
    Ok(value)
}

fn node(limits: &mut Limits) -> Result<()> {
    limits.nodes = limits
        .nodes
        .checked_add(1)
        .ok_or_else(|| err("form node overflow"))?;
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
