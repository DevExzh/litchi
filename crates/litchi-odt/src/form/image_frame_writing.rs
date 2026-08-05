//! Typed writing and lossless mutation for `form:image-frame` controls.

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
const MAX_IMAGE_REFERENCE: usize = 8192;
const MAX_AGGREGATE: usize = 16 * 1024 * 1024;
const MAX_FORMS: usize = 4096;
const MAX_CONTROLS: usize = 65_536;

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ImageFrameControl {
    pub name: String,
    pub xml_id: String,
    pub metadata: GenericControlMetadata,
    pub data_field: Option<String>,
    pub disabled: Option<bool>,
    pub image_data: Option<String>,
    pub printable: Option<bool>,
    pub readonly: Option<bool>,
    pub title: Option<String>,
}

impl ImageFrameControl {
    pub fn new(name: impl Into<String>, xml_id: impl Into<String>) -> Self {
        Self {
            name: name.into(),
            xml_id: xml_id.into(),
            metadata: GenericControlMetadata::default(),
            data_field: None,
            disabled: None,
            image_data: None,
            printable: None,
            readonly: None,
            title: None,
        }
    }

    pub fn to_xml_fragment(&self) -> Result<String> {
        image_frame_xml(self)
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ImageFrameForm {
    pub name: String,
    pub controls: Vec<ImageFrameControl>,
    pub apply_filter: Option<bool>,
}

impl ImageFrameForm {
    pub fn new(name: impl Into<String>) -> Self {
        Self {
            name: name.into(),
            controls: Vec::new(),
            apply_filter: None,
        }
    }

    pub fn add_control(&mut self, control: ImageFrameControl) -> Result<()> {
        validate_control(&control)?;
        if self
            .controls
            .iter()
            .any(|existing| existing.name == control.name)
        {
            return invalid(format!(
                "duplicate image-frame control name '{}'",
                control.name
            ));
        }
        if self
            .controls
            .iter()
            .any(|existing| existing.xml_id == control.xml_id)
        {
            return invalid(format!(
                "duplicate image-frame control xml:id '{}'",
                control.xml_id
            ));
        }
        if self.controls.len() >= MAX_CONTROLS {
            return invalid("too many image-frame controls");
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

pub fn image_frame_controls(xml: &str) -> Result<Vec<ImageFrameControl>> {
    Ok(scan(xml)?
        .controls
        .into_iter()
        .map(|item| item.control)
        .collect())
}

pub fn insert_image_frame_control_xml(
    xml: &str,
    form_index: usize,
    control: &ImageFrameControl,
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

pub fn replace_image_frame_control_xml(
    xml: &str,
    index: usize,
    replacement: &ImageFrameControl,
) -> Result<String> {
    validate_control(replacement)?;
    let scan = scan(xml)?;
    let old = scan.controls.get(index).ok_or_else(|| {
        Error::InvalidFormat(format!("image-frame control {index} is out of bounds"))
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

pub fn remove_image_frame_control_xml(xml: &str, index: usize) -> Result<String> {
    let scan = scan(xml)?;
    let old = scan.controls.get(index).ok_or_else(|| {
        Error::InvalidFormat(format!("image-frame control {index} is out of bounds"))
    })?;
    apply(xml, old.span.clone(), "")
}

fn image_frame_xml(value: &ImageFrameControl) -> Result<String> {
    validate_control(value)?;
    let mut out = format!(
        r#"<form:image-frame form:name="{}" xml:id="{}""#,
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
    push_string(&mut out, "form:data-field", value.data_field.as_deref());
    push_bool(&mut out, "form:disabled", value.disabled);
    push_string(&mut out, "form:image-data", value.image_data.as_deref());
    push_bool(&mut out, "form:printable", value.printable);
    push_bool(&mut out, "form:readonly", value.readonly);
    push_string(&mut out, "form:title", value.title.as_deref());
    out.push_str("/>");
    Ok(out)
}

fn validate_control(value: &ImageFrameControl) -> Result<()> {
    validate_name("image-frame control name", &value.name)?;
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
    validate_optional(
        "image-frame data field",
        value.data_field.as_deref(),
        MAX_STRING,
    )?;
    validate_image_reference(value.image_data.as_deref())?;
    validate_optional("image-frame title", value.title.as_deref(), MAX_STRING)
}

fn validate_controls(controls: &[ImageFrameControl]) -> Result<()> {
    if controls.len() > MAX_CONTROLS {
        return invalid("too many image-frame controls");
    }
    let mut names = Vec::<&str>::new();
    let mut ids = Vec::<&str>::new();
    let mut aggregate = 0usize;
    for control in controls {
        validate_control(control)?;
        if names.contains(&control.name.as_str()) {
            return invalid(format!(
                "duplicate image-frame control name '{}'",
                control.name
            ));
        }
        if ids.contains(&control.xml_id.as_str()) {
            return invalid(format!(
                "duplicate image-frame control xml:id '{}'",
                control.xml_id
            ));
        }
        names.push(&control.name);
        ids.push(&control.xml_id);
        aggregate = aggregate.saturating_add(control_size(control));
        if aggregate > MAX_AGGREGATE {
            return invalid("image-frame control strings exceed 16 MiB");
        }
    }
    Ok(())
}

fn control_size(value: &ImageFrameControl) -> usize {
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
        .saturating_add(value.data_field.as_ref().map_or(0, String::len))
        .saturating_add(value.image_data.as_ref().map_or(0, String::len))
        .saturating_add(value.title.as_ref().map_or(0, String::len))
}

fn validate_image_reference(value: Option<&str>) -> Result<()> {
    let Some(value) = value else { return Ok(()) };
    validate_string("image-frame image reference", value, MAX_IMAGE_REFERENCE)?;
    if value.chars().any(char::is_whitespace) || value.contains('\\') {
        return invalid("image-frame image reference contains invalid URI characters");
    }
    let lower = value.to_ascii_lowercase();
    if ["javascript:", "macro:", "vnd.sun.star.script:", "data:"]
        .iter()
        .any(|scheme| lower.starts_with(scheme))
    {
        return invalid("image-frame image reference uses an active or inline-data URI scheme");
    }
    if !value.contains("://")
        && !value.starts_with('#')
        && value.split('/').any(|part| part == "." || part == "..")
    {
        return invalid("image-frame image reference contains package path traversal");
    }
    Ok(())
}

#[derive(Clone)]
struct ControlLocation {
    span: Range<usize>,
    form: usize,
    control: ImageFrameControl,
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
        return invalid("image-frame form XML exceeds 64 MiB");
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
                Error::InvalidFormat(format!("invalid image-frame form XML: {error}"))
            })?;
        let namespace = match namespace {
            ResolveResult::Bound(uri) => Some(String::from_utf8_lossy(uri.as_ref()).into_owned()),
            _ => None,
        };
        let end = reader.buffer_position() as usize;
        match event {
            Event::Start(ref element) => {
                if stack.len() >= 128 {
                    return invalid("image-frame form XML nesting exceeds 128 levels");
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
                } else if namespace.as_deref() == Some(FORM) && local == b"image-frame" {
                    if stack.iter().any(|open| open.control.is_some()) {
                        return invalid("image-frame controls cannot be nested");
                    }
                    let owner = *form_stack.last().ok_or_else(|| {
                        Error::InvalidFormat("image-frame control has no form owner".to_string())
                    })?;
                    let parsed = parse_control(&reader, element)?;
                    aggregate = aggregate.saturating_add(control_size(&parsed));
                    if aggregate > MAX_AGGREGATE {
                        return invalid("image-frame control strings exceed 16 MiB");
                    }
                    if controls.len() >= MAX_CONTROLS {
                        return invalid("too many image-frame controls");
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
                        "event and macro content is outside the image-frame mutation API",
                    );
                } else if stack.iter().any(|open| open.control.is_some())
                    && !is_property(namespace.as_deref(), local.as_slice())
                {
                    return invalid("unexpected child element in image-frame control");
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
                } else if namespace.as_deref() == Some(FORM) && local == b"image-frame" {
                    if stack.iter().any(|open| open.control.is_some()) {
                        return invalid("image-frame controls cannot be nested");
                    }
                    let owner = *form_stack.last().ok_or_else(|| {
                        Error::InvalidFormat("image-frame control has no form owner".to_string())
                    })?;
                    let parsed = parse_control(&reader, element)?;
                    aggregate = aggregate.saturating_add(control_size(&parsed));
                    if aggregate > MAX_AGGREGATE {
                        return invalid("image-frame control strings exceed 16 MiB");
                    }
                    if controls.len() >= MAX_CONTROLS {
                        return invalid("too many image-frame controls");
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
                        "event and macro content is outside the image-frame mutation API",
                    );
                } else if stack.iter().any(|open| open.control.is_some())
                    && !is_property(namespace.as_deref(), local.as_slice())
                {
                    return invalid("unexpected child element in image-frame control");
                }
            },
            Event::Text(text) if stack.iter().any(|open| open.control.is_some()) => {
                let decoded = text.decode().map_err(|error| {
                    Error::InvalidFormat(format!("invalid image-frame control text: {error}"))
                })?;
                if !decoded.trim().is_empty() {
                    return invalid("image-frame controls cannot contain character data");
                }
            },
            Event::CData(text) if stack.iter().any(|open| open.control.is_some()) => {
                let decoded = text.decode().map_err(|error| {
                    Error::InvalidFormat(format!("invalid image-frame control CDATA: {error}"))
                })?;
                if !decoded.trim().is_empty() {
                    return invalid("image-frame controls cannot contain CDATA");
                }
            },
            Event::GeneralRef(_) if stack.iter().any(|open| open.control.is_some()) => {
                return invalid("image-frame controls cannot contain entity references");
            },
            Event::End(ref element) => {
                let open = stack.pop().ok_or_else(|| {
                    Error::InvalidFormat("image-frame form XML stack underflow".to_string())
                })?;
                if open.local.as_slice() != element.local_name().as_ref() {
                    return invalid("mismatched image-frame form XML elements");
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
            Event::DocType(_) => return invalid("DOCTYPE is not allowed in image-frame form XML"),
            Event::Eof => break,
            _ => {},
        }
        previous = end;
        buffer.clear();
    }
    if !stack.is_empty() {
        return invalid("unclosed image-frame form XML elements");
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

fn parse_control(reader: &NsReader<&[u8]>, element: &BytesStart<'_>) -> Result<ImageFrameControl> {
    let attrs = attributes(reader, element)?;
    validate_allowed(&attrs, IMAGE_FRAME_ATTRS)?;
    let mut value = ImageFrameControl::new(
        required(&attrs, FORM, "name")?,
        required(&attrs, XML, "id")?,
    );
    value.metadata = GenericControlMetadata {
        form_id: optional(&attrs, FORM, "id"),
        control_implementation: optional(&attrs, FORM, "control-implementation"),
        xforms_bind: optional(&attrs, XFORMS, "bind"),
    };
    value.data_field = optional(&attrs, FORM, "data-field");
    value.disabled = optional_bool(&attrs, FORM, "disabled")?;
    value.image_data = optional(&attrs, FORM, "image-data");
    value.printable = optional_bool(&attrs, FORM, "printable")?;
    value.readonly = optional_bool(&attrs, FORM, "readonly")?;
    value.title = optional(&attrs, FORM, "title");
    validate_control(&value)?;
    Ok(value)
}

const IMAGE_FRAME_ATTRS: &[(&str, &str)] = &[
    (FORM, "name"),
    (XML, "id"),
    (FORM, "id"),
    (FORM, "control-implementation"),
    (FORM, "input-required"),
    (XFORMS, "bind"),
    (FORM, "data-field"),
    (FORM, "disabled"),
    (FORM, "image-data"),
    (FORM, "printable"),
    (FORM, "readonly"),
    (FORM, "title"),
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
            Error::InvalidFormat(format!("invalid image-frame control attribute: {error}"))
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
                    "invalid image-frame control attribute value: {error}"
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
                "unexpected image-frame control attribute '{}'",
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
    validate_name("image-frame control xml:id", value)?;
    let mut chars = value.chars();
    let first = chars.next().unwrap();
    if !(first == '_' || first.is_ascii_alphabetic())
        || !chars.all(|ch| ch == '_' || ch == '-' || ch == '.' || ch.is_ascii_alphanumeric())
    {
        return invalid(format!("invalid image-frame control xml:id '{value}'"));
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
    replacement: &ImageFrameControl,
    current: Option<&ImageFrameControl>,
) -> Result<()> {
    for item in &form.controls {
        if current.is_some_and(|value| value.xml_id == item.control.xml_id) {
            continue;
        }
        if item.control.name == replacement.name {
            return invalid(format!(
                "duplicate image-frame control name '{}'",
                replacement.name
            ));
        }
        if item.control.xml_id == replacement.xml_id {
            return invalid(format!(
                "duplicate image-frame control xml:id '{}'",
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
        .map_err(|_| Error::InvalidFormat("invalid image-frame form element name".to_string()))
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

    const ROOT: &str = r#"<o:document-content xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:f="urn:oasis:names:tc:opendocument:xmlns:form:1.0" xmlns:x="http://www.w3.org/2002/xforms"><o:body><o:text><o:forms>"#;
    const END: &str = "</o:forms></o:text></o:body></o:document-content>";

    #[test]
    fn canonical_control_round_trips_without_resolving_image() {
        let mut control = ImageFrameControl::new("Photo", "photo_1");
        control.image_data = Some("Pictures/missing.png".into());
        control.data_field = Some("Portrait".into());
        control.readonly = Some(true);
        control.title = Some("A & B".into());
        let mut form = ImageFrameForm::new("Main");
        form.add_control(control.clone()).unwrap();
        let parsed =
            image_frame_controls(&format!("{ROOT}{}{END}", form.to_xml_fragment().unwrap()))
                .unwrap();
        assert_eq!(parsed, [control]);
    }

    #[test]
    fn odfpy_odfdo_and_libreoffice_shapes_parse() {
        let producer = format!(
            r#"{ROOT}<f:form f:name="Producer"><f:image-frame f:name="odfdo" xml:id="image" f:control-implementation="com.sun.star.form.component.DatabaseImageControl" f:data-field="Photo" f:disabled="false" f:image-data="Pictures/photo.png" f:printable="true" f:readonly="false" f:title="Portrait"/></f:form>{END}"#
        );
        let parsed = image_frame_controls(&producer).unwrap();
        assert_eq!(parsed.len(), 1);
        assert_eq!(parsed[0].image_data.as_deref(), Some("Pictures/photo.png"));
    }

    #[test]
    fn lossless_mutation_and_empty_form_expansion() {
        let xml = format!(
            r#"{ROOT}<f:form f:name="Main"><f:image-frame f:name="Old" xml:id="old"><f:properties><f:property f:property-name="Keep" o:value-type="void"/></f:properties></f:image-frame><!--keep--><f:text f:name="Text" xml:id="text"/></f:form>{END}"#
        );
        let inserted_control = ImageFrameControl::new("Inserted", "inserted");
        let inserted = insert_image_frame_control_xml(&xml, 0, &inserted_control).unwrap();
        assert!(inserted.contains("<!--keep-->") && inserted.contains("f:text"));
        let replacement = ImageFrameControl::new("Replacement", "replacement");
        let replaced = replace_image_frame_control_xml(&inserted, 0, &replacement).unwrap();
        let removed = remove_image_frame_control_xml(&replaced, 1).unwrap();
        assert_eq!(image_frame_controls(&removed).unwrap(), [replacement]);
        let empty = format!(r#"{ROOT}<f:form f:name="Empty"/>{END}"#);
        assert!(
            insert_image_frame_control_xml(&empty, 0, &inserted_control)
                .unwrap()
                .contains("</f:form>")
        );
    }

    #[test]
    fn hostile_namespaces_resources_children_and_active_content_are_rejected() {
        assert!(
            ImageFrameControl::new("I", "1bad")
                .to_xml_fragment()
                .is_err()
        );
        let wrong_attr = format!(
            r#"{ROOT}<f:form f:name="Main"><f:image-frame f:name="I" xml:id="i" o:readonly="true"/></f:form>{END}"#
        );
        assert!(image_frame_controls(&wrong_attr).is_err());
        let foreign = format!(
            r#"{ROOT}<f:form f:name="Main"><f:image-frame f:name="I" xml:id="i"><o:p/></f:image-frame></f:form>{END}"#
        );
        assert!(image_frame_controls(&foreign).is_err());
        let event = format!(r#"{ROOT}<f:form f:name="Main"><o:event-listeners/></f:form>{END}"#);
        assert!(image_frame_controls(&event).is_err());
        for reference in [
            "../Pictures/a.png",
            "Pictures\\a.png",
            "javascript:alert(1)",
            "data:image/png;base64,AA==",
        ] {
            let mut control = ImageFrameControl::new("I", "i");
            control.image_data = Some(reference.into());
            assert!(control.to_xml_fragment().is_err());
        }
        let mut oversized = ImageFrameControl::new("I", "i");
        oversized.image_data = Some("x".repeat(MAX_IMAGE_REFERENCE + 1));
        assert!(oversized.to_xml_fragment().is_err());
    }

    #[test]
    fn builder_and_mutable_document_round_trip_without_image_io() {
        use crate::{Builder, Document, mutable::MutableDocument};
        let mut initial = ImageFrameControl::new("Image", "image");
        initial.image_data = Some("Pictures/intentionally-missing.png".into());
        let mut form = ImageFrameForm::new("Main");
        form.add_control(initial.clone()).unwrap();
        let mut builder = Builder::new();
        builder.add_image_frame_form(&form).unwrap();
        builder.add_paragraph("body").unwrap();
        let document = Document::from_bytes(builder.build().unwrap()).unwrap();
        let mut mutable = MutableDocument::from_document(document).unwrap();
        assert_eq!(mutable.image_frame_controls().unwrap(), [initial]);
        let inserted = ImageFrameControl::new("Inserted", "inserted");
        mutable.insert_image_frame_control(0, &inserted).unwrap();
        let replacement = ImageFrameControl::new("Replacement", "replacement");
        assert_eq!(
            mutable
                .replace_image_frame_control(0, &replacement)
                .unwrap()
                .name,
            "Image"
        );
        assert_eq!(
            mutable.remove_image_frame_control(1).unwrap().name,
            "Inserted"
        );
        assert_eq!(mutable.image_frame_controls().unwrap(), [replacement]);
    }
}
