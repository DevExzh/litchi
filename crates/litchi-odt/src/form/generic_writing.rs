//! Typed writing and lossless mutation for fixed-text, hidden, and generic controls.

use std::ops::Range;

use litchi_core::{Error, Result};
use quick_xml::NsReader;
use quick_xml::events::{BytesStart, Event};
use quick_xml::name::ResolveResult;

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

#[derive(Debug, Clone, PartialEq, Eq, Default)]
pub struct GenericControlMetadata {
    pub form_id: Option<String>,
    pub control_implementation: Option<String>,
    pub xforms_bind: Option<String>,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct FixedTextControl {
    pub name: String,
    pub xml_id: String,
    pub metadata: GenericControlMetadata,
    pub form_for: Option<String>,
    pub disabled: Option<bool>,
    pub label: Option<String>,
    pub printable: Option<bool>,
    pub title: Option<String>,
    pub multi_line: Option<bool>,
}

impl FixedTextControl {
    pub fn new(name: impl Into<String>, xml_id: impl Into<String>) -> Self {
        Self {
            name: name.into(),
            xml_id: xml_id.into(),
            metadata: GenericControlMetadata::default(),
            form_for: None,
            disabled: None,
            label: None,
            printable: None,
            title: None,
            multi_line: None,
        }
    }

    pub fn to_xml_fragment(&self) -> Result<String> {
        fixed_text_xml(self)
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct HiddenControl {
    pub name: String,
    pub xml_id: String,
    pub metadata: GenericControlMetadata,
    pub value: Option<String>,
}

impl HiddenControl {
    pub fn new(name: impl Into<String>, xml_id: impl Into<String>) -> Self {
        Self {
            name: name.into(),
            xml_id: xml_id.into(),
            metadata: GenericControlMetadata::default(),
            value: None,
        }
    }

    pub fn to_xml_fragment(&self) -> Result<String> {
        hidden_xml(self)
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct GenericControl {
    pub name: String,
    pub xml_id: String,
    pub metadata: GenericControlMetadata,
}

impl GenericControl {
    pub fn new(name: impl Into<String>, xml_id: impl Into<String>) -> Self {
        Self {
            name: name.into(),
            xml_id: xml_id.into(),
            metadata: GenericControlMetadata::default(),
        }
    }

    pub fn to_xml_fragment(&self) -> Result<String> {
        generic_xml(self)
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub enum GenericFormControl {
    FixedText(FixedTextControl),
    Hidden(HiddenControl),
    Generic(GenericControl),
}

impl From<FixedTextControl> for GenericFormControl {
    fn from(value: FixedTextControl) -> Self {
        Self::FixedText(value)
    }
}

impl From<HiddenControl> for GenericFormControl {
    fn from(value: HiddenControl) -> Self {
        Self::Hidden(value)
    }
}

impl From<GenericControl> for GenericFormControl {
    fn from(value: GenericControl) -> Self {
        Self::Generic(value)
    }
}

impl GenericFormControl {
    pub fn name(&self) -> &str {
        match self {
            Self::FixedText(value) => &value.name,
            Self::Hidden(value) => &value.name,
            Self::Generic(value) => &value.name,
        }
    }

    pub fn xml_id(&self) -> &str {
        match self {
            Self::FixedText(value) => &value.xml_id,
            Self::Hidden(value) => &value.xml_id,
            Self::Generic(value) => &value.xml_id,
        }
    }

    pub fn to_xml_fragment(&self) -> Result<String> {
        match self {
            Self::FixedText(value) => fixed_text_xml(value),
            Self::Hidden(value) => hidden_xml(value),
            Self::Generic(value) => generic_xml(value),
        }
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct GenericForm {
    pub name: String,
    pub controls: Vec<GenericFormControl>,
    pub apply_filter: Option<bool>,
}

impl GenericForm {
    pub fn new(name: impl Into<String>) -> Self {
        Self {
            name: name.into(),
            controls: Vec::new(),
            apply_filter: None,
        }
    }

    pub fn add_control(&mut self, control: impl Into<GenericFormControl>) -> Result<()> {
        let control = control.into();
        validate_control(&control)?;
        if self
            .controls
            .iter()
            .any(|existing| existing.name() == control.name())
        {
            return invalid(format!(
                "duplicate generic form control name '{}'",
                control.name()
            ));
        }
        if self
            .controls
            .iter()
            .any(|existing| existing.xml_id() == control.xml_id())
        {
            return invalid(format!(
                "duplicate generic form control xml:id '{}'",
                control.xml_id()
            ));
        }
        if let Some(target) = fixed_text_target(&control)
            && self
                .controls
                .iter()
                .any(|existing| fixed_text_target(existing) == Some(target))
        {
            return invalid(format!(
                "multiple fixed-text controls label target '{target}'"
            ));
        }
        if self.controls.len() >= MAX_CONTROLS {
            return invalid("too many generic form controls");
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

pub fn generic_form_controls(xml: &str) -> Result<Vec<GenericFormControl>> {
    Ok(scan(xml)?
        .controls
        .into_iter()
        .map(|item| item.control)
        .collect())
}

pub fn insert_generic_form_control_xml(
    xml: &str,
    form_index: usize,
    control: &GenericFormControl,
) -> Result<String> {
    validate_control(control)?;
    let scan = scan(xml)?;
    let form = scan
        .forms
        .get(form_index)
        .ok_or_else(|| Error::InvalidFormat(format!("form {form_index} is out of bounds")))?;
    reject_duplicate(form, control, None)?;
    if scan.ids.iter().any(|id| id == control.xml_id()) {
        return invalid(format!("duplicate xml:id '{}'", control.xml_id()));
    }
    let fragment = bind_fragment(control.to_xml_fragment()?);
    match form.site.clone() {
        Site::Paired { close_start } => apply(xml, close_start..close_start, &fragment),
        Site::Empty { start, end, qname } => expand_empty(xml, start, end, &qname, &fragment),
    }
}

pub fn replace_generic_form_control_xml(
    xml: &str,
    index: usize,
    replacement: &GenericFormControl,
) -> Result<String> {
    validate_control(replacement)?;
    let scan = scan(xml)?;
    let old = scan.controls.get(index).ok_or_else(|| {
        Error::InvalidFormat(format!("generic form control {index} is out of bounds"))
    })?;
    reject_duplicate(&scan.forms[old.form], replacement, Some(&old.control))?;
    if replacement.xml_id() != old.control.xml_id()
        && scan.ids.iter().any(|id| id == replacement.xml_id())
    {
        return invalid(format!("duplicate xml:id '{}'", replacement.xml_id()));
    }
    apply(
        xml,
        old.span.clone(),
        &bind_fragment(replacement.to_xml_fragment()?),
    )
}

pub fn remove_generic_form_control_xml(xml: &str, index: usize) -> Result<String> {
    let scan = scan(xml)?;
    let old = scan.controls.get(index).ok_or_else(|| {
        Error::InvalidFormat(format!("generic form control {index} is out of bounds"))
    })?;
    apply(xml, old.span.clone(), "")
}

fn fixed_text_xml(value: &FixedTextControl) -> Result<String> {
    validate_fixed_text(value)?;
    let mut out = control_start("fixed-text", &value.name, &value.xml_id, &value.metadata);
    push_string(&mut out, "form:for", value.form_for.as_deref());
    push_bool(&mut out, "form:disabled", value.disabled);
    push_string(&mut out, "form:label", value.label.as_deref());
    push_bool(&mut out, "form:printable", value.printable);
    push_string(&mut out, "form:title", value.title.as_deref());
    push_bool(&mut out, "form:multi-line", value.multi_line);
    out.push_str("/>");
    Ok(out)
}

fn hidden_xml(value: &HiddenControl) -> Result<String> {
    validate_hidden(value)?;
    let mut out = control_start("hidden", &value.name, &value.xml_id, &value.metadata);
    push_string(&mut out, "form:value", value.value.as_deref());
    out.push_str("/>");
    Ok(out)
}

fn generic_xml(value: &GenericControl) -> Result<String> {
    validate_generic(value)?;
    let mut out = control_start(
        "generic-control",
        &value.name,
        &value.xml_id,
        &value.metadata,
    );
    out.push_str("/>");
    Ok(out)
}

fn control_start(
    kind: &str,
    name: &str,
    xml_id: &str,
    metadata: &GenericControlMetadata,
) -> String {
    let mut out = format!(
        r#"<form:{kind} form:name="{}" xml:id="{}""#,
        escape(name),
        escape(xml_id)
    );
    push_string(&mut out, "form:id", metadata.form_id.as_deref());
    push_string(
        &mut out,
        "form:control-implementation",
        metadata.control_implementation.as_deref(),
    );
    push_string(&mut out, "xforms:bind", metadata.xforms_bind.as_deref());
    out
}

fn validate_control(value: &GenericFormControl) -> Result<()> {
    match value {
        GenericFormControl::FixedText(value) => validate_fixed_text(value),
        GenericFormControl::Hidden(value) => validate_hidden(value),
        GenericFormControl::Generic(value) => validate_generic(value),
    }
}

fn validate_fixed_text(value: &FixedTextControl) -> Result<()> {
    validate_identity(&value.name, &value.xml_id, &value.metadata)?;
    validate_optional("fixed-text for", value.form_for.as_deref(), MAX_REFERENCE)?;
    validate_optional("fixed-text label", value.label.as_deref(), MAX_STRING)?;
    validate_optional("fixed-text title", value.title.as_deref(), MAX_STRING)
}

fn validate_hidden(value: &HiddenControl) -> Result<()> {
    validate_identity(&value.name, &value.xml_id, &value.metadata)?;
    validate_optional("hidden value", value.value.as_deref(), MAX_STRING)
}

fn validate_generic(value: &GenericControl) -> Result<()> {
    validate_identity(&value.name, &value.xml_id, &value.metadata)
}

fn validate_identity(name: &str, xml_id: &str, metadata: &GenericControlMetadata) -> Result<()> {
    validate_name("generic form control name", name)?;
    validate_xml_id(xml_id)?;
    if let Some(form_id) = metadata.form_id.as_deref() {
        validate_xml_id(form_id)?;
    }
    validate_optional(
        "control implementation",
        metadata.control_implementation.as_deref(),
        MAX_REFERENCE,
    )?;
    validate_optional(
        "XForms bind",
        metadata.xforms_bind.as_deref(),
        MAX_REFERENCE,
    )
}

fn validate_controls(controls: &[GenericFormControl]) -> Result<()> {
    if controls.len() > MAX_CONTROLS {
        return invalid("too many generic form controls");
    }
    let mut names = Vec::<&str>::new();
    let mut ids = Vec::<&str>::new();
    let mut fixed_targets = Vec::<&str>::new();
    let mut aggregate = 0usize;
    for control in controls {
        validate_control(control)?;
        if names.contains(&control.name()) {
            return invalid(format!(
                "duplicate generic form control name '{}'",
                control.name()
            ));
        }
        if ids.contains(&control.xml_id()) {
            return invalid(format!(
                "duplicate generic form control xml:id '{}'",
                control.xml_id()
            ));
        }
        names.push(control.name());
        ids.push(control.xml_id());
        if let Some(target) = fixed_text_target(control) {
            if fixed_targets.contains(&target) {
                return invalid(format!(
                    "multiple fixed-text controls label target '{target}'"
                ));
            }
            fixed_targets.push(target);
        }
        aggregate = aggregate.saturating_add(control_size(control));
        if aggregate > MAX_AGGREGATE {
            return invalid("generic form control strings exceed 16 MiB");
        }
    }
    Ok(())
}

fn fixed_text_target(control: &GenericFormControl) -> Option<&str> {
    match control {
        GenericFormControl::FixedText(value) => {
            value.form_for.as_deref().filter(|value| !value.is_empty())
        },
        _ => None,
    }
}

fn control_size(control: &GenericFormControl) -> usize {
    let (name, id, metadata, extras): (
        &String,
        &String,
        &GenericControlMetadata,
        &[Option<&String>],
    ) = match control {
        GenericFormControl::FixedText(value) => (
            &value.name,
            &value.xml_id,
            &value.metadata,
            &[
                value.form_for.as_ref(),
                value.label.as_ref(),
                value.title.as_ref(),
            ],
        ),
        GenericFormControl::Hidden(value) => (
            &value.name,
            &value.xml_id,
            &value.metadata,
            &[value.value.as_ref()],
        ),
        GenericFormControl::Generic(value) => (&value.name, &value.xml_id, &value.metadata, &[]),
    };
    name.len()
        .saturating_add(id.len())
        .saturating_add(metadata.form_id.as_ref().map_or(0, String::len))
        .saturating_add(
            metadata
                .control_implementation
                .as_ref()
                .map_or(0, String::len),
        )
        .saturating_add(metadata.xforms_bind.as_ref().map_or(0, String::len))
        .saturating_add(
            extras
                .iter()
                .flatten()
                .fold(0usize, |sum, value| sum.saturating_add(value.len())),
        )
}

#[derive(Clone)]
struct ControlLocation {
    span: Range<usize>,
    form: usize,
    control: GenericFormControl,
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
        return invalid("generic form XML exceeds 64 MiB");
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
            .map_err(|error| Error::InvalidFormat(format!("invalid generic form XML: {error}")))?;
        let namespace = match namespace {
            ResolveResult::Bound(uri) => Some(String::from_utf8_lossy(uri.as_ref()).into_owned()),
            _ => None,
        };
        let end = reader.buffer_position() as usize;
        match event {
            Event::Start(ref element) => {
                if stack.len() >= 128 {
                    return invalid("generic form XML nesting exceeds 128 levels");
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
                } else if is_target(namespace.as_deref(), local.as_slice()) {
                    if stack.iter().any(|open| open.control.is_some()) {
                        return invalid("generic form controls cannot be nested");
                    }
                    let owner = *form_stack.last().ok_or_else(|| {
                        Error::InvalidFormat("generic form control has no form owner".to_string())
                    })?;
                    let parsed = parse_control(&reader, element, local.as_slice())?;
                    aggregate = aggregate.saturating_add(control_size(&parsed));
                    if aggregate > MAX_AGGREGATE {
                        return invalid("generic form control strings exceed 16 MiB");
                    }
                    if controls.len() >= MAX_CONTROLS {
                        return invalid("too many generic form controls");
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
                        "event and macro content is outside the generic-control mutation API",
                    );
                } else if stack.iter().any(|open| open.control.is_some())
                    && !is_property(namespace.as_deref(), local.as_slice())
                {
                    return invalid("unexpected child element in generic form control");
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
                } else if is_target(namespace.as_deref(), local.as_slice()) {
                    if stack.iter().any(|open| open.control.is_some()) {
                        return invalid("generic form controls cannot be nested");
                    }
                    let owner = *form_stack.last().ok_or_else(|| {
                        Error::InvalidFormat("generic form control has no form owner".to_string())
                    })?;
                    let parsed = parse_control(&reader, element, local.as_slice())?;
                    aggregate = aggregate.saturating_add(control_size(&parsed));
                    if aggregate > MAX_AGGREGATE {
                        return invalid("generic form control strings exceed 16 MiB");
                    }
                    if controls.len() >= MAX_CONTROLS {
                        return invalid("too many generic form controls");
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
                        "event and macro content is outside the generic-control mutation API",
                    );
                } else if stack.iter().any(|open| open.control.is_some())
                    && !is_property(namespace.as_deref(), local.as_slice())
                {
                    return invalid("unexpected child element in generic form control");
                }
            },
            Event::Text(text) if stack.iter().any(|open| open.control.is_some()) => {
                let decoded = text.decode().map_err(|error| {
                    Error::InvalidFormat(format!("invalid generic control text: {error}"))
                })?;
                if !decoded.trim().is_empty() {
                    return invalid("generic form controls cannot contain character data");
                }
            },
            Event::CData(text) if stack.iter().any(|open| open.control.is_some()) => {
                let decoded = text.decode().map_err(|error| {
                    Error::InvalidFormat(format!("invalid generic control CDATA: {error}"))
                })?;
                if !decoded.trim().is_empty() {
                    return invalid("generic form controls cannot contain CDATA");
                }
            },
            Event::GeneralRef(_) if stack.iter().any(|open| open.control.is_some()) => {
                return invalid("generic form controls cannot contain entity references");
            },
            Event::End(ref element) => {
                let open = stack.pop().ok_or_else(|| {
                    Error::InvalidFormat("generic form XML stack underflow".to_string())
                })?;
                if open.local.as_slice() != element.local_name().as_ref() {
                    return invalid("mismatched generic form XML elements");
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
            Event::DocType(_) => return invalid("DOCTYPE is not allowed in generic form XML"),
            Event::Eof => break,
            _ => {},
        }
        previous = end;
        buffer.clear();
    }
    if !stack.is_empty() {
        return invalid("unclosed generic form XML elements");
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

fn parse_control(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
    local: &[u8],
) -> Result<GenericFormControl> {
    let attrs = attributes(reader, element)?;
    validate_allowed(
        &attrs,
        match local {
            b"fixed-text" => FIXED_ATTRS,
            b"hidden" => HIDDEN_ATTRS,
            _ => GENERIC_ATTRS,
        },
    )?;
    let name = required(&attrs, FORM, "name")?;
    let xml_id = required(&attrs, XML, "id")?;
    let metadata = GenericControlMetadata {
        form_id: optional(&attrs, FORM, "id"),
        control_implementation: optional(&attrs, FORM, "control-implementation"),
        xforms_bind: optional(&attrs, XFORMS, "bind"),
    };
    match local {
        b"fixed-text" => {
            let mut value = FixedTextControl::new(name, xml_id);
            value.metadata = metadata;
            value.form_for = optional(&attrs, FORM, "for");
            value.disabled = optional_bool(&attrs, FORM, "disabled")?;
            value.label = optional(&attrs, FORM, "label");
            value.printable = optional_bool(&attrs, FORM, "printable")?;
            value.title = optional(&attrs, FORM, "title");
            value.multi_line = optional_bool(&attrs, FORM, "multi-line")?;
            validate_fixed_text(&value)?;
            Ok(value.into())
        },
        b"hidden" => {
            let mut value = HiddenControl::new(name, xml_id);
            value.metadata = metadata;
            value.value = optional(&attrs, FORM, "value");
            validate_hidden(&value)?;
            Ok(value.into())
        },
        _ => {
            let mut value = GenericControl::new(name, xml_id);
            value.metadata = metadata;
            validate_generic(&value)?;
            Ok(value.into())
        },
    }
}

const GENERIC_ATTRS: &[(&str, &str)] = &[
    (FORM, "name"),
    (XML, "id"),
    (FORM, "id"),
    (FORM, "control-implementation"),
    (FORM, "input-required"),
    (XFORMS, "bind"),
];
const HIDDEN_ATTRS: &[(&str, &str)] = &[
    (FORM, "name"),
    (XML, "id"),
    (FORM, "id"),
    (FORM, "control-implementation"),
    (FORM, "input-required"),
    (XFORMS, "bind"),
    (FORM, "value"),
];
const FIXED_ATTRS: &[(&str, &str)] = &[
    (FORM, "name"),
    (XML, "id"),
    (FORM, "id"),
    (FORM, "control-implementation"),
    (FORM, "input-required"),
    (XFORMS, "bind"),
    (FORM, "for"),
    (FORM, "disabled"),
    (FORM, "label"),
    (FORM, "printable"),
    (FORM, "title"),
    (FORM, "multi-line"),
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

fn is_target(namespace: Option<&str>, local: &[u8]) -> bool {
    namespace == Some(FORM) && matches!(local, b"fixed-text" | b"hidden" | b"generic-control")
}

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
            Error::InvalidFormat(format!("invalid generic control attribute: {error}"))
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
                Error::InvalidFormat(format!("invalid generic control attribute value: {error}"))
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
                "unexpected generic form control attribute '{}'",
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
    validate_name("generic form control xml:id", value)?;
    let mut chars = value.chars();
    let first = chars.next().unwrap();
    if !(first == '_' || first.is_ascii_alphabetic())
        || !chars.all(|ch| ch == '_' || ch == '-' || ch == '.' || ch.is_ascii_alphanumeric())
    {
        return invalid(format!("invalid generic form control xml:id '{value}'"));
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
    if value
        .chars()
        .any(|ch| matches!(ch as u32, 0..=8 | 11 | 12 | 14..=31))
    {
        return invalid(format!("{label} contains invalid XML characters"));
    }
    Ok(())
}

fn reject_duplicate(
    form: &FormLocation,
    replacement: &GenericFormControl,
    current: Option<&GenericFormControl>,
) -> Result<()> {
    for item in &form.controls {
        if current.is_some_and(|value| value.xml_id() == item.control.xml_id()) {
            continue;
        }
        if item.control.name() == replacement.name() {
            return invalid(format!(
                "duplicate generic form control name '{}'",
                replacement.name()
            ));
        }
        if item.control.xml_id() == replacement.xml_id() {
            return invalid(format!(
                "duplicate generic form control xml:id '{}'",
                replacement.xml_id()
            ));
        }
        if let Some(target) = fixed_text_target(replacement)
            && fixed_text_target(&item.control) == Some(target)
        {
            return invalid(format!(
                "multiple fixed-text controls label target '{target}'"
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
        .map_err(|_| Error::InvalidFormat("invalid generic form element name".to_string()))
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
    fn canonical_controls_round_trip() {
        let mut fixed = FixedTextControl::new("Label", "fixed_1");
        fixed.label = Some("Name & address".into());
        fixed.form_for = Some("text_1".into());
        fixed.multi_line = Some(true);
        fixed.metadata.control_implementation = Some("com.example.FixedText".into());
        let mut hidden = HiddenControl::new("Token", "hidden_1");
        hidden.value = Some("A < B".into());
        let mut generic = GenericControl::new("Custom", "generic_1");
        generic.metadata.xforms_bind = Some("bind_1".into());
        let mut form = GenericForm::new("Main");
        form.add_control(fixed).unwrap();
        form.add_control(hidden).unwrap();
        form.add_control(generic).unwrap();
        let parsed =
            generic_form_controls(&format!("{ROOT}{}{END}", form.to_xml_fragment().unwrap()))
                .unwrap();
        assert_eq!(parsed.len(), 3);
        assert!(
            matches!(&parsed[0], GenericFormControl::FixedText(value) if value.label.as_deref() == Some("Name & address"))
        );
        assert!(
            matches!(&parsed[1], GenericFormControl::Hidden(value) if value.value.as_deref() == Some("A < B"))
        );
    }

    #[test]
    fn odfpy_odfdo_and_libreoffice_shapes_parse() {
        let producer = format!(
            r#"{ROOT}<f:form f:name="Producer"><f:fixed-text f:name="odfdo" xml:id="fixed" f:control-implementation="com.sun.star.form.component.FixedText" f:label="Label" f:multi-line="false"/><f:hidden f:name="lo-hidden" xml:id="hidden" f:control-implementation="com.sun.star.form.component.HiddenControl" f:value="secret"/><f:generic-control f:name="odfpy" xml:id="generic"/></f:form>{END}"#
        );
        let parsed = generic_form_controls(&producer).unwrap();
        assert_eq!(parsed.len(), 3);
        assert!(matches!(&parsed[2], GenericFormControl::Generic(_)));
    }

    #[test]
    fn lossless_mutation_and_empty_form_expansion() {
        let xml = format!(
            r#"{ROOT}<f:form f:name="Main"><f:hidden f:name="Old" xml:id="old"><f:properties><f:property f:property-name="Keep" o:value-type="void"/></f:properties></f:hidden><!--keep--><f:text f:name="Text" xml:id="text"/></f:form>{END}"#
        );
        let generic: GenericFormControl = GenericControl::new("Generic", "generic").into();
        let inserted = insert_generic_form_control_xml(&xml, 0, &generic).unwrap();
        assert!(inserted.contains("<!--keep-->") && inserted.contains("f:text"));
        let fixed: GenericFormControl = FixedTextControl::new("Fixed", "fixed").into();
        let replaced = replace_generic_form_control_xml(&inserted, 0, &fixed).unwrap();
        let removed = remove_generic_form_control_xml(&replaced, 1).unwrap();
        assert_eq!(generic_form_controls(&removed).unwrap(), [fixed]);
        let empty = format!(r#"{ROOT}<f:form f:name="Empty"/>{END}"#);
        assert!(
            insert_generic_form_control_xml(&empty, 0, &generic)
                .unwrap()
                .contains("</f:form>")
        );
    }

    #[test]
    fn hostile_values_children_duplicates_and_active_content_are_rejected() {
        assert!(HiddenControl::new("H", "1bad").to_xml_fragment().is_err());
        let bad_bool = format!(
            r#"{ROOT}<f:form f:name="Main"><f:fixed-text f:name="F" xml:id="f" f:multi-line="yes"/></f:form>{END}"#
        );
        assert!(generic_form_controls(&bad_bool).is_err());
        let foreign = format!(
            r#"{ROOT}<f:form f:name="Main"><f:hidden f:name="H" xml:id="h"><o:p/></f:hidden></f:form>{END}"#
        );
        assert!(generic_form_controls(&foreign).is_err());
        let event = format!(r#"{ROOT}<f:form f:name="Main"><o:event-listeners/></f:form>{END}"#);
        assert!(generic_form_controls(&event).is_err());
        let duplicate = format!(
            r#"{ROOT}<o:p xml:id="same"/><f:form f:name="Main"><f:hidden f:name="H" xml:id="same"/></f:form>{END}"#
        );
        assert!(generic_form_controls(&duplicate).is_err());
        let mut hidden = HiddenControl::new("H", "h");
        hidden.value = Some("x".repeat(MAX_STRING + 1));
        assert!(hidden.to_xml_fragment().is_err());
    }

    #[test]
    fn builder_and_mutable_document_round_trip() {
        use crate::{Builder, Document, mutable::MutableDocument};
        let mut form = GenericForm::new("Main");
        form.add_control(HiddenControl::new("Hidden", "hidden"))
            .unwrap();
        let mut builder = Builder::new();
        builder.add_generic_form(&form).unwrap();
        builder.add_paragraph("body").unwrap();
        let document = Document::from_bytes(builder.build().unwrap()).unwrap();
        let mut mutable = MutableDocument::from_document(document).unwrap();
        assert_eq!(mutable.generic_form_controls().unwrap().len(), 1);
        let generic: GenericFormControl = GenericControl::new("Generic", "generic").into();
        mutable.insert_generic_form_control(0, &generic).unwrap();
        let fixed: GenericFormControl = FixedTextControl::new("Fixed", "fixed").into();
        assert!(matches!(
            mutable.replace_generic_form_control(0, &fixed).unwrap(),
            GenericFormControl::Hidden(_)
        ));
        assert!(matches!(
            mutable.remove_generic_form_control(1).unwrap(),
            GenericFormControl::Generic(_)
        ));
        assert_eq!(mutable.generic_form_controls().unwrap(), [fixed]);
    }

    #[test]
    fn fixed_text_target_cardinality_is_atomic() {
        let mut first = FixedTextControl::new("First", "first");
        first.form_for = Some("target".into());
        let mut second = FixedTextControl::new("Second", "second");
        second.form_for = Some("target".into());
        let mut form = GenericForm::new("Main");
        form.add_control(first).unwrap();
        assert!(form.add_control(second).is_err());
        assert_eq!(form.controls.len(), 1);
        let malformed = format!(
            r#"{ROOT}<f:form f:name="Main"><f:fixed-text f:name="A" xml:id="a" f:for="target"/><f:fixed-text f:name="B" xml:id="b" f:for="target"/></f:form>{END}"#
        );
        assert!(generic_form_controls(&malformed).is_err());
    }
}
