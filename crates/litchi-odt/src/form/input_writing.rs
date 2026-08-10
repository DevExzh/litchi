//! Typed writing and lossless mutation for password and file controls.

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
pub struct PasswordControl {
    pub name: String,
    pub xml_id: String,
    pub metadata: GenericControlMetadata,
    pub disabled: Option<bool>,
    pub max_length: Option<u64>,
    pub printable: Option<bool>,
    pub tab_index: Option<u64>,
    pub tab_stop: Option<bool>,
    pub title: Option<String>,
    pub value: Option<String>,
    pub convert_empty_to_null: Option<bool>,
    pub linked_cell: Option<String>,
    pub echo_char: Option<char>,
}

impl PasswordControl {
    pub fn new(name: impl Into<String>, xml_id: impl Into<String>) -> Self {
        Self {
            name: name.into(),
            xml_id: xml_id.into(),
            metadata: GenericControlMetadata::default(),
            disabled: None,
            max_length: None,
            printable: None,
            tab_index: None,
            tab_stop: None,
            title: None,
            value: None,
            convert_empty_to_null: None,
            linked_cell: None,
            echo_char: None,
        }
    }

    pub fn to_xml_fragment(&self) -> Result<String> {
        password_xml(self)
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct FileControl {
    pub name: String,
    pub xml_id: String,
    pub metadata: GenericControlMetadata,
    pub current_value: Option<String>,
    pub disabled: Option<bool>,
    pub max_length: Option<u64>,
    pub printable: Option<bool>,
    pub readonly: Option<bool>,
    pub tab_index: Option<u64>,
    pub tab_stop: Option<bool>,
    pub title: Option<String>,
    pub value: Option<String>,
    pub linked_cell: Option<String>,
}

impl FileControl {
    pub fn new(name: impl Into<String>, xml_id: impl Into<String>) -> Self {
        Self {
            name: name.into(),
            xml_id: xml_id.into(),
            metadata: GenericControlMetadata::default(),
            current_value: None,
            disabled: None,
            max_length: None,
            printable: None,
            readonly: None,
            tab_index: None,
            tab_stop: None,
            title: None,
            value: None,
            linked_cell: None,
        }
    }

    pub fn to_xml_fragment(&self) -> Result<String> {
        file_xml(self)
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub enum PasswordFileControl {
    Password(PasswordControl),
    File(FileControl),
}

impl From<PasswordControl> for PasswordFileControl {
    fn from(value: PasswordControl) -> Self {
        Self::Password(value)
    }
}

impl From<FileControl> for PasswordFileControl {
    fn from(value: FileControl) -> Self {
        Self::File(value)
    }
}

impl PasswordFileControl {
    pub fn name(&self) -> &str {
        match self {
            Self::Password(value) => &value.name,
            Self::File(value) => &value.name,
        }
    }

    pub fn xml_id(&self) -> &str {
        match self {
            Self::Password(value) => &value.xml_id,
            Self::File(value) => &value.xml_id,
        }
    }

    pub fn to_xml_fragment(&self) -> Result<String> {
        match self {
            Self::Password(value) => password_xml(value),
            Self::File(value) => file_xml(value),
        }
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct PasswordFileForm {
    pub name: String,
    pub controls: Vec<PasswordFileControl>,
    pub apply_filter: Option<bool>,
}

impl PasswordFileForm {
    pub fn new(name: impl Into<String>) -> Self {
        Self {
            name: name.into(),
            controls: Vec::new(),
            apply_filter: None,
        }
    }

    pub fn add_control(&mut self, control: impl Into<PasswordFileControl>) -> Result<()> {
        let control = control.into();
        validate_control(&control)?;
        if self
            .controls
            .iter()
            .any(|existing| existing.name() == control.name())
        {
            return invalid(format!(
                "duplicate password/file control name '{}'",
                control.name()
            ));
        }
        if self
            .controls
            .iter()
            .any(|existing| existing.xml_id() == control.xml_id())
        {
            return invalid(format!(
                "duplicate password/file control xml:id '{}'",
                control.xml_id()
            ));
        }
        if self.controls.len() >= MAX_CONTROLS {
            return invalid("too many password/file controls");
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

pub fn password_file_controls(xml: &str) -> Result<Vec<PasswordFileControl>> {
    Ok(scan(xml)?
        .controls
        .into_iter()
        .map(|item| item.control)
        .collect())
}

pub fn insert_password_file_control_xml(
    xml: &str,
    form_index: usize,
    control: &PasswordFileControl,
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

pub fn replace_password_file_control_xml(
    xml: &str,
    index: usize,
    replacement: &PasswordFileControl,
) -> Result<String> {
    validate_control(replacement)?;
    let scan = scan(xml)?;
    let old = scan.controls.get(index).ok_or_else(|| {
        Error::InvalidFormat(format!("password/file control {index} is out of bounds"))
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

pub fn remove_password_file_control_xml(xml: &str, index: usize) -> Result<String> {
    let scan = scan(xml)?;
    let old = scan.controls.get(index).ok_or_else(|| {
        Error::InvalidFormat(format!("password/file control {index} is out of bounds"))
    })?;
    apply(xml, old.span.clone(), "")
}

fn password_xml(value: &PasswordControl) -> Result<String> {
    validate_password(value)?;
    let mut out = control_start("password", &value.name, &value.xml_id, &value.metadata);
    push_bool(&mut out, "form:disabled", value.disabled);
    push_u64(&mut out, "form:max-length", value.max_length);
    push_bool(&mut out, "form:printable", value.printable);
    push_u64(&mut out, "form:tab-index", value.tab_index);
    push_bool(&mut out, "form:tab-stop", value.tab_stop);
    push_string(&mut out, "form:title", value.title.as_deref());
    push_string(&mut out, "form:value", value.value.as_deref());
    push_bool(
        &mut out,
        "form:convert-empty-to-null",
        value.convert_empty_to_null,
    );
    push_string(&mut out, "form:linked-cell", value.linked_cell.as_deref());
    if let Some(echo) = value.echo_char {
        let echo = echo.to_string();
        push_string(&mut out, "form:echo-char", Some(&echo));
    }
    out.push_str("/>");
    Ok(out)
}

fn file_xml(value: &FileControl) -> Result<String> {
    validate_file(value)?;
    let mut out = control_start("file", &value.name, &value.xml_id, &value.metadata);
    push_string(
        &mut out,
        "form:current-value",
        value.current_value.as_deref(),
    );
    push_bool(&mut out, "form:disabled", value.disabled);
    push_u64(&mut out, "form:max-length", value.max_length);
    push_bool(&mut out, "form:printable", value.printable);
    push_bool(&mut out, "form:readonly", value.readonly);
    push_u64(&mut out, "form:tab-index", value.tab_index);
    push_bool(&mut out, "form:tab-stop", value.tab_stop);
    push_string(&mut out, "form:title", value.title.as_deref());
    push_string(&mut out, "form:value", value.value.as_deref());
    push_string(&mut out, "form:linked-cell", value.linked_cell.as_deref());
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

fn validate_control(value: &PasswordFileControl) -> Result<()> {
    match value {
        PasswordFileControl::Password(value) => validate_password(value),
        PasswordFileControl::File(value) => validate_file(value),
    }
}

fn validate_password(value: &PasswordControl) -> Result<()> {
    validate_identity(&value.name, &value.xml_id, &value.metadata)?;
    validate_optional("password title", value.title.as_deref(), MAX_STRING)?;
    validate_optional("password value", value.value.as_deref(), MAX_STRING)?;
    validate_optional(
        "password linked cell",
        value.linked_cell.as_deref(),
        MAX_REFERENCE,
    )?;
    if let Some(echo) = value.echo_char
        && !is_xml_char(echo)
    {
        return invalid("password echo-char is not a legal XML character");
    }
    Ok(())
}

fn validate_file(value: &FileControl) -> Result<()> {
    validate_identity(&value.name, &value.xml_id, &value.metadata)?;
    validate_optional(
        "file current value",
        value.current_value.as_deref(),
        MAX_STRING,
    )?;
    validate_optional("file title", value.title.as_deref(), MAX_STRING)?;
    validate_optional("file value", value.value.as_deref(), MAX_STRING)?;
    validate_optional(
        "file linked cell",
        value.linked_cell.as_deref(),
        MAX_REFERENCE,
    )
}

fn validate_identity(name: &str, xml_id: &str, metadata: &GenericControlMetadata) -> Result<()> {
    validate_name("password/file control name", name)?;
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

fn validate_controls(controls: &[PasswordFileControl]) -> Result<()> {
    if controls.len() > MAX_CONTROLS {
        return invalid("too many password/file controls");
    }
    let mut names = Vec::<&str>::new();
    let mut ids = Vec::<&str>::new();
    let mut aggregate = 0usize;
    for control in controls {
        validate_control(control)?;
        if names.contains(&control.name()) {
            return invalid(format!(
                "duplicate password/file control name '{}'",
                control.name()
            ));
        }
        if ids.contains(&control.xml_id()) {
            return invalid(format!(
                "duplicate password/file control xml:id '{}'",
                control.xml_id()
            ));
        }
        names.push(control.name());
        ids.push(control.xml_id());
        aggregate = aggregate.saturating_add(control_size(control));
        if aggregate > MAX_AGGREGATE {
            return invalid("password/file control strings exceed 16 MiB");
        }
    }
    Ok(())
}

fn control_size(control: &PasswordFileControl) -> usize {
    match control {
        PasswordFileControl::Password(value) => {
            identity_size(&value.name, &value.xml_id, &value.metadata).saturating_add(options_size(
                &[
                    value.title.as_ref(),
                    value.value.as_ref(),
                    value.linked_cell.as_ref(),
                ],
            ))
        },
        PasswordFileControl::File(value) => {
            identity_size(&value.name, &value.xml_id, &value.metadata).saturating_add(options_size(
                &[
                    value.current_value.as_ref(),
                    value.title.as_ref(),
                    value.value.as_ref(),
                    value.linked_cell.as_ref(),
                ],
            ))
        },
    }
}

fn identity_size(name: &str, id: &str, metadata: &GenericControlMetadata) -> usize {
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
}

fn options_size(values: &[Option<&String>]) -> usize {
    values
        .iter()
        .flatten()
        .fold(0usize, |sum, value| sum.saturating_add(value.len()))
}

#[derive(Clone)]
struct ControlLocation {
    span: Range<usize>,
    form: usize,
    control: PasswordFileControl,
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
        return invalid("password/file form XML exceeds 64 MiB");
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
                Error::InvalidFormat(format!("invalid password/file form XML: {error}"))
            })?;
        let namespace = match namespace {
            ResolveResult::Bound(uri) => Some(String::from_utf8_lossy(uri.as_ref()).into_owned()),
            _ => None,
        };
        let end = reader.buffer_position() as usize;
        match event {
            Event::Start(ref element) => {
                if stack.len() >= 128 {
                    return invalid("password/file form XML nesting exceeds 128 levels");
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
                    let form_index = forms.len();
                    form = Some(form_index);
                    forms.push(FormLocation {
                        site: Site::Paired { close_start: 0 },
                        controls: Vec::new(),
                    });
                    form_stack.push(form_index);
                } else if is_target(namespace.as_deref(), local.as_slice()) {
                    if stack.iter().any(|open| open.control.is_some()) {
                        return invalid("password/file controls cannot be nested");
                    }
                    let owner = *form_stack.last().ok_or_else(|| {
                        Error::InvalidFormat("password/file control has no form owner".to_string())
                    })?;
                    let parsed = parse_control(&reader, element, local.as_slice())?;
                    aggregate = aggregate.saturating_add(control_size(&parsed));
                    if aggregate > MAX_AGGREGATE {
                        return invalid("password/file control strings exceed 16 MiB");
                    }
                    if controls.len() >= MAX_CONTROLS {
                        return invalid("too many password/file controls");
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
                        "event and macro content is outside the password/file mutation API",
                    );
                } else if stack.iter().any(|open| open.control.is_some())
                    && !is_property(namespace.as_deref(), local.as_slice())
                {
                    return invalid("unexpected child element in password/file control");
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
                        return invalid("password/file controls cannot be nested");
                    }
                    let owner = *form_stack.last().ok_or_else(|| {
                        Error::InvalidFormat("password/file control has no form owner".to_string())
                    })?;
                    let parsed = parse_control(&reader, element, local.as_slice())?;
                    aggregate = aggregate.saturating_add(control_size(&parsed));
                    if aggregate > MAX_AGGREGATE {
                        return invalid("password/file control strings exceed 16 MiB");
                    }
                    if controls.len() >= MAX_CONTROLS {
                        return invalid("too many password/file controls");
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
                        "event and macro content is outside the password/file mutation API",
                    );
                } else if stack.iter().any(|open| open.control.is_some())
                    && !is_property(namespace.as_deref(), local.as_slice())
                {
                    return invalid("unexpected child element in password/file control");
                }
            },
            Event::Text(text) if stack.iter().any(|open| open.control.is_some()) => {
                let decoded = text.decode().map_err(|error| {
                    Error::InvalidFormat(format!("invalid password/file control text: {error}"))
                })?;
                if !decoded.trim().is_empty() {
                    return invalid("password/file controls cannot contain character data");
                }
            },
            Event::CData(text) if stack.iter().any(|open| open.control.is_some()) => {
                let decoded = text.decode().map_err(|error| {
                    Error::InvalidFormat(format!("invalid password/file control CDATA: {error}"))
                })?;
                if !decoded.trim().is_empty() {
                    return invalid("password/file controls cannot contain CDATA");
                }
            },
            Event::GeneralRef(_) if stack.iter().any(|open| open.control.is_some()) => {
                return invalid("password/file controls cannot contain entity references");
            },
            Event::End(ref element) => {
                let open = stack.pop().ok_or_else(|| {
                    Error::InvalidFormat("password/file form XML stack underflow".to_string())
                })?;
                if open.local.as_slice() != element.local_name().as_ref() {
                    return invalid("mismatched password/file form XML elements");
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
            Event::DocType(_) => {
                return invalid("DOCTYPE is not allowed in password/file form XML");
            },
            Event::Eof => break,
            _ => {},
        }
        previous = end;
        buffer.clear();
    }
    if !stack.is_empty() {
        return invalid("unclosed password/file form XML elements");
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
) -> Result<PasswordFileControl> {
    let attrs = attributes(reader, element)?;
    validate_allowed(
        &attrs,
        if local == b"password" {
            PASSWORD_ATTRS
        } else {
            FILE_ATTRS
        },
    )?;
    let name = required(&attrs, FORM, "name")?;
    let xml_id = required(&attrs, XML, "id")?;
    let metadata = GenericControlMetadata {
        form_id: optional(&attrs, FORM, "id"),
        control_implementation: optional(&attrs, FORM, "control-implementation"),
        xforms_bind: optional(&attrs, XFORMS, "bind"),
    };
    if local == b"password" {
        let mut value = PasswordControl::new(name, xml_id);
        value.metadata = metadata;
        value.disabled = optional_bool(&attrs, FORM, "disabled")?;
        value.max_length = optional_u64(&attrs, FORM, "max-length")?;
        value.printable = optional_bool(&attrs, FORM, "printable")?;
        value.tab_index = optional_u64(&attrs, FORM, "tab-index")?;
        value.tab_stop = optional_bool(&attrs, FORM, "tab-stop")?;
        value.title = optional(&attrs, FORM, "title");
        value.value = optional(&attrs, FORM, "value");
        value.convert_empty_to_null = optional_bool(&attrs, FORM, "convert-empty-to-null")?;
        value.linked_cell = optional(&attrs, FORM, "linked-cell");
        value.echo_char = optional(&attrs, FORM, "echo-char")
            .map(|text| parse_echo_char(&text))
            .transpose()?;
        validate_password(&value)?;
        Ok(value.into())
    } else {
        let mut value = FileControl::new(name, xml_id);
        value.metadata = metadata;
        value.current_value = optional(&attrs, FORM, "current-value");
        value.disabled = optional_bool(&attrs, FORM, "disabled")?;
        value.max_length = optional_u64(&attrs, FORM, "max-length")?;
        value.printable = optional_bool(&attrs, FORM, "printable")?;
        value.readonly = optional_bool(&attrs, FORM, "readonly")?;
        value.tab_index = optional_u64(&attrs, FORM, "tab-index")?;
        value.tab_stop = optional_bool(&attrs, FORM, "tab-stop")?;
        value.title = optional(&attrs, FORM, "title");
        value.value = optional(&attrs, FORM, "value");
        value.linked_cell = optional(&attrs, FORM, "linked-cell");
        validate_file(&value)?;
        Ok(value.into())
    }
}

const PASSWORD_ATTRS: &[(&str, &str)] = &[
    (FORM, "name"),
    (XML, "id"),
    (FORM, "id"),
    (FORM, "control-implementation"),
    (FORM, "input-required"),
    (XFORMS, "bind"),
    (FORM, "disabled"),
    (FORM, "max-length"),
    (FORM, "printable"),
    (FORM, "tab-index"),
    (FORM, "tab-stop"),
    (FORM, "title"),
    (FORM, "value"),
    (FORM, "convert-empty-to-null"),
    (FORM, "linked-cell"),
    (FORM, "echo-char"),
];
const FILE_ATTRS: &[(&str, &str)] = &[
    (FORM, "name"),
    (XML, "id"),
    (FORM, "id"),
    (FORM, "control-implementation"),
    (FORM, "input-required"),
    (XFORMS, "bind"),
    (FORM, "current-value"),
    (FORM, "disabled"),
    (FORM, "max-length"),
    (FORM, "printable"),
    (FORM, "readonly"),
    (FORM, "tab-index"),
    (FORM, "tab-stop"),
    (FORM, "title"),
    (FORM, "value"),
    (FORM, "linked-cell"),
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
    namespace == Some(FORM) && matches!(local, b"password" | b"file")
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
            Error::InvalidFormat(format!("invalid password/file control attribute: {error}"))
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
                    "invalid password/file control attribute value: {error}"
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
                "unexpected password/file control attribute '{}'",
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
            value.parse::<u64>().map_err(|_error| {
                Error::InvalidFormat(format!(
                    "invalid non-negative integer '{value}' for {local}"
                ))
            })
        })
        .transpose()
}
fn parse_echo_char(value: &str) -> Result<char> {
    let mut chars = value.chars();
    let first = chars
        .next()
        .ok_or_else(|| Error::InvalidFormat("password echo-char cannot be empty".to_string()))?;
    if chars.next().is_some() || !is_xml_char(first) {
        return invalid("password echo-char must be one legal XML character");
    }
    Ok(first)
}
fn is_xml_char(value: char) -> bool {
    matches!(value as u32, 9 | 10 | 13 | 32..=0xD7FF | 0xE000..=0xFFFD | 0x10000..=0x10FFFF)
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
    validate_name("password/file control xml:id", value)?;
    let mut chars = value.chars();
    let Some(first) = chars.next() else {
        return invalid("password/file control xml:id cannot be empty");
    };
    if !(first == '_' || first.is_ascii_alphabetic())
        || !chars.all(|ch| ch == '_' || ch == '-' || ch == '.' || ch.is_ascii_alphanumeric())
    {
        return invalid(format!("invalid password/file control xml:id '{value}'"));
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
    if value.chars().any(|ch| !is_xml_char(ch)) {
        return invalid(format!("{label} contains invalid XML characters"));
    }
    Ok(())
}

fn reject_duplicate(
    form: &FormLocation,
    replacement: &PasswordFileControl,
    current: Option<&PasswordFileControl>,
) -> Result<()> {
    for item in &form.controls {
        if current.is_some_and(|value| value.xml_id() == item.control.xml_id()) {
            continue;
        }
        if item.control.name() == replacement.name() {
            return invalid(format!(
                "duplicate password/file control name '{}'",
                replacement.name()
            ));
        }
        if item.control.xml_id() == replacement.xml_id() {
            return invalid(format!(
                "duplicate password/file control xml:id '{}'",
                replacement.xml_id()
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
    if value.contains("xforms:") && !value.contains("xmlns:xforms=") {
        value = value.replacen(' ', &format!(" xmlns:xforms=\"{XFORMS}\" "), 1);
    }
    value
}
fn qname(element: &BytesStart<'_>) -> Result<String> {
    String::from_utf8(element.name().as_ref().to_vec()).map_err(|_error| {
        Error::InvalidFormat("invalid password/file form element name".to_string())
    })
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
        let mut password = PasswordControl::new("Password", "password_1");
        password.value = Some("secret & value".into());
        password.echo_char = Some('●');
        password.max_length = Some(64);
        let mut file = FileControl::new("File", "file_1");
        file.value = Some("../../not-opened.txt".into());
        file.current_value = Some("C:\\private\\not-opened.txt".into());
        file.readonly = Some(true);
        let mut form = PasswordFileForm::new("Main");
        form.add_control(password).unwrap();
        form.add_control(file).unwrap();
        let parsed =
            password_file_controls(&format!("{ROOT}{}{END}", form.to_xml_fragment().unwrap()))
                .unwrap();
        assert_eq!(parsed.len(), 2);
        assert!(
            matches!(&parsed[0], PasswordFileControl::Password(value) if value.echo_char == Some('●'))
        );
        assert!(
            matches!(&parsed[1], PasswordFileControl::File(value) if value.value.as_deref() == Some("../../not-opened.txt"))
        );
    }

    #[test]
    fn odfpy_odfdo_and_libreoffice_shapes_parse() {
        let producer = format!(
            r#"{ROOT}<f:form f:name="Producer"><f:password f:name="odfdo" xml:id="password" f:control-implementation="com.sun.star.form.component.TextField" f:echo-char="*" f:max-length="64"/><f:file f:name="odfpy" xml:id="file" f:control-implementation="com.sun.star.form.component.FileControl" f:value="initial" f:current-value="final"/></f:form>{END}"#
        );
        let parsed = password_file_controls(&producer).unwrap();
        assert_eq!(parsed.len(), 2);
        assert!(
            matches!(&parsed[1], PasswordFileControl::File(value) if value.current_value.as_deref() == Some("final"))
        );
    }

    #[test]
    fn lossless_mutation_and_empty_form_expansion() {
        let xml = format!(
            r#"{ROOT}<f:form f:name="Main"><f:password f:name="Old" xml:id="old"><f:properties><f:property f:property-name="Keep" o:value-type="void"/></f:properties></f:password><!--keep--><f:text f:name="Text" xml:id="text"/></f:form>{END}"#
        );
        let file: PasswordFileControl = FileControl::new("File", "file").into();
        let inserted = insert_password_file_control_xml(&xml, 0, &file).unwrap();
        assert!(inserted.contains("<!--keep-->") && inserted.contains("f:text"));
        let password: PasswordFileControl = PasswordControl::new("Password", "password").into();
        let replaced = replace_password_file_control_xml(&inserted, 0, &password).unwrap();
        let removed = remove_password_file_control_xml(&replaced, 1).unwrap();
        assert_eq!(password_file_controls(&removed).unwrap(), [password]);
        let empty = format!(r#"{ROOT}<f:form f:name="Empty"/>{END}"#);
        assert!(
            insert_password_file_control_xml(&empty, 0, &file)
                .unwrap()
                .contains("</f:form>")
        );
    }

    #[test]
    fn hostile_values_children_duplicates_and_active_content_are_rejected() {
        assert!(PasswordControl::new("P", "1bad").to_xml_fragment().is_err());
        let echo = format!(
            r#"{ROOT}<f:form f:name="Main"><f:password f:name="P" xml:id="p" f:echo-char="**"/></f:form>{END}"#
        );
        assert!(password_file_controls(&echo).is_err());
        let bad_bool = format!(
            r#"{ROOT}<f:form f:name="Main"><f:file f:name="F" xml:id="f" f:readonly="yes"/></f:form>{END}"#
        );
        assert!(password_file_controls(&bad_bool).is_err());
        let foreign = format!(
            r#"{ROOT}<f:form f:name="Main"><f:file f:name="F" xml:id="f"><o:p/></f:file></f:form>{END}"#
        );
        assert!(password_file_controls(&foreign).is_err());
        let event = format!(r#"{ROOT}<f:form f:name="Main"><o:event-listeners/></f:form>{END}"#);
        assert!(password_file_controls(&event).is_err());
        let duplicate = format!(
            r#"{ROOT}<o:p xml:id="same"/><f:form f:name="Main"><f:file f:name="F" xml:id="same"/></f:form>{END}"#
        );
        assert!(password_file_controls(&duplicate).is_err());
        let mut file = FileControl::new("F", "f");
        file.value = Some("x".repeat(MAX_STRING + 1));
        assert!(file.to_xml_fragment().is_err());
    }

    #[test]
    fn builder_and_mutable_document_round_trip() {
        use crate::{Builder, Document, mutable::MutableDocument};
        let mut initial = FileControl::new("File", "file");
        initial.current_value = Some("/host/path/is/not/read".into());
        let mut form = PasswordFileForm::new("Main");
        form.add_control(initial).unwrap();
        let mut builder = Builder::new();
        builder.add_password_file_form(&form).unwrap();
        builder.add_paragraph("body").unwrap();
        let document = Document::from_bytes(builder.build().unwrap()).unwrap();
        let mut mutable = MutableDocument::from_document(document).unwrap();
        assert!(
            matches!(&mutable.password_file_controls().unwrap()[0], PasswordFileControl::File(value) if value.current_value.as_deref() == Some("/host/path/is/not/read"))
        );
        let password: PasswordFileControl = PasswordControl::new("Password", "password").into();
        mutable.insert_password_file_control(0, &password).unwrap();
        let replacement: PasswordFileControl = FileControl::new("Other", "other").into();
        assert!(matches!(
            mutable
                .replace_password_file_control(0, &replacement)
                .unwrap(),
            PasswordFileControl::File(_)
        ));
        assert!(matches!(
            mutable.remove_password_file_control(1).unwrap(),
            PasswordFileControl::Password(_)
        ));
        assert_eq!(mutable.password_file_controls().unwrap(), [replacement]);
    }
}
