//! Typed writing and lossless mutation for `form:grid` controls.

use std::ops::Range;

use litchi_core::{Error, Result};
use quick_xml::NsReader;
use quick_xml::events::{BytesStart, Event};
use quick_xml::name::ResolveResult;

use super::generic_writing::OdfGenericControlMetadata;

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
const MAX_COLUMNS: usize = 16_384;
const MAX_COLUMN_CONTROLS: usize = 256;

#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum OdfGridColumnControlKind {
    Text,
    Textarea,
    Password,
    FormattedText,
    Number,
    Date,
    Time,
    Checkbox,
    Listbox,
    Combobox,
}

impl OdfGridColumnControlKind {
    fn from_local(local: &[u8]) -> Option<Self> {
        match local {
            b"text" => Some(Self::Text),
            b"textarea" => Some(Self::Textarea),
            b"password" => Some(Self::Password),
            b"formatted-text" => Some(Self::FormattedText),
            b"number" => Some(Self::Number),
            b"date" => Some(Self::Date),
            b"time" => Some(Self::Time),
            b"checkbox" => Some(Self::Checkbox),
            b"listbox" => Some(Self::Listbox),
            b"combobox" => Some(Self::Combobox),
            _ => None,
        }
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct OdfGridColumnControl {
    pub kind: OdfGridColumnControlKind,
    pub xml: String,
}

impl OdfGridColumnControl {
    pub fn new(kind: OdfGridColumnControlKind, xml: impl Into<String>) -> Result<Self> {
        let value = Self {
            kind,
            xml: xml.into(),
        };
        validate_column_control(&value)?;
        Ok(value)
    }
    pub fn to_xml_fragment(&self) -> &str {
        &self.xml
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct OdfGridColumn {
    pub name: Option<String>,
    pub control_implementation: Option<String>,
    pub label: Option<String>,
    pub text_style_name: Option<String>,
    pub controls: Vec<OdfGridColumnControl>,
}

impl Default for OdfGridColumn {
    fn default() -> Self {
        Self::new()
    }
}

impl OdfGridColumn {
    pub fn new() -> Self {
        Self {
            name: None,
            control_implementation: None,
            label: None,
            text_style_name: None,
            controls: Vec::new(),
        }
    }
    pub fn add_control(&mut self, control: OdfGridColumnControl) -> Result<()> {
        validate_column_control(&control)?;
        if self.controls.len() >= MAX_COLUMN_CONTROLS {
            return invalid("too many controls in form:column");
        }
        self.controls.push(control);
        Ok(())
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct OdfGridControl {
    pub name: String,
    pub xml_id: String,
    pub metadata: OdfGenericControlMetadata,
    pub disabled: Option<bool>,
    pub printable: Option<bool>,
    pub tab_index: Option<OdfGridNonNegativeInteger>,
    pub tab_stop: Option<bool>,
    pub title: Option<String>,
    pub columns: Vec<OdfGridColumn>,
}

impl OdfGridControl {
    pub fn new(name: impl Into<String>, xml_id: impl Into<String>) -> Self {
        Self {
            name: name.into(),
            xml_id: xml_id.into(),
            metadata: OdfGenericControlMetadata::default(),
            disabled: None,
            printable: None,
            tab_index: None,
            tab_stop: None,
            title: None,
            columns: Vec::new(),
        }
    }

    pub fn to_xml_fragment(&self) -> Result<String> {
        grid_xml(self)
    }
    pub fn add_column(&mut self, column: OdfGridColumn) -> Result<()> {
        validate_column(&column)?;
        if self.columns.len() >= MAX_COLUMNS {
            return invalid("too many form:grid columns");
        }
        self.columns.push(column);
        Ok(())
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct OdfGridForm {
    pub name: String,
    pub controls: Vec<OdfGridControl>,
    pub apply_filter: Option<bool>,
}

impl OdfGridForm {
    pub fn new(name: impl Into<String>) -> Self {
        Self {
            name: name.into(),
            controls: Vec::new(),
            apply_filter: None,
        }
    }

    pub fn add_control(&mut self, control: OdfGridControl) -> Result<()> {
        validate_control(&control)?;
        if self
            .controls
            .iter()
            .any(|existing| existing.name == control.name)
        {
            return invalid(format!("duplicate grid control name '{}'", control.name));
        }
        if self
            .controls
            .iter()
            .any(|existing| existing.xml_id == control.xml_id)
        {
            return invalid(format!(
                "duplicate grid control xml:id '{}'",
                control.xml_id
            ));
        }
        if self.controls.len() >= MAX_CONTROLS {
            return invalid("too many grid controls");
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

pub fn grid_controls(xml: &str) -> Result<Vec<OdfGridControl>> {
    Ok(scan(xml)?
        .controls
        .into_iter()
        .map(|item| item.control)
        .collect())
}

pub fn insert_grid_control_xml(
    xml: &str,
    form_index: usize,
    control: &OdfGridControl,
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

pub fn replace_grid_control_xml(
    xml: &str,
    index: usize,
    replacement: &OdfGridControl,
) -> Result<String> {
    validate_control(replacement)?;
    let scan = scan(xml)?;
    let old = scan
        .controls
        .get(index)
        .ok_or_else(|| Error::InvalidFormat(format!("grid control {index} is out of bounds")))?;
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

pub fn remove_grid_control_xml(xml: &str, index: usize) -> Result<String> {
    let scan = scan(xml)?;
    let old = scan
        .controls
        .get(index)
        .ok_or_else(|| Error::InvalidFormat(format!("grid control {index} is out of bounds")))?;
    apply(xml, old.span.clone(), "")
}

fn grid_xml(value: &OdfGridControl) -> Result<String> {
    validate_control(value)?;
    let mut out = format!(
        r#"<form:grid form:name="{}" xml:id="{}""#,
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
    push_bool(&mut out, "form:disabled", value.disabled);
    push_bool(&mut out, "form:printable", value.printable);
    push_string(
        &mut out,
        "form:tab-index",
        value
            .tab_index
            .as_ref()
            .map(OdfGridNonNegativeInteger::as_str),
    );
    push_bool(&mut out, "form:tab-stop", value.tab_stop);
    push_string(&mut out, "form:title", value.title.as_deref());
    if value.columns.is_empty() {
        out.push_str("/>");
    } else {
        out.push('>');
        for column in &value.columns {
            validate_column(column)?;
            out.push_str("<form:column");
            push_string(&mut out, "form:name", column.name.as_deref());
            push_string(
                &mut out,
                "form:control-implementation",
                column.control_implementation.as_deref(),
            );
            push_string(&mut out, "form:label", column.label.as_deref());
            push_string(
                &mut out,
                "form:text-style-name",
                column.text_style_name.as_deref(),
            );
            out.push('>');
            for control in &column.controls {
                out.push_str(control.to_xml_fragment());
            }
            out.push_str("</form:column>");
        }
        out.push_str("</form:grid>");
    }
    Ok(out)
}

fn validate_control(value: &OdfGridControl) -> Result<()> {
    validate_name("grid control name", &value.name)?;
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
    validate_optional("grid title", value.title.as_deref(), MAX_STRING)?;
    if value.columns.len() > MAX_COLUMNS {
        return invalid("too many form:grid columns");
    }
    for column in &value.columns {
        validate_column(column)?;
    }
    Ok(())
}

fn validate_column(value: &OdfGridColumn) -> Result<()> {
    validate_optional("column name", value.name.as_deref(), MAX_STRING)?;
    validate_optional(
        "column implementation",
        value.control_implementation.as_deref(),
        MAX_REFERENCE,
    )?;
    validate_optional("column label", value.label.as_deref(), MAX_STRING)?;
    validate_optional(
        "column text style",
        value.text_style_name.as_deref(),
        MAX_REFERENCE,
    )?;
    if value.controls.is_empty() {
        return invalid("form:column requires at least one column control");
    }
    if value.controls.len() > MAX_COLUMN_CONTROLS {
        return invalid("too many controls in form:column");
    }
    for control in &value.controls {
        validate_column_control(control)?;
    }
    Ok(())
}

fn validate_controls(controls: &[OdfGridControl]) -> Result<()> {
    if controls.len() > MAX_CONTROLS {
        return invalid("too many grid controls");
    }
    let mut names = Vec::<&str>::new();
    let mut ids = Vec::<&str>::new();
    let mut aggregate = 0usize;
    for control in controls {
        validate_control(control)?;
        if names.contains(&control.name.as_str()) {
            return invalid(format!("duplicate grid control name '{}'", control.name));
        }
        if ids.contains(&control.xml_id.as_str()) {
            return invalid(format!(
                "duplicate grid control xml:id '{}'",
                control.xml_id
            ));
        }
        names.push(&control.name);
        ids.push(&control.xml_id);
        aggregate = aggregate.saturating_add(control_size(control));
        if aggregate > MAX_AGGREGATE {
            return invalid("grid control strings exceed 16 MiB");
        }
    }
    Ok(())
}

fn control_size(value: &OdfGridControl) -> usize {
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
        .saturating_add(
            value
                .columns
                .iter()
                .flat_map(|column| column.controls.iter())
                .map(|control| control.xml.len())
                .sum::<usize>(),
        )
}

#[derive(Clone)]
struct ControlLocation {
    span: Range<usize>,
    form: usize,
    control: OdfGridControl,
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
        return invalid("grid form XML exceeds 64 MiB");
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
            .map_err(|error| Error::InvalidFormat(format!("invalid grid form XML: {error}")))?;
        let namespace = match namespace {
            ResolveResult::Bound(uri) => Some(String::from_utf8_lossy(uri.as_ref()).into_owned()),
            _ => None,
        };
        let end = reader.buffer_position() as usize;
        match event {
            Event::Start(ref element) => {
                if stack.len() >= 128 {
                    return invalid("grid form XML nesting exceeds 128 levels");
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
                } else if namespace.as_deref() == Some(FORM) && local == b"grid" {
                    if stack.iter().any(|open| open.control.is_some()) {
                        return invalid("grid controls cannot be nested");
                    }
                    let owner = *form_stack.last().ok_or_else(|| {
                        Error::InvalidFormat("grid control has no form owner".to_string())
                    })?;
                    let parsed = parse_control(&reader, element)?;
                    aggregate = aggregate.saturating_add(control_size(&parsed));
                    if aggregate > MAX_AGGREGATE {
                        return invalid("grid control strings exceed 16 MiB");
                    }
                    if controls.len() >= MAX_CONTROLS {
                        return invalid("too many grid controls");
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
                    return invalid("event and macro content is outside the grid mutation API");
                } else if stack.iter().any(|open| open.control.is_some())
                    && !is_property(namespace.as_deref(), local.as_slice())
                {
                    return invalid("unexpected child element in grid control");
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
                } else if namespace.as_deref() == Some(FORM) && local == b"grid" {
                    if stack.iter().any(|open| open.control.is_some()) {
                        return invalid("grid controls cannot be nested");
                    }
                    let owner = *form_stack.last().ok_or_else(|| {
                        Error::InvalidFormat("grid control has no form owner".to_string())
                    })?;
                    let parsed = parse_control(&reader, element)?;
                    aggregate = aggregate.saturating_add(control_size(&parsed));
                    if aggregate > MAX_AGGREGATE {
                        return invalid("grid control strings exceed 16 MiB");
                    }
                    if controls.len() >= MAX_CONTROLS {
                        return invalid("too many grid controls");
                    }
                    controls.push(ControlLocation {
                        span: previous..end,
                        form: owner,
                        control: parsed,
                    });
                } else if is_active(namespace.as_deref(), local.as_slice())
                    && !form_stack.is_empty()
                {
                    return invalid("event and macro content is outside the grid mutation API");
                } else if stack.iter().any(|open| open.control.is_some())
                    && !is_property(namespace.as_deref(), local.as_slice())
                {
                    return invalid("unexpected child element in grid control");
                }
            },
            Event::Text(text)
                if stack.iter().any(|open| open.control.is_some()) => {
                    let decoded = text.decode().map_err(|error| {
                        Error::InvalidFormat(format!("invalid grid control text: {error}"))
                    })?;
                    if !decoded.trim().is_empty() {
                        return invalid("grid controls cannot contain character data");
                    }
                },
            Event::CData(text)
                if stack.iter().any(|open| open.control.is_some()) => {
                    let decoded = text.decode().map_err(|error| {
                        Error::InvalidFormat(format!("invalid grid control CDATA: {error}"))
                    })?;
                    if !decoded.trim().is_empty() {
                        return invalid("grid controls cannot contain CDATA");
                    }
                },
            Event::GeneralRef(_) if stack.iter().any(|open| open.control.is_some()) => {
                return invalid("grid controls cannot contain entity references");
            },
            Event::End(ref element) => {
                let open = stack.pop().ok_or_else(|| {
                    Error::InvalidFormat("grid form XML stack underflow".to_string())
                })?;
                if open.local.as_slice() != element.local_name().as_ref() {
                    return invalid("mismatched grid form XML elements");
                }
                if let Some(index) = open.control {
                    controls[index].span.end = end;
                    controls[index].control =
                        parse_grid_fragment(&xml[controls[index].span.clone()])?;
                }
                if let Some(index) = open.form {
                    forms[index].site = Site::Paired {
                        close_start: previous,
                    };
                    form_stack.pop();
                }
            },
            Event::DocType(_) => return invalid("DOCTYPE is not allowed in grid form XML"),
            Event::Eof => break,
            _ => {},
        }
        previous = end;
        buffer.clear();
    }
    if !stack.is_empty() {
        return invalid("unclosed grid form XML elements");
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

fn parse_control(reader: &NsReader<&[u8]>, element: &BytesStart<'_>) -> Result<OdfGridControl> {
    let attrs = attributes(reader, element)?;
    validate_allowed(&attrs, GRID_ATTRS)?;
    let mut value = OdfGridControl::new(
        required(&attrs, FORM, "name")?,
        required(&attrs, XML, "id")?,
    );
    value.metadata = OdfGenericControlMetadata {
        form_id: optional(&attrs, FORM, "id"),
        control_implementation: optional(&attrs, FORM, "control-implementation"),
        xforms_bind: optional(&attrs, XFORMS, "bind"),
    };
    value.disabled = optional_bool(&attrs, FORM, "disabled")?;
    value.printable = optional_bool(&attrs, FORM, "printable")?;
    value.tab_index = optional(&attrs, FORM, "tab-index")
        .map(OdfGridNonNegativeInteger::new)
        .transpose()
        .map_err(Error::InvalidFormat)?;
    value.tab_stop = optional_bool(&attrs, FORM, "tab-stop")?;
    value.title = optional(&attrs, FORM, "title");
    validate_control(&value)?;
    Ok(value)
}

const GRID_ATTRS: &[(&str, &str)] = &[
    (FORM, "name"),
    (XML, "id"),
    (FORM, "id"),
    (FORM, "control-implementation"),
    (XFORMS, "bind"),
    (FORM, "disabled"),
    (FORM, "printable"),
    (FORM, "tab-index"),
    (FORM, "tab-stop"),
    (FORM, "title"),
];

const COLUMN_ATTRS: &[(&str, &str)] = &[
    (FORM, "name"),
    (FORM, "control-implementation"),
    (FORM, "label"),
    (FORM, "text-style-name"),
];

fn parse_grid_fragment(raw: &str) -> Result<OdfGridControl> {
    let xml = bind_fragment(raw.to_string());
    let mut reader = NsReader::from_str(&xml);
    let mut buffer = Vec::new();
    let mut previous = 0usize;
    let mut depth = 0usize;
    let mut grid = None;
    let mut column = None::<OdfGridColumn>;
    let mut capture = None::<(OdfGridColumnControlKind, usize, usize)>;
    let mut skip_depth = None::<usize>;
    let mut saw_column = false;
    loop {
        let (namespace, event) = reader
            .read_resolved_event_into(&mut buffer)
            .map_err(|error| {
                Error::InvalidFormat(format!("invalid form:grid fragment: {error}"))
            })?;
        let namespace = match namespace {
            ResolveResult::Bound(uri) => Some(String::from_utf8_lossy(uri.as_ref()).into_owned()),
            _ => None,
        };
        let end = reader.buffer_position() as usize;
        match event {
            Event::Start(ref element) => {
                let local = element.local_name().as_ref().to_vec();
                if is_active(namespace.as_deref(), &local) {
                    return invalid("event and macro content is outside the grid mutation API");
                }
                if depth == 0 {
                    if namespace.as_deref() != Some(FORM) || local != b"grid" {
                        return invalid("expected one form:grid root");
                    }
                    grid = Some(parse_control(&reader, element)?);
                } else if skip_depth.is_some() || capture.is_some() {
                } else if depth == 1
                    && namespace.as_deref() == Some(FORM)
                    && local == b"properties"
                    && !saw_column
                {
                    skip_depth = Some(depth + 1);
                } else if depth == 1 && namespace.as_deref() == Some(FORM) && local == b"column" {
                    saw_column = true;
                    column = Some(parse_column(&reader, element)?);
                } else if depth == 2 && column.is_some() && namespace.as_deref() == Some(FORM) {
                    let kind = OdfGridColumnControlKind::from_local(&local).ok_or_else(|| {
                        Error::InvalidFormat("invalid form:column control kind".to_string())
                    })?;
                    capture = Some((kind, previous, depth + 1));
                } else {
                    return invalid("invalid form:grid child order or nesting");
                }
                depth += 1;
            },
            Event::Empty(ref element) => {
                let local = element.local_name().as_ref().to_vec();
                if is_active(namespace.as_deref(), &local) {
                    return invalid("event and macro content is outside the grid mutation API");
                }
                if skip_depth.is_some() || capture.is_some() {
                } else if depth == 1
                    && namespace.as_deref() == Some(FORM)
                    && local == b"properties"
                    && !saw_column
                {
                } else if depth == 1 && namespace.as_deref() == Some(FORM) && local == b"column" {
                    return invalid("form:column requires at least one column control");
                } else if depth == 2 && column.is_some() && namespace.as_deref() == Some(FORM) {
                    let kind = OdfGridColumnControlKind::from_local(&local).ok_or_else(|| {
                        Error::InvalidFormat("invalid form:column control kind".to_string())
                    })?;
                    let control = OdfGridColumnControl::new(kind, &xml[previous..end])?;
                    column.as_mut().unwrap().add_control(control)?;
                } else {
                    return invalid("invalid empty element in form:grid");
                }
            },
            Event::End(_) => {
                if depth == 0 {
                    return invalid("form:grid XML stack underflow");
                }
                depth -= 1;
                if let Some((kind, start, capture_depth)) = capture {
                    if depth + 1 == capture_depth {
                        let control = OdfGridColumnControl::new(kind, &xml[start..end])?;
                        column.as_mut().unwrap().add_control(control)?;
                        capture = None;
                    }
                } else if skip_depth == Some(depth + 1) {
                    skip_depth = None;
                } else if depth == 1 && column.is_some() {
                    let completed = column.take().unwrap();
                    validate_column(&completed)?;
                    grid.as_mut().unwrap().add_column(completed)?;
                }
            },
            Event::Text(ref text) if capture.is_none() && skip_depth.is_none() => {
                let decoded = text
                    .decode()
                    .map_err(|error| Error::InvalidFormat(format!("invalid grid text: {error}")))?;
                if !decoded.trim().is_empty() {
                    return invalid("form:grid and form:column cannot contain character data");
                }
            },
            Event::CData(_) | Event::GeneralRef(_) if capture.is_none() => {
                return invalid("form:grid cannot contain CDATA or entity references");
            },
            Event::DocType(_) => return invalid("DOCTYPE is not allowed in form:grid XML"),
            Event::Eof => break,
            _ => {},
        }
        previous = end;
        buffer.clear();
    }
    let grid = grid.ok_or_else(|| Error::InvalidFormat("missing form:grid root".to_string()))?;
    validate_control(&grid)?;
    Ok(grid)
}

fn parse_column(reader: &NsReader<&[u8]>, element: &BytesStart<'_>) -> Result<OdfGridColumn> {
    let attrs = attributes(reader, element)?;
    validate_allowed(&attrs, COLUMN_ATTRS)?;
    let mut column = OdfGridColumn::new();
    column.name = optional(&attrs, FORM, "name");
    column.control_implementation = optional(&attrs, FORM, "control-implementation");
    column.label = optional(&attrs, FORM, "label");
    column.text_style_name = optional(&attrs, FORM, "text-style-name");
    Ok(column)
}

fn validate_column_control(value: &OdfGridColumnControl) -> Result<()> {
    validate_string("column control XML", &value.xml, MAX_STRING)?;
    let wrapped = format!(
        r#"<form:form xmlns:form="{FORM}" xmlns:xforms="{XFORMS}" form:name="Column">{}</form:form>"#,
        value.xml
    );
    let count = match value.kind {
        OdfGridColumnControlKind::Text | OdfGridColumnControlKind::Textarea => {
            crate::text_controls(&wrapped)?.len()
        },
        OdfGridColumnControlKind::Password => crate::password_file_controls(&wrapped)?.len(),
        OdfGridColumnControlKind::Checkbox => crate::interactive_controls(&wrapped)?.len(),
        OdfGridColumnControlKind::Listbox | OdfGridColumnControlKind::Combobox => {
            crate::selection_controls(&wrapped)?.len()
        },
        OdfGridColumnControlKind::FormattedText
        | OdfGridColumnControlKind::Number
        | OdfGridColumnControlKind::Date
        | OdfGridColumnControlKind::Time => crate::typed_value_controls(&wrapped)?.len(),
    };
    if count != 1 {
        return invalid("column control fragment does not match its declared kind");
    }
    Ok(())
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

fn is_property(namespace: Option<&str>, _local: &[u8]) -> bool {
    namespace == Some(FORM)
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
            Error::InvalidFormat(format!("invalid grid control attribute: {error}"))
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
                Error::InvalidFormat(format!("invalid grid control attribute value: {error}"))
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
                "unexpected grid control attribute '{}'",
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
    validate_name("grid control xml:id", value)?;
    let mut chars = value.chars();
    let first = chars.next().unwrap();
    if !(first == '_' || first.is_ascii_alphabetic())
        || !chars.all(|ch| ch == '_' || ch == '-' || ch == '.' || ch.is_ascii_alphanumeric())
    {
        return invalid(format!("invalid grid control xml:id '{value}'"));
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
    replacement: &OdfGridControl,
    current: Option<&OdfGridControl>,
) -> Result<()> {
    for item in &form.controls {
        if current.is_some_and(|value| value.xml_id == item.control.xml_id) {
            continue;
        }
        if item.control.name == replacement.name {
            return invalid(format!(
                "duplicate grid control name '{}'",
                replacement.name
            ));
        }
        if item.control.xml_id == replacement.xml_id {
            return invalid(format!(
                "duplicate grid control xml:id '{}'",
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
        .map_err(|_| Error::InvalidFormat("invalid grid form element name".to_string()))
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

use std::fmt;
use std::str::FromStr;

const MAX_GRID_INTEGER_DIGITS: usize = 4096;

fn canonical_nonnegative(value: &str) -> std::result::Result<String, String> {
    if value.is_empty() || value.len() > MAX_GRID_INTEGER_DIGITS + 1 {
        return Err("nonNegativeInteger lexical form is empty or exceeds the safety limit".into());
    }
    let digits = match value.as_bytes()[0] {
        b'+' => &value[1..],
        b'-' => {
            if value[1..].bytes().all(|byte| byte == b'0') {
                return Ok("0".into());
            }
            return Err("nonNegativeInteger cannot be negative".into());
        },
        _ => value,
    };
    if digits.is_empty()
        || digits.len() > MAX_GRID_INTEGER_DIGITS
        || !digits.bytes().all(|byte| byte.is_ascii_digit())
    {
        return Err(format!("invalid nonNegativeInteger lexical form {value:?}"));
    }
    let digits = digits.trim_start_matches('0');
    Ok(if digits.is_empty() {
        "0".into()
    } else {
        digits.into()
    })
}

#[derive(Debug, Clone, PartialEq, Eq, Hash)]
pub struct OdfGridNonNegativeInteger(String);
impl OdfGridNonNegativeInteger {
    pub fn new(value: impl AsRef<str>) -> std::result::Result<Self, String> {
        canonical_nonnegative(value.as_ref()).map(Self)
    }
    pub fn as_str(&self) -> &str {
        &self.0
    }
}
impl fmt::Display for OdfGridNonNegativeInteger {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        f.write_str(&self.0)
    }
}
impl FromStr for OdfGridNonNegativeInteger {
    type Err = String;
    fn from_str(value: &str) -> std::result::Result<Self, Self::Err> {
        Self::new(value)
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    fn document(grid: &str) -> String {
        format!(
            r#"<office:document-content xmlns:office="{OFFICE}" xmlns:form="{FORM}" xmlns:xforms="{XFORMS}"><office:body><office:text><form:form form:name="Data">{grid}</form:form><office:tail/></office:text></office:body></office:document-content>"#
        )
    }

    #[test]
    fn schema_and_producer_shaped_grid_parses_all_column_control_choices() {
        let xml = document(concat!(
            r#"<form:grid form:name="Grid" xml:id="grid1" form:disabled="false" form:printable="true" form:tab-index="+0003" form:tab-stop="true" form:title="Records">"#,
            r#"<form:column form:name="Primary" form:label="Primary" form:text-style-name="GridLabel">"#,
            r#"<form:text form:name="Text" xml:id="c1"/><form:textarea form:name="Area" xml:id="c2"/><form:password form:name="Password" xml:id="c3"/>"#,
            r#"<form:formatted-text form:name="Formatted" xml:id="c4"/><form:number form:name="Number" xml:id="c5"/><form:date form:name="Date" xml:id="c6"/><form:time form:name="Time" xml:id="c7"/>"#,
            r#"<form:checkbox form:name="Check" xml:id="c8"/><form:listbox form:name="List" xml:id="c9"/><form:combobox form:name="Combo" xml:id="c10"/>"#,
            r#"</form:column></form:grid>"#
        ));
        let grid = grid_controls(&xml).unwrap().remove(0);
        assert_eq!(grid.tab_index.as_ref().unwrap().as_str(), "3");
        assert_eq!(grid.columns.len(), 1);
        assert_eq!(
            grid.columns[0]
                .controls
                .iter()
                .map(|c| c.kind)
                .collect::<Vec<_>>(),
            [
                OdfGridColumnControlKind::Text,
                OdfGridColumnControlKind::Textarea,
                OdfGridColumnControlKind::Password,
                OdfGridColumnControlKind::FormattedText,
                OdfGridColumnControlKind::Number,
                OdfGridColumnControlKind::Date,
                OdfGridColumnControlKind::Time,
                OdfGridColumnControlKind::Checkbox,
                OdfGridColumnControlKind::Listbox,
                OdfGridColumnControlKind::Combobox,
            ]
        );
        assert_eq!(grid.columns[0].label.as_deref(), Some("Primary"));
        assert!(grid.to_xml_fragment().unwrap().contains("<form:column"));
    }

    #[test]
    fn rejects_column_cardinality_order_namespaces_active_content_and_wrong_kinds() {
        assert!(grid_controls(&document(r#"<form:grid form:name="G" xml:id="g"><form:column form:label="Empty"/></form:grid>"#)).is_err());
        assert!(grid_controls(&document(r#"<form:grid form:name="G" xml:id="g"><form:column><form:file form:name="F" xml:id="f"/></form:column></form:grid>"#)).is_err());
        assert!(grid_controls(&document(r#"<form:grid form:name="G" xml:id="g"><form:column><form:text form:name="T" xml:id="t"/></form:column><form:properties/></form:grid>"#)).is_err());
        assert!(
            grid_controls(&document(
                r#"<form:grid xmlns:e="urn:evil" form:name="G" xml:id="g" e:title="bad"/>"#
            ))
            .is_err()
        );
        assert!(
            grid_controls(&document(
                r#"<form:grid form:name="G" xml:id="g"><office:event-listeners/></form:grid>"#
            ))
            .is_err()
        );
        assert!(
            grid_controls(&format!(
                "<!DOCTYPE x>{}",
                document(r#"<form:grid form:name="G" xml:id="g"/>"#)
            ))
            .is_err()
        );
        assert!(OdfGridColumnControl::new(OdfGridColumnControlKind::Number, r#"<form:text xmlns:form="urn:oasis:names:tc:opendocument:xmlns:form:1.0" form:name="T" xml:id="t"/>"#).is_err());
        assert_eq!(OdfGridNonNegativeInteger::new("-0").unwrap().as_str(), "0");
        assert!(OdfGridNonNegativeInteger::new("-1").is_err());
    }

    #[test]
    fn lossless_mutation_builder_and_mutable_document_round_trip() {
        let xml = document(
            r#"<form:grid form:name="Old" xml:id="old"><form:column form:label="A"><form:text form:name="Text" xml:id="text"/></form:column></form:grid>"#,
        );
        let nested = OdfGridColumnControl::new(
            OdfGridColumnControlKind::Number,
            r#"<form:number form:name="Number" xml:id="number"/>"#,
        )
        .unwrap();
        let mut column = OdfGridColumn::new();
        column.label = Some("Amount".into());
        column.add_control(nested).unwrap();
        let mut inserted = OdfGridControl::new("Inserted", "inserted");
        inserted.add_column(column).unwrap();
        let updated = insert_grid_control_xml(&xml, 0, &inserted).unwrap();
        assert!(updated.contains("<office:tail/>"));
        let replacement = OdfGridControl::new("Replacement", "replacement");
        let updated = replace_grid_control_xml(&updated, 0, &replacement).unwrap();
        assert_eq!(
            grid_controls(&remove_grid_control_xml(&updated, 1).unwrap()).unwrap(),
            [replacement.clone()]
        );

        let mut form = OdfGridForm::new("Grids");
        form.add_control(inserted.clone()).unwrap();
        let mut builder = crate::DocumentBuilder::new();
        builder.add_grid_form(&form).unwrap();
        let document = crate::Document::from_bytes(builder.build().unwrap()).unwrap();
        let mut mutable = crate::MutableDocument::from_document(document).unwrap();
        assert_eq!(mutable.grid_controls().unwrap(), [inserted]);
        mutable.insert_grid_control(0, &replacement).unwrap();
        assert_eq!(mutable.remove_grid_control(1).unwrap().name, "Replacement");
    }
}
