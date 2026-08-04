//! Package-safe, lossless mutation of form trees and typed controls.

use crate::core::OwnedPackage;
use crate::embedded_chart::{rebuild_package, splice};
use crate::{
    OdfForm, OdfFormNode, OdfFormPart, OdfFormProperty, OdfGenericFormControl, OdfGridControl,
    OdfImageFrameControl, OdfInteractiveControl, OdfPasswordFileControl, OdfSelectionControl,
    OdfTextControl, OdfTypedValueControl, OdfValueRangeControl, OdfVisualControl,
};
use litchi_core::{Error, Result};
use quick_xml::events::Event;
use quick_xml::name::{Namespace, ResolveResult};
use quick_xml::reader::NsReader;
use std::collections::HashSet;

const OFFICE_NS: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:office:1.0";
const FORM_NS: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:form:1.0";
const OFFICE: &str = "urn:oasis:names:tc:opendocument:xmlns:office:1.0";
const FORM: &str = "urn:oasis:names:tc:opendocument:xmlns:form:1.0";
const TEXT: &str = "urn:oasis:names:tc:opendocument:xmlns:text:1.0";
const XLINK: &str = "http://www.w3.org/1999/xlink";
const XFORMS: &str = "http://www.w3.org/2002/xforms";
const SCRIPT: &str = "urn:oasis:names:tc:opendocument:xmlns:script:1.0";
const MAX_XML: usize = 64 * 1024 * 1024;
const MAX_CHILDREN: usize = 65_536;
const MAX_DEPTH: usize = 128;
const MAX_STRING: usize = 1024 * 1024;

#[derive(Debug, Clone, PartialEq)]
#[non_exhaustive]
pub enum OdfAuthoredFormControl {
    Text(OdfTextControl),
    TypedValue(OdfTypedValueControl),
    Selection(OdfSelectionControl),
    Interactive(OdfInteractiveControl),
    Visual(OdfVisualControl),
    PasswordFile(OdfPasswordFileControl),
    Generic(OdfGenericFormControl),
    Grid(OdfGridControl),
    ImageFrame(OdfImageFrameControl),
    ValueRange(OdfValueRangeControl),
}

impl OdfAuthoredFormControl {
    pub fn to_xml_fragment(&self) -> Result<String> {
        let fragment = match self {
            Self::Text(value) => value.to_xml_fragment()?,
            Self::TypedValue(value) => value.to_xml_fragment()?,
            Self::Selection(value) => value.to_xml_fragment()?,
            Self::Interactive(value) => value.to_xml_fragment()?,
            Self::Visual(value) => value.to_xml_fragment()?,
            Self::PasswordFile(value) => value.to_xml_fragment()?,
            Self::Generic(value) => value.to_xml_fragment()?,
            Self::Grid(value) => value.to_xml_fragment()?,
            Self::ImageFrame(value) => value.to_xml_fragment()?,
            Self::ValueRange(value) => value.to_xml_fragment()?,
        };
        bind_fragment(fragment)
    }
}

macro_rules! authored_from {
    ($variant:ident, $type:ty) => {
        impl From<$type> for OdfAuthoredFormControl {
            fn from(value: $type) -> Self {
                Self::$variant(value)
            }
        }
    };
}
authored_from!(Text, OdfTextControl);
authored_from!(TypedValue, OdfTypedValueControl);
authored_from!(Selection, OdfSelectionControl);
authored_from!(Interactive, OdfInteractiveControl);
authored_from!(Visual, OdfVisualControl);
authored_from!(PasswordFile, OdfPasswordFileControl);
authored_from!(Generic, OdfGenericFormControl);
authored_from!(Grid, OdfGridControl);
authored_from!(ImageFrame, OdfImageFrameControl);
authored_from!(ValueRange, OdfValueRangeControl);

#[derive(Debug, Clone, PartialEq)]
#[non_exhaustive]
pub enum OdfAuthoredFormNode {
    Form(OdfAuthoredForm),
    Control(OdfAuthoredFormControl),
}

#[derive(Debug, Clone, PartialEq)]
pub struct OdfAuthoredForm {
    pub name: String,
    pub xml_id: Option<String>,
    pub form_id: Option<String>,
    pub control_implementation: Option<String>,
    pub apply_filter: Option<bool>,
    pub command_type: Option<String>,
    pub command: Option<String>,
    pub datasource: Option<String>,
    pub href: Option<String>,
    pub properties: Vec<OdfFormProperty>,
    pub children: Vec<OdfAuthoredFormNode>,
}

impl OdfAuthoredForm {
    pub fn new(name: impl Into<String>) -> Self {
        Self {
            name: name.into(),
            xml_id: None,
            form_id: None,
            control_implementation: None,
            apply_filter: None,
            command_type: None,
            command: None,
            datasource: None,
            href: None,
            properties: Vec::new(),
            children: Vec::new(),
        }
    }

    pub fn add_control(&mut self, control: impl Into<OdfAuthoredFormControl>) -> Result<()> {
        if self.children.len() >= MAX_CHILDREN {
            return invalid("form exceeds child limit");
        }
        self.children
            .push(OdfAuthoredFormNode::Control(control.into()));
        Ok(())
    }

    pub fn add_form(&mut self, form: OdfAuthoredForm) -> Result<()> {
        if self.children.len() >= MAX_CHILDREN {
            return invalid("form exceeds child limit");
        }
        self.children.push(OdfAuthoredFormNode::Form(form));
        Ok(())
    }

    pub fn to_xml_fragment(&self) -> Result<String> {
        let mut nodes = 0usize;
        serialize_form(self, 1, &mut nodes)
    }
}

#[derive(Clone, Copy)]
#[allow(dead_code, reason = "shared form codec supports every ODF host family")]
pub(crate) enum FormHost {
    Text,
    Spreadsheet,
    Presentation,
}

#[derive(Clone)]
struct Span {
    start: usize,
    end: usize,
    close_start: Option<usize>,
    qname: String,
    parent: Parent,
}

#[derive(Clone, Copy, PartialEq, Eq)]
enum Parent {
    Group(usize),
    Form(usize),
}

struct Scan {
    groups: Vec<Span>,
    forms: Vec<Span>,
    controls: Vec<Span>,
}

pub(crate) fn add_form(
    package: &OwnedPackage,
    content: &str,
    styles: Option<&str>,
    host: FormHost,
    group_index: usize,
    parent_form: Option<usize>,
    form: &OdfAuthoredForm,
) -> Result<(Vec<u8>, usize)> {
    let fragment = form.to_xml_fragment()?;
    let scan = scan(content)?;
    let index = scan.forms.len();
    let updated = if let Some(parent) = parent_form {
        let site = scan
            .forms
            .get(parent)
            .ok_or_else(|| bounds("form", parent, scan.forms.len()))?;
        insert_child(content, site, &fragment)?
    } else if let Some(site) = scan.groups.get(group_index) {
        insert_child(content, site, &fragment)?
    } else if scan.groups.is_empty() && group_index == 0 {
        let group = format!(
            "<office:forms xmlns:office=\"{OFFICE}\" xmlns:form=\"{FORM}\" xmlns:text=\"{TEXT}\" xmlns:xlink=\"{XLINK}\" xmlns:xforms=\"{XFORMS}\" xmlns:script=\"{SCRIPT}\">{fragment}</office:forms>"
        );
        insert_group(content, host, &group)?
    } else {
        return Err(bounds("form group", group_index, scan.groups.len()));
    };
    rebuild_validated(package, &updated, styles).map(|bytes| (bytes, index))
}

pub(crate) fn replace_form(
    package: &OwnedPackage,
    content: &str,
    styles: Option<&str>,
    index: usize,
    form: &OdfAuthoredForm,
) -> Result<Vec<u8>> {
    let scan = scan(content)?;
    let span = scan
        .forms
        .get(index)
        .ok_or_else(|| bounds("form", index, scan.forms.len()))?;
    let updated = splice(content, span.start, span.end, &form.to_xml_fragment()?)?;
    rebuild_validated(package, &updated, styles)
}

pub(crate) fn remove_form(
    package: &OwnedPackage,
    content: &str,
    styles: Option<&str>,
    index: usize,
) -> Result<Vec<u8>> {
    let scan = scan(content)?;
    let span = scan
        .forms
        .get(index)
        .ok_or_else(|| bounds("form", index, scan.forms.len()))?;
    let updated = splice(content, span.start, span.end, "")?;
    rebuild_validated(package, &updated, styles)
}

pub(crate) fn move_form(
    package: &OwnedPackage,
    content: &str,
    styles: Option<&str>,
    from: usize,
    to: usize,
) -> Result<Vec<u8>> {
    let scan = scan(content)?;
    let first = scan
        .forms
        .get(from)
        .ok_or_else(|| bounds("form", from, scan.forms.len()))?;
    let second = scan
        .forms
        .get(to)
        .ok_or_else(|| bounds("form", to, scan.forms.len()))?;
    if first.parent != second.parent {
        return invalid("forms can only be reordered among siblings");
    }
    rebuild_validated(package, &relocate(content, first, second)?, styles)
}

pub(crate) fn add_control(
    package: &OwnedPackage,
    content: &str,
    styles: Option<&str>,
    form_index: usize,
    control: &OdfAuthoredFormControl,
) -> Result<(Vec<u8>, usize)> {
    let scan = scan(content)?;
    let form = scan
        .forms
        .get(form_index)
        .ok_or_else(|| bounds("form", form_index, scan.forms.len()))?;
    let index = scan.controls.len();
    let updated = insert_child(content, form, &control.to_xml_fragment()?)?;
    rebuild_validated(package, &updated, styles).map(|bytes| (bytes, index))
}

pub(crate) fn replace_control(
    package: &OwnedPackage,
    content: &str,
    styles: Option<&str>,
    index: usize,
    control: &OdfAuthoredFormControl,
) -> Result<Vec<u8>> {
    let scan = scan(content)?;
    let span = scan
        .controls
        .get(index)
        .ok_or_else(|| bounds("form control", index, scan.controls.len()))?;
    rebuild_validated(
        package,
        &splice(content, span.start, span.end, &control.to_xml_fragment()?)?,
        styles,
    )
}

pub(crate) fn remove_control(
    package: &OwnedPackage,
    content: &str,
    styles: Option<&str>,
    index: usize,
) -> Result<Vec<u8>> {
    let scan = scan(content)?;
    let span = scan
        .controls
        .get(index)
        .ok_or_else(|| bounds("form control", index, scan.controls.len()))?;
    rebuild_validated(package, &splice(content, span.start, span.end, "")?, styles)
}

pub(crate) fn move_control(
    package: &OwnedPackage,
    content: &str,
    styles: Option<&str>,
    from: usize,
    to: usize,
) -> Result<Vec<u8>> {
    let scan = scan(content)?;
    let first = scan
        .controls
        .get(from)
        .ok_or_else(|| bounds("form control", from, scan.controls.len()))?;
    let second = scan
        .controls
        .get(to)
        .ok_or_else(|| bounds("form control", to, scan.controls.len()))?;
    if first.parent != second.parent {
        return invalid("form controls can only be reordered among siblings");
    }
    rebuild_validated(package, &relocate(content, first, second)?, styles)
}

fn rebuild_validated(
    package: &OwnedPackage,
    content: &str,
    styles: Option<&str>,
) -> Result<Vec<u8>> {
    if content.len() > MAX_XML {
        return invalid("form mutation exceeds XML size limit");
    }
    let mut parts = vec![(content, OdfFormPart::Content)];
    if let Some(styles) = styles {
        parts.push((styles, OdfFormPart::Styles));
    }
    let parsed = crate::form::parse_form_parts(&parts)?;
    validate_unique(&parsed)?;
    rebuild_package(
        package,
        content,
        Vec::new(),
        Vec::new(),
        Vec::new(),
        Vec::new(),
    )
}

fn validate_unique(forms: &crate::OdfForms) -> Result<()> {
    let mut ids = HashSet::new();
    for group in &forms.groups {
        sibling_form_names(&group.forms)?;
        for form in &group.forms {
            unique_form(form, &mut ids)?;
        }
    }
    Ok(())
}

fn sibling_form_names(forms: &[OdfForm]) -> Result<()> {
    let mut names = HashSet::new();
    for form in forms {
        if let Some(name) = &form.name
            && !names.insert(name.as_str())
        {
            return invalid(format!("duplicate sibling form name '{name}'"));
        }
    }
    Ok(())
}

fn unique_form(form: &OdfForm, ids: &mut HashSet<String>) -> Result<()> {
    for id in [form.xml_id.as_deref(), form.form_id.as_deref()]
        .into_iter()
        .flatten()
    {
        if !ids.insert(id.to_string()) {
            return invalid(format!("duplicate form ID '{id}'"));
        }
    }
    let nested: Vec<&OdfForm> = form
        .children
        .iter()
        .filter_map(|node| match node {
            OdfFormNode::Form(value) => Some(value),
            _ => None,
        })
        .collect();
    let mut nested_names = HashSet::new();
    for nested_form in &nested {
        if let Some(name) = &nested_form.name
            && !nested_names.insert(name.as_str())
        {
            return invalid(format!("duplicate sibling form name '{name}'"));
        }
    }
    let mut control_names = HashSet::new();
    for node in &form.children {
        match node {
            OdfFormNode::Form(value) => unique_form(value, ids)?,
            OdfFormNode::Control(value) => {
                if let Some(name) = &value.name
                    && !control_names.insert(name.as_str())
                {
                    return invalid(format!("duplicate sibling form control name '{name}'"));
                }
                for id in [value.xml_id.as_deref(), value.form_id.as_deref()]
                    .into_iter()
                    .flatten()
                {
                    if !ids.insert(id.to_string()) {
                        return invalid(format!("duplicate form control ID '{id}'"));
                    }
                }
            },
        }
    }
    Ok(())
}

fn serialize_form(form: &OdfAuthoredForm, depth: usize, nodes: &mut usize) -> Result<String> {
    if depth > MAX_DEPTH {
        return invalid("authored form nesting exceeds limit");
    }
    *nodes = nodes
        .checked_add(1)
        .ok_or_else(|| Error::InvalidFormat("form node overflow".to_string()))?;
    if *nodes > MAX_CHILDREN {
        return invalid("authored form exceeds node limit");
    }
    validate_name("form name", &form.name)?;
    for value in [
        form.xml_id.as_deref(),
        form.form_id.as_deref(),
        form.control_implementation.as_deref(),
        form.command.as_deref(),
        form.datasource.as_deref(),
        form.href.as_deref(),
    ]
    .into_iter()
    .flatten()
    {
        validate_string(value)?;
    }
    if let Some(value) = &form.command_type
        && !matches!(value.as_str(), "table" | "query" | "command")
    {
        return invalid("form command type must be table, query, or command");
    }
    if form.children.len() > MAX_CHILDREN || form.properties.len() > MAX_CHILDREN {
        return invalid("form exceeds child or property limit");
    }
    let mut out = format!(
        "<form:form xmlns:form=\"{FORM}\" xmlns:office=\"{OFFICE}\" xmlns:text=\"{TEXT}\" xmlns:xlink=\"{XLINK}\" xmlns:xforms=\"{XFORMS}\" xmlns:script=\"{SCRIPT}\""
    );
    attr(&mut out, "form:name", Some(&form.name))?;
    attr(&mut out, "xml:id", form.xml_id.as_deref())?;
    attr(&mut out, "form:id", form.form_id.as_deref())?;
    attr(
        &mut out,
        "form:control-implementation",
        form.control_implementation.as_deref(),
    )?;
    bool_attr(&mut out, "form:apply-filter", form.apply_filter);
    attr(&mut out, "form:command-type", form.command_type.as_deref())?;
    attr(&mut out, "form:command", form.command.as_deref())?;
    attr(&mut out, "form:datasource", form.datasource.as_deref())?;
    if form.href.is_some() {
        out.push_str(" xlink:type=\"simple\" xlink:actuate=\"onRequest\"");
        attr(&mut out, "xlink:href", form.href.as_deref())?;
    }
    if form.properties.is_empty() && form.children.is_empty() {
        out.push_str("/>");
        return Ok(out);
    }
    out.push('>');
    if !form.properties.is_empty() {
        out.push_str("<form:properties>");
        for property in &form.properties {
            out.push_str(&crate::form::property_xml(property)?);
        }
        out.push_str("</form:properties>");
    }
    for child in &form.children {
        match child {
            OdfAuthoredFormNode::Form(value) => {
                out.push_str(&serialize_form(value, depth + 1, nodes)?)
            },
            OdfAuthoredFormNode::Control(value) => {
                *nodes = nodes
                    .checked_add(1)
                    .ok_or_else(|| Error::InvalidFormat("form node overflow".to_string()))?;
                if *nodes > MAX_CHILDREN {
                    return invalid("authored form exceeds node limit");
                }
                out.push_str(&value.to_xml_fragment()?);
            },
        }
    }
    out.push_str("</form:form>");
    Ok(out)
}

fn scan(xml: &str) -> Result<Scan> {
    #[derive(Clone)]
    struct Active {
        depth: usize,
        start: usize,
        qname: String,
        parent: Parent,
    }
    let mut reader = NsReader::from_str(xml);
    let mut buffer = Vec::new();
    let mut depth = 0usize;
    let mut groups = Vec::new();
    let mut forms = Vec::new();
    let mut controls = Vec::new();
    let mut active_groups: Vec<(usize, usize, usize, String)> = Vec::new();
    let mut active_forms: Vec<(usize, usize, Active)> = Vec::new();
    let mut active_control: Option<Active> = None;
    let mut next_group = 0usize;
    let mut next_form = 0usize;
    loop {
        let start = position(&reader)?;
        let (resolved, event) = reader
            .read_resolved_event_into(&mut buffer)
            .map_err(|error| Error::InvalidFormat(format!("invalid form host XML: {error}")))?;
        let ns = namespace(&resolved);
        let end = position(&reader)?;
        match event {
            Event::Start(element) => {
                let local_name = element.local_name();
                let local = local_name.as_ref();
                if ns == NamespaceKind::Office && local == b"forms" {
                    let index = next_group;
                    next_group += 1;
                    active_groups.push((depth, start, index, qname(element.name().as_ref())?));
                } else if ns == NamespaceKind::Form && local == b"form" {
                    let index = next_form;
                    next_form += 1;
                    let parent = active_forms
                        .last()
                        .map(|item| Parent::Form(item.1))
                        .or_else(|| active_groups.last().map(|item| Parent::Group(item.2)))
                        .ok_or_else(|| {
                            Error::InvalidFormat("form:form outside office:forms".to_string())
                        })?;
                    active_forms.push((
                        depth,
                        index,
                        Active {
                            depth,
                            start,
                            qname: qname(element.name().as_ref())?,
                            parent,
                        },
                    ));
                } else if ns == NamespaceKind::Form
                    && control_local(local)
                    && active_control.is_none()
                    && let Some((_, form_index, _)) = active_forms.last()
                {
                    active_control = Some(Active {
                        depth,
                        start,
                        qname: qname(element.name().as_ref())?,
                        parent: Parent::Form(*form_index),
                    });
                }
                depth = depth
                    .checked_add(1)
                    .ok_or_else(|| Error::InvalidFormat("form XML depth overflow".to_string()))?;
                if depth > MAX_DEPTH {
                    return invalid("form XML nesting exceeds limit");
                }
            },
            Event::Empty(element) => {
                let local_name = element.local_name();
                let local = local_name.as_ref();
                if ns == NamespaceKind::Office && local == b"forms" {
                    let index = next_group;
                    next_group += 1;
                    groups.push(Span {
                        start,
                        end,
                        close_start: None,
                        qname: qname(element.name().as_ref())?,
                        parent: Parent::Group(index),
                    });
                } else if ns == NamespaceKind::Form && local == b"form" {
                    let parent = active_forms
                        .last()
                        .map(|item| Parent::Form(item.1))
                        .or_else(|| active_groups.last().map(|item| Parent::Group(item.2)))
                        .ok_or_else(|| {
                            Error::InvalidFormat("form:form outside office:forms".to_string())
                        })?;
                    next_form += 1;
                    forms.push(Span {
                        start,
                        end,
                        close_start: None,
                        qname: qname(element.name().as_ref())?,
                        parent,
                    });
                } else if ns == NamespaceKind::Form
                    && control_local(local)
                    && active_control.is_none()
                    && let Some((_, form_index, _)) = active_forms.last()
                {
                    controls.push(Span {
                        start,
                        end,
                        close_start: None,
                        qname: qname(element.name().as_ref())?,
                        parent: Parent::Form(*form_index),
                    });
                }
            },
            Event::End(element) => {
                depth = depth
                    .checked_sub(1)
                    .ok_or_else(|| Error::InvalidFormat("form XML depth underflow".to_string()))?;
                let local_name = element.local_name();
                let local = local_name.as_ref();
                if active_control
                    .as_ref()
                    .is_some_and(|item| item.depth == depth)
                    && ns == NamespaceKind::Form
                    && control_local(local)
                {
                    let item = active_control.take().expect("active form control");
                    controls.push(Span {
                        start: item.start,
                        end,
                        close_start: Some(start),
                        qname: item.qname,
                        parent: item.parent,
                    });
                } else if active_forms.last().is_some_and(|item| item.0 == depth)
                    && ns == NamespaceKind::Form
                    && local == b"form"
                {
                    let (_, _, item) = active_forms.pop().expect("active form");
                    forms.push(Span {
                        start: item.start,
                        end,
                        close_start: Some(start),
                        qname: item.qname,
                        parent: item.parent,
                    });
                } else if active_groups.last().is_some_and(|item| item.0 == depth)
                    && ns == NamespaceKind::Office
                    && local == b"forms"
                {
                    let (_, group_start, group_index, name) =
                        active_groups.pop().expect("active forms");
                    groups.push(Span {
                        start: group_start,
                        end,
                        close_start: Some(start),
                        qname: name,
                        parent: Parent::Group(group_index),
                    });
                }
            },
            Event::DocType(_) => return invalid("DOCTYPE is not allowed in form host XML"),
            Event::Eof => break,
            _ => {},
        }
        buffer.clear();
    }
    if !active_groups.is_empty() || !active_forms.is_empty() || active_control.is_some() {
        return invalid("unterminated form structure");
    }
    groups.sort_by_key(|span| span.start);
    forms.sort_by_key(|span| span.start);
    controls.sort_by_key(|span| span.start);
    Ok(Scan {
        groups,
        forms,
        controls,
    })
}

fn insert_child(xml: &str, site: &Span, fragment: &str) -> Result<String> {
    if let Some(close) = site.close_start {
        splice(xml, close, close, fragment)
    } else {
        let start_tag = &xml[site.start..site.end];
        let slash = start_tag
            .rfind("/>")
            .ok_or_else(|| Error::InvalidFormat("invalid empty form element".to_string()))?;
        let replacement = format!("{}>{}</{}>", &start_tag[..slash], fragment, site.qname);
        splice(xml, site.start, site.end, &replacement)
    }
}

fn insert_group(xml: &str, host: FormHost, group: &str) -> Result<String> {
    let local = match host {
        FormHost::Text => b"text".as_slice(),
        FormHost::Spreadsheet => b"spreadsheet".as_slice(),
        FormHost::Presentation => b"presentation".as_slice(),
    };
    let mut reader = NsReader::from_str(xml);
    let mut buffer = Vec::new();
    let mut depth = 0usize;
    let mut target_depth = None;
    let mut matches = 0usize;
    loop {
        let start = position(&reader)?;
        let (resolved, event) = reader
            .read_resolved_event_into(&mut buffer)
            .map_err(|error| Error::InvalidFormat(format!("invalid form host XML: {error}")))?;
        let office = namespace(&resolved) == NamespaceKind::Office;
        match event {
            Event::Start(element) => {
                if office && element.local_name().as_ref() == local {
                    matches += 1;
                    target_depth = Some(depth);
                }
                depth = depth
                    .checked_add(1)
                    .ok_or_else(|| Error::InvalidFormat("form host depth overflow".to_string()))?;
            },
            Event::End(element) => {
                depth = depth
                    .checked_sub(1)
                    .ok_or_else(|| Error::InvalidFormat("form host depth underflow".to_string()))?;
                if target_depth == Some(depth) && office && element.local_name().as_ref() == local {
                    if matches != 1 {
                        return invalid("form host is ambiguous");
                    }
                    return splice(xml, start, start, group);
                }
            },
            Event::Empty(element) if office && element.local_name().as_ref() == local => {
                return invalid("form host cannot be empty");
            },
            Event::Eof => break,
            _ => {},
        }
        buffer.clear();
    }
    invalid("form host was not found")
}

fn relocate(xml: &str, first: &Span, second: &Span) -> Result<String> {
    if first.start == second.start {
        return Ok(xml.to_string());
    }
    let mut out = String::with_capacity(xml.len());
    if first.start < second.start {
        out.push_str(&xml[..first.start]);
        out.push_str(&xml[first.end..second.end]);
        out.push_str(&xml[first.start..first.end]);
        out.push_str(&xml[second.end..]);
    } else {
        out.push_str(&xml[..second.start]);
        out.push_str(&xml[first.start..first.end]);
        out.push_str(&xml[second.start..first.start]);
        out.push_str(&xml[first.end..]);
    }
    Ok(out)
}

fn bind_fragment(fragment: String) -> Result<String> {
    let point = fragment
        .find(|ch: char| ch.is_whitespace() || ch == '>' || ch == '/')
        .ok_or_else(|| Error::InvalidFormat("invalid form control fragment".to_string()))?;
    Ok(format!(
        "{} xmlns:form=\"{FORM}\" xmlns:office=\"{OFFICE}\" xmlns:text=\"{TEXT}\" xmlns:xlink=\"{XLINK}\" xmlns:xforms=\"{XFORMS}\" xmlns:script=\"{SCRIPT}\"{}",
        &fragment[..point],
        &fragment[point..]
    ))
}

fn control_local(local: &[u8]) -> bool {
    !matches!(
        local,
        b"form"
            | b"properties"
            | b"property"
            | b"list-property"
            | b"list-value"
            | b"connection-resource"
            | b"item"
            | b"option"
            | b"column"
    )
}

#[derive(Clone, Copy, PartialEq, Eq)]
enum NamespaceKind {
    Office,
    Form,
    Other,
}
fn namespace(value: &ResolveResult<'_>) -> NamespaceKind {
    match value {
        ResolveResult::Bound(Namespace(value)) if *value == OFFICE_NS => NamespaceKind::Office,
        ResolveResult::Bound(Namespace(value)) if *value == FORM_NS => NamespaceKind::Form,
        _ => NamespaceKind::Other,
    }
}
fn position(reader: &NsReader<&[u8]>) -> Result<usize> {
    usize::try_from(reader.buffer_position())
        .map_err(|_| Error::InvalidFormat("form XML position overflow".to_string()))
}
fn qname(value: &[u8]) -> Result<String> {
    std::str::from_utf8(value)
        .map(str::to_string)
        .map_err(|_| Error::InvalidFormat("invalid form qualified name".to_string()))
}
fn validate_name(label: &str, value: &str) -> Result<()> {
    if value.is_empty() {
        return invalid(format!("{label} cannot be empty"));
    }
    validate_string(value)
}
fn validate_string(value: &str) -> Result<()> {
    if value.len() > MAX_STRING {
        return invalid("form string exceeds size limit");
    }
    if value
        .chars()
        .any(|ch| matches!(ch as u32, 0..=8 | 11 | 12 | 14..=31 | 0xFFFE | 0xFFFF))
    {
        return invalid("form string contains an XML-prohibited character");
    }
    Ok(())
}
fn attr(out: &mut String, name: &str, value: Option<&str>) -> Result<()> {
    let Some(value) = value else {
        return Ok(());
    };
    validate_string(value)?;
    out.push(' ');
    out.push_str(name);
    out.push_str("=\"");
    for ch in value.chars() {
        match ch {
            '&' => out.push_str("&amp;"),
            '<' => out.push_str("&lt;"),
            '"' => out.push_str("&quot;"),
            _ => out.push(ch),
        }
    }
    out.push('"');
    Ok(())
}
fn bool_attr(out: &mut String, name: &str, value: Option<bool>) {
    if let Some(value) = value {
        out.push(' ');
        out.push_str(name);
        out.push_str(if value { "=\"true\"" } else { "=\"false\"" });
    }
}
fn bounds(label: &str, index: usize, len: usize) -> Error {
    Error::InvalidFormat(format!(
        "{label} index {index} is out of bounds for {len} entries"
    ))
}
fn invalid<T>(message: impl Into<String>) -> Result<T> {
    Err(Error::InvalidFormat(message.into()))
}
