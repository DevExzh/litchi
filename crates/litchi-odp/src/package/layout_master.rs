//! Lossless custom slide-layout and slide-master editing for packaged ODP files.

use super::Presentation;
use crate::core::PackageWriter;
use crate::model::page_layout::Layout;
use crate::model::page_layout::{
    parse as parse_page_layouts, remove_xml as remove_page_layout_xml,
    set_xml as set_page_layout_xml,
};
use litchi_core::{Error, Result};
use litchi_odf_common::style::master::Master as SharedMasterPage;
use litchi_odf_common::style::master::reader::read as parse_master_pages;
use quick_xml::events::{BytesStart, Event};
use quick_xml::name::{Namespace, ResolveResult};
use quick_xml::{NsReader, XmlVersion};
use std::collections::{HashMap, HashSet};
use std::str;

const OFFICE: &str = "urn:oasis:names:tc:opendocument:xmlns:office:1.0";
const STYLE: &str = "urn:oasis:names:tc:opendocument:xmlns:style:1.0";
const DRAW: &str = "urn:oasis:names:tc:opendocument:xmlns:drawing:1.0";
const PRESENTATION: &str = "urn:oasis:names:tc:opendocument:xmlns:presentation:1.0";
const SCRIPT: &str = "urn:oasis:names:tc:opendocument:xmlns:script:1.0";
const MAX_XML: usize = 64 * 1024 * 1024;
const MAX_ELEMENTS: usize = 262_144;
const MAX_DEFINITIONS: usize = 65_536;
const MAX_NAME: usize = 4_096;

/// A presentation master backed by the shared lossless [`MasterPage`] model.
///
/// Presentation-specific layout and declaration references are typed here;
/// shapes, backgrounds, regions, notes, animations, and extensions remain in
/// `master_page.xml` and are preserved when the definition is edited.
#[derive(Clone, Debug)]
pub struct MasterPage {
    pub master_page: SharedMasterPage,
    pub page_layout_name: Option<String>,
    pub header_name: Option<String>,
    pub footer_name: Option<String>,
    pub date_time_name: Option<String>,
}

impl MasterPage {
    pub fn new(name: impl Into<String>, page_layout_name: impl Into<String>) -> Result<Self> {
        let name = name.into();
        let page_layout_name = page_layout_name.into();
        validate_name(&name, "master page name")?;
        validate_name(&page_layout_name, "page layout name")?;
        Ok(Self {
            master_page: SharedMasterPage {
                name: name.clone(),
                display_name: None,
                page_layout_name: Some(page_layout_name),
                drawing_style_name: None,
                next_style_name: None,
                regions: Vec::new(),
                children: Vec::new(),
                xml: format!(
                    "<style:master-page xmlns:style=\"{STYLE}\" style:name=\"{}\"/>",
                    escape_attr(&name)
                ),
            },
            page_layout_name: None,
            header_name: None,
            footer_name: None,
            date_time_name: None,
        })
    }

    #[must_use]
    pub fn name(&self) -> &str {
        &self.master_page.name
    }

    pub fn validate(&self) -> Result<()> {
        self.fragment().map(|_| ())
    }

    fn validate_fields(&self) -> Result<()> {
        validate_name(&self.master_page.name, "master page name")?;
        for (value, context) in [
            (
                self.master_page.page_layout_name.as_deref(),
                "page layout name",
            ),
            (
                self.master_page.drawing_style_name.as_deref(),
                "drawing-page style name",
            ),
            (
                self.master_page.next_style_name.as_deref(),
                "next master name",
            ),
            (
                self.page_layout_name.as_deref(),
                "presentation page layout name",
            ),
            (self.header_name.as_deref(), "header declaration name"),
            (self.footer_name.as_deref(), "footer declaration name"),
            (self.date_time_name.as_deref(), "date-time declaration name"),
        ] {
            if let Some(value) = value {
                validate_name(value, context)?;
            }
        }
        Ok(())
    }

    fn fragment(&self) -> Result<String> {
        self.validate_fields()?;
        let elements = scan(&self.master_page.xml)?;
        reject_active_content(&elements)?;
        let masters = select(&elements, STYLE, "master-page");
        if masters.len() != 1 {
            return invalid("master XML must contain exactly one style:master-page");
        }
        let changes = [
            Change::new(STYLE, "name", Some(&self.master_page.name), "style"),
            Change::new(
                STYLE,
                "display-name",
                self.master_page.display_name.as_deref(),
                "style",
            ),
            Change::new(
                STYLE,
                "page-layout-name",
                self.master_page.page_layout_name.as_deref(),
                "style",
            ),
            Change::new(
                DRAW,
                "style-name",
                self.master_page.drawing_style_name.as_deref(),
                "draw",
            ),
            Change::new(
                STYLE,
                "next-style-name",
                self.master_page.next_style_name.as_deref(),
                "style",
            ),
            Change::new(
                PRESENTATION,
                "presentation-page-layout-name",
                self.page_layout_name.as_deref(),
                "presentation",
            ),
            Change::new(
                PRESENTATION,
                "use-header-name",
                self.header_name.as_deref(),
                "presentation",
            ),
            Change::new(
                PRESENTATION,
                "use-footer-name",
                self.footer_name.as_deref(),
                "presentation",
            ),
            Change::new(
                PRESENTATION,
                "use-date-time-name",
                self.date_time_name.as_deref(),
                "presentation",
            ),
        ];
        let xml = replace_start(&self.master_page.xml, masters[0], &changes)?;
        let validation_xml = format!(
            "<office:document-styles xmlns:office=\"{OFFICE}\"><office:master-styles>{xml}</office:master-styles></office:document-styles>"
        );
        let parsed = parse_master_pages(&validation_xml)?;
        if parsed.len() != 1 || parsed[0].name != self.master_page.name {
            return invalid("serialized master page did not roundtrip");
        }
        Ok(xml)
    }
}

impl Presentation {
    pub fn master_pages(&self) -> Result<Vec<MasterPage>> {
        masters_from_xml(required_styles(self)?)
    }

    pub fn add_layout(&mut self, layout: &Layout) -> Result<()> {
        layout.validate()?;
        let styles = required_styles(self)?.to_string();
        if definitions(&styles, STYLE, "presentation-page-layout", STYLE, "name")?
            .contains_key(&layout.name)
        {
            return invalid(format!(
                "presentation page layout '{}' already exists",
                layout.name
            ));
        }
        let styles = set_page_layout_xml(&styles, layout)?;
        self.commit_design(styles, self.content_xml().to_string())
    }

    pub fn replace_page_layout(&mut self, layout: &Layout) -> Result<()> {
        layout.validate()?;
        let styles = required_styles(self)?.to_string();
        if !definitions(&styles, STYLE, "presentation-page-layout", STYLE, "name")?
            .contains_key(&layout.name)
        {
            return invalid(format!(
                "presentation page layout '{}' does not exist",
                layout.name
            ));
        }
        let styles = set_page_layout_xml(&styles, layout)?;
        self.commit_design(styles, self.content_xml().to_string())
    }

    pub fn remove_page_layout(&mut self, name: &str, replacement: Option<&str>) -> Result<()> {
        validate_name(name, "presentation page layout name")?;
        if replacement == Some(name) {
            return invalid("layout replacement must differ from removed layout");
        }
        let mut styles = required_styles(self)?.to_string();
        let layouts = definitions(&styles, STYLE, "presentation-page-layout", STYLE, "name")?;
        if !layouts.contains_key(name) {
            return invalid(format!("presentation page layout '{name}' does not exist"));
        }
        if let Some(target) = replacement {
            validate_name(target, "replacement page layout name")?;
            if !layouts.contains_key(target) {
                return invalid(format!("replacement page layout '{target}' does not exist"));
            }
        }
        let content = rewrite_attr(
            self.content_xml(),
            None,
            PRESENTATION,
            "presentation-page-layout-name",
            name,
            replacement,
        )?;
        styles = rewrite_attr(
            &styles,
            None,
            PRESENTATION,
            "presentation-page-layout-name",
            name,
            replacement,
        )?;
        styles = remove_page_layout_xml(&styles, name)?;
        self.commit_design(styles, content)
    }

    pub fn reorder_layouts(&mut self, names: &[String]) -> Result<()> {
        let styles = reorder(
            required_styles(self)?,
            STYLE,
            "presentation-page-layout",
            STYLE,
            "name",
            names,
        )?;
        self.commit_design(styles, self.content_xml().to_string())
    }

    pub fn add_master_page(&mut self, master: &MasterPage) -> Result<()> {
        let fragment = master.fragment()?;
        let styles = required_styles(self)?.to_string();
        if definitions(&styles, STYLE, "master-page", STYLE, "name")?.contains_key(master.name()) {
            return invalid(format!("master page '{}' already exists", master.name()));
        }
        let styles = insert_child(&styles, OFFICE, "master-styles", &fragment)?;
        self.commit_design(styles, self.content_xml().to_string())
    }

    pub fn replace_master_page(&mut self, master: &MasterPage) -> Result<()> {
        let fragment = master.fragment()?;
        let styles = required_styles(self)?.to_string();
        let masters = definitions(&styles, STYLE, "master-page", STYLE, "name")?;
        let span = masters
            .get(master.name())
            .ok_or_else(|| error(format!("master page '{}' does not exist", master.name())))?;
        let styles = replace_range(&styles, span.start, span.end, &fragment)?;
        self.commit_design(styles, self.content_xml().to_string())
    }

    pub fn remove_master_page(&mut self, name: &str, replacement: Option<&str>) -> Result<()> {
        validate_name(name, "master page name")?;
        if replacement == Some(name) {
            return invalid("master replacement must differ from removed master");
        }
        let mut styles = required_styles(self)?.to_string();
        let masters = definitions(&styles, STYLE, "master-page", STYLE, "name")?;
        if !masters.contains_key(name) {
            return invalid(format!("master page '{name}' does not exist"));
        }
        if let Some(target) = replacement {
            validate_name(target, "replacement master page name")?;
            if !masters.contains_key(target) {
                return invalid(format!("replacement master page '{target}' does not exist"));
            }
        }
        let content = rewrite_attr(
            self.content_xml(),
            Some((DRAW, "page")),
            DRAW,
            "master-page-name",
            name,
            replacement,
        )?;
        styles = rewrite_attr(
            &styles,
            Some((STYLE, "master-page")),
            STYLE,
            "next-style-name",
            name,
            replacement,
        )?;
        let refreshed = definitions(&styles, STYLE, "master-page", STYLE, "name")?;
        let span = refreshed
            .get(name)
            .ok_or_else(|| error("master disappeared during staging"))?;
        styles = replace_range(&styles, span.start, span.end, "")?;
        self.commit_design(styles, content)
    }

    pub fn reorder_master_pages(&mut self, names: &[String]) -> Result<()> {
        let styles = reorder(
            required_styles(self)?,
            STYLE,
            "master-page",
            STYLE,
            "name",
            names,
        )?;
        self.commit_design(styles, self.content_xml().to_string())
    }

    pub fn assign_slide_master_page(
        &mut self,
        slide_index: usize,
        master_name: Option<&str>,
    ) -> Result<()> {
        if let Some(name) = master_name {
            validate_name(name, "master page name")?;
            if !definitions(required_styles(self)?, STYLE, "master-page", STYLE, "name")?
                .contains_key(name)
            {
                return invalid(format!("master page '{name}' does not exist"));
            }
        }
        let content = rewrite_index(
            self.content_xml(),
            DRAW,
            "page",
            slide_index,
            &[Change::new(DRAW, "master-page-name", master_name, "draw")],
        )?;
        self.commit_design(required_styles(self)?.to_string(), content)
    }

    pub fn assign_slide_page_layout(
        &mut self,
        slide_index: usize,
        layout_name: Option<&str>,
    ) -> Result<()> {
        if let Some(name) = layout_name {
            validate_name(name, "presentation page layout name")?;
            if !definitions(
                required_styles(self)?,
                STYLE,
                "presentation-page-layout",
                STYLE,
                "name",
            )?
            .contains_key(name)
            {
                return invalid(format!("presentation page layout '{name}' does not exist"));
            }
        }
        let content = rewrite_index(
            self.content_xml(),
            DRAW,
            "page",
            slide_index,
            &[Change::new(
                PRESENTATION,
                "presentation-page-layout-name",
                layout_name,
                "presentation",
            )],
        )?;
        self.commit_design(required_styles(self)?.to_string(), content)
    }

    pub(crate) fn commit_design(&mut self, styles: String, content: String) -> Result<()> {
        validate_references(&styles, &content)?;
        let package = self.owned_package().package()?;
        let mut writer = PackageWriter::new();
        writer.set_mimetype(package.mimetype())?;
        writer.add_file("content.xml", content.as_bytes())?;
        writer.add_file("styles.xml", styles.as_bytes())?;
        for path in ["meta.xml", "settings.xml"] {
            if package.has_file(path) {
                writer.add_file(path, &package.get_file(path)?)?;
            }
        }
        writer.copy_auxiliary_files_from(self.owned_package())?;
        let next = Presentation::from_bytes(writer.finish()?)?;
        *self = next;
        Ok(())
    }
}

fn required_styles(presentation: &Presentation) -> Result<&str> {
    presentation
        .styles_xml()
        .ok_or_else(|| error("ODP package has no styles.xml"))
}

#[derive(Clone, Debug)]
struct Attribute {
    qname: String,
    namespace: Option<String>,
    local: String,
    raw: String,
    value: String,
}

#[derive(Clone, Debug)]
struct Span {
    namespace: Option<String>,
    local: String,
    qname: String,
    attributes: Vec<Attribute>,
    start: usize,
    tag_end: usize,
    end: usize,
    empty: bool,
}

struct Change<'a> {
    namespace: &'static str,
    local: &'static str,
    value: Option<&'a str>,
    prefix: &'static str,
}

impl<'a> Change<'a> {
    const fn new(
        namespace: &'static str,
        local: &'static str,
        value: Option<&'a str>,
        prefix: &'static str,
    ) -> Self {
        Self {
            namespace,
            local,
            value,
            prefix,
        }
    }
}

fn scan(xml: &str) -> Result<Vec<Span>> {
    if xml.len() > MAX_XML {
        return invalid("ODF XML part exceeds 64 MiB");
    }
    let source = xml.strip_prefix('\u{feff}').unwrap_or(xml);
    let source_offset = xml.len() - source.len();
    let mut reader = NsReader::from_str(source);
    reader.config_mut().trim_text(false);
    let mut buffer = Vec::new();
    let mut spans = Vec::<Span>::new();
    let mut open = Vec::new();
    loop {
        let (namespace, event) = reader
            .read_resolved_event_into(&mut buffer)
            .map_err(|e| error(format!("invalid ODF XML: {e}")))?;
        match event {
            Event::Start(element) => {
                let namespace = resolve_namespace(&namespace)?;
                let index = push_span(source, &reader, &mut spans, namespace, &element, false)?;
                open.push(index);
            },
            Event::Empty(element) => {
                let namespace = resolve_namespace(&namespace)?;
                push_span(source, &reader, &mut spans, namespace, &element, true)?;
            },
            Event::End(_) => {
                let index = open
                    .pop()
                    .ok_or_else(|| error("ODF XML element stack underflow"))?;
                spans[index].end = usize::try_from(reader.buffer_position())
                    .map_err(|_err| error("XML position overflow"))?;
            },
            Event::Eof => break,
            _ => {},
        }
        buffer.clear();
    }
    if !open.is_empty() {
        return invalid("ODF XML contains unclosed elements");
    }
    if source_offset != 0 {
        for span in &mut spans {
            span.start += source_offset;
            span.tag_end += source_offset;
            span.end += source_offset;
        }
    }
    Ok(spans)
}

fn push_span(
    xml: &str,
    reader: &NsReader<&[u8]>,
    spans: &mut Vec<Span>,
    namespace: Option<String>,
    element: &BytesStart<'_>,
    empty: bool,
) -> Result<usize> {
    if spans.len() >= MAX_ELEMENTS {
        return invalid("ODF XML exceeds 262144 elements");
    }
    let tag_end =
        usize::try_from(reader.buffer_position()).map_err(|_err| error("XML position overflow"))?;
    let start = xml[..tag_end]
        .rfind('<')
        .ok_or_else(|| error("XML element start is missing"))?;
    let index = spans.len();
    spans.push(Span {
        namespace,
        local: decode(element.local_name().as_ref(), "element local name")?,
        qname: decode(element.name().as_ref(), "element qualified name")?,
        attributes: attributes(reader, element)?,
        start,
        tag_end,
        end: tag_end,
        empty,
    });
    Ok(index)
}

fn attributes(reader: &NsReader<&[u8]>, element: &BytesStart<'_>) -> Result<Vec<Attribute>> {
    let mut result = Vec::new();
    let mut expanded = HashSet::new();
    for raw in element.attributes().with_checks(true) {
        let raw = raw.map_err(|e| error(format!("invalid ODF attribute: {e}")))?;
        let qname = decode(raw.key.as_ref(), "attribute qualified name")?;
        let (namespace, local) = reader.resolver().resolve_attribute(raw.key);
        let namespace = resolve_namespace(&namespace)?;
        let local = decode(local.as_ref(), "attribute local name")?;
        if !qname.starts_with("xmlns") && !expanded.insert((namespace.clone(), local.clone())) {
            return invalid(format!("duplicate expanded XML attribute '{qname}'"));
        }
        let value = raw
            .decoded_and_normalized_value(XmlVersion::Explicit1_0, reader.decoder())
            .map_err(|e| error(format!("invalid ODF attribute value: {e}")))?
            .into_owned();
        if value.len() > 65_536 {
            return invalid("ODF attribute exceeds 64 KiB");
        }
        result.push(Attribute {
            qname,
            namespace,
            local,
            raw: decode(raw.value.as_ref(), "attribute value")?,
            value,
        });
    }
    Ok(result)
}

fn resolve_namespace(value: &ResolveResult<'_>) -> Result<Option<String>> {
    match value {
        ResolveResult::Unbound => Ok(None),
        ResolveResult::Bound(Namespace(uri)) => Ok(Some(decode(uri, "namespace URI")?)),
        ResolveResult::Unknown(prefix) => invalid(format!(
            "unbound XML namespace prefix '{}'",
            String::from_utf8_lossy(prefix.as_ref())
        )),
    }
}

fn select<'a>(spans: &'a [Span], namespace: &str, local: &str) -> Vec<&'a Span> {
    spans
        .iter()
        .filter(|span| span.namespace.as_deref() == Some(namespace) && span.local == local)
        .collect()
}

fn attr<'a>(span: &'a Span, namespace: &str, local: &str) -> Option<&'a str> {
    span.attributes
        .iter()
        .find(|a| a.namespace.as_deref() == Some(namespace) && a.local == local)
        .map(|a| a.value.as_str())
}

fn definitions(
    xml: &str,
    element_ns: &str,
    element_local: &str,
    name_ns: &str,
    name_local: &str,
) -> Result<HashMap<String, Span>> {
    let spans = scan(xml)?;
    let selected = select(&spans, element_ns, element_local);
    if selected.len() > MAX_DEFINITIONS {
        return invalid("ODF definition count exceeds 65536");
    }
    let mut result = HashMap::new();
    for span in selected {
        let name = attr(span, name_ns, name_local)
            .ok_or_else(|| error(format!("{element_local} is missing its name")))?;
        validate_name(name, &format!("{element_local} name"))?;
        if result.insert(name.to_string(), span.clone()).is_some() {
            return invalid(format!("duplicate {element_local} name '{name}'"));
        }
    }
    Ok(result)
}

fn render_start(span: &Span, changes: &[Change<'_>]) -> Result<String> {
    let mut attributes = span.attributes.clone();
    let mut old_prefixes = HashMap::new();
    for change in changes {
        if let Some(old) = attributes
            .iter()
            .find(|a| a.namespace.as_deref() == Some(change.namespace) && a.local == change.local)
            && let Some((prefix, _)) = old.qname.split_once(':')
        {
            old_prefixes.insert(change.namespace, prefix.to_string());
        }
        attributes.retain(|a| {
            !(a.namespace.as_deref() == Some(change.namespace) && a.local == change.local)
        });
    }
    for change in changes {
        let Some(value) = change.value else { continue };
        validate_name(value, change.local)?;
        let prefix = if span.namespace.as_deref() == Some(change.namespace) {
            span.qname
                .split_once(':')
                .map(|(p, _)| p.to_string())
                .unwrap_or_else(|| change.prefix.to_string())
        } else if let Some(prefix) = old_prefixes.get(change.namespace) {
            prefix.clone()
        } else {
            available_prefix(&attributes, change.prefix, change.namespace)
        };
        ensure_namespace(&mut attributes, &prefix, change.namespace)?;
        attributes.push(Attribute {
            qname: format!("{prefix}:{}", change.local),
            namespace: Some(change.namespace.to_string()),
            local: change.local.to_string(),
            raw: escape_attr(value),
            value: value.to_string(),
        });
    }
    let mut output = format!("<{}", span.qname);
    for attribute in attributes {
        output.push(' ');
        output.push_str(&attribute.qname);
        output.push_str("=\"");
        output.push_str(&attribute.raw);
        output.push('"');
    }
    output.push_str(if span.empty { "/>" } else { ">" });
    Ok(output)
}

fn replace_start(xml: &str, span: &Span, changes: &[Change<'_>]) -> Result<String> {
    replace_range(xml, span.start, span.tag_end, &render_start(span, changes)?)
}

fn available_prefix(attributes: &[Attribute], preferred: &str, namespace: &str) -> String {
    for candidate in
        std::iter::once(preferred.to_string()).chain((0..1024).map(|i| format!("litchi{i}")))
    {
        match attributes
            .iter()
            .find(|a| a.qname == format!("xmlns:{candidate}"))
        {
            None => return candidate,
            Some(a) if a.value == namespace => return candidate,
            Some(_) => {},
        }
    }
    "litchi-odp".to_string()
}

fn ensure_namespace(attributes: &mut Vec<Attribute>, prefix: &str, namespace: &str) -> Result<()> {
    let qname = format!("xmlns:{prefix}");
    if let Some(old) = attributes.iter().find(|a| a.qname == qname) {
        if old.value != namespace {
            return invalid(format!("namespace prefix '{prefix}' is already bound"));
        }
    } else {
        attributes.push(Attribute {
            qname,
            namespace: None,
            local: prefix.to_string(),
            raw: namespace.to_string(),
            value: namespace.to_string(),
        });
    }
    Ok(())
}

fn insert_child(xml: &str, namespace: &str, local: &str, fragment: &str) -> Result<String> {
    let spans = scan(xml)?;
    let containers = select(&spans, namespace, local);
    if containers.len() != 1 {
        return invalid(format!(
            "ODF XML must contain exactly one {local} container"
        ));
    }
    let container = containers[0];
    if container.empty {
        let raw = &xml[container.start..container.tag_end];
        let slash = raw
            .rfind("/>")
            .ok_or_else(|| error("empty container marker is missing"))?;
        replace_range(
            xml,
            container.start,
            container.tag_end,
            &format!(
                "{}>{fragment}</{}>",
                raw[..slash].trim_end(),
                container.qname
            ),
        )
    } else {
        let mut output = xml.to_string();
        output.insert_str(container.tag_end, fragment);
        Ok(output)
    }
}

fn reorder(
    xml: &str,
    element_ns: &str,
    element_local: &str,
    name_ns: &str,
    name_local: &str,
    names: &[String],
) -> Result<String> {
    let spans = scan(xml)?;
    let slots = select(&spans, element_ns, element_local);
    if slots.len() != names.len() {
        return invalid(format!(
            "{element_local} reorder requires all {} names",
            slots.len()
        ));
    }
    let mut fragments = HashMap::new();
    for slot in &slots {
        let name = attr(slot, name_ns, name_local)
            .ok_or_else(|| error(format!("{element_local} is missing its name")))?;
        if fragments
            .insert(name.to_string(), xml[slot.start..slot.end].to_string())
            .is_some()
        {
            return invalid(format!("duplicate {element_local} name '{name}'"));
        }
    }
    let mut seen = HashSet::new();
    let mut edits = Vec::new();
    for (slot, name) in slots.iter().zip(names) {
        validate_name(name, "reorder name")?;
        if !seen.insert(name) {
            return invalid(format!("duplicate reorder name '{name}'"));
        }
        let fragment = fragments
            .get(name)
            .ok_or_else(|| error(format!("unknown reorder name '{name}'")))?;
        edits.push((slot.start, slot.end, fragment.clone()));
    }
    apply(xml, edits)
}

fn rewrite_index(
    xml: &str,
    namespace: &str,
    local: &str,
    index: usize,
    changes: &[Change<'_>],
) -> Result<String> {
    let spans = scan(xml)?;
    let selected = select(&spans, namespace, local);
    let span = selected
        .get(index)
        .ok_or_else(|| error(format!("slide index {index} is out of bounds")))?;
    replace_start(xml, span, changes)
}

fn rewrite_attr(
    xml: &str,
    filter: Option<(&str, &str)>,
    namespace: &'static str,
    local: &'static str,
    old: &str,
    replacement: Option<&str>,
) -> Result<String> {
    let spans = scan(xml)?;
    let prefix = if namespace == PRESENTATION {
        "presentation"
    } else if namespace == DRAW {
        "draw"
    } else {
        "style"
    };
    let mut edits = Vec::new();
    for span in &spans {
        if filter
            .is_some_and(|(ns, local)| span.namespace.as_deref() != Some(ns) || span.local != local)
        {
            continue;
        }
        if attr(span, namespace, local) == Some(old) {
            edits.push((
                span.start,
                span.tag_end,
                render_start(span, &[Change::new(namespace, local, replacement, prefix)])?,
            ));
        }
    }
    apply(xml, edits)
}

fn masters_from_xml(xml: &str) -> Result<Vec<MasterPage>> {
    let parsed = parse_master_pages(xml)?;
    if parsed.len() > MAX_DEFINITIONS {
        return invalid("ODP styles exceed 65536 master pages");
    }
    let spans = scan(xml)?;
    let roots = select(&spans, STYLE, "master-page");
    if roots.len() != parsed.len() {
        return invalid("parsed master-page count does not match source XML");
    }
    let mut result = Vec::new();
    let mut names = HashSet::new();
    for (master_page, root) in parsed.into_iter().zip(roots) {
        if !names.insert(master_page.name.clone()) {
            return invalid(format!("duplicate master page '{}'", master_page.name));
        }
        if spans
            .iter()
            .any(|span| span.start >= root.start && span.end <= root.end && is_active_content(span))
        {
            return invalid("scripts and event listeners are not allowed in master XML");
        }
        let model = MasterPage {
            master_page,
            page_layout_name: attr(root, PRESENTATION, "presentation-page-layout-name")
                .map(str::to_string),
            header_name: attr(root, PRESENTATION, "use-header-name").map(str::to_string),
            footer_name: attr(root, PRESENTATION, "use-footer-name").map(str::to_string),
            date_time_name: attr(root, PRESENTATION, "use-date-time-name").map(str::to_string),
        };
        model.validate_fields()?;
        result.push(model);
    }
    Ok(result)
}

fn validate_references(styles: &str, content: &str) -> Result<()> {
    parse_page_layouts(styles)?;
    let layouts: HashSet<String> =
        definitions(styles, STYLE, "presentation-page-layout", STYLE, "name")?
            .into_keys()
            .collect();
    let masters = masters_from_xml(styles)?;
    let master_names: HashSet<String> = masters.iter().map(|m| m.name().to_string()).collect();
    let style_spans = scan(styles)?;
    let content_spans = scan(content)?;
    let page_layouts = named_set(
        style_spans.iter().chain(content_spans.iter()),
        STYLE,
        "page-layout",
        STYLE,
        "name",
    )?;
    let drawing_styles = style_names(
        style_spans.iter().chain(content_spans.iter()),
        "drawing-page",
    )?;
    let headers = named_set(
        content_spans.iter(),
        PRESENTATION,
        "header-decl",
        PRESENTATION,
        "name",
    )?;
    let footers = named_set(
        content_spans.iter(),
        PRESENTATION,
        "footer-decl",
        PRESENTATION,
        "name",
    )?;
    let dates = named_set(
        content_spans.iter(),
        PRESENTATION,
        "date-time-decl",
        PRESENTATION,
        "name",
    )?;
    if let Some(handout) = crate::handout_master::codec::read(styles)? {
        require(
            Some(&handout.page_layout_name),
            &page_layouts,
            "handout physical page layout",
        )?;
        require(
            handout.presentation_page_layout_name.as_deref(),
            &layouts,
            "handout presentation page layout",
        )?;
        require(
            handout.drawing_style_name.as_deref(),
            &drawing_styles,
            "handout drawing-page style",
        )?;
        require(
            handout.header_name.as_deref(),
            &headers,
            "handout header declaration",
        )?;
        require(
            handout.footer_name.as_deref(),
            &footers,
            "handout footer declaration",
        )?;
        require(
            handout.date_time_name.as_deref(),
            &dates,
            "handout date-time declaration",
        )?;
    }
    for master in &masters {
        require(
            master.master_page.page_layout_name.as_deref(),
            &page_layouts,
            "master physical page layout",
        )?;
        require(
            master.master_page.drawing_style_name.as_deref(),
            &drawing_styles,
            "master drawing-page style",
        )?;
        require(
            master.master_page.next_style_name.as_deref(),
            &master_names,
            "next master page",
        )?;
        require(
            master.page_layout_name.as_deref(),
            &layouts,
            "master presentation page layout",
        )?;
        require(
            master.header_name.as_deref(),
            &headers,
            "master header declaration",
        )?;
        require(
            master.footer_name.as_deref(),
            &footers,
            "master footer declaration",
        )?;
        require(
            master.date_time_name.as_deref(),
            &dates,
            "master date-time declaration",
        )?;
    }
    let pages = select(&content_spans, DRAW, "page");
    if pages.len() > MAX_DEFINITIONS {
        return invalid("ODP exceeds 65536 slides");
    }
    let mut page_names = HashSet::new();
    for page in pages {
        if let Some(name) = attr(page, DRAW, "name") {
            validate_name(name, "slide name")?;
            if !page_names.insert(name) {
                return invalid(format!("duplicate slide name '{name}'"));
            }
        }
        require(
            attr(page, DRAW, "master-page-name"),
            &master_names,
            "slide master page",
        )?;
        require(
            attr(page, PRESENTATION, "presentation-page-layout-name"),
            &layouts,
            "slide presentation page layout",
        )?;
        require(
            attr(page, DRAW, "style-name"),
            &drawing_styles,
            "slide drawing-page style",
        )?;
    }
    for spans in [&style_spans, &content_spans] {
        global_refs(
            spans,
            PRESENTATION,
            "presentation-page-layout-name",
            &layouts,
            "presentation page layout",
        )?;
        global_refs(
            spans,
            PRESENTATION,
            "use-header-name",
            &headers,
            "header declaration",
        )?;
        global_refs(
            spans,
            PRESENTATION,
            "use-footer-name",
            &footers,
            "footer declaration",
        )?;
        global_refs(
            spans,
            PRESENTATION,
            "use-date-time-name",
            &dates,
            "date-time declaration",
        )?;
    }
    Ok(())
}

fn named_set<'a>(
    spans: impl Iterator<Item = &'a Span>,
    element_ns: &str,
    element_local: &str,
    name_ns: &str,
    name_local: &str,
) -> Result<HashSet<String>> {
    let mut result = HashSet::new();
    for span in
        spans.filter(|s| s.namespace.as_deref() == Some(element_ns) && s.local == element_local)
    {
        let name = attr(span, name_ns, name_local)
            .ok_or_else(|| error(format!("{element_local} is missing its name")))?;
        validate_name(name, "definition name")?;
        if !result.insert(name.to_string()) {
            return invalid(format!("duplicate {element_local} name '{name}'"));
        }
    }
    Ok(result)
}

fn style_names<'a>(spans: impl Iterator<Item = &'a Span>, family: &str) -> Result<HashSet<String>> {
    let mut result = HashSet::new();
    for span in spans.filter(|s| {
        s.namespace.as_deref() == Some(STYLE)
            && s.local == "style"
            && attr(s, STYLE, "family") == Some(family)
    }) {
        let name =
            attr(span, STYLE, "name").ok_or_else(|| error("style:style is missing style:name"))?;
        validate_name(name, "style name")?;
        result.insert(name.to_string());
    }
    Ok(result)
}

fn global_refs(
    spans: &[Span],
    namespace: &str,
    local: &str,
    names: &HashSet<String>,
    context: &str,
) -> Result<()> {
    for span in spans {
        require(attr(span, namespace, local), names, context)?;
    }
    Ok(())
}

fn require(value: Option<&str>, names: &HashSet<String>, context: &str) -> Result<()> {
    let Some(value) = value else { return Ok(()) };
    validate_name(value, context)?;
    if !names.contains(value) {
        return invalid(format!("{context} references missing name '{value}'"));
    }
    Ok(())
}

fn reject_active_content(spans: &[Span]) -> Result<()> {
    if spans.iter().any(is_active_content) {
        return invalid("scripts and event listeners are not allowed in master XML");
    }
    Ok(())
}

fn is_active_content(span: &Span) -> bool {
    span.namespace.as_deref() == Some(SCRIPT)
        || (span.namespace.as_deref() == Some(OFFICE)
            && matches!(span.local.as_str(), "scripts" | "event-listeners"))
}

fn validate_name(value: &str, context: &str) -> Result<()> {
    if value.is_empty() || value.len() > MAX_NAME {
        return invalid(format!("{context} must contain 1..={MAX_NAME} bytes"));
    }
    if value
        .chars()
        .any(|c| matches!(c, '\0'..='\u{8}' | '\u{b}' | '\u{c}' | '\u{e}'..='\u{1f}'))
    {
        return invalid(format!("{context} contains an XML control character"));
    }
    Ok(())
}

fn apply(xml: &str, mut edits: Vec<(usize, usize, String)>) -> Result<String> {
    edits.sort_by_key(|edit| std::cmp::Reverse(edit.0));
    let mut boundary = xml.len();
    let mut output = xml.to_string();
    for (start, end, replacement) in edits {
        if start > end || end > boundary {
            return invalid("overlapping XML replacement spans");
        }
        output.replace_range(start..end, &replacement);
        boundary = start;
    }
    Ok(output)
}

fn replace_range(xml: &str, start: usize, end: usize, replacement: &str) -> Result<String> {
    if start > end || end > xml.len() || !xml.is_char_boundary(start) || !xml.is_char_boundary(end)
    {
        return invalid("invalid UTF-8 XML replacement span");
    }
    let mut output = String::with_capacity(xml.len() - (end - start) + replacement.len());
    output.push_str(&xml[..start]);
    output.push_str(replacement);
    output.push_str(&xml[end..]);
    Ok(output)
}

fn escape_attr(value: &str) -> String {
    let mut output = String::new();
    for c in value.chars() {
        match c {
            '&' => output.push_str("&amp;"),
            '<' => output.push_str("&lt;"),
            '"' => output.push_str("&quot;"),
            '\r' => output.push_str("&#13;"),
            '\n' => output.push_str("&#10;"),
            '\t' => output.push_str("&#9;"),
            _ => output.push(c),
        }
    }
    output
}

fn decode(bytes: &[u8], context: &str) -> Result<String> {
    str::from_utf8(bytes)
        .map(str::to_string)
        .map_err(|_err| error(format!("{context} is not UTF-8")))
}

fn error(message: impl Into<String>) -> Error {
    Error::InvalidFormat(message.into())
}
fn invalid<T>(message: impl Into<String>) -> Result<T> {
    Err(error(message))
}
