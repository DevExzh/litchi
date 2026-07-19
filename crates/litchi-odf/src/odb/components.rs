//! Typed inert database form and report component collections.

use super::document::{
    DatabaseContent, DatabaseDocument, DatabaseElement, parse_database_content,
    validate_database_root,
};
use super::query::parse_database_query_model_xml;
use litchi_core::{Error, Result};
use quick_xml::events::Event;
use quick_xml::name::{Namespace, ResolveResult};
use quick_xml::reader::NsReader;
use std::collections::{BTreeMap, BTreeSet};

const DB: &str = "urn:oasis:names:tc:opendocument:xmlns:database:1.0";
const OFFICE: &str = "urn:oasis:names:tc:opendocument:xmlns:office:1.0";
const XLINK: &str = "http://www.w3.org/1999/xlink";
const MATH: &str = "http://www.w3.org/1998/Math/MathML";
const XML: &str = "http://www.w3.org/XML/1998/namespace";
const SCRIPT: &str = "urn:oasis:names:tc:opendocument:xmlns:script:1.0";
const MAX_VALUE: usize = 1024 * 1024;
const MAX_AGGREGATE: usize = 16 * 1024 * 1024;
const MAX_ITEMS: usize = 4096;
const MAX_DEPTH: usize = 64;
const MAX_PAYLOAD_NODES: usize = 65_536;
const MAX_PAYLOAD_ATTRIBUTES: usize = 262_144;

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct OdfDatabaseInertAttribute {
    pub namespace_uri: Option<String>,
    pub local_name: String,
    pub value: String,
}
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum OdfDatabaseInertContent {
    Text(String),
    Element(OdfDatabaseInertElement),
}
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct OdfDatabaseInertElement {
    pub namespace_uri: Option<String>,
    pub local_name: String,
    pub attributes: Vec<OdfDatabaseInertAttribute>,
    pub content: Vec<OdfDatabaseInertContent>,
}
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum OdfDatabaseComponentPayload {
    OfficeDocument(OdfDatabaseInertElement),
    Math(OdfDatabaseInertElement),
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct OdfDatabaseComponentLink {
    pub href: String,
    pub explicit_simple_type: bool,
    pub show_none: bool,
    pub actuate_on_request: bool,
}
impl OdfDatabaseComponentLink {
    pub fn new(href: impl Into<String>) -> Self {
        Self {
            href: href.into(),
            explicit_simple_type: false,
            show_none: false,
            actuate_on_request: false,
        }
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct OdfDatabaseComponent {
    pub name: String,
    pub title: Option<String>,
    pub description: Option<String>,
    pub link: Option<OdfDatabaseComponentLink>,
    pub as_template: Option<bool>,
    pub payload: Option<OdfDatabaseComponentPayload>,
}
impl OdfDatabaseComponent {
    pub fn new(name: impl Into<String>) -> Self {
        Self {
            name: name.into(),
            title: None,
            description: None,
            link: None,
            as_template: None,
            payload: None,
        }
    }
}
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct OdfDatabaseComponentCollection {
    pub name: String,
    pub title: Option<String>,
    pub description: Option<String>,
    pub items: Vec<OdfDatabaseComponentItem>,
}
impl OdfDatabaseComponentCollection {
    pub fn new(name: impl Into<String>) -> Self {
        Self {
            name: name.into(),
            title: None,
            description: None,
            items: Vec::new(),
        }
    }
}
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum OdfDatabaseComponentItem {
    Component(OdfDatabaseComponent),
    Collection(OdfDatabaseComponentCollection),
}
#[derive(Debug, Clone, PartialEq, Eq, Default)]
pub struct OdfDatabaseForms {
    pub items: Vec<OdfDatabaseComponentItem>,
}
#[derive(Debug, Clone, PartialEq, Eq, Default)]
pub struct OdfDatabaseReports {
    pub items: Vec<OdfDatabaseComponentItem>,
}
#[derive(Debug, Clone, PartialEq, Eq, Default)]
pub struct OdfDatabaseComponentModel {
    pub forms: Option<OdfDatabaseForms>,
    pub reports: Option<OdfDatabaseReports>,
}

impl OdfDatabaseForms {
    pub fn to_xml_fragment(&self) -> Result<String> {
        write_container("forms", &self.items)
    }
}
impl OdfDatabaseReports {
    pub fn to_xml_fragment(&self) -> Result<String> {
        write_container("reports", &self.items)
    }
}
impl DatabaseDocument {
    pub fn database_forms(&self) -> Result<Option<OdfDatabaseForms>> {
        Ok(model_from_root(self.database())?.forms)
    }
    pub fn database_reports(&self) -> Result<Option<OdfDatabaseReports>> {
        Ok(model_from_root(self.database())?.reports)
    }
    pub fn database_components(&self) -> Result<OdfDatabaseComponentModel> {
        model_from_root(self.database())
    }
}

pub fn parse_database_components_xml(xml: &str) -> Result<OdfDatabaseComponentModel> {
    parse_database_query_model_xml(xml)?;
    let root = parse_database_content(xml)?;
    validate_database_root(&root)?;
    model_from_root(&root)
}
pub fn parse_database_forms_xml(xml: &str) -> Result<Option<OdfDatabaseForms>> {
    Ok(parse_database_components_xml(xml)?.forms)
}
pub fn parse_database_reports_xml(xml: &str) -> Result<Option<OdfDatabaseReports>> {
    Ok(parse_database_components_xml(xml)?.reports)
}
pub fn set_database_forms_xml(xml: &str, value: Option<&OdfDatabaseForms>) -> Result<String> {
    mutate(
        xml,
        Target::Forms,
        value.map(OdfDatabaseForms::to_xml_fragment).transpose()?,
    )
}
pub fn set_database_reports_xml(xml: &str, value: Option<&OdfDatabaseReports>) -> Result<String> {
    mutate(
        xml,
        Target::Reports,
        value.map(OdfDatabaseReports::to_xml_fragment).transpose()?,
    )
}

#[derive(Default)]
struct Budget {
    aggregate: usize,
    items: usize,
    nodes: usize,
    attributes: usize,
}
impl Budget {
    fn string(&mut self, label: &str, v: &str) -> Result<String> {
        if v.len() > MAX_VALUE {
            return invalid(format!("database component {label} is too large"));
        }
        self.aggregate = self
            .aggregate
            .checked_add(v.len())
            .ok_or_else(|| Error::InvalidFormat("component string budget overflow".into()))?;
        if self.aggregate > MAX_AGGREGATE {
            return invalid("database component strings are too large");
        }
        Ok(v.into())
    }
    fn item(&mut self) -> Result<()> {
        self.items += 1;
        if self.items > MAX_ITEMS {
            invalid("too many database components or collections")
        } else {
            Ok(())
        }
    }
    fn node(&mut self) -> Result<()> {
        self.nodes += 1;
        if self.nodes > MAX_PAYLOAD_NODES {
            invalid("too many component payload nodes")
        } else {
            Ok(())
        }
    }
    fn attribute(&mut self) -> Result<()> {
        self.attributes += 1;
        if self.attributes > MAX_PAYLOAD_ATTRIBUTES {
            invalid("too many component payload attributes")
        } else {
            Ok(())
        }
    }
}

fn model_from_root(root: &DatabaseElement) -> Result<OdfDatabaseComponentModel> {
    let mut b = Budget::default();
    let forms = root
        .children()
        .find(|v| v.namespace_uri() == Some(DB) && v.local_name() == "forms")
        .map(|v| parse_container(v, "forms", &mut b).map(|items| OdfDatabaseForms { items }))
        .transpose()?;
    let reports = root
        .children()
        .find(|v| v.namespace_uri() == Some(DB) && v.local_name() == "reports")
        .map(|v| parse_container(v, "reports", &mut b).map(|items| OdfDatabaseReports { items }))
        .transpose()?;
    Ok(OdfDatabaseComponentModel { forms, reports })
}
fn parse_container(
    e: &DatabaseElement,
    name: &str,
    b: &mut Budget,
) -> Result<Vec<OdfDatabaseComponentItem>> {
    expect(e, name)?;
    attrs(e, &[])?;
    children(e)?
        .into_iter()
        .map(|v| parse_item(v, 1, b))
        .collect()
}
fn parse_item(
    e: &DatabaseElement,
    depth: usize,
    b: &mut Budget,
) -> Result<OdfDatabaseComponentItem> {
    if depth > MAX_DEPTH {
        return invalid("component collections are too deeply nested");
    }
    b.item()?;
    match e.local_name() {
        "component-collection" => {
            expect(e, "component-collection")?;
            attrs(e, &[(DB, "name"), (DB, "title"), (DB, "description")])?;
            let mut items = Vec::new();
            for v in children(e)? {
                items.push(parse_item(v, depth + 1, b)?);
            }
            Ok(OdfDatabaseComponentItem::Collection(
                OdfDatabaseComponentCollection {
                    name: req(e, DB, "name", b)?,
                    title: opt(e, DB, "title", b)?,
                    description: opt(e, DB, "description", b)?,
                    items,
                },
            ))
        },
        "component" => Ok(OdfDatabaseComponentItem::Component(parse_component(e, b)?)),
        n => invalid(format!("unexpected db:{n} in component collection")),
    }
}
fn parse_component(e: &DatabaseElement, b: &mut Budget) -> Result<OdfDatabaseComponent> {
    expect(e, "component")?;
    attrs(
        e,
        &[
            (DB, "name"),
            (DB, "title"),
            (DB, "description"),
            (DB, "as-template"),
            (XLINK, "type"),
            (XLINK, "href"),
            (XLINK, "show"),
            (XLINK, "actuate"),
        ],
    )?;
    let href = e.attribute(Some(XLINK), "href");
    let explicit_type = e.attribute(Some(XLINK), "type");
    let show = e.attribute(Some(XLINK), "show");
    let actuate = e.attribute(Some(XLINK), "actuate");
    if explicit_type.is_some_and(|v| v != "simple")
        || show.is_some_and(|v| v != "none")
        || actuate.is_some_and(|v| v != "onRequest")
    {
        return invalid("invalid fixed xlink component attribute");
    }
    if href.is_none() && (explicit_type.is_some() || show.is_some() || actuate.is_some()) {
        return invalid("component xlink attributes require xlink:href");
    }
    let link = href
        .map(|v| {
            b.string("href", v).map(|href| OdfDatabaseComponentLink {
                href,
                explicit_simple_type: explicit_type.is_some(),
                show_none: show.is_some(),
                actuate_on_request: actuate.is_some(),
            })
        })
        .transpose()?;
    let kids = children(e)?;
    if kids.len() > 1 {
        return invalid("db:component has multiple inline payloads");
    }
    let payload = kids.first().map(|v| parse_payload(v, 1, b)).transpose()?;
    Ok(OdfDatabaseComponent {
        name: req(e, DB, "name", b)?,
        title: opt(e, DB, "title", b)?,
        description: opt(e, DB, "description", b)?,
        link,
        as_template: e
            .attribute(Some(DB), "as-template")
            .map(strict_bool)
            .transpose()?,
        payload,
    })
}
fn parse_payload(
    e: &DatabaseElement,
    depth: usize,
    b: &mut Budget,
) -> Result<OdfDatabaseComponentPayload> {
    let root = parse_inert(e, depth, b)?;
    match (e.namespace_uri(), e.local_name()) {
        (Some(OFFICE), "document") => Ok(OdfDatabaseComponentPayload::OfficeDocument(root)),
        (Some(MATH), "math") => Ok(OdfDatabaseComponentPayload::Math(root)),
        _ => invalid("component payload must be office:document or math:math"),
    }
}
fn parse_inert(
    e: &DatabaseElement,
    depth: usize,
    b: &mut Budget,
) -> Result<OdfDatabaseInertElement> {
    if depth > MAX_DEPTH {
        return invalid("component payload is too deeply nested");
    }
    b.node()?;
    if e.namespace_uri() == Some(SCRIPT)
        || (e.namespace_uri() == Some(OFFICE)
            && matches!(e.local_name(), "scripts" | "event-listeners"))
    {
        return invalid("active content in component payload");
    }
    let namespace_uri = e.namespace_uri().map(str::to_string);
    let local_name = b.string("payload name", e.local_name())?;
    let mut attributes = Vec::new();
    for a in e.attributes() {
        b.attribute()?;
        attributes.push(OdfDatabaseInertAttribute {
            namespace_uri: a.namespace_uri().map(str::to_string),
            local_name: b.string("payload attribute name", a.local_name())?,
            value: b.string("payload attribute", a.value())?,
        });
    }
    let mut content = Vec::new();
    for v in e.content() {
        content.push(match v {
            DatabaseContent::Text(v) => OdfDatabaseInertContent::Text(b.string("payload text", v)?),
            DatabaseContent::Element(v) => {
                OdfDatabaseInertContent::Element(parse_inert(v, depth + 1, b)?)
            },
        });
    }
    Ok(OdfDatabaseInertElement {
        namespace_uri,
        local_name,
        attributes,
        content,
    })
}

fn write_container(name: &str, items: &[OdfDatabaseComponentItem]) -> Result<String> {
    let mut b = Budget::default();
    let mut x = format!(r#"<db:{name} xmlns:db="{DB}" xmlns:xlink="{XLINK}">"#);
    for v in items {
        write_item(&mut x, v, 1, &mut b)?;
    }
    x.push_str("</db:");
    x.push_str(name);
    x.push('>');
    Ok(x)
}
fn write_item(
    x: &mut String,
    v: &OdfDatabaseComponentItem,
    depth: usize,
    b: &mut Budget,
) -> Result<()> {
    if depth > MAX_DEPTH {
        return invalid("component collections are too deeply nested");
    }
    b.item()?;
    match v {
        OdfDatabaseComponentItem::Component(v) => write_component(x, v, b),
        OdfDatabaseComponentItem::Collection(v) => {
            x.push_str("<db:component-collection");
            out(x, "db:name", &v.name, "collection name", b)?;
            opt_out(x, "db:title", v.title.as_deref(), "collection title", b)?;
            opt_out(
                x,
                "db:description",
                v.description.as_deref(),
                "collection description",
                b,
            )?;
            if v.items.is_empty() {
                x.push_str("/>")
            } else {
                x.push('>');
                for c in &v.items {
                    write_item(x, c, depth + 1, b)?;
                }
                x.push_str("</db:component-collection>")
            }
            Ok(())
        },
    }
}
fn write_component(x: &mut String, v: &OdfDatabaseComponent, b: &mut Budget) -> Result<()> {
    x.push_str("<db:component");
    if let Some(link) = &v.link {
        if link.explicit_simple_type {
            lit(x, "xlink:type", "simple")
        }
        out(x, "xlink:href", &link.href, "href", b)?;
        if link.show_none {
            lit(x, "xlink:show", "none")
        }
        if link.actuate_on_request {
            lit(x, "xlink:actuate", "onRequest")
        }
    }
    out(x, "db:name", &v.name, "component name", b)?;
    opt_out(x, "db:title", v.title.as_deref(), "component title", b)?;
    opt_out(
        x,
        "db:description",
        v.description.as_deref(),
        "component description",
        b,
    )?;
    bool_out(x, "db:as-template", v.as_template);
    if let Some(payload) = &v.payload {
        x.push('>');
        write_payload(x, payload, b)?;
        x.push_str("</db:component>")
    } else {
        x.push_str("/>")
    }
    Ok(())
}
fn write_payload(x: &mut String, v: &OdfDatabaseComponentPayload, b: &mut Budget) -> Result<()> {
    let root = match v {
        OdfDatabaseComponentPayload::OfficeDocument(v) => {
            if v.namespace_uri.as_deref() != Some(OFFICE) || v.local_name != "document" {
                return invalid("invalid office document payload root");
            }
            v
        },
        OdfDatabaseComponentPayload::Math(v) => {
            if v.namespace_uri.as_deref() != Some(MATH) || v.local_name != "math" {
                return invalid("invalid MathML payload root");
            }
            v
        },
    };
    let mut namespaces = BTreeSet::new();
    collect_namespaces(root, &mut namespaces);
    namespaces.remove(XML);
    let map: BTreeMap<String, String> = namespaces
        .into_iter()
        .enumerate()
        .map(|(i, v)| (v, format!("n{i}")))
        .collect();
    write_inert(x, root, 1, b, &map, true)
}
fn collect_namespaces(e: &OdfDatabaseInertElement, set: &mut BTreeSet<String>) {
    if let Some(v) = &e.namespace_uri {
        set.insert(v.clone());
    }
    for a in &e.attributes {
        if let Some(v) = &a.namespace_uri {
            set.insert(v.clone());
        }
    }
    for c in &e.content {
        if let OdfDatabaseInertContent::Element(v) = c {
            collect_namespaces(v, set)
        }
    }
}
fn write_inert(
    x: &mut String,
    e: &OdfDatabaseInertElement,
    depth: usize,
    b: &mut Budget,
    map: &BTreeMap<String, String>,
    root: bool,
) -> Result<()> {
    if depth > MAX_DEPTH {
        return invalid("component payload is too deeply nested");
    }
    b.node()?;
    if e.namespace_uri.as_deref() == Some(SCRIPT)
        || (e.namespace_uri.as_deref() == Some(OFFICE)
            && matches!(e.local_name.as_str(), "scripts" | "event-listeners"))
    {
        return invalid("active component payload");
    }
    name(x, e.namespace_uri.as_deref(), &e.local_name, map, true)?;
    if root {
        for (ns, prefix) in map {
            x.push_str(" xmlns:");
            x.push_str(prefix);
            x.push_str("=\"");
            x.push_str(&escape(ns));
            x.push('"');
        }
    }
    let mut attributes = e.attributes.iter().collect::<Vec<_>>();
    attributes
        .sort_by(|a, b| (&a.namespace_uri, &a.local_name).cmp(&(&b.namespace_uri, &b.local_name)));
    for a in attributes {
        b.attribute()?;
        x.push(' ');
        name(x, a.namespace_uri.as_deref(), &a.local_name, map, false)?;
        x.push_str("=\"");
        x.push_str(&escape(&b.string("payload attribute", &a.value)?));
        x.push('"');
    }
    if e.content.is_empty() {
        x.push_str("/>");
        return Ok(());
    }
    x.push('>');
    for c in &e.content {
        match c {
            OdfDatabaseInertContent::Text(v) => x.push_str(&escape(&b.string("payload text", v)?)),
            OdfDatabaseInertContent::Element(v) => write_inert(x, v, depth + 1, b, map, false)?,
        }
    }
    x.push_str("</");
    name(x, e.namespace_uri.as_deref(), &e.local_name, map, false)?;
    x.push('>');
    Ok(())
}
fn name(
    x: &mut String,
    ns: Option<&str>,
    local: &str,
    map: &BTreeMap<String, String>,
    open: bool,
) -> Result<()> {
    if open {
        x.push('<');
    }
    if ns == Some(XML) {
        x.push_str("xml:")
    } else if let Some(ns) = ns {
        let prefix = map
            .get(ns)
            .ok_or_else(|| Error::InvalidFormat("unbound inert namespace".into()))?;
        x.push_str(prefix);
        x.push(':');
    }
    x.push_str(local);
    Ok(())
}

fn children(e: &DatabaseElement) -> Result<Vec<&DatabaseElement>> {
    let mut r = Vec::new();
    for v in e.content() {
        match v {
            DatabaseContent::Element(v) => r.push(v),
            DatabaseContent::Text(v) if v.trim().is_empty() => {},
            DatabaseContent::Text(_) => return invalid(format!("text in db:{}", e.local_name())),
        }
    }
    Ok(r)
}
fn expect(e: &DatabaseElement, n: &str) -> Result<()> {
    if e.namespace_uri() == Some(DB) && e.local_name() == n {
        Ok(())
    } else {
        invalid(format!("expected db:{n}"))
    }
}
fn attrs(e: &DatabaseElement, a: &[(&str, &str)]) -> Result<()> {
    for v in e.attributes() {
        if !a
            .iter()
            .any(|(ns, n)| v.namespace_uri() == Some(*ns) && v.local_name() == *n)
        {
            return invalid(format!("unexpected component attribute {}", v.local_name()));
        }
    }
    Ok(())
}
fn req(e: &DatabaseElement, ns: &str, n: &str, b: &mut Budget) -> Result<String> {
    let v = e
        .attribute(Some(ns), n)
        .ok_or_else(|| Error::InvalidFormat(format!("missing component {n}")))?;
    b.string(n, v)
}
fn opt(e: &DatabaseElement, ns: &str, n: &str, b: &mut Budget) -> Result<Option<String>> {
    e.attribute(Some(ns), n).map(|v| b.string(n, v)).transpose()
}
fn strict_bool(v: &str) -> Result<bool> {
    match v {
        "true" => Ok(true),
        "false" => Ok(false),
        _ => invalid("invalid strict component boolean"),
    }
}
fn out(x: &mut String, n: &str, v: &str, label: &str, b: &mut Budget) -> Result<()> {
    b.string(label, v)?;
    lit(x, n, v);
    Ok(())
}
fn opt_out(x: &mut String, n: &str, v: Option<&str>, label: &str, b: &mut Budget) -> Result<()> {
    if let Some(v) = v {
        out(x, n, v, label, b)?
    }
    Ok(())
}
fn bool_out(x: &mut String, n: &str, v: Option<bool>) {
    if let Some(v) = v {
        lit(x, n, if v { "true" } else { "false" })
    }
}
fn lit(x: &mut String, n: &str, v: &str) {
    x.push(' ');
    x.push_str(n);
    x.push_str("=\"");
    x.push_str(&escape(v));
    x.push('"')
}
fn escape(v: &str) -> String {
    v.replace('&', "&amp;")
        .replace('<', "&lt;")
        .replace('"', "&quot;")
        .replace('\'', "&apos;")
}

#[derive(Clone, Copy)]
enum Target {
    Forms,
    Reports,
}
#[derive(Default)]
struct Spans {
    close: usize,
    forms: Option<(usize, usize)>,
    reports: Option<(usize, usize)>,
    queries: Option<(usize, usize)>,
    tables: Option<(usize, usize)>,
    schema: Option<(usize, usize)>,
}
fn mutate(xml: &str, target: Target, replacement: Option<String>) -> Result<String> {
    parse_database_components_xml(xml)?;
    let s = locate(xml)?;
    let current = match target {
        Target::Forms => s.forms,
        Target::Reports => s.reports,
    };
    let (start, end) = current.unwrap_or_else(|| {
        let at = match target {
            Target::Forms => s
                .reports
                .or(s.queries)
                .or(s.tables)
                .or(s.schema)
                .map(|v| v.0)
                .unwrap_or(s.close),
            Target::Reports => s
                .queries
                .or(s.tables)
                .or(s.schema)
                .map(|v| v.0)
                .unwrap_or(s.close),
        };
        (at, at)
    });
    let replacement = replacement.unwrap_or_default();
    let mut out = String::with_capacity(xml.len() - (end - start) + replacement.len());
    out.push_str(&xml[..start]);
    out.push_str(&replacement);
    out.push_str(&xml[end..]);
    Ok(out)
}
fn locate(xml: &str) -> Result<Spans> {
    let mut r = NsReader::from_str(xml);
    let mut b = Vec::new();
    let mut stack: Vec<(Option<String>, String, usize)> = Vec::new();
    let mut depth = None;
    let mut s = Spans::default();
    loop {
        let start = r.buffer_position() as usize;
        let (ns, event) = r
            .read_resolved_event_into(&mut b)
            .map_err(|e| Error::InvalidFormat(format!("invalid component XML: {e}")))?;
        let resolved = namespace(&ns)?;
        drop(ns);
        let end = r.buffer_position() as usize;
        match event {
            Event::Start(ref e) => {
                let local = owned(e.local_name().as_ref())?;
                if resolved.as_deref() == Some(OFFICE) && local == "database" {
                    depth = Some(stack.len())
                }
                stack.push((resolved, local, start));
            },
            Event::Empty(ref e) => {
                let local = owned(e.local_name().as_ref())?;
                if depth.is_some_and(|d| stack.len() == d + 1) && resolved.as_deref() == Some(DB) {
                    set_span(&mut s, &local, start, end)
                }
            },
            Event::End(ref e) => {
                let local = owned(e.local_name().as_ref())?;
                if let Some((ns, opened, at)) = stack.pop() {
                    if opened != local {
                        return invalid("mismatched component XML");
                    }
                    if depth.is_some_and(|d| stack.len() == d + 1) && ns.as_deref() == Some(DB) {
                        set_span(&mut s, &opened, at, end)
                    }
                    if depth == Some(stack.len())
                        && ns.as_deref() == Some(OFFICE)
                        && opened == "database"
                    {
                        s.close = start;
                        depth = None
                    }
                }
            },
            Event::Eof => break,
            _ => {},
        }
        b.clear();
    }
    if s.close == 0 {
        invalid("could not locate database close")
    } else {
        Ok(s)
    }
}
fn set_span(s: &mut Spans, n: &str, a: usize, b: usize) {
    let v = Some((a, b));
    match n {
        "forms" => s.forms = v,
        "reports" => s.reports = v,
        "queries" => s.queries = v,
        "table-representations" => s.tables = v,
        "schema-definition" => s.schema = v,
        _ => {},
    }
}
fn namespace(v: &ResolveResult<'_>) -> Result<Option<String>> {
    match v {
        ResolveResult::Unbound => Ok(None),
        ResolveResult::Bound(Namespace(v)) => std::str::from_utf8(v)
            .map(|v| Some(v.into()))
            .map_err(|_| Error::InvalidFormat("non-UTF-8 namespace".into())),
        ResolveResult::Unknown(v) => {
            invalid(format!("unknown prefix {}", String::from_utf8_lossy(v)))
        },
    }
}
fn owned(v: &[u8]) -> Result<String> {
    std::str::from_utf8(v)
        .map(str::to_string)
        .map_err(|_| Error::InvalidFormat("non-UTF-8 name".into()))
}
fn invalid<T>(v: impl Into<String>) -> Result<T> {
    Err(Error::InvalidFormat(v.into()))
}

#[cfg(test)]
mod tests {
    use super::*;
    fn wrap(forms: &str, reports: &str) -> String {
        format!(
            r#"<o:document-content xmlns:o="{OFFICE}" xmlns:d="{DB}" xmlns:x="{XLINK}" xmlns:m="{MATH}"><o:body><o:database><d:data-source><d:connection-data><d:connection-resource x:type="simple" x:href="db"/></d:connection-data></d:data-source><!--keep-->{forms}{reports}<d:queries/><d:table-representations/><d:schema-definition><d:table-definitions/></d:schema-definition></o:database></o:body></o:document-content>"#
        )
    }
    #[test]
    fn nested_components_payloads_and_links_roundtrip() {
        let forms = r#"<d:forms><d:component-collection d:name="Folder" d:title="Forms"><d:component x:type="simple" x:href="Forms/Main" x:show="none" x:actuate="onRequest" d:as-template="false" d:name="Main"><m:math display="block"><m:mi>x</m:mi><m:mo>+</m:mo><m:mn>1</m:mn></m:math></d:component></d:component-collection></d:forms>"#;
        let reports = r#"<d:reports><d:component d:name="Summary"><o:document o:version="1.2"><o:body><o:text/></o:body></o:document></d:component></d:reports>"#;
        let parsed = parse_database_components_xml(&wrap(forms, reports)).unwrap();
        let f = parsed.forms.as_ref().unwrap().to_xml_fragment().unwrap();
        let r = parsed.reports.as_ref().unwrap().to_xml_fragment().unwrap();
        assert!(f.contains("xmlns:n0=") && f.contains("xlink:href=\"Forms/Main\""));
        assert_eq!(
            parse_database_components_xml(&wrap(&f, &r)).unwrap(),
            parsed
        );
    }
    #[test]
    fn rejects_fixed_links_wrong_children_active_content_and_bounds() {
        let bad = [
            r#"<d:forms><d:component d:name="x" x:type="extended" x:href="x"/></d:forms>"#,
            r#"<d:forms><d:component d:name="x" x:show="none"/></d:forms>"#,
            r#"<d:forms><d:component d:name="x"><m:math/><m:math/></d:component></d:forms>"#,
            r#"<d:forms><d:component d:name="x"><o:document><o:scripts/></o:document></d:component></d:forms>"#,
            r#"<d:forms><d:bad/></d:forms>"#,
            r#"<d:forms d:name="x"/>"#,
        ];
        for v in bad {
            assert!(
                parse_database_components_xml(&wrap(v, "")).is_err(),
                "accepted {v}"
            );
        }
        let huge = "x".repeat(MAX_VALUE + 1);
        assert!(
            parse_database_components_xml(&wrap(
                &format!(r#"<d:forms><d:component d:name="{huge}"/></d:forms>"#),
                ""
            ))
            .is_err()
        );
    }
    #[test]
    fn setters_preserve_bytes_and_direct_order() {
        let original = wrap("", r#"<d:reports><d:component d:name="old"/></d:reports>"#);
        let forms = OdfDatabaseForms {
            items: vec![OdfDatabaseComponentItem::Component(
                OdfDatabaseComponent::new("new"),
            )],
        };
        let inserted = set_database_forms_xml(&original, Some(&forms)).unwrap();
        assert!(
            inserted.contains("<!--keep--><db:forms")
                && inserted.find("<db:forms").unwrap() < inserted.find("<d:reports").unwrap()
        );
        let reports = OdfDatabaseReports::default();
        let replaced = set_database_reports_xml(&inserted, Some(&reports)).unwrap();
        assert!(!replaced.contains("d:name=\"old\"") && replaced.contains("<db:reports"));
        let removed = set_database_forms_xml(&replaced, None).unwrap();
        assert!(removed.contains("<!--keep--><db:reports") && !removed.contains("<db:forms"));
    }
}
