//! Bounded `SpreadsheetML` revision XML codecs.

use crate::error::Result;
use litchi_ooxml_common::mce::process_ooxml;
use quick_xml::XmlVersion;
use quick_xml::events::{BytesStart, Event};
use quick_xml::name::{Namespace, ResolveResult};
use quick_xml::reader::NsReader;
use std::collections::HashSet;

use super::model::{
    MAX_DEPTH, MAX_HEADERS, MAX_NODES, MAX_PART_BYTES, MAX_RECORDS_PER_LOG, MAX_SHEET_IDS,
    MAX_STRING_BYTES, MAX_USERS, NS, REL_NS, RevisionAttribute, RevisionAttributeNamespace,
    RevisionConformance, RevisionHeader, RevisionHeaderProperties, RevisionHeaders, RevisionLog,
    RevisionRecord, RevisionRecordKind, RevisionUser, RevisionUsers, RevisionXmlElement, STRICT_NS,
    STRICT_REL_NS, bounded, valid_date, validate_attrs, validate_element, validate_guid,
    validate_headers, validate_log, validate_users,
};
use super::{invalid, limit, xml_error};

#[derive(Clone, Debug)]
struct Node {
    ns: String,
    name: String,
    attrs: Vec<RevisionAttribute>,
    children: Vec<Node>,
    text: String,
}

pub fn parse_revision_users(xml: &[u8]) -> Result<RevisionUsers> {
    let root = parse_document(xml)?;
    let ns = root_ns(&root, "users")?;
    whitespace(&root)?;
    let count = required_u32_attr(&root, "count")? as usize;
    only_attrs(&root, &[(RevisionAttributeNamespace::Unqualified, "count")])?;
    if root.children.len() != count {
        return Err(invalid("users count does not match userInfo children"));
    }
    if count > MAX_USERS {
        return Err(limit("revision users"));
    }
    let mut users = Vec::with_capacity(count);
    let mut ids = HashSet::new();
    let mut guids = HashSet::new();
    for node in &root.children {
        require_name(node, ns, "userInfo")?;
        whitespace(node)?;
        only_attrs(
            node,
            &[
                (RevisionAttributeNamespace::Unqualified, "guid"),
                (RevisionAttributeNamespace::Unqualified, "name"),
                (RevisionAttributeNamespace::Unqualified, "id"),
                (RevisionAttributeNamespace::Unqualified, "dateTime"),
            ],
        )?;
        let user = RevisionUser {
            guid: required_attr(node, &RevisionAttributeNamespace::Unqualified, "guid")?.into(),
            name: required_attr(node, &RevisionAttributeNamespace::Unqualified, "name")?.into(),
            id: required_attr(node, &RevisionAttributeNamespace::Unqualified, "id")?
                .parse()
                .map_err(|_| invalid("invalid revision user id"))?,
            date_time: required_attr(node, &RevisionAttributeNamespace::Unqualified, "dateTime")?
                .into(),
            extension_elements: node.children.iter().map(to_public).collect::<Result<_>>()?,
        };
        validate_guid(&user.guid)?;
        valid_date(&user.date_time)?;
        bounded(&user.name)?;
        if !ids.insert(user.id) || !guids.insert(user.guid.clone()) {
            return Err(invalid("duplicate revision user id or GUID"));
        }
        users.push(user);
    }
    Ok(RevisionUsers { users })
}

pub fn parse_revision_headers(xml: &[u8]) -> Result<RevisionHeaders> {
    let root = parse_document(xml)?;
    let ns = root_ns(&root, "headers")?;
    whitespace(&root)?;
    let properties = RevisionHeaderProperties {
        guid: req(&root, "guid")?.into(),
        last_guid: opt(&root, "lastGuid").map(Into::into),
        shared: opt_bool(&root, "shared")?,
        disk_revisions: opt_bool(&root, "diskRevisions")?,
        history: opt_bool(&root, "history")?,
        track_revisions: opt_bool(&root, "trackRevisions")?,
        exclusive: opt_bool(&root, "exclusive")?,
        keep_change_history: opt_bool(&root, "keepChangeHistory")?,
        protected: opt_bool(&root, "protected")?,
        preserve_history: opt_u32(&root, "preserveHistory")?,
        revision_id: opt_u32(&root, "revisionId")?,
        version: opt(&root, "version")
            .map(|v| v.parse().map_err(|_| invalid("invalid headers version")))
            .transpose()?,
    };
    only_attrs(
        &root,
        &[
            (RevisionAttributeNamespace::Unqualified, "guid"),
            (RevisionAttributeNamespace::Unqualified, "lastGuid"),
            (RevisionAttributeNamespace::Unqualified, "shared"),
            (RevisionAttributeNamespace::Unqualified, "diskRevisions"),
            (RevisionAttributeNamespace::Unqualified, "history"),
            (RevisionAttributeNamespace::Unqualified, "trackRevisions"),
            (RevisionAttributeNamespace::Unqualified, "exclusive"),
            (RevisionAttributeNamespace::Unqualified, "keepChangeHistory"),
            (RevisionAttributeNamespace::Unqualified, "protected"),
            (RevisionAttributeNamespace::Unqualified, "preserveHistory"),
            (RevisionAttributeNamespace::Unqualified, "revisionId"),
            (RevisionAttributeNamespace::Unqualified, "version"),
        ],
    )?;
    validate_guid(&properties.guid)?;
    if let Some(v) = &properties.last_guid {
        validate_guid(v)?;
    }
    if root.children.len() > MAX_HEADERS {
        return Err(limit("revision headers"));
    }
    let mut headers = Vec::with_capacity(root.children.len());
    let mut guids = HashSet::new();
    let mut rels = HashSet::new();
    let mut total_sheets = 0usize;
    for node in &root.children {
        require_name(node, ns, "header")?;
        whitespace(node)?;
        only_attrs(
            node,
            &[
                (RevisionAttributeNamespace::Unqualified, "guid"),
                (RevisionAttributeNamespace::Unqualified, "dateTime"),
                (RevisionAttributeNamespace::Unqualified, "maxSheetId"),
                (RevisionAttributeNamespace::Unqualified, "userName"),
                (RevisionAttributeNamespace::Relationships, "id"),
                (RevisionAttributeNamespace::Unqualified, "minRId"),
                (RevisionAttributeNamespace::Unqualified, "maxRId"),
            ],
        )?;
        let map = node
            .children
            .first()
            .ok_or_else(|| invalid("revision header requires sheetIdMap"))?;
        require_name(map, ns, "sheetIdMap")?;
        whitespace(map)?;
        only_attrs(map, &[(RevisionAttributeNamespace::Unqualified, "count")])?;
        let count = required_u32_attr(map, "count")? as usize;
        if map.children.len() != count {
            return Err(invalid("sheetIdMap count mismatch"));
        }
        total_sheets = total_sheets
            .checked_add(count)
            .ok_or_else(|| limit("sheet ids"))?;
        if total_sheets > MAX_SHEET_IDS {
            return Err(limit("sheet ids"));
        }
        let mut sheet_ids = Vec::with_capacity(count);
        let mut seen = HashSet::new();
        for item in &map.children {
            require_name(item, ns, "sheetId")?;
            require_empty(item)?;
            only_attrs(item, &[(RevisionAttributeNamespace::Unqualified, "val")])?;
            let id = required_u32_attr(item, "val")?;
            if !seen.insert(id) {
                return Err(invalid("duplicate sheet id in revision header"));
            }
            sheet_ids.push(id);
        }
        let header = RevisionHeader {
            guid: req(node, "guid")?.into(),
            date_time: req(node, "dateTime")?.into(),
            max_sheet_id: req(node, "maxSheetId")?
                .parse()
                .map_err(|_| invalid("invalid maxSheetId"))?,
            user_name: req(node, "userName")?.into(),
            relationship_id: required_attr(node, &RevisionAttributeNamespace::Relationships, "id")?
                .into(),
            min_revision_id: opt_u32(node, "minRId")?,
            max_revision_id: opt_u32(node, "maxRId")?,
            sheet_ids,
            trailing_elements: node
                .children
                .iter()
                .skip(1)
                .map(to_public)
                .collect::<Result<_>>()?,
        };
        validate_guid(&header.guid)?;
        valid_date(&header.date_time)?;
        bounded(&header.user_name)?;
        if header.relationship_id.is_empty() {
            return Err(invalid("empty revision-log relationship id"));
        }
        if header.sheet_ids.iter().any(|id| *id >= header.max_sheet_id) {
            return Err(invalid("sheetIdMap value must be below maxSheetId"));
        }
        if header
            .min_revision_id
            .zip(header.max_revision_id)
            .is_some_and(|(a, b)| a > b)
        {
            return Err(invalid("header minRId exceeds maxRId"));
        }
        if !guids.insert(header.guid.clone()) || !rels.insert(header.relationship_id.clone()) {
            return Err(invalid("duplicate revision header GUID or relationship id"));
        }
        headers.push(header);
    }
    if headers.last().is_some_and(|h| h.guid != properties.guid) {
        return Err(invalid("headers guid must match the most recent header"));
    }
    Ok(RevisionHeaders {
        properties,
        headers,
    })
}

pub fn parse_revision_log(xml: &[u8]) -> Result<RevisionLog> {
    let root = parse_document(xml)?;
    let ns = root_ns(&root, "revisions")?;
    whitespace(&root)?;
    only_attrs(&root, &[])?;
    if root.children.len() > MAX_RECORDS_PER_LOG {
        return Err(limit("revision records"));
    }
    let mut records = Vec::with_capacity(root.children.len());
    for node in &root.children {
        if node.ns != ns {
            return Err(invalid("revision record has wrong namespace"));
        }
        let kind = RevisionRecordKind::parse(&node.name)
            .ok_or_else(|| invalid(format!("unknown revision record '{}'", node.name)))?;
        let mut attributes = node.attrs.clone();
        let revision_id = take_u32(
            &mut attributes,
            &RevisionAttributeNamespace::Unqualified,
            "rId",
        )?;
        let sheet_id = take_u32(
            &mut attributes,
            &RevisionAttributeNamespace::Unqualified,
            "sId",
        )?;
        records.push(RevisionRecord {
            kind,
            revision_id,
            sheet_id,
            attributes,
            children: node.children.iter().map(to_public).collect::<Result<_>>()?,
            text: node.text.clone(),
        });
    }
    let log = RevisionLog { records };
    validate_log(&log)?;
    Ok(log)
}

pub fn write_revision_users(
    value: &RevisionUsers,
    conformance: RevisionConformance,
) -> Result<Vec<u8>> {
    validate_users(value)?;
    let mut out = start("users", conformance);
    attr(&mut out, "count", &value.users.len().to_string());
    out.push('>');
    for user in &value.users {
        out.push_str("<userInfo");
        attr(&mut out, "guid", &user.guid);
        attr(&mut out, "name", &user.name);
        attr(&mut out, "id", &user.id.to_string());
        attr(&mut out, "dateTime", &user.date_time);
        if user.extension_elements.is_empty() {
            out.push_str("/>");
        } else {
            out.push('>');
            for e in &user.extension_elements {
                write_element(&mut out, e)?;
            }
            out.push_str("</userInfo>");
        }
    }
    out.push_str("</users>");
    finish(out)
}

pub fn write_revision_headers(
    value: &RevisionHeaders,
    conformance: RevisionConformance,
) -> Result<Vec<u8>> {
    validate_headers(value)?;
    let mut out = start("headers", conformance);
    attr(&mut out, "guid", &value.properties.guid);
    opt_attr(&mut out, "lastGuid", value.properties.last_guid.as_deref());
    bool_attr(&mut out, "shared", value.properties.shared);
    bool_attr(&mut out, "diskRevisions", value.properties.disk_revisions);
    bool_attr(&mut out, "history", value.properties.history);
    bool_attr(&mut out, "trackRevisions", value.properties.track_revisions);
    bool_attr(&mut out, "exclusive", value.properties.exclusive);
    bool_attr(
        &mut out,
        "keepChangeHistory",
        value.properties.keep_change_history,
    );
    bool_attr(&mut out, "protected", value.properties.protected);
    num_attr(
        &mut out,
        "preserveHistory",
        value.properties.preserve_history,
    );
    num_attr(&mut out, "revisionId", value.properties.revision_id);
    if let Some(v) = value.properties.version {
        attr(&mut out, "version", &v.to_string());
    }
    out.push('>');
    for h in &value.headers {
        out.push_str("<header");
        attr(&mut out, "guid", &h.guid);
        attr(&mut out, "dateTime", &h.date_time);
        attr(&mut out, "maxSheetId", &h.max_sheet_id.to_string());
        attr(&mut out, "userName", &h.user_name);
        rel_attr(&mut out, "id", &h.relationship_id);
        num_attr(&mut out, "minRId", h.min_revision_id);
        num_attr(&mut out, "maxRId", h.max_revision_id);
        out.push_str("><sheetIdMap");
        attr(&mut out, "count", &h.sheet_ids.len().to_string());
        out.push('>');
        for id in &h.sheet_ids {
            out.push_str("<sheetId");
            attr(&mut out, "val", &id.to_string());
            out.push_str("/>");
        }
        out.push_str("</sheetIdMap>");
        for e in &h.trailing_elements {
            write_element(&mut out, e)?;
        }
        out.push_str("</header>");
    }
    out.push_str("</headers>");
    finish(out)
}

pub fn write_revision_log(
    value: &RevisionLog,
    conformance: RevisionConformance,
) -> Result<Vec<u8>> {
    validate_log(value)?;
    let mut out = start("revisions", conformance);
    if value.records.is_empty() {
        out.push_str("/>");
        return finish(out);
    }
    out.push('>');
    for record in &value.records {
        out.push('<');
        out.push_str(record.kind.name());
        if let Some(v) = record.revision_id {
            attr(&mut out, "rId", &v.to_string());
        }
        if let Some(v) = record.sheet_id {
            attr(&mut out, "sId", &v.to_string());
        }
        write_attributes(&mut out, &record.attributes)?;
        if record.children.is_empty() && record.text.is_empty() {
            out.push_str("/>");
            continue;
        }
        out.push('>');
        text(&mut out, &record.text);
        for child in &record.children {
            write_element(&mut out, child)?;
        }
        out.push_str("</");
        out.push_str(record.kind.name());
        out.push('>');
    }
    out.push_str("</revisions>");
    finish(out)
}

fn parse_document(xml: &[u8]) -> Result<Node> {
    if xml.len() > MAX_PART_BYTES {
        return Err(limit("revision part bytes"));
    }
    let processed = process_ooxml(xml)?;
    if processed.len() > MAX_PART_BYTES {
        return Err(limit("MCE-expanded revision part bytes"));
    }
    let mut reader = NsReader::from_reader(processed.as_ref());
    reader.config_mut().trim_text(false);
    let (mut stack, mut root, mut nodes, mut strings, mut buffer) =
        (Vec::new(), None, 0usize, 0usize, Vec::new());
    loop {
        let decoder = reader.decoder();
        let event = reader
            .read_event_into(&mut buffer)
            .map_err(xml_error)?
            .into_owned();
        let resolver = reader.resolver().clone();
        let (ns, event) = resolver.resolve_event(event);
        match event {
            Event::Start(e) => {
                if stack.len() >= MAX_DEPTH {
                    return Err(limit("revision XML depth"));
                }
                let n = make_node(&e, &ns, &reader, decoder, &mut strings)?;
                nodes += 1;
                if nodes > MAX_NODES {
                    return Err(limit("revision XML nodes"));
                }
                stack.push(n);
            },
            Event::Empty(e) => {
                let n = make_node(&e, &ns, &reader, decoder, &mut strings)?;
                nodes += 1;
                if nodes > MAX_NODES {
                    return Err(limit("revision XML nodes"));
                }
                attach(n, &mut stack, &mut root)?;
            },
            Event::End(_) => {
                let n = stack
                    .pop()
                    .ok_or_else(|| invalid("unexpected revision XML end"))?;
                attach(n, &mut stack, &mut root)?;
            },
            Event::Text(t) => {
                let d = t.decode().map_err(xml_error)?;
                let d = quick_xml::escape::unescape(&d).map_err(xml_error)?;
                append_text(stack.last_mut(), &d, &mut strings)?;
            },
            Event::GeneralRef(r) => {
                let d = litchi_ooxml_common::xml::decode_xml_reference(&r)?;
                append_text(stack.last_mut(), &d, &mut strings)?;
            },
            Event::CData(_) => return Err(invalid("CDATA is rejected in revision parts")),
            Event::DocType(_) | Event::PI(_) => {
                return Err(invalid("DTD and processing instructions are rejected"));
            },
            Event::Decl(_) | Event::Comment(_) => {},
            Event::Eof => break,
        }
        buffer.clear();
    }
    if !stack.is_empty() {
        return Err(invalid("unterminated revision XML"));
    }
    root.ok_or_else(|| invalid("revision XML has no root"))
}
fn make_node(
    e: &BytesStart<'_>,
    ns: &ResolveResult<'_>,
    reader: &NsReader<&[u8]>,
    decoder: quick_xml::encoding::Decoder,
    strings: &mut usize,
) -> Result<Node> {
    let ns = ns_text(ns)?;
    if !matches!(ns.as_str(), NS | STRICT_NS) {
        return Err(invalid("revision element has unsupported namespace"));
    }
    let name = std::str::from_utf8(e.local_name().as_ref())
        .map_err(xml_error)?
        .into();
    let mut attrs = Vec::new();
    let mut seen = HashSet::new();
    for a in e.attributes().with_checks(true) {
        let a = a.map_err(xml_error)?;
        if a.key.as_ref() == b"xmlns" || a.key.as_ref().starts_with(b"xmlns:") {
            continue;
        }
        let (resolved, local) = reader.resolver().resolve_attribute(a.key);
        let namespace = match resolved {
            ResolveResult::Unbound => RevisionAttributeNamespace::Unqualified,
            ResolveResult::Bound(Namespace(v))
                if v == REL_NS.as_bytes() || v == STRICT_REL_NS.as_bytes() =>
            {
                RevisionAttributeNamespace::Relationships
            },
            ResolveResult::Bound(Namespace(v)) if v == b"http://www.w3.org/XML/1998/namespace" => {
                RevisionAttributeNamespace::Xml
            },
            _ => return Err(invalid("revision attribute has unsupported namespace")),
        };
        let local = std::str::from_utf8(local.as_ref())
            .map_err(xml_error)?
            .to_owned();
        if !seen.insert((namespace.clone(), local.clone())) {
            return Err(invalid("duplicate expanded revision attribute"));
        }
        let value = a
            .decoded_and_normalized_value(XmlVersion::Implicit1_0, decoder)
            .map_err(xml_error)?
            .into_owned();
        *strings = strings
            .checked_add(local.len() + value.len())
            .ok_or_else(|| limit("revision strings"))?;
        if *strings > MAX_STRING_BYTES {
            return Err(limit("revision strings"));
        }
        attrs.push(RevisionAttribute {
            namespace,
            name: local,
            value,
        });
    }
    Ok(Node {
        ns,
        name,
        attrs,
        children: Vec::new(),
        text: String::new(),
    })
}
fn append_text(node: Option<&mut Node>, value: &str, strings: &mut usize) -> Result<()> {
    *strings = strings
        .checked_add(value.len())
        .ok_or_else(|| limit("revision strings"))?;
    if *strings > MAX_STRING_BYTES {
        return Err(limit("revision strings"));
    }
    if let Some(n) = node {
        n.text.push_str(value);
    } else if !value.trim().is_empty() {
        return Err(invalid("text outside revision root"));
    }
    Ok(())
}
fn attach(node: Node, stack: &mut [Node], root: &mut Option<Node>) -> Result<()> {
    if let Some(p) = stack.last_mut() {
        p.children.push(node);
    } else if root.replace(node).is_some() {
        return Err(invalid("multiple revision XML roots"));
    }
    Ok(())
}
fn to_public(n: &Node) -> Result<RevisionXmlElement> {
    if !matches!(n.ns.as_str(), NS | STRICT_NS) {
        return Err(invalid("unsupported payload namespace"));
    }
    let e = RevisionXmlElement {
        name: n.name.clone(),
        attributes: n.attrs.clone(),
        children: n.children.iter().map(to_public).collect::<Result<_>>()?,
        text: n.text.clone(),
    };
    validate_element(&e, 1)?;
    Ok(e)
}

fn start(name: &str, c: RevisionConformance) -> String {
    format!(
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><{name} xmlns=\"{}\" xmlns:r=\"{}\"",
        c.namespace(),
        c.relationship_namespace()
    )
}
fn finish(s: String) -> Result<Vec<u8>> {
    if s.len() > MAX_PART_BYTES {
        Err(limit("serialized revision part bytes"))
    } else {
        Ok(s.into_bytes())
    }
}
fn write_element(out: &mut String, e: &RevisionXmlElement) -> Result<()> {
    validate_element(e, 1)?;
    out.push('<');
    out.push_str(&e.name);
    write_attributes(out, &e.attributes)?;
    if e.children.is_empty() && e.text.is_empty() {
        out.push_str("/>");
        return Ok(());
    }
    out.push('>');
    text(out, &e.text);
    for c in &e.children {
        write_element(out, c)?;
    }
    out.push_str("</");
    out.push_str(&e.name);
    out.push('>');
    Ok(())
}
fn write_attributes(out: &mut String, attrs: &[RevisionAttribute]) -> Result<()> {
    validate_attrs(attrs)?;
    let mut attrs = attrs.to_vec();
    attrs.sort_by(|a, b| {
        a.namespace
            .cmp(&b.namespace)
            .then(a.name.cmp(&b.name))
            .then(a.value.cmp(&b.value))
    });
    for a in attrs {
        match a.namespace {
            RevisionAttributeNamespace::Unqualified => attr(out, &a.name, &a.value),
            RevisionAttributeNamespace::Relationships => rel_attr(out, &a.name, &a.value),
            RevisionAttributeNamespace::Xml => {
                let n = format!("xml:{}", a.name);
                attr(out, &n, &a.value);
            },
        }
    }
    Ok(())
}
fn attr(o: &mut String, n: &str, v: &str) {
    o.push(' ');
    o.push_str(n);
    o.push_str("=\"");
    escape(o, v, true);
    o.push('"');
}
fn rel_attr(o: &mut String, n: &str, v: &str) {
    let n = format!("r:{n}");
    attr(o, &n, v);
}
fn opt_attr(o: &mut String, n: &str, v: Option<&str>) {
    if let Some(v) = v {
        attr(o, n, v);
    }
}
fn bool_attr(o: &mut String, n: &str, v: Option<bool>) {
    if let Some(v) = v {
        attr(o, n, if v { "1" } else { "0" });
    }
}
fn num_attr(o: &mut String, n: &str, v: Option<u32>) {
    if let Some(v) = v {
        attr(o, n, &v.to_string());
    }
}
fn text(o: &mut String, v: &str) {
    escape(o, v, false);
}
fn escape(o: &mut String, v: &str, attribute: bool) {
    for c in v.chars() {
        match c {
            '&' => o.push_str("&amp;"),
            '<' => o.push_str("&lt;"),
            '>' => o.push_str("&gt;"),
            '"' if attribute => o.push_str("&quot;"),
            '\r' if attribute => o.push_str("&#xD;"),
            '\n' if attribute => o.push_str("&#xA;"),
            '\t' if attribute => o.push_str("&#x9;"),
            _ => o.push(c),
        }
    }
}

fn root_ns<'a>(n: &'a Node, name: &str) -> Result<&'a str> {
    if n.name == name && matches!(n.ns.as_str(), NS | STRICT_NS) {
        Ok(&n.ns)
    } else {
        Err(invalid(format!("expected SpreadsheetML {name} root")))
    }
}
fn require_name(n: &Node, ns: &str, name: &str) -> Result<()> {
    if n.ns == ns && n.name == name {
        Ok(())
    } else {
        Err(invalid(format!("expected {name}")))
    }
}
fn whitespace(n: &Node) -> Result<()> {
    if n.text.trim().is_empty() {
        Ok(())
    } else {
        Err(invalid(format!("unexpected text in {}", n.name)))
    }
}
fn require_empty(n: &Node) -> Result<()> {
    if n.children.is_empty() && n.text.trim().is_empty() {
        Ok(())
    } else {
        Err(invalid(format!("{} must be empty", n.name)))
    }
}
fn only_attrs(n: &Node, allowed: &[(RevisionAttributeNamespace, &str)]) -> Result<()> {
    if let Some(a) = n.attrs.iter().find(|a| {
        !allowed
            .iter()
            .any(|(ns, name)| *ns == a.namespace && *name == a.name)
    }) {
        Err(invalid(format!(
            "unexpected attribute '{}' on {}",
            a.name, n.name
        )))
    } else {
        Ok(())
    }
}
fn required_attr<'a>(n: &'a Node, ns: &RevisionAttributeNamespace, name: &str) -> Result<&'a str> {
    n.attrs
        .iter()
        .find(|a| &a.namespace == ns && a.name == name)
        .map(|a| a.value.as_str())
        .ok_or_else(|| invalid(format!("{} requires attribute {name}", n.name)))
}
fn req<'a>(n: &'a Node, name: &str) -> Result<&'a str> {
    required_attr(n, &RevisionAttributeNamespace::Unqualified, name)
}
fn opt<'a>(n: &'a Node, name: &str) -> Option<&'a str> {
    n.attrs
        .iter()
        .find(|a| a.namespace == RevisionAttributeNamespace::Unqualified && a.name == name)
        .map(|a| a.value.as_str())
}
fn required_u32_attr(n: &Node, name: &str) -> Result<u32> {
    req(n, name)?
        .parse()
        .map_err(|_| invalid(format!("invalid {name}")))
}
fn opt_u32(n: &Node, name: &str) -> Result<Option<u32>> {
    opt(n, name)
        .map(|v| v.parse().map_err(|_| invalid(format!("invalid {name}"))))
        .transpose()
}
fn opt_bool(n: &Node, name: &str) -> Result<Option<bool>> {
    opt(n, name).map(parse_bool).transpose()
}
fn parse_bool(v: &str) -> Result<bool> {
    match v {
        "1" | "true" => Ok(true),
        "0" | "false" => Ok(false),
        _ => Err(invalid("invalid XML boolean")),
    }
}
fn take_u32(
    a: &mut Vec<RevisionAttribute>,
    ns: &RevisionAttributeNamespace,
    name: &str,
) -> Result<Option<u32>> {
    if let Some(i) = a.iter().position(|a| &a.namespace == ns && a.name == name) {
        let v = a
            .remove(i)
            .value
            .parse()
            .map_err(|_| invalid(format!("invalid {name}")))?;
        Ok(Some(v))
    } else {
        Ok(None)
    }
}

fn ns_text(v: &ResolveResult<'_>) -> Result<String> {
    match v {
        ResolveResult::Bound(Namespace(v)) => Ok(std::str::from_utf8(v).map_err(xml_error)?.into()),
        _ => Err(invalid("unbound revision element namespace")),
    }
}
