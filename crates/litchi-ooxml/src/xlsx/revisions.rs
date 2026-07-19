//! Typed, inert SpreadsheetML shared-workbook revision parts.
//!
//! Revision payload formulas, rich text, and formatting are data only. They
//! are never calculated, executed, or used to retrieve external resources.

use crate::common::mce::process_ooxml;
use crate::error::{OoxmlError, Result};
use chrono::{DateTime, NaiveDateTime};
use litchi_opc::{BlobPart, OpcPackage, PackURI, Part};
use quick_xml::XmlVersion;
use quick_xml::events::{BytesStart, Event};
use quick_xml::name::{Namespace, ResolveResult};
use quick_xml::reader::NsReader;
use std::collections::{HashMap, HashSet};

const NS: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const STRICT_NS: &str = "http://purl.oclc.org/ooxml/spreadsheetml/main";
const REL_NS: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const STRICT_REL_NS: &str = "http://purl.oclc.org/ooxml/officeDocument/relationships";
const USERS_REL: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/usernames";
const STRICT_USERS_REL: &str = "http://purl.oclc.org/ooxml/officeDocument/relationships/usernames";
const HEADERS_REL: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/revisionHeaders";
const STRICT_HEADERS_REL: &str =
    "http://purl.oclc.org/ooxml/officeDocument/relationships/revisionHeaders";
const LOG_REL: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/revisionLog";
const STRICT_LOG_REL: &str = "http://purl.oclc.org/ooxml/officeDocument/relationships/revisionLog";
const USERS_CT: &str = "application/vnd.openxmlformats-officedocument.spreadsheetml.userNames+xml";
const HEADERS_CT: &str =
    "application/vnd.openxmlformats-officedocument.spreadsheetml.revisionHeaders+xml";
const LOG_CT: &str = "application/vnd.openxmlformats-officedocument.spreadsheetml.revisionLog+xml";
const MAX_PART_BYTES: usize = 16 * 1024 * 1024;
const MAX_DEPTH: usize = 128;
const MAX_NODES: usize = 500_000;
const MAX_STRING_BYTES: usize = 8 * 1024 * 1024;
const MAX_HEADERS: usize = 4096;
const MAX_USERS: usize = 65_536;
const MAX_LOGS: usize = 4096;
const MAX_RECORDS_PER_LOG: usize = 262_144;
const MAX_TOTAL_RECORDS: usize = 1_000_000;
const MAX_SHEET_IDS: usize = 1_048_576;

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum RevisionConformance {
    Transitional,
    Strict,
}
impl RevisionConformance {
    fn namespace(self) -> &'static str {
        match self {
            Self::Transitional => NS,
            Self::Strict => STRICT_NS,
        }
    }
    fn relationship_namespace(self) -> &'static str {
        match self {
            Self::Transitional => REL_NS,
            Self::Strict => STRICT_REL_NS,
        }
    }
    fn users_relationship(self) -> &'static str {
        match self {
            Self::Transitional => USERS_REL,
            Self::Strict => STRICT_USERS_REL,
        }
    }
    fn headers_relationship(self) -> &'static str {
        match self {
            Self::Transitional => HEADERS_REL,
            Self::Strict => STRICT_HEADERS_REL,
        }
    }
    fn log_relationship(self) -> &'static str {
        match self {
            Self::Transitional => LOG_REL,
            Self::Strict => STRICT_LOG_REL,
        }
    }
}

#[derive(Clone, Debug, PartialEq, Eq, PartialOrd, Ord, Hash)]
pub enum RevisionAttributeNamespace {
    Unqualified,
    Relationships,
    Xml,
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub struct RevisionAttribute {
    pub namespace: RevisionAttributeNamespace,
    pub name: String,
    pub value: String,
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub struct RevisionXmlElement {
    pub name: String,
    pub attributes: Vec<RevisionAttribute>,
    pub children: Vec<RevisionXmlElement>,
    pub text: String,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum RevisionRecordKind {
    RowColumn,
    Move,
    CustomView,
    SheetName,
    InsertSheet,
    CellChange,
    Format,
    AutoFormat,
    DefinedName,
    Comment,
    QueryTable,
    Conflict,
}
impl RevisionRecordKind {
    fn name(self) -> &'static str {
        match self {
            Self::RowColumn => "rrc",
            Self::Move => "rm",
            Self::CustomView => "rcv",
            Self::SheetName => "rsnm",
            Self::InsertSheet => "ris",
            Self::CellChange => "rcc",
            Self::Format => "rfmt",
            Self::AutoFormat => "raf",
            Self::DefinedName => "rdn",
            Self::Comment => "rcmt",
            Self::QueryTable => "rqt",
            Self::Conflict => "rcft",
        }
    }
    fn parse(value: &str) -> Option<Self> {
        Some(match value {
            "rrc" => Self::RowColumn,
            "rm" => Self::Move,
            "rcv" => Self::CustomView,
            "rsnm" => Self::SheetName,
            "ris" => Self::InsertSheet,
            "rcc" => Self::CellChange,
            "rfmt" => Self::Format,
            "raf" => Self::AutoFormat,
            "rdn" => Self::DefinedName,
            "rcmt" => Self::Comment,
            "rqt" => Self::QueryTable,
            "rcft" => Self::Conflict,
            _ => return None,
        })
    }
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub struct RevisionRecord {
    pub kind: RevisionRecordKind,
    pub revision_id: Option<u32>,
    pub sheet_id: Option<u32>,
    pub attributes: Vec<RevisionAttribute>,
    pub children: Vec<RevisionXmlElement>,
    pub text: String,
}

#[derive(Clone, Debug, PartialEq, Eq, Default)]
pub struct RevisionHeaderProperties {
    pub guid: String,
    pub last_guid: Option<String>,
    pub shared: Option<bool>,
    pub disk_revisions: Option<bool>,
    pub history: Option<bool>,
    pub track_revisions: Option<bool>,
    pub exclusive: Option<bool>,
    pub keep_change_history: Option<bool>,
    pub protected: Option<bool>,
    pub preserve_history: Option<u32>,
    pub revision_id: Option<u32>,
    pub version: Option<i32>,
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub struct RevisionHeader {
    pub guid: String,
    pub date_time: String,
    pub max_sheet_id: u32,
    pub user_name: String,
    pub relationship_id: String,
    pub min_revision_id: Option<u32>,
    pub max_revision_id: Option<u32>,
    pub sheet_ids: Vec<u32>,
    pub trailing_elements: Vec<RevisionXmlElement>,
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub struct RevisionHeaders {
    pub properties: RevisionHeaderProperties,
    pub headers: Vec<RevisionHeader>,
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub struct RevisionUser {
    pub guid: String,
    pub name: String,
    pub id: i32,
    pub date_time: String,
    pub extension_elements: Vec<RevisionXmlElement>,
}

#[derive(Clone, Debug, PartialEq, Eq, Default)]
pub struct RevisionUsers {
    pub users: Vec<RevisionUser>,
}

#[derive(Clone, Debug, PartialEq, Eq, Default)]
pub struct RevisionLog {
    pub records: Vec<RevisionRecord>,
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub struct RevisionLogPart {
    pub relationship_id: String,
    pub part_name: String,
    pub log: RevisionLog,
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub struct WorkbookRevisions {
    pub users_relationship_id: String,
    pub users_part_name: String,
    pub headers_relationship_id: String,
    pub headers_part_name: String,
    pub users: RevisionUsers,
    pub headers: RevisionHeaders,
    pub logs: Vec<RevisionLogPart>,
}

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
        valid_guid(&user.guid)?;
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
    valid_guid(&properties.guid)?;
    if let Some(v) = &properties.last_guid {
        valid_guid(v)?;
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
        valid_guid(&header.guid)?;
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

pub fn load_workbook_revisions(package: &OpcPackage) -> Result<Option<WorkbookRevisions>> {
    let workbook = package.main_document_part()?;
    require_workbook_type(workbook.content_type())?;
    let workbook_name = workbook.partname().to_string();
    validate_sources(package, &workbook_name, None)?;
    let users_rels: Vec<_> = workbook
        .rels()
        .iter()
        .filter(|r| is_users_rel(r.reltype()))
        .collect();
    let headers_rels: Vec<_> = workbook
        .rels()
        .iter()
        .filter(|r| is_headers_rel(r.reltype()))
        .collect();
    if users_rels.len() > 1 || headers_rels.len() > 1 {
        return Err(invalid("workbook has duplicate revision relationships"));
    }
    if users_rels.is_empty() && headers_rels.is_empty() {
        reject_orphans(package, None, None, &HashSet::new())?;
        return Ok(None);
    }
    if users_rels.len() != 1 || headers_rels.len() != 1 {
        return Err(invalid(
            "userNames and revisionHeaders relationships must occur together",
        ));
    }
    let users_rel = users_rels[0];
    let headers_rel = headers_rels[0];
    if users_rel.is_external() || headers_rel.is_external() {
        return Err(invalid("workbook revision relationships must be internal"));
    }
    let users_uri = users_rel.target_partname()?;
    let headers_uri = headers_rel.target_partname()?;
    let users_part = package.get_part(&users_uri)?;
    let headers_part = package.get_part(&headers_uri)?;
    require_ct(users_part, USERS_CT)?;
    require_ct(headers_part, HEADERS_CT)?;
    no_rels(users_part, "userNames")?;
    validate_sources(package, &workbook_name, Some(headers_uri.as_str()))?;
    let headers = parse_revision_headers(headers_part.blob())?;
    let mut logs = Vec::new();
    let mut targets = HashSet::new();
    for header in &headers.headers {
        let rel = headers_part
            .rels()
            .get(&header.relationship_id)
            .ok_or_else(|| {
                invalid(format!(
                    "missing revisionLog relationship '{}'",
                    header.relationship_id
                ))
            })?;
        if !is_log_rel(rel.reltype()) || rel.is_external() {
            return Err(invalid(
                "revision header relationship must target an internal revisionLog",
            ));
        }
        let uri = rel.target_partname()?;
        if !targets.insert(uri.to_string()) {
            return Err(invalid("duplicate revisionLog target"));
        }
        let part = package.get_part(&uri)?;
        require_ct(part, LOG_CT)?;
        no_rels(part, "revisionLog")?;
        logs.push(RevisionLogPart {
            relationship_id: header.relationship_id.clone(),
            part_name: uri.to_string(),
            log: parse_revision_log(part.blob())?,
        });
    }
    if headers_part.rels().iter().any(|r| !is_log_rel(r.reltype()))
        || headers_part.rels().iter().count() != headers.headers.len()
    {
        return Err(invalid(
            "revisionHeaders has unmatched or invalid relationships",
        ));
    }
    reject_orphans(
        package,
        Some(users_uri.as_str()),
        Some(headers_uri.as_str()),
        &targets,
    )?;
    let value = WorkbookRevisions {
        users_relationship_id: users_rel.r_id().into(),
        users_part_name: users_uri.to_string(),
        headers_relationship_id: headers_rel.r_id().into(),
        headers_part_name: headers_uri.to_string(),
        users: parse_revision_users(users_part.blob())?,
        headers,
        logs,
    };
    validate_package(&value)?;
    Ok(Some(value))
}

pub fn store_workbook_revisions(
    package: &mut OpcPackage,
    value: &WorkbookRevisions,
    conformance: RevisionConformance,
) -> Result<()> {
    validate_package(value)?;
    if load_workbook_revisions(package)?.is_some() {
        return Err(invalid("package already contains workbook revisions"));
    }
    let workbook = package.main_document_part()?;
    require_workbook_type(workbook.content_type())?;
    let workbook_uri = workbook.partname().clone();
    if workbook.rels().get(&value.users_relationship_id).is_some()
        || workbook
            .rels()
            .get(&value.headers_relationship_id)
            .is_some()
    {
        return Err(invalid("workbook revision relationship id already exists"));
    }
    let users_uri = PackURI::new(&value.users_part_name).map_err(OoxmlError::InvalidUri)?;
    let headers_uri = PackURI::new(&value.headers_part_name).map_err(OoxmlError::InvalidUri)?;
    let mut names = HashSet::from([users_uri.to_string(), headers_uri.to_string()]);
    if names.len() != 2 {
        return Err(invalid("revision part names must be unique"));
    }
    let users_xml = write_revision_users(&value.users, conformance)?;
    let headers_xml = write_revision_headers(&value.headers, conformance)?;
    let mut pending = Vec::new();
    for log in &value.logs {
        let uri = PackURI::new(&log.part_name).map_err(OoxmlError::InvalidUri)?;
        if !names.insert(uri.to_string()) {
            return Err(invalid("duplicate revision part name"));
        }
        pending.push((
            uri,
            log.relationship_id.clone(),
            write_revision_log(&log.log, conformance)?,
        ));
    }
    for name in &names {
        let uri = PackURI::new(name).map_err(OoxmlError::InvalidUri)?;
        if package.iter_parts().any(|p| p.partname() == &uri) {
            return Err(invalid(format!("revision part '{uri}' already exists")));
        }
    }
    package.add_part(Box::new(BlobPart::new(
        users_uri.clone(),
        USERS_CT.into(),
        users_xml,
    )));
    package.add_part(Box::new(BlobPart::new(
        headers_uri.clone(),
        HEADERS_CT.into(),
        headers_xml,
    )));
    let users_target = users_uri.relative_ref(workbook_uri.base_uri());
    let headers_target = headers_uri.relative_ref(workbook_uri.base_uri());
    {
        let rels = package.get_part_mut(&workbook_uri)?.rels_mut();
        rels.add_relationship(
            conformance.users_relationship().into(),
            users_target,
            value.users_relationship_id.clone(),
            false,
        );
        rels.add_relationship(
            conformance.headers_relationship().into(),
            headers_target,
            value.headers_relationship_id.clone(),
            false,
        );
    }
    for (uri, rid, xml) in pending {
        let target = uri.relative_ref(headers_uri.base_uri());
        package.add_part(Box::new(BlobPart::new(uri, LOG_CT.into(), xml)));
        package
            .get_part_mut(&headers_uri)?
            .rels_mut()
            .add_relationship(conformance.log_relationship().into(), target, rid, false);
    }
    Ok(())
}

fn validate_users(v: &RevisionUsers) -> Result<()> {
    if v.users.len() > MAX_USERS {
        return Err(limit("revision users"));
    }
    let (mut ids, mut guids) = (HashSet::new(), HashSet::new());
    for u in &v.users {
        valid_guid(&u.guid)?;
        valid_date(&u.date_time)?;
        bounded(&u.name)?;
        if !ids.insert(u.id) || !guids.insert(&u.guid) {
            return Err(invalid("duplicate revision user"));
        }
        for e in &u.extension_elements {
            validate_element(e, 1)?;
        }
    }
    Ok(())
}
fn validate_headers(v: &RevisionHeaders) -> Result<()> {
    valid_guid(&v.properties.guid)?;
    if let Some(g) = &v.properties.last_guid {
        valid_guid(g)?;
    }
    if v.headers.len() > MAX_HEADERS {
        return Err(limit("revision headers"));
    }
    let (mut gs, mut rs) = (HashSet::new(), HashSet::new());
    for h in &v.headers {
        valid_guid(&h.guid)?;
        valid_date(&h.date_time)?;
        bounded(&h.user_name)?;
        if h.relationship_id.is_empty() || !gs.insert(&h.guid) || !rs.insert(&h.relationship_id) {
            return Err(invalid("duplicate/empty revision header identifier"));
        }
        if h.sheet_ids.len() > MAX_SHEET_IDS || h.sheet_ids.iter().any(|id| *id >= h.max_sheet_id) {
            return Err(invalid("invalid revision header sheet map"));
        }
        if h.min_revision_id
            .zip(h.max_revision_id)
            .is_some_and(|(a, b)| a > b)
        {
            return Err(invalid("header minRId exceeds maxRId"));
        }
        for e in &h.trailing_elements {
            validate_element(e, 1)?;
        }
    }
    if v.headers
        .last()
        .is_some_and(|h| h.guid != v.properties.guid)
    {
        return Err(invalid("headers guid must match newest header"));
    }
    Ok(())
}
fn validate_log(v: &RevisionLog) -> Result<()> {
    if v.records.len() > MAX_RECORDS_PER_LOG {
        return Err(limit("revision records"));
    }
    for r in &v.records {
        validate_attrs(&r.attributes)?;
        bounded(&r.text)?;
        for e in &r.children {
            validate_element(e, 1)?;
        }
    }
    Ok(())
}
fn validate_package(v: &WorkbookRevisions) -> Result<()> {
    validate_users(&v.users)?;
    validate_headers(&v.headers)?;
    if v.logs.len() > MAX_LOGS || v.logs.len() != v.headers.headers.len() {
        return Err(invalid("revision log/header count mismatch"));
    }
    if v.users_relationship_id.is_empty()
        || v.headers_relationship_id.is_empty()
        || v.users_relationship_id == v.headers_relationship_id
    {
        return Err(invalid("invalid workbook revision relationship ids"));
    }
    let headers: HashMap<_, _> = v
        .headers
        .headers
        .iter()
        .map(|h| (&h.relationship_id, h))
        .collect();
    let mut rels = HashSet::new();
    let mut ids = HashSet::new();
    let mut total = 0usize;
    for p in &v.logs {
        if !rels.insert(&p.relationship_id) {
            return Err(invalid("duplicate revisionLog relationship id"));
        }
        let h = headers
            .get(&p.relationship_id)
            .ok_or_else(|| invalid("revisionLog lacks matching header"))?;
        validate_log(&p.log)?;
        total = total
            .checked_add(p.log.records.len())
            .ok_or_else(|| limit("revision records"))?;
        if total > MAX_TOTAL_RECORDS {
            return Err(limit("total revision records"));
        }
        for r in &p.log.records {
            if let Some(id) = r.revision_id {
                if !ids.insert(id) {
                    return Err(invalid("duplicate revision id across logs"));
                }
                if h.min_revision_id.is_some_and(|min| id < min)
                    || h.max_revision_id.is_some_and(|max| id > max)
                {
                    return Err(invalid("revision id outside header range"));
                }
            }
            if r.sheet_id.is_some_and(|id| !h.sheet_ids.contains(&id)) {
                return Err(invalid("revision record sheet id absent from header map"));
            }
        }
    }
    Ok(())
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
                let d = crate::common::xml::decode_xml_reference(&r)?;
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
                if v.as_ref() == REL_NS.as_bytes() || v.as_ref() == STRICT_REL_NS.as_bytes() =>
            {
                RevisionAttributeNamespace::Relationships
            },
            ResolveResult::Bound(Namespace(v))
                if v.as_ref() == b"http://www.w3.org/XML/1998/namespace" =>
            {
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
fn validate_element(e: &RevisionXmlElement, depth: usize) -> Result<()> {
    if depth > MAX_DEPTH || !ncname(&e.name) {
        return Err(invalid("invalid revision payload element/depth"));
    }
    bounded(&e.text)?;
    validate_attrs(&e.attributes)?;
    for c in &e.children {
        validate_element(c, depth + 1)?;
    }
    Ok(())
}
fn validate_attrs(a: &[RevisionAttribute]) -> Result<()> {
    let mut seen = HashSet::new();
    for x in a {
        if !ncname(&x.name) || !seen.insert((x.namespace.clone(), &x.name)) {
            return Err(invalid("invalid/duplicate revision payload attribute"));
        }
        bounded(&x.value)?;
    }
    Ok(())
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
                attr(out, &n, &a.value)
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
        attr(o, n, v)
    }
}
fn bool_attr(o: &mut String, n: &str, v: Option<bool>) {
    if let Some(v) = v {
        attr(o, n, if v { "1" } else { "0" })
    }
}
fn num_attr(o: &mut String, n: &str, v: Option<u32>) {
    if let Some(v) = v {
        attr(o, n, &v.to_string())
    }
}
fn text(o: &mut String, v: &str) {
    escape(o, v, false)
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
fn valid_guid(v: &str) -> Result<()> {
    let b = v.as_bytes();
    if b.len() != 38
        || b[0] != b'{'
        || b[37] != b'}'
        || [9, 14, 19, 24].iter().any(|i| b[*i] != b'-')
        || b[1..37]
            .iter()
            .enumerate()
            .any(|(i, c)| ![8, 13, 18, 23].contains(&i) && !c.is_ascii_hexdigit())
    {
        Err(invalid(format!("invalid GUID '{v}'")))
    } else {
        Ok(())
    }
}
fn valid_date(v: &str) -> Result<()> {
    if DateTime::parse_from_rfc3339(v).is_ok()
        || NaiveDateTime::parse_from_str(v, "%Y-%m-%dT%H:%M:%S%.f").is_ok()
    {
        Ok(())
    } else {
        Err(invalid(format!("invalid XML dateTime '{v}'")))
    }
}
fn ncname(v: &str) -> bool {
    let mut c = v.chars();
    c.next().is_some_and(|x| x == '_' || x.is_alphabetic())
        && c.all(|x| x == '_' || x == '-' || x == '.' || x.is_alphanumeric())
}
fn bounded(v: &str) -> Result<()> {
    if v.len() > MAX_STRING_BYTES {
        Err(limit("revision string bytes"))
    } else {
        Ok(())
    }
}
fn ns_text(v: &ResolveResult<'_>) -> Result<String> {
    match v {
        ResolveResult::Bound(Namespace(v)) => {
            Ok(std::str::from_utf8(v.as_ref()).map_err(xml_error)?.into())
        },
        _ => Err(invalid("unbound revision element namespace")),
    }
}

fn require_workbook_type(v: &str) -> Result<()> {
    if matches!(
        v,
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"
            | "application/vnd.openxmlformats-officedocument.spreadsheetml.template.main+xml"
            | "application/vnd.ms-excel.sheet.macroEnabled.main+xml"
            | "application/vnd.ms-excel.template.macroEnabled.main+xml"
            | "application/vnd.ms-excel.addin.macroEnabled.main+xml"
    ) {
        Ok(())
    } else {
        Err(invalid("revision source is not a workbook main part"))
    }
}
fn require_ct(p: &dyn Part, v: &str) -> Result<()> {
    if p.content_type() == v {
        Ok(())
    } else {
        Err(OoxmlError::InvalidContentType {
            expected: v.into(),
            got: p.content_type().into(),
        })
    }
}
fn no_rels(p: &dyn Part, label: &str) -> Result<()> {
    if p.rels().iter().next().is_none() {
        Ok(())
    } else {
        Err(invalid(format!("{label} part must not have relationships")))
    }
}
fn is_users_rel(v: &str) -> bool {
    matches!(v, USERS_REL | STRICT_USERS_REL)
}
fn is_headers_rel(v: &str) -> bool {
    matches!(v, HEADERS_REL | STRICT_HEADERS_REL)
}
fn is_log_rel(v: &str) -> bool {
    matches!(v, LOG_REL | STRICT_LOG_REL)
}
fn special(v: &str) -> bool {
    is_users_rel(v) || is_headers_rel(v) || is_log_rel(v)
}
fn validate_sources(package: &OpcPackage, workbook: &str, headers: Option<&str>) -> Result<()> {
    if package.rels().iter().any(|r| special(r.reltype())) {
        return Err(invalid("package root cannot source revision relationships"));
    }
    for p in package.iter_parts() {
        for r in p.rels().iter() {
            if (is_users_rel(r.reltype()) || is_headers_rel(r.reltype()))
                && p.partname().as_str() != workbook
            {
                return Err(invalid("workbook revision relationship has invalid source"));
            }
            if headers.is_some()
                && is_log_rel(r.reltype())
                && Some(p.partname().as_str()) != headers
            {
                return Err(invalid("revisionLog relationship has invalid source"));
            }
        }
    }
    Ok(())
}
fn reject_orphans(
    package: &OpcPackage,
    users: Option<&str>,
    headers: Option<&str>,
    logs: &HashSet<String>,
) -> Result<()> {
    for p in package.iter_parts() {
        let n = p.partname().as_str();
        if p.content_type() == USERS_CT && Some(n) != users {
            return Err(invalid("orphan userNames part"));
        }
        if p.content_type() == HEADERS_CT && Some(n) != headers {
            return Err(invalid("orphan revisionHeaders part"));
        }
        if p.content_type() == LOG_CT && !logs.contains(n) {
            return Err(invalid("orphan revisionLog part"));
        }
    }
    Ok(())
}
fn invalid(v: impl Into<String>) -> OoxmlError {
    OoxmlError::InvalidFormat(v.into())
}
fn limit(v: &str) -> OoxmlError {
    invalid(format!("{v} exceeds configured limit"))
}
fn xml_error(v: impl std::fmt::Display) -> OoxmlError {
    OoxmlError::Xml(v.to_string())
}

#[cfg(test)]
mod tests {
    use super::*;
    const LO: &[u8] = include_bytes!(
        "../../../../3rdparty/libreoffice-core/sc/qa/unit/data/xlsx/track-changes/simple-cell-changes.xlsx"
    );
    const POI: &[u8] = include_bytes!(
        "../../../../3rdparty/poi/test-data/spreadsheet/workbookProtection_workbook_revision_protected.xlsx"
    );
    fn guid(n: u8) -> String {
        format!("{{00000000-0000-0000-0000-{n:012X}}}")
    }
    fn log() -> RevisionLog {
        RevisionLog {
            records: vec![RevisionRecord {
                kind: RevisionRecordKind::CellChange,
                revision_id: Some(1),
                sheet_id: Some(1),
                attributes: vec![],
                children: vec![RevisionXmlElement {
                    name: "nc".into(),
                    attributes: vec![RevisionAttribute {
                        namespace: RevisionAttributeNamespace::Unqualified,
                        name: "r".into(),
                        value: "A1".into(),
                    }],
                    children: vec![RevisionXmlElement {
                        name: "v".into(),
                        attributes: vec![],
                        children: vec![],
                        text: "=1+1".into(),
                    }],
                    text: String::new(),
                }],
                text: String::new(),
            }],
        }
    }
    fn value() -> WorkbookRevisions {
        let h = RevisionHeader {
            guid: guid(2),
            date_time: "2026-07-17T12:00:00Z".into(),
            max_sheet_id: 2,
            user_name: "Reviewer".into(),
            relationship_id: "rId1".into(),
            min_revision_id: Some(1),
            max_revision_id: Some(1),
            sheet_ids: vec![1],
            trailing_elements: vec![],
        };
        WorkbookRevisions {
            users_relationship_id: "rId2".into(),
            users_part_name: "/xl/revisions/userNames.xml".into(),
            headers_relationship_id: "rId3".into(),
            headers_part_name: "/xl/revisions/revisionHeaders.xml".into(),
            users: RevisionUsers::default(),
            headers: RevisionHeaders {
                properties: RevisionHeaderProperties {
                    guid: h.guid.clone(),
                    disk_revisions: Some(true),
                    ..Default::default()
                },
                headers: vec![h],
            },
            logs: vec![RevisionLogPart {
                relationship_id: "rId1".into(),
                part_name: "/xl/revisions/revisionLog1.xml".into(),
                log: log(),
            }],
        }
    }
    fn package() -> OpcPackage {
        let mut p = OpcPackage::new();
        p.rels_mut().add_relationship(
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"
                .into(),
            "xl/workbook.xml".into(),
            "rId1".into(),
            false,
        );
        p.add_part(Box::new(BlobPart::new(
            PackURI::new("/xl/workbook.xml").unwrap(),
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml".into(),
            Vec::new(),
        )));
        p
    }
    #[test]
    fn loads_libreoffice_and_poi_reference_packages() {
        let a = load_workbook_revisions(&OpcPackage::from_bytes(LO).unwrap())
            .unwrap()
            .unwrap();
        assert_eq!(a.logs.len(), 5);
        assert!(
            a.logs
                .iter()
                .flat_map(|l| &l.log.records)
                .any(|r| r.kind == RevisionRecordKind::RowColumn)
        );
        assert!(
            a.logs
                .iter()
                .flat_map(|l| &l.log.records)
                .any(|r| r.kind == RevisionRecordKind::CustomView)
        );
        let b = load_workbook_revisions(&OpcPackage::from_bytes(POI).unwrap())
            .unwrap()
            .unwrap();
        assert_eq!(b.logs.len(), 3);
        assert_eq!(b.logs.iter().map(|l| l.log.records.len()).sum::<usize>(), 2);
    }
    #[test]
    fn strict_writers_are_deterministic_and_round_trip() {
        let v = value();
        let u = write_revision_users(&v.users, RevisionConformance::Strict).unwrap();
        let h = write_revision_headers(&v.headers, RevisionConformance::Strict).unwrap();
        let l = write_revision_log(&v.logs[0].log, RevisionConformance::Strict).unwrap();
        assert_eq!(
            u,
            write_revision_users(&v.users, RevisionConformance::Strict).unwrap()
        );
        assert_eq!(parse_revision_users(&u).unwrap(), v.users);
        assert_eq!(parse_revision_headers(&h).unwrap(), v.headers);
        assert_eq!(parse_revision_log(&l).unwrap(), v.logs[0].log);
    }
    #[test]
    fn mce_fallback() {
        let x = format!(
            r#"<revisions xmlns="{NS}" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:z="urn:z" mc:Ignorable="z"><mc:AlternateContent><mc:Choice Requires="z"><z:x/></mc:Choice><mc:Fallback><rcv guid="{}" action="add"/></mc:Fallback></mc:AlternateContent></revisions>"#,
            guid(3)
        );
        assert_eq!(
            parse_revision_log(x.as_bytes()).unwrap().records[0].kind,
            RevisionRecordKind::CustomView
        );
    }
    #[test]
    fn package_writer_round_trip() {
        let mut p = package();
        let v = value();
        store_workbook_revisions(&mut p, &v, RevisionConformance::Strict).unwrap();
        assert_eq!(load_workbook_revisions(&p).unwrap().unwrap(), v);
    }
    #[test]
    fn malformed_and_caps() {
        for x in [
            format!(r#"<users xmlns="{NS}" count="2"/>"#),
            format!(r#"<headers xmlns="{NS}" guid="bad"/>"#),
            format!(r#"<revisions xmlns="{NS}"><bad/></revisions>"#),
            format!(r#"<!DOCTYPE x><revisions xmlns="{NS}"/>"#),
        ] {
            assert!(if x.contains("<users") {
                parse_revision_users(x.as_bytes()).is_err()
            } else if x.contains("<headers") {
                parse_revision_headers(x.as_bytes()).is_err()
            } else {
                parse_revision_log(x.as_bytes()).is_err()
            });
        }
        assert!(parse_revision_log(&vec![b' '; MAX_PART_BYTES + 1]).is_err());
        let deep = format!(
            r#"<revisions xmlns="{NS}"><rcc>{}{}</rcc></revisions>"#,
            "<x>".repeat(MAX_DEPTH),
            "</x>".repeat(MAX_DEPTH)
        );
        assert!(parse_revision_log(deep.as_bytes()).is_err());
    }
    #[test]
    fn graph_and_reference_errors() {
        let mut v = value();
        v.logs[0].log.records[0].sheet_id = Some(9);
        assert!(
            store_workbook_revisions(&mut package(), &v, RevisionConformance::Transitional)
                .is_err()
        );
        let mut p = package();
        p.get_part_mut(&PackURI::new("/xl/workbook.xml").unwrap())
            .unwrap()
            .rels_mut()
            .add_relationship(
                USERS_REL.into(),
                "https://invalid.example/users".into(),
                "rId2".into(),
                true,
            );
        assert!(load_workbook_revisions(&p).is_err());
    }
}
