//! Semantic workbook revision values and invariants.

use crate::error::Result;
use chrono::{DateTime, NaiveDateTime};
use litchi_ooxml_common::custom_xml::valid_guid as is_valid_guid;
use std::collections::{HashMap, HashSet};

use super::{invalid, limit};

pub(super) const NS: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
pub(super) const STRICT_NS: &str = "http://purl.oclc.org/ooxml/spreadsheetml/main";
pub(super) const REL_NS: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
pub(super) const STRICT_REL_NS: &str = "http://purl.oclc.org/ooxml/officeDocument/relationships";
pub(super) const USERS_REL: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/usernames";
pub(super) const STRICT_USERS_REL: &str =
    "http://purl.oclc.org/ooxml/officeDocument/relationships/usernames";
pub(super) const HEADERS_REL: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/revisionHeaders";
pub(super) const STRICT_HEADERS_REL: &str =
    "http://purl.oclc.org/ooxml/officeDocument/relationships/revisionHeaders";
pub(super) const LOG_REL: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/revisionLog";
pub(super) const STRICT_LOG_REL: &str =
    "http://purl.oclc.org/ooxml/officeDocument/relationships/revisionLog";
pub(super) const USERS_CT: &str =
    "application/vnd.openxmlformats-officedocument.spreadsheetml.userNames+xml";
pub(super) const HEADERS_CT: &str =
    "application/vnd.openxmlformats-officedocument.spreadsheetml.revisionHeaders+xml";
pub(super) const LOG_CT: &str =
    "application/vnd.openxmlformats-officedocument.spreadsheetml.revisionLog+xml";
pub(super) const MAX_PART_BYTES: usize = 16 * 1024 * 1024;
pub(super) const MAX_DEPTH: usize = 128;
pub(super) const MAX_NODES: usize = 500_000;
pub(super) const MAX_STRING_BYTES: usize = 8 * 1024 * 1024;
pub(super) const MAX_HEADERS: usize = 4096;
pub(super) const MAX_USERS: usize = 65_536;
pub(super) const MAX_LOGS: usize = 4096;
pub(super) const MAX_RECORDS_PER_LOG: usize = 262_144;
pub(super) const MAX_TOTAL_RECORDS: usize = 1_000_000;
pub(super) const MAX_SHEET_IDS: usize = 1_048_576;

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum RevisionConformance {
    Transitional,
    Strict,
}
impl RevisionConformance {
    pub(super) fn namespace(self) -> &'static str {
        match self {
            Self::Transitional => NS,
            Self::Strict => STRICT_NS,
        }
    }
    pub(super) fn relationship_namespace(self) -> &'static str {
        match self {
            Self::Transitional => REL_NS,
            Self::Strict => STRICT_REL_NS,
        }
    }
    pub(super) fn users_relationship(self) -> &'static str {
        match self {
            Self::Transitional => USERS_REL,
            Self::Strict => STRICT_USERS_REL,
        }
    }
    pub(super) fn headers_relationship(self) -> &'static str {
        match self {
            Self::Transitional => HEADERS_REL,
            Self::Strict => STRICT_HEADERS_REL,
        }
    }
    pub(super) fn log_relationship(self) -> &'static str {
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
    pub(super) fn name(self) -> &'static str {
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
    pub(super) fn parse(value: &str) -> Option<Self> {
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
pub struct Revisions {
    pub users_relationship_id: String,
    pub users_part_name: String,
    pub headers_relationship_id: String,
    pub headers_part_name: String,
    pub users: RevisionUsers,
    pub headers: RevisionHeaders,
    pub logs: Vec<RevisionLogPart>,
}

pub(super) fn validate_users(v: &RevisionUsers) -> Result<()> {
    if v.users.len() > MAX_USERS {
        return Err(limit("revision users"));
    }
    let (mut ids, mut guids) = (HashSet::new(), HashSet::new());
    for u in &v.users {
        validate_guid(&u.guid)?;
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
pub(super) fn validate_headers(v: &RevisionHeaders) -> Result<()> {
    validate_guid(&v.properties.guid)?;
    if let Some(g) = &v.properties.last_guid {
        validate_guid(g)?;
    }
    if v.headers.len() > MAX_HEADERS {
        return Err(limit("revision headers"));
    }
    let (mut gs, mut rs) = (HashSet::new(), HashSet::new());
    for h in &v.headers {
        validate_guid(&h.guid)?;
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
pub(super) fn validate_log(v: &RevisionLog) -> Result<()> {
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
pub(super) fn validate_package(v: &Revisions) -> Result<()> {
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

pub(super) fn validate_element(e: &RevisionXmlElement, depth: usize) -> Result<()> {
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
pub(super) fn validate_attrs(a: &[RevisionAttribute]) -> Result<()> {
    let mut seen = HashSet::new();
    for x in a {
        if !ncname(&x.name) || !seen.insert((x.namespace.clone(), &x.name)) {
            return Err(invalid("invalid/duplicate revision payload attribute"));
        }
        bounded(&x.value)?;
    }
    Ok(())
}

pub(super) fn validate_guid(value: &str) -> Result<()> {
    if !is_valid_guid(value) {
        Err(invalid(format!("invalid GUID '{value}'")))
    } else {
        Ok(())
    }
}
pub(super) fn valid_date(v: &str) -> Result<()> {
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
pub(super) fn bounded(v: &str) -> Result<()> {
    if v.len() > MAX_STRING_BYTES {
        Err(limit("revision string bytes"))
    } else {
        Ok(())
    }
}
