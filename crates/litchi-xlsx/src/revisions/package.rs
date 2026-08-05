//! OPC graph ownership for workbook revision parts.

use crate::error::Result;
use litchi_opc::{BlobPart, OpcPackage, PackURI, Part};
use std::collections::HashSet;

use super::codec::{
    parse_revision_headers, parse_revision_log, parse_revision_users, write_revision_headers,
    write_revision_log, write_revision_users,
};
use super::invalid;
use super::model::*;

pub fn load_workbook_revisions(package: &OpcPackage) -> Result<Option<Revisions>> {
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
    let value = Revisions {
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
    value: &Revisions,
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
    let users_uri = PackURI::new(&value.users_part_name).map_err(invalid)?;
    let headers_uri = PackURI::new(&value.headers_part_name).map_err(invalid)?;
    let mut names = HashSet::from([users_uri.to_string(), headers_uri.to_string()]);
    if names.len() != 2 {
        return Err(invalid("revision part names must be unique"));
    }
    let users_xml = write_revision_users(&value.users, conformance)?;
    let headers_xml = write_revision_headers(&value.headers, conformance)?;
    let mut pending = Vec::new();
    for log in &value.logs {
        let uri = PackURI::new(&log.part_name).map_err(invalid)?;
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
        let uri = PackURI::new(name).map_err(invalid)?;
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
        Err(invalid(format!(
            "invalid content type: expected {v}, got {}",
            p.content_type()
        )))
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
