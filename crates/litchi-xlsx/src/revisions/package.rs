//! OPC graph ownership for workbook revision parts.

use crate::error::Result;
use litchi_opc::{BlobPart, OpcPackage, PackURI, Part};
use std::collections::{HashMap, HashSet};

use super::codec::{
    parse_revision_headers, parse_revision_log, parse_revision_users, write_revision_headers,
    write_revision_log, write_revision_users,
};
use super::invalid;
use super::model::*;
use super::snapshot::{Snapshot, SourceRelationship};

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
    let mut candidate = package.clone();
    store_workbook_revisions_inner(&mut candidate, value, conformance)?;
    *package = candidate;
    Ok(())
}

fn store_workbook_revisions_inner(
    package: &mut OpcPackage,
    value: &Revisions,
    conformance: RevisionConformance,
) -> Result<()> {
    validate_physical_value(value)?;
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
        package
            .validate_new_part_name(&uri)
            .map_err(|error| invalid(error.to_string()))?;
    }
    package.try_add_part(Box::new(BlobPart::new(
        users_uri.clone(),
        USERS_CT.into(),
        users_xml,
    )))?;
    package.try_add_part(Box::new(BlobPart::new(
        headers_uri.clone(),
        HEADERS_CT.into(),
        headers_xml,
    )))?;
    let users_target = users_uri.relative_ref(workbook_uri.base_uri());
    let headers_target = headers_uri.relative_ref(workbook_uri.base_uri());
    {
        let rels = package.get_part_mut(&workbook_uri)?.rels_mut();
        rels.try_add_relationship(
            conformance.users_relationship().into(),
            users_target,
            value.users_relationship_id.clone(),
            litchi_opc::TargetMode::Internal,
        )?;
        rels.try_add_relationship(
            conformance.headers_relationship().into(),
            headers_target,
            value.headers_relationship_id.clone(),
            litchi_opc::TargetMode::Internal,
        )?;
    }
    for (uri, rid, xml) in pending {
        let target = uri.relative_ref(headers_uri.base_uri());
        package.try_add_part(Box::new(BlobPart::new(uri, LOG_CT.into(), xml)))?;
        package
            .get_part_mut(&headers_uri)?
            .rels_mut()
            .try_add_relationship(
                conformance.log_relationship().into(),
                target,
                rid,
                litchi_opc::TargetMode::Internal,
            )?;
    }
    package.unsign();
    Ok(())
}

/// Replace the complete revision graph on a clone-staged package.
///
/// Existing parts whose typed value is unchanged are left byte-for-byte
/// untouched. A topology change is planned as a remove-and-create operation;
/// callers publish this helper only on a clone so every failure is atomic.
pub(crate) fn replace_workbook_revisions(
    package: &mut OpcPackage,
    value: Option<&Revisions>,
    conformance: RevisionConformance,
) -> Result<()> {
    let existing = load_workbook_revisions(package)?;
    match (existing.as_ref(), value) {
        (None, None) => return Ok(()),
        (Some(existing), None) => remove_workbook_revisions_inner(package, existing)?,
        (None, Some(value)) => store_workbook_revisions_inner(package, value, conformance)?,
        (Some(existing), Some(value)) => {
            validate_package(value)?;
            validate_physical_value(value)?;
            if can_update_in_place(existing, value) {
                update_in_place(package, existing, value, conformance)?;
            } else {
                remove_workbook_revisions_inner(package, existing)?;
                store_workbook_revisions_inner(package, value, conformance)?;
            }
        },
    }
    Ok(())
}

/// Remove the complete workbook revision owner atomically.
pub fn remove_workbook_revisions(package: &mut OpcPackage) -> Result<bool> {
    let mut candidate = package.clone();
    let Some(value) = load_workbook_revisions(&candidate)? else {
        return Ok(false);
    };
    remove_workbook_revisions_inner(&mut candidate, &value)?;
    validate_graph(&candidate)?;
    *package = candidate;
    Ok(true)
}

fn remove_workbook_revisions_inner(package: &mut OpcPackage, value: &Revisions) -> Result<()> {
    validate_physical_value(value)?;
    let workbook = package.main_document_part()?.partname().clone();
    let users_uri = PackURI::new(&value.users_part_name).map_err(invalid)?;
    let headers_uri = PackURI::new(&value.headers_part_name).map_err(invalid)?;
    let log_uris = value
        .logs
        .iter()
        .map(|log| PackURI::new(&log.part_name).map_err(invalid))
        .collect::<Result<Vec<_>>>()?;

    let mut names = vec![users_uri, headers_uri];
    names.extend(log_uris);
    for name in &names {
        let actual = package.get_part(name)?.partname().clone();
        if incoming_references(package, &actual)? != 1 {
            return Err(invalid(format!(
                "revision part '{}' has unexpected incoming relationships",
                actual
            )));
        }
    }

    let headers = package.get_part(&PackURI::new(&value.headers_part_name).map_err(invalid)?)?;
    let header_log_ids = headers
        .rels()
        .iter()
        .filter(|relationship| is_log_rel(relationship.reltype()))
        .map(|relationship| relationship.r_id().to_owned())
        .collect::<Vec<_>>();
    {
        let workbook_part = package.get_part_mut(&workbook)?;
        workbook_part
            .rels_mut()
            .remove(&value.users_relationship_id);
        workbook_part
            .rels_mut()
            .remove(&value.headers_relationship_id);
    }
    let headers_uri = PackURI::new(&value.headers_part_name).map_err(invalid)?;
    let headers_part = package.get_part_mut(&headers_uri)?;
    for relationship_id in header_log_ids {
        headers_part.rels_mut().remove(&relationship_id);
    }
    for name in names {
        let actual = package.get_part(&name)?.partname().clone();
        package.remove_part(&actual);
    }
    package.unsign();
    Ok(())
}

fn can_update_in_place(current: &Revisions, value: &Revisions) -> bool {
    if current.users_relationship_id != value.users_relationship_id
        || current.users_part_name != value.users_part_name
        || current.headers_relationship_id != value.headers_relationship_id
        || current.headers_part_name != value.headers_part_name
    {
        return false;
    }
    let current_logs: HashMap<_, _> = current
        .logs
        .iter()
        .map(|log| (log.relationship_id.as_str(), log.part_name.as_str()))
        .collect();
    let current_parts: HashSet<_> = current
        .logs
        .iter()
        .map(|log| log.part_name.as_str())
        .collect();
    value
        .logs
        .iter()
        .all(|log| match current_logs.get(log.relationship_id.as_str()) {
            Some(part_name) => *part_name == log.part_name,
            None => !current_parts.contains(log.part_name.as_str()),
        })
}

fn update_in_place(
    package: &mut OpcPackage,
    current: &Revisions,
    value: &Revisions,
    conformance: RevisionConformance,
) -> Result<()> {
    for name in [&current.users_part_name, &current.headers_part_name] {
        let uri = PackURI::new(name).map_err(invalid)?;
        if incoming_references(package, &uri)? != 1 {
            return Err(invalid(format!(
                "revision part '{}' has unexpected incoming relationships",
                uri
            )));
        }
    }
    let current_conformance = revision_conformance(package)?
        .ok_or_else(|| invalid("revision owner has no relationship conformance"))?;
    let conformance_changed = current_conformance != conformance;
    let users_changed = conformance_changed || current.users != value.users;
    let headers_changed = conformance_changed || current.headers != value.headers;
    let current_logs: HashMap<_, _> = current
        .logs
        .iter()
        .map(|log| (log.relationship_id.as_str(), log))
        .collect();
    let value_logs: HashMap<_, _> = value
        .logs
        .iter()
        .map(|log| (log.relationship_id.as_str(), log))
        .collect();

    let users_xml = users_changed
        .then(|| write_revision_users(&value.users, conformance))
        .transpose()?;
    let headers_xml = headers_changed
        .then(|| write_revision_headers(&value.headers, conformance))
        .transpose()?;
    let mut log_xml = HashMap::new();
    for log in &value.logs {
        let changed = conformance_changed
            || current_logs
                .get(log.relationship_id.as_str())
                .is_none_or(|old| *old != log);
        if changed {
            log_xml.insert(
                log.relationship_id.clone(),
                write_revision_log(&log.log, conformance)?,
            );
        }
    }

    let current_log_ids: HashSet<_> = current_logs.keys().copied().collect();
    for log in &value.logs {
        if !current_log_ids.contains(log.relationship_id.as_str()) {
            let uri = PackURI::new(&log.part_name).map_err(invalid)?;
            package
                .validate_new_part_name(&uri)
                .map_err(|error| invalid(error.to_string()))?;
            let headers_part =
                package.get_part(&PackURI::new(&value.headers_part_name).map_err(invalid)?)?;
            if headers_part.rels().get(&log.relationship_id).is_some() {
                return Err(invalid("revision log relationship id already exists"));
            }
        }
    }
    for old in current
        .logs
        .iter()
        .filter(|log| !value_logs.contains_key(log.relationship_id.as_str()))
    {
        let uri = PackURI::new(&old.part_name).map_err(invalid)?;
        if incoming_references(package, &uri)? != 1 {
            return Err(invalid(format!(
                "revision log part '{}' has unexpected incoming relationships",
                uri
            )));
        }
    }

    let users_uri = PackURI::new(&value.users_part_name).map_err(invalid)?;
    let headers_uri = PackURI::new(&value.headers_part_name).map_err(invalid)?;
    if let Some(xml) = users_xml {
        package.get_part_mut(&users_uri)?.set_blob(xml);
    }
    if let Some(xml) = headers_xml {
        package.get_part_mut(&headers_uri)?.set_blob(xml);
    }
    let old_removed = current
        .logs
        .iter()
        .filter(|log| !value_logs.contains_key(log.relationship_id.as_str()))
        .map(|log| log.relationship_id.as_str())
        .collect::<HashSet<_>>();
    {
        let headers_part = package.get_part_mut(&headers_uri)?;
        for relationship_id in &old_removed {
            headers_part.rels_mut().remove(relationship_id);
        }
        if conformance_changed {
            for log in &value.logs {
                replace_relationship_type(
                    headers_part,
                    &log.relationship_id,
                    conformance.log_relationship(),
                )?;
            }
        }
    }
    for log in &value.logs {
        if let Some(xml) = log_xml.remove(&log.relationship_id) {
            let uri = PackURI::new(&log.part_name).map_err(invalid)?;
            if current_log_ids.contains(log.relationship_id.as_str()) {
                package.get_part_mut(&uri)?.set_blob(xml);
            } else {
                package.try_add_part(Box::new(BlobPart::new(uri.clone(), LOG_CT.into(), xml)))?;
                package
                    .get_part_mut(&headers_uri)?
                    .rels_mut()
                    .try_add_relationship(
                        conformance.log_relationship().into(),
                        uri.relative_ref(headers_uri.base_uri()),
                        log.relationship_id.clone(),
                        litchi_opc::TargetMode::Internal,
                    )?;
            }
        }
    }
    for old in current
        .logs
        .iter()
        .filter(|log| old_removed.contains(log.relationship_id.as_str()))
    {
        let uri = PackURI::new(&old.part_name).map_err(invalid)?;
        package.remove_part(&uri);
    }
    if conformance_changed {
        let workbook_uri = package.main_document_part()?.partname().clone();
        let workbook_part = package.get_part_mut(&workbook_uri)?;
        replace_relationship_type(
            workbook_part,
            &value.users_relationship_id,
            conformance.users_relationship(),
        )?;
        replace_relationship_type(
            workbook_part,
            &value.headers_relationship_id,
            conformance.headers_relationship(),
        )?;
    }
    if users_changed || headers_changed || current.logs != value.logs || conformance_changed {
        package.unsign();
    }
    validate_graph(package)
}

fn replace_relationship_type(
    source: &mut dyn Part,
    relationship_id: &str,
    relationship_type: &str,
) -> Result<()> {
    let relationship = source
        .rels()
        .get(relationship_id)
        .ok_or_else(|| invalid(format!("missing revision relationship '{relationship_id}'")))?;
    if relationship.reltype() == relationship_type {
        return Ok(());
    }
    let target = relationship.target_ref().to_owned();
    let mode = relationship.target_mode();
    source.rels_mut().remove(relationship_id);
    source.rels_mut().try_add_relationship(
        relationship_type.to_owned(),
        target,
        relationship_id.to_owned(),
        mode,
    )?;
    Ok(())
}

/// Restore exact source bytes and relationship metadata captured by a patch.
pub(crate) fn restore_snapshot_source(package: &mut OpcPackage, snapshot: &Snapshot) -> Result<()> {
    let workbook_uri = package.main_document_part()?.partname().clone();
    if workbook_uri.as_str() != snapshot.source().workbook_part_name()
        || package.main_document_part()?.content_type() != snapshot.source().workbook_content_type()
    {
        return Err(invalid("revision patch targets a different workbook"));
    }
    apply_relationship_set(
        package.get_part_mut(&workbook_uri)?.rels_mut(),
        snapshot.source().workbook_relationships(),
        Some(is_revision_relationship),
    )?;
    for part in snapshot.source().parts() {
        let uri = PackURI::new(part.part_name()).map_err(invalid)?;
        let target = package.get_part_mut(&uri)?;
        if target.content_type() != part.content_type() {
            return Err(invalid("revision patch content type changed"));
        }
        if target.blob() != part.bytes() {
            target.set_blob_shared(part.bytes_arc());
        }
        apply_relationship_set(target.rels_mut(), part.relationships(), None)?;
    }
    validate_graph(package)
}

fn apply_relationship_set(
    relationships: &mut litchi_opc::Relationships,
    expected: &[SourceRelationship],
    filter: Option<fn(&str) -> bool>,
) -> Result<()> {
    let expected_ids: HashSet<_> = expected
        .iter()
        .map(|relationship| relationship.id())
        .collect();
    let remove = relationships
        .iter()
        .filter(|relationship| filter.map_or(true, |f| f(relationship.reltype())))
        .filter(|relationship| !expected_ids.contains(relationship.r_id()))
        .map(|relationship| relationship.r_id().to_owned())
        .collect::<Vec<_>>();
    for relationship_id in remove {
        relationships.remove(&relationship_id);
    }
    for expected in expected {
        let same = relationships
            .get(expected.id())
            .is_some_and(|relationship| {
                relationship.reltype() == expected.relationship_type()
                    && relationship.target_ref() == expected.target()
                    && relationship.target_mode() == expected.mode()
            });
        if !same {
            if relationships.get(expected.id()).is_some() {
                relationships.remove(expected.id());
            }
            relationships.try_add_relationship(
                expected.relationship_type().to_owned(),
                expected.target().to_owned(),
                expected.id().to_owned(),
                expected.mode(),
            )?;
        }
    }
    Ok(())
}

/// Validate the complete revision relationship graph without mutating it.
pub(super) fn validate_graph(package: &OpcPackage) -> Result<()> {
    let value = load_workbook_revisions(package)?;
    if let Some(value) = value {
        validate_owner_references(package, &value)?;
    }
    Ok(())
}

fn validate_owner_references(package: &OpcPackage, value: &Revisions) -> Result<()> {
    let mut names = Vec::with_capacity(value.logs.len().saturating_add(2));
    names.push(value.users_part_name.as_str());
    names.push(value.headers_part_name.as_str());
    names.extend(value.logs.iter().map(|log| log.part_name.as_str()));
    for name in names {
        let uri = PackURI::new(name).map_err(invalid)?;
        if incoming_references(package, &uri)? != 1 {
            return Err(invalid(format!(
                "revision part '{}' has unexpected incoming relationships",
                uri
            )));
        }
    }
    Ok(())
}

fn validate_physical_value(value: &Revisions) -> Result<()> {
    validate_package(value)?;
    let users = PackURI::new(&value.users_part_name).map_err(invalid)?;
    let headers = PackURI::new(&value.headers_part_name).map_err(invalid)?;
    let mut names = HashSet::from([users.to_string(), headers.to_string()]);
    for (header, log) in value.headers.headers.iter().zip(&value.logs) {
        if header.relationship_id != log.relationship_id {
            return Err(invalid("revision header/log order does not match"));
        }
        let uri = PackURI::new(&log.part_name).map_err(invalid)?;
        if !names.insert(uri.to_string()) {
            return Err(invalid("revision part names must be unique"));
        }
    }
    Ok(())
}

fn incoming_references(package: &OpcPackage, target: &PackURI) -> Result<usize> {
    let mut count = 0;
    for relationship in package.rels().iter() {
        if !relationship.is_external()
            && relationship
                .target_partname()
                .is_ok_and(|candidate| same_uri(&candidate, target))
        {
            count += 1;
        }
    }
    for part in package.iter_parts() {
        for relationship in part.rels().iter() {
            if !relationship.is_external()
                && relationship
                    .target_partname()
                    .is_ok_and(|candidate| same_uri(&candidate, target))
            {
                count += 1;
            }
        }
    }
    Ok(count)
}

fn same_uri(a: &PackURI, b: &PackURI) -> bool {
    a.as_str().eq_ignore_ascii_case(b.as_str())
}

pub(crate) fn is_revision_relationship(value: &str) -> bool {
    is_users_rel(value) || is_headers_rel(value) || is_log_rel(value)
}

pub(crate) fn revision_conformance(package: &OpcPackage) -> Result<Option<RevisionConformance>> {
    let mut transitional = false;
    let mut strict = false;
    let mut observe = |relationship: &litchi_opc::Relationship| match relationship.reltype() {
        USERS_REL | HEADERS_REL | LOG_REL => transitional = true,
        STRICT_USERS_REL | STRICT_HEADERS_REL | STRICT_LOG_REL => strict = true,
        _ => {},
    };
    for relationship in package.rels().iter() {
        observe(relationship);
    }
    for part in package.iter_parts() {
        for relationship in part.rels().iter() {
            observe(relationship);
        }
    }
    if transitional && strict {
        return Err(invalid(
            "mixed transitional and strict revision relationships",
        ));
    }
    Ok(if strict {
        Some(RevisionConformance::Strict)
    } else if transitional {
        Some(RevisionConformance::Transitional)
    } else {
        None
    })
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
