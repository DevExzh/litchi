//! OPC graph ownership for XLSB shared-workbook metadata.

use std::collections::HashSet;

use litchi_opc::constants::content_type;
use litchi_opc::{BlobPart, OpcPackage, PackURI, Part};

use crate::package::error::{Error, Result};

use super::codec::{parse_headers, parse_log, parse_users, write_headers, write_log, write_users};
use super::model::Catalog;
use super::{invalid, validation};

/// User-names part content type from `[MS-XLSB]` section 2.1.7.55.
pub const USERS_CONTENT_TYPE: &str = "application/vnd.ms-excel.userNames";
/// Revision-headers part content type from `[MS-XLSB]` section 2.1.7.43.
pub const HEADERS_CONTENT_TYPE: &str = "application/vnd.ms-excel.revisionHeaders";
/// Revision-log part content type from `[MS-XLSB]` section 2.1.7.44.
pub const LOG_CONTENT_TYPE: &str = "application/vnd.ms-excel.revisionLog";
/// Workbook-to-user-names relationship type.
pub const USERS_RELATIONSHIP: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/usernames";
/// Workbook-to-revision-headers relationship type.
pub const HEADERS_RELATIONSHIP: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/revisionHeaders";
/// Revision-headers-to-revision-log relationship type.
pub const LOG_RELATIONSHIP: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/revisionLog";
const STRICT_USERS_RELATIONSHIP: &str =
    "http://purl.oclc.org/ooxml/officeDocument/relationships/usernames";
const STRICT_HEADERS_RELATIONSHIP: &str =
    "http://purl.oclc.org/ooxml/officeDocument/relationships/revisionHeaders";
const STRICT_LOG_RELATIONSHIP: &str =
    "http://purl.oclc.org/ooxml/officeDocument/relationships/revisionLog";

const USERS_PART_TEMPLATE: &str = "/xl/sharedWorkbook/userNames%d.bin";
const HEADERS_PART_TEMPLATE: &str = "/xl/sharedWorkbook/revisionHeaders%d.bin";
const LOG_PART_TEMPLATE: &str = "/xl/sharedWorkbook/revisionLog%d.bin";

/// Load and validate the complete shared-workbook metadata graph.
pub fn load(package: &OpcPackage) -> Result<Catalog> {
    let workbook = package.main_document_part()?;
    require_workbook(workbook)?;
    validate_sources(package, workbook.partname())?;

    let user_relationships = relationships_of(workbook, is_users_relationship);
    let header_relationships = relationships_of(workbook, is_headers_relationship);
    if user_relationships.len() > 1 || header_relationships.len() > 1 {
        return Err(invalid(
            "workbook has duplicate shared-workbook relationships",
        ));
    }
    if user_relationships.is_empty() && header_relationships.is_empty() {
        reject_orphans(package, None, None, &HashSet::new())?;
        return Ok(Catalog::empty());
    }
    if user_relationships.len() != 1 || header_relationships.len() != 1 {
        return Err(invalid(
            "userNames and revisionHeaders relationships must occur together",
        ));
    }

    let user_relationship = user_relationships[0];
    let header_relationship = header_relationships[0];
    if user_relationship.is_external() || header_relationship.is_external() {
        return Err(invalid("shared-workbook relationships must be internal"));
    }
    let user_uri = user_relationship.target_partname()?;
    let header_uri = header_relationship.target_partname()?;
    let user_part = package.get_part(&user_uri)?;
    let header_part = package.get_part(&header_uri)?;
    require_content_type(user_part, USERS_CONTENT_TYPE)?;
    require_content_type(header_part, HEADERS_CONTENT_TYPE)?;
    require_no_relationships(user_part, "userNames")?;

    let mut users = parse_users(user_part.blob())?;
    users.relationship_id = user_relationship.r_id().to_string();
    users.relationship_type = user_relationship.reltype().to_string();
    users.part_name = user_uri.to_string();

    let mut headers = parse_headers(header_part.blob())?;
    headers.relationship_id = header_relationship.r_id().to_string();
    headers.relationship_type = header_relationship.reltype().to_string();
    headers.part_name = header_uri.to_string();

    let log_relationships: Vec<_> = header_part
        .rels()
        .iter()
        .filter(|relationship| is_log_relationship(relationship.reltype()))
        .collect();
    if log_relationships.len() != headers.headers.len()
        || header_part
            .rels()
            .iter()
            .any(|relationship| !is_log_relationship(relationship.reltype()))
    {
        return Err(invalid(
            "revision-headers relationships do not match BrtRRHeader records",
        ));
    }

    let mut logs = Vec::with_capacity(headers.headers.len());
    let mut log_targets = HashSet::new();
    for header in &headers.headers {
        let relationship = header_part
            .rels()
            .get(&header.relationship_id)
            .ok_or_else(|| invalid("BrtRRHeader relationship ID has no revision log"))?;
        if !is_log_relationship(relationship.reltype()) || relationship.is_external() {
            return Err(invalid("revision-log relationship is invalid"));
        }
        let uri = relationship.target_partname()?;
        if !log_targets.insert(uri.to_string()) {
            return Err(invalid(
                "multiple BrtRRHeader records target one revision log",
            ));
        }
        let part = package.get_part(&uri)?;
        require_content_type(part, LOG_CONTENT_TYPE)?;
        require_no_relationships(part, "revisionLog")?;
        let mut log = parse_log(part.blob())?;
        log.relationship_id = relationship.r_id().to_string();
        log.relationship_type = relationship.reltype().to_string();
        log.part_name = uri.to_string();
        logs.push(log);
    }
    reject_orphans(
        package,
        Some(user_uri.as_str()),
        Some(header_uri.as_str()),
        &log_targets,
    )?;

    let catalog = Catalog {
        users: Some(users),
        headers: Some(headers),
        logs,
    };
    validation::validate_catalog(&catalog)?;
    Ok(catalog)
}

/// Atomically replace the shared-workbook owner on a cloned OPC graph.
pub fn store(package: &mut OpcPackage, catalog: &Catalog) -> Result<()> {
    let _ = load(package)?;
    let mut candidate = package.clone();
    store_inner(&mut candidate, catalog)?;
    let _ = load(&candidate)?;
    *package = candidate;
    Ok(())
}

fn store_inner(package: &mut OpcPackage, source: &Catalog) -> Result<()> {
    let mut catalog = source.clone();
    normalize_identity(package, &mut catalog)?;
    validation::validate_catalog(&catalog)?;
    remove_owner(package)?;
    let Some(users) = catalog.users.as_ref() else {
        return Ok(());
    };
    let headers = catalog
        .headers
        .as_ref()
        .ok_or_else(|| invalid("shared-workbook headers are missing"))?;

    let workbook_uri = package.main_document_part()?.partname().clone();
    package.try_add_part(Box::new(BlobPart::new(
        PackURI::new(&users.part_name).map_err(Error::InvalidUri)?,
        USERS_CONTENT_TYPE.to_string(),
        write_users(users)?,
    )))?;
    package.try_add_part(Box::new(BlobPart::new(
        PackURI::new(&headers.part_name).map_err(Error::InvalidUri)?,
        HEADERS_CONTENT_TYPE.to_string(),
        write_headers(headers)?,
    )))?;
    {
        let workbook = package.get_part_mut(&workbook_uri)?;
        workbook.rels_mut().add_relationship(
            users.relationship_type.clone(),
            relative_target(&workbook_uri, &users.part_name)?,
            users.relationship_id.clone(),
            false,
        );
        workbook.rels_mut().add_relationship(
            headers.relationship_type.clone(),
            relative_target(&workbook_uri, &headers.part_name)?,
            headers.relationship_id.clone(),
            false,
        );
    }
    let header_uri = PackURI::new(&headers.part_name).map_err(Error::InvalidUri)?;
    for log in &catalog.logs {
        let log_uri = PackURI::new(&log.part_name).map_err(Error::InvalidUri)?;
        package.try_add_part(Box::new(BlobPart::new(
            log_uri.clone(),
            LOG_CONTENT_TYPE.to_string(),
            write_log(log)?,
        )))?;
        package
            .get_part_mut(&header_uri)?
            .rels_mut()
            .add_relationship(
                log.relationship_type.clone(),
                relative_target(&header_uri, &log.part_name)?,
                log.relationship_id.clone(),
                false,
            );
    }
    Ok(())
}

fn normalize_identity(package: &OpcPackage, catalog: &mut Catalog) -> Result<()> {
    match (&mut catalog.users, &mut catalog.headers) {
        (None, None) if catalog.logs.is_empty() => return Ok(()),
        (Some(_), Some(_)) => {},
        _ => {
            return Err(invalid(
                "shared-workbook users and headers must occur together",
            ));
        },
    }
    let workbook = package.main_document_part()?;
    let mut used = workbook
        .rels()
        .iter()
        .map(|relationship| relationship.r_id().to_string())
        .collect::<HashSet<_>>();
    let users = catalog
        .users
        .as_mut()
        .ok_or_else(|| invalid("shared-workbook users are missing"))?;
    let headers = catalog
        .headers
        .as_mut()
        .ok_or_else(|| invalid("shared-workbook headers are missing"))?;
    if users.relationship_id.is_empty() {
        users.relationship_id = next_relationship_id(&used, "rIdUsers");
    }
    used.insert(users.relationship_id.clone());
    if headers.relationship_id.is_empty() {
        headers.relationship_id = next_relationship_id(&used, "rIdRevisionHeaders");
    }
    used.insert(headers.relationship_id.clone());
    if users.relationship_type.is_empty() {
        users.relationship_type = USERS_RELATIONSHIP.to_string();
    }
    if headers.relationship_type.is_empty() {
        headers.relationship_type = HEADERS_RELATIONSHIP.to_string();
    }
    if users.part_name.is_empty() {
        users.part_name = package.next_partname(USERS_PART_TEMPLATE)?.to_string();
    }
    if headers.part_name.is_empty() {
        headers.part_name = package.next_partname(HEADERS_PART_TEMPLATE)?.to_string();
    }
    if catalog.logs.len() != headers.headers.len() {
        return Err(invalid("revision log/header count mismatch"));
    }
    let mut names = HashSet::new();
    for (index, log) in catalog.logs.iter_mut().enumerate() {
        let header = headers
            .headers
            .get(index)
            .ok_or_else(|| invalid("revision log/header count mismatch"))?;
        if log.relationship_id.is_empty() {
            log.relationship_id = header.relationship_id.clone();
        }
        if log.relationship_id != header.relationship_id {
            return Err(invalid(
                "revision log relationship does not match its header",
            ));
        }
        if log.relationship_type.is_empty() {
            log.relationship_type = LOG_RELATIONSHIP.to_string();
        }
        if log.part_name.is_empty() {
            log.part_name = package.next_partname(LOG_PART_TEMPLATE)?.to_string();
        }
        if !names.insert(log.part_name.clone()) {
            return Err(invalid("revision log part names must be unique"));
        }
    }
    Ok(())
}

fn remove_owner(package: &mut OpcPackage) -> Result<()> {
    let special_relationships: Vec<_> = package
        .iter_parts()
        .flat_map(|part| {
            part.rels()
                .iter()
                .filter(|relationship| is_special_relationship(relationship.reltype()))
                .map(|relationship| (part.partname().clone(), relationship.r_id().to_string()))
                .collect::<Vec<_>>()
        })
        .collect();
    for (source, relationship_id) in special_relationships {
        package
            .get_part_mut(&source)?
            .rels_mut()
            .remove(&relationship_id);
    }
    let owner_parts: Vec<_> = package
        .iter_parts()
        .filter(|part| is_owner_content_type(part.content_type()))
        .map(|part| part.partname().clone())
        .collect();
    for part in owner_parts {
        package.remove_part(&part);
    }
    Ok(())
}

fn validate_sources(package: &OpcPackage, workbook: &PackURI) -> Result<()> {
    if package
        .rels()
        .iter()
        .any(|relationship| is_special_relationship(relationship.reltype()))
    {
        return Err(invalid(
            "shared-workbook relationships cannot originate at package root",
        ));
    }
    for part in package.iter_parts() {
        for relationship in part.rels().iter() {
            let valid_source = if is_users_relationship(relationship.reltype())
                || is_headers_relationship(relationship.reltype())
            {
                part.partname() == workbook
            } else if is_log_relationship(relationship.reltype()) {
                part.content_type() == HEADERS_CONTENT_TYPE
            } else {
                true
            };
            if !valid_source {
                return Err(invalid(
                    "shared-workbook relationship has an invalid source",
                ));
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
    for part in package.iter_parts() {
        let name = part.partname().as_str();
        if part.content_type() == USERS_CONTENT_TYPE && Some(name) != users
            || part.content_type() == HEADERS_CONTENT_TYPE && Some(name) != headers
            || part.content_type() == LOG_CONTENT_TYPE && !logs.contains(name)
        {
            return Err(invalid(
                "shared-workbook package contains an orphan owner part",
            ));
        }
    }
    Ok(())
}

fn relationships_of<'a>(
    part: &'a dyn Part,
    predicate: impl Fn(&str) -> bool,
) -> Vec<&'a litchi_opc::Relationship> {
    part.rels()
        .iter()
        .filter(|relationship| predicate(relationship.reltype()))
        .collect()
}

fn require_workbook(part: &dyn Part) -> Result<()> {
    if part.content_type() == content_type::XLSB_BIN {
        Ok(())
    } else {
        Err(invalid("main document is not an XLSB workbook"))
    }
}

fn require_content_type(part: &dyn Part, expected: &str) -> Result<()> {
    if part.content_type() == expected {
        Ok(())
    } else {
        Err(Error::InvalidContentType {
            expected: expected.to_string(),
            got: part.content_type().to_string(),
        })
    }
}

fn require_no_relationships(part: &dyn Part, label: &str) -> Result<()> {
    if part.rels().is_empty() {
        Ok(())
    } else {
        Err(invalid(format!("{label} part must not have relationships")))
    }
}

fn relative_target(source: &PackURI, target: &str) -> Result<String> {
    let target = PackURI::new(target).map_err(Error::InvalidUri)?;
    Ok(target.relative_ref(source.base_uri()))
}

fn next_relationship_id(used: &HashSet<String>, preferred: &str) -> String {
    if !used.contains(preferred) {
        return preferred.to_string();
    }
    let mut index = 1_u32;
    loop {
        let candidate = format!("{preferred}{index}");
        if !used.contains(&candidate) {
            return candidate;
        }
        index = index.saturating_add(1);
    }
}

fn is_owner_content_type(value: &str) -> bool {
    matches!(
        value,
        USERS_CONTENT_TYPE | HEADERS_CONTENT_TYPE | LOG_CONTENT_TYPE
    )
}

fn is_special_relationship(value: &str) -> bool {
    is_users_relationship(value) || is_headers_relationship(value) || is_log_relationship(value)
}

fn is_users_relationship(value: &str) -> bool {
    matches!(value, USERS_RELATIONSHIP | STRICT_USERS_RELATIONSHIP)
}

fn is_headers_relationship(value: &str) -> bool {
    matches!(value, HEADERS_RELATIONSHIP | STRICT_HEADERS_RELATIONSHIP)
}

fn is_log_relationship(value: &str) -> bool {
    matches!(value, LOG_RELATIONSHIP | STRICT_LOG_RELATIONSHIP)
}
