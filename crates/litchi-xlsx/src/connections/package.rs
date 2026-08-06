//! Workbook package integration for the SpreadsheetML connections owner.

use super::codec::patch_connections_source;
use super::model::*;
use super::{codec, invalid};
use litchi_core::sheet::Result;
use std::collections::HashSet;
use std::sync::Arc;

use litchi_opc::{OpcPackage, PackURI, Part};

pub fn store_in_package(package: &mut OpcPackage, value: &Connections, strict: bool) -> Result<()> {
    store_in_package_with_query_table_validator(package, value, strict, query_table_connection_id)
}

/// Store connections while allowing the migration host to retain its complete
/// query-table parser for cross-part validation.
#[doc(hidden)]
pub fn store_in_package_with_query_table_validator<F>(
    package: &mut OpcPackage,
    value: &Connections,
    strict: bool,
    query_table_connection_id: F,
) -> Result<()>
where
    F: Fn(&[u8]) -> Result<u32>,
{
    let xml = value.to_xml(strict)?;
    validate_query_table_connection_ids(package, value, query_table_connection_id)?;
    let workbook_name = package.main_document_part()?.partname().clone();
    let existing = {
        let workbook = package.get_part(&workbook_name)?;
        let mut found = workbook.rels().iter().filter(|relationship| {
            matches!(
                relationship.reltype(),
                CONNECTIONS_RELATIONSHIP | STRICT_CONNECTIONS_RELATIONSHIP
            )
        });
        let first = found
            .next()
            .map(|relationship| {
                if relationship.is_external() {
                    return Err(invalid("connections relationship cannot be external"));
                }
                Ok((
                    relationship.r_id().to_string(),
                    relationship.target_partname()?,
                ))
            })
            .transpose()?;
        if found.next().is_some() {
            return Err(invalid("workbook has multiple connections relationships"));
        }
        first
    };
    if let Some((_, part_name)) = existing {
        let part = package.get_part(&part_name)?;
        if part.content_type() != CONNECTIONS_CONTENT_TYPE {
            return Err(invalid(
                "existing connections part has invalid content type",
            ));
        }
        package.get_part_mut(&part_name)?.set_blob(xml);
    } else {
        let part_name = next_connections_part_name(package)?;
        let relationship_id = next_connections_relationship_id(package, &workbook_name)?;
        package.try_add_part(Box::new(litchi_opc::part::BlobPart::new(
            part_name.clone(),
            CONNECTIONS_CONTENT_TYPE.into(),
            xml,
        )))?;
        package
            .get_part_mut(&workbook_name)?
            .rels_mut()
            .add_relationship(
                if strict {
                    STRICT_CONNECTIONS_RELATIONSHIP
                } else {
                    CONNECTIONS_RELATIONSHIP
                }
                .into(),
                part_name.relative_ref(workbook_name.base_uri()),
                relationship_id,
                false,
            );
    }
    package.unsign();
    Ok(())
}

pub fn remove_from_package(package: &mut OpcPackage) -> Result<bool> {
    if package
        .iter_parts()
        .any(|part| part.content_type() == QUERY_TABLE_CONTENT_TYPE)
    {
        return Err(invalid(
            "cannot remove connections while query-table parts remain",
        ));
    }
    let workbook_name = package.main_document_part()?.partname().clone();
    let relationship = package
        .get_part(&workbook_name)?
        .rels()
        .iter()
        .find(|relationship| {
            matches!(
                relationship.reltype(),
                CONNECTIONS_RELATIONSHIP | STRICT_CONNECTIONS_RELATIONSHIP
            )
        })
        .map(|relationship| {
            relationship
                .target_partname()
                .map(|part_name| (relationship.r_id().to_string(), part_name))
        })
        .transpose()?;
    let Some((relationship_id, part_name)) = relationship else {
        return Ok(false);
    };
    package
        .get_part_mut(&workbook_name)?
        .rels_mut()
        .remove(&relationship_id);
    if !package_part_is_referenced(package, &part_name) {
        package.remove_part(&part_name);
    }
    package.unsign();
    Ok(true)
}

fn validate_query_table_connection_ids<F>(
    package: &OpcPackage,
    value: &Connections,
    query_table_connection_id: F,
) -> Result<()>
where
    F: Fn(&[u8]) -> Result<u32>,
{
    let ids = value
        .connections
        .iter()
        .map(|connection| connection.id)
        .collect::<HashSet<_>>();
    for part in package
        .iter_parts()
        .filter(|part| part.content_type() == QUERY_TABLE_CONTENT_TYPE)
    {
        let connection_id = query_table_connection_id(part.blob())?;
        if !ids.contains(&connection_id) {
            return Err(invalid(format!(
                "query-table part '{}' references missing connection ID {}",
                part.partname(),
                connection_id
            )));
        }
    }
    Ok(())
}

fn query_table_connection_id(xml: &[u8]) -> Result<u32> {
    if xml.len() > 8 * 1024 * 1024 {
        return Err(invalid("query-table part exceeds 8 MiB"));
    }
    let processed = litchi_ooxml_common::mce::process_ooxml(xml)?;
    if processed.len() > 8 * 1024 * 1024 {
        return Err(invalid("processed query-table part exceeds 8 MiB"));
    }
    let root = codec::parse_dom(processed.as_ref())?;
    codec::expect(&root, "queryTable")?;
    let _name = codec::req(&root, "name")?;
    let connection_id = codec::u32req(&root, "connectionId")?;
    codec::only_unqualified(
        &root,
        &[
            "name",
            "headers",
            "rowNumbers",
            "disableRefresh",
            "backgroundRefresh",
            "firstBackgroundRefresh",
            "refreshOnLoad",
            "growShrinkType",
            "fillFormulas",
            "removeDataOnSave",
            "disableEdit",
            "preserveFormatting",
            "adjustColumnWidth",
            "intermediate",
            "connectionId",
            "autoFormatId",
            "applyNumberFormats",
            "applyBorderFormats",
            "applyFontFormats",
            "applyPatternFormats",
            "applyAlignmentFormats",
            "applyWidthHeightFormats",
        ],
    )?;
    codec::kids(&root)?;
    Ok(connection_id)
}

fn next_connections_part_name(package: &OpcPackage) -> Result<PackURI> {
    for suffix in 0..=65_536u32 {
        let name = if suffix == 0 {
            "/xl/connections.xml".into()
        } else {
            format!("/xl/connections{suffix}.xml")
        };
        let candidate = PackURI::new(&name)?;
        if package.get_part(&candidate).is_err() {
            return Ok(candidate);
        }
    }
    Err(invalid("no free connections part name"))
}

fn next_connections_relationship_id(package: &OpcPackage, workbook: &PackURI) -> Result<String> {
    let relationships = package.get_part(workbook)?.rels();
    for suffix in 1..=65_537u32 {
        let candidate = format!("rIdConnections{suffix}");
        if relationships.get(&candidate).is_none() {
            return Ok(candidate);
        }
    }
    Err(invalid("no free connections relationship ID"))
}

fn package_part_is_referenced(package: &OpcPackage, target: &PackURI) -> bool {
    package.iter_parts().any(|part| {
        part.rels().iter().any(|relationship| {
            !relationship.is_external()
                && relationship
                    .target_partname()
                    .is_ok_and(|name| name == *target)
        })
    }) || package.rels().iter().any(|relationship| {
        !relationship.is_external()
            && relationship
                .target_partname()
                .is_ok_and(|name| name == *target)
    })
}
pub fn load_from_package(package: &OpcPackage) -> Result<Option<Connections>> {
    let workbook = package.main_document_part()?;
    let mut found = workbook.rels().iter().filter(|x| {
        matches!(
            x.reltype(),
            CONNECTIONS_RELATIONSHIP | STRICT_CONNECTIONS_RELATIONSHIP
        )
    });
    let Some(rel) = found.next() else {
        return Ok(None);
    };
    if found.next().is_some() {
        return Err(invalid("workbook has multiple connections relationships"));
    }
    if rel.is_external() {
        return Err(invalid("connections relationship cannot be external"));
    }
    let uri: PackURI = rel.target_partname()?;
    let part = package.get_part(&uri)?;
    if part.content_type() != CONNECTIONS_CONTENT_TYPE {
        return Err(invalid(format!(
            "connections part '{uri}' has invalid content type '{}'",
            part.content_type()
        )));
    }
    if part.rels().iter().next().is_some() {
        return Err(invalid("connections part must not have relationships"));
    }
    Ok(Some(Connections::parse(part.blob())?))
}

/// Validate the workbook connection/query-table graph and return its typed
/// connection catalog. No query-table payload is interpreted beyond its
/// inert connection ID, and no target is opened or refreshed.
pub fn validate_graph(package: &OpcPackage) -> Result<Option<Connections>> {
    if package
        .rels()
        .iter()
        .any(|relationship| is_connections_relationship(relationship.reltype()))
    {
        return Err(invalid(
            "package root cannot source a workbook connections relationship",
        ));
    }
    let workbook = package.main_document_part()?;
    let owners = workbook
        .rels()
        .iter()
        .filter(|relationship| is_connections_relationship(relationship.reltype()))
        .collect::<Vec<_>>();
    if owners.len() > 1 {
        return Err(invalid("workbook has multiple connections relationships"));
    }
    let owner_target = owners
        .first()
        .map(|relationship| {
            if relationship.is_external() {
                return Err(invalid("connections relationship cannot be external"));
            }
            Ok(relationship.target_partname()?)
        })
        .transpose()?;
    if let Some(target) = owner_target.as_ref() {
        let part = package.get_part(target)?;
        if part.content_type() != CONNECTIONS_CONTENT_TYPE {
            return Err(invalid("connections relationship targets an invalid part"));
        }
        if part.rels().iter().next().is_some() {
            return Err(invalid("connections part must not have relationships"));
        }
    }
    for part in package
        .iter_parts()
        .filter(|part| part.content_type() == CONNECTIONS_CONTENT_TYPE)
    {
        if owner_target.as_ref() != Some(part.partname()) {
            return Err(invalid(format!(
                "connections part '{}' has no workbook owner",
                part.partname()
            )));
        }
    }

    let query_parts = package
        .iter_parts()
        .filter(|part| part.content_type() == QUERY_TABLE_CONTENT_TYPE)
        .collect::<Vec<_>>();
    let values = if owner_target.is_some() {
        Some(load_from_package(package)?.ok_or_else(|| invalid("connections owner disappeared"))?)
    } else {
        None
    };
    if !query_parts.is_empty() && values.is_none() {
        return Err(invalid(
            "query-table parts require a workbook connections part",
        ));
    }
    if let Some(values) = values.as_ref() {
        validate_query_table_connection_ids(package, values, query_table_connection_id)?;
    }
    for part in query_parts {
        if part.rels().iter().next().is_some() {
            return Err(invalid("query-table parts must not have relationships"));
        }
        let mut owners = 0usize;
        for candidate in package.iter_parts() {
            for relationship in candidate.rels().iter() {
                if relationship.is_external()
                    || !is_query_table_relationship(relationship.reltype())
                    || relationship.target_partname().ok().as_ref() != Some(part.partname())
                {
                    continue;
                }
                owners += 1;
            }
        }
        if owners != 1 {
            return Err(invalid(format!(
                "query-table part '{}' must have exactly one worksheet owner",
                part.partname()
            )));
        }
    }
    Ok(values)
}

fn is_connections_relationship(value: &str) -> bool {
    matches!(
        value,
        CONNECTIONS_RELATIONSHIP | STRICT_CONNECTIONS_RELATIONSHIP
    )
}

fn is_query_table_relationship(value: &str) -> bool {
    matches!(
        value,
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/queryTable"
            | "http://purl.oclc.org/ooxml/officeDocument/relationships/queryTable"
    )
}

/// Immutable source snapshot used by connection transactions and patches.
#[derive(Clone, Debug, PartialEq)]
pub struct Snapshot {
    connections: Option<Connections>,
    source: SourceState,
    conformance: Conformance,
}

impl Snapshot {
    pub fn load(package: &OpcPackage) -> Result<Self> {
        let connections = validate_graph(package)?;
        let source = SourceState::capture(package)?;
        let conformance = if let Some(part) = &source.connection {
            detect_conformance(part.bytes())
        } else {
            detect_conformance(&source.workbook_bytes)
        };
        Ok(Self {
            connections,
            source,
            conformance,
        })
    }

    pub fn read(package: &OpcPackage) -> Result<Self> {
        Self::load(package)
    }

    pub fn connections(&self) -> Option<&Connections> {
        self.connections.as_ref()
    }

    pub fn catalog(&self) -> Option<&Connections> {
        self.connections()
    }

    pub fn source_xml(&self) -> Option<&[u8]> {
        self.source.connection.as_ref().map(SourcePart::bytes)
    }

    pub fn query_table_xml(&self, part_uri: &PackURI) -> Option<&[u8]> {
        self.source
            .query_tables
            .iter()
            .find(|part| part.part_uri == *part_uri)
            .map(SourcePart::bytes)
    }

    pub fn query_table_parts(&self) -> impl Iterator<Item = &PackURI> {
        self.source.query_tables.iter().map(|part| &part.part_uri)
    }

    pub fn workbook_part_name(&self) -> &str {
        &self.source.workbook_part_name
    }

    pub fn conformance(&self) -> Conformance {
        self.conformance
    }

    pub fn is_empty(&self) -> bool {
        self.connections.is_none()
    }

    fn same_source(&self, other: &Self) -> bool {
        self.source == other.source
    }
}

#[derive(Clone, Debug, PartialEq)]
struct SourceState {
    workbook_part_name: String,
    workbook_content_type: String,
    workbook_bytes: Arc<Vec<u8>>,
    workbook_relationships: Vec<SourceRelationship>,
    root_relationships: Vec<SourceRelationship>,
    connection: Option<SourcePart>,
    query_tables: Vec<SourcePart>,
}

impl SourceState {
    fn capture(package: &OpcPackage) -> Result<Self> {
        let workbook = package.main_document_part()?;
        let mut workbook_relationships = workbook
            .rels()
            .iter()
            .map(SourceRelationship::from_relationship)
            .collect::<Vec<_>>();
        workbook_relationships.sort_by(|left, right| left.id.cmp(&right.id));
        let mut root_relationships = package
            .rels()
            .iter()
            .map(SourceRelationship::from_relationship)
            .collect::<Vec<_>>();
        root_relationships.sort_by(|left, right| left.id.cmp(&right.id));
        let connection_uri = workbook
            .rels()
            .iter()
            .find(|relationship| is_connections_relationship(relationship.reltype()))
            .map(|relationship| relationship.target_partname())
            .transpose()?;
        let connection = connection_uri
            .as_ref()
            .map(|uri| package.get_part(uri).map(SourcePart::from_part))
            .transpose()?;
        let mut query_tables = package
            .iter_parts()
            .filter(|part| part.content_type() == QUERY_TABLE_CONTENT_TYPE)
            .map(SourcePart::from_part)
            .collect::<Vec<_>>();
        query_tables.sort_by(|left, right| left.part_uri.as_str().cmp(right.part_uri.as_str()));
        Ok(Self {
            workbook_part_name: workbook.partname().to_string(),
            workbook_content_type: workbook.content_type().to_owned(),
            workbook_bytes: workbook.blob_arc(),
            workbook_relationships,
            root_relationships,
            connection,
            query_tables,
        })
    }

    fn connection_relationship(&self) -> Option<&SourceRelationship> {
        self.workbook_relationships
            .iter()
            .find(|relationship| is_connections_relationship(&relationship.relationship_type))
    }
}

#[derive(Clone, Debug, PartialEq)]
struct SourcePart {
    part_uri: PackURI,
    content_type: String,
    bytes: Arc<Vec<u8>>,
    relationships: Vec<SourceRelationship>,
}

impl SourcePart {
    fn from_part(part: &dyn Part) -> Self {
        let mut relationships = part
            .rels()
            .iter()
            .map(SourceRelationship::from_relationship)
            .collect::<Vec<_>>();
        relationships.sort_by(|left, right| left.id.cmp(&right.id));
        Self {
            part_uri: part.partname().clone(),
            content_type: part.content_type().to_owned(),
            bytes: part.blob_arc(),
            relationships,
        }
    }

    fn bytes(&self) -> &[u8] {
        self.bytes.as_slice()
    }
}

#[derive(Clone, Debug, PartialEq)]
struct SourceRelationship {
    id: String,
    relationship_type: String,
    target: String,
    external: bool,
}

impl SourceRelationship {
    fn from_relationship(relationship: &litchi_opc::Relationship) -> Self {
        Self {
            id: relationship.r_id().to_owned(),
            relationship_type: relationship.reltype().to_owned(),
            target: relationship.target_ref().to_owned(),
            external: relationship.is_external(),
        }
    }
}

fn detect_conformance(xml: &[u8]) -> Conformance {
    if xml
        .windows(STRICT_NAMESPACE.len())
        .any(|window| window == STRICT_NAMESPACE.as_bytes())
    {
        Conformance::Strict
    } else {
        Conformance::Transitional
    }
}

/// Failure-atomic edits over the workbook's inert connection catalog.
pub struct Transaction<'a> {
    target: &'a mut OpcPackage,
    before: Snapshot,
    draft: Option<Connections>,
    strict: bool,
}

impl<'a> Transaction<'a> {
    pub fn new(target: &'a mut OpcPackage) -> Result<Self> {
        let before = Snapshot::load(target)?;
        let strict = before.conformance.strict();
        Ok(Self {
            draft: before.connections.clone(),
            target,
            before,
            strict,
        })
    }

    pub fn before(&self) -> &Snapshot {
        &self.before
    }

    pub fn connections(&self) -> Option<&Connections> {
        self.draft.as_ref()
    }

    pub fn replace(&mut self, value: Option<Connections>) -> Result<bool> {
        validate_draft(self.target, value.as_ref())?;
        if self.draft == value {
            return Ok(false);
        }
        self.draft = value;
        Ok(true)
    }

    pub fn edit(
        &mut self,
        id: u32,
        edit: impl FnOnce(&mut Connection) -> Result<()>,
    ) -> Result<bool> {
        let mut draft = self
            .draft
            .clone()
            .ok_or_else(|| invalid("cannot edit an absent connections part"))?;
        let connection = draft
            .connections
            .iter_mut()
            .find(|connection| connection.id == id)
            .ok_or_else(|| invalid(format!("connection ID {id} was not found")))?;
        edit(connection)?;
        if connection.id != id {
            return Err(invalid("connection ID is immutable inside a transaction"));
        }
        validate_draft(self.target, Some(&draft))?;
        if self.draft.as_ref() == Some(&draft) {
            return Ok(false);
        }
        self.draft = Some(draft);
        Ok(true)
    }

    pub fn set(&mut self, connection: Connection) -> Result<bool> {
        let mut draft = self.draft.clone().unwrap_or(Connections {
            connections: Vec::new(),
        });
        if let Some(existing) = draft
            .connections
            .iter_mut()
            .find(|existing| existing.id == connection.id)
        {
            if *existing == connection {
                return Ok(false);
            }
            *existing = connection;
        } else {
            draft.add(connection)?;
        }
        validate_draft(self.target, Some(&draft))?;
        self.draft = Some(draft);
        Ok(true)
    }

    pub fn remove(&mut self, id: u32) -> Result<Option<Connection>> {
        let mut draft = self
            .draft
            .clone()
            .ok_or_else(|| invalid("cannot remove from an absent connections part"))?;
        let Some(index) = draft
            .connections
            .iter()
            .position(|connection| connection.id == id)
        else {
            return Ok(None);
        };
        let removed = draft.connections.remove(index);
        if draft.connections.is_empty() {
            validate_draft(self.target, None)?;
            self.draft = None;
        } else {
            validate_draft(self.target, Some(&draft))?;
            self.draft = Some(draft);
        }
        Ok(Some(removed))
    }

    pub fn is_changed(&self) -> bool {
        self.before.connections.as_ref() != self.draft.as_ref()
    }

    pub fn commit(self) -> Result<Commit> {
        if !self.is_changed() {
            let patch = Patch::new(self.before.clone(), self.before.clone());
            return Ok(Commit::new(self.before, patch, false));
        }
        let current = Snapshot::load(self.target)?;
        if !current.same_source(&self.before) {
            return Err(invalid("connections transaction source is stale"));
        }
        let mut candidate = self.target.clone();
        apply_connections(
            &mut candidate,
            self.before.connections.as_ref(),
            self.draft.as_ref(),
            self.strict,
        )?;
        let snapshot = Snapshot::load(&candidate)?;
        if snapshot.connections.as_ref() != self.draft.as_ref() {
            return Err(invalid("connection publication changed the staged model"));
        }
        let patch = Patch::new(self.before, snapshot.clone());
        *self.target = candidate;
        Ok(Commit::new(snapshot, patch, true))
    }
}

/// A reversible source-checked package edit.
#[derive(Clone, Debug, PartialEq)]
pub struct Patch {
    before: Snapshot,
    after: Snapshot,
}

impl Patch {
    fn new(before: Snapshot, after: Snapshot) -> Self {
        Self { before, after }
    }

    pub fn before(&self) -> &Snapshot {
        &self.before
    }

    pub fn after(&self) -> &Snapshot {
        &self.after
    }

    pub fn is_empty(&self) -> bool {
        self.before.same_source(&self.after)
    }

    pub fn inverse(&self) -> Self {
        Self {
            before: self.after.clone(),
            after: self.before.clone(),
        }
    }

    pub fn apply(&self, target: &mut OpcPackage) -> Result<()> {
        let current = Snapshot::load(target)?;
        if !current.same_source(&self.before) {
            return Err(invalid("connections patch source is stale"));
        }
        if self.is_empty() {
            return Ok(());
        }
        let mut candidate = target.clone();
        restore_snapshot(&mut candidate, &self.after)?;
        let resulting = Snapshot::load(&candidate)?;
        if !resulting.same_source(&self.after) {
            return Err(invalid("connections patch publication changed its source"));
        }
        *target = candidate;
        Ok(())
    }
}

#[derive(Debug)]
pub struct Commit {
    snapshot: Snapshot,
    patch: Patch,
    changed: bool,
}

impl Commit {
    fn new(snapshot: Snapshot, patch: Patch, changed: bool) -> Self {
        Self {
            snapshot,
            patch,
            changed,
        }
    }

    pub fn changed(&self) -> bool {
        self.changed
    }

    pub fn snapshot(&self) -> &Snapshot {
        &self.snapshot
    }

    pub fn patch(&self) -> &Patch {
        &self.patch
    }

    pub fn into_parts(self) -> (Snapshot, Patch) {
        (self.snapshot, self.patch)
    }
}

fn validate_draft(package: &OpcPackage, value: Option<&Connections>) -> Result<()> {
    if let Some(value) = value {
        value.to_xml(false)?;
        validate_query_table_connection_ids(package, value, query_table_connection_id)?;
    } else if package
        .iter_parts()
        .any(|part| part.content_type() == QUERY_TABLE_CONTENT_TYPE)
    {
        return Err(invalid(
            "cannot remove connections while query-table parts remain",
        ));
    }
    Ok(())
}

fn apply_connections(
    package: &mut OpcPackage,
    before: Option<&Connections>,
    after: Option<&Connections>,
    strict: bool,
) -> Result<()> {
    validate_draft(package, after)?;
    let workbook_name = package.main_document_part()?.partname().clone();
    let owner = package
        .get_part(&workbook_name)?
        .rels()
        .iter()
        .find(|relationship| is_connections_relationship(relationship.reltype()))
        .map(|relationship| {
            relationship
                .target_partname()
                .map(|target| (relationship.r_id().to_owned(), target))
        })
        .transpose()?;
    match (owner, after) {
        (Some((relationship_id, part_name)), Some(after)) => {
            let source = package.get_part(&part_name)?.blob().to_vec();
            let updated = if let Some(before) = before {
                patch_connections_source(&source, before, after, strict)?
            } else {
                after.to_xml(strict)?
            };
            package.get_part_mut(&part_name)?.set_blob(updated);
            if package.get_part(&part_name)?.rels().iter().next().is_some() {
                return Err(invalid("connections part must not have relationships"));
            }
            let _ = relationship_id;
        },
        (None, Some(after)) => {
            let part_name = next_connections_part_name(package)?;
            let relationship_id = next_connections_relationship_id(package, &workbook_name)?;
            package.try_add_part(Box::new(litchi_opc::part::BlobPart::new(
                part_name.clone(),
                CONNECTIONS_CONTENT_TYPE.into(),
                after.to_xml(strict)?,
            )))?;
            package
                .get_part_mut(&workbook_name)?
                .rels_mut()
                .add_relationship(
                    if strict {
                        STRICT_CONNECTIONS_RELATIONSHIP
                    } else {
                        CONNECTIONS_RELATIONSHIP
                    }
                    .into(),
                    part_name.relative_ref(workbook_name.base_uri()),
                    relationship_id,
                    false,
                );
        },
        (Some((relationship_id, part_name)), None) => {
            package
                .get_part_mut(&workbook_name)?
                .rels_mut()
                .remove(&relationship_id);
            if !package_part_is_referenced(package, &part_name) {
                package.remove_part(&part_name);
            }
        },
        (None, None) => {},
    }
    package.unsign();
    validate_graph(package).map(|_| ())
}

fn restore_snapshot(package: &mut OpcPackage, snapshot: &Snapshot) -> Result<()> {
    let workbook_name = package.main_document_part()?.partname().clone();
    let existing = package
        .get_part(&workbook_name)?
        .rels()
        .iter()
        .find(|relationship| is_connections_relationship(relationship.reltype()))
        .map(|relationship| {
            relationship
                .target_partname()
                .map(|target| (relationship.r_id().to_owned(), target))
        })
        .transpose()?;
    if let Some((relationship_id, part_name)) = existing {
        package
            .get_part_mut(&workbook_name)?
            .rels_mut()
            .remove(&relationship_id);
        if !package_part_is_referenced(package, &part_name) {
            package.remove_part(&part_name);
        }
    }
    if let Some(part) = &snapshot.source.connection {
        package.try_add_part(Box::new(litchi_opc::part::BlobPart::new(
            part.part_uri.clone(),
            part.content_type.clone(),
            part.bytes().to_vec(),
        )))?;
        let relationship = snapshot
            .source
            .connection_relationship()
            .ok_or_else(|| invalid("connection snapshot is missing its workbook relationship"))?;
        package
            .get_part_mut(&workbook_name)?
            .rels_mut()
            .add_relationship(
                relationship.relationship_type.clone(),
                relationship.target.clone(),
                relationship.id.clone(),
                relationship.external,
            );
    }
    validate_graph(package).map(|_| ())
}
