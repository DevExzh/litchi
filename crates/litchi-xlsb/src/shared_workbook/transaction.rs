//! Source-checked snapshot edits for shared-workbook revision metadata.
//!
//! The transaction owns metadata and opaque BIFF12 record envelopes only. It
//! never coordinates users, acquires locks, or applies revision records.

use std::collections::HashSet;
use std::sync::Arc;

use litchi_opc::{OpcPackage, TargetMode};

use super::model::{
    Catalog, Header, Info, RawRecord, RevisionHeaders, RevisionLog, User, UserNames,
};
use super::{Result, package, validation};
use crate::package::error::Error;

/// An immutable shared-workbook metadata snapshot bound to package source.
#[derive(Debug, Clone)]
pub struct Snapshot {
    catalog: Catalog,
    source: Option<Arc<SourceState>>,
}

impl Snapshot {
    /// Read and validate the shared-workbook owner from an OPC package.
    pub fn read(package: &OpcPackage) -> Result<Self> {
        let catalog = package::load(package)?;
        let source = SourceState::capture(package, &catalog)?;
        Ok(Self {
            catalog,
            source: Some(Arc::new(source)),
        })
    }

    /// Borrow the package-neutral metadata graph.
    #[must_use]
    pub const fn catalog(&self) -> &Catalog {
        &self.catalog
    }

    /// Borrow the current-user log, when the workbook is shared.
    #[must_use]
    pub const fn users(&self) -> Option<&UserNames> {
        self.catalog.users.as_ref()
    }

    /// Borrow the revision-header metadata, when present.
    #[must_use]
    pub const fn headers(&self) -> Option<&RevisionHeaders> {
        self.catalog.headers.as_ref()
    }

    /// Borrow revision logs in the same order as their headers.
    #[must_use]
    pub fn logs(&self) -> &[RevisionLog] {
        &self.catalog.logs
    }

    /// Borrow exact source bytes and relationship context for this owner.
    #[must_use]
    pub fn source_parts(&self) -> &[SourcePart] {
        self.source
            .as_deref()
            .map_or(&[], |source| source.parts.as_slice())
    }

    /// Whether this snapshot can produce a source-checked package patch.
    #[must_use]
    pub const fn is_source_bound(&self) -> bool {
        self.source.is_some()
    }

    /// Whether the package does not contain shared-workbook metadata.
    #[must_use]
    pub fn is_empty(&self) -> bool {
        self.catalog == Catalog::empty()
    }

    /// Start a detached transaction against this snapshot.
    #[must_use]
    pub fn edit(&self) -> Transaction {
        Transaction {
            base: self.clone(),
            catalog: self.catalog.clone(),
        }
    }
}

impl PartialEq for Snapshot {
    fn eq(&self, other: &Self) -> bool {
        self.catalog == other.catalog && self.source == other.source
    }
}

impl Eq for Snapshot {}

/// A detached, failure-atomic metadata transaction.
#[derive(Debug, Clone)]
pub struct Transaction {
    base: Snapshot,
    catalog: Catalog,
}

impl Transaction {
    /// Borrow the currently staged metadata graph.
    #[must_use]
    pub const fn catalog(&self) -> &Catalog {
        &self.catalog
    }

    /// Replace the complete metadata graph after validating its local shape.
    pub fn replace(&mut self, catalog: Catalog) -> Result<()> {
        self.apply_candidate(catalog)
    }

    /// Replace the current-user metadata.
    pub fn set_users(&mut self, users: Option<UserNames>) -> Result<()> {
        let mut candidate = self.catalog.clone();
        candidate.users = users;
        self.apply_candidate(candidate)
    }

    /// Replace revision-header metadata while retaining the existing log list.
    pub fn set_headers(&mut self, headers: Option<RevisionHeaders>) -> Result<()> {
        let mut candidate = self.catalog.clone();
        candidate.headers = headers;
        if candidate.headers.is_none() {
            candidate.logs.clear();
        }
        self.apply_candidate(candidate)
    }

    /// Replace all revision logs in header order.
    pub fn set_logs(&mut self, logs: Vec<RevisionLog>) -> Result<()> {
        let mut candidate = self.catalog.clone();
        candidate.logs = logs;
        self.apply_candidate(candidate)
    }

    /// Replace the typed `BrtInfo` metadata.
    pub fn set_info(&mut self, info: Info) -> Result<()> {
        let mut candidate = self.catalog.clone();
        let headers = candidate
            .headers
            .as_mut()
            .ok_or_else(|| super::invalid("cannot set BrtInfo without revision headers"))?;
        headers.info = info;
        self.apply_candidate(candidate)
    }

    /// Add or replace a currently open user by stable id or GUID.
    pub fn upsert_user(&mut self, user: User) -> Result<()> {
        let mut candidate = self.catalog.clone();
        let users = candidate
            .users
            .as_mut()
            .ok_or_else(|| super::invalid("cannot edit users without a user-names part"))?;
        if let Some(index) = users
            .users
            .iter()
            .position(|existing| existing.id == user.id || existing.guid == user.guid)
        {
            if users.users.iter().enumerate().any(|(other, existing)| {
                other != index && (existing.id == user.id || existing.guid == user.guid)
            }) {
                return Err(super::invalid(
                    "BrtUsr id and GUID identify different users",
                ));
            }
            users.users[index] = user;
        } else {
            users.users.push(user);
        }
        self.apply_candidate(candidate)
    }

    /// Remove a currently open user by id.
    pub fn remove_user(&mut self, id: u32) -> Result<Option<User>> {
        let mut candidate = self.catalog.clone();
        let Some(users) = candidate.users.as_mut() else {
            return Ok(None);
        };
        let Some(index) = users.users.iter().position(|user| user.id == id) else {
            return Ok(None);
        };
        let user = users.users.remove(index);
        self.apply_candidate(candidate)?;
        Ok(Some(user))
    }

    /// Replace one revision header and its corresponding opaque log.
    pub fn set_revision(&mut self, index: usize, header: Header, log: RevisionLog) -> Result<()> {
        let mut candidate = self.catalog.clone();
        let headers = candidate
            .headers
            .as_mut()
            .ok_or_else(|| super::invalid("cannot edit a revision without headers"))?;
        let existing = headers
            .headers
            .get_mut(index)
            .ok_or_else(|| super::invalid("revision-header index is out of range"))?;
        *existing = header;
        let existing_log = candidate
            .logs
            .get_mut(index)
            .ok_or_else(|| super::invalid("revision-log index is out of range"))?;
        *existing_log = log;
        self.apply_candidate(candidate)
    }

    /// Add a revision header and its opaque log at the end of the owner.
    pub fn push_revision(&mut self, header: Header, log: RevisionLog) -> Result<()> {
        let mut candidate = self.catalog.clone();
        candidate
            .headers
            .as_mut()
            .ok_or_else(|| super::invalid("cannot add a revision without headers"))?
            .headers
            .push(header);
        candidate.logs.push(log);
        self.apply_candidate(candidate)
    }

    /// Remove one revision header and its corresponding log.
    pub fn remove_revision(&mut self, index: usize) -> Result<Option<(Header, RevisionLog)>> {
        let mut candidate = self.catalog.clone();
        let headers = candidate
            .headers
            .as_mut()
            .ok_or_else(|| super::invalid("cannot remove a revision without headers"))?;
        if index >= headers.headers.len() || index >= candidate.logs.len() {
            return Ok(None);
        }
        let header = headers.headers.remove(index);
        let log = candidate.logs.remove(index);
        self.apply_candidate(candidate)?;
        Ok(Some((header, log)))
    }

    /// Replace the opaque records of one revision log.
    pub fn set_log_records(&mut self, index: usize, records: Vec<RawRecord>) -> Result<()> {
        let mut candidate = self.catalog.clone();
        let log = candidate
            .logs
            .get_mut(index)
            .ok_or_else(|| super::invalid("revision-log index is out of range"))?;
        log.records = records;
        self.apply_candidate(candidate)
    }

    /// Append one opaque BIFF12 record to a revision log.
    pub fn push_log_record(&mut self, index: usize, record: RawRecord) -> Result<()> {
        let mut candidate = self.catalog.clone();
        let log = candidate
            .logs
            .get_mut(index)
            .ok_or_else(|| super::invalid("revision-log index is out of range"))?;
        log.records.push(record);
        self.apply_candidate(candidate)
    }

    /// Remove one opaque BIFF12 record from a revision log.
    pub fn remove_log_record(&mut self, index: usize, record: usize) -> Result<Option<RawRecord>> {
        let mut candidate = self.catalog.clone();
        let log = candidate
            .logs
            .get_mut(index)
            .ok_or_else(|| super::invalid("revision-log index is out of range"))?;
        if record >= log.records.len() {
            return Ok(None);
        }
        let value = log.records.remove(record);
        self.apply_candidate(candidate)?;
        Ok(Some(value))
    }

    /// Validate and produce a source-checked commit.
    pub fn commit(self) -> Result<Commit> {
        self.validate_candidate(&self.catalog)?;
        let source =
            Arc::clone(self.base.source.as_ref().ok_or_else(|| {
                super::unsupported("shared-workbook snapshot is not source bound")
            })?);
        let changed = self.catalog != self.base.catalog;
        Ok(Commit {
            snapshot: Snapshot {
                catalog: self.catalog.clone(),
                source: (!changed).then_some(Arc::clone(&source)),
            },
            patch: Patch {
                before: source,
                after: self.catalog,
                changed,
            },
        })
    }

    fn apply_candidate(&mut self, candidate: Catalog) -> Result<()> {
        self.validate_candidate(&candidate)?;
        self.catalog = candidate;
        Ok(())
    }

    fn validate_candidate(&self, candidate: &Catalog) -> Result<()> {
        validation::validate_local(candidate)?;
        match (&candidate.users, &candidate.headers) {
            (None, None) if candidate.logs.is_empty() => Ok(()),
            (Some(_), Some(headers)) if candidate.logs.len() == headers.headers.len() => {
                if self.base.is_empty() {
                    validate_new_graph(candidate)
                } else {
                    validation::validate_catalog(candidate)
                }
            },
            (Some(_), Some(_)) => Err(super::invalid("revision log/header count mismatch")),
            _ => Err(super::invalid(
                "shared-workbook users and headers must occur together",
            )),
        }
    }
}

/// A successful immutable transaction result.
#[derive(Debug, Clone)]
pub struct Commit {
    snapshot: Snapshot,
    patch: Patch,
}

impl Commit {
    /// Borrow the committed semantic snapshot.
    #[must_use]
    pub const fn snapshot(&self) -> &Snapshot {
        &self.snapshot
    }

    /// Borrow the source-checked package patch.
    #[must_use]
    pub const fn patch(&self) -> &Patch {
        &self.patch
    }

    /// Consume the commit into its snapshot and patch.
    #[must_use]
    pub fn into_parts(self) -> (Snapshot, Patch) {
        (self.snapshot, self.patch)
    }
}

/// A clone-staged patch guarded by the exact owner source image.
#[derive(Debug, Clone)]
pub struct Patch {
    before: Arc<SourceState>,
    after: Catalog,
    changed: bool,
}

impl Patch {
    /// Whether applying this patch is a source-preserving no-op.
    #[must_use]
    pub const fn is_empty(&self) -> bool {
        !self.changed
    }

    /// Apply the patch atomically after checking its owner source image.
    pub fn apply(&self, package: &mut OpcPackage) -> Result<Snapshot> {
        let current = SourceState::capture_current(package)?;
        if current != *self.before {
            return Err(Error::UnsupportedFeature(
                "shared-workbook patch source snapshot does not match".to_string(),
            ));
        }
        if !self.changed {
            return Snapshot::read(package);
        }
        let mut candidate = package.clone();
        package::store(&mut candidate, &self.after)?;
        let snapshot = Snapshot::read(&candidate)?;
        *package = candidate;
        Ok(snapshot)
    }

    /// Commit this source-checked patch to the package.
    pub fn commit(&self, package: &mut OpcPackage) -> Result<Snapshot> {
        self.apply(package)
    }
}

/// Read shared-workbook metadata from an OPC package.
pub fn read(package: &OpcPackage) -> Result<Snapshot> {
    Snapshot::read(package)
}

/// Apply a source-checked shared-workbook patch atomically.
pub fn apply(package: &mut OpcPackage, patch: &Patch) -> Result<Snapshot> {
    patch.apply(package)
}

fn validate_new_graph(catalog: &Catalog) -> Result<()> {
    let users = catalog
        .users
        .as_ref()
        .ok_or_else(|| super::invalid("shared-workbook users are missing"))?;
    let headers = catalog
        .headers
        .as_ref()
        .ok_or_else(|| super::invalid("shared-workbook headers are missing"))?;
    let header_guids: HashSet<_> = headers.headers.iter().map(|header| header.guid).collect();
    if users
        .users
        .iter()
        .any(|user| !header_guids.contains(&user.guid))
    {
        return Err(super::invalid(
            "BrtUsr GUID does not identify a revision header",
        ));
    }
    if !headers.headers.is_empty()
        && (headers.info.guid != headers.headers.last().map(|header| header.guid).unwrap()
            || !header_guids.contains(&headers.info.root_guid))
    {
        return Err(super::invalid(
            "BrtInfo GUIDs do not identify revision headers",
        ));
    }
    Ok(())
}

#[derive(Debug, Clone, PartialEq, Eq)]
struct SourceState {
    parts: Vec<SourcePart>,
}

impl SourceState {
    fn capture(package: &OpcPackage, catalog: &Catalog) -> Result<Self> {
        let workbook = package.main_document_part()?;
        let mut names = HashSet::new();
        if let Some(users) = &catalog.users {
            names.insert(users.part_name.clone());
        }
        if let Some(headers) = &catalog.headers {
            names.insert(headers.part_name.clone());
        }
        names.extend(catalog.logs.iter().map(|log| log.part_name.clone()));

        let owner_relationship_ids: HashSet<&str> = catalog
            .users
            .iter()
            .flat_map(|users| [users.relationship_id.as_str()])
            .chain(
                catalog
                    .headers
                    .iter()
                    .flat_map(|headers| [headers.relationship_id.as_str()]),
            )
            .collect();
        let mut parts = Vec::new();
        if !owner_relationship_ids.is_empty() {
            parts.push(SourcePart {
                part_name: workbook.partname().to_string(),
                content_type: workbook.content_type().to_string(),
                blob: None,
                relationships: capture_relationships(workbook, None),
            });
        }
        for name in names {
            let uri = litchi_opc::PackURI::new(&name).map_err(Error::InvalidUri)?;
            let part = package.get_part(&uri)?;
            parts.push(SourcePart {
                part_name: part.partname().to_string(),
                content_type: part.content_type().to_string(),
                blob: Some(part.blob_arc()),
                relationships: capture_relationships(part, Some(&owner_relationship_ids)),
            });
        }
        parts.sort_by(|left, right| left.part_name.cmp(&right.part_name));
        Ok(Self { parts })
    }

    fn capture_current(package: &OpcPackage) -> Result<Self> {
        let catalog = package::load(package)?;
        Self::capture(package, &catalog)
    }
}

fn capture_relationships(
    part: &dyn litchi_opc::Part,
    _owner_relationship_ids: Option<&HashSet<&str>>,
) -> Vec<SourceRelationship> {
    let mut relationships = part
        .rels()
        .iter()
        .filter(|relationship| {
            relationship.reltype() == package::USERS_RELATIONSHIP
                || relationship.reltype() == package::HEADERS_RELATIONSHIP
                || relationship.reltype() == package::LOG_RELATIONSHIP
                || relationship.reltype().ends_with("/usernames")
                || relationship.reltype().ends_with("/revisionHeaders")
                || relationship.reltype().ends_with("/revisionLog")
        })
        .map(|relationship| SourceRelationship {
            id: relationship.r_id().to_string(),
            relationship_type: relationship.reltype().to_string(),
            target: relationship.target_ref().to_string(),
            mode: relationship.target_mode(),
        })
        .collect::<Vec<_>>();
    relationships.sort_by(|left, right| left.id.cmp(&right.id));
    relationships
}

/// Exact source information for one owner or workbook relationship context.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct SourcePart {
    part_name: String,
    content_type: String,
    blob: Option<Arc<Vec<u8>>>,
    relationships: Vec<SourceRelationship>,
}

impl SourcePart {
    /// Absolute OPC part name.
    #[must_use]
    pub fn part_name(&self) -> &str {
        &self.part_name
    }

    /// Part content type.
    #[must_use]
    pub fn content_type(&self) -> &str {
        &self.content_type
    }

    /// Exact source bytes, when this entry represents an owner part.
    #[must_use]
    pub fn bytes(&self) -> &[u8] {
        self.blob.as_deref().map_or(&[], Vec::as_slice)
    }

    /// Whether exact source bytes are present.
    #[must_use]
    pub fn has_bytes(&self) -> bool {
        self.blob.is_some()
    }

    /// Owner relationships in source order by relationship id.
    #[must_use]
    pub fn relationships(&self) -> &[SourceRelationship] {
        &self.relationships
    }
}

/// Exact source relationship metadata for one owner part.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct SourceRelationship {
    id: String,
    relationship_type: String,
    target: String,
    mode: TargetMode,
}

impl SourceRelationship {
    /// Relationship id.
    #[must_use]
    pub fn id(&self) -> &str {
        &self.id
    }

    /// Relationship type URI.
    #[must_use]
    pub fn relationship_type(&self) -> &str {
        &self.relationship_type
    }

    /// Original relationship target reference.
    #[must_use]
    pub fn target(&self) -> &str {
        &self.target
    }

    /// Relationship target mode.
    #[must_use]
    pub const fn mode(&self) -> TargetMode {
        self.mode
    }
}
