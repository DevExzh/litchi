//! Immutable source snapshots for the workbook revision owner.

use std::sync::Arc;

use litchi_opc::{OpcPackage, PackURI, Part, TargetMode};

use super::model::{
    RevisionConformance, RevisionHeaders, RevisionLogPart, RevisionUsers, Revisions,
};
use super::package::load_workbook_revisions;
use crate::error::{Result, invalid};

/// An immutable semantic and physical snapshot of workbook revision metadata.
///
/// The typed graph and the exact bytes of every owned part are captured
/// together. A transaction can therefore recognize a semantic no-op without
/// serializing any XML, and a patch can reject a stale owner before planning
/// package changes.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct Snapshot {
    revisions: Option<Revisions>,
    source: SourceState,
}

impl Snapshot {
    /// Load and validate the revision graph from an OPC package.
    pub fn load(package: &OpcPackage) -> Result<Self> {
        let revisions = load_workbook_revisions(package)?;
        let source = SourceState::capture(package, revisions.as_ref())?;
        Ok(Self { revisions, source })
    }

    /// Alias for [`Self::load`] that emphasizes the source-bound result.
    pub fn read(package: &OpcPackage) -> Result<Self> {
        Self::load(package)
    }

    /// Borrow the complete typed revision owner, when present.
    pub fn revisions(&self) -> Option<&Revisions> {
        self.revisions.as_ref()
    }

    /// Borrow the revision users part, when present.
    pub fn users(&self) -> Option<&RevisionUsers> {
        self.revisions.as_ref().map(|value| &value.users)
    }

    /// Borrow the revision headers part, when present.
    pub fn headers(&self) -> Option<&RevisionHeaders> {
        self.revisions.as_ref().map(|value| &value.headers)
    }

    /// Borrow revision log parts in the same order as their headers.
    pub fn logs(&self) -> &[RevisionLogPart] {
        self.revisions
            .as_ref()
            .map_or(&[], |value| value.logs.as_slice())
    }

    /// Return the conformance used by the source revision relationships.
    pub fn conformance(&self) -> Option<RevisionConformance> {
        self.source.conformance
    }

    /// Return the workbook part that owns the revision relationships.
    pub fn workbook_part_name(&self) -> &str {
        &self.source.workbook_part_name
    }

    /// Return the workbook part that owns the revision relationships.
    pub fn workbook(&self) -> &str {
        self.workbook_part_name()
    }

    /// Return the conformance used for writes, defaulting new owners to
    /// Transitional SpreadsheetML.
    pub fn desired_conformance(&self) -> RevisionConformance {
        self.conformance()
            .unwrap_or(RevisionConformance::Transitional)
    }

    /// Return exact source bytes for one owned revision part.
    pub fn source_xml(&self, part_name: &str) -> Option<&[u8]> {
        self.source
            .parts
            .iter()
            .find(|part| part.part_name == part_name)
            .map(SourcePart::bytes)
    }

    /// Whether this workbook has no revision metadata.
    pub fn is_empty(&self) -> bool {
        self.revisions.is_none()
    }

    pub(crate) fn same_source(&self, other: &Self) -> bool {
        self.source == other.source
    }

    pub(crate) fn source(&self) -> &SourceState {
        &self.source
    }
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub(crate) struct SourceState {
    workbook_part_name: String,
    workbook_content_type: String,
    workbook_relationships: Vec<SourceRelationship>,
    parts: Vec<SourcePart>,
    conformance: Option<RevisionConformance>,
}

impl SourceState {
    fn capture(package: &OpcPackage, revisions: Option<&Revisions>) -> Result<Self> {
        let workbook = package.main_document_part()?;
        let workbook_relationships = workbook
            .rels()
            .iter()
            .filter(|relationship| super::package::is_revision_relationship(relationship.reltype()))
            .map(SourceRelationship::from_relationship)
            .collect::<Vec<_>>();
        let mut workbook_relationships = workbook_relationships;
        workbook_relationships.sort_by(|left, right| left.id.cmp(&right.id));
        let mut parts = Vec::new();
        if let Some(revisions) = revisions {
            let mut names = Vec::with_capacity(revisions.logs.len().saturating_add(2));
            names.push(revisions.users_part_name.as_str());
            names.push(revisions.headers_part_name.as_str());
            names.extend(revisions.logs.iter().map(|log| log.part_name.as_str()));
            for name in names {
                let uri = PackURI::new(name).map_err(invalid)?;
                let part = package.get_part(&uri)?;
                parts.push(SourcePart::from_part(part));
            }
            parts.sort_by(|left, right| left.part_name.cmp(&right.part_name));
        }
        Ok(Self {
            workbook_part_name: workbook.partname().to_string(),
            workbook_content_type: workbook.content_type().to_owned(),
            workbook_relationships,
            parts,
            conformance: super::package::revision_conformance(package)?,
        })
    }

    pub(crate) fn workbook_part_name(&self) -> &str {
        &self.workbook_part_name
    }

    pub(crate) fn workbook_content_type(&self) -> &str {
        &self.workbook_content_type
    }

    pub(crate) fn workbook_relationships(&self) -> &[SourceRelationship] {
        &self.workbook_relationships
    }

    pub(crate) fn parts(&self) -> &[SourcePart] {
        &self.parts
    }
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub(crate) struct SourcePart {
    part_name: String,
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
            part_name: part.partname().to_string(),
            content_type: part.content_type().to_owned(),
            bytes: part.blob_arc(),
            relationships,
        }
    }

    pub(crate) fn bytes(&self) -> &[u8] {
        self.bytes.as_slice()
    }

    pub(crate) fn part_name(&self) -> &str {
        &self.part_name
    }

    pub(crate) fn content_type(&self) -> &str {
        &self.content_type
    }

    pub(crate) fn bytes_arc(&self) -> Arc<Vec<u8>> {
        Arc::clone(&self.bytes)
    }

    pub(crate) fn relationships(&self) -> &[SourceRelationship] {
        &self.relationships
    }
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub(crate) struct SourceRelationship {
    id: String,
    relationship_type: String,
    target: String,
    mode: TargetMode,
}

impl SourceRelationship {
    fn from_relationship(relationship: &litchi_opc::Relationship) -> Self {
        Self {
            id: relationship.r_id().to_owned(),
            relationship_type: relationship.reltype().to_owned(),
            target: relationship.target_ref().to_owned(),
            mode: relationship.target_mode(),
        }
    }

    pub(crate) fn id(&self) -> &str {
        &self.id
    }

    pub(crate) fn relationship_type(&self) -> &str {
        &self.relationship_type
    }

    pub(crate) fn target(&self) -> &str {
        &self.target
    }

    pub(crate) fn mode(&self) -> TargetMode {
        self.mode
    }
}
