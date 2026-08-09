//! Immutable workbook snapshots for the external-link owner.

use std::sync::Arc;

use litchi_opc::{OpcPackage, PackURI, Part, TargetMode};

use super::model::Conformance;
use super::package::{Entry, load_external_links};
use super::{invalid, validation};
use crate::error::Result;

/// An immutable semantic and physical snapshot of every external link owned
/// by the workbook part.
///
/// Target URLs, DDE topics, and OLE metadata are retained as inert values. No
/// external package is opened and no host refresh/activation behavior is
/// reachable through this owner.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct Snapshot {
    entries: Vec<Entry>,
    source: SourceState,
    conformance: Conformance,
    next_relationship_id: String,
    next_part_uri: PackURI,
}

impl Snapshot {
    /// Capture and validate the complete workbook external-link graph.
    pub fn load(package: &OpcPackage) -> Result<Self> {
        let entries = load_external_links(package)?;
        let source = SourceState::capture(package, &entries)?;
        let conformance = source.conformance;
        validation::entries(&entries, conformance)?;
        let next_relationship_id = next_relationship_id(&source.workbook_relationships);
        let next_part_uri = next_part_uri(package)?;
        Ok(Self {
            entries,
            source,
            conformance,
            next_relationship_id,
            next_part_uri,
        })
    }

    /// Alias for [`Self::load`] emphasizing the source-bound result.
    pub fn read(package: &OpcPackage) -> Result<Self> {
        Self::load(package)
    }

    /// Borrow the workbook's external-link entries in stable relationship-ID
    /// order.
    #[must_use]
    pub fn entries(&self) -> &[Entry] {
        &self.entries
    }

    /// Contextual alias for [`Self::entries`].
    #[must_use]
    pub fn links(&self) -> &[Entry] {
        self.entries()
    }

    /// The workbook part that owns the external-link relationships.
    #[must_use]
    pub fn workbook_part_name(&self) -> &str {
        &self.source.workbook_part_name
    }

    /// Namespace conformance used for newly authored external-link parts.
    #[must_use]
    pub fn conformance(&self) -> Conformance {
        self.conformance
    }

    /// Exact source XML for one owned external-link part.
    pub fn source_xml(&self, part_uri: &PackURI) -> Option<&[u8]> {
        self.source
            .parts
            .iter()
            .find(|part| part.part_uri == *part_uri)
            .map(SourcePart::bytes)
    }

    /// Whether this workbook owns no external links.
    #[must_use]
    pub fn is_empty(&self) -> bool {
        self.entries.is_empty()
    }

    pub(crate) fn same_source(&self, other: &Self) -> bool {
        self.source == other.source
    }

    pub(crate) fn next_relationship_id(&self) -> &str {
        &self.next_relationship_id
    }
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub(crate) struct SourceState {
    pub(crate) workbook_part_name: String,
    pub(crate) workbook_content_type: String,
    pub(crate) workbook_relationships: Vec<SourceRelationship>,
    pub(crate) parts: Vec<SourcePart>,
    pub(crate) conformance: Conformance,
}

impl SourceState {
    fn capture(package: &OpcPackage, entries: &[Entry]) -> Result<Self> {
        let workbook = package.main_document_part()?;
        let mut workbook_relationships = workbook
            .rels()
            .iter()
            .map(SourceRelationship::from_relationship)
            .collect::<Vec<_>>();
        workbook_relationships.sort_by(|left, right| left.id.cmp(&right.id));

        let mut parts = Vec::with_capacity(entries.len());
        for entry in entries {
            parts.push(SourcePart::from_part(package.get_part(&entry.part_uri)?));
        }
        parts.sort_by(|left, right| left.part_uri.as_str().cmp(right.part_uri.as_str()));

        let conformance = entries.first().map_or_else(
            || detect_workbook_conformance(workbook.blob()),
            |entry| detect_conformance(&parts, &entry.part_uri),
        );
        Ok(Self {
            workbook_part_name: workbook.partname().to_string(),
            workbook_content_type: workbook.content_type().to_owned(),
            workbook_relationships,
            parts,
            conformance,
        })
    }
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub(crate) struct SourcePart {
    pub(crate) part_uri: PackURI,
    pub(crate) content_type: String,
    pub(crate) bytes: Arc<Vec<u8>>,
    pub(crate) relationships: Vec<SourceRelationship>,
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

    pub(crate) fn bytes(&self) -> &[u8] {
        self.bytes.as_slice()
    }
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub(crate) struct SourceRelationship {
    pub(crate) id: String,
    pub(crate) relationship_type: String,
    pub(crate) target: String,
    pub(crate) mode: TargetMode,
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
}

fn next_relationship_id(relationships: &[SourceRelationship]) -> String {
    let mut candidate = 1u32;
    loop {
        let id = format!("rId{candidate}");
        if !relationships
            .iter()
            .any(|relationship| relationship.id == id)
        {
            return id;
        }
        candidate = candidate.saturating_add(1);
    }
}

pub(crate) fn next_part_uri(package: &OpcPackage) -> Result<PackURI> {
    let mut candidate = 1u32;
    loop {
        let uri = PackURI::new(format!("/xl/externalLinks/externalLink{candidate}.xml"))
            .map_err(invalid)?;
        if package.get_part(&uri).is_err() {
            return Ok(uri);
        }
        candidate = candidate
            .checked_add(1)
            .ok_or_else(|| invalid("external-link part name space exhausted"))?;
    }
}

fn detect_conformance(parts: &[SourcePart], part_uri: &PackURI) -> Conformance {
    parts
        .iter()
        .find(|part| part.part_uri == *part_uri)
        .map(|part| detect_workbook_conformance(part.bytes()))
        .unwrap_or_default()
}

fn detect_workbook_conformance(xml: &[u8]) -> Conformance {
    if xml
        .windows(super::model::STRICT_SML.len())
        .any(|window| window == super::model::STRICT_SML.as_bytes())
    {
        Conformance::Strict
    } else {
        Conformance::Transitional
    }
}
