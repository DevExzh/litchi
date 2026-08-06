//! Immutable source snapshots for the workbook Custom XML Maps owner.

use std::sync::Arc;

use litchi_core::sheet::Result;
use litchi_opc::{OpcPackage, PackURI, Part, TargetMode};

use super::model::{REL, STRICT_REL, XmlMapConformance, XmlMapInfo};
use super::{invalid, package};

/// An immutable semantic and physical snapshot of the workbook's Custom XML
/// Maps owner.
///
/// The source workbook bytes, workbook relationship topology, and owned XML
/// part are retained so a transaction can reject stale sources and publish a
/// known-field edit without rebuilding opaque XML payloads.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct Snapshot {
    info: Option<XmlMapInfo>,
    source: SourceState,
    conformance: XmlMapConformance,
}

impl Snapshot {
    /// Capture and validate the complete Custom XML Maps graph.
    pub fn load(package: &OpcPackage) -> Result<Self> {
        let loaded = package::load_from_package_with_conformance(package)?;
        let info = loaded.as_ref().map(|(value, _)| value.clone());
        let conformance = loaded
            .as_ref()
            .map(|(_, conformance)| *conformance)
            .unwrap_or_default();
        let workbook = package.main_document_part()?;
        let source = SourceState::capture(package, workbook)?;
        Ok(Self {
            info,
            source,
            conformance,
        })
    }

    /// Alias for [`Self::load`] emphasizing the source-bound result.
    pub fn read(package: &OpcPackage) -> Result<Self> {
        Self::load(package)
    }

    /// Borrow the typed Custom XML Maps value when the workbook owns one.
    pub fn info(&self) -> Option<&XmlMapInfo> {
        self.info.as_ref()
    }

    /// Contextual alias for [`Self::info`].
    pub fn xml_maps(&self) -> Option<&XmlMapInfo> {
        self.info()
    }

    /// Contextual catalog alias for [`Self::info`].
    pub fn catalog(&self) -> Option<&XmlMapInfo> {
        self.info()
    }

    /// Explicit alias for the typed `MapInfo` value.
    pub fn map_info(&self) -> Option<&XmlMapInfo> {
        self.info()
    }

    /// Contextual value alias for [`Self::info`].
    pub fn value(&self) -> Option<&XmlMapInfo> {
        self.info()
    }

    /// The namespace conformance used by the owned part, or Transitional for
    /// an absent owner.
    pub fn conformance(&self) -> XmlMapConformance {
        self.conformance
    }

    /// The workbook part that owns the Custom XML Maps relationship.
    pub fn workbook_part_name(&self) -> &str {
        &self.source.workbook_part_name
    }

    /// The owned Custom XML Maps part identity, when present.
    pub fn part_name(&self) -> Option<&PackURI> {
        self.source.part.as_ref().map(|part| &part.part_uri)
    }

    /// Exact source XML bytes for the owned Custom XML Maps part.
    pub fn source_xml(&self) -> Option<&[u8]> {
        self.source.part.as_ref().map(SourcePart::bytes)
    }

    /// Whether this workbook currently has no Custom XML Maps owner.
    pub fn is_empty(&self) -> bool {
        self.info.is_none()
    }

    pub(crate) fn same_source(&self, other: &Self) -> bool {
        self.source == other.source
    }

    pub(crate) fn restore_into(&self, package: &mut OpcPackage) -> Result<()> {
        let workbook_uri = package.main_document_part()?.partname().clone();
        if workbook_uri.as_str() != self.source.workbook_part_name {
            return Err(invalid(
                "custom XML maps patch targets a different workbook",
            ));
        }

        let existing = package
            .get_part(&workbook_uri)?
            .rels()
            .iter()
            .filter(|relationship| matches!(relationship.reltype(), REL | STRICT_REL))
            .map(|relationship| {
                Ok((
                    relationship.r_id().to_owned(),
                    relationship.target_partname()?,
                ))
            })
            .collect::<Result<Vec<_>>>()?;
        if existing.len() > 1 {
            return Err(invalid(
                "workbook has multiple custom XML maps relationships",
            ));
        }

        for (relationship_id, part_uri) in existing {
            package
                .get_part_mut(&workbook_uri)?
                .rels_mut()
                .remove(&relationship_id);
            if !part_is_referenced(package, &part_uri) {
                package.remove_part(&part_uri);
            }
        }

        if let Some(part) = &self.source.part {
            if let Ok(existing) = package.get_part(&part.part_uri) {
                if existing.content_type() != part.content_type {
                    return Err(invalid(format!(
                        "custom XML maps part '{}' has an incompatible content type",
                        part.part_uri
                    )));
                }
                if existing.rels().iter().next().is_some() {
                    return Err(invalid("custom XML maps part must not have relationships"));
                }
                package
                    .get_part_mut(&part.part_uri)?
                    .set_blob(part.bytes().to_vec());
            } else {
                let mut replacement = litchi_opc::part::BlobPart::new(
                    part.part_uri.clone(),
                    part.content_type.clone(),
                    part.bytes().to_vec(),
                );
                for relationship in &part.relationships {
                    replacement.rels_mut().add_relationship(
                        relationship.relationship_type.clone(),
                        relationship.target.clone(),
                        relationship.id.clone(),
                        relationship.mode == TargetMode::External,
                    );
                }
                package.try_add_part(Box::new(replacement))?;
            }

            let owner = self
                .source
                .workbook_relationship()
                .ok_or_else(|| invalid("custom XML maps snapshot is missing its owner"))?;
            package
                .get_part_mut(&workbook_uri)?
                .rels_mut()
                .add_relationship(
                    owner.relationship_type.clone(),
                    owner.target.clone(),
                    owner.id.clone(),
                    owner.mode == TargetMode::External,
                );
        }
        package.unsign();
        Ok(())
    }
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub(crate) struct SourceState {
    pub(crate) workbook_part_name: String,
    pub(crate) workbook_content_type: String,
    pub(crate) workbook_bytes: Arc<Vec<u8>>,
    pub(crate) workbook_relationships: Vec<SourceRelationship>,
    pub(crate) root_relationships: Vec<SourceRelationship>,
    pub(crate) part: Option<SourcePart>,
}

impl SourceState {
    fn capture(package: &OpcPackage, workbook: &dyn Part) -> Result<Self> {
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

        let part = if let Some(relationship) = workbook
            .rels()
            .iter()
            .find(|relationship| matches!(relationship.reltype(), REL | STRICT_REL))
        {
            let part_uri = relationship.target_partname()?;
            Some(SourcePart::from_part(package.get_part(&part_uri)?))
        } else {
            None
        };

        Ok(Self {
            workbook_part_name: workbook.partname().to_string(),
            workbook_content_type: workbook.content_type().to_owned(),
            workbook_bytes: workbook.blob_arc(),
            workbook_relationships,
            root_relationships,
            part,
        })
    }

    fn workbook_relationship(&self) -> Option<&SourceRelationship> {
        self.workbook_relationships.iter().find(|relationship| {
            matches!(relationship.relationship_type.as_str(), REL | STRICT_REL)
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

    fn bytes(&self) -> &[u8] {
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

fn part_is_referenced(package: &OpcPackage, target: &PackURI) -> bool {
    package.iter_parts().any(|part| {
        part.rels().iter().any(|relationship| {
            !relationship.is_external()
                && relationship
                    .target_partname()
                    .is_ok_and(|part_name| part_name == *target)
        })
    }) || package.rels().iter().any(|relationship| {
        !relationship.is_external()
            && relationship
                .target_partname()
                .is_ok_and(|part_name| part_name == *target)
    })
}
