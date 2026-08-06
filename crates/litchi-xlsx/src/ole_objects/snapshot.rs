//! Immutable source snapshots for one worksheet's OLE-object owner.

use std::sync::Arc;

use litchi_opc::{OpcPackage, PackURI, Part, TargetMode};

use super::model::OleObjects;
use super::{codec, package::validate_graph};
use crate::error::{Result, invalid};

/// An immutable semantic and physical snapshot of one worksheet OLE graph.
///
/// The worksheet XML and every referenced opaque payload/preview part are
/// retained exactly. Transactions can therefore patch known metadata without
/// rebuilding extension markup or opening embedded content, while patches can
/// reject a stale source before touching the package.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct Snapshot {
    worksheet: PackURI,
    objects: Option<OleObjects>,
    source: SourceState,
    conformance: super::model::OleObjectConformance,
}

impl Snapshot {
    /// Capture and validate the OLE graph owned by one worksheet.
    pub fn load(package: &OpcPackage, worksheet: &PackURI) -> Result<Self> {
        let objects = validate_graph(package, worksheet)?;
        let worksheet_part = package.get_part(worksheet)?;
        let root = codec::parse_document(worksheet_part.blob())?;
        let conformance = codec::crate_conformance(&root)?;
        let source = SourceState::capture(package, worksheet_part, objects.as_ref())?;
        Ok(Self {
            worksheet: worksheet.clone(),
            objects,
            source,
            conformance,
        })
    }

    /// Alias for [`Self::load`] emphasizing the source-bound result.
    pub fn read(package: &OpcPackage, worksheet: &PackURI) -> Result<Self> {
        Self::load(package, worksheet)
    }

    /// Worksheet selected by this snapshot.
    pub fn worksheet(&self) -> &PackURI {
        &self.worksheet
    }

    /// Borrow the complete typed OLE graph, when the worksheet has one.
    pub fn objects(&self) -> Option<&OleObjects> {
        self.objects.as_ref()
    }

    /// Contextual alias for [`Self::objects`].
    pub fn ole_objects(&self) -> Option<&OleObjects> {
        self.objects()
    }

    /// Exact worksheet XML captured by this snapshot.
    pub fn source_xml(&self) -> &[u8] {
        self.source.xml.as_slice()
    }

    /// Conformance namespace selected by the worksheet source.
    pub fn conformance(&self) -> super::model::OleObjectConformance {
        self.conformance
    }

    /// Whether this worksheet has no OLE-object collection.
    pub fn is_empty(&self) -> bool {
        self.objects.is_none()
    }

    pub(crate) fn same_source(&self, other: &Self) -> bool {
        self.worksheet == other.worksheet && self.source == other.source
    }
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub(crate) struct SourceState {
    xml: Arc<Vec<u8>>,
    content_type: String,
    relationships: Vec<SourceRelationship>,
    resources: Vec<SourcePart>,
}

impl SourceState {
    fn capture(
        package: &OpcPackage,
        worksheet: &dyn Part,
        objects: Option<&OleObjects>,
    ) -> Result<Self> {
        let mut relationships = worksheet
            .rels()
            .iter()
            .map(SourceRelationship::from_relationship)
            .collect::<Vec<_>>();
        relationships.sort_by(|left, right| left.id.cmp(&right.id));

        let mut names = std::collections::HashSet::new();
        if let Some(objects) = objects {
            for object in &objects.objects {
                if let Some(super::model::OleObjectTarget::Internal(resource)) = &object.target {
                    names.insert(resource.part_name.clone());
                }
                if let Some(properties) = &object.properties {
                    if let Some(preview) = &properties.preview {
                        names.insert(preview.part_name.clone());
                    }
                }
            }
        }

        let mut resources = Vec::with_capacity(names.len());
        for name in names {
            let uri = PackURI::new(&name).map_err(invalid)?;
            resources.push(SourcePart::from_part(package.get_part(&uri)?));
        }
        resources.sort_by(|left, right| left.part_name.cmp(&right.part_name));

        Ok(Self {
            xml: worksheet.blob_arc(),
            content_type: worksheet.content_type().to_owned(),
            relationships,
            resources,
        })
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
}
