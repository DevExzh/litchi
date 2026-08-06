//! Immutable source snapshot for one worksheet's classic comments owner.

use std::sync::Arc;

use litchi_opc::{OpcPackage, PackURI};

use super::model::{Comments, Part};
use super::{load_from_worksheet, package::validate_graph};
use crate::error::{Result, invalid};

/// Immutable semantic and source context for one worksheet comments graph.
///
/// The source XML is retained exactly. A transaction can therefore recognize
/// semantic no-ops without serializing or invalidating package signatures.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct Snapshot {
    worksheet: PackURI,
    part: Option<Part>,
    relationship_type: Option<String>,
    source: Option<Arc<[u8]>>,
}

impl Snapshot {
    /// Capture and validate the comments graph owned by `worksheet`.
    pub fn load(package: &OpcPackage, worksheet: &PackURI) -> Result<Self> {
        validate_graph(package)?;
        let part = load_from_worksheet(package, worksheet)?;
        let relationship_type = part
            .as_ref()
            .map(|part| {
                package
                    .get_part(worksheet)?
                    .rels()
                    .get(&part.relationship_id)
                    .map(|relationship| relationship.reltype().to_owned())
                    .ok_or_else(|| {
                        invalid("classic comments relationship disappeared during snapshot")
                    })
            })
            .transpose()?;
        let source = part
            .as_ref()
            .map(|part| -> Result<Arc<[u8]>> {
                let name = PackURI::new(&part.part_name).map_err(invalid)?;
                let resource = package.get_part(&name)?;
                Ok(Arc::from(resource.blob()))
            })
            .transpose()?;
        Ok(Self {
            worksheet: worksheet.clone(),
            part,
            relationship_type,
            source,
        })
    }

    /// Worksheet part selected by this snapshot.
    pub fn worksheet(&self) -> &PackURI {
        &self.worksheet
    }

    /// Physical comments-part context, when the worksheet has legacy notes.
    pub fn part(&self) -> Option<&Part> {
        self.part.as_ref()
    }

    /// Borrow the typed comments graph, if present.
    pub fn comments(&self) -> Option<&Comments> {
        self.part.as_ref().map(|part| &part.comments)
    }

    /// Exact source bytes of the comments part, if present.
    pub fn source_xml(&self) -> Option<&[u8]> {
        self.source.as_deref()
    }

    /// Exact worksheet-to-comments relationship type, when present.
    pub fn relationship_type(&self) -> Option<&str> {
        self.relationship_type.as_deref()
    }

    /// Whether the worksheet has no classic comments part.
    pub fn is_empty(&self) -> bool {
        self.part.is_none()
    }

    pub(crate) fn same_source(&self, other: &Self) -> bool {
        self.worksheet == other.worksheet
            && self.part == other.part
            && self.relationship_type == other.relationship_type
            && self.source == other.source
    }
}
