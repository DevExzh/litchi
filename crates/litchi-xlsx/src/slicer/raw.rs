//! Native-identity values for typed slicer graph edits.
//!
//! The ordinary slicer model generates relationship IDs and part names. This
//! module is the explicit raw escape for callers that must preserve those
//! physical identities. Its values are accepted only by the package-scoped
//! [`super::Edit`] insertion methods, which still validate and publish the
//! complete XLSX package through the ordinary edit commit.

use super::{Definition, Slicers};

/// A Slicer Cache definition with caller-selected OPC identity.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Cache {
    /// The workbook relationship ID that points at this cache.
    pub relationship_id: String,
    /// The absolute OPC part name for this cache.
    pub part_name: String,
    /// The typed cache definition.
    pub definition: Definition,
}

impl Cache {
    pub(crate) fn into_model(self) -> crate::slicer_cache::Cache {
        crate::slicer_cache::Cache {
            relationship_id: self.relationship_id,
            part_name: self.part_name,
            definition: self.definition,
        }
    }
}

/// A worksheet Slicers part with caller-selected OPC identity.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Part {
    /// The worksheet relationship ID that points at this part.
    pub relationship_id: String,
    /// The absolute OPC part name for this worksheet Slicers part.
    pub part_name: String,
    /// The typed Slicers XML value.
    pub slicers: Slicers,
}

impl Part {
    pub(crate) fn into_model(self) -> crate::slicer_cache::views::SlicerPart {
        crate::slicer_cache::views::SlicerPart {
            relationship_id: self.relationship_id,
            part_name: self.part_name,
            slicers: self.slicers,
        }
    }
}
