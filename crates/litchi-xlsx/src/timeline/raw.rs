//! Native-identity values for typed timeline graph edits.
//!
//! The ordinary timeline model generates relationship IDs and part names.
//! This module is the explicit raw escape for callers that must preserve
//! those physical identities. Its values are accepted only by the
//! package-scoped [`super::Edit`] insertion methods, which still validate and
//! publish the complete XLSX package through the ordinary edit commit.

use super::{CacheDefinition, Views};

/// A View Cache definition with caller-selected OPC identity.
#[derive(Debug, Clone, PartialEq)]
pub struct Cache {
    /// The workbook relationship ID that points at this cache.
    pub relationship_id: String,
    /// The absolute OPC part name for this cache.
    pub part_name: String,
    /// The typed cache definition.
    pub definition: CacheDefinition,
}

impl Cache {
    pub(crate) fn into_model(self) -> crate::timelines::Cache {
        crate::timelines::Cache {
            relationship_id: self.relationship_id,
            part_name: self.part_name,
            definition: self.definition,
        }
    }
}

/// A worksheet Timelines part with caller-selected worksheet and OPC identity.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Part {
    /// The worksheet part that owns the relationship.
    pub worksheet_part_name: String,
    /// The worksheet relationship ID that points at this part.
    pub relationship_id: String,
    /// The absolute OPC part name for this worksheet Timelines part.
    pub part_name: String,
    /// The typed Timelines XML value.
    pub timelines: Views,
}

impl Part {
    pub(crate) fn into_model(self) -> crate::timelines::Part {
        crate::timelines::Part {
            worksheet_part_name: self.worksheet_part_name,
            relationship_id: self.relationship_id,
            part_name: self.part_name,
            timelines: self.timelines,
        }
    }
}
