//! Bounded validation for the family content part.

mod content;
mod structure;

pub(crate) use content::{
    BlockOrder, ReplacementSite, compact_for_publication, project, project_styles, resource_sites,
    validate_authored,
};
pub(crate) use structure::{
    BodyStructures, EditableStructureKind, StructureSite, has_semantic_reference,
    project_structures, structure_sites,
};
