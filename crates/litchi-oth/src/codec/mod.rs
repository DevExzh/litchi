//! Bounded validation for the family content part.

mod content;

pub(crate) use content::{
    BlockOrder, ReplacementSite, compact_for_publication, project, project_styles, resource_sites,
    validate_authored,
};
