//! Presentation-structure OPC facade.

use std::sync::Arc;

use litchi_opc::OpcPackage;

pub use super::codec::{
    add_custom_show, add_custom_show_slide, add_section, add_section_slide, find_custom_show,
    find_section, load, remove_custom_show, remove_custom_show_slide, remove_section,
    remove_section_slide, reorder_custom_show_slides, reorder_custom_shows, reorder_section_slides,
    reorder_sections, replace_custom_show, replace_section, store,
    synchronize_after_slide_mutation, update_custom_show, update_section,
};

use super::transaction::{Commit, Patch, Snapshot};
use crate::{Error, Result};

/// Capture the validated presentation graph and its exact source context.
pub fn load_snapshot(package: &OpcPackage) -> Result<Snapshot> {
    let graph = load(package)?;
    let presentation = package.main_document_part()?;
    Snapshot::from_wire(
        presentation.partname().to_string(),
        presentation.content_type().to_owned(),
        presentation.blob_arc(),
        Snapshot::capture_relationships(presentation),
        graph,
    )
}

/// Apply a committed presentation-structure patch atomically.
pub fn apply_patch(package: &mut OpcPackage, patch: &Patch) -> Result<Snapshot> {
    let current = load_snapshot(package)?;
    if !current.same_source(patch.before()) {
        return Err(invalid("presentation structure source is stale"));
    }
    if patch.is_empty() {
        return Ok(current);
    }

    let presentation_name = package.main_document_part()?.partname().clone();
    if presentation_name.as_str() != patch.after().presentation_part_name() {
        return Err(invalid(
            "presentation structure patch changes the owning part",
        ));
    }
    let mut candidate = package.clone();
    candidate
        .get_part_mut(&presentation_name)?
        .set_blob_shared(Arc::clone(patch.after().source_arc()));
    let result = load_snapshot(&candidate)?;
    if !result.same_source(patch.after()) || result.graph() != patch.after().graph() {
        return Err(invalid(
            "published presentation structure differs from the commit",
        ));
    }
    candidate.unsign();
    *package = candidate;
    Ok(result)
}

/// Apply a committed presentation-structure edit atomically.
pub fn apply_commit(package: &mut OpcPackage, commit: Commit) -> Result<Snapshot> {
    apply_patch(package, commit.patch())
}

fn invalid(message: impl Into<String>) -> Error {
    Error::Invalid(message.into())
}
