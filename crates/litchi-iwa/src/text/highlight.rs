//! Native plain-highlight CRUD for TSWP text storages.

use std::collections::HashSet;

use crate::{IWorkPackage, Result};

use super::annotation::{
    AnnotationKind, add_annotation, annotations, remove_annotation,
    remove_unreferenced_annotation_objects, update_annotation,
};
use super::highlight_types::{TextHighlight, TextHighlightId};
use super::position::TextRange;

/// Read every native plain highlight in a storage, ordered by text position.
pub(crate) fn text_highlights(
    package: &IWorkPackage,
    storage_id: u64,
) -> Result<Vec<TextHighlight>> {
    annotations(package, storage_id)?
        .into_iter()
        .filter(|annotation| annotation.graph.is_plain_highlight())
        .map(|annotation| {
            Ok(TextHighlight::new(
                TextHighlightId::from_native(annotation.object_id),
                annotation.range,
            ))
        })
        .collect()
}

/// Create a plain highlight over a currently unoccupied annotation range.
pub(crate) fn add_text_highlight(
    package: &mut IWorkPackage,
    storage_id: u64,
    range: TextRange,
) -> Result<TextHighlight> {
    let annotation = add_annotation(
        package,
        storage_id,
        range,
        AnnotationKind::PlainHighlight,
        String::new(),
    )?;
    Ok(TextHighlight::new(
        TextHighlightId::from_native(annotation.object_id),
        annotation.range,
    ))
}

/// Atomically move a plain highlight without changing its ID or owned metadata.
pub(crate) fn update_text_highlight(
    package: &mut IWorkPackage,
    storage_id: u64,
    id: TextHighlightId,
    range: TextRange,
) -> Result<TextHighlight> {
    let annotation = update_annotation(
        package,
        storage_id,
        id.object_id(),
        range,
        AnnotationKind::PlainHighlight,
        "",
    )?;
    Ok(TextHighlight::new(id, annotation.range))
}

/// Delete one plain highlight and its owned empty annotation graph.
pub(crate) fn remove_text_highlight(
    package: &mut IWorkPackage,
    storage_id: u64,
    id: TextHighlightId,
) -> Result<TextHighlight> {
    let annotation = remove_annotation(
        package,
        storage_id,
        id.object_id(),
        AnnotationKind::PlainHighlight,
    )?;
    Ok(TextHighlight::new(id, annotation.range))
}

/// Reclaim annotation graphs whose table references disappeared during text replacement.
pub(crate) fn remove_unreferenced_highlight_objects(
    package: &mut IWorkPackage,
    archive_name: &str,
    candidates: &HashSet<u64>,
) -> Result<()> {
    remove_unreferenced_annotation_objects(package, archive_name, candidates)
}

#[cfg(test)]
#[path = "highlight_internal_tests.rs"]
mod tests;
