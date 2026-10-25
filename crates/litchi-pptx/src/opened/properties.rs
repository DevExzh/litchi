//! Failure-atomic composition of positional shape text and geometry edits.

use std::collections::HashSet;
use std::ops::Range;

use crate::{Error, Result};

pub(crate) fn stage_shape_properties_with_limits(
    xml: &[u8],
    replacements: &[super::ShapePropertiesReplacement<'_>],
    max_text_bytes: usize,
    max_output_bytes: usize,
    shape_limits: crate::shape::Limits,
) -> Result<(Option<Vec<u8>>, usize)> {
    stage_shape_properties_admitted_with_limits(
        xml,
        replacements,
        max_text_bytes,
        max_output_bytes,
        super::MAX_SHAPE_PROPERTIES_REPLACEMENTS,
        max_text_bytes,
        shape_limits,
    )
}

pub(super) fn stage_shape_properties_admitted_with_limits(
    xml: &[u8],
    replacements: &[super::ShapePropertiesReplacement<'_>],
    max_text_bytes: usize,
    max_output_bytes: usize,
    selection_limit: usize,
    aggregate_text_limit: usize,
    shape_limits: crate::shape::Limits,
) -> Result<(Option<Vec<u8>>, usize)> {
    if replacements.is_empty() {
        return Ok((None, 0));
    }
    if replacements.len() > selection_limit {
        return Err(Error::Limit {
            kind: crate::opened::LimitKind::Codec(crate::opened::CodecLimitKind::ShapeProperties),
            observed: replacements.len(),
            maximum: selection_limit,
        });
    }
    let aggregate_text_bytes = replacements.iter().try_fold(0usize, |total, replacement| {
        let Some(text) = replacement.text() else {
            return Ok(total);
        };
        total.checked_add(text.len()).ok_or_else(|| {
            Error::Invalid("opened-presentation aggregate shape text size overflow".into())
        })
    })?;
    if aggregate_text_bytes > aggregate_text_limit {
        return Err(Error::Limit {
            kind: crate::opened::LimitKind::Codec(crate::opened::CodecLimitKind::TextBytes),
            observed: aggregate_text_bytes,
            maximum: aggregate_text_limit,
        });
    }

    let scene = crate::shape::Scene::read(xml, shape_limits)?;
    if scene.is_rewritten() {
        return Err(Error::refused(
            litchi_core::edit::RefusalKind::UnsupportedContent,
            "set_shape_properties",
            "atomic shape property batches do not support markup-compatibility branch selection",
        ));
    }
    if replacements.len() > scene.len() {
        return Err(Error::Limit {
            kind: crate::opened::LimitKind::Codec(crate::opened::CodecLimitKind::ShapeProperties),
            observed: replacements.len(),
            maximum: scene.len(),
        });
    }

    let mut selected = HashSet::new();
    selected
        .try_reserve(replacements.len())
        .map_err(|source| Error::Allocation {
            resource: "opened-presentation shape property identities",
            source,
        })?;
    let mut text_replacements = Vec::new();
    text_replacements
        .try_reserve_exact(replacements.len())
        .map_err(|source| Error::Allocation {
            resource: "opened-presentation shape property text plan",
            source,
        })?;
    let mut bounds_replacements = Vec::new();
    bounds_replacements
        .try_reserve_exact(replacements.len())
        .map_err(|source| Error::Allocation {
            resource: "opened-presentation shape property geometry plan",
            source,
        })?;
    let mut changed_shapes = HashSet::new();
    changed_shapes
        .try_reserve(replacements.len())
        .map_err(|source| Error::Allocation {
            resource: "opened-presentation changed shape identities",
            source,
        })?;
    let mut selected_spans: Vec<Range<usize>> = Vec::new();
    selected_spans
        .try_reserve_exact(replacements.len())
        .map_err(|source| Error::Allocation {
            resource: "opened-presentation shape property owner ranges",
            source,
        })?;

    for replacement in replacements {
        let position = replacement.position();
        let index = position.get();
        if !selected.insert(index) {
            return Err(Error::refused(
                litchi_core::edit::RefusalKind::InvalidValue,
                "set_shape_properties",
                "a shape is selected more than once",
            ));
        }
        let shape = scene.at(index)?;
        selected_spans.push(shape.common().span()?.range(xml.len())?);
        let current_text = shape.common().text();
        let current_bounds = shape
            .common()
            .bounds()
            .map(|bounds| (bounds.x(), bounds.y(), bounds.width(), bounds.height()));
        if let Some(text) = replacement.text() {
            text_replacements.push(super::ShapeTextReplacement::new(
                crate::shape::Key::Index(index),
                text,
            ));
            if current_text != Some(text) {
                changed_shapes.insert(index);
            }
        }
        if let Some(bounds) = replacement.bounds()
            && current_bounds != Some(bounds)
        {
            changed_shapes.insert(index);
        }
        if let Some(bounds) = replacement.bounds() {
            bounds_replacements.push((position, bounds));
        }
    }
    selected_spans.sort_unstable_by_key(|span| (span.start, span.end));
    for pair in selected_spans.windows(2) {
        if pair[0].end > pair[1].start {
            return Err(Error::Invalid(
                "opened-presentation shape property selections overlap".into(),
            ));
        }
    }
    let changed = changed_shapes.len();

    let mut candidate = if text_replacements.is_empty() {
        None
    } else {
        super::xml::stage_shape_texts_admitted_with_limits(
            xml,
            &text_replacements,
            max_text_bytes,
            max_output_bytes,
            selection_limit,
            aggregate_text_limit,
            shape_limits,
        )?
        .0
    };

    if !bounds_replacements.is_empty() {
        let input = candidate.as_deref().unwrap_or(xml);
        if let Some(next) = super::geometry::rewrite_shape_bounds_many_with_limits(
            input,
            &bounds_replacements,
            max_output_bytes,
            shape_limits,
        )? {
            candidate = Some(next);
        }
    }

    let Some(candidate) = candidate else {
        if changed != 0 {
            return Err(Error::Invalid(
                "opened-presentation shape property batch produced no candidate for a changed shape"
                    .into(),
            ));
        }
        return Ok((None, 0));
    };
    let published = crate::shape::Scene::read(&candidate, shape_limits)?;
    if published.is_rewritten() || published.len() != scene.len() {
        return Err(Error::Invalid(
            "opened-presentation shape property batch changed the semantic shape graph".into(),
        ));
    }
    for index in 0..scene.len() {
        if scene.at(index)?.id() != published.at(index)?.id() {
            return Err(Error::Invalid(
                "opened-presentation shape property batch changed shape identities".into(),
            ));
        }
    }
    for replacement in replacements {
        let shape = published.at(replacement.position().get())?;
        if let Some(text) = replacement.text()
            && shape.common().text() != Some(text)
        {
            return Err(Error::Invalid(
                "opened-presentation shape property text did not round-trip semantically".into(),
            ));
        }
        if let Some(bounds) = replacement.bounds() {
            let actual = shape
                .common()
                .bounds()
                .map(|value| (value.x(), value.y(), value.width(), value.height()));
            if actual != Some(bounds) {
                return Err(Error::Invalid(
                    "opened-presentation shape property bounds did not round-trip semantically"
                        .into(),
                ));
            }
        }
    }
    Ok((Some(candidate), changed))
}
