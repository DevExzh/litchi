//! Lexical proof for reversible top-level removal and canonical text-box additions.
use super::xml::{shape_reference_ids, span_bytes, validate_shape_removal_dependencies};
use crate::{Error, Result};
use std::collections::BTreeSet;
use std::ops::Range;

pub(crate) struct NormalizedPair {
    pub(crate) before: Vec<u8>,
    pub(crate) after: Vec<u8>,
}

pub(crate) fn normalize_shape_removal_pair_with_limits(
    before: &[u8],
    after: &[u8],
    forward: bool,
    replacements: bool,
    shape_limits: crate::shape::Limits,
) -> Result<Option<NormalizedPair>> {
    let (complete, candidate) = if forward {
        (before, after)
    } else {
        (after, before)
    };
    let scene = crate::shape::Scene::read(complete, shape_limits)?;
    let next = crate::shape::Scene::read(candidate, shape_limits)?;
    if scene.is_rewritten() || next.is_rewritten() {
        return Ok(None);
    }
    for value in [&scene, &next] {
        let mut ids = BTreeSet::new();
        for shape in value.iter() {
            let Some(id) = shape.id() else {
                return Ok(None);
            };
            if !ids.insert(id) {
                return Ok(None);
            }
        }
    }
    // Keep the root sequences as values so that we can first match the
    // retained subsequence by identity and then peel off a canonical suffix.
    // The latter is needed when a removed highest root is replaced by the
    // source authoring path and therefore reuses its identity.
    let mut complete_roots = Vec::new();
    complete_roots
        .try_reserve_exact(scene.roots().count())
        .map_err(allocation)?;
    complete_roots.extend(scene.roots());
    let mut candidate_roots = Vec::new();
    candidate_roots
        .try_reserve_exact(next.roots().count())
        .map_err(allocation)?;
    candidate_roots.extend(next.roots());

    let conformance = crate::media_parts::document_conformance(candidate)?;
    let mut retained_indices = Vec::new();
    retained_indices
        .try_reserve_exact(candidate_roots.len())
        .map_err(allocation)?;
    let mut complete_cursor = 0;
    for candidate_root in &candidate_roots {
        while complete_cursor < complete_roots.len()
            && complete_roots[complete_cursor].id() != candidate_root.id()
        {
            complete_cursor += 1;
        }
        if complete_cursor == complete_roots.len() {
            break;
        }
        retained_indices.push(complete_cursor);
        complete_cursor += 1;
    }

    // A candidate root with the same identity as an old root is normally a
    // retained root. If it is the boundary root and has exactly the fragment
    // emitted by the native text-box authoring path, classify it as an
    // authored suffix instead. This admits delete-and-append replacement
    // while keeping arbitrary metadata edits outside the proof.
    let mut retained_count = retained_indices.len();
    if replacements {
        while retained_count > 0 {
            let candidate_root = candidate_roots[retained_count - 1];
            let complete_root = complete_roots[retained_indices[retained_count - 1]];
            if candidate_root.xml()? == complete_root.xml()? {
                break;
            }
            if !canonical_text_box(candidate_root, conformance)? {
                break;
            }
            retained_count -= 1;
            retained_indices.pop();
        }
    }

    let mut spans = Vec::new();
    spans
        .try_reserve_exact(scene.roots().count())
        .map_err(allocation)?;
    let mut retained_cursor = 0;
    for (index, root) in complete_roots.iter().enumerate() {
        if retained_indices.get(retained_cursor) == Some(&index) {
            retained_cursor += 1;
        } else {
            spans.push(root.common().span()?.range(complete.len())?);
        }
    }
    if spans.is_empty() {
        return Ok(None);
    }
    let mut removed_ids = BTreeSet::new();
    let mut range_index = 0;
    let mut retained_max = 0;
    for shape in scene.iter() {
        let Some(id) = shape.id() else {
            return Ok(None);
        };
        let range = shape.common().span()?.range(complete.len())?;
        while spans
            .get(range_index)
            .is_some_and(|removed| removed.end <= range.start)
        {
            range_index += 1;
        }
        if spans
            .get(range_index)
            .is_some_and(|removed| removed.start <= range.start && range.end <= removed.end)
        {
            removed_ids.insert(id);
        } else {
            retained_max = retained_max.max(id);
        }
    }
    let normalized_complete = strip_ranges(complete, &spans)?;
    let mut added = Vec::new();
    added
        .try_reserve_exact(candidate_roots.len().saturating_sub(retained_count))
        .map_err(allocation)?;
    for shape in candidate_roots.iter().skip(retained_count).copied() {
        if !canonical_text_box(shape, conformance)? {
            return Ok(None);
        }
        added.push(shape.common().span()?.range(candidate.len())?);
    }
    let normalized_candidate = strip_ranges(candidate, &added)?;

    // The allocator may leave gaps after temporary additions, but every
    // authored ID still has to be above all identities that survive in the
    // normalized candidate. Otherwise a forged suffix could collide with a
    // dangling timing/control/connector reference.
    let retained_reference_max = if added.is_empty() {
        0
    } else {
        match shape_reference_ids(&normalized_candidate) {
            Ok(references) => references.into_iter().max().unwrap_or(0),
            Err(Error::Refused(_)) => return Ok(None),
            Err(error) => return Err(error),
        }
    };
    let mut allocation_floor = retained_max.max(retained_reference_max);
    for shape in candidate_roots.iter().skip(retained_count) {
        let Some(id) = shape.id() else {
            return Ok(None);
        };
        if id <= allocation_floor {
            return Ok(None);
        }
        allocation_floor = id;
    }

    // Check the real candidate, including the authored suffix. An unsafe
    // direction is not a proof; the caller may still prove its inverse.
    match validate_shape_removal_dependencies(candidate, &removed_ids) {
        Ok(()) => {},
        Err(Error::Refused(_)) => return Ok(None),
        Err(error) => return Err(error),
    }
    Ok(Some(if forward {
        NormalizedPair {
            before: normalized_complete,
            after: normalized_candidate,
        }
    } else {
        NormalizedPair {
            before: normalized_candidate,
            after: normalized_complete,
        }
    }))
}

fn canonical_text_box(
    shape: crate::shape::Shape<'_>,
    conformance: crate::media_parts::Conformance,
) -> Result<bool> {
    let crate::shape::Shape::Auto(_) = shape else {
        return Ok(false);
    };
    let Some(id) = shape.id() else {
        return Ok(false);
    };
    if shape.placeholder().is_some() {
        return Ok(false);
    }
    let (Some(bounds), Some(text)) = (shape.bounds(), shape.text()) else {
        return Ok(false);
    };
    let name = format!("Text Box {id}");
    if shape.name() != Some(name.as_str()) {
        return Ok(false);
    }
    let fragment = super::text_box_fragment(
        text,
        (bounds.x(), bounds.y(), bounds.width(), bounds.height()),
        id,
        conformance,
        super::Limits::default(),
    )?;
    Ok(shape.xml()? == fragment.as_bytes())
}

fn strip_ranges(xml: &[u8], spans: &[Range<usize>]) -> Result<Vec<u8>> {
    let mut output = Vec::new();
    output.try_reserve_exact(xml.len()).map_err(allocation)?;
    let mut end = 0;
    for span in spans {
        output.extend_from_slice(span_bytes(xml, end..span.start)?);
        end = span.end;
    }
    output.extend_from_slice(span_bytes(xml, end..xml.len())?);
    Ok(output)
}

fn allocation(source: std::collections::TryReserveError) -> Error {
    Error::Allocation {
        resource: "source shape-removal proof",
        source,
    }
}
