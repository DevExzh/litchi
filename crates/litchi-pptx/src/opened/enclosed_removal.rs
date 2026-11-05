//! Top-level shape removal on slides whose markup compatibility is enclosed
//! in shape-tree children (record 0303, follow-up of record 0290 F3).
//!
//! PowerPoint writes an embedded OLE object as a `p:graphicFrame` whose
//! `a:graphicData` holds an `mc:AlternateContent` (a VML choice and a
//! picture fallback). The shape reader resolves that markup, so the scene's
//! spans describe processed bytes, not the source. A whole top-level child
//! of `p:spTree` that encloses the markup can still be removed exactly:
//!
//! 1. the shape is selected and its dependencies are checked on the
//!    processed scene, as for any removal;
//! 2. the k-th processed root is the k-th source child of `p:spTree`
//!    (the two child lists must have the same length, and the processed
//!    children must be the scene's roots);
//! 3. the candidate is the source with that child's bytes deleted, and its
//!    shapes must be the source's shapes without the removed ones, in order
//!    and by identity, so the deleted child carried exactly them.
//!
//! Replay and publication prove a pair of slide images the same way: one is
//! the other with whole top-level children deleted.

use std::collections::BTreeSet;
use std::ops::Range;

use litchi_core::xml::ReaderOrigin;
use quick_xml::events::Event;
use quick_xml::reader::NsReader;

use super::xml::{
    MAX_XML_DEPTH, is_presentation_namespace, position, shape_removal_plan_in, span_bytes,
    stage_shape_removal_with_limits, validate_shape_removal_dependencies,
};
use crate::presentation::transition::{SlideMce, admit_source_with_limits, enclosed_edit_refusal};
use crate::{Error, Result};

/// The source spans of the shape children of `p:sld/p:cSld/p:spTree`, in
/// document order. The tree's own properties and extension list are not
/// shapes.
fn shape_tree_children(xml: &[u8]) -> Result<Vec<Range<usize>>> {
    let mut reader = NsReader::from_reader(xml);
    let origin = ReaderOrigin::of(xml);
    let mut depth = 0usize;
    // Whether the open elements at depths 2 and 3 are `p:cSld` and its
    // `p:spTree`.
    let mut common_slide = false;
    let mut shape_tree = false;
    let mut open_child: Option<usize> = None;
    let mut children = Vec::new();
    loop {
        let start = position(&reader, origin)?;
        let (namespace, event) = reader
            .read_resolved_event()
            .map_err(|error| Error::Xml(error.to_string()))?;
        let pml = is_presentation_namespace(&namespace);
        let end = position(&reader, origin)?;
        match event {
            Event::Start(element) => {
                depth += 1;
                if depth > MAX_XML_DEPTH {
                    return Err(Error::Limit {
                        kind: crate::opened::LimitKind::Codec(
                            crate::opened::CodecLimitKind::XmlDepth,
                        ),
                        observed: depth,
                        maximum: MAX_XML_DEPTH,
                    });
                }
                let local = element.local_name();
                match depth {
                    2 => common_slide = pml && local.as_ref() == b"cSld",
                    3 => shape_tree = common_slide && pml && local.as_ref() == b"spTree",
                    4 if shape_tree && !is_tree_property(pml, local.as_ref()) => {
                        open_child = Some(start);
                    },
                    _ => {},
                }
            },
            Event::Empty(element) => {
                if depth == 3 && shape_tree && !is_tree_property(pml, element.local_name().as_ref())
                {
                    children.push(start..end);
                }
            },
            Event::End(_) => {
                match depth {
                    2 => common_slide = false,
                    3 => shape_tree = false,
                    4 => {
                        if let Some(start) = open_child.take() {
                            children.push(start..end);
                        }
                    },
                    _ => {},
                }
                depth = depth
                    .checked_sub(1)
                    .ok_or_else(|| Error::Xml("unbalanced slide XML".to_owned()))?;
            },
            Event::Eof => break,
            _ => {},
        }
    }
    Ok(children)
}

/// `p:nvGrpSpPr`, `p:grpSpPr` and `p:extLst` of the shape tree.
fn is_tree_property(pml: bool, local: &[u8]) -> bool {
    pml && matches!(local, b"nvGrpSpPr" | b"grpSpPr" | b"extLst")
}

fn without(xml: &[u8], spans: &[Range<usize>]) -> Result<Vec<u8>> {
    let mut output = Vec::new();
    output
        .try_reserve_exact(xml.len())
        .map_err(|source| Error::Allocation {
            resource: "enclosed shape removal XML",
            source,
        })?;
    let mut cursor = 0usize;
    for span in spans {
        if span.start < cursor {
            return Err(refusal("removed shape-tree children overlap"));
        }
        output.extend_from_slice(span_bytes(xml, cursor..span.start)?);
        cursor = span.end;
    }
    output.extend_from_slice(span_bytes(xml, cursor..xml.len())?);
    Ok(output)
}

fn refusal(reason: &'static str) -> Error {
    Error::refused(
        litchi_core::edit::RefusalKind::UnsupportedContent,
        "remove_shape",
        reason,
    )
}

pub(crate) fn stage_enclosed_shape_removal_with_limits(
    xml: &[u8],
    key: crate::shape::Key<'_>,
    max_output_bytes: usize,
    shape_limits: crate::shape::Limits,
) -> Result<(Option<Vec<u8>>, usize)> {
    let scene = crate::shape::Scene::read(xml, shape_limits)?;
    if !scene.is_rewritten() {
        return stage_shape_removal_with_limits(xml, key, max_output_bytes, shape_limits);
    }
    let processed = scene.owner_xml();
    let (processed_span, removed, shape_count) = shape_removal_plan_in(&scene, processed, key)?;
    let processed_children = shape_tree_children(processed)?;
    let roots = scene
        .roots()
        .map(|root| root.common().span()?.range(processed.len()))
        .collect::<Result<Vec<_>>>()?;
    if processed_children != roots {
        return Err(refusal(
            "the processed shape tree has children that are not shapes",
        ));
    }
    let children = shape_tree_children(xml)?;
    if children.len() != roots.len() {
        return Err(refusal(
            "markup compatibility changes the slide's top-level shapes",
        ));
    }
    let index = roots
        .iter()
        .position(|root| *root == processed_span)
        .ok_or_else(|| refusal("shape removal requires a direct top-level shape"))?;
    let span = children[index].clone();
    if xml.len() - span.len() > max_output_bytes {
        return Err(Error::Limit {
            kind: crate::opened::LimitKind::Codec(crate::opened::CodecLimitKind::XmlBytes),
            observed: xml.len() - span.len(),
            maximum: max_output_bytes,
        });
    }
    let candidate = without(xml, std::slice::from_ref(&span))?;
    validate_shape_removal_dependencies(&candidate, &removed)?;
    // The deleted source child must carry exactly the removed shapes: the
    // candidate's shapes are the source's shapes without them, in order.
    let published = crate::shape::Scene::read(&candidate, shape_limits)?;
    let mut retained = scene
        .iter()
        .filter(|shape| shape.id().is_none_or(|id| !removed.contains(&id)));
    for shape in published.iter() {
        if retained.next().map(|previous| previous.id()) != Some(shape.id()) {
            return Err(refusal(
                "removing the source shape does not remove exactly its processed shape",
            ));
        }
    }
    if retained.next().is_some() || published.len().checked_add(removed.len()) != Some(shape_count)
    {
        return Err(refusal(
            "removing the source shape does not remove exactly its processed shape",
        ));
    }
    Ok((Some(candidate), 1))
}

/// The shape identities of a scene.
fn identities(scene: &crate::shape::Scene<'_>) -> Result<BTreeSet<u32>> {
    let mut ids = BTreeSet::new();
    for shape in scene.iter() {
        let id = shape
            .id()
            .ok_or_else(|| refusal("shape removal requires identifiable shapes"))?;
        if !ids.insert(id) {
            return Err(refusal("shape removal requires unique shape identities"));
        }
    }
    Ok(ids)
}

/// Prove a pair of slide images where at least one encloses markup
/// compatibility in shape-tree children: the smaller image must be the
/// larger with whole top-level shape-tree children deleted, and nothing left
/// in it may target a removed shape. Returns `false` when neither image
/// encloses markup compatibility, so ordinary proofs apply.
///
/// # Errors
///
pub(crate) fn validate_enclosed_removal_pair_with_limits(
    before: &[u8],
    after: &[u8],
    operation: &'static str,
    shape_limits: crate::shape::Limits,
    transition_limits: crate::transition::Limits,
) -> Result<bool> {
    let before_mce = admit_source_with_limits(before, operation, shape_limits, transition_limits)?;
    let after_mce = admit_source_with_limits(after, operation, shape_limits, transition_limits)?;
    if before_mce == SlideMce::Plain && after_mce == SlideMce::Plain {
        return Ok(false);
    }
    let (larger, smaller) = if before.len() >= after.len() {
        (before, after)
    } else {
        (after, before)
    };
    let larger_children = shape_tree_children(larger)?;
    let smaller_children = shape_tree_children(smaller)?;
    // Match the retained children in order by their exact bytes; the
    // unmatched children are the removed ones.
    let mut removed = Vec::new();
    let mut retained = smaller_children.iter().peekable();
    for child in &larger_children {
        let bytes = span_bytes(larger, child.clone())?;
        match retained.peek() {
            Some(next) if span_bytes(smaller, (*next).clone())? == bytes => {
                retained.next();
            },
            _ => removed.push(child.clone()),
        }
    }
    if retained.next().is_some() || removed.is_empty() || without(larger, &removed)? != smaller {
        return Err(enclosed_edit_refusal(operation));
    }
    let larger_ids = identities(&crate::shape::Scene::read(larger, shape_limits)?)?;
    let smaller_ids = identities(&crate::shape::Scene::read(smaller, shape_limits)?)?;
    let gone: BTreeSet<u32> = larger_ids.difference(&smaller_ids).copied().collect();
    if gone.is_empty() || !smaller_ids.is_subset(&larger_ids) {
        return Err(enclosed_edit_refusal(operation));
    }
    validate_shape_removal_dependencies(smaller, &gone)?;
    Ok(true)
}
