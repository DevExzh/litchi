//! Bounded proof for the two private style edges owned by a movie caption.
//!
//! A source-built Keynote movie may list the styles used by its title and
//! caption in the movie archive header even though the `TSD.MovieArchive`
//! payload only carries the two caption-info edges.  The lifecycle graph must
//! admit those edges only after following the known schema path all the way to
//! a `TSWP.ShapeStyleArchive`.  This module is intentionally narrow: it does
//! not discover arbitrary descendants or treat unknown archive-header edges as
//! proof of ownership.

#![allow(
    clippy::arbitrary_source_item_ordering,
    reason = "The witness keeps schema admission and its fixed-size result together."
)]

use litchi_iwa_common::WireLimits;
use litchi_iwa_core::ArchiveObject;
use litchi_iwa_protos::{keynote_movie_caption_codec, pages_movie_caption_codec};

use super::budget::LifecycleBudget;
use super::{Package, SlideMediaLifecycleError};

const MOVIE_MESSAGE_TYPE: u32 = 3_007;
const CAPTION_INFO_MESSAGE_TYPE: u32 = 633;
const SHAPE_STYLE_MESSAGE_TYPE: u32 = 2_025;
const STANDIN_MESSAGE_TYPE: u32 = 3_097;

/// Prove the private style identifiers transitively owned by one movie.
///
/// The result has one slot for the title style and one for the caption style;
/// equal identifiers are de-duplicated.  `None` means that the corresponding
/// movie edge was absent.  The fixed result avoids an allocation while making
/// the maximum of two transitive witnesses explicit to callers.
pub(super) fn prove_movie_caption_style_witness(
    package: &Package,
    source_component: &str,
    source_movie: &ArchiveObject,
    limits: WireLimits,
    budget: &mut LifecycleBudget,
) -> Result<[Option<u64>; 2], SlideMediaLifecycleError> {
    let Some(movie_identifier) = source_movie.archive_info.identifier else {
        return Ok([None, None]);
    };
    let Some((actual_component, actual_movie)) = package.object_with_component(movie_identifier)
    else {
        return Ok([None, None]);
    };
    if actual_component != source_component || !std::ptr::eq(actual_movie, source_movie) {
        return Ok([None, None]);
    }
    if source_movie.messages.len() != source_movie.archive_info.message_infos.len() {
        return Ok([None, None]);
    }

    let mut movie_index = None;
    for (index, message) in source_movie.messages.iter().enumerate() {
        if message.type_ == MOVIE_MESSAGE_TYPE {
            if movie_index.is_some() {
                return Ok([None, None]);
            }
            movie_index = Some(index);
        }
    }
    let Some(movie_index) = movie_index else {
        return Ok([None, None]);
    };
    let Some(movie_payload) = source_movie
        .messages
        .get(movie_index)
        .map(|message| message.data.as_slice())
    else {
        return Ok([None, None]);
    };
    charge_codec_pass(movie_payload, budget)?;
    let Ok(movie_snapshot) = keynote_movie_caption_codec::decode_movie_caption(
        movie_payload,
        movie_decode_options(limits)?,
    ) else {
        return Ok([None, None]);
    };
    let Some(movie_info) = source_movie.archive_info.message_infos.get(movie_index) else {
        return Ok([None, None]);
    };
    charge_header_metadata(movie_info, budget)?;

    let mut styles = [None, None];
    admit_caption_style(
        package,
        source_component,
        movie_info,
        movie_identifier,
        movie_snapshot.title_identifier(),
        &mut styles,
        limits,
        budget,
    )?;
    admit_caption_style(
        package,
        source_component,
        movie_info,
        movie_identifier,
        movie_snapshot.caption_identifier(),
        &mut styles,
        limits,
        budget,
    )?;
    Ok(styles)
}

fn admit_caption_style(
    package: &Package,
    source_component: &str,
    movie_info: &litchi_iwa_core::MessageInfo,
    movie_identifier: u64,
    caption_identifier: Option<u64>,
    styles: &mut [Option<u64>; 2],
    limits: WireLimits,
    budget: &mut LifecycleBudget,
) -> Result<(), SlideMediaLifecycleError> {
    let Some(caption_identifier) = caption_identifier else {
        return Ok(());
    };
    if caption_identifier == 0
        || movie_info.data_references.contains(&caption_identifier)
        || movie_info
            .object_references
            .iter()
            .filter(|identifier| **identifier == caption_identifier)
            .count()
            != 1
    {
        return Ok(());
    }

    let Some((caption_component, caption_object)) =
        package.object_with_component(caption_identifier)
    else {
        return Ok(());
    };
    if caption_component != source_component
        || caption_object.messages.len() != caption_object.archive_info.message_infos.len()
    {
        return Ok(());
    }
    let Some(caption_index) = unique_caption_message_index(caption_object) else {
        return Ok(());
    };
    let Some(caption_message) = caption_object.messages.get(caption_index) else {
        return Ok(());
    };
    if caption_message.type_ == STANDIN_MESSAGE_TYPE {
        return Ok(());
    }
    let Some(caption_info) = caption_object.archive_info.message_infos.get(caption_index) else {
        return Ok(());
    };
    charge_header_metadata(caption_info, budget)?;

    charge_codec_pass(&caption_message.data, budget)?;
    let Ok(caption_snapshot) = pages_movie_caption_codec::decode_caption_info(
        &caption_message.data,
        caption_decode_options(limits)?,
    ) else {
        return Ok(());
    };
    if caption_snapshot.parent_identifier() != movie_identifier {
        return Ok(());
    }
    let Some(style_identifier) = caption_snapshot
        .style_identifier()
        .filter(|identifier| *identifier != 0)
    else {
        return Ok(());
    };

    // A style is a transitive witness only when this movie's aggregate
    // metadata explicitly carries the same identifier. External styles and
    // caption-only styles remain owned by their existing graph and therefore
    // deliberately produce no witness here.
    if movie_info.data_references.contains(&style_identifier)
        || movie_info
            .object_references
            .iter()
            .filter(|identifier| **identifier == style_identifier)
            .count()
            != 1
    {
        return Ok(());
    }

    let Some((style_component, style_object)) = package.object_with_component(style_identifier)
    else {
        return Ok(());
    };
    if style_component != source_component {
        return Ok(());
    }
    if unique_message_index(style_object, SHAPE_STYLE_MESSAGE_TYPE).is_err() {
        return Ok(());
    }
    if strict_style_header_reference(caption_info, style_identifier).is_err() {
        return Ok(());
    }

    if styles[0] == Some(style_identifier) || styles[1] == Some(style_identifier) {
        return Ok(());
    }
    let slot = styles
        .iter_mut()
        .find(|slot| slot.is_none())
        .ok_or(SlideMediaLifecycleError::InvalidSource)?;
    *slot = Some(style_identifier);
    Ok(())
}

/// Require exactly one aggregate style edge in a caption-info header.
///
/// Native source-built archives encode the edge in `MessageInfo.object_references`.
/// When a producer also carries field-level metadata, only the schema path
/// `[1, 1, 2]` (with the older `[1, 2]` spelling accepted for compatibility)
/// can contribute the same edge.  Data references and unrelated field paths
/// never establish ownership.
fn strict_style_header_reference(
    caption_info: &litchi_iwa_core::MessageInfo,
    style_identifier: u64,
) -> Result<u64, SlideMediaLifecycleError> {
    let aggregate_count = caption_info
        .object_references
        .iter()
        .filter(|identifier| **identifier == style_identifier)
        .count();
    if aggregate_count != 1 {
        return Err(SlideMediaLifecycleError::InvalidSource);
    }

    let mut field_count = 0usize;
    for field in &caption_info.field_infos {
        if field.data_references.contains(&style_identifier) {
            // A data edge can never prove private style ownership.
            return Err(SlideMediaLifecycleError::InvalidSource);
        }
        let references = field
            .object_references
            .iter()
            .filter(|identifier| **identifier == style_identifier)
            .count();
        if references != 0 && !matches!(field.path.path.as_slice(), [1, 1, 2] | [1, 2]) {
            return Err(SlideMediaLifecycleError::InvalidSource);
        }
        field_count = field_count
            .checked_add(references)
            .ok_or(SlideMediaLifecycleError::InvalidSource)?;
    }
    if field_count > 1 {
        return Err(SlideMediaLifecycleError::InvalidSource);
    }
    Ok(style_identifier)
}

fn unique_message_index(
    object: &ArchiveObject,
    message_type: u32,
) -> Result<usize, SlideMediaLifecycleError> {
    let mut index = None;
    for (candidate, message) in object.messages.iter().enumerate() {
        if message.type_ == message_type {
            if index.is_some() {
                return Err(SlideMediaLifecycleError::InvalidSource);
            }
            index = Some(candidate);
        }
    }
    index.ok_or(SlideMediaLifecycleError::InvalidSource)
}

fn unique_caption_message_index(object: &ArchiveObject) -> Option<usize> {
    let mut index = None;
    for (candidate, message) in object.messages.iter().enumerate() {
        if message.type_ == CAPTION_INFO_MESSAGE_TYPE || message.type_ == STANDIN_MESSAGE_TYPE {
            if index.is_some() {
                return None;
            }
            index = Some(candidate);
        }
    }
    index
}

fn charge_codec_pass(
    payload: &[u8],
    budget: &mut LifecycleBudget,
) -> Result<(), SlideMediaLifecycleError> {
    let fields = payload.len().max(1);
    let work = payload.len().saturating_mul(32).max(1);
    budget.charge_wire_fields(fields)?;
    budget.charge_wire_work(work)?;
    budget.charge_allocation_plan(payload.len().max(1), 1)
}

fn charge_header_metadata(
    info: &litchi_iwa_core::MessageInfo,
    budget: &mut LifecycleBudget,
) -> Result<(), SlideMediaLifecycleError> {
    let fields = info
        .object_references
        .len()
        .checked_add(info.data_references.len())
        .and_then(|count| count.checked_add(info.field_infos.len()))
        .ok_or(SlideMediaLifecycleError::InvalidSource)?;
    let references = info
        .object_references
        .len()
        .checked_add(info.data_references.len())
        .and_then(|count| {
            info.field_infos.iter().try_fold(count, |count, field| {
                count
                    .checked_add(field.object_references.len())
                    .and_then(|count| count.checked_add(field.data_references.len()))
            })
        })
        .ok_or(SlideMediaLifecycleError::InvalidSource)?;
    budget.charge_wire_fields(fields)?;
    budget.charge_references(references)?;
    budget.charge_wire_work(fields.saturating_add(references))
}

fn movie_decode_options(
    limits: WireLimits,
) -> Result<keynote_movie_caption_codec::DecodeOptions, SlideMediaLifecycleError> {
    let recursion =
        u32::try_from(limits.max_nesting()).map_err(|_| SlideMediaLifecycleError::InvalidSource)?;
    Ok(keynote_movie_caption_codec::DecodeOptions::new(
        limits.max_input_bytes(),
        limits.max_fields(),
        limits.max_rewrite_work(),
        recursion,
    ))
}

fn caption_decode_options(
    limits: WireLimits,
) -> Result<pages_movie_caption_codec::DecodeOptions, SlideMediaLifecycleError> {
    let recursion =
        u32::try_from(limits.max_nesting()).map_err(|_| SlideMediaLifecycleError::InvalidSource)?;
    Ok(pages_movie_caption_codec::DecodeOptions::new(
        limits.max_input_bytes(),
        limits.max_fields(),
        limits.max_rewrite_work(),
        recursion,
    ))
}
