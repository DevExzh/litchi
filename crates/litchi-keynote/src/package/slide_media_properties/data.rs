//! Strict media-data witnesses for the media-properties transaction.
//!
//! Media properties live in the drawable archive, but admitting that archive
//! without checking its materialized assets would let a malformed Package
//! Metadata graph pass the semantic edit path.  The replacement owner already
//! has the source-order selector, lazy PackageMetadata closure, digest/length,
//! owner, archive-header, and ZIP-entry checks required here.  This adapter
//! deliberately reuses that path for both the content and (when present) the
//! poster instead of growing a second metadata parser in the properties
//! owner.

use litchi_core::Position;

use super::super::slide_media_replacement::{
    Error as SlideMediaDataError, LimitKind as SlideMediaDataLimitKind, MediaBudget,
    MediaBudgetUsage, MediaPart, read_selected_media, select_media,
};
use super::super::slide_movie_geometry::GeometryBudget;
use super::{Error, LimitKind, Package};
use crate::{MovieSelector, SlideSelector};

/// Borrowed materialized assets belonging to one selected MovieArchive.
///
/// The witness contains no native identifiers or generated values.  Its
/// lifetime is tied to the immutable package that supplied the bytes, and the
/// optional poster is absent for a validated audio control.
#[derive(Debug, Clone, Copy)]
pub(super) struct MediaAssets<'a> {
    pub(super) content: &'a [u8],
    pub(super) poster: Option<&'a [u8]>,
}

/// Validate the selected media's content and optional poster under the
/// caller's existing operation budget.
///
/// The replacement selector owns all physical admission rules.  We run it on
/// a private ledger, then debit the consuming geometry ledger on a copy before
/// publishing the debit.  A limit failure therefore leaves the caller's
/// budget and package state unchanged, while all work performed by the lazy
/// metadata passes remains accounted for.
pub(super) fn validate_selected_media_assets<'a>(
    package: &'a Package,
    slide: Position,
    movie: Position,
    operation_budget: &mut GeometryBudget,
) -> Result<MediaAssets<'a>, Error> {
    let wire = operation_budget.residual(package)?;
    let retained = operation_budget.remaining_retained()?;
    let scratch = operation_budget.remaining_scratch()?;
    let media_bytes = retained.min(scratch);
    let references = operation_budget.remaining_references()?;
    let allocations = operation_budget.remaining_allocations()?;
    let mut media_budget = MediaBudget::for_package_with_caps(
        package,
        wire.max_input_bytes(),
        wire.max_fields(),
        wire.max_nesting(),
        wire.max_rewrite_work(),
        references,
        allocations,
        media_bytes,
    )
    .map_err(map_media_error)?;
    let slide_selector = SlideSelector::position(slide);
    let movie_selector = MovieSelector::position(movie);
    let content_selection = select_media(
        package,
        slide_selector,
        movie_selector,
        MediaPart::Content,
        &mut media_budget,
    )
    .map_err(map_media_error)?;
    let content = read_selected_media(
        package,
        &content_selection,
        MediaPart::Content,
        &mut media_budget,
    )
    .map_err(map_media_error)?;

    let poster = if !content_selection.has_poster() {
        None
    } else {
        match select_media(
            package,
            SlideSelector::position(slide),
            MovieSelector::position(movie),
            MediaPart::Poster,
            &mut media_budget,
        ) {
            Ok(poster_selection) => {
                if !content_selection.same_identity(&poster_selection) {
                    return Err(Error::InvalidSource);
                }
                Some(
                    read_selected_media(
                        package,
                        &poster_selection,
                        MediaPart::Poster,
                        &mut media_budget,
                    )
                    .map_err(map_media_error)?,
                )
            },
            // A poster edge on an audio control is malformed.  An audio
            // control with no edge never enters this branch and is accepted.
            Err(SlideMediaDataError::Refused(crate::slide::media::data::Refusal::AudioPoster)) => {
                return Err(Error::InvalidSource);
            },
            Err(error) => return Err(map_media_error(error)),
        }
    };

    debit_geometry_budget(operation_budget, media_budget.usage())?;
    Ok(MediaAssets { content, poster })
}

fn debit_geometry_budget(
    operation_budget: &mut GeometryBudget,
    usage: MediaBudgetUsage,
) -> Result<(), Error> {
    let mut next = *operation_budget;
    next.source(usage.input_bytes)?;
    next.fields(usage.fields)?;
    next.nesting(usage.nesting)?;
    next.work(usage.work)?;
    next.references(usage.references)?;
    next.allocations(usage.allocations)?;
    // The replacement ledger's media-bytes axis covers transient metadata
    // decompression, bounded fact vectors, and copied names.  The returned
    // assets borrow the package, but the consuming owner keeps both retained
    // and scratch ceilings conservative for the full transaction.
    next.retained(usage.media_bytes)?;
    next.scratch(usage.media_bytes)?;
    *operation_budget = next;
    Ok(())
}

fn map_media_error(error: SlideMediaDataError) -> Error {
    match error {
        SlideMediaDataError::Refused(crate::slide::media::data::Refusal::UnsupportedSource) => {
            Error::Refused(crate::slide::media::properties::Refusal::UnsupportedSource)
        },
        SlideMediaDataError::Limit {
            kind,
            observed,
            maximum,
        } => Error::Limit {
            kind: map_limit_kind(kind),
            observed,
            maximum,
        },
        SlideMediaDataError::Allocation { amount } => Error::Allocation { amount },
        SlideMediaDataError::Refused(crate::slide::media::data::Refusal::EmptySlideName)
        | SlideMediaDataError::Refused(crate::slide::media::data::Refusal::SlideNameNotFound)
        | SlideMediaDataError::Refused(
            crate::slide::media::data::Refusal::SlidePositionNotFound { .. },
        )
        | SlideMediaDataError::Refused(crate::slide::media::data::Refusal::AmbiguousSelector)
        | SlideMediaDataError::Refused(
            crate::slide::media::data::Refusal::MoviePositionNotFound { .. },
        ) => Error::InvalidSource,
        SlideMediaDataError::Refused(crate::slide::media::data::Refusal::AudioPoster)
        | SlideMediaDataError::InvalidSource
        | SlideMediaDataError::Read
        | SlideMediaDataError::Refused(crate::slide::media::data::Refusal::EmptyReplacement)
        | SlideMediaDataError::Refused(crate::slide::media::data::Refusal::ReplacementTooLarge)
        | SlideMediaDataError::Refused(crate::slide::media::data::Refusal::ReplacementType)
        | SlideMediaDataError::Refused(crate::slide::media::data::Refusal::AssetNotReachable)
        | SlideMediaDataError::Refused(crate::slide::media::data::Refusal::AssetNotMaterialized)
        | SlideMediaDataError::Verification
        | SlideMediaDataError::PatchConflict => Error::InvalidSource,
    }
}

fn map_limit_kind(kind: SlideMediaDataLimitKind) -> LimitKind {
    match kind {
        SlideMediaDataLimitKind::InputBytes => LimitKind::InputBytes,
        SlideMediaDataLimitKind::OutputBytes => LimitKind::OutputBytes,
        SlideMediaDataLimitKind::Entries => LimitKind::Entries,
        SlideMediaDataLimitKind::EntryBytes => LimitKind::EntryBytes,
        SlideMediaDataLimitKind::TotalBytes => LimitKind::TotalBytes,
        SlideMediaDataLimitKind::Slides => LimitKind::Slides,
        SlideMediaDataLimitKind::References => LimitKind::References,
        SlideMediaDataLimitKind::MediaBytes => LimitKind::Scratch,
        SlideMediaDataLimitKind::WireFields => LimitKind::WireFields,
        SlideMediaDataLimitKind::WireNesting => LimitKind::WireNesting,
        SlideMediaDataLimitKind::WireWork => LimitKind::WireWork,
        SlideMediaDataLimitKind::Allocations => LimitKind::Allocations,
    }
}
