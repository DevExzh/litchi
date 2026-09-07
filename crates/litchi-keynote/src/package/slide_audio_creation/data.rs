//! Borrowed media-data verification under the creation transaction's residual limits.

use litchi_core::Position;

use super::super::slide_media_replacement::{
    MediaBudget, MediaPart, SlideMediaDataError, SlideMediaDataLimitKind, read_selected_media,
    select_media,
};
use super::{CreationBudget, Package, SlideAudioCreationError, SlideAudioCreationLimitKind};
use crate::{MovieSelector, SlideSelector};

pub(super) fn read_content<'a>(
    package: &'a Package,
    slide: Position,
    movie: Position,
    budget: &mut CreationBudget,
) -> Result<&'a [u8], SlideAudioCreationError> {
    let mut media_budget = MediaBudget::for_package_with_caps(
        package,
        budget.max_input.saturating_sub(budget.input),
        budget.remaining_wire_fields(),
        budget.max_nesting,
        budget.remaining_work(),
        budget.max_references.saturating_sub(budget.references),
        budget.remaining_allocations(),
        budget.max_media_bytes.saturating_sub(budget.media_bytes),
    )
    .map_err(map_error)?;
    let selection = select_media(
        package,
        SlideSelector::position(slide),
        MovieSelector::position(movie),
        MediaPart::Content,
        &mut media_budget,
    )
    .map_err(map_error)?;
    let content = read_selected_media(package, &selection, MediaPart::Content, &mut media_budget)
        .map_err(map_error)?;
    let usage = media_budget.usage();
    budget.charge_input(usage.input_bytes)?;
    budget.charge_wire_fields(usage.fields)?;
    budget.charge_nesting(usage.nesting)?;
    budget.charge_work(usage.work)?;
    budget.charge_references(usage.references)?;
    budget.charge_allocation_events(usage.allocations)?;
    budget.charge_media_bytes(usage.media_bytes)?;
    Ok(content)
}

fn map_error(error: SlideMediaDataError) -> SlideAudioCreationError {
    match error {
        SlideMediaDataError::LimitExceeded {
            kind,
            observed,
            maximum,
        } => {
            let kind = match kind {
                SlideMediaDataLimitKind::InputBytes => SlideAudioCreationLimitKind::InputBytes,
                SlideMediaDataLimitKind::OutputBytes => SlideAudioCreationLimitKind::OutputBytes,
                SlideMediaDataLimitKind::Entries => SlideAudioCreationLimitKind::Entries,
                SlideMediaDataLimitKind::EntryBytes => SlideAudioCreationLimitKind::EntryBytes,
                SlideMediaDataLimitKind::TotalBytes => SlideAudioCreationLimitKind::TotalBytes,
                SlideMediaDataLimitKind::Slides => SlideAudioCreationLimitKind::Slides,
                SlideMediaDataLimitKind::References => SlideAudioCreationLimitKind::References,
                SlideMediaDataLimitKind::MediaBytes => SlideAudioCreationLimitKind::MediaBytes,
                SlideMediaDataLimitKind::WireFields => SlideAudioCreationLimitKind::WireFields,
                SlideMediaDataLimitKind::WireNesting => SlideAudioCreationLimitKind::WireNesting,
                SlideMediaDataLimitKind::WireWork => SlideAudioCreationLimitKind::WireWork,
                SlideMediaDataLimitKind::Allocations => SlideAudioCreationLimitKind::Allocations,
            };
            SlideAudioCreationError::LimitExceeded {
                kind,
                observed,
                maximum,
            }
        },
        SlideMediaDataError::Allocation { amount } => {
            SlideAudioCreationError::Allocation { amount }
        },
        _ => SlideAudioCreationError::Verification,
    }
}
