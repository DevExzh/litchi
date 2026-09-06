//! Physical-object proof for releasing a removed identifier suffix.
//!
//! The metadata codec can rewrite the scalar watermark, but only the format
//! owner has the immutable package needed to prove that no surviving archive
//! object occupies the released range. Shared objects omitted from the actual
//! removal set participate in this census like every other survivor.

use super::budget::LifecycleBudget;
use super::{Package, SlideMediaLifecycleError};

/// Match the host allocation policy only when the current watermark itself is
/// physically removed. An unused metadata-only suffix is otherwise preserved.
pub(super) fn plan_release(
    source: &Package,
    removed_ids: &[u64],
    expected_last_identifier: u64,
    budget: &mut LifecycleBudget,
) -> Result<Option<u64>, SlideMediaLifecycleError> {
    budget.charge_wire_work(removed_ids.len().max(1))?;
    if removed_ids.contains(&0) || removed_ids.windows(2).any(|window| window[0] >= window[1]) {
        return Err(SlideMediaLifecycleError::InvalidSource);
    }
    if removed_ids
        .binary_search(&expected_last_identifier)
        .is_err()
    {
        return Ok(None);
    }

    let mut maximum_remaining = 0;
    let mut removed_last = false;
    for component in source.state.source.components().iter() {
        let objects = &component.archive().objects;
        budget.charge_entries(objects.len())?;
        let work = objects
            .len()
            .checked_mul(usize::BITS as usize)
            .ok_or(SlideMediaLifecycleError::InvalidSource)?;
        budget.charge_wire_work(work.max(1))?;
        for object in objects {
            let identifier = object
                .archive_info
                .identifier
                .filter(|identifier| *identifier != 0)
                .ok_or(SlideMediaLifecycleError::InvalidSource)?;
            if removed_ids.binary_search(&identifier).is_ok() {
                removed_last |= identifier == expected_last_identifier;
                continue;
            }
            if identifier >= expected_last_identifier {
                return Err(SlideMediaLifecycleError::InvalidSource);
            }
            maximum_remaining = maximum_remaining.max(identifier);
        }
    }
    if !removed_last || maximum_remaining == 0 {
        return Err(SlideMediaLifecycleError::InvalidSource);
    }
    Ok(Some(maximum_remaining))
}

/// Verify the emitted package still has exactly the physical maximum used by
/// the source-bound release proof. Metadata scalar readback is separately
/// verified by the neutral rewrite and the focused metadata adapter.
pub(super) fn verify_release(
    candidate: &Package,
    new_last_identifier: u64,
    budget: &mut LifecycleBudget,
) -> Result<(), SlideMediaLifecycleError> {
    let mut maximum_remaining = 0;
    for component in candidate.state.source.components().iter() {
        let objects = &component.archive().objects;
        budget.charge_entries(objects.len())?;
        budget.charge_wire_work(objects.len().max(1))?;
        for object in objects {
            let identifier = object
                .archive_info
                .identifier
                .filter(|identifier| *identifier != 0)
                .ok_or(SlideMediaLifecycleError::Verification)?;
            maximum_remaining = maximum_remaining.max(identifier);
        }
    }
    if maximum_remaining != new_last_identifier {
        return Err(SlideMediaLifecycleError::Verification);
    }
    Ok(())
}
