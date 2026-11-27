//! Shared checked-allocation plumbing for codec layers.

use litchi_cfb::OleError;

pub(super) fn allocation(
    resource: &'static str,
    source: std::collections::TryReserveError,
) -> OleError {
    OleError::Allocation { resource, source }
}
