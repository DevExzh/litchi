//! Public validation boundary for workbook revision metadata.

use crate::error::Result;

use super::model::{RevisionHeaders, RevisionLog, RevisionUsers, Revisions};

/// Validate one complete revision package model before publication.
pub fn revisions(value: &Revisions) -> Result<()> {
    super::package::validate_value(value)
}

/// Validate the typed user catalog independently.
pub fn users(value: &RevisionUsers) -> Result<()> {
    super::model::validate_users(value)
}

/// Validate the typed header catalog independently.
pub fn headers(value: &RevisionHeaders) -> Result<()> {
    super::model::validate_headers(value)
}

/// Validate one inert revision log independently.
pub fn log(value: &RevisionLog) -> Result<()> {
    super::model::validate_log(value)
}
