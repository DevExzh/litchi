//! Structural validation for the bounded XLSB timeline XML owner.

use super::model::{Cache, Views, validate_cache, validate_views};
use crate::package::error::Result;

pub(crate) fn cache(value: &Cache) -> Result<()> {
    validate_cache(value)
}

pub(crate) fn views(value: &Views) -> Result<()> {
    validate_views(value)
}
