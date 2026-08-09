//! Central validation facade for semantic web-extension models.

mod model;

#[allow(
    unused_imports,
    reason = "the validation facade preserves model validators for sibling owners"
)]
pub(in crate::web) use model::*;
