//! Contextual pivot reader facade.
//!
//! Package traversal, XML codecs, semantic parser state, and validation are
//! kept behind this small facade so callers only depend on the four stable
//! reader operations exported by `pivot`.

mod codec;
mod model;
mod package;
mod validation;

#[cfg(test)]
mod tests;

pub use codec::{
    read_pivot_cache_definition, read_pivot_cache_records, read_pivot_table_definition,
};
pub use package::read_pivot_tables;
