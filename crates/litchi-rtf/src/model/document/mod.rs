//! Retained native RTF document representation.

mod model;

#[cfg(test)]
mod tests;

pub use model::RtfDocument;
pub(crate) use model::owned_table;
