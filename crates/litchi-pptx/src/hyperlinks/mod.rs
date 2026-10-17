//! Layered, inert `PresentationML` hyperlink values.
//!
//! `model` owns the contextual target variants and constructors. `codec`
//! parses the small target grammar with bounded strings; it never follows a
//! relationship or accesses package storage. `link` reads and locates the
//! click actions of shapes and text runs (record 0321).

mod codec;
pub(crate) mod link;
mod model;

#[cfg(test)]
mod tests;

pub use link::{Click, Link, Target, mentions_click};
pub use model::Hyperlink;
