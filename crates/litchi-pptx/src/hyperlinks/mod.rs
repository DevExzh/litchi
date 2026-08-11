//! Layered, inert `PresentationML` hyperlink values.
//!
//! `model` owns the contextual target variants and constructors. `codec`
//! parses the small target grammar with bounded strings; it never follows a
//! relationship or accesses package storage.

mod codec;
mod model;

#[cfg(test)]
mod tests;

pub use model::Hyperlink;
