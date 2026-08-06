//! Bounded physical ingress for Apple iWork bundles.
//!
//! This crate owns only the ZIP container boundary: central-directory limits,
//! legacy nested `Index.zip` handling, and the checksum-free Snappy/IWA
//! component stream. It deliberately does not depend
//! on the iWork facade, semantic format crates, or the archive-neutral package
//! entry store.
//!
//! Application readers should consume [`ComponentCatalog::iter`] and perform
//! message decoding in their own adapter layer. Raw ZIP implementation types
//! remain private to this crate.

#![forbid(unsafe_code)]

mod catalog;
mod error;
mod limits;
pub mod package;
mod zip;

pub use catalog::{Component, ComponentCatalog};
pub use error::{Error, LimitKind, Result};
pub use limits::Limits;
