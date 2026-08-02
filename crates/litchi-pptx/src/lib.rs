//! Typed PowerPoint Open XML documents.
//!
//! The crate is being extracted one semantic capability at a time. The
//! [`transition`] module owns slide-transition values and their bounded
//! PresentationML codec. Package anchoring remains in the migration host for
//! now.

#![forbid(unsafe_code)]

mod error;
pub mod transition;

pub use error::{Error, Result};
