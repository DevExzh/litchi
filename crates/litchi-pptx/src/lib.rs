//! Typed PowerPoint Open XML documents.
//!
//! The crate is being extracted one semantic capability at a time. The
//! [`transition`] module owns slide-transition values and its bounded
//! PresentationML codec. [`tag`] owns inert programmable tag lists, their safe
//! CRUD model, and slide relationship discovery.

#![forbid(unsafe_code)]

mod error;
pub mod tag;
pub mod transition;

pub use error::{Error, Result};
