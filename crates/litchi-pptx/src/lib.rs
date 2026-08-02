//! Typed PowerPoint Open XML documents.
//!
//! The crate is being extracted one semantic capability at a time. The
//! [`transition`] owns slide-transition values and its bounded codec. [`tag`]
//! owns inert programmable tag lists and package CRUD. [`notes`] owns bounded
//! speaker-notes graphs, text encoding, and transactional package mutation.

#![forbid(unsafe_code)]

mod error;
pub mod notes;
pub mod shape;
pub mod tag;
pub mod transition;

pub use error::{Error, Result};
