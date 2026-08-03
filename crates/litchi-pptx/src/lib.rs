//! Typed PowerPoint Open XML documents.
//!
//! The crate is being extracted one semantic capability at a time. The
//! [`transition`] owns slide-transition values and its bounded codec. [`font`]
//! owns embedded-font values and atomic package CRUD. [`tag`]
//! owns inert programmable tag lists and package CRUD. [`notes`] owns bounded
//! speaker-notes graphs, text encoding, and transactional package mutation.
//! [`table::style`] owns typed table-style catalogs and their package graph.

#![forbid(unsafe_code)]

mod error;
pub mod font;
pub mod notes;
pub mod shape;
pub mod table;
pub mod tag;
pub mod transition;

pub use error::{Error, Result};
