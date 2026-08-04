//! Typed PowerPoint Open XML documents.
//!
//! The crate is being extracted one semantic capability at a time. The
//! transition module owns slide-transition values and its bounded codec. The
//! laser module owns inert laser-trace values and their bounded codec. The
//! font module owns embedded-font values and atomic package CRUD. The tag
//! module owns inert programmable tag lists and package CRUD. The notes module
//! owns bounded speaker-notes graphs, text encoding, and transactional package
//! mutation.
//! [`table::style`] owns typed table-style catalogs and their package graph.

#![forbid(unsafe_code)]

mod error;
pub mod font;
pub mod format;
pub mod hyperlinks;
pub mod laser;
pub mod notes;
pub mod shape;
pub mod table;
pub mod tag;
pub mod time;
pub mod transition;

pub use error::{Error, Result};
pub use format::{ImageFormat, TextFormat};
pub use hyperlinks::Hyperlink;
