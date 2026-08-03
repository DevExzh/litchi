//! Canonical WordprocessingML (`.docx`) APIs.
//!
//! The concise modules own format semantics while [`litchi_opc`] remains the
//! explicit low-level package graph.

#![forbid(unsafe_code)]

mod error;

pub mod alt;
pub mod color;
pub mod font;
pub mod web;

pub use error::{Error, Result};
