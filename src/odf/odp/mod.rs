//! OpenDocument Presentation (.odp) implementation.
//!
//! This module provides comprehensive support for parsing and working with
//! OpenDocument Presentation documents (.odp files), which are the open standard
//! equivalent of Microsoft PowerPoint presentations.

mod parser;
mod presentation;
mod slide;

pub use presentation::Presentation;
pub use slide::{Shape, Slide};
