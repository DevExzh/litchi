//! Standalone `OpenDocument` Formula (`.odf` and `.otf`) support.
//!
//! The crate keeps `MathML` tree ownership, XML validation, package handling,
//! authoring, and the public facade in separate layers. Formula markup is
//! inert data: this crate never evaluates it.

#![forbid(unsafe_code)]

pub mod authoring;
pub mod codec;
pub mod facade;
pub mod model;
pub mod package;

pub use authoring::Builder;
pub use facade::Formula;
pub use model::{Attribute, Content, Element, Kind};
