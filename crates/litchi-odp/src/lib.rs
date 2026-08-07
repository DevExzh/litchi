//! `OpenDocument` Presentation (`.odp`) support.
//!
//! The crate is organized by responsibility: semantic value types live in
//! [`model`], XML parsing in [`codec`], package access in [`package`], document
//! construction and mutation in [`authoring`], and the concise public surface
//! in [`facade`].

#![forbid(unsafe_code)]

pub mod annotation;
pub mod authoring;
pub mod charts;
pub mod codec;
pub mod facade;
pub mod handout_master;
pub mod model;
pub mod package;

pub use facade::slide::{Shape, Slide};
pub use facade::{Builder, MasterPage, Presentation};
pub use facade::{layout, master, page, slide};

// Keep implementation modules ergonomic internally without flattening their
// semantic vocabulary into the public crate root.
pub(crate) use model::*;

pub use litchi_odf_common::core;
pub use litchi_odf_common::rdf;
pub use litchi_odf_common::{constants, datatype, namespace};
