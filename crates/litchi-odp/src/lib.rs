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
pub mod content;
pub mod facade;
pub mod handout_master;
pub mod model;
pub mod package;
pub mod streaming;

pub use facade::settings as show;
pub use facade::slide::{Shape, Slide};
pub use facade::{
    Builder, FlatPresentation, MasterPage, Presentation, SlideCatalogEntry,
    SourceBackedPresentation, SourceBackedPresentationCatalog,
};
pub use facade::{edit, embedded, image, layout, master, page, slide};
pub use package::ReadLimits;
pub use package::{
    SourceBackedTailAppendEdit, SourceBackedTailAppendPublicationPlan, TailAppendError,
    TailAppendLimits, TailAppendSourceProof,
};

// Keep implementation modules ergonomic internally without flattening their
// semantic vocabulary into the public crate root.
#[allow(
    clippy::wildcard_imports,
    reason = "internal-only glob keeps model vocabulary ergonomic without affecting the public API"
)]
pub(crate) use model::*;

pub use litchi_odf_common::core;
pub use litchi_odf_common::drawing;
pub use litchi_odf_common::rdf;
pub use litchi_odf_common::{constants, datatype, namespace};

// Bounded fresh plain-slide publication.
pub use streaming::{
    PlainSlide, PublicationError, PublicationFailureKind, SlideStreamReport, StreamingError,
    StreamingLimits, XmlAuditLimits, stream_plain_slides_to, try_stream_plain_slides_to,
};
