//! `OpenDocument` Drawing support with semantic responsibility layers.
//!
//! [`Package`] is the immutable owner of one ODG/OTG package and owns every
//! state verb (open, edit, apply, save). [`Drawing`] is its read-only
//! projection. The primary transaction family lives in [`edit`].
#![forbid(unsafe_code)]
#![allow(
    clippy::arbitrary_source_item_ordering,
    reason = "drawing package helpers stay in manifest and content traversal order"
)]

mod authoring;
mod codec;
mod facade;
pub mod flat;
pub mod metadata;
mod model;
mod package;
#[path = "model/transition.rs"]
pub mod transition;

pub use facade::edit;
pub use facade::{Builder, Drawing};
pub use litchi_odf_common::{
    Error, LimitKind, OpenOptions, SemanticLimitKind, SemanticLimits, SemanticLimitsBuilder,
    SourceLimitKind, SourceLimits, SourceLimitsBuilder,
};
pub use litchi_odf_common::{
    core::Profile as EncryptionProfile,
    signature::{
        CanonicalizationAlgorithm, DigitalSignature, DigitalSignatures, DocumentSigner,
        SignatureAlgorithm, SignatureValidity, SignatureVerification,
    },
};
pub use model::{
    Contour, ContourKind, DrawingAttribute, DrawingAttributeNamespace, EnhancedGeometry,
    EnhancedGeometryChild, EnhancedGeometryChildKind, FormControl, GluePoint, ImageMap,
    ImageMapArea, ImageMapAreaShape, auxiliary, enhanced, form, group, layer, page, resource,
    shape, style, style_resource,
};
pub use package::Package;
pub use transition::{Sound as TransitionSound, Transition};
