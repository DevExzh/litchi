//! Immutable semantic values for this document family.

pub mod auxiliary;
pub mod enhanced;
pub mod form;
pub mod group;
pub mod layer;
#[path = "style_resource.rs"]
pub mod named_resource;
pub mod page;
pub mod resource;
pub mod shape;
pub mod style;
pub use named_resource as style_resource;

pub use auxiliary::{Contour, ContourKind, GluePoint, ImageMap, ImageMapArea, ImageMapAreaShape};
pub use enhanced::{
    DrawingAttribute, DrawingAttributeNamespace, EnhancedGeometry, EnhancedGeometryChild,
    EnhancedGeometryChildKind,
};
pub use form::Control as FormControl;
