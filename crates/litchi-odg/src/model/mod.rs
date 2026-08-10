//! Immutable semantic values for this document family.

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

pub use form::Control as FormControl;
