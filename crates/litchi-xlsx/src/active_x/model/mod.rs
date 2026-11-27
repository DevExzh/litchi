//! Layered `ActiveX` value models.
//!
//! The submodules keep worksheet placement, descriptor metadata, and package
//! payloads distinct while the parent re-exports the ergonomic semantic API.

mod control;
mod descriptor;
mod resource;

pub use control::{Control, ControlProperties, Controls, Marker, ObjectAnchor};
pub use descriptor::{Descriptor, Font, Persistence, Picture, Property, PropertyObject};
pub use resource::{Binary, ControlSet, LoadedControl, PreviewImage};
