//! Strict, inert metadata for legacy PowerPoint OLE objects.
//!
//! This module never loads an embedded storage, invokes COM, starts an OLE
//! server, follows a link, or executes object content.

mod codec;
mod model;
mod validation;

pub mod editor;

pub use editor::Editor;
pub use model::{
    Collection, ColorFollow, ContainerKind, Control, Definition, DimensionPolicy, DrawAspect,
    EmbedPreferences, ExternalObject, LinkInfo, Metadata, ObjectSubtype, ObjectType, UnknownRecord,
    UpdateMode,
};

#[cfg(test)]
mod tests;
