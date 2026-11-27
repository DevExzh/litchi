//! Strict, inert metadata for legacy `PowerPoint` OLE objects and their
//! package-level edits.
//!
//! [`crate::Package::edit_objects`] stages edits of the live document's
//! embedded-object list and storages; see [`Edit`]. This module never loads an embedded storage, invokes COM, starts an OLE
//! server, follows a link, or executes object content.

mod codec;
mod edit;
mod model;
mod validation;

pub(crate) mod editor;

pub use edit::Edit;
pub(crate) use editor::Editor;
pub use model::{
    Collection, ColorFollow, ContainerKind, Control, Definition, DimensionPolicy, DrawAspect,
    EmbedPreferences, ExternalObject, LinkInfo, Metadata, ObjectSubtype, ObjectType, UnknownRecord,
    UpdateMode,
};

crate::scoped::family!("embedded-object");

#[cfg(test)]
mod tests;
