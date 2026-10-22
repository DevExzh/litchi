//! Package implementation for legacy Word documents (.doc).

mod body;
mod codec;
pub mod edit;
pub mod metadata;
mod model;
pub mod property_set;
/// Low-level cursor-backed CFB access for format-specific callers.
pub(crate) mod raw {
    pub use super::codec::{OpenOutcome, open_package};
    pub use super::model::Reader;
}
#[cfg(test)]
mod tests;

pub use body::{History, JoinedEdits, ThreeWayPlan};
pub use edit::{Commit, Diagnostics, Edit, EditOptions, LimitKind, Patch};
pub(crate) use model::preflight_declared_streams;
pub use model::{
    EncryptionKind, Error, OpenOptions, Package, Password, Result, SourceLimitKind, SourceLimits,
    SourceLimitsBuilder,
};

/// Collects one common OLE object package through its streaming output API.
///
/// The common owner deliberately exposes `write_to` as its only package
/// output operation.  DOC's legacy package layers still need an owned buffer
/// at a few CFB hand-off boundaries, so keep that collection in one private
/// adapter instead of reintroducing `to_bytes`/`finish` aliases on the owner.
pub(crate) fn object_package_bytes(
    package: &litchi_ole_common::object::Package,
) -> Result<Vec<u8>> {
    let mut bytes = Vec::new();
    package.write_to(&mut bytes).map_err(Error::from)?;
    Ok(bytes)
}

/// Commits a cloned common OLE edit and collects its immutable package output.
pub(crate) fn object_edit_bytes(editor: &litchi_ole_common::object::Edit) -> Result<Vec<u8>> {
    let commit = editor.clone().commit().map_err(Error::from)?;
    object_package_bytes(commit.package())
}
