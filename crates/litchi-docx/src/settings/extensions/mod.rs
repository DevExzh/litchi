//! Word 2010/2012 settings extensions.

mod codec;
mod model;
mod package;
mod transaction;
mod validation;

#[cfg(test)]
mod tests;

pub use model::{
    DocumentId, Extension, Extensions, Guid, MAX_EXTENSIONS, MAX_OPAQUE_BYTES, OnOff,
    OpaqueExtension, WORD_2010_NAMESPACE, WORD_2012_NAMESPACE,
};
pub use transaction::{Commit, Patch, Snapshot, Transaction};

pub(crate) use package::process_part;
