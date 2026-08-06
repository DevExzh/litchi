//! Typed, inert inspection of MS-XLDM section 2.5 table metadata.
//!
//! The owner is intentionally layered: this facade exposes the stable typed
//! model, the codec owns bounded XML projection and exact-byte snapshots, and
//! validation owns the cross-file verification entry point. No metadata is
//! executed, decompressed, or used for I/O.

mod codec;
mod model;
mod validation;

#[cfg(test)]
mod tests;

pub use codec::{inspect, parse_file, write_file};
pub use model::{
    ColumnPolicy, DictionaryPolicy, HierarchyPolicy, MetadataClass, MetadataCollection,
    MetadataDataObject, MetadataError, MetadataFile, MetadataFileKind, MetadataMember,
    MetadataModel, MetadataObject, MetadataProperty, MetadataResult, RelationshipIndexKind,
    RelationshipPolicy,
};
pub use validation::validate_files;
