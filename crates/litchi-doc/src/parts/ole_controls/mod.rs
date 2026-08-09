//! Inert Word OLE-control and `ObjectPool` metadata.
//!
//! This context owns the table records described by [MS-DOC] sections
//! 2.1.4, 2.9.161, 2.9.165-2.9.167, and 2.9.229. It does not resolve,
//! activate, render, or otherwise execute a control.

mod codec;
mod editor;
mod model;
mod package;
mod validation;

#[cfg(test)]
mod tests;

pub use codec::{parse_bytes, parse_metadata, to_bytes, to_metadata_bytes};
pub use editor::{Change, Commit, Editor, Patch};
pub use model::{
    ActiveX, EPRINT_STREAM, Entry, FieldCounts, Flags, Format, Metadata, OBJ_INFO_STREAM,
    OBJECT_POOL_STORAGE, OCX_DATA_STREAM, ObjectPool, OcxInfo, PRINT_STREAM, Persist1, Persist2,
    RgxOcxInfo, StorageName, Story,
};
pub use package::{FIB_INDEX_PLC_OCX, parse, parse_object, parse_object_metadata, parse_pool};
