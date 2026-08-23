//! Checked, borrowed `OfficeArt` BLIP stores and image records.
//!
//! This facade owns the format grammar shared by DOC, PPT, and XLS.  The
//! implementation is layered by responsibility: `model` contains borrowed
//! typed objects, `codec` performs wire decoding and lazy traversal, and
//! `validation` enforces cross-record `OfficeArt` invariants.  Image payloads
//! are never decompressed or rendered here.

mod codec;
mod file;
mod model;
mod validation;

#[cfg(test)]
mod tests;

pub mod write;

pub use file::{File, all, container, delay, get, record, record_with_delay, scan, store};
pub use model::{
    Bitmap, Blip, Block, Blocks, Compression, Context, Delay, DelayBlocks, Embedded, Entry, Id,
    JpegFlavor, Kind, Limits, Meta, MetaHeader, Name, Offset, Opaque, Point, Rect, Storage, Store,
    Uid, Uids,
};
