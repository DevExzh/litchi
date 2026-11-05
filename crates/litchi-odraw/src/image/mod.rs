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

pub mod builder;
pub mod write;

pub(super) fn map_traversal_limit(error: crate::Error) -> crate::Error {
    match error {
        crate::Error::LimitExceeded {
            limit: crate::LimitKind::RecordBytes,
            observed,
            maximum,
        } => crate::Error::ImageLimitExceeded {
            limit: LimitKind::RecordBytes,
            observed,
            maximum: u64::from(maximum),
        },
        crate::Error::LimitExceeded {
            limit: crate::LimitKind::Depth,
            observed,
            maximum,
        } => crate::Error::ImageLimitExceeded {
            limit: LimitKind::Depth,
            observed,
            maximum: u64::from(maximum),
        },
        crate::Error::LimitExceeded {
            limit: crate::LimitKind::Records,
            observed,
            maximum,
        } => crate::Error::ImageLimitExceeded {
            limit: LimitKind::Records,
            observed,
            maximum: u64::from(maximum),
        },
        other => other,
    }
}

pub use file::{File, all, container, delay, get, record, record_with_delay, scan, store};
pub use model::{
    Bitmap, Blip, Block, Blocks, Compression, Context, Delay, DelayBlocks, Embedded, Entry, Id,
    JpegFlavor, Kind, LimitKind, Limits, Meta, MetaHeader, Name, Offset, Opaque, Point, Rect,
    Storage, Store, Uid, Uids,
};
