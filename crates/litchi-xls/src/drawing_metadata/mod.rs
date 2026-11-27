//! Layered worksheet drawing metadata for legacy XLS.
//!
//! The XLS host owns the meaning of OfficeArt `ClientAnchor` records.  The
//! generic OfficeArt crate intentionally leaves that payload opaque, so this
//! owner keeps the wire codec, semantic model, and invariants together while
//! exposing only the small typed facade needed by worksheet callers.

mod codec;
mod model;
mod validation;

pub use model::{AnchorBehavior, AnchorPoint, SheetAnchor};

pub(crate) use codec::decode_sheet_anchor;

#[cfg(test)]
mod tests;
