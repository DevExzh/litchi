//! Contextual publication of one legacy `PowerPoint` slide payload.
//!
//! The package owner is deliberately narrower than an OLE2 presentation
//! editor. It accepts one complete `SlideContainer`, discovers the slide's
//! `PPDrawing` sibling and the `___PPT10` `BinaryTagData` that owns its
//! `BuildList`, and publishes only fixed-width diagram metadata edits. The
//! surrounding slide record, opaque tag records, and `OfficeArt` bytes remain
//! source-owned.

mod codec;
mod model;
#[allow(
    dead_code,
    reason = "the slide-level owner behind the package-level Edit; its full surface stays exercised by the unit tests (record 0357)"
)]
mod transaction;
mod validation;

#[cfg(test)]
mod tests;

pub use model::{SlideLimits, SlideRevision, SlideSnapshot};
pub(crate) use transaction::SlideEditor;
