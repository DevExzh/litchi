//! Typed, inert worksheet hyperlink projections.
//!
//! Hyperlink targets are metadata only. Parsing never resolves, opens, or
//! fetches an external target.

pub(crate) mod codec;
pub(crate) mod model;
mod patch;
mod snapshot;
pub mod source;

pub use model::{Hyperlink, HyperlinkReference};
pub(crate) use patch::Patch;
pub use snapshot::Snapshot;

pub(crate) use codec::parse;
