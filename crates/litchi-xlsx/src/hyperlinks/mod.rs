//! Typed, inert worksheet hyperlink projections.
//!
//! Hyperlink targets are metadata only. Parsing never resolves, opens, or
//! fetches an external target.

pub(crate) mod codec;
pub(crate) mod model;

pub use model::{Hyperlink, HyperlinkReference};

pub(crate) use codec::parse;
