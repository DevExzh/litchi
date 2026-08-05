//! Deterministic RTF output owners.
//!
//! Each child module contributes an implementation slice to the same
//! [`RtfWriter`](super::RtfWriter) facade. Keeping the slices as inherent
//! implementations preserves the existing API while making the output order
//! and ownership boundaries explicit.

mod content;
mod drawing;
mod header;
mod metadata;
mod primitives;
mod resources;
mod story;
mod table;
