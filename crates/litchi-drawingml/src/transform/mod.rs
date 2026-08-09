//! Typed `DrawingML` two-dimensional transforms.
//!
//! [`Transform`] owns the format-neutral `a:CT_Transform2D` vocabulary from
//! `[MS-ODRAWXML]` and the `DrawingML` core schema: offsets, extents, group
//! child coordinate spaces, flips, and rotation. DOCX, PPTX, XLSX, and XLSB
//! retain their shape, anchor, and package semantics and consume this owner
//! for the shared `a:xfrm` subtree.
//!
//! The detached [`Snapshot`] retains the exact source bytes for no-op reads
//! and edits. A changed transaction emits only the modeled, validated
//! transform grammar; unsupported children or attributes are rejected instead
//! of being silently discarded.

pub mod codec;
mod model;
mod transaction;
mod validation;

#[cfg(test)]
mod tests;

pub use model::{Angle, Point, Size, Transform};
pub use transaction::{Commit, Patch, Snapshot, Transaction};
pub use validation::validate;
