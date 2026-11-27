//! Shared ODF master-page and header/footer semantics.
//!
//! Package-family crates retain their own archive and facade orchestration;
//! this module owns the family-neutral style vocabulary and bounded XML codec.

pub mod content;
mod model;
pub mod reader;
pub mod region;
pub mod writer;

pub use model::{Child, ChildKind, Master};
pub use region::{Kind, Region};
