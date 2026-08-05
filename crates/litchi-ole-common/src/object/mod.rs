//! Bounded, inert OLE object ownership.
//!
//! The object layer owns only target-selected CFB storage capture and
//! transactional byte-preserving rewrites. It does not know which document
//! format named a storage, how a host classifies an object, or how any OLE
//! payload is activated. The selected storage's streams remain raw bytes so
//! format crates can interpret their own metadata without a common-owned
//! classification leak.

mod codec;
mod discovery;
mod editor;
mod model;
pub mod target;

pub use discovery::discover;
pub use editor::Editor;
pub use model::{Limits, Object, Objects, Storage, Stream};
pub use target::{Target, Targets};
