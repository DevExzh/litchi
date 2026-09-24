//! Bounded, inert OLE object ownership.
//!
//! The object layer owns only target-selected CFB storage capture and
//! transactional byte-preserving rewrites. It does not know which document
//! format named a storage, how a host classifies an object, or how any OLE
//! payload is activated. Captured stream bytes remain the source-preserving
//! representation; standard OLEDS link, presentation, and native-data
//! metadata also have inert typed projections so format crates do not need to
//! rescan CFB names or reimplement bounded stream validation.

mod cfb_path;
mod codec;
pub mod directory;
mod discovery;
mod editor;
pub mod link;
mod model;
mod patch;
mod snapshot;
pub mod target;

pub use directory::{EntryKind, Links, Metadata, Sid};
pub use discovery::discover;
#[cfg(feature = "performance-diagnostics")]
pub use editor::{CfbParseEvent, CfbParseOutcome};
pub use editor::{Editor, MAX_STREAM_REMOVALS, PreparedReplacement};
pub use link::Link;
pub use litchi_cfb::{SectorLayoutFallback, SectorLayoutPolicy, SectorLayoutReport};
pub use model::{Limits, Object, Objects, Storage, Stream};
pub use patch::{Commit, Patch};
pub use snapshot::Snapshot;
pub use target::{Target, Targets};
