//! PowerPoint `ExOleObjStg` storage and compression model.
//!
//! This is the PowerPoint host layer for persisted OLE, VBA, and ActiveX
//! payloads.  Generic CFB object discovery belongs to `litchi-ole-common`;
//! this module owns only the PowerPoint record that points at those bytes.

mod codec;
mod editor;
mod model;
mod snapshot;

#[cfg(test)]
mod tests;

pub use editor::Editor;
pub use model::{Compression, Kind, MAX_DECLARED_BYTES, MAX_STORED_BYTES, Storage};
pub use snapshot::{Metadata, Snapshot};

#[cfg(any(test, feature = "vba-inspection"))]
#[allow(unused_imports)]
pub(crate) use model::Ref;
