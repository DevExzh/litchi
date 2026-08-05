//! PowerPoint `ExOleObjStg` storage and compression model.
//!
//! This is the PowerPoint host layer for persisted OLE, VBA, and ActiveX
//! payloads.  Generic CFB object discovery belongs to `litchi-ole-common`;
//! this module owns only the PowerPoint record that points at those bytes.

mod codec;
mod model;

#[cfg(test)]
mod tests;

pub use model::{Compression, Kind, Storage};

pub(crate) use model::{Metadata, Ref};
