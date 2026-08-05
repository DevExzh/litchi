//! Neutral, inert chart views for legacy PowerPoint OLE objects.
//!
//! PPT owns only external-object discovery, frame attribution, storage
//! decompression, and host metadata. `[MS-OGRAPH]` Workbook validation and
//! chart traversal belong to [`litchi_ograph`]. Linked objects are never opened.

mod codec;
mod model;
mod package;

#[cfg(test)]
mod tests;

pub use model::{Chart, Excel, Failure, Frame, Graph, Info, Inventory, Kind};

pub(crate) use package::enumerate;
