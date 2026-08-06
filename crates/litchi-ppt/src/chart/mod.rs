//! Neutral, inert chart views and bounded Graph-package transactions for
//! legacy PowerPoint OLE objects.
//!
//! PPT owns only external-object discovery, frame attribution, storage
//! decompression, and host metadata. `[MS-OGRAPH]` Workbook validation and
//! chart traversal belong to [`litchi_ograph`]. [`PackageEditor`] only replaces
//! an existing standalone Graph chart stream; linked objects are never opened.

mod codec;
mod model;
mod package;
mod transaction;

#[cfg(test)]
mod tests;

pub use model::{Chart, Excel, Failure, Frame, Graph, Info, Inventory, Kind};
pub use transaction::{PackageEditor, Snapshot};

pub(crate) use package::enumerate;
