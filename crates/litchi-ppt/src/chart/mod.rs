//! Neutral, inert chart views and bounded Graph-package drafts for
//! legacy `PowerPoint` OLE objects.
//!
//! PPT owns only external-object discovery, frame attribution, storage
//! decompression, and host metadata. `[MS-OGRAPH]` Workbook validation and
//! chart traversal belong to [`litchi_ograph`]. [`crate::builder::graph::Draft`] only replaces
//! an existing standalone Graph chart stream; linked objects are never opened.

/// Host-neutral typed MS-OGRAPH semantic chart model and validation
/// primitives.
///
/// The types are re-exported here so callers can build producer-specific
/// semantic requests without reaching through the PPT host implementation.
/// Fresh encoding and mutation of parsed charts still follow the proof
/// boundaries documented by `litchi-ograph`; an untouched parsed chart can be
/// replayed losslessly.
pub mod semantic {
    pub use super::model::Models;
    pub use litchi_ograph::chart::{
        Ai, Binding, Cache, CellRef, Context, Count, DataKind, Family, Group, GroupId, ItemsMut,
        Label, Legend, Link, Model, Order, Owner, Props, Raw, Rect, Role, RowCol, Series, Source,
        Value, ValueRef, XlValue, axis, cache, format, group, layout,
    };
}

mod codec;
mod model;
mod package;

#[cfg(test)]
mod tests;

pub use model::{Chart, Excel, Failure, Frame, Graph, Info, Inventory, Kind};

pub(crate) use package::enumerate;
