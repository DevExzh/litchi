//! Semantic Word shape facade.
//!
//! [`model`] owns the typed, owned DOC-facing shape tree, lossless unknown
//! records, and exact OfficeArt container snapshots. [`codec`] owns the
//! FIB/table-stream and textbox-story decoding required to populate it.

mod codec;
mod model;

pub use codec::{count_shapes, extract_drawing_shapes, extract_shape_text, extract_shapes};
pub use model::{
    Anchor, Bounds, ClientAnchor, Flags, Group, Kind, Native, PropertyTable, Shape, ShapeId,
    UnknownProperty, UnknownRecord,
};

pub(crate) use codec::{FIB_INDEX_DGG_INFO, extract_dgg_shapes};
