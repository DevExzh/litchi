//! Layered MS-DOC route-slip lifecycle facade.
//!
//! The binary `RouteSlip` owner remains below `parts`; this contextual facade
//! exposes its typed values and package/document lifecycle operations without
//! requiring callers to navigate FIB indexes or table-stream offsets.

pub use crate::parts::route_slip::{
    Commit, DeliveryOption, Diagnostics, Edit, EditKind, Error, FIB_INDEX_ROUTE_SLIP, LimitKind,
    Metadata, NarrowString, Patch, Protection, Recipient, RecipientSelectionError,
    RecipientSelector, Result, parse, parse_bytes, record, to_bytes,
};
