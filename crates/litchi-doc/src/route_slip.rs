//! Layered MS-DOC route-slip lifecycle facade.
//!
//! The binary `RouteSlip` owner remains below `parts`; this contextual facade
//! exposes its typed values and package/document lifecycle operations without
//! requiring callers to navigate FIB indexes or table-stream offsets.

pub use crate::parts::route_slip::{
    Commit, DeliveryOption, EditKind, Editor, FIB_INDEX_ROUTE_SLIP, Metadata, NarrowString,
    PackageCommit, PackageSnapshot, Patch, Protection, Recipient, RecipientSelectionError,
    RecipientSelector, Snapshot, Transaction, TransactionError, parse, parse_bytes, to_bytes,
};
