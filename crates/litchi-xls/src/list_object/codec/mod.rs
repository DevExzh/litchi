//! Layered BIFF8 codecs for worksheet tables and List12 records.
//!
//! The binary seam owns bounded BIFF primitives and retained wire fragments,
//! the semantic seam translates `ListObject` models to and from Feature11/12
//! payloads, and the package seam emits the List12/following-record sequence.
//! This facade deliberately re-exports the small helper surface used by the
//! neighboring model, package, and regression-test owners.

mod binary;
mod package;
mod semantic;

pub(super) use binary::{
    PendingFeature, append_frt, append_string, parse_string, record, u16_at, u32_at, validate_frt,
    validate_frt_any,
};
