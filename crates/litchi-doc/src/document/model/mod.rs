//! Layered document state and semantic query facade.
//!
//! `Document` owns the parsed OLE2-backed state in [`state`], exposes the
//! contextual reading API from [`semantic`], and keeps table-shape invariants
//! in [`validation`]. The crate root re-exports the facade directly, so
//! callers never need to know this internal topology.

mod semantic;
mod state;
mod validation;

pub use state::Document;
