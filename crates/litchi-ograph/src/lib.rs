//! Host-neutral Microsoft Office Graph (`[MS-OGRAPH]`) primitives.
//!
//! The crate deliberately separates three layers:
//!
//! - [`raw`] provides bounded borrowed and move-owned BIFF record framing;
//! - [`record`] contains small typed record codecs;
//! - [`chart`] discovers standalone and Excel-hosted chart substreams;
//! - [`PackageRef`] and [`Package`] validate standalone OGraph compound files.
//!
//! Parsing preserves record order, unknown records, reserved bits, and unused
//! bytes. Editing hosts such as XLS and PPT remain responsible for their own
//! object IDs, directory entries, and embedding metadata.

#![forbid(unsafe_code)]

pub mod chart;
mod error;
mod limits;
mod package;
pub mod raw;
pub mod record;

pub use error::{Error, Result};
pub use limits::{Limits, MAX_BIFF_RECORD_BYTES};
pub use package::{Package, PackageRef, Payload, Topology, Workbook, WorkbookRef};
