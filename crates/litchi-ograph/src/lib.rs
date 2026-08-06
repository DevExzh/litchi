//! Host-neutral Microsoft Office Graph (`[MS-OGRAPH]`) primitives.
//!
//! The crate deliberately separates three layers:
//!
//! - [`litchi_biff`] provides bounded borrowed and move-owned BIFF record framing;
//! - [`record`] contains small typed record codecs;
//! - [`chart`] discovers standalone and Excel-hosted chart substreams;
//! - [`chart::Editor`] provides bounded, source-checked cache, chart-area, and
//!   fixed-metadata patches;
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
pub mod record;

pub use error::{Error, Result};
pub use limits::Limits;
pub use package::{Package, PackageRef, Payload, Topology, Workbook, WorkbookRef};
