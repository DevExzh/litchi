//! Typed worksheet cell watches and phonetic defaults.
//!
//! The owner follows `[MS-XLSB]` sections 2.4.21, 2.4.331, 2.4.378, and
//! 2.4.744. [`worksheet`] owns the BIFF12 stream snapshot and transaction;
//! [`workbook`] owns package traversal and the relationship between a logical
//! worksheet selector and its binary part. No watch is evaluated or monitored
//! by litchi.

mod codec;
mod model;
pub mod phonetic;
mod validation;
pub mod workbook;
pub mod worksheet;

#[cfg(test)]
mod tests;

pub use model::{Reference, UnknownRecord, Watch, Watches};
pub use phonetic::{Alignment, Info, Type};
pub use worksheet::{Commit, Edit, Patch, Snapshot};

pub use crate::package::error::{Error, Result};

/// Maximum number of typed `BrtCellWatch` records in one worksheet.
pub const MAX_WATCHES: usize = 65_536;
/// Maximum number of opaque records retained inside one watch collection.
pub const MAX_OPAQUE_RECORDS: usize = 65_536;
/// Maximum aggregate payload bytes retained for opaque records in one
/// collection.
pub const MAX_OPAQUE_PAYLOAD: usize = 64 * 1024 * 1024;
/// Maximum number of BIFF12 records accepted in one worksheet snapshot.
pub const MAX_RECORDS: usize = 1_000_000;
