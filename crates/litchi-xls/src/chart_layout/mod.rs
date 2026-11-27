//! BIFF8 chart layout future records of the Chart Sheet substream
//! (MS-XLS 2.1):
//!
//! - **CrtLayout12** (0x089D): layout information for an attached label or
//!   legend (MS-XLS 2.4.66).
//! - **CrtLayout12A** (0x08A7): layout information for a plot area
//!   (MS-XLS 2.4.67).
//!
//! Everything in this module is INERT: layout values and checksums are stored
//! verbatim and no chart layout is computed or applied. The `dwCheckSum`
//! fields are preserved as read; the checksum inputs include data owned by
//! other records, so they are not recomputed here.
//!
//! # References
//!
//! - MS-XLS 2.4.66 (CrtLayout12), 2.4.67 (CrtLayout12A), 2.5.62
//!   (CrtLayout12Mode), 2.5.134 (FrtFlags), 2.5.135 (FrtHeader), 2.5.342
//!   (Xnum)

mod codec;
mod model;

#[cfg(test)]
mod tests;

pub use model::{CrtLayout12, CrtLayout12A, CrtLayout12Mode};
