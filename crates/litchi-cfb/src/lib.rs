//! Microsoft Compound File Binary (CFB / OLE2) container parser and writer.
//!
//! This crate provides the CFB substrate consumed by the legacy Office binary
//! format crates (`litchi-doc`, `litchi-ppt`, and `litchi-xls`) and by
//! encrypted OOXML package support.
//!
//! See `[MS-CFB]: Compound File Binary File Format` for the format spec.

#![allow(
    missing_docs,
    reason = "the VT_* property-type constants mirror self-explanatory MS-CFB specification names"
)]
#![allow(
    non_ascii_idents,
    reason = "zerocopy's RawDirectoryEntry derive expansion emits internal identifiers outside this crate's source"
)]

pub mod consts;
mod directory_name;
mod file;
mod overlay;
mod shared;
mod shared_bulk;
mod splice;
mod stream_move;
mod validation;
pub mod writer;

pub use file::{DirectoryEntry, OleError, OleFile, OleFileLimits, is_ole_file};
pub use overlay::{
    ArtifactFingerprint, ComposedOverlaySource, OutputProgress, OverlayError, OverlayLimits,
    PublishReport, SameLengthStreamOverlay, ValidatedOverlayPlan,
};
pub use shared::{SharedOleFile, SharedOleFileLimits};
pub use shared_bulk::{SharedOleBulkError, SharedOleBulkRead};
pub use splice::{SameLengthStreamSplice, StreamSpliceLimits};
pub use stream_move::{ExistingStreamMove, StreamMoveLimits, ValidatedStreamMovePlan};
pub use validation::{CfbValidationError, validate_source, validate_source_with_limits};
pub use writer::{
    OleWriter, SequentialOleWriter, SequentialWriteError, SequentialWriteProgress,
    SequentialWriteReport, SequentialWriterLimits, SequentialWriterOptions,
};

#[cfg(test)]
mod allocation_validation_tests;
#[cfg(test)]
mod directory_validation_tests;
#[cfg(test)]
mod overlay_tests;
#[cfg(test)]
mod splice_tests;
#[cfg(test)]
mod stream_move_tests;
#[cfg(test)]
mod validation_tests;
