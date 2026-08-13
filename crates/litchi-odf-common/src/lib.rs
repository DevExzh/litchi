//! Shared `OpenDocument` vocabulary and scalar codecs.
//!
//! This crate contains functionality shared by every ODF document family:
//! constants, spreadsheet coordinates, namespace/qualified-name vocabulary,
//! and lexical data types.

#![forbid(unsafe_code)]

pub mod annotation;
pub mod calculation;
pub mod chart;
pub mod compact_xml;
pub mod constants;
pub mod coordinates;
pub mod core;
pub mod datatype;
pub mod detect;
pub mod drawing;
pub mod embedded;
pub mod media;
pub mod namespace;
pub mod package;
pub mod rdf;
pub mod repair;
pub mod signature;
pub mod style;
pub mod validation;

pub use repair::{
    MIMETYPE_LOCAL_EXTRA_ISSUE, MIMETYPE_LOCAL_EXTRA_REPAIR, MimetypeRepairPlan, OdfRepairLimits,
    OutputProgress as RepairOutputProgress, RemoveMimetypeLocalExtra, RepairError,
    RepairFingerprint, RepairPublication, plan_mimetype_local_extra, plan_mimetype_repair,
};
pub use validation::{
    DEFAULT_ODF_VALIDATION_LIMITS, OdfValidationError, OdfValidationLimits, validate_package,
    validate_package_with_limits,
};
