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
mod error;
pub mod flat;
pub mod leniency;
pub mod limits;
pub mod media;
pub mod metadata;
pub mod namespace;
pub mod package;
pub mod publish;
pub mod rdf;
pub mod repair;
pub mod signature;
pub mod style;
pub mod validation;

pub use core::{
    ArchiveLimits, OwnedPackage, PreparedPackage, SourceBackedPackage, SourceMemberReaderError,
};
pub use core::{
    GeneratedXmlEnvelope, GeneratedXmlLimitExceeded, GeneratedXmlLimitResource, GeneratedXmlLimits,
    GeneratedXmlReport, XmlStreamLimits,
};
pub use core::{ResolvedReader, map_resolved_reader_error};
pub use core::{
    SourceContentInsertionCallbackError, SourceContentInsertionError, SourceContentInsertionPlan,
    SourceContentPublicationError, SourceContentPublicationOptions,
    SourceContentPublicationProgress, SourceContentPublicationReport, write_content_xml_to_stream,
    write_content_xml_to_stream_with_options,
};
pub use error::Error;
pub use leniency::{
    ContentDefect, Defect, DefectLocation, Leniency, PackageDefect, ToleranceReport,
    ToleratedDefect,
};
pub use limits::{
    LimitKind, OpenOptions, RawXmlLimitKind, SemanticLimitKind, SemanticLimits,
    SemanticLimitsBuilder, SourceLimitKind, SourceLimits, SourceLimitsBuilder,
};
pub use repair::{
    Destructive, MIMETYPE_LOCAL_EXTRA_ISSUE, MIMETYPE_LOCAL_EXTRA_REPAIR,
    MIMETYPE_REPAIR_PLAN_SCHEMA, MimetypeRepairPlan, NonDestructive, OdfRepairLimits,
    OutputProgress as RepairOutputProgress, RemoveMimetypeLocalExtra, RepairChangedRegion,
    RepairEffects, RepairError, RepairFingerprint, RepairIntent, RepairIntentKind,
    RepairInversePatch, RepairPatch, RepairPlan, RepairPreview, RepairPublication,
    plan_mimetype_local_extra, plan_mimetype_repair, plan_odf_repair, plan_repair,
};
pub use validation::{
    DEFAULT_ODF_VALIDATION_LIMITS, OdfValidationError, OdfValidationLimits, validate_package,
};
