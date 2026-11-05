//! `OpenDocument` HTML Template support with semantic responsibility layers.
#![forbid(unsafe_code)]

mod authoring;
mod codec;
mod facade;
pub mod metadata;
mod model;
mod package;

pub use facade::{
    Block, Builder, CapabilityState, Commit, Diagnostics, Document, Edit, FormsChange,
    HeadingChange, History, InlineBlock, InlineChange, JoinError, JoinFailure, ListChange,
    MergeConflict, MergePlan, Package, ParagraphChange, Patch, ResourceChange, ResourceMember,
    ResourcePayloadChange, SecurityCapabilities, SecurityPolicy, SecurityReport, StructureChange,
    StructureKind, StructureLifecycleChange, StructureSelector, TextBody, TransferPlan,
    TransferPolicy, TransferSelector, ValidationCapabilities,
};
/// The family error, limit dimension and caller-selected ingress options of
/// the OTH package verbs.
pub use litchi_odf_common::{
    Error, LimitKind, OpenOptions, SemanticLimitKind, SemanticLimits, SemanticLimitsBuilder,
    SourceLimitKind, SourceLimits, SourceLimitsBuilder,
};
pub use model::block::Content as ContentBlock;
pub use model::{
    annotation, block, bookmark, change, field, form, formatting, frame, heading, index, inline,
    link, list, note, paragraph, resource, ruby, section, style, table,
};
