//! `OpenDocument` HTML Template support with semantic responsibility layers.
#![forbid(unsafe_code)]

mod authoring;
mod codec;
mod facade;
mod model;
mod package;

pub use facade::{
    Block, Builder, CapabilityState, Commit, Edit, FormsChange, HeadingChange, History,
    InlineBlock, InlineChange, JoinError, JoinFailure, ListChange, MergeConflict, MergePlan,
    ParagraphChange, Patch, ResourceChange, ResourceMember, ResourcePayloadChange,
    SecurityCapabilities, SecurityPolicy, SecurityReport, Template, TextBody, TransferPlan,
    TransferPolicy, TransferSelector, ValidationCapabilities,
};
pub use model::block::Content as ContentBlock;
pub use model::{
    block, bookmark, field, form, formatting, heading, inline, link, list, paragraph, resource,
    style,
};
