//! `OpenDocument` HTML Template support with semantic responsibility layers.
#![forbid(unsafe_code)]

mod authoring;
mod codec;
mod facade;
mod model;
mod package;

pub use facade::{
    Block, Builder, Commit, Edit, HeadingChange, History, JoinError, JoinFailure, ListChange,
    MergeConflict, MergePlan, ParagraphChange, Patch, SecurityPolicy, SecurityReport, Template,
    TextBody, TransferPlan, TransferPolicy, TransferSelector,
};
pub use model::block::Content as ContentBlock;
pub use model::{
    block, bookmark, field, form, formatting, heading, link, list, paragraph, resource, style,
};
