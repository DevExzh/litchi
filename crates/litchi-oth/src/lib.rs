//! `OpenDocument` HTML Template support with semantic responsibility layers.
#![forbid(unsafe_code)]

mod authoring;
mod codec;
mod facade;
mod model;
mod package;

pub use facade::{
    Block, Builder, Commit, Edit, History, JoinError, JoinFailure, ParagraphChange, Patch,
    Template, TextBody,
};
pub use model::block::Content as ContentBlock;
pub use model::{
    block, bookmark, field, form, formatting, heading, link, list, paragraph, resource, style,
};
