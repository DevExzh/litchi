//! `OpenDocument` HTML Template support with semantic responsibility layers.
#![forbid(unsafe_code)]

mod authoring;
mod codec;
mod facade;
mod model;
mod package;

pub use facade::{Block, Builder, Commit, Edit, ParagraphChange, Patch, Template, TextBody};
pub use model::{heading, link, paragraph};
