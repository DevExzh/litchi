//! `OpenDocument` Drawing support with semantic responsibility layers.
#![forbid(unsafe_code)]

mod authoring;
mod codec;
mod facade;
mod flat;
mod model;
mod package;

pub use facade::{Builder, Drawing};
pub use flat::{
    FlatDrawing, FlatDrawingCommit, FlatDrawingEdit, FlatDrawingPatch, FlatPage, FlatShape,
    TextChange,
};
pub use model::{layer, page};
