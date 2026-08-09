//! `OpenDocument` Drawing support with semantic responsibility layers.
#![forbid(unsafe_code)]

mod authoring;
mod codec;
mod facade;
mod flat;
mod model;
mod package;

pub use facade::{Builder, Drawing};
pub use facade::{
    Commit as PackageCommit, LayerChange as PackageLayerChange, NameChange as PackageNameChange,
    Patch as PackagePatch, Snapshot as PackageSnapshot, TextChange as PackageTextChange,
    Transaction as PackageTransaction,
};
pub use flat::{
    FlatDrawing, FlatDrawingCommit, FlatDrawingEdit, FlatDrawingPatch, FlatPage, FlatShape,
    TextChange,
};
pub use model::{layer, page, shape};
