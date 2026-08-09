//! `OpenDocument` Database support with semantic responsibility layers.
#![forbid(unsafe_code)]

mod authoring;
mod codec;
mod facade;
mod model;
mod package;

pub use authoring::{Commit, Edit, Patch, QueryChange};
pub use facade::{Builder, Database};
pub use model::{
    Catalog, Column, Component, ComponentKind, DataType, Index, IndexColumn, Key, KeyColumn,
    KeyKind, Limits, OwnedCatalog, ReferentialAction, Table, TableKind, connection, query,
};
