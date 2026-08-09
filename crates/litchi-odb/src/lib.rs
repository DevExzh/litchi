//! `OpenDocument` Database support with semantic responsibility layers.
#![forbid(unsafe_code)]
#![allow(
    clippy::needless_pass_by_value,
    reason = "database edit operations own specification-shaped values so patches remain self-contained"
)]
#![allow(
    clippy::arbitrary_source_item_ordering,
    clippy::format_push_string,
    clippy::shadow_reuse,
    clippy::shadow_unrelated,
    clippy::unnecessary_sort_by,
    reason = "database XML emitters are bounded, ordered by ODF traversal, and retain short-lived semantic projection names"
)]

mod authoring;
mod codec;
mod facade;
mod model;
mod package;

pub use authoring::{Change, ChangeKind, Commit, Edit, History, HistoryLimits, Patch, QueryChange};
pub use facade::{Builder, Database};
pub use model::connection::Connection;
pub use model::query::Query;
pub use model::{
    Catalog, Column, Component, ComponentKind, DataType, Index, IndexColumn, Key, KeyColumn,
    KeyKind, Limits, OwnedCatalog, ProducerExtension, ReferentialAction, Relation, Table,
    TableKind, connection, query,
};
