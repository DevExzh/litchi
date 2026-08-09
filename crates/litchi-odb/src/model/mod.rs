//! Immutable semantic values for this document family.

mod catalog;
mod component;
pub mod connection;
mod extension;
pub mod query;
mod table;

pub use catalog::{Catalog, Limits, OwnedCatalog};
pub use component::{Component, ComponentKind};
pub use extension::ProducerExtension;
pub use table::{
    Column, DataType, Index, IndexColumn, Key, KeyColumn, KeyKind, ReferentialAction, Relation,
    Table, TableKind,
};
