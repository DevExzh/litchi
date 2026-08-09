//! Immutable semantic values for this document family.

mod catalog;
mod component;
pub mod connection;
pub mod query;
mod table;

pub use catalog::{Catalog, Limits, OwnedCatalog};
pub use component::{Component, ComponentKind};
pub use table::{
    Column, DataType, Index, IndexColumn, Key, KeyColumn, KeyKind, ReferentialAction, Table,
    TableKind,
};
