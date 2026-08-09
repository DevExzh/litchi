//! Immutable semantic values for this document family.

mod catalog;
pub mod connection;
pub mod query;
mod table;

pub use catalog::{Catalog, Limits, OwnedCatalog};
pub use table::{Column, Table, TableKind};
