//! Immutable semantic values for this document family.

mod active;
mod catalog;
mod component;
pub mod connection;
mod extension;
#[path = "query.rs"]
pub mod stored_query;
pub use stored_query as query;
mod table;

pub use active::{ActiveContentEntry, ActiveContentInventory, ActiveContentKind};
pub use catalog::{Catalog, Limits, OwnedCatalog};
pub use component::{
    Component, ComponentDependency, ComponentDependencyInventory, ComponentDependencyKind,
    ComponentKind, ComponentLinkKind, ComponentTransferRefusal, ComponentTransferSupport,
};
pub use extension::ProducerExtension;
pub use table::{
    Column, DataType, Index, IndexColumn, Key, KeyColumn, KeyKind, ReferentialAction, Relation,
    RelationResolution, Table, TableKind,
};
