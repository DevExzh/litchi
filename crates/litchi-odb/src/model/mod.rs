//! Immutable semantic values for this document family.

mod active;
mod catalog;
mod component;
pub mod connection;
mod extension;
#[path = "query.rs"]
pub mod stored_query;
pub use stored_query as query;
mod settings;
mod table;

pub use active::{ActiveContentEntry, ActiveContentInventory, ActiveContentKind};
pub use catalog::{Catalog, Limits, OwnedCatalog};
pub use component::{
    Component, ComponentDependency, ComponentDependencyInventory, ComponentDependencyKind,
    ComponentKind, ComponentLinkKind, ComponentTransferRefusal, ComponentTransferSupport,
};
pub use connection::{FileDatabaseTarget, ServerDatabaseAddress, ServerDatabaseTarget};
pub use extension::ProducerExtension;
pub use settings::{
    ApplicationConnectionSettings, AutoIncrementSettings, BooleanComparisonMode,
    CharacterSetSettings, DataSourceSetting, DataSourceSettingType, DatabaseSettings,
    DelimiterSettings, DriverSettings, LoginSettings, Settings, TableFilter, TableSetting,
    TableTypeFilter,
};
#[cfg(test)]
pub(crate) use settings::{parse_count, parse_test_lock, reset_parse_count};
pub use table::{
    Column, DataType, Index, IndexColumn, Key, KeyColumn, KeyKind, ReferentialAction, Relation,
    RelationResolution, Table, TableKind,
};
