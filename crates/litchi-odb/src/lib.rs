//! `OpenDocument` Database support with semantic responsibility layers.
#![forbid(unsafe_code)]
#![allow(
    clippy::unnecessary_sort_by,
    reason = "database XML emitters are bounded, ordered by ODF traversal, and retain short-lived semantic projection names"
)]

mod authoring;
mod codec;
mod facade;
pub mod metadata;
mod model;
mod package;

pub use authoring::{
    ActiveContentDisposition, Change, ChangeAction, ChangeKind, Commit, CompositionLimits,
    DependencyDisposition, Diagnostics, DurablePatch, Edit, EditPolicy, EncryptionCapability,
    EncryptionPolicy, History, HistoryLimits, JoinError, JoinedEdits, Lineage, MergeChoice,
    MergePlan, MergePlanError, Patch, ProtectionCapabilities, ProtectionDecision,
    ProtectionOperation, ProtectionRefusal, ProtectionStatus, ProtectionSupport,
    ProtectionTransition, QueryChange, SealedPatch, SignatureCapability, SignaturePolicy, SubEdit,
};
pub use facade::{Builder, Database, Package};
pub use litchi_odf_common::signature::{
    DigitalSignature, DigitalSignatures, SignatureValidity, SignatureVerification,
};
pub use litchi_odf_common::{
    Error, LimitKind, OpenOptions, PreparedPackage, SemanticLimitKind, SemanticLimits,
    SourceLimitKind, SourceLimits,
};
pub use model::connection::Connection;
pub use model::query::{
    Query, QueryCommandInventory, QueryCommandSemantics, QueryDependencyReason,
    QueryDependencySupport, QueryJoin, QueryJoinKind, QueryParameter, QueryParameterKind,
    QueryRelation, QueryRelationRole, QueryStatementKind, QueryUpdateTarget,
};
pub use model::{
    ActiveContentEntry, ActiveContentInventory, ActiveContentKind, ApplicationConnectionSettings,
    AutoIncrementSettings, BooleanComparisonMode, Catalog, CharacterSetSettings, Column, Component,
    ComponentDependency, ComponentDependencyInventory, ComponentDependencyKind, ComponentKind,
    ComponentLinkKind, ComponentTransferRefusal, ComponentTransferSupport, DataSourceSetting,
    DataSourceSettingType, DataType, DatabaseSettings, DelimiterSettings, DriverSettings,
    FileDatabaseTarget, Index, IndexColumn, Key, KeyColumn, KeyKind, Limits, LoginSettings,
    OwnedCatalog, ProducerExtension, ReferentialAction, Relation, RelationResolution,
    ServerDatabaseAddress, ServerDatabaseTarget, Settings, Table, TableFilter, TableKind,
    TableSetting, TableTypeFilter, connection, query,
};
