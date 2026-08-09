//! `OpenDocument` Drawing support with semantic responsibility layers.
#![forbid(unsafe_code)]
#![allow(
    clippy::arbitrary_source_item_ordering,
    reason = "drawing package helpers stay in manifest and content traversal order"
)]

mod authoring;
mod codec;
mod facade;
mod flat;
mod model;
mod package;

pub use facade::{Builder, Drawing};
pub use facade::{
    Change as PackageChange, Commit as PackageCommit,
    ControlReferenceChange as PackageControlReferenceChange, DurablePatch as PackageDurablePatch,
    GeometryChange as PackageGeometryChange, History as PackageHistory,
    HistoryLimits as PackageHistoryLimits, JoinedEdits as PackageJoinedEdits,
    LayerChange as PackageLayerChange, Lineage as PackageLineage, MergePlan as PackageMergePlan,
    NameChange as PackageNameChange, PageNameChange as PackagePageNameChange,
    PageStyleChange as PackagePageStyleChange, Patch as PackagePatch,
    PathChange as PackagePathChange, PreparedEdit as PackagePreparedEdit,
    ResourceChange as PackageResourceChange, SecurityStatus as PackageSecurityStatus,
    ShapeTransfer as PackageShapeTransfer, Snapshot as PackageSnapshot,
    StructureChange as PackageStructureChange, StyleChange as PackageStyleChange,
    TextChange as PackageTextChange, Transaction as PackageTransaction,
    TransferResource as PackageTransferResource,
};
pub use flat::{
    FlatDrawing, FlatDrawingCommit, FlatDrawingEdit, FlatDrawingPatch, FlatPage, FlatShape,
    TextChange,
};
pub use model::{FormControl, form, layer, page, resource, shape};
