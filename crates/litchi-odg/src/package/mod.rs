//! Validated package ownership for this family.

mod snapshot;

pub use litchi_core::patch::HistoryLimits;
pub use snapshot::{
    ActiveContentStatus, ActiveContentWritePolicy, Change, Commit, ControlReferenceChange,
    DurablePatch, GeometryChange, JoinedEdits, LayerChange, Lineage, MergePlan, NameChange,
    PageNameChange, PageStyleChange, Patch, PathChange, PreparedEdit, ResourceChange,
    SecurityCapabilities, SecurityStatus, SecurityWritePolicy, ShapeTransfer, Snapshot,
    SnapshotHistory as History, StructureChange, StyleChange, TextChange, Transaction,
    TransferControl, TransferResource, TransferStyle, TransferStyleResource,
};
pub(crate) use snapshot::{MIMETYPE, TEMPLATE_MIMETYPE};
