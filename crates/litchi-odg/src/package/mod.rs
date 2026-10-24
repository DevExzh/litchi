//! Validated package ownership for this family.

mod snapshot;

pub use litchi_core::patch::HistoryLimits;
pub use snapshot::{
    ActiveContentStatus, ActiveContentWritePolicy, Change, Commit, ControlReferenceChange,
    Diagnostics, DurablePatch, Edit, GeometryChange, History, JoinError, JoinedEdits, LayerChange,
    Lineage, MergePlan, NameChange, Package, PageNameChange, PageStyleChange, PageTransitionChange,
    Patch, PathChange, ResourceChange, SecurityCapabilities, SecurityLifecycleDisposition,
    SecurityLifecycleOperation, SecurityLifecycleRefusal, SecurityStatus, SecurityWritePolicy,
    ShapeTransfer, StructureChange, StyleChange, SubEdit, TextChange, TransferControl,
    TransferResource, TransferStyle, TransferStyleResource,
};
pub(crate) use snapshot::{MIMETYPE, TEMPLATE_MIMETYPE};
