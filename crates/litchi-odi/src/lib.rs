//! `OpenDocument` Image support with semantic responsibility layers.
#![forbid(unsafe_code)]

mod authoring;
mod codec;
mod editing;
mod facade;
pub mod flat;
mod history;
pub mod metadata;
mod model;
mod package;
mod semantic;

pub use editing::{FrameEditor, MetadataEditor};
pub use facade::{
    Builder, Commit, Diagnostics, Edit, Image, MetadataChange, MetadataFields, Package, Patch,
    ResourceChange, StyleChange,
};
pub use flat::FrameChange;
pub use history::{CommittedTransition, History, HistoryArtifact};
pub use litchi_odf_common::{
    Error, LimitKind, OpenOptions, SemanticLimitKind, SemanticLimits, SemanticLimitsBuilder,
    SourceLimitKind, SourceLimits, SourceLimitsBuilder,
};
pub use model::{active, frame, map, resource, security, source, surface};
pub use security::{ProtectionDisposition, ProtectionInventory};
pub use semantic::{
    ArtifactKind, CapabilityState, Conflict, ConflictKind, FrameProperty, MetadataProperty,
    OperationKey, PublicationState, ResourceValue, RewriteBlocker, RewriteCapability,
    SecurityCapabilities, SecurityPolicy, SemanticOperation, SemanticPatch, SemanticPlan,
    SemanticValue, StyleDependencyState,
};
pub use surface::{Surface, SurfaceDisposition, SurfaceKind, SurfaceLocation};
