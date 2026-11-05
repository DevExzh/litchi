//! `OpenDocument` Image support with semantic responsibility layers.
#![forbid(unsafe_code)]
#![allow(
    clippy::missing_errors_doc,
    clippy::module_name_repetitions,
    reason = "the image edit trait and facade share one documented typed-error contract and ODF vocabulary"
)]
#![allow(
    clippy::arbitrary_source_item_ordering,
    clippy::shadow_reuse,
    reason = "image codecs follow ODF traversal order and keep source/target names aligned with ODF traversal"
)]

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
