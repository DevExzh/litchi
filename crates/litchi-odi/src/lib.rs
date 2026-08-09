//! `OpenDocument` Image support with semantic responsibility layers.
#![forbid(unsafe_code)]
#![allow(
    clippy::missing_errors_doc,
    clippy::module_name_repetitions,
    reason = "the image edit trait and facade share one documented typed-error contract and ODF vocabulary"
)]
#![allow(
    clippy::arbitrary_source_item_ordering,
    clippy::map_err_ignore,
    clippy::shadow_reuse,
    reason = "image codecs follow ODF traversal order and normalize dependency errors at the facade boundary"
)]

mod authoring;
mod codec;
mod editing;
mod facade;
mod flat;
mod history;
mod model;
mod package;
mod semantic;

pub use editing::{FrameEditor, MetadataEditor};
pub use facade::{
    Builder, Commit, Edit, Image, MetadataChange, MetadataFields, Patch, ResourceChange,
    StyleChange,
};
pub use flat::{FlatImage, FlatImageCommit, FlatImagePatch, FlatImageTransaction, FrameChange};
pub use history::{CommittedTransition, History, HistoryArtifact};
pub use model::{frame, map, resource, source};
pub use semantic::{
    ArtifactKind, Conflict, ConflictKind, FrameProperty, MetadataProperty, OperationKey,
    ResourceValue, SecurityPolicy, SemanticOperation, SemanticPatch, SemanticPlan, SemanticValue,
};
