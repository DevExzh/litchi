//! Semantic slide-master and slide-layout authoring facade.

mod codec;
mod model;
mod package;
pub mod placeholder;

#[cfg(test)]
mod tests;

pub use crate::shape::{PLACEHOLDER_TYPE_EXTENSION_URI, PlaceholderTypeExtension};
pub use model::{
    AuthoredSlideLayout, AuthoredSlideMaster, MIN_MASTER_OR_LAYOUT_ID, PlaceholderKind,
    PlaceholderSpec, SlideLayoutKind,
};
pub use package::{
    add_slide_layout, add_slide_master, remove_slide_layout, store_placeholder_shape,
    validate_master_layout_graph,
};

pub(crate) use package::preflight_add_slide_layout;
pub use placeholder::{
    Commit as PlaceholderTypeExtensionCommit, Edit as PlaceholderTypeExtensionEdit,
    Limits as PlaceholderTypeExtensionLimits, Patch as PlaceholderTypeExtensionPatch,
    SignaturePolicy as PlaceholderSignaturePolicy,
    SlotCommit as PlaceholderTypeExtensionSlotCommit, SlotEdit as PlaceholderTypeExtensionSlotEdit,
    SlotPatch as PlaceholderTypeExtensionSlotPatch,
    SlotSnapshot as PlaceholderTypeExtensionSlotSnapshot,
    Snapshot as PlaceholderTypeExtensionSnapshot,
    apply_commit as apply_placeholder_type_extension_commit,
    apply_patch as apply_placeholder_type_extension_patch,
    apply_patch_with_policy as apply_placeholder_type_extension_patch_with_policy,
    apply_slot_commit as apply_placeholder_type_extension_slot_commit,
    apply_slot_patch as apply_placeholder_type_extension_slot_patch,
    apply_slot_patch_with_policy as apply_placeholder_type_extension_slot_patch_with_policy,
    load_slot_snapshot as load_placeholder_type_extension_slot_snapshot,
    load_slot_snapshot_with_limits as load_placeholder_type_extension_slot_snapshot_with_limits,
    load_snapshot as load_placeholder_type_extension_snapshot,
    load_snapshot_with_limits as load_placeholder_type_extension_snapshot_with_limits,
};

pub(crate) use placeholder::preflight_store_placeholder_shape;
