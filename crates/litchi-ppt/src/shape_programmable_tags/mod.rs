//! Shape-scoped programmable tags from OfficeArt `ClientData` records.
//!
//! The owner is split by responsibility: [`model`] contains the typed,
//! semantic tag values, while [`codec`] owns the MS-PPT record boundaries,
//! validation, and lossless serialization.

mod codec;
mod model;
#[cfg(test)]
mod tests;

pub use model::{
    PresentationShapeProgrammableTagsEntry, ShapeBinaryTag, ShapeBinaryTagPayload,
    ShapeBinaryTagVersion, ShapeProgrammableTag, ShapeProgrammableTagLimits, ShapeProgrammableTags,
    ShapeProgrammableTagsEntry, ShapeStringTag, ShapeStyleAtom,
};
