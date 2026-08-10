//! Typed normal-view and notes-text-view state for `PowerPoint` presentations.
//!
//! The owner is split by responsibility: `model` contains the semantic
//! records, while `codec` owns the MS-PPT binary record boundaries and
//! validation. Unknown atom payloads remain available through the model's
//! opaque variant so a read/write cycle does not discard data.

mod codec;
mod model;
#[cfg(test)]
mod tests;

#[allow(
    clippy::module_name_repetitions,
    reason = "`NormalViewSetInfo` is the established public type name matching the MS-PPT record name; renaming it would break downstream crates"
)]
pub use model::{
    NormalViewSet, NormalViewSetInfo, NormalViewSetPayload, NotesTextViewInfo, ViewBarState,
};
