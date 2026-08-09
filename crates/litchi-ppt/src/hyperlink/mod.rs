//! Hyperlink definitions and `PowerPoint` 9 hyperlink extensions.

mod codec;
mod model;
#[cfg(test)]
mod tests;

#[allow(
    clippy::module_name_repetitions,
    reason = "`HyperlinkExtension` is the established public API name for `PowerPoint` 9 hyperlink metadata; renaming it would break downstream crates"
)]
pub use model::{
    Hyperlink, HyperlinkExtension, Hyperlinks, Interaction, InteractionAction, InteractionJump,
    InteractionLimits, InteractionLinkTarget, InteractionTrigger, InteractiveInfoAtom,
    MacroNameAtom, ShapeInteractionEntry,
};

pub(crate) use codec::encode_record;
