//! Hyperlink definitions and PowerPoint 9 hyperlink extensions.

mod codec;
mod model;
#[cfg(test)]
mod tests;

pub use model::{
    Hyperlink, HyperlinkExtension, Hyperlinks, Interaction, InteractionAction, InteractionJump,
    InteractionLimits, InteractionLinkTarget, InteractionTrigger, InteractiveInfoAtom,
    MacroNameAtom, ShapeInteractionEntry,
};

pub(crate) use codec::encode_record;
