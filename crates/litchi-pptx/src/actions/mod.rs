//! Layered, inert PresentationML action-setting discovery.
//!
//! The model contains the semantic action vocabulary, the codec performs a
//! bounded XML scan, and the package layer resolves declared OPC targets
//! without opening, following, or executing them.

mod codec;
mod model;
mod package;

#[cfg(test)]
mod tests;

pub use model::{Jump, Kind, Setting, Target, Trigger};
pub use package::{Limits, load_slide_action_settings};

// Compatibility aliases retain the historical facade while the canonical
// names remain concise inside the actions context.
pub type PptxActionTrigger = Trigger;
pub type PptxSlideShowJump = Jump;
pub type PptxActionKind = Kind;
pub type PptxActionTarget = Target;
pub type PptxActionSetting = Setting;

pub(super) fn invalid(message: impl Into<String>) -> crate::Error {
    crate::Error::Invalid(message.into())
}

pub(super) fn limit(resource: &'static str, limit: usize) -> crate::Error {
    crate::Error::Limit { resource, limit }
}
