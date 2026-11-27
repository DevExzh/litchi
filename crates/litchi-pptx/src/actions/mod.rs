//! Layered, inert `PresentationML` action-setting discovery.
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

pub(super) fn invalid(message: impl Into<String>) -> crate::Error {
    crate::Error::Invalid(message.into())
}

pub(super) fn limit(
    kind: crate::error::CodecLimitKind,
    observed: usize,
    maximum: usize,
) -> crate::Error {
    crate::Error::Limit {
        kind: crate::error::LimitKind::Codec(kind),
        observed,
        maximum,
    }
}
