//! Contextual package-bound facade for the main WordprocessingML document.
//!
//! The owner is split by responsibility while keeping every method on the
//! public [`super::Document`] model: semantic queries live in `model`, OPC
//! relationship traversal lives in `package`, XML and typed-part decoding
//! lives in `codec`, and graph/state checks live in `validation`.

mod codec;
mod model;
mod package;
mod validation;

#[cfg(test)]
mod tests;
