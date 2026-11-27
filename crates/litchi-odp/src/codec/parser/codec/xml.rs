//! Layered XML parser facade.
//!
//! The implementation is deliberately split by responsibility: event-stream
//! traversal lives in the codec family, model assembly in the semantic family,
//! and bounded namespace/attribute checks in the validation family. The Parser
//! methods remain exposed through the same crate-local facade.

mod codec;
#[cfg(test)]
mod oracle;
mod semantic;
mod validation;
