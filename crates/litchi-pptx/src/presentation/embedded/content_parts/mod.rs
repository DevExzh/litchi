//! Inert, lossless PresentationML `p:contentPart` owners.
//!
//! The module is deliberately layered: [`model`] exposes contextual semantic
//! values, [`codec`] validates bounded slide XML and retains anchor bytes, and
//! [`package`] resolves only the owning slide relationship graph. Referenced
//! payload vocabularies remain opaque and are never executed.

mod codec;
mod model;
mod package;

pub use litchi_opc::TargetMode;
pub use model::{Anchor, ContentPart, Payload, Relationship, RelationshipMetadata, Target};
pub use package::{Limits, load_slide};

#[cfg(test)]
mod tests;
