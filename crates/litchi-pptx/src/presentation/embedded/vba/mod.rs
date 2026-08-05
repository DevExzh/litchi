//! Inert VBA-project relationship metadata.
//!
//! The project payload is deliberately opaque in the PPTX owner crate. The
//! shared graph service owns relationship/content-type mutation; callers that
//! need MS-OVBA decoding can pass the borrowed payload to `litchi-vba`.

mod model;
mod package;

pub use model::Project;
pub use package::{discover, remove, store};

pub(crate) const MAX_PAYLOAD_BYTES: usize = 64 * 1024 * 1024;
