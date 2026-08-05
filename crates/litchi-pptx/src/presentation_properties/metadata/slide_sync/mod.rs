//! Contextual slide-library synchronization metadata.
//!
//! The model contains validated values, the codec owns bounded XML, and the
//! package layer owns only the slide-to-data-part relationship graph.

mod codec;
mod model;
mod package;

pub use model::{DateTime, Offset, Part, Properties};
pub use package::{CONTENT_TYPE, RELATIONSHIP_TYPE, load, store};
