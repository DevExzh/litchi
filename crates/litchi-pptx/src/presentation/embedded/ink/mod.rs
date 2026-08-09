//! Inert `InkML` content-part discovery and storage.

mod codec;
mod model;
mod package;

pub use model::{Annotation, StoredAnnotation};
pub use package::{Limits, load_slide, store_slide};

/// OPC content type of an `InkML` part.
pub const CONTENT_TYPE: &str = "application/inkml+xml";
