//! Inert OLE object inventories and safe binary embedding.

mod authoring;
mod codec;
mod model;
mod package;

/// Contextual, source-checked edits owned by one PresentationML slide.
pub mod slide;

pub use authoring::add;
pub use model::{Authored, Frame, Kind, Mode, Object, Target};
pub use package::{Limits, load_slide};
