//! Inert OLE object inventories and safe binary embedding.

mod authoring;
mod codec;
mod model;
mod package;

pub use authoring::add;
pub use model::{Authored, Frame, Kind, Mode, Object, Target};
pub use package::{Limits, load_slide};
