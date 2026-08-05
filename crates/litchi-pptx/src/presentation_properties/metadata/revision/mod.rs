//! Revision Information values, bounded XML, and OPC lifecycle.

mod codec;
mod model;
mod package;

pub use codec::{CONTENT_TYPE, RELATIONSHIP_TYPE};
pub use model::{Client, Info, Namespace, Part};
pub use package::{load, store};
