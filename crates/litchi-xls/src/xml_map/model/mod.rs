//! Typed XML-map values, split by identity, opaque payload, and owner.

mod binding;
mod identity;
mod info;
mod map;
mod opaque;
mod schema;

pub use binding::{DataBinding, LoadMode};
pub use identity::{MapId, SchemaId, XPath};
pub use info::MapInfo;
pub use map::Map;
pub use opaque::OpaqueXml;
pub use schema::{NamespaceDeclaration, Schema};
