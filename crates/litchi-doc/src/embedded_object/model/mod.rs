//! Semantic values and transactional state for DOC embedded objects.

pub(in crate::embedded_object) mod comp_obj;
mod editor;
mod info;
mod inventory;
mod metadata;
pub(in crate::embedded_object) mod ole;
mod options;
mod reference;
mod unknown;

pub use comp_obj::{Clipboard, CompObj};
pub use editor::Editor;
pub(in crate::embedded_object) use editor::{FieldMarker, RawPiece};
pub use info::Info;
pub use inventory::{Entry, Inventory};
pub use metadata::Metadata;
pub use ole::{Kind, Ole};
pub use options::WriteOptions;
pub use reference::Reference;
pub use unknown::Unknown;
