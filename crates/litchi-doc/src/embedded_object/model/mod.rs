//! Semantic values and transactional state for DOC embedded objects.

mod editor;
mod info;
mod options;
mod reference;

pub use editor::Editor;
pub(in crate::embedded_object) use editor::{FieldMarker, RawPiece};
pub use info::Info;
pub use options::WriteOptions;
pub use reference::Reference;
