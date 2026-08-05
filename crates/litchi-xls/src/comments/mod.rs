//! Layered BIFF8 worksheet comments/notes owner.
//!
//! The facade exposes contextual comment values while keeping BIFF8 record
//! parsing, continuation state, and regression fixtures in dedicated layers.

mod codec;
mod model;
#[cfg(test)]
mod tests;

pub(crate) use codec::CommentCollector;
pub use model::{CONTINUE_TYPE, MSODRAWING_TYPE, OBJ_TYPE, RECORD_TYPE, TXO_TYPE};
pub use model::{
    Comment, CommentRecord, HorizontalAlignment, NoteMetadata, ObjectIdentity, ObjectPadding,
    ObjectProperties, ObjectSubrecord, RecordKind, TextOrientation, TextProperties, TextRun,
    VerticalAlignment, Visibility,
};
