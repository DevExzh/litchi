//! Layered BIFF8 worksheet comments/notes owner.
//!
//! The facade exposes contextual comment values while keeping BIFF8 record
//! parsing, continuation state, and regression fixtures in dedicated layers.

mod codec;
mod model;
#[cfg(test)]
mod tests;
mod transaction;

/// The comments family error: the workbook error of the source or
/// republished package.
pub use crate::Error;
pub(crate) use codec::CommentCollector;
pub use model::{CONTINUE_TYPE, MSODRAWING_TYPE, OBJ_TYPE, RECORD_TYPE, TXO_TYPE};
pub use model::{
    Comment, CommentRecord, HorizontalAlignment, NoteMetadata, ObjectIdentity, ObjectPadding,
    ObjectProperties, ObjectSubrecord, RecordKind, TextOrientation, TextProperties, TextRun,
    VerticalAlignment, Visibility,
};
pub(crate) use transaction::source_diagnostics;
pub use transaction::{
    Commit, Diagnostics, Edit, LimitKind, Operation, Patch, Snapshot, Update, Value, Worksheet,
};
