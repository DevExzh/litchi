//! Typed, package-neutral XLSX threaded comments.
//!
//! The facade keeps the data graph and bounded XML codec together while OPC
//! relationships and physical part lifecycle remain in the host adapter.

mod codec;
mod model;

#[cfg(test)]
mod tests;

pub use codec::{
    parse_comments, parse_persons, validate_comments, validate_graph, validate_guid,
    validate_people, validate_timestamp, write_comments, write_persons,
};
pub use model::{Comment, Comments, Graph, Mention, People, Person, SheetPart, WorkbookPart};

/// Maximum serialized size accepted for one threaded-comments XML part.
pub const MAX_PART_BYTES: usize = 64 * 1024 * 1024;
/// Maximum number of people accepted in one people part.
pub const MAX_PERSONS: usize = 100_000;
/// Maximum number of comments accepted in one comments part.
pub const MAX_COMMENTS: usize = 100_000;
/// Maximum number of mentions accepted in one comments part.
pub const MAX_MENTIONS: usize = 100_000;
/// Maximum UTF-16 code units accepted in one comment's text.
pub const MAX_TEXT_UTF16: usize = 1_048_576;
/// Maximum byte length accepted for person identity metadata.
pub const MAX_IDENTITY_BYTES: usize = 16_384;
