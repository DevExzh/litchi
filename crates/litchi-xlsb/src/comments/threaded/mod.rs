//! Typed, inert XLSB threaded comments and persons.
//!
//! Legacy BIFF12 notes remain owned by [`super`].  This owner follows the
//! separate XML parts from `[MS-XLSB]` 2.1.17--2.1.18 and keeps semantic values,
//! XML conversion, package CRUD, and graph validation in distinct layers.

pub mod codec;
pub mod edit;
pub mod package;
pub mod semantic;
pub mod validation;

#[cfg(test)]
mod tests;

pub use codec::{parse_comments, parse_persons, write_comments, write_persons};
pub use edit::{Commit, Patch, Snapshot, SourcePart, SourceRelationship, Transaction, apply, read};
pub use package::{
    load_from_worksheet, load_graph, load_people, remove_from_worksheet, remove_graph,
    remove_people, store_graph, store_on_worksheet, store_people,
    validate_graph as validate_package_graph,
};
pub use semantic::{
    Comment, Comments, CommentsPart, Graph, Mention, People, PeoplePart, Person, RawAttribute,
    RawXml, Thread,
};
pub use validation::{
    Error, Result, group_threads, validate_comments, validate_graph, validate_people,
};

/// Maximum serialized bytes accepted for one threaded-comments or persons
/// part.  The bound applies before XML parsing and before package mutation.
pub const MAX_PART_BYTES: usize = 64 * 1024 * 1024;
/// Maximum persons in one workbook persons part.
pub const MAX_PERSONS: usize = 100_000;
/// Maximum comments in one worksheet threaded-comments part.
pub const MAX_COMMENTS: usize = 100_000;
/// Maximum mentions across one worksheet threaded-comments part.
pub const MAX_MENTIONS: usize = 100_000;
/// Maximum UTF-16 code units in one comment's text.
pub const MAX_TEXT_UTF16: usize = 1_048_576;
/// Maximum byte length of one person identity string or preserved attribute.
pub const MAX_IDENTITY_BYTES: usize = 16 * 1024;
/// Maximum preserved unknown attributes on one semantic element.
pub const MAX_ATTRIBUTES: usize = 256;
/// Maximum preserved unknown child elements on one semantic element.
pub const MAX_EXTENSIONS: usize = 256;
/// Maximum total bytes of preserved unknown children on one semantic element.
pub const MAX_EXTENSION_BYTES: usize = 8 * 1024 * 1024;
/// Maximum nesting depth accepted while preserving unknown XML children.
pub const MAX_XML_DEPTH: usize = 128;
