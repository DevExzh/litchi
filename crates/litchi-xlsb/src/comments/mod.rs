//! Typed, bounded XLSB comments.
//!
//! The comments stream follows the `COMMENTS`, `COMMENTAUTHORS`,
//! `COMMENTLIST`, and `COMMENT` rules in [MS-XLSB] section 2.1.7.8. The
//! semantic model is kept separate from the BIFF12 record codec so callers
//! can use concise owner types without importing host-specific rich-string
//! or package abstractions.

mod codec;
mod model;

pub use codec::{
    Error, MAX_AUTHOR_UNITS, MAX_COLUMNS, MAX_ROWS, MAX_TEXT_RUNS, MAX_TEXT_UNITS, Result,
};
pub use codec::{read, write};
pub use model::{Comment, CommentRun};
