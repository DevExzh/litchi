//! Word 2003 range-level protection metadata.
//!
//! The `protection` context models the bookmark-delimited editable ranges
//! described by [MS-DOC]. It is deliberately inert: usernames are never
//! authenticated, protection is never bypassed, and no document content is
//! changed while decoding the tables.

mod codec;
mod model;

#[cfg(test)]
mod tests;

pub use model::{Mode, Range, Ranges, Reserved, Role, Selector, User};
