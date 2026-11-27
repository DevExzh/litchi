//! Semantic model for an inert ODF alphabetical-index auto-mark reference.

use crate::variable_declaration::{Part, Scope};

/// One inert `text:alphabetical-index-auto-mark-file` reference.
///
/// The referenced concordance file remains external: it is never opened,
/// fetched, or parsed.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct AlphabeticalIndexAutoMarkFile {
    /// The package part (content or styles) declaring the reference.
    pub part: Part,
    /// The `office:text` scope declaring the reference.
    pub scope: Scope,
    /// The external concordance file URI from `xlink:href`.
    pub href: String,
}
