//! The alternative-format import family (record 0355a): body
//! `w:altChunk` anchors of an opened or built package, their main-part
//! relationships and internal payload parts.
//!
//! [`Package::edit_alt_chunks`](crate::Package::edit_alt_chunks) stages an
//! [`Edit`] whose commit publishes a new package; the anchor element is
//! spliced into the main-document XML, which is otherwise kept byte for
//! byte. [`Package::apply_alt_chunks`](crate::Package::apply_alt_chunks)
//! replays its [`Patch`].
//!
//! ```
//! use litchi_docx::Package;
//! use litchi_docx::alt::{Data, Import};
//!
//! let package = Package::new()?;
//! let mut edit = package.edit_alt_chunks();
//! edit.add(Import::data(Data::Html(b"<p>opaque</p>".to_vec())), None)?;
//! let commit = edit.commit()?;
//! assert_eq!(commit.package().document()?.alts()?.len(), 1);
//! assert!(package.document()?.alts()?.is_empty());
//! # Ok::<(), litchi_docx::Error>(())
//! ```

use super::{Chunk, Import};
use crate::package::family::{graph_edit, graph_family};

graph_family!("alternative-format", own_limits);
graph_edit!("alternative-format", edit_alt_chunks, apply_alt_chunks);

/// The budgets of the alternative-format family, reported by
/// [`crate::Error::AltChunkLimit`].
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum LimitKind {
    /// Anchors in one document ([`super::MAX_CHUNKS`]).
    Anchors,
    /// Generated `rIdAltChunkN` relationship identifiers.
    RelationshipIds,
    /// Generated `afchunkN` payload part names.
    PartNames,
}

impl std::fmt::Display for LimitKind {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        formatter.write_str(match self {
            Self::Anchors => "anchor",
            Self::RelationshipIds => "relationship identifier",
            Self::PartNames => "part name",
        })
    }
}

impl Edit {
    /// The active body anchors, in document order, staged changes included.
    ///
    /// # Errors
    ///
    /// Returns an XML error for a malformed anchor.
    pub fn chunks(&self) -> Result<Vec<Chunk>> {
        self.package().alt_chunks()
    }

    /// Append an anchor for `import` after the body content and return it.
    ///
    /// # Errors
    ///
    /// As [`Self::insert`].
    pub fn add(&mut self, import: Import, match_source: Option<bool>) -> Result<Chunk> {
        let index = self.chunks()?.len();
        self.insert(index, import, match_source)
    }

    /// Insert an anchor for `import` before anchor `index` (after the body
    /// content for `index` equal to the anchor count) and return it.
    ///
    /// # Errors
    ///
    /// Returns [`crate::Error::OutOfBounds`] for an index past the anchor
    /// count, [`crate::Error::AltChunkLimit`] for an exhausted budget, or an
    /// invalid-import or package error.
    pub fn insert(
        &mut self,
        index: usize,
        import: Import,
        match_source: Option<bool>,
    ) -> Result<Chunk> {
        self.stage(|package| {
            package.with_alt_body(|package, body| {
                package.insert_alt(body, index, import, match_source)
            })
        })
    }

    /// Replace anchor `index` with one for `import`; returns the replaced
    /// anchor. Its relationship and payload go as [`Self::remove`] removes
    /// them.
    ///
    /// # Errors
    ///
    /// Returns [`crate::Error::OutOfBounds`] for a missing anchor, or the
    /// errors of [`Self::insert`].
    pub fn replace(
        &mut self,
        index: usize,
        import: Import,
        match_source: Option<bool>,
    ) -> Result<Chunk> {
        self.stage(|package| {
            package.with_alt_body(|package, body| {
                package.replace_alt(body, index, import, match_source)
            })
        })
    }

    /// Remove anchor `index` and return it. Its relationship goes when no
    /// other anchor uses it, and an internal payload part when nothing else
    /// in the package relates to it.
    ///
    /// # Errors
    ///
    /// Returns [`crate::Error::OutOfBounds`] for a missing anchor, or an
    /// error for a relationship that is not an alternative-format import.
    pub fn remove(&mut self, index: usize) -> Result<Chunk> {
        self.stage(|package| package.with_alt_body(|package, body| package.remove_alt(body, index)))
    }

    /// Move anchor `from` to anchor position `to`, keeping its
    /// relationship.
    ///
    /// # Errors
    ///
    /// Returns [`crate::Error::OutOfBounds`] for a missing anchor position.
    pub fn move_to(&mut self, from: usize, to: usize) -> Result<&mut Self> {
        self.stage(|package| package.with_alt_body(|_, body| body.move_alt(from, to)))?;
        Ok(self)
    }
}
