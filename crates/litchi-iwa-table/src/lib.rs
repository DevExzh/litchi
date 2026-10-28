//! Format-neutral table-family engines shared by the iWork table owners.
//!
//! Numbers sheets, Pages body tables and Keynote slide tables store the same
//! TST table graph: a `TableModelArchive`, its tiles of BNC cells, and the
//! `TableDataList` sidecars that hold formats, controls, styles and text.
//! The engines in this crate operate on one resolved table model of that
//! graph. Each leaf owner (`litchi-numbers`, `litchi-pages`,
//! `litchi-keynote`) resolves its own semantic selector to a [`cell_format::Target`],
//! supplies its package through the small [`TableSource`] trait, and keeps
//! publication (ZIP reassembly, candidate reopen, previews) and its public
//! transaction vocabulary.
//!
//! This crate is a private implementation layer of those owners (record
//! 0379, ADR 0024). It is never re-exported by a facade: its signatures name
//! native archive structures and resolved object positions, which the
//! supported APIs do not expose.

#![forbid(unsafe_code)]

use litchi_iwa_archive::{ComponentCatalog, SourceCatalog};
use litchi_iwa_core::RawMessage;

/// Cell data formats and cell controls of an existing table cell: the eleven
/// display families, the four scalar controls and the Pop-Up Menu graph
/// (moved from `litchi-numbers`, record 0379).
pub mod cell_format;

/// Cell fill and text layout through the copy-on-write cell-style graph
/// (record 0371).
pub mod cell_appearance;

/// Format-neutral whole-cell character-style graph planning.
pub mod cell_text_style;

pub mod hidden_axes;

/// Format-neutral table-cell comment and direct-reply transaction kernel.
pub mod comments;

/// Format-neutral native graph/rewrite engine for root table-cell comments.
///
/// The public facades only provide selector evidence and map the shared
/// transaction result into their package error vocabulary.  All copy-on-write
/// archive mutation lives here so the three owners cannot drift in their
/// ownership, list-key, or locality rules.
pub mod comments_native;

pub mod structure;

/// A located native object: its messages and its position in the component
/// catalog.
#[derive(Debug, Clone, Copy)]
pub struct ResolvedObject<'a> {
    /// The object's messages.
    pub messages: &'a [RawMessage],
    /// Index of the owning component in [`TableSource::catalog`].
    pub component_index: usize,
    /// Index of the object inside that component's archive.
    pub object_index: usize,
}

/// An object index lookup failed because the leaf's index is inconsistent.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct ResolveError;

/// The package view a table engine reads.
///
/// A leaf implements it for its parsed package: the decoded component
/// catalog, the exact physical source when the package was opened from
/// bytes, its object index, and the ingress limits the package was opened
/// with. The engines never publish through this trait.
pub trait TableSource {
    /// Every decoded IWA component of the package, in catalog order.
    fn catalog(&self) -> &ComponentCatalog;

    /// The exact physical source, when the package was opened from bytes.
    fn physical(&self) -> Option<&SourceCatalog>;

    /// Locate one object by identifier through the leaf's object index.
    ///
    /// # Errors
    ///
    /// [`ResolveError`] when the index names a position the catalog does not
    /// hold.
    fn resolve_object(&self, identifier: u64) -> Result<Option<ResolvedObject<'_>>, ResolveError>;

    /// The exact source bytes the package was opened from.
    fn source_bytes(&self) -> &[u8];

    /// The physical ingress limits the package was opened with.
    fn archive_limits(&self) -> litchi_iwa_archive::Limits;

    /// The semantic object ceiling the package was opened with.
    fn max_objects(&self) -> usize;

    /// The semantic reference ceiling the package was opened with.
    fn max_references(&self) -> usize;

    /// The package-level proofs the leaf memoizes for this package (record
    /// 0370k); `None` recomputes every proof. The proofs depend only on the
    /// immutable package, so a leaf keeps one set per opened package.
    fn proofs(&self) -> Option<&cell_format::control::PackageProofs> {
        None
    }
}
