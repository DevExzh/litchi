//! [`Matches`]: the streaming result of [`File::query`](crate::File::query)
//! (record 0010 §5.1, §8.1).

// In a build without any adapter, `Matches` is uninhabited and its methods
// are unreachable; those lints are allowed in such builds only. Each adapter
// adds its feature to this list.
#![cfg_attr(
    not(any(
        feature = "doc",
        feature = "docx",
        feature = "xlsx",
        feature = "pptx",
        feature = "xlsb"
    )),
    allow(
        dead_code,
        unreachable_code,
        unused_imports,
        unused_macros,
        unused_mut,
        unused_variables
    )
)]

use std::convert::Infallible;
use std::fmt;
use std::iter::FusedIterator;
use std::marker::PhantomData;

use super::plan::{Diagnostic, Plan};
#[cfg(any(
    feature = "doc",
    feature = "docx",
    feature = "xlsx",
    feature = "pptx",
    feature = "xlsb"
))]
use super::walk::Walk;
use crate::Result;
use crate::node::Node;

/// The nodes a query matches, in document order: a lazy, fallible iterator.
///
/// Nothing is read until the first match is pulled, and the walk visits
/// nothing beyond the match the caller last pulled. [`Matches::limit`] and
/// [`Matches::first`] stop it early; zero matches is an empty iterator. Each
/// item is fallible, because a part read lazily during the walk can fail;
/// after the first error the iterator is fused.
///
/// ```
/// # #[cfg(feature = "docx")]
/// # {
/// let file = litchi_ops::open(concat!(env!("CARGO_MANIFEST_DIR"),
///     "/../../test-data/ooxml/docx/drawing.docx"))?;
/// let first = file.query("p")?.limit(1).collect::<litchi_ops::Result<Vec<_>>>()?;
/// assert!(first.len() <= 1);
/// # }
/// # Ok::<(), litchi_ops::Error>(())
/// ```
/// A query cannot outlive its file:
///
/// ```compile_fail
/// fn escape(file: litchi_ops::File) -> litchi_ops::query::Matches<'static, 'static> {
///     file.query("p").unwrap()
/// }
/// ```
pub struct Matches<'f, 'q> {
    inner: Inner<'f, 'q>,
}

/// One variant per format adapter (static dispatch, ADR 0004).
#[allow(
    clippy::large_enum_variant,
    reason = "a walk is built once per query and iterated in place; boxing it would add an \
              allocation to every query, and the other variant is uninhabited"
)]
enum Inner<'f, 'q> {
    #[cfg(feature = "docx")]
    Docx(Walk<'q, 'f, crate::format::docx::Nav<'f>>),
    #[cfg(feature = "doc")]
    Doc(Walk<'q, 'f, crate::format::doc::Nav<'f>>),
    #[cfg(feature = "xlsx")]
    Xlsx(Walk<'q, 'f, crate::format::xlsx::Nav<'f>>),
    #[cfg(feature = "pptx")]
    Pptx(Walk<'q, 'f, crate::format::pptx::Nav<'f>>),
    #[cfg(feature = "xlsb")]
    Xlsb(Walk<'q, 'f, crate::format::xlsb::Nav<'f>>),
    /// Keeps the lifetimes in use in builds without an adapter; uninhabited.
    #[allow(dead_code, reason = "uninhabited: only keeps the lifetimes in use")]
    Never(Infallible, PhantomData<(&'f (), &'q ())>),
}

impl<'f, 'q> Matches<'f, 'q> {
    /// The matches of a DOC walk.
    #[cfg(feature = "doc")]
    pub(crate) fn doc(walk: Walk<'q, 'f, crate::format::doc::Nav<'f>>) -> Self {
        Self {
            inner: Inner::Doc(walk),
        }
    }

    /// The matches of an DOCX walk.
    #[cfg(feature = "docx")]
    pub(crate) fn docx(walk: Walk<'q, 'f, crate::format::docx::Nav<'f>>) -> Self {
        Self {
            inner: Inner::Docx(walk),
        }
    }

    /// The matches of an XLSX walk.
    #[cfg(feature = "xlsx")]
    pub(crate) fn xlsx(walk: Walk<'q, 'f, crate::format::xlsx::Nav<'f>>) -> Self {
        Self {
            inner: Inner::Xlsx(walk),
        }
    }
    #[cfg(feature = "pptx")]
    pub(crate) fn pptx(walk: Walk<'q, 'f, crate::format::pptx::Nav<'f>>) -> Self {
        Self {
            inner: Inner::Pptx(walk),
        }
    }
    #[cfg(feature = "xlsb")]
    pub(crate) fn xlsb(walk: Walk<'q, 'f, crate::format::xlsb::Nav<'f>>) -> Self {
        Self {
            inner: Inner::Xlsb(walk),
        }
    }

    /// The same matches, stopping after `n` in all: the walk ends there,
    /// reading nothing more.
    ///
    #[must_use]
    pub fn limit(self, n: usize) -> Self {
        match self.inner {
            #[cfg(feature = "doc")]
            Inner::Doc(walk) => Self {
                inner: Inner::Doc(walk.limit(n)),
            },
            #[cfg(feature = "docx")]
            Inner::Docx(walk) => Self {
                inner: Inner::Docx(walk.limit(n)),
            },
            #[cfg(feature = "xlsx")]
            Inner::Xlsx(walk) => Self {
                inner: Inner::Xlsx(walk.limit(n)),
            },
            #[cfg(feature = "pptx")]
            Inner::Pptx(walk) => Self {
                inner: Inner::Pptx(walk.limit(n)),
            },
            #[cfg(feature = "xlsb")]
            Inner::Xlsb(walk) => Self {
                inner: Inner::Xlsb(walk.limit(n)),
            },
            Inner::Never(never, _) => match never {},
        }
    }

    /// The first match, or `None`; the walk stops there.
    ///
    /// # Errors
    ///
    /// The error that ended the walk before a first match.
    ///
    pub fn first(self) -> Result<Option<Node<'f>>> {
        self.limit(1).next().transpose()
    }

    /// The compiled plan the walk runs.
    #[must_use]
    pub fn plan(&self) -> &Plan<'q> {
        match &self.inner {
            #[cfg(feature = "doc")]
            Inner::Doc(walk) => walk.plan(),
            #[cfg(feature = "docx")]
            Inner::Docx(walk) => walk.plan(),
            #[cfg(feature = "xlsx")]
            Inner::Xlsx(walk) => walk.plan(),
            #[cfg(feature = "pptx")]
            Inner::Pptx(walk) => walk.plan(),
            #[cfg(feature = "xlsb")]
            Inner::Xlsb(walk) => walk.plan(),
            Inner::Never(never, _) => match *never {},
        }
    }

    /// The compile-time findings that do not stop the query, such as an
    /// unknown key ([`Plan::diagnostics`]).
    #[must_use]
    pub fn diagnostics(&self) -> &[Diagnostic] {
        self.plan().diagnostics()
    }
}

impl<'f> Iterator for Matches<'f, '_> {
    type Item = Result<Node<'f>>;

    fn next(&mut self) -> Option<Self::Item> {
        match &mut self.inner {
            #[cfg(feature = "doc")]
            Inner::Doc(walk) => walk.next().map(|node| node.map(Node::Doc)),
            #[cfg(feature = "docx")]
            Inner::Docx(walk) => walk.next().map(|node| node.map(Node::Docx)),
            #[cfg(feature = "xlsx")]
            Inner::Xlsx(walk) => walk.next().map(|node| node.map(Node::Xlsx)),
            #[cfg(feature = "pptx")]
            Inner::Pptx(walk) => walk.next().map(|node| node.map(Node::Pptx)),
            #[cfg(feature = "xlsb")]
            Inner::Xlsb(walk) => walk.next().map(|node| node.map(Node::Xlsb)),
            Inner::Never(never, _) => match *never {},
        }
    }
}

impl FusedIterator for Matches<'_, '_> {}

impl fmt::Debug for Matches<'_, '_> {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        f.debug_struct("Matches")
            .field("format", &self.plan().format())
            .finish_non_exhaustive()
    }
}
