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
        feature = "ppt",
        feature = "xlsb",
        feature = "xls",
        feature = "rtf",
        feature = "ods",
        feature = "odt",
        feature = "odp",
        odf_objects,
        feature = "iwork-adapter"
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
    feature = "ppt",
    feature = "xlsb",
    feature = "xls",
    feature = "rtf",
    feature = "ods",
    feature = "odt",
    feature = "odp"
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
    /// The walk takes its plan's fast step (record 0315), decided once
    /// here so the general route's code stays as it was.
    fast: bool,
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
    #[cfg(feature = "ppt")]
    Ppt(Walk<'q, 'f, crate::format::ppt::Nav<'f>>),
    #[cfg(feature = "xlsb")]
    Xlsb(Walk<'q, 'f, crate::format::xlsb::Nav<'f>>),
    #[cfg(feature = "xls")]
    Xls(Walk<'q, 'f, crate::format::xls::Nav<'f>>),
    #[cfg(feature = "odt")]
    Odt(Walk<'q, 'f, crate::format::odt::Nav<'f>>),
    #[cfg(feature = "rtf")]
    Rtf(Walk<'q, 'f, crate::format::rtf::Nav<'f>>),
    #[cfg(feature = "ods")]
    Ods(Walk<'q, 'f, crate::format::ods::Nav<'f>>),
    #[cfg(feature = "odp")]
    Odp(Walk<'q, 'f, crate::format::odp::Nav<'f>>),
    #[cfg(odf_objects)]
    Object(crate::format::odf_tree::Walks<'q, 'f>),
    #[cfg(feature = "iwork-adapter")]
    Iwork(crate::format::iwork::Walks<'q, 'f>),
    /// Keeps the lifetimes in use in builds without an adapter; uninhabited.
    #[allow(dead_code, reason = "uninhabited: only keeps the lifetimes in use")]
    Never(Infallible, PhantomData<(&'f (), &'q ())>),
}

impl<'f, 'q> Matches<'f, 'q> {
    /// The matches of a DOC walk.
    #[cfg(feature = "doc")]
    pub(crate) fn doc(walk: Walk<'q, 'f, crate::format::doc::Nav<'f>>) -> Self {
        Self {
            fast: walk.is_fast(),
            inner: Inner::Doc(walk),
        }
    }

    /// The matches of an DOCX walk.
    #[cfg(feature = "docx")]
    pub(crate) fn docx(walk: Walk<'q, 'f, crate::format::docx::Nav<'f>>) -> Self {
        Self {
            fast: walk.is_fast(),
            inner: Inner::Docx(walk),
        }
    }

    /// The matches of an XLSX walk.
    #[cfg(feature = "xlsx")]
    pub(crate) fn xlsx(walk: Walk<'q, 'f, crate::format::xlsx::Nav<'f>>) -> Self {
        Self {
            fast: walk.is_fast(),
            inner: Inner::Xlsx(walk),
        }
    }
    #[cfg(feature = "pptx")]
    pub(crate) fn pptx(walk: Walk<'q, 'f, crate::format::pptx::Nav<'f>>) -> Self {
        Self {
            fast: walk.is_fast(),
            inner: Inner::Pptx(walk),
        }
    }
    #[cfg(feature = "ppt")]
    pub(crate) fn ppt(walk: Walk<'q, 'f, crate::format::ppt::Nav<'f>>) -> Self {
        Self {
            fast: walk.is_fast(),
            inner: Inner::Ppt(walk),
        }
    }
    #[cfg(feature = "xlsb")]
    pub(crate) fn xlsb(walk: Walk<'q, 'f, crate::format::xlsb::Nav<'f>>) -> Self {
        Self {
            fast: walk.is_fast(),
            inner: Inner::Xlsb(walk),
        }
    }
    #[cfg(feature = "xls")]
    pub(crate) fn xls(walk: Walk<'q, 'f, crate::format::xls::Nav<'f>>) -> Self {
        Self {
            fast: walk.is_fast(),
            inner: Inner::Xls(walk),
        }
    }
    /// The matches of an RTF walk.
    #[cfg(feature = "rtf")]
    pub(crate) fn rtf(walk: Walk<'q, 'f, crate::format::rtf::Nav<'f>>) -> Self {
        Self {
            fast: walk.is_fast(),
            inner: Inner::Rtf(walk),
        }
    }
    /// The matches of an ODS walk.
    #[cfg(feature = "ods")]
    pub(crate) fn ods(walk: Walk<'q, 'f, crate::format::ods::Nav<'f>>) -> Self {
        Self {
            fast: walk.is_fast(),
            inner: Inner::Ods(walk),
        }
    }

    /// The matches of an ODT walk.
    #[cfg(feature = "odt")]
    pub(crate) fn odt(walk: Walk<'q, 'f, crate::format::odt::Nav<'f>>) -> Self {
        Self {
            fast: walk.is_fast(),
            inner: Inner::Odt(walk),
        }
    }

    /// The matches of an ODP walk.
    #[cfg(feature = "odp")]
    pub(crate) fn odp(walk: Walk<'q, 'f, crate::format::odp::Nav<'f>>) -> Self {
        Self {
            fast: walk.is_fast(),
            inner: Inner::Odp(walk),
        }
    }

    /// The matches of an object-tree walk (ODG, ODC, ODI, ODB, ODF formula).
    #[cfg(odf_objects)]
    pub(crate) fn object(walks: crate::format::odf_tree::Walks<'q, 'f>) -> Self {
        Self {
            inner: Inner::Object(walks),
            fast: false,
        }
    }

    /// The matches of an iWork object-tree walk (Pages, Keynote, Numbers).
    #[cfg(feature = "iwork-adapter")]
    pub(crate) fn iwork(walks: crate::format::iwork::Walks<'q, 'f>) -> Self {
        Self {
            inner: Inner::Iwork(walks),
            fast: false,
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
                fast: self.fast,
            },
            #[cfg(feature = "docx")]
            Inner::Docx(walk) => Self {
                inner: Inner::Docx(walk.limit(n)),
                fast: self.fast,
            },
            #[cfg(feature = "xlsx")]
            Inner::Xlsx(walk) => Self {
                inner: Inner::Xlsx(walk.limit(n)),
                fast: self.fast,
            },
            #[cfg(feature = "pptx")]
            Inner::Pptx(walk) => Self {
                inner: Inner::Pptx(walk.limit(n)),
                fast: self.fast,
            },
            #[cfg(feature = "ppt")]
            Inner::Ppt(walk) => Self {
                inner: Inner::Ppt(walk.limit(n)),
                fast: self.fast,
            },
            #[cfg(feature = "xlsb")]
            Inner::Xlsb(walk) => Self {
                inner: Inner::Xlsb(walk.limit(n)),
                fast: self.fast,
            },
            #[cfg(feature = "xls")]
            Inner::Xls(walk) => Self {
                inner: Inner::Xls(walk.limit(n)),
                fast: self.fast,
            },
            #[cfg(feature = "rtf")]
            Inner::Rtf(walk) => Self {
                inner: Inner::Rtf(walk.limit(n)),
                fast: self.fast,
            },
            #[cfg(feature = "ods")]
            Inner::Ods(walk) => Self {
                inner: Inner::Ods(walk.limit(n)),
                fast: self.fast,
            },
            #[cfg(feature = "odt")]
            Inner::Odt(walk) => Self {
                inner: Inner::Odt(walk.limit(n)),
                fast: self.fast,
            },
            #[cfg(feature = "odp")]
            Inner::Odp(walk) => Self {
                inner: Inner::Odp(walk.limit(n)),
                fast: self.fast,
            },
            #[cfg(odf_objects)]
            Inner::Object(walks) => Self {
                inner: Inner::Object(walks.limit(n)),
                fast: self.fast,
            },
            #[cfg(feature = "iwork-adapter")]
            Inner::Iwork(walks) => Self {
                inner: Inner::Iwork(walks.limit(n)),
                fast: self.fast,
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
            #[cfg(feature = "ppt")]
            Inner::Ppt(walk) => walk.plan(),
            #[cfg(feature = "xlsb")]
            Inner::Xlsb(walk) => walk.plan(),
            #[cfg(feature = "xls")]
            Inner::Xls(walk) => walk.plan(),
            #[cfg(feature = "rtf")]
            Inner::Rtf(walk) => walk.plan(),
            #[cfg(feature = "odt")]
            Inner::Odt(walk) => walk.plan(),
            #[cfg(feature = "ods")]
            Inner::Ods(walk) => walk.plan(),
            #[cfg(feature = "odp")]
            Inner::Odp(walk) => walk.plan(),
            #[cfg(odf_objects)]
            Inner::Object(walks) => walks.plan(),
            #[cfg(feature = "iwork-adapter")]
            Inner::Iwork(walks) => walks.plan(),
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
        if self.fast {
            return self.next_fast();
        }
        match &mut self.inner {
            #[cfg(feature = "doc")]
            Inner::Doc(walk) => walk.next_map(Node::Doc),
            #[cfg(feature = "docx")]
            Inner::Docx(walk) => walk.next_map(Node::Docx),
            #[cfg(feature = "xlsx")]
            Inner::Xlsx(walk) => walk.next_map(Node::Xlsx),
            #[cfg(feature = "pptx")]
            Inner::Pptx(walk) => walk.next_map(Node::Pptx),
            #[cfg(feature = "ppt")]
            Inner::Ppt(walk) => walk.next_map(Node::Ppt),
            #[cfg(feature = "xlsb")]
            Inner::Xlsb(walk) => walk.next_map(Node::Xlsb),
            #[cfg(feature = "xls")]
            Inner::Xls(walk) => walk.next_map(Node::Xls),
            #[cfg(feature = "odt")]
            Inner::Odt(walk) => walk.next_map(Node::Odt),
            #[cfg(feature = "rtf")]
            Inner::Rtf(walk) => walk.next_map(Node::Rtf),
            #[cfg(feature = "ods")]
            Inner::Ods(walk) => walk.next_map(Node::Ods),
            #[cfg(feature = "odp")]
            Inner::Odp(walk) => walk.next_map(Node::Odp),
            #[cfg(odf_objects)]
            Inner::Object(walks) => walks.next(),
            #[cfg(feature = "iwork-adapter")]
            Inner::Iwork(walks) => walks.next(),
            Inner::Never(never, _) => match *never {},
        }
    }
}

impl<'f> Matches<'f, '_> {
    /// [`Iterator::next`] for a walk on its plan's fast step (record 0315).
    fn next_fast(&mut self) -> Option<Result<Node<'f>>> {
        match &mut self.inner {
            #[cfg(feature = "doc")]
            Inner::Doc(walk) => walk.next_fast_map(Node::Doc),
            #[cfg(feature = "docx")]
            Inner::Docx(walk) => walk.next_fast_map(Node::Docx),
            #[cfg(feature = "xlsx")]
            Inner::Xlsx(walk) => walk.next_fast_map(Node::Xlsx),
            #[cfg(feature = "pptx")]
            Inner::Pptx(walk) => walk.next_fast_map(Node::Pptx),
            #[cfg(feature = "ppt")]
            Inner::Ppt(walk) => walk.next_fast_map(Node::Ppt),
            #[cfg(feature = "xlsb")]
            Inner::Xlsb(walk) => walk.next_fast_map(Node::Xlsb),
            #[cfg(feature = "xls")]
            Inner::Xls(walk) => walk.next_fast_map(Node::Xls),
            #[cfg(feature = "odt")]
            Inner::Odt(walk) => walk.next_fast_map(Node::Odt),
            #[cfg(feature = "rtf")]
            Inner::Rtf(walk) => walk.next_fast_map(Node::Rtf),
            #[cfg(feature = "ods")]
            Inner::Ods(walk) => walk.next_fast_map(Node::Ods),
            #[cfg(feature = "odp")]
            Inner::Odp(walk) => walk.next_fast_map(Node::Odp),
            #[cfg(odf_objects)]
            Inner::Object(walks) => walks.next(),
            #[cfg(feature = "iwork-adapter")]
            Inner::Iwork(walks) => walks.next(),
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
