//! Typed nodes and the generic node view (design §5; record 0010 §2–§3, §7.2).
//!
//! The planned `get`, `find`, `query` and `children` APIs yield a [`Node`]: one flat,
//! data-bearing enum with a variant per format, each wrapping that format's
//! own node, which carries the leaf's ADR 0007 values (decision D2). Typed
//! reads never build the generic view.
//!
//! On demand, any node projects into a [`View`]: OfficeCLI's eight-member
//! `DocumentNode` (`path`, `type`, `text`, `preview`, `style`, `childCount`,
//! `format`, `children`) with OfficeCLI's member order and omission rules
//! (decision D1). The view is a borrowed projection, never a model:
//!
//! - its `serde::Serialize` implementation streams straight into the
//!   serializer, with no intermediate `DocumentNode`, no per-node map and no
//!   path `String`;
//! - `format` is produced from a static per-kind property table whose keys
//!   are `&'static str`, with typed [`Value`]s that keep OfficeCLI's JSON
//!   types and print through the canonical [`crate::value`] renderers;
//! - [`ViewOptions`] bound the walk: serialization refuses `depth` above
//!   [`MAX_VIEW_DEPTH`] and more than `max_nodes` nodes is
//!   `output_too_large`.
//!
//! Writing JSON is `serde_json::to_writer(sink, &node.view())`;
//! [`View::serialize_into`] does the same and keeps the unified error typed.
//!
//! # Status
//!
//! DOCX, XLSX, XLSB and PPTX carry concrete leaf-backed nodes. Other formats retain the
//! uninhabited [`Pending`] placeholder until their adapters land. These nodes
//! are available through File navigation and lazy queries.

use std::borrow::Cow;
use std::convert::Infallible;
use std::fmt;
use std::marker::PhantomData;

use litchi_core::FileFormat;

use crate::error::Result;

/// Matches every [`Node`] variant, binding its payload to `$inner`, which
/// implements [`Project`]. The single place that lists the variants' cfgs.
macro_rules! dispatch {
    ($node:expr, $inner:ident => $body:expr) => {
        match $node {
            #[cfg(feature = "docx")]
            Node::Docx($inner) => $body,
            #[cfg(feature = "xlsx")]
            Node::Xlsx($inner) => $body,
            #[cfg(feature = "pptx")]
            Node::Pptx($inner) => $body,
            #[cfg(feature = "xlsb")]
            Node::Xlsb($inner) => $body,
            #[cfg(feature = "doc")]
            Node::Doc($inner) => $body,
            #[cfg(feature = "xls")]
            Node::Xls($inner) => $body,
            #[cfg(feature = "ppt")]
            Node::Ppt($inner) => $body,
            #[cfg(feature = "rtf")]
            Node::Rtf($inner) => $body,
            #[cfg(feature = "odt")]
            Node::Odt($inner) => $body,
            #[cfg(feature = "ods")]
            Node::Ods($inner) => $body,
            #[cfg(feature = "odp")]
            Node::Odp($inner) => $body,
            #[cfg(any(test, feature = "test-support"))]
            Node::Fake($inner) => $body,
            #[cfg(not(any(
                feature = "doc",
                feature = "docx",
                feature = "ppt",
                feature = "pptx",
                feature = "xls",
                feature = "xlsx",
                feature = "xlsb",
                feature = "rtf",
                feature = "odt",
                feature = "ods",
                feature = "odp",
                test,
                feature = "test-support"
            )))]
            Node::Unavailable($inner) => $body,
        }
    };
}

#[cfg(any(test, feature = "test-support"))]
#[doc(hidden)]
pub mod fake;
mod path;
mod serialize;
pub(crate) mod spec;
#[cfg(test)]
mod tests;
mod value;
mod view;

#[cfg(any(test, feature = "test-support"))]
pub(crate) use path::Ident;
#[cfg(test)]
pub(crate) use path::Select;
pub use path::{NodePath, PathStep};
/// Explicit native-identity path construction. Ordinary adapters use semantic names and positions.
pub mod raw {
    pub use super::path::Ident;
    /// Convert a grammar step, accepting a single native identity predicate.
    #[must_use]
    pub fn step_from_grammar<'a>(
        step: &'a litchi_core::selector::Step<'_>,
    ) -> Option<super::PathStep<'a>> {
        super::PathStep::from_raw_grammar(step)
    }

    /// Construct a step using a native identity; external targets are never followed.
    #[must_use]
    pub const fn identity_step<'a>(
        name: &'a str,
        key: &'a str,
        value: Ident<'a>,
    ) -> super::PathStep<'a> {
        super::PathStep::identity(name, key, value)
    }
}
pub use serialize::SerializeError;
pub use value::{Key, Value};
pub(crate) use view::Project;
// The query engine's fake format and the format adapters (R4-R7) read
// properties through it.
#[cfg(any(
    test,
    feature = "test-support",
    feature = "docx",
    feature = "xlsx",
    feature = "pptx",
    feature = "xlsb"
))]
pub(crate) use view::Source;
pub use view::{
    Children, DEFAULT_MAX_NODES, Format, MAX_VIEW_DEPTH, PREVIEW_CHARS, Preview, Prop, View,
    ViewOptions,
};

/// A typed node of an open document: one variant per format, each wrapping
/// that format's node (record 0010 §3.1).
///
/// A node borrows the `File` it came from (`'f`). Its generic accessors
/// dispatch with one `match`, statically; [`Node::view`] projects it into the
/// OfficeCLI-shaped [`View`] on demand.
///
/// ```
/// use litchi_ops::node::{Node, ViewOptions};
///
/// // Nodes come from `File::get`, `find`, `query` and `children`.
/// fn outline(node: &Node<'_>) -> serde_json::Result<String> {
///     let view = node.view_with(ViewOptions::default().depth(2));
///     serde_json::to_string_pretty(&view)
/// }
/// # let _ = outline;
/// ```
#[non_exhaustive]
#[allow(
    clippy::large_enum_variant,
    reason = "format nodes stay inline so spreadsheet traversal does not allocate a box per cell"
)]
pub enum Node<'f> {
    /// A WordprocessingML node (DOCX adapter, step R4a).
    #[cfg(feature = "docx")]
    Docx(crate::format::docx::Node<'f>),
    /// A SpreadsheetML node (XLSX adapter, step R4b).
    #[cfg(feature = "xlsx")]
    Xlsx(crate::format::xlsx::Node<'f>),
    /// A PresentationML node (PPTX adapter, step R4c).
    #[cfg(feature = "pptx")]
    Pptx(crate::format::pptx::Node<'f>),
    /// A BIFF12 workbook node (XLSB adapter, step R4d).
    #[cfg(feature = "xlsb")]
    Xlsb(crate::format::xlsb::Node<'f>),
    /// A legacy Word node (DOC adapter, step R6a).
    #[cfg(feature = "doc")]
    Doc(crate::format::doc::Node<'f>),
    /// A legacy Excel node (XLS adapter, step R6b).
    #[cfg(feature = "xls")]
    Xls(Pending<'f, crate::format::xls::Xls>),
    /// A legacy PowerPoint node (PPT adapter, step R6c).
    #[cfg(feature = "ppt")]
    Ppt(Pending<'f, crate::format::ppt::Ppt>),
    /// A Rich Text Format node (tier-3 adapter, step R7).
    #[cfg(feature = "rtf")]
    Rtf(Pending<'f, crate::format::rtf::Rtf>),
    /// An OpenDocument text node (tier-3 adapter, step R7).
    #[cfg(feature = "odt")]
    Odt(Pending<'f, crate::format::odt::Odt>),
    /// An OpenDocument spreadsheet node (tier-3 adapter, step R7).
    #[cfg(feature = "ods")]
    Ods(Pending<'f, crate::format::ods::Ods>),
    /// An OpenDocument presentation node (tier-3 adapter, step R7).
    #[cfg(feature = "odp")]
    Odp(Pending<'f, crate::format::odp::Odp>),
    /// A node of the test-only fake format.
    #[cfg(any(test, feature = "test-support"))]
    #[doc(hidden)]
    Fake(fake::Node<'f>),
    /// Keeps `'f` in use when no format is compiled in; uninhabited.
    #[cfg(not(any(
        feature = "doc",
        feature = "docx",
        feature = "ppt",
        feature = "pptx",
        feature = "xls",
        feature = "xlsx",
        feature = "xlsb",
        feature = "rtf",
        feature = "odt",
        feature = "ods",
        feature = "odp",
        test,
        feature = "test-support"
    )))]
    #[doc(hidden)]
    Unavailable(Pending<'f, ()>),
}

impl<'f> Node<'f> {
    /// The OfficeCLI `type` of the node (`paragraph`, `cell`, `slide`,
    /// `textbox`).
    #[must_use]
    pub fn kind(&self) -> &'static str {
        Project::kind(self)
    }

    /// The format of the document the node belongs to.
    #[must_use]
    pub fn format(&self) -> FileFormat {
        Project::file_format(self)
    }

    /// The canonical path, preferring stable identity where the element has
    /// one (`/body/p[@paraId=1A2B3C4D]`, `/slide[1]/shape[@id=5]`).
    #[must_use]
    pub fn path(&self) -> NodePath<'_> {
        Project::path(self)
    }

    /// The positional path (`/body/p[3]`), a Litchi extension: it never uses
    /// an identity form.
    #[must_use]
    pub fn positional_path(&self) -> NodePath<'_> {
        Project::positional_path(self)
    }

    /// Every child, in document order. Each item is fallible because a
    /// lazily decoded part can fail mid-walk; the iterator is fused after the
    /// first error.
    pub fn children(&self) -> Children<'_, 'f> {
        Project::children(self, &ViewOptions::DEFAULT)
    }

    /// The node view with the default options: one level of children.
    #[must_use]
    pub fn view(&self) -> View<'_, 'f> {
        View::new(self, ViewOptions::DEFAULT)
    }

    /// The node view with `options`.
    #[must_use]
    pub fn view_with(&self, options: ViewOptions) -> View<'_, 'f> {
        View::new(self, options)
    }

    /// The DOCX node, when the node belongs to a DOCX document.
    ///
    /// ```
    /// use litchi_ops::format::docx::Kind;
    /// use litchi_ops::node::Node;
    ///
    /// // Whether a node is a DOCX table cell.
    /// fn is_docx_cell(node: &Node<'_>) -> bool {
    ///     node.as_docx().is_some_and(|node| node.element_kind() == Kind::Cell)
    /// }
    /// # let _ = is_docx_cell;
    /// ```
    #[cfg(feature = "docx")]
    #[must_use]
    pub fn as_docx(&self) -> Option<&crate::format::docx::Node<'f>> {
        match self {
            Self::Docx(node) => Some(node),
            #[allow(
                unreachable_patterns,
                reason = "other formats' variants exist only with their features"
            )]
            _ => None,
        }
    }

    /// The legacy Word node, when the node belongs to a DOC document.
    #[cfg(feature = "doc")]
    #[must_use]
    pub fn as_doc(&self) -> Option<&crate::format::doc::Node<'f>> {
        match self {
            Self::Doc(node) => Some(node),
            #[allow(
                unreachable_patterns,
                reason = "other formats' variants exist only with their features"
            )]
            _ => None,
        }
    }

    /// The XLSX node, when the node belongs to an XLSX workbook.
    ///
    /// ```
    /// use litchi_ops::node::Node;
    ///
    /// // The sheet a workbook node belongs to.
    /// fn sheet_of<'f>(node: &Node<'f>) -> litchi_ops::Result<Option<&'f str>> {
    ///     match node.as_xlsx() { Some(node) => node.sheet_name(), None => Ok(None) }
    /// }
    /// # let _ = sheet_of;
    /// ```
    #[cfg(feature = "xlsx")]
    #[must_use]
    pub fn as_xlsx(&self) -> Option<&crate::format::xlsx::Node<'f>> {
        match self {
            Self::Xlsx(node) => Some(node),
            #[allow(
                unreachable_patterns,
                reason = "other formats' variants exist only with their features"
            )]
            _ => None,
        }
    }
    /// Borrow the format-specific presentation node.
    #[cfg(feature = "pptx")]
    #[must_use]
    pub fn as_pptx(&self) -> Option<&crate::format::pptx::Node<'f>> {
        match self {
            Self::Pptx(node) => Some(node),
            #[allow(unreachable_patterns, reason = "format variants depend on features")]
            _ => None,
        }
    }
}

impl<'f> Project<'f> for Node<'f> {
    fn kind(&self) -> &'static str {
        dispatch!(self, node => node.kind())
    }

    fn file_format(&self) -> FileFormat {
        dispatch!(self, node => node.file_format())
    }

    fn path(&self) -> NodePath<'_> {
        dispatch!(self, node => node.path())
    }

    fn positional_path(&self) -> NodePath<'_> {
        dispatch!(self, node => node.positional_path())
    }

    fn text(&self) -> Result<Option<Cow<'_, str>>> {
        dispatch!(self, node => node.text())
    }

    fn preview(&self) -> Result<Option<Preview<'_>>> {
        dispatch!(self, node => node.preview())
    }

    fn style(&self) -> Result<Option<Cow<'_, str>>> {
        dispatch!(self, node => node.style())
    }

    fn child_count(&self, options: &ViewOptions) -> Result<Option<u32>> {
        dispatch!(self, node => node.child_count(options))
    }

    fn format(&self) -> Result<Format<'_>> {
        dispatch!(self, node => Project::format(node))
    }

    fn children(&self, options: &ViewOptions) -> Children<'_, 'f> {
        dispatch!(self, node => Project::children(node, options))
    }
}

/// The payload of a [`Node`] variant whose format adapter has not landed
/// yet. It is uninhabited, so the variant cannot occur; the adapter replaces
/// it with the format's own node enum (record 0010 §3.1).
pub struct Pending<'f, H> {
    never: Infallible,
    _file: PhantomData<&'f H>,
}

impl<H> fmt::Debug for Pending<'_, H> {
    fn fmt(&self, _: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self.never {}
    }
}

impl<'f, H> Project<'f> for Pending<'f, H> {
    fn kind(&self) -> &'static str {
        match self.never {}
    }

    fn file_format(&self) -> FileFormat {
        match self.never {}
    }

    fn path(&self) -> NodePath<'_> {
        match self.never {}
    }

    fn text(&self) -> Result<Option<Cow<'_, str>>> {
        match self.never {}
    }

    fn child_count(&self, _: &ViewOptions) -> Result<Option<u32>> {
        match self.never {}
    }

    fn format(&self) -> Result<Format<'_>> {
        match self.never {}
    }

    fn children(&self, _: &ViewOptions) -> Children<'_, 'f> {
        match self.never {}
    }
}

const fn assert_send_sync<T: Send + Sync>() {}
const _: () = {
    assert_send_sync::<Node<'static>>();
    assert_send_sync::<View<'static, 'static>>();
    assert_send_sync::<Value<'static>>();
    assert_send_sync::<NodePath<'static>>();
};

impl fmt::Debug for Node<'_> {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        f.debug_struct("Node").finish_non_exhaustive()
    }
}
