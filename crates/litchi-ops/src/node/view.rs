//! The node view: a borrowed projection of one typed node into OfficeCLI's
//! `DocumentNode` members (record 0010 §2.1–§2.5).

use std::borrow::Cow;
use std::fmt;
use std::iter::FusedIterator;
use std::marker::PhantomData;

use litchi_core::FileFormat;
use serde::ser::{Serialize, Serializer};

use super::Node;
use super::path::NodePath;
use super::spec::PropSpec;
use super::value::{Key, Value};
use crate::error::{Error, Result};

/// The deepest `depth` a view honours; larger requests are refused. It bounds
/// the serializer's recursion (record 0010 §2.5; OfficeCLI clamps to 256 and
/// has no output cap, M1-Δ7).
pub const MAX_VIEW_DEPTH: u32 = 64;

/// The default node budget of a view: serializing more nodes is
/// `output_too_large` (ADR 0006 bounded output).
pub const DEFAULT_MAX_NODES: u32 = 100_000;

/// How many characters a [`Preview::Prefix`] keeps before `...`, as OfficeCLI
/// does for Word paragraphs and slide shapes.
pub const PREVIEW_CHARS: usize = 50;

/// Options of a node view.
///
/// ```
/// use litchi_ops::node::{DEFAULT_MAX_NODES, MAX_VIEW_DEPTH, ViewOptions};
///
/// let options = ViewOptions::default();
/// assert_eq!((options.depth, options.max_nodes, options.dense), (1, DEFAULT_MAX_NODES, None));
///
/// let options = ViewOptions::default().depth(1_000).max_nodes(500).dense(256);
/// assert_eq!(options.depth, 1_000);
/// assert_eq!((options.max_nodes, options.dense), (500, Some(256)));
/// ```
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub struct ViewOptions {
    /// Child levels listed below the node: 0 lists none, 1 (the default, as
    /// in OfficeCLI's CLI) lists the direct children. Values above
    /// [`MAX_VIEW_DEPTH`] are refused during serialization.
    pub depth: u32,
    /// The most nodes one serialization emits, the viewed node included;
    /// one more is `output_too_large`. Defaults to [`DEFAULT_MAX_NODES`].
    pub max_nodes: u32,
    /// For ranges: list empty-cell stubs up to this many cells, instead of
    /// stored cells only (ADR 0007's budgeted `grid(range)` versus sparse
    /// `cells(range)`). `None` by default.
    pub dense: Option<u32>,
}

impl ViewOptions {
    /// The default options: depth 1, [`DEFAULT_MAX_NODES`], stored cells only.
    pub const DEFAULT: Self = Self {
        depth: 1,
        max_nodes: DEFAULT_MAX_NODES,
        dense: None,
    };

    /// The same options with `depth` child levels, refused during serialization above
    /// [`MAX_VIEW_DEPTH`].
    #[must_use]
    pub const fn depth(mut self, depth: u32) -> Self {
        self.depth = depth;
        self
    }

    /// The same options with a budget of `max_nodes` nodes.
    #[must_use]
    pub const fn max_nodes(mut self, max_nodes: u32) -> Self {
        self.max_nodes = max_nodes;
        self
    }

    /// The same options listing ranges densely, up to `cells` cells.
    #[must_use]
    pub const fn dense(mut self, cells: u32) -> Self {
        self.dense = Some(cells);
        self
    }

    /// The options a child view uses: one level less.
    const fn child(self) -> Self {
        Self {
            depth: self.depth.saturating_sub(1),
            ..self
        }
    }
}

impl Default for ViewOptions {
    fn default() -> Self {
        Self::DEFAULT
    }
}

/// The `preview` member: a short label for the node, written without
/// allocating.
///
/// Word paragraphs and slide shapes preview their text cut to
/// [`PREVIEW_CHARS`] characters plus `...` (OfficeCLI cuts UTF-16 code units;
/// Litchi cuts characters); a shape without text previews its name, a slide
/// its title or `(untitled)`, a cell its A1 reference (record 0010 §2.4).
///
/// ```
/// use std::borrow::Cow;
///
/// use litchi_ops::node::Preview;
///
/// let long = "a".repeat(60);
/// let preview = Preview::Prefix(Cow::Borrowed(&long));
/// assert_eq!(preview.to_string(), format!("{}...", "a".repeat(50)));
/// assert_eq!(Preview::Prefix(Cow::Borrowed("short")).to_string(), "short");
/// assert_eq!(Preview::Text(Cow::Borrowed("(untitled)")).to_string(), "(untitled)");
/// ```
#[derive(Clone, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum Preview<'a> {
    /// The label as it is.
    Text(Cow<'a, str>),
    /// The text, cut to its first [`PREVIEW_CHARS`] characters followed by
    /// `...` when it is longer.
    Prefix(Cow<'a, str>),
}

impl Preview<'_> {
    /// The same preview, borrowing its text. Never allocates.
    #[must_use]
    pub fn borrowed(&self) -> Preview<'_> {
        match self {
            Self::Text(text) => Preview::Text(Cow::Borrowed(text)),
            Self::Prefix(text) => Preview::Prefix(Cow::Borrowed(text)),
        }
    }
}

impl fmt::Display for Preview<'_> {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Self::Text(text) => f.write_str(text),
            Self::Prefix(text) => match text.char_indices().nth(PREVIEW_CHARS) {
                Some((end, _)) => {
                    f.write_str(&text[..end])?;
                    f.write_str("...")
                },
                None => f.write_str(text),
            },
        }
    }
}

/// What a per-format node supplies to the generic view. Each format
/// adapter's node implements it; [`Node`] dispatches with one `match`.
///
/// Each method reads only what its member needs: `child_count` returns
/// `None` rather than decode a part the request did not touch (M1-Δ1), and
/// `format` parses the node's property set once.
pub(crate) trait Project<'f> {
    /// The OfficeCLI `type`.
    fn kind(&self) -> &'static str;

    /// The format of the node's document.
    fn file_format(&self) -> FileFormat;

    /// The canonical path.
    fn path(&self) -> NodePath<'_>;

    /// The positional path; the canonical one when the kind has no identity
    /// form.
    fn positional_path(&self) -> NodePath<'_> {
        self.path()
    }

    /// The `text` member; `None` when the element has no text.
    fn text(&self) -> Result<Option<Cow<'_, str>>>;

    /// The `preview` member, per the OfficeCLI rule for the kind.
    fn preview(&self) -> Result<Option<Preview<'_>>> {
        Ok(None)
    }

    /// The `style` member, set for Word paragraphs only.
    fn style(&self) -> Result<Option<Cow<'_, str>>> {
        Ok(None)
    }

    /// The number of children the view would list at depth one or more, or
    /// `None` when counting would decode an untouched part.
    fn child_count(&self, options: &ViewOptions) -> Result<Option<u32>>;

    /// The `format` entries, in the order of the kind's property table.
    fn format(&self) -> Result<Format<'_>>;

    /// The children, in document order; `options` carries `dense` for
    /// ranges.
    fn children(&self, options: &ViewOptions) -> Children<'_, 'f>;
}

/// The node view: a borrowed projection of one [`Node`] into OfficeCLI's
/// `DocumentNode` members, read through accessors or serialized.
///
/// A view is never stored and builds no model: each accessor reads the
/// typed node, and `serde::Serialize` streams the members (`path`, `type`,
/// `text`, `preview`, `style`, `childCount`, `format`, `children`) in
/// OfficeCLI's order, omitting `text`, `preview` and `style` when absent, as
/// OfficeCLI's `WhenWritingNull` does. `childCount` is omitted only when
/// counting would decode a part the request did not touch (M1-Δ1).
///
/// ```
/// use litchi_ops::node::{Node, ViewOptions};
///
/// fn summary(node: &Node<'_>) -> litchi_ops::Result<String> {
///     let view = node.view();
///     let mut out = format!("{} {}", view.kind(), view.path());
///     for prop in view.format()? {
///         out.push_str(&format!(" {}={}", prop.key(), prop));
///     }
///     Ok(out)
/// }
/// # let _ = summary;
/// ```
#[derive(Debug, Clone, Copy)]
pub struct View<'n, 'f> {
    node: &'n Node<'f>,
    options: ViewOptions,
}

impl<'n, 'f> View<'n, 'f> {
    /// The view of `node` with `options`; `depth` is refused during serialization above
    /// [`MAX_VIEW_DEPTH`].
    #[must_use]
    pub fn new(node: &'n Node<'f>, options: ViewOptions) -> Self {
        let depth = options.depth;
        Self {
            node,
            options: options.depth(depth),
        }
    }

    /// The viewed node.
    #[must_use]
    pub const fn node(&self) -> &'n Node<'f> {
        self.node
    }

    /// The options, as requested.
    #[must_use]
    pub const fn options(&self) -> ViewOptions {
        self.options
    }

    /// The options of this view's child views: one level less.
    #[must_use]
    pub const fn child_options(&self) -> ViewOptions {
        self.options.child()
    }

    /// The `path` member: the canonical path, written without allocating.
    #[must_use]
    pub fn path(&self) -> NodePath<'n> {
        Project::path(self.node)
    }

    /// The `type` member.
    #[must_use]
    pub fn kind(&self) -> &'static str {
        Project::kind(self.node)
    }

    /// The `text` member, borrowed where the leaf allows.
    ///
    /// # Errors
    ///
    /// The leaf's error when reading the text decodes a part that fails.
    pub fn text(&self) -> Result<Option<Cow<'n, str>>> {
        Project::text(self.node)
    }

    /// The `preview` member.
    ///
    /// # Errors
    ///
    /// The leaf's error when reading the preview decodes a part that fails.
    pub fn preview(&self) -> Result<Option<Preview<'n>>> {
        Project::preview(self.node)
    }

    /// The `style` member (Word paragraphs: the style id).
    ///
    /// # Errors
    ///
    /// The leaf's error when reading the style decodes a part that fails.
    pub fn style(&self) -> Result<Option<Cow<'n, str>>> {
        Project::style(self.node)
    }

    /// The `childCount` member: the children the view lists at depth one or
    /// more, or `None` when counting would decode an untouched part.
    ///
    /// # Errors
    ///
    /// The leaf's error when counting decodes a part that fails.
    pub fn child_count(&self) -> Result<Option<u32>> {
        Project::child_count(self.node, &self.options)
    }

    /// The `format` member: typed entries in the order of the kind's
    /// property table, local (not effective) properties only.
    ///
    /// # Errors
    ///
    /// The leaf's error when parsing the node's property set fails.
    pub fn format(&self) -> Result<Format<'n>> {
        Project::format(self.node)
    }

    /// The `children` member: the child nodes, in document order, or nothing
    /// at depth 0. View each one with [`View::child_options`].
    pub fn children(&self) -> Children<'n, 'f> {
        if self.options.depth == 0 {
            Children::empty()
        } else {
            Project::children(self.node, &self.options)
        }
    }
}

/// The entries of a node's `format` member, in the order of the kind's
/// property table. Reading them allocates nothing beyond what the leaf
/// accessor returns.
///
/// ```
/// use litchi_ops::node::{Node, Value};
///
/// // The `format` entries as `key=value` text, as OfficeCLI shows them.
/// fn properties(node: &Node<'_>) -> litchi_ops::Result<Vec<String>> {
///     let format = node.view().format()?;
///     Ok(format.map(|prop| format!("{}={prop}", prop.key())).collect())
/// }
/// # let _ = properties;
/// ```
pub struct Format<'a> {
    props: &'static [PropSpec],
    hidden: &'static [&'static str],
    next: usize,
    source: Source<'a>,
}

/// Where a [`Format`] reads its values: the node's parsed property set, one
/// variant per format adapter.
#[derive(Debug)]
#[cfg_attr(
    not(test),
    allow(dead_code, reason = "constructed by the format adapters (R4-R7)")
)]
pub(crate) enum Source<'a> {
    /// A node without readable properties.
    Empty(PhantomData<&'a ()>),
    /// The fake format: values aligned with the property table.
    #[cfg(any(test, feature = "test-support"))]
    Fake(&'a [Option<Value<'static>>]),
    /// An RTF node's borrowed native values.
    #[cfg(feature = "rtf")]
    Rtf(crate::format::rtf::Props<'a>),
    /// A DOCX node's parsed property set.
    #[cfg(feature = "docx")]
    Docx(crate::format::docx::Props<'a>),
    /// An XLSX node's leaf values.
    #[cfg(feature = "xlsx")]
    Xlsx(crate::format::xlsx::Props<'a>),
    #[cfg(feature = "pptx")]
    Pptx(crate::format::pptx::Props<'a>),
    #[cfg(feature = "ppt")]
    Ppt(crate::format::ppt::Props<'a>),
    #[cfg(feature = "xlsb")]
    Xlsb(crate::format::xlsb::Props<'a>),
    /// An XLS node's leaf values.
    #[cfg(feature = "xls")]
    Xls(crate::format::xls::Props<'a>),
    /// A DOC node's leaf values.
    #[cfg(feature = "doc")]
    Doc(crate::format::doc::NodeProps<'a>),
    /// An ODS node's leaf values.
    #[cfg(feature = "ods")]
    Ods(crate::format::ods::Props<'a>),
    /// An ODT node's leaf values.
    #[cfg(feature = "odt")]
    Odt(crate::format::odt::Props<'a>),
    /// An ODP node's scene values.
    #[cfg(feature = "odp")]
    Odp(crate::format::odp::Props<'a>),
    /// An object-tree node's values, aligned with its property table.
    #[cfg(any(odf_objects, feature = "iwork-adapter"))]
    Tree(&'a [Option<Value<'static>>]),
}

impl<'a> Source<'a> {
    /// The value of the property at `index`, if the node has it. Values
    /// borrow the node (`'a`), never the iterator, so a `Format` can hand
    /// them out.
    fn get(&self, index: usize) -> Option<Value<'a>> {
        match *self {
            Self::Empty(_) => {
                // Nothing is set; the index only matters to a format's source.
                let _ = index;
                None
            },
            #[cfg(any(test, feature = "test-support"))]
            Self::Fake(values) => values.get(index)?.as_ref().map(Value::borrowed),
            #[cfg(feature = "rtf")]
            Self::Rtf(ref props) => props.get(index),
            #[cfg(feature = "docx")]
            Self::Docx(ref props) => props.get(index),
            #[cfg(feature = "xlsx")]
            Self::Xlsx(ref props) => props.get(index),
            #[cfg(feature = "pptx")]
            Self::Pptx(ref props) => props.get(index),
            #[cfg(feature = "ppt")]
            Self::Ppt(ref props) => props.get(index),
            #[cfg(feature = "xlsb")]
            Self::Xlsb(ref props) => props.get(index),
            #[cfg(feature = "xls")]
            Self::Xls(ref props) => props.get(index),
            #[cfg(feature = "doc")]
            Self::Doc(ref props) => props.get(index),
            #[cfg(feature = "ods")]
            Self::Ods(ref props) => props.get(index),
            #[cfg(feature = "odt")]
            Self::Odt(ref props) => props.get(index),
            #[cfg(feature = "odp")]
            Self::Odp(ref props) => props.get(index),
            #[cfg(any(odf_objects, feature = "iwork-adapter"))]
            Self::Tree(values) => values.get(index)?.as_ref().map(Value::borrowed),
        }
    }
}

#[cfg_attr(
    not(test),
    allow(dead_code, reason = "constructed by the format adapters (R4-R7)")
)]
impl<'a> Format<'a> {
    /// The entries of `source` for a kind whose property table is `props`.
    pub(crate) const fn new(props: &'static [PropSpec], source: Source<'a>) -> Self {
        Self {
            props,
            hidden: &[],
            next: 0,
            source,
        }
    }

    /// Omit raw properties from ordinary projection; indexed query access remains explicit.
    #[cfg(any(feature = "doc", feature = "docx", feature = "rtf", feature = "odt"))]
    pub(crate) const fn without(mut self, hidden: &'static [&'static str]) -> Self {
        self.hidden = hidden;
        self
    }

    /// No entries.
    pub(crate) const fn empty() -> Self {
        Self::new(&[], Source::Empty(PhantomData))
    }

    /// The value of the property at `index` in the kind's table, if the
    /// node has it, whatever the iterator has already yielded: the query
    /// walker's direct read of one key (record 0010 §5.3). Never allocates
    /// beyond what the leaf accessor returns.
    pub(crate) fn get(&self, index: usize) -> Option<Value<'a>> {
        if index < self.props.len() {
            self.source.get(index)
        } else {
            None
        }
    }
}

impl<'a> Iterator for Format<'a> {
    type Item = Prop<'a>;

    fn next(&mut self) -> Option<Self::Item> {
        while let Some(spec) = self.props.get(self.next) {
            let index = self.next;
            self.next += 1;
            if spec.raw || self.hidden.contains(&spec.name) {
                continue;
            }
            if let Some(value) = self.source.get(index) {
                return Some(Prop { spec, value });
            }
        }
        None
    }

    fn size_hint(&self) -> (usize, Option<usize>) {
        (0, Some(self.props.len().saturating_sub(self.next)))
    }
}

impl FusedIterator for Format<'_> {}

impl fmt::Debug for Format<'_> {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        f.debug_struct("Format")
            .field("props", &self.props.len())
            .field("next", &self.next)
            .finish_non_exhaustive()
    }
}

/// One entry of a node's `format` member: a key and its typed value.
///
/// `Display` writes the value as OfficeCLI shows it for this key: the
/// canonical text of [`crate::value`], with a length in the style of the
/// key's value type. `serde::Serialize` writes it with its JSON type.
///
/// ```
/// use litchi_ops::node::{Node, Value};
///
/// fn bold(node: &Node<'_>) -> litchi_ops::Result<Option<bool>> {
///     for prop in node.view().format()? {
///         if prop.key() == "bold" {
///             return Ok(match prop.value() {
///                 Value::Bool(bold) => Some(*bold),
///                 _ => None,
///             });
///         }
///     }
///     Ok(None)
/// }
/// # let _ = bold;
/// ```
#[derive(Clone, PartialEq)]
pub struct Prop<'a> {
    spec: &'static PropSpec,
    value: Value<'a>,
}

impl<'a> Prop<'a> {
    /// The key.
    #[must_use]
    pub fn key(&self) -> Key {
        self.spec.name
    }

    /// The typed value.
    #[must_use]
    pub fn value(&self) -> &Value<'a> {
        &self.value
    }

    /// The typed value, by value.
    #[must_use]
    pub fn into_value(self) -> Value<'a> {
        self.value
    }

    /// The key and the typed value.
    #[must_use]
    pub fn into_parts(self) -> (Key, Value<'a>) {
        (self.spec.name, self.value)
    }
}

impl fmt::Display for Prop<'_> {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        self.value.fmt_typed(Some(self.spec.ty), f)
    }
}

impl Serialize for Prop<'_> {
    fn serialize<S: Serializer>(&self, serializer: S) -> std::result::Result<S::Ok, S::Error> {
        self.value.serialize_typed(Some(self.spec.ty), serializer)
    }
}

/// The children of a node, in document order.
///
/// Each item is fallible, because a lazily decoded part can fail mid-walk;
/// after the first error the iterator is fused.
///
/// ```
/// use litchi_ops::node::{Node, View};
///
/// // The canonical paths of a view's listed children, each viewed one level
/// // down.
/// fn child_paths(view: View<'_, '_>) -> litchi_ops::Result<Vec<String>> {
///     let mut paths = Vec::new();
///     for child in view.children() {
///         let child: Node<'_> = child?;
///         paths.push(child.view_with(view.child_options()).path().to_string());
///     }
///     Ok(paths)
/// }
/// # let _ = child_paths;
/// ```
pub struct Children<'n, 'f> {
    inner: Inner<'n, 'f>,
}

#[cfg_attr(
    not(test),
    allow(dead_code, reason = "constructed by the format adapters (R4-R7)")
)]
enum Inner<'n, 'f> {
    Empty(PhantomData<(&'n (), Node<'f>)>),
    Failed(Option<Error>),
    #[cfg(any(test, feature = "test-support"))]
    Fake(super::fake::Children<'n, 'f>),
    /// Boxed: the DOCX iterator carries its parent's address.
    #[cfg(feature = "docx")]
    Docx(Box<crate::format::docx::Children<'f>>),
    /// The legacy DOC iterator carries the retained document handle.
    #[cfg(feature = "doc")]
    Doc(crate::format::doc::children::Children<'f>),
    /// The XLSX iterator carries the workbook handle through its nodes.
    #[cfg(feature = "xlsx")]
    Xlsx(crate::format::xlsx::Children<'f>),
    #[cfg(feature = "pptx")]
    Pptx(crate::format::pptx::Children<'f>),
    #[cfg(feature = "ppt")]
    Ppt(crate::format::ppt::Children<'f>),
    #[cfg(feature = "xlsb")]
    Xlsb(crate::format::xlsb::Children<'f>),
    #[cfg(feature = "xls")]
    Xls(crate::format::xls::Children<'f>),
    /// The RTF iterator carries the native RTF handle through its nodes.
    #[cfg(feature = "rtf")]
    Rtf(crate::format::rtf::children::Children<'f>),
    /// The ODS iterator borrows the handle's cached content index.
    #[cfg(feature = "ods")]
    Ods(crate::format::ods::Children<'f>),
    /// The ODT iterator carries per-kind ordinals through the body index.
    #[cfg(feature = "odt")]
    Odt(Box<crate::format::odt::children::Children<'f>>),
    /// The ODP iterator carries its parent node and scene cursor.
    #[cfg(feature = "odp")]
    Odp(Box<crate::format::odp::children::Children<'f>>),
    /// The object-tree iterator walks one item's children.
    #[cfg(odf_objects)]
    Object(crate::format::odf_tree::children::Children<'f>),
    #[cfg(feature = "iwork-adapter")]
    Iwork(crate::format::iwork::children::Children<'f>),
}

impl<'n, 'f> Children<'n, 'f> {
    /// No children.
    pub(crate) const fn empty() -> Self {
        Self {
            inner: Inner::Empty(PhantomData),
        }
    }

    /// Children that could not be listed: yields `error`, then nothing.
    #[cfg_attr(
        not(test),
        allow(dead_code, reason = "constructed by the format adapters (R4-R7)")
    )]
    pub(crate) const fn failed(error: Error) -> Self {
        Self {
            inner: Inner::Failed(Some(error)),
        }
    }

    /// The fake format's children.
    #[cfg(any(test, feature = "test-support"))]
    pub(crate) const fn fake(children: super::fake::Children<'n, 'f>) -> Self {
        Self {
            inner: Inner::Fake(children),
        }
    }

    /// A DOCX node's children.
    #[cfg(feature = "docx")]
    pub(crate) fn docx(children: crate::format::docx::Children<'f>) -> Self {
        Self {
            inner: Inner::Docx(Box::new(children)),
        }
    }

    /// A DOC node's children.
    #[cfg(feature = "doc")]
    pub(crate) fn doc(children: crate::format::doc::children::Children<'f>) -> Self {
        Self {
            inner: Inner::Doc(children),
        }
    }

    /// An XLSX node's children.
    #[cfg(feature = "xlsx")]
    pub(crate) const fn xlsx(children: crate::format::xlsx::Children<'f>) -> Self {
        Self {
            inner: Inner::Xlsx(children),
        }
    }
    #[cfg(feature = "pptx")]
    pub(crate) const fn pptx(children: crate::format::pptx::Children<'f>) -> Self {
        Self {
            inner: Inner::Pptx(children),
        }
    }
    #[cfg(feature = "ppt")]
    pub(crate) const fn ppt(children: crate::format::ppt::Children<'f>) -> Self {
        Self {
            inner: Inner::Ppt(children),
        }
    }
    #[cfg(feature = "xlsb")]
    pub(crate) const fn xlsb(children: crate::format::xlsb::Children<'f>) -> Self {
        Self {
            inner: Inner::Xlsb(children),
        }
    }
    #[cfg(feature = "xls")]
    pub(crate) const fn xls(children: crate::format::xls::Children<'f>) -> Self {
        Self {
            inner: Inner::Xls(children),
        }
    }

    /// An ODT node's children.
    #[cfg(feature = "odt")]
    pub(crate) fn odt(children: crate::format::odt::children::Children<'f>) -> Self {
        Self {
            inner: Inner::Odt(Box::new(children)),
        }
    }

    /// An ODP node's children.
    #[cfg(feature = "odp")]
    pub(crate) fn odp(children: crate::format::odp::children::Children<'f>) -> Self {
        Self {
            inner: Inner::Odp(Box::new(children)),
        }
    }

    /// An object-tree node's children.
    #[cfg(odf_objects)]
    pub(crate) const fn odf_tree(
        children: crate::format::odf_tree::children::Children<'f>,
    ) -> Self {
        Self {
            inner: Inner::Object(children),
        }
    }

    /// An iWork object-tree node's children.
    #[cfg(feature = "iwork-adapter")]
    pub(crate) const fn iwork(children: crate::format::iwork::children::Children<'f>) -> Self {
        Self {
            inner: Inner::Iwork(children),
        }
    }

    /// An RTF node's children.
    #[cfg(feature = "rtf")]
    pub(crate) const fn rtf(children: crate::format::rtf::children::Children<'f>) -> Self {
        Self {
            inner: Inner::Rtf(children),
        }
    }

    /// An ODS node's children.
    #[cfg(feature = "ods")]
    pub(crate) const fn ods(children: crate::format::ods::Children<'f>) -> Self {
        Self {
            inner: Inner::Ods(children),
        }
    }
}

impl<'f> Iterator for Children<'_, 'f> {
    type Item = Result<Node<'f>>;

    fn next(&mut self) -> Option<Self::Item> {
        let item = match &mut self.inner {
            Inner::Empty(_) => return None,
            Inner::Failed(error) => error.take().map(Err),
            #[cfg(any(test, feature = "test-support"))]
            Inner::Fake(children) => children.next(),
            #[cfg(feature = "docx")]
            Inner::Docx(children) => children.next().map(|child| child.map(Node::Docx)),
            #[cfg(feature = "doc")]
            Inner::Doc(children) => children.next().map(|child| child.map(Node::Doc)),
            #[cfg(feature = "xlsx")]
            Inner::Xlsx(children) => children.next().map(|child| child.map(Node::Xlsx)),
            #[cfg(feature = "pptx")]
            Inner::Pptx(children) => children.next().map(|child| child.map(Node::Pptx)),
            #[cfg(feature = "ppt")]
            Inner::Ppt(children) => children.next().map(|child| child.map(Node::Ppt)),
            #[cfg(feature = "xlsb")]
            Inner::Xlsb(children) => children.next().map(|child| child.map(Node::Xlsb)),
            #[cfg(feature = "xls")]
            Inner::Xls(children) => children.next().map(|child| child.map(Node::Xls)),
            #[cfg(feature = "rtf")]
            Inner::Rtf(children) => children.next().map(|child| child.map(Node::Rtf)),
            #[cfg(feature = "ods")]
            Inner::Ods(children) => children.next().map(|child| child.map(Node::Ods)),
            #[cfg(feature = "odt")]
            Inner::Odt(children) => children.next().map(|child| child.map(Node::Odt)),
            #[cfg(feature = "odp")]
            Inner::Odp(children) => children.next().map(|child| child.map(Node::Odp)),
            #[cfg(odf_objects)]
            Inner::Object(children) => children.next().map(|child| child.map(Node::Object)),
            #[cfg(feature = "iwork-adapter")]
            Inner::Iwork(children) => children.next().map(|child| child.map(Node::Iwork)),
        };
        if matches!(item, Some(Err(_)) | None) {
            self.inner = Inner::Empty(PhantomData);
        }
        item
    }
}

impl FusedIterator for Children<'_, '_> {}

impl fmt::Debug for Children<'_, '_> {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        f.debug_struct("Children").finish_non_exhaustive()
    }
}

impl fmt::Debug for Preview<'_> {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        f.debug_struct("Preview").finish_non_exhaustive()
    }
}

impl fmt::Debug for Prop<'_> {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        f.debug_struct("Prop").finish_non_exhaustive()
    }
}
