//! A read-only structural index of the text body, with lossless edits.
//!
//! [`Tree`] indexes `office:body/office:text` of `content.xml` in one pass:
//! every paragraph, heading, run, hyperlink, list, list item, table, row,
//! cell, section, frame, image, text box, note, annotation and bookmark,
//! with the exact byte range it occupies. Nothing is decoded until asked:
//! text, attributes and style properties are read from the source bytes on
//! demand, so the index costs one pass and a few words per element.
//!
//! The tree is the addressing model of the body edits staged on
//! [`crate::transaction::Edit`] (`set_text`, `set_attribute`,
//! `set_text_properties`, `set_paragraph_properties`, `insert`, `remove`,
//! `relocate`). Every edit names nodes of the tree of its source package, is
//! applied as a byte splice of the exact source, and is published through one
//! audited content-part replacement, so everything an edit does not touch
//! keeps its source bytes.
//!
//! ```no_run
//! use litchi_odt::Package;
//! use litchi_odt::body::Kind;
//!
//! # fn main() -> litchi_core::Result<()> {
//! let package = Package::open("report.odt")?;
//! let document = package.document()?;
//! let body = document.body()?;
//! for node in body.root().children() {
//!     if node.kind() == Kind::Paragraph {
//!         println!("{}", node.text()?);
//!     }
//! }
//! # Ok(())
//! # }
//! ```

mod edit;
mod fields;
mod index;
mod kept;
pub mod parts;
mod properties;
mod text;
pub(crate) mod xml;

use std::borrow::Cow;
use std::ops::Range;

use litchi_core::Result;

pub(crate) use edit::{BodyEdits, apply as apply_body_edits};
pub use edit::{Location, NewNode, Placement, admits};
pub use fields::{TextField, fields, text_fields};
pub use parts::{Parts, PartsView};
pub use properties::{
    CellProperties, ParagraphProperties, ParagraphPropertyUpdate, TabStopSpec, TextProperties,
    TextPropertyUpdate,
};
use xml::{Ns, Prefixes};

/// The largest text one body edit stages.
pub(crate) const MAX_TEXT_BYTES: usize = 1024 * 1024;
const NONE: u32 = u32::MAX;

/// The kind of an indexed body element.
#[derive(Clone, Copy, Debug, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum Kind {
    /// `office:text`, the root of the index.
    Body,
    /// `text:p`.
    Paragraph,
    /// `text:h`.
    Heading,
    /// `text:span`: a run with its own style. Nested spans belong to it.
    Span,
    /// Character data (with `text:s`, `text:tab` and `text:line-break`)
    /// directly inside a paragraph, heading or hyperlink: a run without a
    /// style of its own.
    Text,
    /// `text:a`.
    Hyperlink,
    /// `text:list`.
    List,
    /// `text:list-item` or `text:list-header`.
    ListItem,
    /// `table:table`.
    Table,
    /// `table:table-row`, wherever its row group puts it.
    Row,
    /// `table:table-cell`.
    Cell,
    /// `table:covered-table-cell`: a cell hidden by a spanning neighbour.
    CoveredCell,
    /// `text:section`.
    Section,
    /// `draw:frame`.
    Frame,
    /// `draw:image`.
    Image,
    /// `draw:text-box`.
    TextBox,
    /// `text:note` with `text:note-class="footnote"` (the default).
    Footnote,
    /// `text:note` with `text:note-class="endnote"`.
    Endnote,
    /// `office:annotation`.
    Annotation,
    /// `text:bookmark` or `text:bookmark-start`.
    Bookmark,
}

impl Kind {
    /// Whether runs (spans and text) are direct children of this kind.
    #[must_use]
    pub const fn holds_runs(self) -> bool {
        matches!(self, Self::Paragraph | Self::Heading | Self::Hyperlink)
    }

    /// Whether the kind is a run.
    #[must_use]
    pub const fn is_run(self) -> bool {
        matches!(self, Self::Span | Self::Text)
    }
}

/// The identity of a node within one [`Tree`]. Identities are dense,
/// assigned in document order, and valid only for the tree that issued
/// them.
#[derive(Clone, Copy, Debug, PartialEq, Eq, Hash, PartialOrd, Ord)]
pub struct NodeId(u32);

impl NodeId {
    /// The body node.
    pub const BODY: Self = Self(0);

    /// The zero-based position in document order.
    #[must_use]
    pub const fn index(self) -> usize {
        self.0 as usize
    }
}

#[derive(Clone, Debug)]
struct NodeData {
    kind: Kind,
    parent: u32,
    first_child: u32,
    last_child: u32,
    next_sibling: u32,
    /// Start of the start tag (or of the character data).
    start: u32,
    /// End of the start tag; `start` for character data.
    open_end: u32,
    /// Start of the end tag; `end` for empty elements and character data.
    close_start: u32,
    /// End of the end tag.
    end: u32,
}

/// The part of the package a style is declared in.
#[derive(Clone, Copy, Debug, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum StylePart {
    /// `content.xml` `office:automatic-styles`.
    Automatic,
    /// `styles.xml` `office:styles`.
    Common,
}

#[derive(Clone, Debug)]
struct StyleData {
    part: StylePart,
    start: u32,
    open_end: u32,
    end: u32,
    /// Start tags of the property children, by family of properties.
    text: Option<Range<u32>>,
    paragraph: Option<Range<u32>>,
    cell: Option<Range<u32>>,
}

/// The structural index of one document body. Built once per document by
/// [`crate::Document::body`].
#[derive(Debug)]
pub struct Tree {
    prefixes: Prefixes,
    style_prefixes: Prefixes,
    nodes: Vec<NodeData>,
    styles: Vec<StyleData>,
    /// Automatic style indexes sorted by name.
    automatic_names: Vec<(Box<str>, u32)>,
    /// Common style indexes sorted by name.
    common_names: Vec<(Box<str>, u32)>,
    /// Where new automatic styles are inserted in `content.xml`.
    automatic_insert: AutomaticInsert,
}

/// How a new automatic style reaches `content.xml`.
#[derive(Clone, Copy, Debug)]
enum AutomaticInsert {
    /// Before this offset (the `</office:automatic-styles>` end tag).
    Before(u32),
    /// Replace the empty `<office:automatic-styles/>` at this range.
    Expand(u32, u32),
    /// Insert a new container before `office:body` at this offset.
    Create(u32),
}

/// One indexed document body: the [`Tree`] with the parts it indexes.
#[derive(Clone, Copy)]
pub struct Body<'a> {
    tree: &'a Tree,
    content: &'a str,
    styles: &'a str,
}

impl std::fmt::Debug for Body<'_> {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        formatter
            .debug_struct("Body")
            .field("nodes", &self.tree.nodes.len())
            .finish_non_exhaustive()
    }
}

impl Tree {
    /// The number of automatic styles; common styles follow them.
    fn automatic_count(&self) -> u32 {
        self.styles
            .iter()
            .take_while(|style| style.part == StylePart::Automatic)
            .count() as u32
    }

    /// Index `content` (`content.xml`) and the common styles of `styles`
    /// (`styles.xml`).
    ///
    /// # Errors
    ///
    /// Returns an error when `content` is not well-formed XML, has no
    /// `office:text` body, or exceeds the index limits.
    pub fn build(content: &str, styles: Option<&str>) -> Result<Self> {
        index::build(content, styles)
    }

    /// The number of indexed nodes, the body included.
    #[must_use]
    pub fn len(&self) -> usize {
        self.nodes.len()
    }

    /// Whether the index holds only the body.
    #[must_use]
    pub fn is_empty(&self) -> bool {
        self.nodes.len() <= 1
    }

    /// A borrowed view of this tree over the parts it was built from.
    #[must_use]
    pub fn view<'a>(&'a self, content: &'a str, styles: Option<&'a str>) -> Body<'a> {
        Body {
            tree: self,
            content,
            styles: styles.unwrap_or(""),
        }
    }
}

impl<'a> Body<'a> {
    /// The body node (`office:text`).
    #[must_use]
    pub fn root(self) -> Node<'a> {
        Node {
            body: self,
            id: NodeId::BODY,
        }
    }

    /// The node with `id`, if this tree issued it.
    #[must_use]
    pub fn node(self, id: NodeId) -> Option<Node<'a>> {
        (id.index() < self.tree.nodes.len()).then_some(Node { body: self, id })
    }

    /// Every node in document order, the body first.
    pub fn nodes(self) -> impl Iterator<Item = Node<'a>> + 'a {
        (0..self.tree.nodes.len()).map(move |index| Node {
            body: self,
            id: NodeId(index as u32),
        })
    }

    /// The number of indexed nodes, the body included.
    #[must_use]
    pub fn len(self) -> usize {
        self.tree.nodes.len()
    }

    /// Whether the index holds only the body.
    #[must_use]
    pub fn is_empty(self) -> bool {
        self.tree.is_empty()
    }

    /// The automatic style `name` (from `content.xml`).
    #[must_use]
    pub fn automatic_style(self, name: &str) -> Option<Style<'a>> {
        lookup(&self.tree.automatic_names, name).map(|index| Style { body: self, index })
    }

    /// The common style `name` (from `styles.xml`).
    #[must_use]
    pub fn common_style(self, name: &str) -> Option<Style<'a>> {
        lookup(&self.tree.common_names, name).map(|index| Style { body: self, index })
    }

    /// The common styles of `styles.xml`, in declaration order.
    #[must_use]
    pub fn common_styles(self) -> CommonStyles<'a> {
        CommonStyles {
            body: self,
            next: self.tree.automatic_count(),
        }
    }

    pub(crate) fn tree(self) -> &'a Tree {
        self.tree
    }

    pub(crate) fn content(self) -> &'a str {
        self.content
    }

    pub(crate) fn prefixes(self) -> &'a Prefixes {
        &self.tree.prefixes
    }

    fn data(self, id: NodeId) -> &'a NodeData {
        &self.tree.nodes[id.index()]
    }
}

fn lookup(names: &[(Box<str>, u32)], name: &str) -> Option<u32> {
    names
        .binary_search_by(|(held, _)| (**held).cmp(name))
        .ok()
        .map(|index| names[index].1)
}

/// A typed attribute of a body element. Each resolves to the attribute
/// OpenDocument defines for the element's kind.
#[derive(Clone, Copy, Debug, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum Attribute {
    /// The style: `text:style-name`, `table:style-name` or
    /// `draw:style-name`, by kind.
    StyleName,
    /// The name: `text:name`, `table:name` or `draw:name`, by kind.
    Name,
    /// `text:outline-level` of a heading.
    OutlineLevel,
    /// `xlink:href` of a hyperlink or image.
    Href,
    /// `office:title` of a hyperlink.
    Title,
    /// `office:target-frame-name` of a hyperlink.
    TargetFrame,
    /// `text:note-class` of a note.
    NoteClass,
    /// `table:number-columns-spanned` of a cell.
    ColumnsSpanned,
    /// `table:number-rows-spanned` of a cell.
    RowsSpanned,
    /// `table:number-columns-repeated` of a cell or covered cell: how many
    /// identical columns it stands for.
    ColumnsRepeated,
    /// `office:value-type` of a cell.
    ValueType,
    /// `svg:width` of a frame.
    Width,
    /// `svg:height` of a frame.
    Height,
    /// `svg:x` of a frame.
    X,
    /// `svg:y` of a frame.
    Y,
    /// `text:anchor-type` of a frame.
    AnchorType,
    /// `text:continue-numbering` of a list.
    ContinueNumbering,
    /// `text:is-list-header` of a heading.
    IsListHeader,
    /// `text:protected` of a section.
    Protected,
    /// `text:display` of a section (`none` hides it).
    Display,
}

impl Attribute {
    /// The namespace and local name of this attribute on `kind`, or `None`
    /// when the kind does not carry it.
    pub(crate) const fn qualified(self, kind: Kind) -> Option<(Ns, &'static str)> {
        use Kind as K;
        Some(match (self, kind) {
            (Self::StyleName, K::Table | K::Row | K::Cell | K::CoveredCell) => {
                (Ns::Table, "style-name")
            },
            (Self::StyleName, K::Frame) => (Ns::Draw, "style-name"),
            (
                Self::StyleName,
                K::Paragraph | K::Heading | K::Span | K::Hyperlink | K::List | K::Section,
            ) => (Ns::Text, "style-name"),
            (Self::Name, K::Table) => (Ns::Table, "name"),
            (Self::Name, K::Frame) => (Ns::Draw, "name"),
            (Self::Name, K::Section | K::Bookmark) => (Ns::Text, "name"),
            (Self::Name, K::Annotation) => (Ns::Office, "name"),
            (Self::OutlineLevel, K::Heading) => (Ns::Text, "outline-level"),
            (Self::Href, K::Hyperlink | K::Image) => (Ns::Xlink, "href"),
            (Self::Title, K::Hyperlink) => (Ns::Office, "title"),
            (Self::TargetFrame, K::Hyperlink) => (Ns::Office, "target-frame-name"),
            (Self::NoteClass, K::Footnote | K::Endnote) => (Ns::Text, "note-class"),
            (Self::ColumnsSpanned, K::Cell) => (Ns::Table, "number-columns-spanned"),
            (Self::RowsSpanned, K::Cell) => (Ns::Table, "number-rows-spanned"),
            (Self::ColumnsRepeated, K::Cell | K::CoveredCell) => {
                (Ns::Table, "number-columns-repeated")
            },
            (Self::ValueType, K::Cell) => (Ns::Office, "value-type"),
            (Self::Width, K::Frame) => (Ns::Svg, "width"),
            (Self::Height, K::Frame) => (Ns::Svg, "height"),
            (Self::X, K::Frame) => (Ns::Svg, "x"),
            (Self::Y, K::Frame) => (Ns::Svg, "y"),
            (Self::AnchorType, K::Frame) => (Ns::Text, "anchor-type"),
            (Self::ContinueNumbering, K::List) => (Ns::Text, "continue-numbering"),
            (Self::IsListHeader, K::Heading) => (Ns::Text, "is-list-header"),
            (Self::Protected, K::Section) => (Ns::Text, "protected"),
            (Self::Display, K::Section) => (Ns::Text, "display"),
            _ => return None,
        })
    }
}

/// One indexed node, borrowed from its [`Body`].
#[derive(Clone, Copy)]
pub struct Node<'a> {
    body: Body<'a>,
    id: NodeId,
}

impl std::fmt::Debug for Node<'_> {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        formatter
            .debug_struct("Node")
            .field("id", &self.id)
            .field("kind", &self.kind())
            .finish()
    }
}

impl<'a> Node<'a> {
    /// The node's identity in its tree.
    #[must_use]
    pub const fn id(self) -> NodeId {
        self.id
    }

    /// The node's kind.
    #[must_use]
    pub fn kind(self) -> Kind {
        self.body.data(self.id).kind
    }

    /// The body this node belongs to.
    #[must_use]
    pub const fn body(self) -> Body<'a> {
        self.body
    }

    /// The nearest indexed ancestor; `None` for the body.
    #[must_use]
    pub fn parent(self) -> Option<Self> {
        let parent = self.body.data(self.id).parent;
        (parent != NONE).then_some(Self {
            body: self.body,
            id: NodeId(parent),
        })
    }

    /// The indexed children, in document order.
    #[must_use]
    pub fn children(self) -> Children<'a> {
        Children {
            body: self.body,
            next: self.body.data(self.id).first_child,
        }
    }

    /// Whether another indexed node starts inside this node's bytes: its
    /// children, or nodes the index attached to an ancestor (notes or
    /// bookmarks inside a run).
    #[must_use]
    pub fn encloses_indexed(self) -> bool {
        let end = self.body.data(self.id).end;
        self.id
            .0
            .checked_add(1)
            .and_then(|next| self.body.tree.nodes.get(next as usize))
            .is_some_and(|next| next.start < end)
    }

    /// The next indexed sibling.
    #[must_use]
    pub fn next_sibling(self) -> Option<Self> {
        let next = self.body.data(self.id).next_sibling;
        (next != NONE).then_some(Self {
            body: self.body,
            id: NodeId(next),
        })
    }

    /// The byte range of the element (or character data) in `content.xml`.
    #[must_use]
    pub fn range(self) -> Range<usize> {
        let data = self.body.data(self.id);
        data.start as usize..data.end as usize
    }

    /// The byte range of the element's content, between its tags; empty for
    /// an empty element.
    #[must_use]
    pub fn content_range(self) -> Range<usize> {
        let data = self.body.data(self.id);
        if data.kind == Kind::Text {
            return data.start as usize..data.end as usize;
        }
        data.open_end as usize..data.close_start.max(data.open_end) as usize
    }

    /// The source markup of the element.
    #[must_use]
    pub fn xml(self) -> &'a str {
        &self.body.content[self.range()]
    }

    pub(crate) fn start_tag(self) -> &'a str {
        let data = self.body.data(self.id);
        &self.body.content[data.start as usize..data.open_end as usize]
    }

    pub(crate) fn is_empty_element(self) -> bool {
        let data = self.body.data(self.id);
        data.kind != Kind::Text && data.close_start == data.end
    }

    /// The decoded value of `attribute`, when the element carries it.
    #[must_use]
    pub fn attribute(self, attribute: Attribute) -> Option<Cow<'a, str>> {
        let (ns, local) = attribute.qualified(self.kind())?;
        self.raw_attribute(ns, local)
    }

    pub(crate) fn raw_attribute(self, ns: Ns, local: &str) -> Option<Cow<'a, str>> {
        if self.kind() == Kind::Text {
            return None;
        }
        let prefixes = self.body.prefixes();
        xml::attributes(self.start_tag()).find_map(|(name, value)| {
            prefixes
                .is(name.as_bytes(), ns, local)
                .then(|| xml::unescape(value))
        })
    }

    /// The node's visible text: paragraphs of a container are separated by
    /// line feeds, `text:s`, `text:tab` and `text:line-break` are expanded,
    /// and notes, annotations, frames and tracked deletions inside the node
    /// are left out.
    ///
    /// # Errors
    ///
    /// Returns an error when the node's markup cannot be read.
    pub fn text(self) -> Result<Cow<'a, str>> {
        text::text(self)
    }

    /// The style the node names, if any.
    #[must_use]
    pub fn style_name(self) -> Option<Cow<'a, str>> {
        self.attribute(Attribute::StyleName)
    }

    /// The automatic style the node names, if its style is automatic.
    #[must_use]
    pub fn automatic_style(self) -> Option<Style<'a>> {
        self.body.automatic_style(&self.style_name()?)
    }

    /// The common style the node is based on: the named style itself when it
    /// is common, or the parent of its automatic style.
    #[must_use]
    pub fn common_style_name(self) -> Option<Cow<'a, str>> {
        let name = self.style_name()?;
        match self.body.automatic_style(&name) {
            Some(style) => style.parent_name(),
            None => Some(name),
        }
    }

    /// The character properties the node's automatic style sets locally.
    #[must_use]
    pub fn text_properties(self) -> TextProperties<'a> {
        self.automatic_style()
            .map(Style::text_properties)
            .unwrap_or_default()
    }

    /// The paragraph properties the node's automatic style sets locally.
    #[must_use]
    pub fn paragraph_properties(self) -> ParagraphProperties<'a> {
        self.automatic_style()
            .map(Style::paragraph_properties)
            .unwrap_or_default()
    }

    /// The cell properties the node's automatic style sets locally.
    #[must_use]
    pub fn cell_properties(self) -> CellProperties<'a> {
        self.automatic_style()
            .map(Style::cell_properties)
            .unwrap_or_default()
    }

    /// The decoded text of the first direct child element `ns:local`, such
    /// as an annotation's `dc:creator`.
    fn child_element_text(self, ns: Ns, local: &str) -> Option<Cow<'a, str>> {
        text::child_element_text(self, ns, local)
    }

    /// An annotation's author (`dc:creator`).
    #[must_use]
    pub fn author(self) -> Option<Cow<'a, str>> {
        (self.kind() == Kind::Annotation)
            .then(|| self.child_element_text(Ns::Dc, "creator"))
            .flatten()
    }

    /// An annotation's date (`dc:date`).
    #[must_use]
    pub fn date(self) -> Option<Cow<'a, str>> {
        (self.kind() == Kind::Annotation)
            .then(|| self.child_element_text(Ns::Dc, "date"))
            .flatten()
    }

    /// A note's citation (`text:note-citation`).
    #[must_use]
    pub fn citation(self) -> Option<Cow<'a, str>> {
        matches!(self.kind(), Kind::Footnote | Kind::Endnote)
            .then(|| self.child_element_text(Ns::Text, "note-citation"))
            .flatten()
    }
}

/// The indexed children of a node.
#[derive(Clone)]
pub struct Children<'a> {
    body: Body<'a>,
    next: u32,
}

impl<'a> Iterator for Children<'a> {
    type Item = Node<'a>;

    fn next(&mut self) -> Option<Self::Item> {
        if self.next == NONE {
            return None;
        }
        let id = NodeId(self.next);
        self.next = self.body.data(id).next_sibling;
        Some(Node {
            body: self.body,
            id,
        })
    }
}

impl std::iter::FusedIterator for Children<'_> {}

/// The common styles of one body, in declaration order.
#[derive(Clone)]
pub struct CommonStyles<'a> {
    body: Body<'a>,
    next: u32,
}

impl<'a> Iterator for CommonStyles<'a> {
    type Item = Style<'a>;

    fn next(&mut self) -> Option<Self::Item> {
        let index = self.next;
        if index as usize >= self.body.tree.styles.len() {
            return None;
        }
        self.next += 1;
        Some(Style {
            body: self.body,
            index,
        })
    }
}

impl std::iter::FusedIterator for CommonStyles<'_> {}

/// One style declaration.
#[derive(Clone, Copy)]
pub struct Style<'a> {
    body: Body<'a>,
    index: u32,
}

impl std::fmt::Debug for Style<'_> {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        formatter
            .debug_struct("Style")
            .field("name", &self.name())
            .finish_non_exhaustive()
    }
}

impl<'a> Style<'a> {
    fn data(self) -> &'a StyleData {
        &self.body.tree.styles[self.index as usize]
    }

    fn source(self) -> (&'a str, &'a Prefixes) {
        match self.data().part {
            StylePart::Automatic => (self.body.content, &self.body.tree.prefixes),
            StylePart::Common => (self.body.styles, &self.body.tree.style_prefixes),
        }
    }

    fn start_tag(self) -> &'a str {
        let (xml, _) = self.source();
        let data = self.data();
        &xml[data.start as usize..data.open_end as usize]
    }

    fn attribute(self, ns: Ns, local: &str) -> Option<Cow<'a, str>> {
        let (_, prefixes) = self.source();
        xml::attributes(self.start_tag()).find_map(|(name, value)| {
            prefixes
                .is(name.as_bytes(), ns, local)
                .then(|| xml::unescape(value))
        })
    }

    /// Where the style is declared.
    #[must_use]
    pub fn part(self) -> StylePart {
        self.data().part
    }

    /// `style:name`.
    #[must_use]
    pub fn name(self) -> Option<Cow<'a, str>> {
        self.attribute(Ns::Style, "name")
    }

    /// `style:display-name`, falling back to the name.
    #[must_use]
    pub fn display_name(self) -> Option<Cow<'a, str>> {
        self.attribute(Ns::Style, "display-name")
            .or_else(|| self.name())
    }

    /// `style:family`.
    #[must_use]
    pub fn family(self) -> Option<Cow<'a, str>> {
        self.attribute(Ns::Style, "family")
    }

    /// `style:parent-style-name`.
    #[must_use]
    pub fn parent_name(self) -> Option<Cow<'a, str>> {
        self.attribute(Ns::Style, "parent-style-name")
    }

    /// The source markup of the declaration.
    pub(crate) fn xml(self) -> &'a str {
        let (xml, _) = self.source();
        let data = self.data();
        &xml[data.start as usize..data.end as usize]
    }

    fn property_tag(self, range: Option<&Range<u32>>) -> Option<(&'a str, &'a Prefixes)> {
        let (xml, prefixes) = self.source();
        range.map(|range| (&xml[range.start as usize..range.end as usize], prefixes))
    }

    /// The `style:text-properties` the declaration sets.
    #[must_use]
    pub fn text_properties(self) -> TextProperties<'a> {
        self.property_tag(self.data().text.as_ref())
            .map(|(tag, prefixes)| TextProperties::read(tag, prefixes))
            .unwrap_or_default()
    }

    /// The `style:paragraph-properties` the declaration sets.
    #[must_use]
    pub fn paragraph_properties(self) -> ParagraphProperties<'a> {
        self.property_tag(self.data().paragraph.as_ref())
            .map(|(tag, prefixes)| ParagraphProperties::read(tag, prefixes))
            .unwrap_or_default()
    }

    /// The `style:table-cell-properties` the declaration sets.
    #[must_use]
    pub fn cell_properties(self) -> CellProperties<'a> {
        self.property_tag(self.data().cell.as_ref())
            .map(|(tag, prefixes)| CellProperties::read(tag, prefixes))
            .unwrap_or_default()
    }
}

#[cfg(test)]
mod tests;
