//! Lossless body edits: typed changes to nodes of one source [`super::Tree`],
//! compiled into non-overlapping byte splices of the exact source parts.
//!
//! Every staged edit names nodes of the tree of the edit's source package.
//! Staging validates the edit against that tree and refuses an edit whose
//! bytes overlap another node's pending change (a node inside a replaced or
//! removed element), so a later failure never depends on edit order. Edits
//! to one node merge; insertions at one offset keep sequential semantics.
//! Compilation derives the automatic styles formatting edits need, splices
//! `content.xml` (and `meta.xml` for metadata), and publishes one package
//! through audited source-provenance splices: every byte outside the
//! changed ranges keeps its source spelling.

use std::collections::{BTreeMap, HashSet};
use std::ops::Range;

use litchi_core::{Error, Result};
use quick_xml::Reader;
use quick_xml::events::Event;

use super::properties::{Metadata, ParagraphPropertyEdit, TextPropertyEdit, check_length};
use super::xml::{self, Ns, Prefixes};
use super::{Attribute, AutomaticInsert, Body, Kind, Node, NodeId, Style};
use crate::Document;
use crate::core::{AuthoredXmlFragment, XmlSourcePart, XmlSplicePublication};

/// The largest number of cells one inserted table may create.
const MAX_GENERATED_CELLS: usize = 16_384;
/// The largest package one body publication writes.
const MAX_PACKAGE_BYTES: usize = 64 * 1024 * 1024;

/// Where an inserted node goes among its parent's content.
#[derive(Clone, Copy, Debug, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum Placement {
    /// Before the parent's first child.
    Start,
    /// After the parent's last child.
    End,
    /// Immediately before this child of the parent.
    Before(NodeId),
    /// Immediately after this child of the parent.
    After(NodeId),
}

/// A parent and a placement within it.
#[derive(Clone, Copy, Debug, PartialEq, Eq, Hash)]
pub struct Location {
    /// The node that receives the new child.
    pub parent: NodeId,
    /// Where among its children.
    pub placement: Placement,
}

impl Location {
    /// `placement` within `parent`.
    #[must_use]
    pub const fn new(parent: NodeId, placement: Placement) -> Self {
        Self { parent, placement }
    }
}

/// A node an insertion creates.
#[derive(Clone, Debug, PartialEq, Eq)]
#[non_exhaustive]
pub enum NewNode {
    /// A paragraph (`text:p`) with plain text and an optional style.
    Paragraph {
        /// The text; spaces, tabs and line feeds are kept.
        text: String,
        /// An existing paragraph style.
        style: Option<String>,
    },
    /// A heading (`text:h`) of outline level 1–10.
    Heading {
        /// The text.
        text: String,
        /// The outline level.
        level: u8,
        /// An existing paragraph style.
        style: Option<String>,
    },
    /// A run (`text:span`) with optional character formatting.
    Span {
        /// The text.
        text: String,
        /// Formatting for a new automatic text style.
        properties: TextPropertyEdit,
    },
    /// An inert hyperlink (`text:a`).
    Hyperlink {
        /// The target IRI.
        href: String,
        /// The text.
        text: String,
    },
    /// A point bookmark (`text:bookmark`).
    Bookmark {
        /// The bookmark name.
        name: String,
    },
    /// A table of empty cells.
    Table {
        /// Rows, at least one.
        rows: usize,
        /// Columns, at least one.
        columns: usize,
    },
    /// An empty row with as many cells as the table's first row spans.
    Row,
    /// A list of plain items.
    List {
        /// One paragraph per item.
        items: Vec<String>,
    },
    /// A list item with one paragraph.
    ListItem {
        /// The text.
        text: String,
    },
    /// A section holding one empty paragraph.
    Section {
        /// The section name.
        name: String,
    },
    /// A footnote with one paragraph.
    Footnote {
        /// The note text.
        text: String,
    },
    /// An endnote with one paragraph.
    Endnote {
        /// The note text.
        text: String,
    },
    /// An annotation (comment) with one paragraph.
    Annotation {
        /// The comment text.
        text: String,
        /// The author (`dc:creator`).
        author: Option<String>,
    },
}

impl NewNode {
    /// The kind of node the insertion creates.
    #[must_use]
    pub const fn kind(&self) -> Kind {
        match self {
            Self::Paragraph { .. } => Kind::Paragraph,
            Self::Heading { .. } => Kind::Heading,
            Self::Span { .. } => Kind::Span,
            Self::Hyperlink { .. } => Kind::Hyperlink,
            Self::Bookmark { .. } => Kind::Bookmark,
            Self::Table { .. } => Kind::Table,
            Self::Row => Kind::Row,
            Self::List { .. } => Kind::List,
            Self::ListItem { .. } => Kind::ListItem,
            Self::Section { .. } => Kind::Section,
            Self::Footnote { .. } => Kind::Footnote,
            Self::Endnote { .. } => Kind::Endnote,
            Self::Annotation { .. } => Kind::Annotation,
        }
    }
}

/// Whether a node of kind `child` may be inserted into, or moved into, a
/// node of kind `parent`.
#[must_use]
pub const fn admits(parent: Kind, child: Kind) -> bool {
    use Kind as K;
    match child {
        K::Paragraph => matches!(
            parent,
            K::Body
                | K::Section
                | K::Cell
                | K::TextBox
                | K::Footnote
                | K::Endnote
                | K::ListItem
                | K::Annotation
        ),
        K::Heading | K::List => matches!(
            parent,
            K::Body | K::Section | K::Cell | K::TextBox | K::Footnote | K::Endnote | K::ListItem
        ),
        K::Table | K::Section => matches!(
            parent,
            K::Body | K::Section | K::Cell | K::TextBox | K::Footnote | K::Endnote
        ),
        K::Span | K::Text | K::Bookmark => {
            matches!(parent, K::Paragraph | K::Heading | K::Hyperlink)
        },
        K::Hyperlink | K::Footnote | K::Endnote | K::Annotation | K::Frame => {
            matches!(parent, K::Paragraph | K::Heading)
        },
        K::Row => matches!(parent, K::Table),
        K::ListItem => matches!(parent, K::List),
        _ => false,
    }
}

#[derive(Clone, Debug, Default)]
struct NodeEdit {
    attributes: Vec<(Ns, &'static str, Option<String>)>,
    text: Option<String>,
    text_properties: Option<TextPropertyEdit>,
    paragraph_properties: Option<ParagraphPropertyEdit>,
    /// Raw replacement of the whole element (move source, swap).
    replacement: Option<Replacement>,
}

impl NodeEdit {
    fn replaces_content(&self) -> bool {
        self.text.is_some() || self.replacement.is_some()
    }
}

#[derive(Clone, Debug)]
enum Replacement {
    Remove,
    Bytes(Range<usize>),
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
enum Claim {
    /// The start tag only.
    StartTag,
    /// The whole element, content included.
    Whole,
}

#[derive(Clone, Debug)]
enum Fragment {
    Markup(String),
    Copy(Range<usize>),
}

/// Body edits staged against one source tree.
#[derive(Debug, Default)]
pub(crate) struct BodyEdits {
    nodes: BTreeMap<u32, NodeEdit>,
    claims: Vec<(u32, Range<usize>, Claim)>,
    /// Inserted fragments per offset, in final order.
    inserts: BTreeMap<usize, Vec<Fragment>>,
    /// Content appended to empty parent elements, per parent.
    expansions: BTreeMap<u32, Vec<Fragment>>,
    /// End markers (`text:bookmark-end`, `office:annotation-end`) removed
    /// with their start.
    marker_removals: Vec<Range<usize>>,
    metadata: BTreeMap<u8, (Metadata, Option<String>)>,
    styles: StyleAllocator,
    generated: Generated,
}

/// Names and counters for generated tables, sections and notes.
#[derive(Debug, Default)]
struct Generated {
    tables: usize,
    notes: usize,
    footnotes: usize,
    endnotes: usize,
}

fn invalid(message: impl Into<String>) -> Error {
    Error::InvalidFormat(message.into())
}

fn node(body: Body<'_>, id: NodeId) -> Result<Node<'_>> {
    body.node(id)
        .ok_or_else(|| invalid(format!("ODT body node {} does not exist", id.index())))
}

fn start_tag_range(node: Node<'_>) -> Range<usize> {
    let start = node.range().start;
    start..start + node.start_tag().len()
}

fn overlaps(left: &Range<usize>, right: &Range<usize>) -> bool {
    left.start < right.end && right.start < left.end
}

impl BodyEdits {
    fn check_claim(&self, body: Body<'_>, id: NodeId, range: &Range<usize>) -> Result<()> {
        if self
            .nodes
            .get(&id.0)
            .is_some_and(|edit| edit.replacement.is_some())
        {
            return Err(invalid(
                "an ODT body node that is removed, moved or swapped cannot be edited again in the same transaction",
            ));
        }
        for (other, held, _) in &self.claims {
            if *other != id.0 && overlaps(range, held) {
                return Err(invalid(
                    "ODT body edits overlap: one edited node lies inside another node this transaction replaces",
                ));
            }
        }
        for offset in self.inserts.keys() {
            if range.start < *offset && *offset < range.end {
                return Err(invalid(
                    "ODT body edits overlap: an insertion lies inside a node this transaction replaces",
                ));
            }
        }
        for parent in self.expansions.keys() {
            if *parent != id.0
                && body
                    .node(NodeId(*parent))
                    .is_some_and(|parent| overlaps(&parent.range(), range))
            {
                return Err(invalid(
                    "ODT body edits overlap: an insertion grows an element this transaction replaces",
                ));
            }
        }
        for marker in &self.marker_removals {
            if overlaps(marker, range) {
                return Err(invalid(
                    "ODT body edits overlap: a removed bookmark or comment end lies inside an edited node",
                ));
            }
        }
        Ok(())
    }

    fn claim(
        &mut self,
        body: Body<'_>,
        id: NodeId,
        range: Range<usize>,
        claim: Claim,
    ) -> Result<()> {
        self.check_claim(body, id, &range)?;
        self.claims.push((id.0, range, claim));
        Ok(())
    }

    fn edit_mut(&mut self, id: NodeId) -> &mut NodeEdit {
        self.nodes.entry(id.0).or_default()
    }

    /// Replace the text of a paragraph, heading, run, hyperlink, cell, list
    /// item, text box, note or annotation.
    pub(crate) fn set_text(&mut self, body: Body<'_>, id: NodeId, text: String) -> Result<()> {
        let target = node(body, id)?;
        xml::check_text(&text, "text")?;
        match target.kind() {
            Kind::Paragraph
            | Kind::Heading
            | Kind::Span
            | Kind::Text
            | Kind::Hyperlink
            | Kind::Cell
            | Kind::ListItem
            | Kind::TextBox
            | Kind::Footnote
            | Kind::Endnote
            | Kind::Annotation => {},
            other => {
                return Err(invalid(format!(
                    "ODT {other:?} nodes have no settable text"
                )));
            },
        }
        if self.expansions.contains_key(&id.0) {
            return Err(invalid(
                "ODT body edits overlap: a node that received an insertion cannot have its text replaced",
            ));
        }
        self.claim(body, id, target.range(), Claim::Whole)?;
        self.edit_mut(id).text = Some(text);
        Ok(())
    }

    /// Set or remove one typed attribute.
    pub(crate) fn set_attribute(
        &mut self,
        body: Body<'_>,
        id: NodeId,
        attribute: Attribute,
        value: Option<String>,
    ) -> Result<()> {
        let target = node(body, id)?;
        let kind = target.kind();
        let (ns, local) = attribute.qualified(kind).ok_or_else(|| {
            invalid(format!(
                "ODT {kind:?} nodes have no {attribute:?} attribute"
            ))
        })?;
        validate_attribute(body, attribute, value.as_deref())?;
        self.claim(body, id, start_tag_range(target), Claim::StartTag)?;
        let edit = self.edit_mut(id);
        edit.attributes
            .retain(|(held_ns, held_local, _)| !(*held_ns == ns && *held_local == local));
        edit.attributes.push((ns, local, value));
        Ok(())
    }

    /// Set character formatting through a derived automatic style.
    pub(crate) fn set_text_properties(
        &mut self,
        body: Body<'_>,
        id: NodeId,
        properties: &TextPropertyEdit,
    ) -> Result<()> {
        let target = node(body, id)?;
        properties.validate()?;
        let range = match target.kind() {
            Kind::Span | Kind::Paragraph | Kind::Heading => start_tag_range(target),
            Kind::Text => target.range(),
            other => {
                return Err(invalid(format!(
                    "ODT {other:?} nodes have no character formatting"
                )));
            },
        };
        let claim = if target.kind() == Kind::Text {
            Claim::Whole
        } else {
            Claim::StartTag
        };
        self.claim(body, id, range, claim)?;
        self.edit_mut(id)
            .text_properties
            .get_or_insert_with(TextPropertyEdit::default)
            .merge(properties);
        Ok(())
    }

    /// Set paragraph formatting through a derived automatic style.
    pub(crate) fn set_paragraph_properties(
        &mut self,
        body: Body<'_>,
        id: NodeId,
        properties: &ParagraphPropertyEdit,
    ) -> Result<()> {
        let target = node(body, id)?;
        properties.validate()?;
        if !matches!(target.kind(), Kind::Paragraph | Kind::Heading) {
            return Err(invalid(format!(
                "ODT {:?} nodes have no paragraph formatting",
                target.kind()
            )));
        }
        self.claim(body, id, start_tag_range(target), Claim::StartTag)?;
        self.edit_mut(id)
            .paragraph_properties
            .get_or_insert_with(ParagraphPropertyEdit::default)
            .merge(properties);
        Ok(())
    }

    /// Remove one node and everything inside it.
    pub(crate) fn remove(&mut self, body: Body<'_>, id: NodeId) -> Result<()> {
        let target = node(body, id)?;
        check_detachable(target)?;
        self.detach(body, target)?;
        self.edit_mut(id).replacement = Some(Replacement::Remove);
        Ok(())
    }

    fn detach(&mut self, body: Body<'_>, target: Node<'_>) -> Result<()> {
        let id = target.id();
        if self.nodes.contains_key(&id.0) || self.expansions.contains_key(&id.0) {
            return Err(invalid(
                "an ODT body node edited in this transaction cannot also be removed, moved or swapped",
            ));
        }
        self.claim(body, id, target.range(), Claim::Whole)?;
        // A range bookmark or annotation loses its end marker with it.
        if let Some(end) = end_marker(target) {
            for (_, held, _) in &self.claims {
                if overlaps(held, &end) {
                    return Err(invalid(
                        "ODT body edits overlap: the end marker of a removed node lies inside another edited node",
                    ));
                }
            }
            self.marker_removals.push(end);
        }
        Ok(())
    }

    /// Insert a new node.
    pub(crate) fn insert(
        &mut self,
        body: Body<'_>,
        location: Location,
        new: &NewNode,
    ) -> Result<()> {
        let parent = node(body, location.parent)?;
        if !admits(parent.kind(), new.kind()) {
            return Err(invalid(format!(
                "ODT {:?} nodes cannot hold a new {:?}",
                parent.kind(),
                new.kind()
            )));
        }
        let markup = self.render(body, parent, new)?;
        self.place(body, parent, location.placement, Fragment::Markup(markup))?;
        Ok(())
    }

    /// Move one node to `location`, keeping its bytes.
    pub(crate) fn relocate(
        &mut self,
        body: Body<'_>,
        id: NodeId,
        location: Location,
    ) -> Result<()> {
        let target = node(body, id)?;
        check_detachable(target)?;
        let parent = node(body, location.parent)?;
        let same_parent = target.parent().map(Node::id) == Some(parent.id());
        if !same_parent && !admits(parent.kind(), target.kind()) {
            return Err(invalid(format!(
                "ODT {:?} nodes cannot hold a moved {:?}",
                parent.kind(),
                target.kind()
            )));
        }
        let range = target.range();
        if overlaps(&parent.range(), &range) && parent.range().start >= range.start {
            return Err(invalid("an ODT body node cannot move into itself"));
        }
        if let Placement::Before(anchor) | Placement::After(anchor) = location.placement {
            if anchor == id {
                return Ok(());
            }
        }
        if self.nodes.contains_key(&id.0) || self.expansions.contains_key(&id.0) {
            return Err(invalid(
                "an ODT body node edited in this transaction cannot also be moved",
            ));
        }
        self.check_claim(body, id, &range)?;
        self.claims.push((id.0, range.clone(), Claim::Whole));
        if let Err(error) = self.place(body, parent, location.placement, Fragment::Copy(range)) {
            self.claims.pop();
            return Err(error);
        }
        self.edit_mut(id).replacement = Some(Replacement::Remove);
        Ok(())
    }

    /// Exchange two nodes' bytes.
    pub(crate) fn swap(&mut self, body: Body<'_>, left: NodeId, right: NodeId) -> Result<()> {
        if left == right {
            return Ok(());
        }
        let a = node(body, left)?;
        let b = node(body, right)?;
        check_detachable(a)?;
        check_detachable(b)?;
        let (ra, rb) = (a.range(), b.range());
        if overlaps(&ra, &rb) {
            return Err(invalid(
                "one ODT body node cannot swap with a node it contains",
            ));
        }
        let fits =
            |node: Node<'_>, kind: Kind| node.parent().is_some_and(|p| admits(p.kind(), kind));
        if a.parent().map(Node::id) != b.parent().map(Node::id)
            && !(fits(a, b.kind()) && fits(b, a.kind()))
        {
            return Err(invalid(format!(
                "ODT {:?} and {:?} nodes cannot exchange places",
                a.kind(),
                b.kind()
            )));
        }
        for id in [left, right] {
            if self.nodes.contains_key(&id.0) || self.expansions.contains_key(&id.0) {
                return Err(invalid(
                    "an ODT body node edited in this transaction cannot also be swapped",
                ));
            }
        }
        self.check_claim(body, left, &ra)?;
        self.check_claim(body, right, &rb)?;
        self.claims.push((left.0, ra.clone(), Claim::Whole));
        self.claims.push((right.0, rb.clone(), Claim::Whole));
        self.edit_mut(left).replacement = Some(Replacement::Bytes(rb));
        self.edit_mut(right).replacement = Some(Replacement::Bytes(ra));
        Ok(())
    }

    /// Set or remove one `meta.xml` property.
    pub(crate) fn set_metadata(&mut self, field: Metadata, value: Option<String>) -> Result<()> {
        if let Some(value) = &value {
            xml::check_text(value, "metadata value")?;
        }
        self.metadata.insert(field as u8, (field, value));
        Ok(())
    }

    fn place(
        &mut self,
        body: Body<'_>,
        parent: Node<'_>,
        placement: Placement,
        fragment: Fragment,
    ) -> Result<()> {
        let (offset, prepend) = match placement {
            Placement::Start => (first_offset(parent), true),
            Placement::End => (last_offset(parent), false),
            Placement::Before(anchor) => {
                (Some(child_of(body, parent, anchor)?.range().start), false)
            },
            Placement::After(anchor) => (Some(child_of(body, parent, anchor)?.range().end), true),
        };
        if self
            .nodes
            .get(&parent.id().0)
            .is_some_and(NodeEdit::replaces_content)
        {
            return Err(invalid(
                "ODT body edits overlap: an insertion targets a node this transaction replaces",
            ));
        }
        let Some(offset) = offset else {
            // An empty element grows content of its own.
            for (other, held, claim) in &self.claims {
                if *other != parent.id().0
                    && *claim == Claim::Whole
                    && overlaps(held, &parent.range())
                {
                    return Err(invalid(
                        "ODT body edits overlap: an insertion targets a node inside another replaced node",
                    ));
                }
            }
            let list = self.expansions.entry(parent.id().0).or_default();
            if prepend {
                list.insert(0, fragment);
            } else {
                list.push(fragment);
            }
            return Ok(());
        };
        for (_, held, _) in &self.claims {
            if held.start < offset && offset < held.end {
                return Err(invalid(
                    "ODT body edits overlap: an insertion lies inside a node this transaction replaces",
                ));
            }
        }
        for marker in &self.marker_removals {
            if marker.start < offset && offset < marker.end {
                return Err(invalid("ODT body insertion splits a removed marker"));
            }
        }
        let list = self.inserts.entry(offset).or_default();
        if prepend {
            list.insert(0, fragment);
        } else {
            list.push(fragment);
        }
        Ok(())
    }

    fn render(&mut self, body: Body<'_>, parent: Node<'_>, new: &NewNode) -> Result<String> {
        let prefixes = body.prefixes();
        let t = prefixes.prefix(Ns::Text);
        let mut out = String::new();
        match new {
            NewNode::Paragraph { text, style } | NewNode::Heading { text, style, .. } => {
                xml::check_text(text, "paragraph text")?;
                let (local, level) = match new {
                    NewNode::Heading { level, .. } => {
                        if !(1..=10).contains(level) {
                            return Err(invalid("ODT heading levels run from 1 to 10"));
                        }
                        ("h", Some(*level))
                    },
                    _ => ("p", None),
                };
                out.push('<');
                out.push_str(t);
                out.push(':');
                out.push_str(local);
                if let Some(level) = level {
                    push_attribute(
                        &mut out,
                        prefixes,
                        Ns::Text,
                        "outline-level",
                        &level.to_string(),
                    );
                }
                if let Some(style) = style {
                    check_style(body, style)?;
                    push_attribute(&mut out, prefixes, Ns::Text, "style-name", style);
                }
                close_element(&mut out, prefixes, Ns::Text, local, |out| {
                    xml::encode_text(out, prefixes, text);
                });
            },
            NewNode::Span { text, properties } => {
                xml::check_text(text, "run text")?;
                properties.validate()?;
                out.push('<');
                out.push_str(t);
                out.push_str(":span");
                if !properties.is_empty() {
                    let name = self.styles.derive(
                        body,
                        StyleFamily::Text,
                        None,
                        Some(properties),
                        None,
                    )?;
                    push_attribute(&mut out, prefixes, Ns::Text, "style-name", &name);
                }
                close_element(&mut out, prefixes, Ns::Text, "span", |out| {
                    xml::encode_text(out, prefixes, text);
                });
            },
            NewNode::Hyperlink { href, text } => {
                validate_attribute(body, Attribute::Href, Some(href))?;
                xml::check_text(text, "hyperlink text")?;
                require_bound(prefixes, Ns::Xlink)?;
                out.push('<');
                out.push_str(t);
                out.push_str(":a");
                push_attribute(&mut out, prefixes, Ns::Xlink, "type", "simple");
                push_attribute(&mut out, prefixes, Ns::Xlink, "href", href);
                close_element(&mut out, prefixes, Ns::Text, "a", |out| {
                    xml::encode_text(out, prefixes, text);
                });
            },
            NewNode::Bookmark { name } => {
                validate_attribute(body, Attribute::Name, Some(name))?;
                out.push('<');
                out.push_str(t);
                out.push_str(":bookmark");
                push_attribute(&mut out, prefixes, Ns::Text, "name", name);
                out.push_str("/>");
            },
            NewNode::Table { rows, columns } => {
                if *rows == 0
                    || *columns == 0
                    || rows.saturating_mul(*columns) > MAX_GENERATED_CELLS
                {
                    return Err(invalid(format!(
                        "an inserted ODT table holds 1 to {MAX_GENERATED_CELLS} cells"
                    )));
                }
                let ta = prefixes.prefix(Ns::Table);
                let name = self.table_name(body);
                out.push('<');
                out.push_str(ta);
                out.push_str(":table");
                push_attribute(&mut out, prefixes, Ns::Table, "name", &name);
                out.push('>');
                out.push('<');
                out.push_str(ta);
                out.push_str(":table-column");
                if *columns > 1 {
                    push_attribute(
                        &mut out,
                        prefixes,
                        Ns::Table,
                        "number-columns-repeated",
                        &columns.to_string(),
                    );
                }
                out.push_str("/>");
                for _ in 0..*rows {
                    render_row(&mut out, prefixes, *columns);
                }
                out.push_str("</");
                out.push_str(ta);
                out.push_str(":table>");
            },
            NewNode::Row => {
                let columns = table_columns(parent);
                if columns == 0 || columns > MAX_GENERATED_CELLS {
                    return Err(invalid(
                        "a row can only be added to an ODT table whose first row has cells",
                    ));
                }
                render_row(&mut out, prefixes, columns);
            },
            NewNode::List { items } => {
                if items.is_empty() || items.len() > MAX_GENERATED_CELLS {
                    return Err(invalid(format!(
                        "an inserted ODT list holds 1 to {MAX_GENERATED_CELLS} items"
                    )));
                }
                out.push('<');
                out.push_str(t);
                out.push_str(":list>");
                for item in items {
                    xml::check_text(item, "list item text")?;
                    render_item(&mut out, prefixes, item);
                }
                out.push_str("</");
                out.push_str(t);
                out.push_str(":list>");
            },
            NewNode::ListItem { text } => {
                xml::check_text(text, "list item text")?;
                render_item(&mut out, prefixes, text);
            },
            NewNode::Section { name } => {
                validate_attribute(body, Attribute::Name, Some(name))?;
                out.push('<');
                out.push_str(t);
                out.push_str(":section");
                push_attribute(&mut out, prefixes, Ns::Text, "name", name);
                out.push_str("><");
                out.push_str(t);
                out.push_str(":p/></");
                out.push_str(t);
                out.push_str(":section>");
            },
            NewNode::Footnote { text } | NewNode::Endnote { text } => {
                xml::check_text(text, "note text")?;
                let endnote = matches!(new, NewNode::Endnote { .. });
                let (class, prefix) = if endnote {
                    ("endnote", "edn")
                } else {
                    ("footnote", "ftn")
                };
                let citation = self.note_citation(body, endnote);
                let id = self.note_id(body, prefix);
                out.push('<');
                out.push_str(t);
                out.push_str(":note");
                push_attribute(&mut out, prefixes, Ns::Text, "id", &id);
                push_attribute(&mut out, prefixes, Ns::Text, "note-class", class);
                out.push_str("><");
                out.push_str(t);
                out.push_str(":note-citation>");
                out.push_str(&citation.to_string());
                out.push_str("</");
                out.push_str(t);
                out.push_str(":note-citation><");
                out.push_str(t);
                out.push_str(":note-body>");
                render_paragraph(&mut out, prefixes, None, text);
                out.push_str("</");
                out.push_str(t);
                out.push_str(":note-body></");
                out.push_str(t);
                out.push_str(":note>");
            },
            NewNode::Annotation { text, author } => {
                xml::check_text(text, "comment text")?;
                let o = prefixes.prefix(Ns::Office);
                out.push('<');
                out.push_str(o);
                out.push_str(":annotation>");
                if let Some(author) = author {
                    xml::check_text(author, "comment author")?;
                    require_bound(prefixes, Ns::Dc)?;
                    let dc = prefixes.prefix(Ns::Dc);
                    out.push('<');
                    out.push_str(dc);
                    out.push_str(":creator>");
                    xml::escape_into(&mut out, author);
                    out.push_str("</");
                    out.push_str(dc);
                    out.push_str(":creator>");
                }
                render_paragraph(&mut out, prefixes, None, text);
                out.push_str("</");
                out.push_str(o);
                out.push_str(":annotation>");
            },
        }
        Ok(out)
    }

    fn table_name(&mut self, body: Body<'_>) -> String {
        let existing: HashSet<String> = body
            .nodes()
            .filter(|node| node.kind() == Kind::Table)
            .filter_map(|node| {
                node.attribute(Attribute::Name)
                    .map(|name| name.into_owned())
            })
            .collect();
        loop {
            self.generated.tables += 1;
            let name = format!("Table{}", existing.len() + self.generated.tables);
            if !existing.contains(&name) {
                return name;
            }
        }
    }

    fn note_citation(&mut self, body: Body<'_>, endnote: bool) -> usize {
        let kind = if endnote {
            Kind::Endnote
        } else {
            Kind::Footnote
        };
        let existing = body.nodes().filter(|node| node.kind() == kind).count();
        let staged = if endnote {
            self.generated.endnotes += 1;
            self.generated.endnotes
        } else {
            self.generated.footnotes += 1;
            self.generated.footnotes
        };
        existing + staged
    }

    fn note_id(&mut self, body: Body<'_>, prefix: &str) -> String {
        let existing: HashSet<String> = body
            .nodes()
            .filter(|node| matches!(node.kind(), Kind::Footnote | Kind::Endnote))
            .filter_map(|node| node.raw_attribute(Ns::Text, "id").map(|id| id.into_owned()))
            .collect();
        loop {
            self.generated.notes += 1;
            let id = format!("{prefix}{}", existing.len() + self.generated.notes);
            if !existing.contains(&id) {
                return id;
            }
        }
    }
}

fn require_bound(prefixes: &Prefixes, ns: Ns) -> Result<()> {
    if prefixes.is_bound(ns) {
        Ok(())
    } else {
        Err(invalid(format!(
            "ODT content.xml does not declare the {} namespace this edit writes",
            ns.uri()
        )))
    }
}

fn check_style(body: Body<'_>, name: &str) -> Result<()> {
    if body.automatic_style(name).is_some() || body.common_style(name).is_some() {
        Ok(())
    } else {
        Err(invalid(format!("ODT style `{name}` does not exist")))
    }
}

fn validate_attribute(body: Body<'_>, attribute: Attribute, value: Option<&str>) -> Result<()> {
    let required = || invalid(format!("the ODT {attribute:?} attribute cannot be removed"));
    match attribute {
        Attribute::StyleName => {
            if let Some(value) = value {
                check_style(body, value)?;
            }
        },
        Attribute::Name => {
            let value = value.ok_or_else(required)?;
            if value.is_empty() || value.len() > 1024 {
                return Err(invalid("ODT names hold 1 to 1024 bytes"));
            }
            xml::check_text(value, "name")?;
        },
        Attribute::OutlineLevel => {
            let value = value.ok_or_else(required)?;
            if !value
                .parse::<u8>()
                .is_ok_and(|level| (1..=10).contains(&level))
            {
                return Err(invalid("ODT heading levels run from 1 to 10"));
            }
        },
        Attribute::Href => {
            let value = value.ok_or_else(required)?;
            if value.is_empty() || value.len() > 8192 {
                return Err(invalid("ODT hyperlink targets hold 1 to 8192 bytes"));
            }
            xml::check_text(value, "hyperlink target")?;
            let scheme = value.split(':').next().unwrap_or_default();
            if value.contains(':') && scheme.eq_ignore_ascii_case("javascript") {
                return Err(invalid(
                    "ODT hyperlinks stay inert: script targets are refused",
                ));
            }
        },
        Attribute::Title | Attribute::TargetFrame => {
            if let Some(value) = value {
                xml::check_text(value, "hyperlink attribute")?;
            }
        },
        Attribute::NoteClass => {
            let value = value.ok_or_else(required)?;
            if !matches!(value, "footnote" | "endnote") {
                return Err(invalid("ODT note classes are footnote and endnote"));
            }
        },
        Attribute::ColumnsSpanned
        | Attribute::RowsSpanned
        | Attribute::ColumnsRepeated
        | Attribute::ValueType => {
            return Err(invalid(format!(
                "the ODT {attribute:?} attribute shapes the table grid and is not set directly"
            )));
        },
        Attribute::Width | Attribute::Height | Attribute::X | Attribute::Y => {
            let value = value.ok_or_else(required)?;
            check_length(
                value,
                "frame geometry",
                matches!(attribute, Attribute::X | Attribute::Y),
            )?;
        },
        Attribute::AnchorType => {
            let value = value.ok_or_else(required)?;
            if !matches!(value, "paragraph" | "char" | "as-char" | "page" | "frame") {
                return Err(invalid(
                    "ODT anchors are paragraph, char, as-char, page or frame",
                ));
            }
        },
        Attribute::ContinueNumbering | Attribute::IsListHeader | Attribute::Protected => {
            if let Some(value) = value {
                if !matches!(value, "true" | "false") {
                    return Err(invalid(format!(
                        "the ODT {attribute:?} attribute is true or false"
                    )));
                }
            }
        },
        Attribute::Display => {
            if let Some(value) = value {
                if !matches!(value, "true" | "none" | "condition") {
                    return Err(invalid("ODT section display is true, none or condition"));
                }
            }
        },
    }
    Ok(())
}

/// Whether a node can leave its place without breaking its container.
fn check_detachable(node: Node<'_>) -> Result<()> {
    match node.kind() {
        Kind::Body => Err(invalid("the ODT body cannot be removed or moved")),
        Kind::Cell | Kind::CoveredCell => Err(invalid(
            "ODT table cells keep the table grid: remove or move the row instead",
        )),
        Kind::Row => {
            let spans = node.children().any(|cell| {
                cell.kind() == Kind::CoveredCell
                    || cell
                        .attribute(Attribute::RowsSpanned)
                        .is_some_and(|rows| rows.parse::<u32>().is_ok_and(|rows| rows > 1))
            });
            if spans {
                Err(invalid(
                    "an ODT row with vertically merged cells cannot be removed or moved",
                ))
            } else {
                Ok(())
            }
        },
        _ => Ok(()),
    }
}

/// The end marker that closes a range bookmark or annotation.
fn end_marker(node: Node<'_>) -> Option<Range<usize>> {
    let prefixes = node.body().prefixes();
    let (ns, end_local, name_ns) = match node.kind() {
        Kind::Bookmark => {
            let tag_name = xml::tag_name(node.start_tag());
            if !prefixes.is(tag_name.as_bytes(), Ns::Text, "bookmark-start") {
                return None;
            }
            (Ns::Text, "bookmark-end", Ns::Text)
        },
        Kind::Annotation => (Ns::Office, "annotation-end", Ns::Office),
        _ => return None,
    };
    let name = node.raw_attribute(name_ns, "name")?;
    let content = node.body().content();
    let mut needle = String::from("<");
    needle.push_str(prefixes.prefix(ns));
    needle.push(':');
    needle.push_str(end_local);
    let mut cursor = node.range().end;
    while let Some(found) = content[cursor..].find(&needle) {
        let start = cursor + found;
        let end = xml::tag_end(content.as_bytes(), start)?;
        let tag = &content[start..end];
        if xml::tag_name(tag).len() + 1 == needle.len()
            && xml::attributes(tag).any(|(attr, value)| {
                prefixes.is(attr.as_bytes(), name_ns, "name") && xml::unescape(value) == name
            })
        {
            return Some(start..end);
        }
        cursor = end;
    }
    None
}

fn child_of<'a>(body: Body<'a>, parent: Node<'a>, anchor: NodeId) -> Result<Node<'a>> {
    let anchor = node(body, anchor)?;
    if anchor.parent().map(Node::id) == Some(parent.id()) {
        Ok(anchor)
    } else {
        Err(invalid(
            "an ODT placement anchor must be a child of the insertion parent",
        ))
    }
}

/// Where content inserted first goes; `None` for an empty element.
fn first_offset(parent: Node<'_>) -> Option<usize> {
    if parent.is_empty_element() {
        return None;
    }
    if parent.kind().holds_runs() {
        return Some(parent.content_range().start);
    }
    match parent.children().next() {
        Some(child) => Some(child.range().start),
        None => last_offset(parent),
    }
}

/// Where content inserted last goes; `None` for an empty element.
fn last_offset(parent: Node<'_>) -> Option<usize> {
    if parent.is_empty_element() {
        return None;
    }
    if parent.kind().holds_runs() {
        return Some(parent.content_range().end);
    }
    if let Some(last) = parent.children().last() {
        return Some(last.range().end);
    }
    if matches!(parent.kind(), Kind::Footnote | Kind::Endnote) {
        let prefixes = parent.body().prefixes();
        let content = parent.content_range();
        let xml = &parent.body().content()[content.clone()];
        let mut needle = String::from("</");
        needle.push_str(prefixes.prefix(Ns::Text));
        needle.push_str(":note-body>");
        return xml.rfind(&needle).map(|offset| content.start + offset);
    }
    Some(parent.content_range().end)
}

/// The columns a table's first row spans.
fn table_columns(table: Node<'_>) -> usize {
    let Some(row) = table.children().find(|child| child.kind() == Kind::Row) else {
        return 0;
    };
    row.children()
        .filter(|cell| matches!(cell.kind(), Kind::Cell | Kind::CoveredCell))
        .map(|cell| {
            cell.attribute(Attribute::ColumnsRepeated)
                .and_then(|repeat| repeat.parse::<usize>().ok())
                .unwrap_or(1)
                .max(1)
        })
        .fold(0usize, usize::saturating_add)
}

fn push_attribute(out: &mut String, prefixes: &Prefixes, ns: Ns, local: &str, value: &str) {
    out.push(' ');
    prefixes.push_name(out, ns, local);
    out.push_str("=\"");
    xml::escape_into(out, value);
    out.push('"');
}

/// Finish an open start tag with `content` and the end tag, or as an empty
/// element when `content` writes nothing.
fn close_element(
    out: &mut String,
    prefixes: &Prefixes,
    ns: Ns,
    local: &str,
    content: impl FnOnce(&mut String),
) {
    let mark = out.len();
    out.push('>');
    let body_start = out.len();
    content(out);
    if out.len() == body_start {
        out.truncate(mark);
        out.push_str("/>");
    } else {
        out.push_str("</");
        prefixes.push_name(out, ns, local);
        out.push('>');
    }
}

fn render_paragraph(out: &mut String, prefixes: &Prefixes, style: Option<&str>, text: &str) {
    out.push('<');
    prefixes.push_name(out, Ns::Text, "p");
    if let Some(style) = style {
        push_attribute(out, prefixes, Ns::Text, "style-name", style);
    }
    close_element(out, prefixes, Ns::Text, "p", |out| {
        xml::encode_text(out, prefixes, text);
    });
}

fn render_item(out: &mut String, prefixes: &Prefixes, text: &str) {
    out.push('<');
    prefixes.push_name(out, Ns::Text, "list-item");
    out.push('>');
    render_paragraph(out, prefixes, None, text);
    out.push_str("</");
    prefixes.push_name(out, Ns::Text, "list-item");
    out.push('>');
}

fn render_row(out: &mut String, prefixes: &Prefixes, columns: usize) {
    out.push('<');
    prefixes.push_name(out, Ns::Table, "table-row");
    out.push('>');
    for _ in 0..columns {
        out.push('<');
        prefixes.push_name(out, Ns::Table, "table-cell");
        push_attribute(out, prefixes, Ns::Office, "value-type", "string");
        out.push_str("><");
        prefixes.push_name(out, Ns::Text, "p");
        out.push_str("/></");
        prefixes.push_name(out, Ns::Table, "table-cell");
        out.push('>');
    }
    out.push_str("</");
    prefixes.push_name(out, Ns::Table, "table-row");
    out.push('>');
}

/// Rebuild a start tag with single spaces and double quotes, so a tag from
/// a pretty-printed producer is compact authored markup.
fn normalize_tag(tag: &str) -> String {
    let empty = tag.trim_end().ends_with("/>");
    let mut out = String::with_capacity(tag.len());
    out.push('<');
    out.push_str(xml::tag_name(tag));
    for (name, value) in xml::attributes(tag) {
        out.push(' ');
        out.push_str(name);
        out.push_str("=\"");
        if value.contains('"') {
            xml::escape_into(&mut out, &xml::unescape(value));
        } else {
            out.push_str(value);
        }
        out.push('"');
    }
    out.push_str(if empty { "/>" } else { ">" });
    out
}

/// Apply attribute changes to a normalized start tag.
fn rewrite_tag(tag: &str, prefixes: &Prefixes, changes: &[(Ns, &str, Option<String>)]) -> String {
    let mut out = normalize_tag(tag);
    for (ns, local, value) in changes {
        let found = xml::attributes(&out)
            .find(|(name, _)| prefixes.is(name.as_bytes(), *ns, local))
            .map(|(name, _)| name.to_owned());
        match (found, value) {
            (Some(name), Some(value)) => {
                if let Some((_, value_range)) = xml::attribute_span(&out, &name) {
                    let mut escaped = String::new();
                    xml::escape_into(&mut escaped, value);
                    out.replace_range(value_range, &escaped);
                }
            },
            (Some(name), None) => {
                if let Some((whole, _)) = xml::attribute_span(&out, &name) {
                    // Remove the pair with the space before it.
                    out.replace_range(whole.start - 1..whole.end, "");
                }
            },
            (None, Some(value)) => {
                let close = if out.ends_with("/>") {
                    out.len() - 2
                } else {
                    out.len() - 1
                };
                let mut pair = String::new();
                push_attribute(&mut pair, prefixes, *ns, local, value);
                out.insert_str(close, &pair);
            },
            (None, None) => {},
        }
    }
    out
}

/// The open form of a start tag (`<x a="b">` for `<x a="b"/>`).
fn open_form(tag: &str) -> String {
    match tag.strip_suffix("/>") {
        Some(open) => format!("{open}>"),
        None => tag.to_owned(),
    }
}

fn end_tag(tag: &str) -> String {
    format!("</{}>", xml::tag_name(tag))
}

/// The direct child elements of `content` with their ranges (relative).
fn direct_children(content: &str) -> Vec<(String, Range<usize>)> {
    let mut out = Vec::new();
    let mut reader = Reader::from_str(content);
    reader.config_mut().trim_text(false);
    reader.config_mut().check_end_names = false;
    let mut depth = 0usize;
    let mut open: Option<(String, usize)> = None;
    loop {
        let start = usize::try_from(reader.buffer_position()).unwrap_or(0);
        let Ok(event) = reader.read_event() else {
            break;
        };
        let end = usize::try_from(reader.buffer_position()).unwrap_or(0);
        match event {
            Event::Start(tag) => {
                if depth == 0 {
                    open = Some((
                        String::from_utf8_lossy(tag.name().as_ref()).into_owned(),
                        start,
                    ));
                }
                depth += 1;
            },
            Event::Empty(tag) => {
                if depth == 0 {
                    out.push((
                        String::from_utf8_lossy(tag.name().as_ref()).into_owned(),
                        start..end,
                    ));
                }
            },
            Event::End(_) => {
                depth = depth.saturating_sub(1);
                if depth == 0 {
                    if let Some((name, begin)) = open.take() {
                        out.push((name, begin..end));
                    }
                }
            },
            Event::Eof => break,
            _ => {},
        }
    }
    out
}

/// The family of a derived automatic style.
#[derive(Clone, Copy, Debug, PartialEq, Eq, Hash)]
enum StyleFamily {
    Text,
    Paragraph,
}

/// Derived automatic styles, deduplicated by base and edit.
#[derive(Debug, Default)]
struct StyleAllocator {
    keys: Vec<(String, String)>,
    definitions: Vec<String>,
    names: HashSet<String>,
    counter: usize,
}

impl StyleAllocator {
    fn derive(
        &mut self,
        body: Body<'_>,
        family: StyleFamily,
        base: Option<&str>,
        text: Option<&TextPropertyEdit>,
        paragraph: Option<&ParagraphPropertyEdit>,
    ) -> Result<String> {
        let key = format!("{family:?}|{base:?}|{text:?}|{paragraph:?}");
        if let Some((_, name)) = self.keys.iter().find(|(held, _)| *held == key) {
            return Ok(name.clone());
        }
        let prefixes = body.prefixes();
        let name = loop {
            self.counter += 1;
            let candidate = match family {
                StyleFamily::Text => format!("litchi_T{}", self.counter),
                StyleFamily::Paragraph => format!("litchi_P{}", self.counter),
            };
            if body.automatic_style(&candidate).is_none()
                && body.common_style(&candidate).is_none()
                && !self.names.contains(&candidate)
            {
                break candidate;
            }
        };
        let text_changes = text.map(TextPropertyEdit::attributes).unwrap_or_default();
        let paragraph_changes = paragraph
            .map(ParagraphPropertyEdit::attributes)
            .unwrap_or_default();
        let automatic = base.and_then(|base| body.automatic_style(base));
        let definition = match automatic {
            Some(style) => clone_style(style, prefixes, &name, &text_changes, &paragraph_changes),
            None => {
                let mut out = String::from("<");
                prefixes.push_name(&mut out, Ns::Style, "style");
                push_attribute(&mut out, prefixes, Ns::Style, "name", &name);
                let family_name = match family {
                    StyleFamily::Text => "text",
                    StyleFamily::Paragraph => "paragraph",
                };
                push_attribute(&mut out, prefixes, Ns::Style, "family", family_name);
                if let Some(base) = base {
                    if body.common_style(base).is_some() {
                        push_attribute(&mut out, prefixes, Ns::Style, "parent-style-name", base);
                    }
                }
                out.push('>');
                if !paragraph_changes.is_empty() {
                    push_property(
                        &mut out,
                        prefixes,
                        "paragraph-properties",
                        &paragraph_changes,
                    );
                }
                if !text_changes.is_empty() {
                    push_property(&mut out, prefixes, "text-properties", &text_changes);
                }
                out.push_str("</");
                prefixes.push_name(&mut out, Ns::Style, "style");
                out.push('>');
                out
            },
        };
        self.names.insert(name.clone());
        self.keys.push((key, name.clone()));
        self.definitions.push(definition);
        Ok(name)
    }
}

fn push_property(
    out: &mut String,
    prefixes: &Prefixes,
    local: &str,
    attributes: &[(Ns, &'static str, String)],
) {
    out.push('<');
    prefixes.push_name(out, Ns::Style, local);
    for (ns, name, value) in attributes {
        push_attribute(out, prefixes, *ns, name, value);
    }
    out.push_str("/>");
}

/// Copy an automatic style under a new name with changed properties. The
/// copy is rebuilt from its tags, so it is compact whatever the source
/// formatting.
fn clone_style(
    style: Style<'_>,
    prefixes: &Prefixes,
    name: &str,
    text: &[(Ns, &'static str, String)],
    paragraph: &[(Ns, &'static str, String)],
) -> String {
    let source = style.xml();
    let mut reader = Reader::from_str(source);
    reader.config_mut().trim_text(false);
    reader.config_mut().check_end_names = false;
    let mut out = String::with_capacity(source.len() + 128);
    let mut depth = 0usize;
    let mut seen_text = false;
    let mut seen_paragraph = false;
    let changes = |list: &[(Ns, &'static str, String)]| -> Vec<(Ns, &'static str, Option<String>)> {
        list.iter()
            .map(|(ns, local, value)| (*ns, *local, Some(value.clone())))
            .collect()
    };
    loop {
        let start = usize::try_from(reader.buffer_position()).unwrap_or(0);
        let Ok(event) = reader.read_event() else {
            break;
        };
        let end = usize::try_from(reader.buffer_position()).unwrap_or(0);
        let raw = &source[start..end];
        match event {
            Event::Start(_) | Event::Empty(_) => {
                let empty = matches!(event, Event::Empty(_));
                let tag_name = xml::tag_name(raw);
                let tag = if depth == 0 {
                    let renamed =
                        rewrite_tag(raw, prefixes, &[(Ns::Style, "name", Some(name.to_owned()))]);
                    if empty && (!text.is_empty() || !paragraph.is_empty()) {
                        // Grow the empty declaration to hold properties.
                        out.push_str(&open_form(&renamed));
                        if !paragraph.is_empty() {
                            push_property(&mut out, prefixes, "paragraph-properties", paragraph);
                        }
                        if !text.is_empty() {
                            push_property(&mut out, prefixes, "text-properties", text);
                        }
                        out.push_str(&end_tag(&renamed));
                        break;
                    }
                    renamed
                } else if depth == 1
                    && prefixes.is(tag_name.as_bytes(), Ns::Style, "text-properties")
                {
                    seen_text = true;
                    rewrite_tag(raw, prefixes, &changes(text))
                } else if depth == 1
                    && prefixes.is(tag_name.as_bytes(), Ns::Style, "paragraph-properties")
                {
                    seen_paragraph = true;
                    rewrite_tag(raw, prefixes, &changes(paragraph))
                } else {
                    normalize_tag(raw)
                };
                if depth == 0 && !seen_paragraph && !paragraph.is_empty() {
                    out.push_str(&tag);
                    push_property(&mut out, prefixes, "paragraph-properties", paragraph);
                    seen_paragraph = true;
                } else {
                    out.push_str(&tag);
                }
                if !empty {
                    depth += 1;
                }
            },
            Event::End(_) => {
                depth = depth.saturating_sub(1);
                if depth == 0 && !seen_text && !text.is_empty() {
                    push_property(&mut out, prefixes, "text-properties", text);
                }
                out.push_str(raw.trim());
                if depth == 0 {
                    break;
                }
            },
            Event::Text(text_event) => {
                if !text_event.iter().all(u8::is_ascii_whitespace) {
                    out.push_str(raw);
                }
            },
            Event::Eof => break,
            _ => out.push_str(raw),
        }
    }
    out
}

/// One splice of a source part.
struct Splice {
    range: Range<usize>,
    bytes: String,
    /// The replacement is exactly one start tag.
    start_tag: bool,
}

/// Compile and publish `edits` against `document`, returning the new package
/// bytes, or `None` when the edits change nothing.
pub(crate) fn apply(document: &Document, edits: BodyEdits) -> Result<Option<Vec<u8>>> {
    let body = document.body()?;
    let content = body.content();
    let prefixes = body.prefixes();
    let mut edits = edits;
    let mut splices: Vec<Splice> = Vec::new();
    let node_edits = std::mem::take(&mut edits.nodes);
    for (id, edit) in &node_edits {
        let target = node(body, NodeId(*id))?;
        let expansion = edits.expansions.remove(id);
        node_splices(
            body,
            target,
            edit,
            expansion.as_deref(),
            &mut edits.styles,
            &mut splices,
        )?;
    }
    for (id, fragments) in std::mem::take(&mut edits.expansions) {
        let target = node(body, NodeId(id))?;
        node_splices(
            body,
            target,
            &NodeEdit::default(),
            Some(&fragments),
            &mut edits.styles,
            &mut splices,
        )?;
    }
    for (offset, fragments) in &edits.inserts {
        if fragments.is_empty() {
            continue;
        }
        let mut bytes = String::new();
        for fragment in fragments {
            push_fragment(&mut bytes, content, fragment);
        }
        splices.push(Splice {
            range: *offset..*offset,
            bytes,
            start_tag: false,
        });
    }
    for marker in &edits.marker_removals {
        splices.push(Splice {
            range: marker.clone(),
            bytes: String::new(),
            start_tag: false,
        });
    }
    if !edits.styles.definitions.is_empty() {
        let definitions = edits.styles.definitions.concat();
        let o = prefixes.prefix(Ns::Office);
        let splice = match body.tree().automatic_insert {
            AutomaticInsert::Before(offset) => Splice {
                range: offset as usize..offset as usize,
                bytes: definitions,
                start_tag: false,
            },
            AutomaticInsert::Expand(start, end) => {
                let tag = &content[start as usize..end as usize];
                Splice {
                    range: start as usize..end as usize,
                    bytes: format!(
                        "{}{definitions}{}",
                        open_form(&normalize_tag(tag)),
                        end_tag(tag)
                    ),
                    start_tag: false,
                }
            },
            AutomaticInsert::Create(offset) => Splice {
                range: offset as usize..offset as usize,
                bytes: format!("<{o}:automatic-styles>{definitions}</{o}:automatic-styles>"),
                start_tag: false,
            },
        };
        splices.push(splice);
    }
    splices.retain(|splice| content[splice.range.clone()] != splice.bytes);
    // Insertions at one offset were merged; distinct splices must not overlap.
    splices.sort_by_key(|splice| (splice.range.start, splice.range.end));
    for pair in splices.windows(2) {
        let (left, right) = (&pair[0], &pair[1]);
        if right.range.start < left.range.end
            || (left.range.is_empty()
                && right.range.is_empty()
                && left.range.start == right.range.start)
        {
            return Err(invalid("ODT body edits compiled to overlapping splices"));
        }
    }
    let package = document.transaction_package();
    let mut publications = Vec::new();
    if !splices.is_empty() {
        let part = XmlSourcePart::load(package, crate::constants::ODF_CONTENT)?;
        if part.bytes() != content.as_bytes() {
            return Err(invalid(
                "ODT content.xml changed between indexing and publication",
            ));
        }
        let mut publication = XmlSplicePublication::new(part.clone());
        for splice in splices {
            let source = &content.as_bytes()[splice.range.clone()];
            let proof = part.checked_range(splice.range.clone(), source)?;
            publication.replace(proof, classify(splice.bytes, splice.start_tag)?)?;
        }
        publications.push(publication);
    }
    if !edits.metadata.is_empty() {
        if let Some(publication) = metadata_publication(document, &edits.metadata)? {
            publications.push(publication);
        }
    }
    if publications.is_empty() {
        return Ok(None);
    }
    crate::core::rebuild_package_with_xml_splices(package, publications, MAX_PACKAGE_BYTES)
        .map(Some)
}

fn classify(bytes: String, start_tag: bool) -> Result<AuthoredXmlFragment> {
    if bytes.is_empty() {
        Ok(AuthoredXmlFragment::deletion())
    } else if start_tag {
        AuthoredXmlFragment::start_tag(bytes.into_bytes())
    } else if bytes.starts_with('<') {
        AuthoredXmlFragment::markup(bytes.into_bytes())
    } else {
        AuthoredXmlFragment::text(bytes.into_bytes())
    }
}

fn push_fragment(out: &mut String, content: &str, fragment: &Fragment) {
    match fragment {
        Fragment::Markup(markup) => out.push_str(markup),
        Fragment::Copy(range) => out.push_str(&content[range.clone()]),
    }
}

/// The splices one node edit compiles to.
fn node_splices(
    body: Body<'_>,
    target: Node<'_>,
    edit: &NodeEdit,
    expansion: Option<&[Fragment]>,
    styles: &mut StyleAllocator,
    out: &mut Vec<Splice>,
) -> Result<()> {
    let content = body.content();
    let prefixes = body.prefixes();
    match &edit.replacement {
        Some(Replacement::Remove) => {
            out.push(Splice {
                range: target.range(),
                bytes: String::new(),
                start_tag: false,
            });
            return Ok(());
        },
        Some(Replacement::Bytes(range)) => {
            out.push(Splice {
                range: target.range(),
                bytes: content[range.clone()].to_owned(),
                start_tag: false,
            });
            return Ok(());
        },
        None => {},
    }
    let kind = target.kind();
    let mut changes: Vec<(Ns, &str, Option<String>)> = edit
        .attributes
        .iter()
        .map(|(ns, local, value)| (*ns, *local, value.clone()))
        .collect();
    let formatting = edit.text_properties.is_some() || edit.paragraph_properties.is_some();
    let new_style = if formatting {
        let base = target.style_name();
        let pending = changes
            .iter()
            .find(|(ns, local, _)| *ns == Ns::Text && *local == "style-name")
            .map(|(_, _, value)| value.clone());
        let base = match pending {
            Some(value) => value,
            None => base.map(std::borrow::Cow::into_owned),
        };
        let family = if matches!(kind, Kind::Paragraph | Kind::Heading) {
            StyleFamily::Paragraph
        } else {
            StyleFamily::Text
        };
        Some(styles.derive(
            body,
            family,
            base.as_deref(),
            edit.text_properties.as_ref(),
            edit.paragraph_properties.as_ref(),
        )?)
    } else {
        None
    };
    if kind == Kind::Text {
        let inner = match &edit.text {
            Some(text) => {
                let mut encoded = String::new();
                xml::encode_text(&mut encoded, prefixes, text);
                encoded
            },
            None => content[target.range()].to_owned(),
        };
        let bytes = match &new_style {
            Some(style) => {
                let mut wrapped = String::from("<");
                prefixes.push_name(&mut wrapped, Ns::Text, "span");
                push_attribute(&mut wrapped, prefixes, Ns::Text, "style-name", style);
                wrapped.push('>');
                wrapped.push_str(&inner);
                wrapped.push_str("</");
                prefixes.push_name(&mut wrapped, Ns::Text, "span");
                wrapped.push('>');
                wrapped
            },
            None if inner.contains('<') => {
                let mut wrapped = String::from("<");
                prefixes.push_name(&mut wrapped, Ns::Text, "span");
                wrapped.push('>');
                wrapped.push_str(&inner);
                wrapped.push_str("</");
                prefixes.push_name(&mut wrapped, Ns::Text, "span");
                wrapped.push('>');
                wrapped
            },
            None => inner,
        };
        out.push(Splice {
            range: target.range(),
            bytes,
            start_tag: false,
        });
        return Ok(());
    }
    if let Some(style) = &new_style {
        let (ns, local) = Attribute::StyleName
            .qualified(kind)
            .unwrap_or((Ns::Text, "style-name"));
        changes.retain(|(held_ns, held_local, _)| !(*held_ns == ns && *held_local == local));
        changes.push((ns, local, Some(style.clone())));
    }
    if kind == Kind::Cell && edit.text.is_some() {
        changes.push((Ns::Office, "value-type", Some("string".to_owned())));
        for local in [
            "value",
            "date-value",
            "time-value",
            "boolean-value",
            "string-value",
            "currency",
        ] {
            changes.push((Ns::Office, local, None));
        }
        changes.push((Ns::Table, "formula", None));
    }
    let original = target.start_tag();
    let tag = if changes.is_empty() {
        normalize_tag(original)
    } else {
        rewrite_tag(original, prefixes, &changes)
    };
    if edit.text.is_none() && expansion.is_none() {
        if !changes.is_empty() {
            let empty = target.is_empty_element();
            out.push(Splice {
                range: start_tag_range(target),
                bytes: tag,
                start_tag: !empty,
            });
        }
        return Ok(());
    }
    // A whole replacement: start tag, new content, end tag.
    let mut bytes = open_form(&tag);
    if let Some(text) = &edit.text {
        whole_content(target, text, &mut bytes)?;
    } else if let Some(fragments) = expansion {
        for fragment in fragments {
            push_fragment(&mut bytes, content, fragment);
        }
    }
    bytes.push_str(&end_tag(&tag));
    out.push(Splice {
        range: target.range(),
        bytes,
        start_tag: false,
    });
    Ok(())
}

/// The new content of an element whose text is replaced.
fn whole_content(target: Node<'_>, text: &str, out: &mut String) -> Result<()> {
    let body = target.body();
    let prefixes = body.prefixes();
    let content = body.content();
    match target.kind() {
        Kind::Paragraph | Kind::Heading | Kind::Span | Kind::Hyperlink => {
            xml::encode_text(out, prefixes, text);
        },
        Kind::Cell | Kind::ListItem | Kind::TextBox => {
            let style = target
                .children()
                .find(|child| matches!(child.kind(), Kind::Paragraph | Kind::Heading))
                .and_then(Node::style_name);
            render_paragraph(out, prefixes, style.as_deref(), text);
        },
        Kind::Footnote | Kind::Endnote => {
            let range = target.content_range();
            for (name, child) in direct_children(&content[range.clone()]) {
                if prefixes.is(name.as_bytes(), Ns::Text, "note-citation") {
                    let child = range.start + child.start..range.start + child.end;
                    out.push_str(&content[child]);
                }
            }
            out.push('<');
            prefixes.push_name(out, Ns::Text, "note-body");
            out.push('>');
            let style = target.children().next().and_then(Node::style_name);
            render_paragraph(out, prefixes, style.as_deref(), text);
            out.push_str("</");
            prefixes.push_name(out, Ns::Text, "note-body");
            out.push('>');
        },
        Kind::Annotation => {
            let range = target.content_range();
            for (name, child) in direct_children(&content[range.clone()]) {
                let keep = prefixes
                    .resolve(name.as_bytes())
                    .is_some_and(|(ns, local)| {
                        ns == Ns::Dc || (ns == Ns::Meta && local == b"date-string")
                    });
                if keep {
                    let child = range.start + child.start..range.start + child.end;
                    out.push_str(&content[child]);
                }
            }
            render_paragraph(out, prefixes, None, text);
        },
        other => {
            return Err(invalid(format!(
                "ODT {other:?} nodes have no settable text"
            )));
        },
    }
    Ok(())
}

/// One `meta.xml` publication for the staged metadata edits.
fn metadata_publication(
    document: &Document,
    fields: &BTreeMap<u8, (Metadata, Option<String>)>,
) -> Result<Option<XmlSplicePublication>> {
    let package = document.transaction_package();
    if !package.has_file("meta.xml")? {
        return Err(invalid(
            "the ODT package has no meta.xml for metadata edits",
        ));
    }
    let part = XmlSourcePart::load(package, "meta.xml")?;
    let source = std::str::from_utf8(part.bytes())
        .map_err(|_error| invalid("ODT meta.xml is not UTF-8"))?
        .to_owned();
    let prefixes = Prefixes::of_root(&source);
    for ns in [Ns::Dc, Ns::Meta] {
        if fields
            .values()
            .any(|(field, value)| value.is_some() && field.qualified().0 == ns)
        {
            require_bound(&prefixes, ns)?;
        }
    }
    // Locate office:meta and its children.
    let mut reader = Reader::from_str(&source);
    reader.config_mut().trim_text(false);
    let mut depth = 0usize;
    let mut meta: Option<(Range<usize>, bool)> = None;
    let mut meta_close: Option<usize> = None;
    let mut children: Vec<(Ns, String, Range<usize>)> = Vec::new();
    let mut open_child: Option<(Ns, String, usize)> = None;
    loop {
        let start = usize::try_from(reader.buffer_position()).unwrap_or(0);
        let event = reader
            .read_event()
            .map_err(|error| invalid(format!("ODT meta.xml: {error}")))?;
        let end = usize::try_from(reader.buffer_position()).unwrap_or(0);
        match event {
            Event::Start(tag) | Event::Empty(tag) => {
                let empty = source.as_bytes().get(end.saturating_sub(2)) == Some(&b'/');
                let name = tag.name();
                let resolved = prefixes.resolve(name.as_ref());
                if depth == 1
                    && resolved.is_some_and(|(ns, local)| ns == Ns::Office && local == b"meta")
                {
                    meta = Some((start..end, empty));
                } else if depth == 2 && meta.is_some() && meta_close.is_none() {
                    if let Some((ns, local)) = resolved {
                        let local = String::from_utf8_lossy(local).into_owned();
                        if empty {
                            children.push((ns, local, start..end));
                        } else {
                            open_child = Some((ns, local, start));
                        }
                    }
                }
                if !empty {
                    depth += 1;
                }
            },
            Event::End(_) => {
                depth = depth.saturating_sub(1);
                if depth == 2 {
                    if let Some((ns, local, begin)) = open_child.take() {
                        children.push((ns, local, begin..end));
                    }
                }
                if depth == 1 && meta.is_some() && meta_close.is_none() {
                    meta_close = Some(start);
                }
            },
            Event::Eof => break,
            _ => {},
        }
    }
    let Some((meta_range, meta_empty)) = meta else {
        return Err(invalid("ODT meta.xml has no office:meta element"));
    };
    let mut splices: Vec<Splice> = Vec::new();
    let mut appended = String::new();
    for (field, value) in fields.values() {
        let (ns, local) = field.qualified();
        let existing: Vec<&Range<usize>> = children
            .iter()
            .filter(|(held_ns, held_local, _)| *held_ns == ns && held_local == local)
            .map(|(_, _, range)| range)
            .collect();
        let mut rendered = String::new();
        if let Some(value) = value {
            let values: Vec<&str> = if *field == Metadata::Keywords {
                value
                    .split(',')
                    .map(str::trim)
                    .filter(|v| !v.is_empty())
                    .collect()
            } else {
                vec![value.as_str()]
            };
            for value in values {
                rendered.push('<');
                prefixes.push_name(&mut rendered, ns, local);
                if value.is_empty() {
                    rendered.push_str("/>");
                } else {
                    rendered.push('>');
                    xml::escape_into(&mut rendered, value);
                    rendered.push_str("</");
                    prefixes.push_name(&mut rendered, ns, local);
                    rendered.push('>');
                }
            }
        }
        match existing.split_first() {
            Some((first, rest)) => {
                splices.push(Splice {
                    range: (*first).clone(),
                    bytes: rendered,
                    start_tag: false,
                });
                for range in rest {
                    splices.push(Splice {
                        range: (*range).clone(),
                        bytes: String::new(),
                        start_tag: false,
                    });
                }
            },
            None => appended.push_str(&rendered),
        }
    }
    if !appended.is_empty() {
        if meta_empty {
            let tag = &source[meta_range.clone()];
            splices.push(Splice {
                range: meta_range,
                bytes: format!(
                    "{}{appended}{}",
                    open_form(&normalize_tag(tag)),
                    end_tag(tag)
                ),
                start_tag: false,
            });
        } else {
            let close = meta_close.ok_or_else(|| invalid("ODT office:meta is not closed"))?;
            splices.push(Splice {
                range: close..close,
                bytes: appended,
                start_tag: false,
            });
        }
    }
    splices.retain(|splice| source[splice.range.clone()] != splice.bytes);
    if splices.is_empty() {
        return Ok(None);
    }
    splices.sort_by_key(|splice| (splice.range.start, splice.range.end));
    let mut publication = XmlSplicePublication::new(part.clone());
    for splice in splices {
        let proof = part.checked_range(
            splice.range.clone(),
            &source.as_bytes()[splice.range.clone()],
        )?;
        publication.replace(proof, classify(splice.bytes, false)?)?;
    }
    Ok(Some(publication))
}
