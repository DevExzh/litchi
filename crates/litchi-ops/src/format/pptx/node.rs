//! PPTX typed nodes and their node view (record 0010 §2–§3, §4.6).
//!
//! A [`Node`] is one element of an open presentation: the presentation, a
//! slide, a shape-tree element (shape, picture, table, chart, group,
//! connector, SmartArt or OLE frame), a paragraph, run or line break of a
//! shape's text, a table row or cell, a slide's notes or one of its
//! comments. It borrows the [`Pptx`] handle and carries the leaf values it
//! reads:
//!
//! - a slide keeps its position, and pins its shape scene
//!   ([`SharedScene`]) once a read of its shapes needs it;
//! - a shape-tree element keeps a [`Handle`] into that scene, and pins
//!   its parsed text body once a read of its text needs it;
//! - a paragraph, run or line break shares its shape's parsed text body;
//! - notes and comments keep their owned text and records.
//!
//! No node keeps a `SourceSlide` (record 0010 §3.3, risk 6): scenes and
//! handles hold no package, so dropping every node lets the handle build its
//! editor. The handle itself caches nothing for the adapter; the leaf's
//! weighted scene memo is the only retention policy.
//!
//! The view's `format` member reads typed values from those leaf values with
//! OfficeCLI's keys ([`super::spec`]), local properties only.

use std::borrow::Cow;
use std::fmt;
use std::sync::{Arc, OnceLock};

use litchi_core::{FileFormat, Metadata, Position};
use litchi_drawingml::text::body::{Body, Paragraph, Run, RunProperties};
use litchi_drawingml::transform::Transform;
use litchi_pptx::master_layout::SlideLayoutKind;
use litchi_pptx::modern_comments::Status as ModernStatus;
use litchi_pptx::presentation::modern_comments::{
    Comment as ModernComment, List as ModernList, Reply as ModernReply,
};
use litchi_pptx::presentation::related::{Comment, Comments};
use litchi_pptx::shape::Shape;
use litchi_pptx::shape::owned::{Handle, Handles, Scene as SharedScene};
use litchi_pptx::table::Properties as TableProperties;

use super::Pptx;
use super::children::{Children, Listing};
use super::owner::SlideFacts;
use super::spec::{
    CELL_PROPS, COMMENT_PROPS, CellKey, CommentKey, FRAME_PROPS, FrameKey, Kind, NOTES_PROPS,
    NotesKey, PARAGRAPH_PROPS, PLACEHOLDER_TYPE, PRESENTATION_PROPS, ParagraphKey, RUN_PROPS,
    RootKey, RunKey, SHAPE_PROPS, SLIDE_PROPS, SLIDE_SIZE, SMARTART_PROPS, ShapeKey, SlideKey,
    SmartArtKey, TABLE_PROPS, TableKey,
};
use super::steps::{self, Builder, Steps};
use super::text;
use crate::error::{Error, Result};
use crate::node::{
    Children as ViewChildren, Format, NodePath, PathStep, Preview, Project, Source, Value,
    ViewOptions,
};
use crate::value::Length;

/// What a slide shows when it has no title placeholder (OfficeCLI).
const UNTITLED: &str = "(untitled)";

/// A node of a PPTX presentation (record 0010 §3.1).
///
/// Get one from the unified `get`, `find`, `query` and `children` reads; it
/// borrows the [`Pptx`] handle it came from. The generic accessors live on
/// [`crate::node::Node`]; the typed ones here give the leaf values.
///
/// ```
/// use litchi_ops::format::pptx::Node;
///
/// // The shape a node belongs to, and the slide it sits on.
/// fn locate(node: &Node<'_>) -> Option<(usize, bool)> {
///     Some((node.slide()?, node.shape().is_some()))
/// }
/// # let _ = locate;
/// ```
#[derive(Clone)]
pub struct Node<'f> {
    file: &'f Pptx,
    item: Item,
    path: OnceLock<Steps>,
    positional: OnceLock<Box<[PathStep<'static>]>>,
}

/// The element a node is, with the leaf values it reads.
#[derive(Clone)]
pub(super) enum Item {
    Presentation,
    Slide(SlideItem),
    Master(MasterItem),
    Layout(LayoutItem),
    Tree(TreeItem),
    Paragraph(TextItem),
    Run(TextItem),
    LineBreak(TextItem),
    Row(TableItem),
    Cell(TableItem),
    Notes(NotesItem),
    Comment(CommentItem),
    ModernComment(ModernCommentItem),
    Reply(ReplyItem),
}

/// A slide: its position, and its scene and facts once a read needed them.
#[derive(Clone)]
pub(super) struct SlideItem {
    pub(super) position: u32,
    /// Listed at the leaf level of a view: the node shows catalog data only
    /// and decodes nothing (record 0010 M1-Δ1).
    pub(super) listed: bool,
    scene: OnceLock<SharedScene>,
    facts: OnceLock<SlideFacts>,
}

impl SlideItem {
    pub(super) fn new(position: u32, listed: bool) -> Self {
        Self {
            position,
            listed,
            scene: OnceLock::new(),
            facts: OnceLock::new(),
        }
    }

    /// The slide's scene, read and pinned on first use.
    pub(super) fn scene(&self, file: &Pptx) -> Result<&SharedScene> {
        if let Some(scene) = self.scene.get() {
            return Ok(scene);
        }
        let scene = file
            .slide_scene(self.position as usize)?
            .ok_or_else(|| missing_slide(self.position))?;
        Ok(self.scene.get_or_init(|| scene))
    }

    fn facts(&self, file: &Pptx) -> Result<&SlideFacts> {
        if let Some(facts) = self.facts.get() {
            return Ok(facts);
        }
        let facts = file
            .slide_facts(self.position as usize)?
            .ok_or_else(|| missing_slide(self.position))?;
        Ok(self.facts.get_or_init(|| facts))
    }
}

/// A slide master catalog entry and its detached scene/facts once selected.
#[derive(Clone)]
pub(super) struct MasterItem {
    pub(super) position: u32,
    pub(super) listed: bool,
    scene: OnceLock<SharedScene>,
    facts: OnceLock<MasterFacts>,
}

impl MasterItem {
    pub(super) fn new(position: u32, listed: bool) -> Self {
        Self {
            position,
            listed,
            scene: OnceLock::new(),
            facts: OnceLock::new(),
        }
    }

    pub(super) fn scene(&self, file: &Pptx) -> Result<&SharedScene> {
        if let Some(scene) = self.scene.get() {
            return Ok(scene);
        }
        let scene = file
            .master_scene(self.position as usize)?
            .ok_or_else(|| missing_master(self.position))?;
        Ok(self.scene.get_or_init(|| scene))
    }

    fn facts(&self, file: &Pptx) -> Result<&MasterFacts> {
        if let Some(facts) = self.facts.get() {
            return Ok(facts);
        }
        let facts = file
            .master_facts(self.position as usize)?
            .ok_or_else(|| missing_master(self.position))?;
        Ok(self.facts.get_or_init(|| facts))
    }
}

/// A slide layout catalog entry and its detached scene/facts once selected.
#[derive(Clone)]
pub(super) struct LayoutItem {
    pub(super) master: u32,
    pub(super) position: u32,
    pub(super) listed: bool,
    scene: OnceLock<SharedScene>,
    facts: OnceLock<LayoutFacts>,
}

impl LayoutItem {
    pub(super) fn new(master: u32, position: u32, listed: bool) -> Self {
        Self {
            master,
            position,
            listed,
            scene: OnceLock::new(),
            facts: OnceLock::new(),
        }
    }

    pub(super) fn scene(&self, file: &Pptx) -> Result<&SharedScene> {
        if let Some(scene) = self.scene.get() {
            return Ok(scene);
        }
        let scene = file
            .layout_scene(self.master as usize, self.position as usize)?
            .ok_or_else(|| missing_layout(self.master, self.position))?;
        Ok(self.scene.get_or_init(|| scene))
    }

    fn facts(&self, file: &Pptx) -> Result<&LayoutFacts> {
        if let Some(facts) = self.facts.get() {
            return Ok(facts);
        }
        let facts = file
            .layout_facts(self.master as usize, self.position as usize)?
            .ok_or_else(|| missing_layout(self.master, self.position))?;
        Ok(self.facts.get_or_init(|| facts))
    }
}

/// Semantic facts shared by eager and source-backed master reads.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub(super) struct MasterFacts {
    pub(super) name: Option<String>,
    pub(super) preserved: bool,
}

/// Semantic facts shared by eager and source-backed layout reads.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub(super) struct LayoutFacts {
    pub(super) name: Option<String>,
    pub(super) kind: Option<SlideLayoutKind>,
}

/// A slide the catalog listed but the leaf no longer finds: the source
/// changed under the handle, which the leaf reports first; this is a guard.
fn missing_slide(position: u32) -> Error {
    Error::new(
        crate::Code::InternalError,
        format!("slide {} vanished from the catalog", position + 1),
    )
    .with_format(FileFormat::Pptx)
}

fn missing_master(position: u32) -> Error {
    Error::new(
        crate::Code::InternalError,
        format!(
            "slide master at position {} vanished from the catalog",
            position + 1
        ),
    )
    .with_format(FileFormat::Pptx)
}

fn missing_layout(master: u32, position: u32) -> Error {
    Error::new(
        crate::Code::InternalError,
        format!(
            "slide layout at position {} of master {} vanished from the catalog",
            position + 1,
            master + 1
        ),
    )
    .with_format(FileFormat::Pptx)
}

/// The scene container owning a shape-tree element.
#[derive(Clone, Copy)]
pub(super) enum TreeContainer {
    Slide(u32),
    Master(u32),
    Layout { master: u32, position: u32 },
}

impl TreeContainer {
    pub(super) const fn slide(self) -> Option<u32> {
        match self {
            Self::Slide(position) => Some(position),
            Self::Master(_) | Self::Layout { .. } => None,
        }
    }
}

/// A shape-tree element: its scene container, kind, leaf handle and position
/// among its container's elements of the kind, and its text body once read.
#[derive(Clone)]
pub(super) struct TreeItem {
    pub(super) container: TreeContainer,
    pub(super) kind: Kind,
    pub(super) handle: Handle,
    pub(super) ordinal: u32,
    body: OnceLock<std::result::Result<Option<Arc<Body>>, Error>>,
}

impl TreeItem {
    pub(super) fn new(slide: u32, kind: Kind, handle: Handle, ordinal: u32) -> Self {
        Self {
            container: TreeContainer::Slide(slide),
            kind,
            handle,
            ordinal,
            body: OnceLock::new(),
        }
    }

    pub(super) fn new_master(position: u32, kind: Kind, handle: Handle, ordinal: u32) -> Self {
        Self {
            container: TreeContainer::Master(position),
            kind,
            handle,
            ordinal,
            body: OnceLock::new(),
        }
    }

    pub(super) fn new_layout(
        master: u32,
        position: u32,
        kind: Kind,
        handle: Handle,
        ordinal: u32,
    ) -> Self {
        Self {
            container: TreeContainer::Layout { master, position },
            kind,
            handle,
            ordinal,
            body: OnceLock::new(),
        }
    }

    pub(super) const fn slide_position(&self) -> Option<u32> {
        self.container.slide()
    }

    /// The shape's text body with local properties, parsed and pinned on
    /// first use; `None` for an element without one.
    pub(super) fn body(&self) -> Result<Option<&Arc<Body>>> {
        let read = self.body.get_or_init(|| {
            self.handle
                .view()
                .text_body()
                .map(|body| body.map(Arc::new))
                .map_err(Error::from)
        });
        match read {
            Ok(body) => Ok(body.as_ref()),
            Err(error) => Err(error.clone()),
        }
    }

    /// The placeholder type token of a `p:sp` with a `p:ph`, as OfficeCLI
    /// names it; an absent `type` is the schema default, `obj`.
    pub(super) fn placeholder_type(&self) -> Option<&'static str> {
        placeholder_type(&self.handle)
    }

    /// Whether a `p:sp` is a title or centred-title placeholder.
    pub(super) fn is_title(&self) -> bool {
        is_title(&self.handle)
    }

    fn steps(&self, canonical: bool) -> Builder {
        let mut out = Builder::new();
        match self.container {
            TreeContainer::Slide(slide) => steps::push_tree(
                &mut out,
                slide,
                &self.handle,
                self.kind,
                self.ordinal,
                canonical,
            ),
            TreeContainer::Master(master) => {
                out.push(PathStep::at("slidemaster", Position::new(master as usize)));
                push_tree_tail(&mut out, &self.handle, self.kind, self.ordinal, canonical);
            },
            TreeContainer::Layout { master, position } => {
                out.push(PathStep::at("slidemaster", Position::new(master as usize)));
                out.push(PathStep::at(
                    "slidelayout",
                    Position::new(position as usize),
                ));
                push_tree_tail(&mut out, &self.handle, self.kind, self.ordinal, canonical);
            },
        }
        out
    }
}

fn push_tree_tail(out: &mut Builder, handle: &Handle, kind: Kind, ordinal: u32, canonical: bool) {
    let mut groups = Vec::new();
    let mut cursor = handle.parent();
    while let Some(group) = cursor {
        cursor = group.parent();
        groups.push(group);
    }
    for group in groups.iter().rev() {
        out.push(steps::tree_step(group, Kind::Group, None, canonical));
    }
    out.push(steps::tree_step(handle, kind, Some(ordinal), canonical));
}

/// The placeholder type of a shape, as OfficeCLI names it.
pub(super) fn placeholder_type(handle: &Handle) -> Option<&'static str> {
    let placeholder = handle.view().placeholder()?;
    match placeholder.kind() {
        // `obj` is the schema default only when the attribute was omitted.
        None => Some("obj"),
        // Preserve an authored token that this adapter cannot classify as
        // absent from the ordinary property vocabulary; raw shape metadata
        // still retains the source spelling.
        Some(token) => PLACEHOLDER_TYPE.get(token),
    }
}

/// Whether a shape is a title or centred-title placeholder (OfficeCLI
/// `IsTitle`).
pub(super) fn is_title(handle: &Handle) -> bool {
    matches!(handle.view(), Shape::Auto(_))
        && handle
            .view()
            .placeholder()
            .and_then(|placeholder| placeholder.kind())
            .is_some_and(|kind| matches!(kind.trim(), "title" | "ctrTitle"))
}

/// A paragraph, run or line break: its shape, the shared text body, and
/// where it sits in it.
#[derive(Clone)]
pub(super) struct TextItem {
    pub(super) shape: TreeItem,
    pub(super) body: Arc<Body>,
    pub(super) paragraph: u32,
    /// The index in the paragraph's runs (runs and line breaks).
    pub(super) run: u32,
    /// The position among the paragraph's runs, or among its line breaks.
    pub(super) ordinal: u32,
}

impl TextItem {
    pub(super) fn paragraph(&self) -> Option<&Paragraph> {
        self.body.paragraphs.get(self.paragraph as usize)
    }

    pub(super) fn run(&self) -> Option<&Run> {
        self.paragraph()?.runs.get(self.run as usize)
    }
}

/// A table row or cell: its table and its position.
#[derive(Clone)]
pub(super) struct TableItem {
    pub(super) table: TreeItem,
    pub(super) row: u32,
    pub(super) column: u32,
}

/// A slide's speaker notes.
#[derive(Clone)]
pub(super) struct NotesItem {
    pub(super) slide: u32,
    pub(super) text: Arc<str>,
}

/// One legacy comment of a slide.
#[derive(Clone)]
pub(super) struct CommentItem {
    pub(super) slide: u32,
    pub(super) list: Comments,
    pub(super) index: u32,
}

impl CommentItem {
    pub(super) fn comment(&self) -> Option<&Comment> {
        self.list.entries().get(self.index as usize)
    }
}

/// One modern threaded comment selected from a slide's detached comment list.
#[derive(Clone)]
pub(super) struct ModernCommentItem {
    pub(super) slide: u32,
    pub(super) list: ModernList,
    pub(super) index: u32,
}

impl ModernCommentItem {
    pub(super) fn comment(&self) -> Option<&ModernComment> {
        self.list.entries().get(self.index as usize)
    }
}

/// One reply selected from a modern comment's detached reply list.
#[derive(Clone)]
pub(super) struct ReplyItem {
    pub(super) slide: u32,
    pub(super) list: ModernList,
    pub(super) comment: u32,
    pub(super) index: u32,
}

impl ReplyItem {
    pub(super) fn comment(&self) -> Option<&ModernComment> {
        self.list.entries().get(self.comment as usize)
    }

    pub(super) fn reply(&self) -> Option<&ModernReply> {
        self.comment()?.replies().get(self.index as usize)
    }
}

impl<'f> Node<'f> {
    pub(super) fn new(file: &'f Pptx, item: Item) -> Self {
        Self {
            file,
            item,
            path: OnceLock::new(),
            positional: OnceLock::new(),
        }
    }

    pub(super) fn file(&self) -> &'f Pptx {
        self.file
    }

    pub(super) fn item(&self) -> &Item {
        &self.item
    }

    /// The zero-based position of the slide the node sits on; `None` for
    /// the presentation.
    #[must_use]
    pub fn slide(&self) -> Option<usize> {
        match &self.item {
            Item::Presentation | Item::Master(_) | Item::Layout(_) => None,
            Item::Slide(slide) => Some(slide.position as usize),
            Item::Tree(tree) => tree.slide_position().map(|position| position as usize),
            Item::Paragraph(text) | Item::Run(text) | Item::LineBreak(text) => text
                .shape
                .slide_position()
                .map(|position| position as usize),
            Item::Row(table) | Item::Cell(table) => table
                .table
                .slide_position()
                .map(|position| position as usize),
            Item::Notes(notes) => Some(notes.slide as usize),
            Item::Comment(comment) => Some(comment.slide as usize),
            Item::ModernComment(comment) => Some(comment.slide as usize),
            Item::Reply(reply) => Some(reply.slide as usize),
        }
    }

    /// The leaf shape the node is or belongs to: the element of a shape-tree
    /// node, the shape of a paragraph or run, the table of a row or cell.
    #[must_use]
    pub fn shape(&self) -> Option<&Handle> {
        match &self.item {
            Item::Tree(tree) => Some(&tree.handle),
            Item::Paragraph(text) | Item::Run(text) | Item::LineBreak(text) => {
                Some(&text.shape.handle)
            },
            Item::Row(table) | Item::Cell(table) => Some(&table.table.handle),
            _ => None,
        }
    }

    /// The leaf paragraph of a paragraph node, or the paragraph a run or line
    /// break sits in.
    #[must_use]
    pub fn paragraph(&self) -> Option<&Paragraph> {
        match &self.item {
            Item::Paragraph(text) | Item::Run(text) | Item::LineBreak(text) => text.paragraph(),
            _ => None,
        }
    }

    /// The leaf run of a run or line-break node.
    #[must_use]
    pub fn run(&self) -> Option<&Run> {
        match &self.item {
            Item::Run(text) | Item::LineBreak(text) => text.run(),
            _ => None,
        }
    }

    /// The node's element kind.
    pub(super) fn element(&self) -> Kind {
        match &self.item {
            Item::Presentation => Kind::Presentation,
            Item::Slide(_) => Kind::Slide,
            Item::Master(_) => Kind::SlideMaster,
            Item::Layout(_) => Kind::SlideLayout,
            Item::Tree(tree) => tree.kind,
            Item::Paragraph(_) => Kind::Paragraph,
            Item::Run(_) => Kind::Run,
            Item::LineBreak(_) => Kind::LineBreak,
            Item::Row(_) => Kind::Row,
            Item::Cell(_) => Kind::Cell,
            Item::Notes(_) => Kind::Notes,
            Item::Comment(_) => Kind::Comment,
            Item::ModernComment(_) => Kind::ModernComment,
            Item::Reply(_) => Kind::Reply,
        }
    }

    /// The OfficeCLI `type` of the node.
    pub(super) fn type_name(&self) -> &'static str {
        match &self.item {
            Item::Presentation => "presentation",
            Item::Slide(_) => "slide",
            Item::Master(_) => "slideMaster",
            Item::Layout(_) => "slideLayout",
            Item::Tree(tree) => match tree.kind {
                Kind::Shape if tree.is_title() => "title",
                Kind::Shape if tree.handle.view().placeholder().is_some() => "placeholder",
                Kind::Shape => "shape",
                Kind::Picture => "picture",
                Kind::Table => "table",
                Kind::Chart => "chart",
                Kind::Group => "group",
                Kind::Connector => "connector",
                Kind::SmartArt => "smartart",
                Kind::Ole => "ole",
                _ => tree.kind.spec().name,
            },
            Item::Paragraph(_) => "paragraph",
            Item::Run(_) => "run",
            Item::LineBreak(_) => "linebreak",
            Item::Row(_) => "tr",
            Item::Cell(_) => "tc",
            Item::Notes(_) => "notes",
            Item::Comment(_) => "comment",
            Item::ModernComment(_) => "modernComment",
            Item::Reply(_) => "reply",
        }
    }

    /// The node's address steps: canonical (identities where they
    /// re-resolve) or positional.
    fn steps(&self, canonical: bool) -> Builder {
        match &self.item {
            Item::Presentation => Builder::new(),
            Item::Slide(slide) => {
                let mut out = Builder::new();
                out.push(steps::slide(slide.position));
                out
            },
            Item::Master(master) => {
                let mut out = Builder::new();
                out.push(PathStep::at(
                    "slidemaster",
                    Position::new(master.position as usize),
                ));
                out
            },
            Item::Layout(layout) => {
                let mut out = Builder::new();
                out.push(PathStep::at(
                    "slidemaster",
                    Position::new(layout.master as usize),
                ));
                out.push(PathStep::at(
                    "slidelayout",
                    Position::new(layout.position as usize),
                ));
                out
            },
            Item::Tree(tree) => tree.steps(canonical),
            Item::Paragraph(text) => {
                let mut out = text.shape.steps(canonical);
                out.push(at("paragraph", text.paragraph));
                out
            },
            Item::Run(text) => {
                let mut out = text.shape.steps(canonical);
                out.push(at("paragraph", text.paragraph));
                out.push(at("run", text.ordinal));
                out
            },
            Item::LineBreak(text) => {
                let mut out = text.shape.steps(canonical);
                out.push(at("paragraph", text.paragraph));
                out.push(at("linebreak", text.ordinal));
                out
            },
            Item::Row(table) => {
                let mut out = table.table.steps(canonical);
                out.push(at("tr", table.row));
                out
            },
            Item::Cell(table) => {
                let mut out = table.table.steps(canonical);
                out.push(at("tr", table.row));
                out.push(at("tc", table.column));
                out
            },
            Item::Notes(notes) => {
                let mut out = Builder::new();
                out.push(steps::slide(notes.slide));
                out.push(PathStep::new("notes"));
                out
            },
            Item::Comment(comment) => {
                let mut out = Builder::new();
                out.push(steps::slide(comment.slide));
                out.push(at("comment", comment.index));
                out
            },
            Item::ModernComment(comment) => {
                let mut out = Builder::new();
                out.push(steps::slide(comment.slide));
                out.push(at("moderncomment", comment.index));
                out
            },
            Item::Reply(reply) => {
                let mut out = Builder::new();
                out.push(steps::slide(reply.slide));
                out.push(at("moderncomment", reply.comment));
                out.push(at("reply", reply.index));
                out
            },
        }
    }

    /// The `text` member (and the query text of every node but a slide).
    pub(super) fn text_member(&self) -> Result<Option<Cow<'_, str>>> {
        Ok(match &self.item {
            Item::Presentation
            | Item::Slide(_)
            | Item::Master(_)
            | Item::Layout(_)
            | Item::Row(_) => None,
            // OfficeCLI's `GetShapeText` is "" for a shape without text.
            Item::Tree(tree) if tree.kind == Kind::Shape => {
                Some(Cow::Borrowed(tree.handle.view().text().unwrap_or("")))
            },
            Item::Tree(_) => None,
            Item::Paragraph(text) => text.paragraph().map(text::paragraph_text),
            Item::Run(text) => text.run().map(|run| Cow::Borrowed(run.text.as_str())),
            Item::LineBreak(_) => Some(Cow::Borrowed("")),
            Item::Cell(table) => Some(Cow::Owned(text::cell_text(table)?)),
            Item::Notes(notes) => Some(Cow::Borrowed(&notes.text)),
            Item::Comment(comment) => comment
                .comment()
                .map(|comment| Cow::Borrowed(comment.text())),
            Item::ModernComment(comment) => comment
                .comment()
                .and_then(|comment| comment.text().map(Cow::Borrowed)),
            Item::Reply(reply) => reply
                .reply()
                .and_then(|reply| reply.text().map(Cow::Borrowed)),
        })
    }

    /// The `preview` member (OfficeCLI's per-kind rule).
    fn preview_member(&self) -> Result<Option<Preview<'_>>> {
        Ok(match &self.item {
            Item::Slide(slide) if slide.listed => None,
            Item::Slide(slide) => Some(Preview::Text(Cow::Owned(slide_title(
                slide.scene(self.file)?,
            )))),
            Item::Master(master) if master.listed => None,
            Item::Master(master) => Some(Preview::Text(Cow::Borrowed(
                master.facts(self.file)?.name.as_deref().unwrap_or("master"),
            ))),
            Item::Layout(layout) if layout.listed => None,
            Item::Layout(layout) => Some(Preview::Text(Cow::Borrowed(
                layout.facts(self.file)?.name.as_deref().unwrap_or("layout"),
            ))),
            Item::Tree(tree) => Some(tree_preview(tree)?),
            Item::ModernComment(comment) => comment.comment().and_then(|comment| {
                comment
                    .text()
                    .map(|text| Preview::Prefix(Cow::Borrowed(text)))
            }),
            Item::Reply(reply) => reply.reply().and_then(|reply| {
                reply
                    .text()
                    .map(|text| Preview::Prefix(Cow::Borrowed(text)))
            }),
            _ => None,
        })
    }

    /// The `childCount` member: the children the view lists.
    fn child_count_member(&self) -> Result<Option<u32>> {
        let count = match &self.item {
            // The presentation view lists both catalog axes: slides first,
            // then slide masters.  Keep this in step with `Children::of`,
            // which transitions from `State::Slides` to `State::Masters`.
            Item::Presentation => self
                .file
                .slides_in_catalog()?
                .saturating_add(self.file.masters_in_catalog()?),
            Item::Slide(slide) if slide.listed => return Ok(None),
            Item::Slide(slide) => listed_tree(slide.scene(self.file)?.roots()),
            Item::Master(master) if master.listed => return Ok(None),
            // A master lists its layouts before its direct shape-tree
            // elements.  The selected-node count must include both phases.
            Item::Master(master) => self
                .file
                .layouts_in_catalog(master.position as usize)?
                .saturating_add(listed_tree(master.scene(self.file)?.roots())),
            Item::Layout(layout) if layout.listed => return Ok(None),
            Item::Layout(layout) => listed_tree(layout.scene(self.file)?.roots()),
            Item::Tree(tree) => match tree.kind {
                Kind::Shape => tree.body()?.map_or(0, |body| body.paragraphs.len()),
                Kind::Group => listed_tree(tree.handle.children()),
                Kind::Table => text::table_shape(&tree.handle)?.0,
                _ => 0,
            },
            Item::Paragraph(text) => text.paragraph().map_or(0, |paragraph| {
                paragraph
                    .runs
                    .iter()
                    .filter(|run| text::listed_run(run).is_some())
                    .count()
            }),
            Item::Row(table) => text::row_cells(&table.table.handle, table.row)?,
            Item::Run(_)
            | Item::LineBreak(_)
            | Item::Cell(_)
            | Item::Notes(_)
            | Item::Comment(_)
            | Item::Reply(_) => 0,
            Item::ModernComment(comment) => comment
                .comment()
                .map_or(0, |comment| comment.replies().len()),
        };
        Ok(Some(u32::try_from(count).unwrap_or(u32::MAX)))
    }

    /// The `format` member's values.
    pub(super) fn format_member(&self) -> Result<Format<'_>> {
        let (table, kind) = match &self.item {
            Item::Presentation => {
                let (width, height) = self.file.slide_size()?;
                let metadata = self.file.cached_metadata()?;
                (
                    PRESENTATION_PROPS,
                    PropKind::Root {
                        size: (width, height),
                        metadata,
                    },
                )
            },
            Item::Slide(slide) if slide.listed => (SLIDE_PROPS, PropKind::Empty),
            Item::Slide(slide) => (SLIDE_PROPS, PropKind::Slide(slide.facts(self.file)?)),
            Item::Master(master) if master.listed => {
                (Kind::SlideMaster.spec().props, PropKind::Empty)
            },
            Item::Master(master) => (
                Kind::SlideMaster.spec().props,
                PropKind::Master(master.facts(self.file)?),
            ),
            Item::Layout(layout) if layout.listed => {
                (Kind::SlideLayout.spec().props, PropKind::Empty)
            },
            Item::Layout(layout) => (
                Kind::SlideLayout.spec().props,
                PropKind::Layout(layout.facts(self.file)?),
            ),
            Item::Tree(tree) => match tree.kind {
                Kind::Shape => (
                    SHAPE_PROPS,
                    PropKind::Shape(Box::new(TreeProps::read(tree, true)?)),
                ),
                Kind::Table => {
                    let (rows, cols) = text::table_shape(&tree.handle)?;
                    let properties = text::table_properties(&tree.handle)?;
                    (
                        TABLE_PROPS,
                        PropKind::Table(Box::new(TableProps {
                            tree: TreeProps::read(tree, false)?,
                            rows,
                            cols,
                            properties,
                        })),
                    )
                },
                Kind::SmartArt => (
                    SMARTART_PROPS,
                    PropKind::SmartArt(Box::new(TreeProps::read(tree, false)?)),
                ),
                _ => (
                    FRAME_PROPS,
                    PropKind::Frame(Box::new(TreeProps::read(tree, false)?)),
                ),
            },
            Item::Paragraph(text) => match text.paragraph() {
                Some(paragraph) => (PARAGRAPH_PROPS, PropKind::Paragraph(paragraph)),
                None => (PARAGRAPH_PROPS, PropKind::Empty),
            },
            Item::Run(text) => match text.run() {
                Some(run) => (RUN_PROPS, PropKind::Run(&run.properties)),
                None => (RUN_PROPS, PropKind::Empty),
            },
            Item::Cell(cell) => (CELL_PROPS, PropKind::Cell(text::cell_fill(cell)?)),
            Item::LineBreak(_) | Item::Row(_) => (&[][..], PropKind::Empty),
            Item::Notes(notes) => (NOTES_PROPS, PropKind::Notes(&notes.text)),
            Item::Comment(comment) => match comment.comment() {
                Some(record) => (COMMENT_PROPS, PropKind::Comment(record)),
                None => (COMMENT_PROPS, PropKind::Empty),
            },
            Item::ModernComment(comment) => match comment.comment() {
                Some(record) => (
                    Kind::ModernComment.spec().props,
                    PropKind::ModernComment(record),
                ),
                None => (Kind::ModernComment.spec().props, PropKind::Empty),
            },
            Item::Reply(reply) => match reply.reply() {
                Some(record) => (Kind::Reply.spec().props, PropKind::Reply(record)),
                None => (Kind::Reply.spec().props, PropKind::Empty),
            },
        };
        Ok(Format::new(table, Source::Pptx(Props { kind })))
    }
}

/// `name[N]` for a zero-based position.
fn at(name: &'static str, position: u32) -> PathStep<'static> {
    PathStep::at(name, Position::new(position as usize))
}

/// The number of elements a container lists.
fn listed_tree(handles: Handles) -> usize {
    handles
        .filter(|handle| steps::tree_kind(handle).is_some())
        .count()
}

/// A slide's preview: the text of its first title placeholder among the
/// shape tree's direct shapes, else `(untitled)` (OfficeCLI).
pub(super) fn slide_title(scene: &SharedScene) -> String {
    scene
        .roots()
        .find(|handle| {
            matches!(handle.view(), Shape::Auto(_))
                && handle
                    .view()
                    .placeholder()
                    .and_then(|placeholder| placeholder.kind())
                    .is_some_and(|kind| matches!(kind.trim(), "title" | "ctrTitle"))
        })
        .map_or_else(
            || UNTITLED.to_owned(),
            |handle| handle.view().text().unwrap_or("").to_owned(),
        )
}

/// A shape-tree element's preview: a shape's text cut to 50 characters, or
/// its name; a table's name and size; any other element's name.
fn tree_preview(tree: &TreeItem) -> Result<Preview<'_>> {
    if tree.kind == Kind::Shape {
        if let Some(text) = tree.handle.view().text().filter(|text| !text.is_empty()) {
            return Ok(Preview::Prefix(Cow::Borrowed(text)));
        }
    }
    let name = element_name(tree);
    if tree.kind == Kind::Table {
        let (rows, cols) = text::table_shape(&tree.handle)?;
        return Ok(Preview::Text(Cow::Owned(format!("{name} ({rows}x{cols})"))));
    }
    Ok(Preview::Text(Cow::Borrowed(name)))
}

/// The element's `cNvPr` name, or OfficeCLI's default for its kind.
fn element_name(tree: &TreeItem) -> &str {
    tree.handle.view().name().unwrap_or(match tree.kind {
        Kind::Shape => "?",
        Kind::Picture => "Picture",
        Kind::Table => "Table",
        Kind::Chart => "Chart",
        Kind::Group => "Group",
        Kind::Connector => "Connector",
        Kind::SmartArt => "SmartArt",
        _ => "Object",
    })
}

impl<'f> Project<'f> for Node<'f> {
    fn kind(&self) -> &'static str {
        self.type_name()
    }

    fn file_format(&self) -> FileFormat {
        FileFormat::Pptx
    }

    fn path(&self) -> NodePath<'_> {
        NodePath::new(
            self.path
                .get_or_init(|| self.steps(true).finish())
                .as_slice(),
        )
    }

    fn positional_path(&self) -> NodePath<'_> {
        NodePath::new(
            self.positional
                .get_or_init(|| self.steps(false).finish().into_boxed()),
        )
    }

    fn text(&self) -> Result<Option<Cow<'_, str>>> {
        self.file.finish(|| self.text_member())
    }

    fn preview(&self) -> Result<Option<Preview<'_>>> {
        self.file.finish(|| self.preview_member())
    }

    fn child_count(&self, _: &ViewOptions) -> Result<Option<u32>> {
        self.file.finish(|| self.child_count_member())
    }

    fn format(&self) -> Result<Format<'_>> {
        self.file.finish(|| self.format_member())
    }

    fn children(&self, options: &ViewOptions) -> ViewChildren<'_, 'f> {
        // A child listed at depth 0 is a leaf of this view: a slide there
        // shows catalog data only (record 0010 M1-Δ1).
        let listing = Listing::View {
            leaves: options.depth <= 1,
        };
        ViewChildren::pptx(Children::of(
            self,
            crate::path::navigate::KindSet::ALL,
            listing,
        ))
    }
}

impl fmt::Debug for Node<'_> {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("Node")
            .field("type", &self.type_name())
            .field("slide", &self.slide())
            .finish_non_exhaustive()
    }
}

// ---------------------------------------------------------------------------
// The `format` member's values
// ---------------------------------------------------------------------------

/// Where a PPTX node's `format` values come from: the leaf values the node
/// holds, read by key index.
pub(crate) struct Props<'a> {
    kind: PropKind<'a>,
}

enum PropKind<'a> {
    Root {
        size: (i64, i64),
        metadata: Option<&'a Metadata>,
    },
    Slide(&'a SlideFacts),
    Master(&'a MasterFacts),
    Layout(&'a LayoutFacts),
    Shape(Box<TreeProps<'a>>),
    Frame(Box<TreeProps<'a>>),
    SmartArt(Box<TreeProps<'a>>),
    Table(Box<TableProps<'a>>),
    Paragraph(&'a Paragraph),
    Run(&'a RunProperties),
    Notes(&'a str),
    Cell(Option<[u8; 3]>),
    Comment(&'a Comment),
    ModernComment(&'a ModernComment),
    Reply(&'a ModernReply),
    Empty,
}

/// A shape-tree element's values. Fallible leaf properties are decoded while
/// constructing the node format, so a malformed transform, non-visual tag or
/// text body is returned as a typed read error instead of becoming an absent
/// ordinary property.
struct TreeProps<'a> {
    tree: &'a TreeItem,
    transform: Option<Transform>,
    description: Option<Cow<'a, str>>,
    text: Option<text::Facts<'a>>,
}

impl<'a> TreeProps<'a> {
    fn read(tree: &'a TreeItem, include_text: bool) -> Result<Self> {
        let shape = tree.handle.view();
        let transform = shape.transform().map_err(Error::from)?;
        let description = shape.description().map_err(Error::from)?;
        let text = if include_text {
            tree.body()?.map(|body| text::Facts::of(body.as_ref()))
        } else {
            None
        };
        Ok(Self {
            tree,
            transform,
            description,
            text,
        })
    }

    fn transform(&self) -> Option<&Transform> {
        self.transform.as_ref()
    }

    fn text(&self) -> Option<&text::Facts<'a>> {
        self.text.as_ref()
    }

    fn name(&self) -> Value<'a> {
        Value::Text(Cow::Borrowed(element_name(self.tree)))
    }

    fn alt(&self) -> Option<Value<'a>> {
        let description = self.description.as_ref()?;
        (!description.is_empty()).then_some(Value::Text(description.clone()))
    }

    fn x(&self) -> Option<Value<'a>> {
        match self.tree.handle.view().bounds() {
            Some(bounds) => Some(length(bounds.x())),
            None => self.transform()?.offset()?.x().as_emu().map(length),
        }
    }

    fn y(&self) -> Option<Value<'a>> {
        match self.tree.handle.view().bounds() {
            Some(bounds) => Some(length(bounds.y())),
            None => self.transform()?.offset()?.y().as_emu().map(length),
        }
    }

    fn width(&self) -> Option<Value<'a>> {
        match self.tree.handle.view().bounds() {
            Some(bounds) => Some(length(bounds.width())),
            None => Some(length(self.transform()?.extent()?.width().as_emu())),
        }
    }

    fn height(&self) -> Option<Value<'a>> {
        match self.tree.handle.view().bounds() {
            Some(bounds) => Some(length(bounds.height())),
            None => Some(length(self.transform()?.extent()?.height().as_emu())),
        }
    }

    fn rotation(&self) -> Option<Value<'a>> {
        let angle = self.transform()?.authored_rotation()?.value();
        (angle != 0).then(|| Value::Number(f64::from(angle) / 60_000.0))
    }

    fn flip_h(&self) -> Option<Value<'a>> {
        self.transform()?
            .authored_flip_horizontal()
            .filter(|flip| *flip)
            .map(Value::Bool)
    }

    fn flip_v(&self) -> Option<Value<'a>> {
        self.transform()?
            .authored_flip_vertical()
            .filter(|flip| *flip)
            .map(Value::Bool)
    }

    /// The 1-based stacking position among the shape tree's direct content
    /// elements; only for direct children of the tree (OfficeCLI).
    fn zorder(&self) -> Option<Value<'a>> {
        let handle = &self.tree.handle;
        if handle.parent().is_some() {
            return None;
        }
        let mut position = 0_i64;
        for root in handle.scene().roots() {
            if is_content(&root) {
                position += 1;
            }
            if root.position() == handle.position() {
                return Some(Value::Int(position));
            }
        }
        None
    }

    fn shape(&self, key: ShapeKey) -> Option<Value<'a>> {
        let tree = self.tree;
        match key {
            ShapeKey::Name => Some(self.name()),
            ShapeKey::PhType => tree.placeholder_type().map(Value::Token),
            ShapeKey::PhIndex | ShapeKey::Id => None,
            ShapeKey::Alt => self.alt(),
            ShapeKey::IsTitle => Some(Value::Bool(tree.is_title())),
            ShapeKey::X => self.x(),
            ShapeKey::Y => self.y(),
            ShapeKey::Width => self.width(),
            ShapeKey::Height => self.height(),
            ShapeKey::FlipH => self.flip_h(),
            ShapeKey::FlipV => self.flip_v(),
            ShapeKey::Zorder => self.zorder(),
            ShapeKey::Rotation => self.rotation(),
            ShapeKey::Font
            | ShapeKey::FontLatin
            | ShapeKey::FontEa
            | ShapeKey::FontCs
            | ShapeKey::Size
            | ShapeKey::Bold
            | ShapeKey::Italic
            | ShapeKey::Cap
            | ShapeKey::Underline
            | ShapeKey::Strike
            | ShapeKey::Highlight
            | ShapeKey::Spacing
            | ShapeKey::Color
            | ShapeKey::Align
            | ShapeKey::LineSpacing
            | ShapeKey::SpaceBefore
            | ShapeKey::SpaceAfter
            | ShapeKey::Indent
            | ShapeKey::MarginLeft
            | ShapeKey::MarginRight
            | ShapeKey::Direction => self.text()?.value(key),
        }
    }

    fn frame(&self, key: FrameKey) -> Option<Value<'a>> {
        match key {
            FrameKey::Name => Some(self.name()),
            FrameKey::Id => None,
            FrameKey::Alt => self.alt(),
            FrameKey::X => self.x(),
            FrameKey::Y => self.y(),
            FrameKey::Width => self.width(),
            FrameKey::Height => self.height(),
            FrameKey::Rotation => self.rotation(),
            FrameKey::FlipH => self.flip_h(),
            FrameKey::FlipV => self.flip_v(),
            FrameKey::Zorder => self.zorder(),
        }
    }

    fn smartart(&self, key: SmartArtKey) -> Option<Value<'a>> {
        match key {
            SmartArtKey::Name => Some(self.name()),
            SmartArtKey::Id => None,
            SmartArtKey::X => self.x(),
            SmartArtKey::Y => self.y(),
            SmartArtKey::Width => self.width(),
            SmartArtKey::Height => self.height(),
            SmartArtKey::Zorder => self.zorder(),
        }
    }
}

/// Whether a root element counts in OfficeCLI's z-order: a shape, picture,
/// graphic frame, group or connector.
fn is_content(handle: &Handle) -> bool {
    matches!(
        handle.view(),
        Shape::Auto(_)
            | Shape::Picture(_)
            | Shape::Table(_)
            | Shape::Chart(_)
            | Shape::Diagram(_)
            | Shape::Ole(_)
            | Shape::Frame(_)
            | Shape::Group(_)
            | Shape::Connector(_)
    )
}

/// A table's values: the frame's record and the table's dimensions and
/// style switches.
struct TableProps<'a> {
    tree: TreeProps<'a>,
    rows: usize,
    cols: usize,
    properties: Option<TableProperties>,
}

impl<'a> TableProps<'a> {
    fn value(&self, key: TableKey) -> Option<Value<'a>> {
        let switch = |on: bool| on.then_some(Value::Bool(true));
        let properties = self.properties.as_ref();
        match key {
            TableKey::Name => Some(self.tree.name()),
            TableKey::Id => None,
            TableKey::Rows => Some(count(self.rows)),
            TableKey::Cols => Some(count(self.cols)),
            // The table-style GUID is producer/native metadata. Keep it in
            // the leaf table model and expose it only through its raw API.
            TableKey::Style => None,
            // OfficeCLI always reports the header-row and banded-row switches
            // of a table with properties, and the others when they are on.
            TableKey::FirstRow => properties.map(|properties| Value::Bool(properties.first_row)),
            TableKey::BandedRows => properties.map(|properties| Value::Bool(properties.band_row)),
            TableKey::LastRow => switch(properties?.last_row),
            TableKey::FirstCol => switch(properties?.first_col),
            TableKey::LastCol => switch(properties?.last_col),
            TableKey::BandedCols => switch(properties?.band_col),
            TableKey::X => self.tree.x(),
            TableKey::Y => self.tree.y(),
            TableKey::Width => self.tree.width(),
            TableKey::Height => self.tree.height(),
            TableKey::Zorder => self.tree.zorder(),
        }
    }
}

fn count(value: usize) -> Value<'static> {
    Value::Int(i64::try_from(value).unwrap_or(i64::MAX))
}

fn length<'a>(emu: i64) -> Value<'a> {
    Value::Length(Length::from_emu(emu))
}

/// The slide-size preset OfficeCLI derives from the aspect ratio when the
/// size has no `type`; Litchi derives it from the ratio always.
fn slide_size(width: i64, height: i64) -> Option<&'static str> {
    if width <= 0 || height <= 0 {
        return None;
    }
    #[allow(
        clippy::cast_precision_loss,
        reason = "slide sizes are far below 2^53 EMU"
    )]
    let ratio = width as f64 / height as f64;
    let token = if (ratio - 16.0 / 9.0).abs() < 0.02 {
        "widescreen"
    } else if (ratio - 4.0 / 3.0).abs() < 0.02 {
        "standard"
    } else if (ratio - 16.0 / 10.0).abs() < 0.02 {
        "16:10"
    } else {
        "custom"
    };
    SLIDE_SIZE.get(token)
}

impl<'a> Props<'a> {
    /// The value of the property at `index` of the node's table, if set.
    pub(crate) fn get(&self, index: usize) -> Option<Value<'a>> {
        match &self.kind {
            PropKind::Root { size, metadata } => {
                let count = |value: Option<u32>| value.map(|value| Value::Int(i64::from(value)));
                let text = |value: Option<&'a String>| {
                    value
                        .filter(|value| !value.is_empty())
                        .map(|value| Value::Text(Cow::Borrowed(value.as_str())))
                };
                match RootKey::at(index)? {
                    RootKey::SlideWidth => Some(length(size.0)),
                    RootKey::SlideHeight => Some(length(size.1)),
                    RootKey::SlideSize => slide_size(size.0, size.1).map(Value::Token),
                    RootKey::Title => text(metadata.and_then(|m| m.title.as_ref())),
                    RootKey::Author => text(metadata.and_then(|m| m.author.as_ref())),
                    RootKey::Subject => text(metadata.and_then(|m| m.subject.as_ref())),
                    RootKey::Keywords => text(metadata.and_then(|m| m.keywords.as_ref())),
                    RootKey::Description => text(metadata.and_then(|m| m.description.as_ref())),
                    RootKey::LastModifiedBy => {
                        text(metadata.and_then(|m| m.last_modified_by.as_ref()))
                    },
                    RootKey::RevisionNumber => text(metadata.and_then(|m| m.revision.as_ref())),
                    RootKey::Category => text(metadata.and_then(|m| m.category.as_ref())),
                    RootKey::Created => metadata
                        .and_then(|m| m.created)
                        .map(|at| Value::Text(Cow::Owned(crate::format::metadata::round_trip(at)))),
                    RootKey::Modified => metadata
                        .and_then(|m| m.modified)
                        .map(|at| Value::Text(Cow::Owned(crate::format::metadata::round_trip(at)))),
                    RootKey::ExtendedTemplate => text(metadata.and_then(|m| m.template.as_ref())),
                    RootKey::ExtendedManager => text(metadata.and_then(|m| m.manager.as_ref())),
                    RootKey::ExtendedCompany => text(metadata.and_then(|m| m.company.as_ref())),
                    RootKey::ExtendedApplication => {
                        text(metadata.and_then(|m| m.application.as_ref()))
                    },
                    RootKey::ExtendedPages => count(metadata.and_then(|m| m.page_count)),
                    RootKey::ExtendedWords => count(metadata.and_then(|m| m.word_count)),
                    RootKey::ExtendedCharacters => count(metadata.and_then(|m| m.character_count)),
                    RootKey::ExtendedTotalTime => {
                        count(metadata.and_then(|m| m.editing_time_minutes))
                    },
                }
            },
            PropKind::Slide(facts) => match SlideKey::at(index)? {
                SlideKey::Layout => facts
                    .layout
                    .as_deref()
                    .map(|name| Value::Text(Cow::Borrowed(name))),
                SlideKey::LayoutType => facts
                    .layout_type
                    .as_deref()
                    .map(|kind| Value::Text(Cow::Borrowed(kind))),
                SlideKey::Name => facts
                    .name
                    .as_deref()
                    .map(|name| Value::Text(Cow::Borrowed(name))),
                // OfficeCLI reports a hidden slide only.
                SlideKey::Hidden => facts.hidden.then_some(Value::Bool(true)),
            },
            PropKind::Shape(props) => props.shape(ShapeKey::at(index)?),
            PropKind::Frame(props) => props.frame(FrameKey::at(index)?),
            PropKind::SmartArt(props) => props.smartart(SmartArtKey::at(index)?),
            PropKind::Table(props) => props.value(TableKey::at(index)?),
            PropKind::Paragraph(paragraph) => {
                text::paragraph_value(paragraph, ParagraphKey::at(index)?)
            },
            PropKind::Run(properties) => text::run_value(properties, RunKey::at(index)?),
            PropKind::Cell(fill) => match CellKey::at(index)? {
                CellKey::Fill => fill.map(|[red, green, blue]| {
                    Value::Color(crate::value::Color::rgb(red, green, blue))
                }),
            },
            PropKind::Notes(notes) => match NotesKey::at(index)? {
                NotesKey::Text => Some(Value::Text(Cow::Borrowed(notes))),
            },
            PropKind::Comment(comment) => match CommentKey::at(index)? {
                CommentKey::Text => Some(Value::Text(Cow::Borrowed(comment.text()))),
                CommentKey::Author => comment
                    .author_name()
                    .map(|name| Value::Text(Cow::Borrowed(name))),
                CommentKey::Initials => comment
                    .initials()
                    .map(|initials| Value::Text(Cow::Borrowed(initials))),
                // The producer comment index is native metadata. It remains
                // available through the related raw surface only.
                CommentKey::Index => None,
                CommentKey::Date => comment
                    .date_time()
                    .map(|date| Value::Text(Cow::Borrowed(date))),
                CommentKey::X => Some(length(comment.location().0)),
                CommentKey::Y => Some(length(comment.location().1)),
            },
            PropKind::Master(facts) => master_value(facts, index),
            PropKind::Layout(facts) => layout_value(facts, index),
            PropKind::ModernComment(comment) => modern_comment_value(comment, index),
            PropKind::Reply(reply) => reply_value(reply, index),
            PropKind::Empty => None,
        }
    }
}

fn master_value(facts: &MasterFacts, index: usize) -> Option<Value<'_>> {
    let key = Kind::SlideMaster.spec().props.get(index)?.name;
    match key {
        "name" => facts
            .name
            .as_deref()
            .filter(|name| !name.is_empty())
            .map(|name| Value::Text(Cow::Borrowed(name))),
        "preserved" | "isPreserved" => facts.preserved.then_some(Value::Bool(true)),
        _ => None,
    }
}

fn layout_value(facts: &LayoutFacts, index: usize) -> Option<Value<'_>> {
    let key = Kind::SlideLayout.spec().props.get(index)?.name;
    match key {
        "name" => facts
            .name
            .as_deref()
            .filter(|name| !name.is_empty())
            .map(|name| Value::Text(Cow::Borrowed(name))),
        "type" | "layoutType" => facts.kind.map(|kind| Value::Token(kind.as_str())),
        _ => None,
    }
}

fn modern_comment_value(comment: &ModernComment, index: usize) -> Option<Value<'_>> {
    let key = Kind::ModernComment.spec().props.get(index)?.name;
    match key {
        "text" => comment.text().map(|text| Value::Text(Cow::Borrowed(text))),
        "author" | "authorName" => comment
            .author_name()
            .map(|name| Value::Text(Cow::Borrowed(name))),
        "created" | "date" => Some(Value::Text(Cow::Borrowed(comment.created()))),
        "startDate" => comment
            .start_date()
            .map(|date| Value::Text(Cow::Borrowed(date))),
        "dueDate" => comment
            .due_date()
            .map(|date| Value::Text(Cow::Borrowed(date))),
        "title" => comment
            .title()
            .map(|title| Value::Text(Cow::Borrowed(title))),
        "status" => Some(Value::Token(status_token(comment.status()))),
        "progress" => Value::number(f64::from(comment.progress().thousandths()) / 1_000.0),
        "x" => comment.location().map(|(x, _)| length(x)),
        "y" => comment.location().map(|(_, y)| length(y)),
        _ => None,
    }
}

fn reply_value(reply: &ModernReply, index: usize) -> Option<Value<'_>> {
    let key = Kind::Reply.spec().props.get(index)?.name;
    match key {
        "text" => reply.text().map(|text| Value::Text(Cow::Borrowed(text))),
        "author" | "authorName" => reply
            .author_name()
            .map(|name| Value::Text(Cow::Borrowed(name))),
        "created" | "date" => Some(Value::Text(Cow::Borrowed(reply.created()))),
        "status" => Some(Value::Token(status_token(reply.status()))),
        _ => None,
    }
}

fn status_token(status: ModernStatus) -> &'static str {
    match status {
        ModernStatus::Active => "active",
        ModernStatus::Resolved => "resolved",
        ModernStatus::Closed => "closed",
    }
}

impl fmt::Debug for Props<'_> {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter.debug_struct("Props").finish_non_exhaustive()
    }
}

// Nodes cross threads (record 0010 §3.3).
const _: () = {
    const fn assert_send_sync<T: Send + Sync>() {}
    assert_send_sync::<Node<'static>>();
};
