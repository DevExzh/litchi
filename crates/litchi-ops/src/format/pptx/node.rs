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
    AXIS_PROPS, AxisKey, CELL_PROPS, COLUMN_PROPS, COMMENT_PROPS, CellKey, ChartKey, CommentKey,
    EQUATION_PROPS, FRAME_PROPS, FrameKey, HYPERLINK_PROPS, Kind, LINKED_FRAME_PROPS,
    LinkedFrameKey, MEDIA_PROPS, MediaKey, NODE_PROPS, NOTES_PROPS, NodeKey, NotesKey,
    PARAGRAPH_PROPS, PLACEHOLDER_TYPE, PRESENTATION_PROPS, ParagraphKey, ROW_PROPS, RUN_PROPS,
    RootKey, RowKey, RunKey, SERIES_PROPS, SHAPE_PROPS, SLIDE_PROPS, SLIDE_SIZE, SMARTART_PROPS,
    SeriesKey, ShapeKey, SlideKey, SmartArtKey, TABLE_PROPS, THEME_PROPS, TableKey,
};
use super::steps::{self, Builder};
use super::text;
use crate::error::{Error, Result};
use crate::format::chart::Summary;
use crate::node::{
    Children as ViewChildren, Format, NodePath, PathStep, Preview, Project, Source, Value,
    ViewOptions,
};
use crate::value::Length;
use litchi_drawingml::diagram::edit::Outline;

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
    /// The paths of an address of more than four steps, written on first
    /// use; shorter paths are written on each request (record 0262).
    deep: OnceLock<Box<DeepPaths>>,
}

/// The paths of a deep address, each written on first use.
#[derive(Clone, Default)]
struct DeepPaths {
    canonical: OnceLock<Box<[PathStep<'static>]>>,
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
    Series(ChartItem),
    Axis(ChartItem),
    /// A SmartArt content node (record 0284).
    DiagramNode(DiagramItem),
    Notes(NotesItem),
    Comment(CommentItem),
    ModernComment(ModernCommentItem),
    Reply(ReplyItem),
    /// The presentation theme (record 0321).
    Theme(super::theme::ThemeItem),
    /// A table grid column (record 0321). Boxed, as are the other two
    /// record 0321 elements, so the item stays as large as a text item and
    /// every query that never names them moves no more bytes per node
    /// (record 0326).
    Column(Box<super::columns::ColumnItem>),
    /// A shape's or run's click action (record 0321).
    Hyperlink(Box<super::links::LinkItem>),
    /// An Office Math equation (record 0321).
    Equation(Box<super::equations::EquationItem>),
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
    Layout {
        master: u32,
        position: u32,
    },
    /// The notes page of a slide: the container of the notes body shape a
    /// notes paragraph or run belongs to (record 0333).
    Notes(u32),
}

impl TreeContainer {
    /// The slide whose shape tree holds the element; `None` for masters,
    /// layouts and notes pages, so slide-shape writes refuse them.
    pub(super) const fn slide(self) -> Option<u32> {
        match self {
            Self::Slide(position) => Some(position),
            Self::Master(_) | Self::Layout { .. } | Self::Notes(_) => None,
        }
    }

    /// The slide of a notes page.
    pub(super) const fn notes(self) -> Option<u32> {
        match self {
            Self::Notes(slide) => Some(slide),
            Self::Slide(_) | Self::Master(_) | Self::Layout { .. } => None,
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
    /// A chart frame's chart part (record 0270) or a SmartArt frame's data
    /// model (record 0284), read on first use. One cell for both, since a
    /// frame is one kind or the other: every node of a query is moved
    /// several times, and a second cell made each move 24 bytes longer
    /// (record 0326).
    frame: OnceLock<std::result::Result<Option<FramePart>, Error>>,
}

/// What a chart or SmartArt frame's part reads as.
#[derive(Clone)]
enum FramePart {
    Chart(Arc<Summary>),
    Diagram(Arc<Outline>),
    /// A table frame's cell count per row (record 0333), so resolving a
    /// row and then a cell reads the table once.
    Table(Arc<Vec<u32>>),
}

impl TreeItem {
    pub(super) fn new(slide: u32, kind: Kind, handle: Handle, ordinal: u32) -> Self {
        Self {
            container: TreeContainer::Slide(slide),
            kind,
            handle,
            ordinal,
            body: OnceLock::new(),
            frame: OnceLock::new(),
        }
    }

    pub(super) fn new_master(position: u32, kind: Kind, handle: Handle, ordinal: u32) -> Self {
        Self {
            container: TreeContainer::Master(position),
            kind,
            handle,
            ordinal,
            body: OnceLock::new(),
            frame: OnceLock::new(),
        }
    }

    /// A shape-tree element of any container.
    pub(super) fn in_container(
        container: TreeContainer,
        kind: Kind,
        handle: Handle,
        ordinal: u32,
    ) -> Self {
        Self {
            container,
            kind,
            handle,
            ordinal,
            body: OnceLock::new(),
            frame: OnceLock::new(),
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
            frame: OnceLock::new(),
        }
    }

    /// A table frame's cell count per row, read on first use; `None` for
    /// another element.
    pub(super) fn table_grid(&self) -> Result<Option<Arc<Vec<u32>>>> {
        if self.kind != Kind::Table {
            return Ok(None);
        }
        let read = self.frame.get_or_init(|| {
            let table = self.handle.table().map_err(Error::from)?;
            let counts = table.row_cell_counts().map_err(Error::from)?;
            Ok(Some(FramePart::Table(Arc::new(
                counts
                    .into_iter()
                    .map(|count| u32::try_from(count).unwrap_or(u32::MAX))
                    .collect(),
            ))))
        });
        match read {
            Ok(Some(FramePart::Table(grid))) => Ok(Some(Arc::clone(grid))),
            Ok(_) => Ok(None),
            Err(error) => Err(error.clone()),
        }
    }

    pub(super) const fn slide_position(&self) -> Option<u32> {
        match self.container {
            TreeContainer::Notes(slide) => Some(slide),
            container => container.slide(),
        }
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

    /// A chart frame's chart, read from its chart part on first use; `None`
    /// for another element, a master or layout chart, or a frame that names
    /// no chart part.
    pub(super) fn chart(&self, file: &Pptx) -> Result<Option<Arc<Summary>>> {
        if self.kind != Kind::Chart {
            return Ok(None);
        }
        let read = self.frame.get_or_init(|| {
            super::chart::summary(file, self).map(|chart| chart.map(FramePart::Chart))
        });
        match read {
            Ok(Some(FramePart::Chart(chart))) => Ok(Some(Arc::clone(chart))),
            Ok(_) => Ok(None),
            Err(error) => Err(error.clone()),
        }
    }

    /// A SmartArt frame's data-model outline, read from its data part on
    /// first use; `None` for another element, a master or layout SmartArt,
    /// or a frame that names no data part.
    pub(super) fn diagram(&self, file: &Pptx) -> Result<Option<Arc<Outline>>> {
        if self.kind != Kind::SmartArt {
            return Ok(None);
        }
        let read = self.frame.get_or_init(|| {
            super::smartart::outline(file, self).map(|outline| outline.map(FramePart::Diagram))
        });
        match read {
            Ok(Some(FramePart::Diagram(outline))) => Ok(Some(Arc::clone(outline))),
            Ok(_) => Ok(None),
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
            // A notes body is addressed through the notes node: its text
            // nodes are `/slide[N]/notes/paragraph[K]` (record 0333).
            TreeContainer::Notes(slide) => {
                out.push(steps::slide(slide));
                out.push(PathStep::new("notes"));
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
    /// The shape whose body holds the text: a slide shape, the table frame
    /// of a cell's text ([`Self::cell`]), or a notes page's body shape (a
    /// [`TreeContainer::Notes`] container).
    pub(super) shape: TreeItem,
    pub(super) body: Arc<Body>,
    pub(super) paragraph: u32,
    /// The index in the paragraph's runs (runs and line breaks).
    pub(super) run: u32,
    /// The position among the paragraph's runs, or among its line breaks.
    pub(super) ordinal: u32,
    /// The table cell whose body holds the text; [`CellRef::NONE`] for a
    /// shape's or notes page's text (record 0333). It fills the padding
    /// after the three positions, so the item keeps its size.
    pub(super) cell: CellRef,
}

/// A table cell's row and column, packed into one word: rows and columns
/// up to 65,534 are addressable (`PowerPoint` tables hold 75 of each).
#[derive(Clone, Copy, PartialEq, Eq, Debug)]
pub(super) struct CellRef(u32);

impl CellRef {
    /// No cell: the text of a shape or notes page.
    pub(super) const NONE: Self = Self(u32::MAX);

    /// The cell at `row` and `column`; `None` past the packed range.
    pub(super) const fn new(row: u32, column: u32) -> Option<Self> {
        if row < 0xFFFF && column < 0xFFFF {
            Some(Self((row << 16) | column))
        } else {
            None
        }
    }

    /// The cell's row and column.
    pub(super) const fn get(self) -> Option<(u32, u32)> {
        if self.0 == u32::MAX {
            None
        } else {
            Some((self.0 >> 16, self.0 & 0xFFFF))
        }
    }
}

impl TextItem {
    /// The steps of the node whose body holds the text: the shape, the
    /// table cell, or the notes node.
    pub(super) fn owner_steps(&self, canonical: bool) -> Builder {
        let mut out = self.shape.steps(canonical);
        if let Some((row, column)) = self.cell.get() {
            out.push(PathStep::at("tr", Position::new(row as usize)));
            out.push(PathStep::at("tc", Position::new(column as usize)));
        }
        out
    }

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

/// A series or axis of a chart: the chart frame, its chart summary and the
/// position in the summary's series or axes (record 0270).
#[derive(Clone)]
pub(super) struct ChartItem {
    pub(super) chart: TreeItem,
    pub(super) summary: Arc<Summary>,
    pub(super) index: u32,
}

/// A content node of a SmartArt: the frame, its outline and the node's
/// index in the outline's entries (record 0284).
#[derive(Clone)]
pub(super) struct DiagramItem {
    pub(super) smartart: TreeItem,
    pub(super) outline: Arc<Outline>,
    pub(super) index: u32,
}

impl DiagramItem {
    pub(super) fn entry(&self) -> Option<&litchi_drawingml::diagram::edit::Entry> {
        self.outline.entries.get(self.index as usize)
    }

    /// The node's position among its siblings, then its ancestors'.
    pub(super) fn positions(&self) -> Vec<u32> {
        let mut out = Vec::new();
        let mut current = Some(self.index as usize);
        while let Some(index) = current {
            let parent = self
                .outline
                .entries
                .get(index)
                .and_then(|entry| entry.parent);
            let position = self.outline.entries[..index]
                .iter()
                .filter(|entry| entry.parent == parent)
                .count();
            out.push(u32::try_from(position).unwrap_or(u32::MAX));
            current = parent;
        }
        out.reverse();
        out
    }
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
            deep: OnceLock::new(),
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
            Item::Presentation | Item::Master(_) | Item::Layout(_) | Item::Theme(_) => None,
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
            Item::Column(column) => column
                .table
                .slide_position()
                .map(|position| position as usize),
            Item::Hyperlink(item) => link_shape(item)
                .slide_position()
                .map(|position| position as usize),
            Item::Equation(item) => Some(item.slide as usize),
            Item::Series(item) | Item::Axis(item) => item
                .chart
                .slide_position()
                .map(|position| position as usize),
            Item::DiagramNode(item) => item
                .smartart
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
            Item::Column(column) => Some(&column.table.handle),
            Item::Hyperlink(item) => Some(&link_shape(item).handle),
            Item::Equation(item) => item.shape.as_ref().map(|tree| &tree.handle),
            Item::Series(item) | Item::Axis(item) => Some(&item.chart.handle),
            Item::DiagramNode(item) => Some(&item.smartart.handle),
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
            Item::Series(_) => Kind::Series,
            Item::DiagramNode(_) => Kind::SmartArtNode,
            Item::Axis(_) => Kind::Axis,
            Item::Notes(_) => Kind::Notes,
            Item::Comment(_) => Kind::Comment,
            Item::ModernComment(_) => Kind::ModernComment,
            Item::Reply(_) => Kind::Reply,
            Item::Theme(_) => Kind::Theme,
            Item::Column(_) => Kind::Col,
            Item::Hyperlink(_) => Kind::Hyperlink,
            Item::Equation(_) => Kind::Equation,
        }
    }

    /// Whether the node is a `p:sp` with a placeholder binding: what a bare
    /// `placeholder` selector matches.
    pub(super) fn is_placeholder(&self) -> bool {
        matches!(&self.item, Item::Tree(tree)
            if tree.kind == Kind::Shape && tree.handle.view().placeholder().is_some())
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
            Item::Series(_) => "series",
            Item::DiagramNode(_) => "node",
            Item::Axis(_) => "axis",
            Item::Notes(_) => "notes",
            Item::Comment(_) => "comment",
            Item::ModernComment(_) => "modernComment",
            Item::Reply(_) => "reply",
            Item::Theme(_) => "theme",
            Item::Column(_) => "col",
            Item::Hyperlink(_) => "hyperlink",
            Item::Equation(_) => "equation",
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
                let mut out = text.owner_steps(canonical);
                out.push(at("paragraph", text.paragraph));
                out
            },
            Item::Run(text) => {
                let mut out = text.owner_steps(canonical);
                out.push(at("paragraph", text.paragraph));
                out.push(at("run", text.ordinal));
                out
            },
            Item::LineBreak(text) => {
                let mut out = text.owner_steps(canonical);
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
            Item::Series(item) => {
                let mut out = item.chart.steps(canonical);
                out.push(at("series", item.index));
                out
            },
            Item::Axis(item) => {
                let mut out = item.chart.steps(canonical);
                out.push(at("axis", item.index));
                out
            },
            Item::DiagramNode(item) => {
                let mut out = item.smartart.steps(canonical);
                for position in item.positions() {
                    out.push(at("node", position));
                }
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
            Item::Theme(_) => {
                let mut out = Builder::new();
                out.push(PathStep::new("theme"));
                out
            },
            Item::Column(column) => {
                let mut out = column.table.steps(canonical);
                out.push(at("col", column.index));
                out
            },
            Item::Hyperlink(item) => {
                let mut out = match &item.owner {
                    super::links::Owner::Shape(tree) => tree.steps(canonical),
                    super::links::Owner::Run(text) => {
                        let mut out = text.owner_steps(canonical);
                        out.push(at("paragraph", text.paragraph));
                        out.push(at("run", text.ordinal));
                        out
                    },
                };
                out.push(PathStep::new("hyperlink"));
                out
            },
            Item::Equation(item) => {
                let mut out = match &item.shape {
                    Some(tree) => tree.steps(canonical),
                    None => {
                        let mut out = Builder::new();
                        out.push(steps::slide(item.slide));
                        out
                    },
                };
                out.push(at("equation", item.index));
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
            | Item::Row(_)
            | Item::Series(_)
            | Item::Axis(_)
            | Item::Theme(_)
            | Item::Column(_) => None,
            Item::Hyperlink(item) => Some(super::links::link_text(&item.link)),
            Item::Equation(item) => Some(Cow::Borrowed(item.equation.text())),
            // OfficeCLI's `GetShapeText` is "" for a shape without text.
            Item::Tree(tree) if tree.kind == Kind::Shape => {
                Some(Cow::Borrowed(tree.handle.view().text().unwrap_or("")))
            },
            Item::Tree(tree) if tree.kind == Kind::SmartArt => tree
                .diagram(self.file)?
                .map(|outline| Cow::Owned(super::smartart::outline_text(&outline))),
            Item::Tree(_) => None,
            Item::DiagramNode(item) => item.entry().map(|entry| Cow::Owned(entry.text.clone())),
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
                Kind::Chart => tree
                    .chart(self.file)?
                    .map_or(0, |chart| chart.series.len() + chart.axes.len()),
                Kind::SmartArt => tree
                    .diagram(self.file)?
                    .map_or(0, |outline| outline.children(None).count()),
                _ => 0,
            },
            Item::DiagramNode(item) => item.outline.children(Some(item.index as usize)).count(),
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
            | Item::Series(_)
            | Item::Axis(_)
            | Item::Notes(_)
            | Item::Comment(_)
            | Item::Reply(_)
            | Item::Theme(_)
            | Item::Column(_)
            | Item::Hyperlink(_)
            | Item::Equation(_) => 0,
            Item::ModernComment(comment) => comment
                .comment()
                .map_or(0, |comment| comment.replies().len()),
        };
        Ok(Some(u32::try_from(count).unwrap_or(u32::MAX)))
    }

    /// The value at `index` of the node's property table: what
    /// [`Self::format_member`] gives there. A shape-tree element other than
    /// a table reads only that key's inputs (its transform, alternative
    /// text or text body, each when the key needs it), so a geometry filter
    /// does not decode the text body (record 0262); a read error of an input
    /// the key does not use is not reported here, but by `format`.
    pub(super) fn prop_member(&self, index: usize) -> Result<Option<Value<'_>>> {
        let Item::Tree(tree) = &self.item else {
            return Ok(self.format_member()?.get(index));
        };
        Ok(match tree.kind {
            Kind::Shape => match ShapeKey::at(index) {
                Some(key @ (ShapeKey::Link | ShapeKey::Tooltip)) => super::links::click_value(
                    super::links::of_shape(self.file, tree)?
                        .as_ref()
                        .map(|item| &*item.link),
                    key == ShapeKey::Tooltip,
                ),
                Some(key) => TreeProps::read_for(tree, Needs::shape(tree, key))?.shape(key),
                None => None,
            },
            Kind::Picture | Kind::Group => match LinkedFrameKey::at(index) {
                Some(key @ (LinkedFrameKey::Link | LinkedFrameKey::Tooltip)) => {
                    super::links::click_value(
                        super::links::of_frame(self.file, tree)?
                            .as_ref()
                            .map(|item| &*item.link),
                        key == LinkedFrameKey::Tooltip,
                    )
                },
                Some(_) => match FrameKey::at(index) {
                    Some(key) => TreeProps::read_for(tree, Needs::frame(tree, key))?.frame(key),
                    None => None,
                },
                None => None,
            },
            Kind::Table | Kind::Chart | Kind::Video | Kind::Audio | Kind::Media => {
                self.format_member()?.get(index)
            },
            Kind::SmartArt => match SmartArtKey::at(index) {
                Some(SmartArtKey::Layout | SmartArtKey::LayoutId | SmartArtKey::NodeCount) => {
                    self.format_member()?.get(index)
                },
                Some(key) => TreeProps::read_for(tree, Needs::smartart(tree, key))?.smartart(key),
                None => None,
            },
            _ => match FrameKey::at(index) {
                Some(key) => TreeProps::read_for(tree, Needs::frame(tree, key))?.frame(key),
                None => None,
            },
        })
    }

    /// The canonical or positional path: written now and carried by value
    /// when it has at most four steps, else written once and kept.
    fn address(&self, canonical: bool) -> NodePath<'_> {
        fn kept(deep: &DeepPaths, canonical: bool) -> &OnceLock<Box<[PathStep<'static>]>> {
            if canonical {
                &deep.canonical
            } else {
                &deep.positional
            }
        }
        if let Some(steps) = self.deep.get().and_then(|deep| kept(deep, canonical).get()) {
            return NodePath::new(steps);
        }
        let steps = self.steps(canonical).finish();
        if let Some(path) = NodePath::inline(steps.as_slice()) {
            return path;
        }
        let deep = self.deep.get_or_init(Box::default);
        NodePath::new(kept(deep, canonical).get_or_init(|| steps.into_boxed()))
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
                Kind::Shape => {
                    let mut props = TreeProps::read(tree, true)?;
                    props.link = super::links::of_shape(self.file, tree)?.map(|item| item.link);
                    (SHAPE_PROPS, PropKind::Shape(Box::new(props)))
                },
                Kind::Picture | Kind::Group => {
                    let mut props = TreeProps::read(tree, false)?;
                    props.link = super::links::of_frame(self.file, tree)?.map(|item| item.link);
                    (LINKED_FRAME_PROPS, PropKind::LinkedFrame(Box::new(props)))
                },
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
                Kind::Video | Kind::Audio | Kind::Media => (
                    MEDIA_PROPS,
                    PropKind::Media(Box::new(MediaProps {
                        tree: TreeProps::read(tree, false)?,
                        trim: (
                            super::media::trim(tree, true)?,
                            super::media::trim(tree, false)?,
                        ),
                    })),
                ),
                Kind::SmartArt => (
                    SMARTART_PROPS,
                    PropKind::SmartArt(Box::new(SmartArtProps {
                        tree: TreeProps::read(tree, false)?,
                        outline: tree.diagram(self.file)?,
                    })),
                ),
                Kind::Chart => (
                    super::spec::CHART_CAPABILITY_PROPS,
                    PropKind::Chart(Box::new(ChartProps {
                        tree: TreeProps::read(tree, false)?,
                        chart: tree.chart(self.file)?,
                    })),
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
                Some(run) => (
                    RUN_PROPS,
                    PropKind::Run(
                        &run.properties,
                        super::links::of_run(self.file, text)?.map(|item| item.link),
                    ),
                ),
                None => (RUN_PROPS, PropKind::Empty),
            },
            Item::Cell(cell) => (CELL_PROPS, PropKind::Cell(text::cell_properties(cell)?)),
            Item::Row(row) => (ROW_PROPS, PropKind::Row(text::row_height(row)?)),
            Item::LineBreak(_) => (&[][..], PropKind::Empty),
            Item::Series(item) => (SERIES_PROPS, PropKind::Series(item)),
            Item::DiagramNode(item) => (NODE_PROPS, PropKind::DiagramNode(item)),
            Item::Axis(item) => (AXIS_PROPS, PropKind::Axis(item)),
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
                Some(record) => (
                    Kind::Reply.spec().props,
                    PropKind::Reply(record, reply.slide, reply.comment),
                ),
                None => (Kind::Reply.spec().props, PropKind::Empty),
            },
            Item::Theme(item) => (THEME_PROPS, PropKind::Theme(&item.theme)),
            Item::Column(item) => (COLUMN_PROPS, PropKind::Column(item)),
            Item::Hyperlink(item) => (HYPERLINK_PROPS, PropKind::Hyperlink(item)),
            Item::Equation(item) => (EQUATION_PROPS, PropKind::Equation(item)),
        };
        Ok(Format::new(table, Source::Pptx(Props { kind })))
    }
}

/// A media picture's values: its frame's, then its playback trim.
struct MediaProps<'a> {
    tree: TreeProps<'a>,
    trim: (Option<String>, Option<String>),
}

/// The shape a hyperlink belongs to.
fn link_shape(item: &super::links::LinkItem) -> &TreeItem {
    match &item.owner {
        super::links::Owner::Shape(tree) => tree,
        super::links::Owner::Run(text) => &text.shape,
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
        self.address(true)
    }

    fn positional_path(&self) -> NodePath<'_> {
        self.address(false)
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
    /// A picture or a group (record 0325).
    LinkedFrame(Box<TreeProps<'a>>),
    SmartArt(Box<SmartArtProps<'a>>),
    DiagramNode(&'a DiagramItem),
    Chart(Box<ChartProps<'a>>),
    Series(&'a ChartItem),
    Axis(&'a ChartItem),
    Table(Box<TableProps<'a>>),
    Paragraph(&'a Paragraph),
    Run(
        &'a RunProperties,
        Option<Arc<litchi_pptx::hyperlinks::Link>>,
    ),
    Notes(&'a str),
    Cell(litchi_pptx::table::CellProperties),
    /// A table row's height (record 0333).
    Row(Option<i64>),
    Comment(&'a Comment),
    ModernComment(&'a ModernComment),
    /// A reply, its slide and its thread's position (record 0346).
    Reply(&'a ModernReply, u32, u32),
    Theme(&'a litchi_drawingml::theme::Theme),
    Column(&'a super::columns::ColumnItem),
    Hyperlink(&'a super::links::LinkItem),
    Media(Box<MediaProps<'a>>),
    Equation(&'a super::equations::EquationItem),
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
    /// The click action of a shape, picture or group, when the caller
    /// resolved it (record 0325).
    link: Option<Arc<litchi_pptx::hyperlinks::Link>>,
}

/// The inputs a [`TreeProps`] reads.
#[derive(Debug, Clone, Copy)]
struct Needs {
    transform: bool,
    description: bool,
    text: bool,
}

impl Needs {
    /// What `format` reads: everything (the text body for shapes only).
    const fn all(text: bool) -> Self {
        Self {
            transform: true,
            description: true,
            text,
        }
    }

    /// The inputs of a geometry key: the transform when the scene has no
    /// bounds for the element.
    fn geometry(tree: &TreeItem) -> Self {
        Self {
            transform: tree.handle.view().bounds().is_none(),
            description: false,
            text: false,
        }
    }

    const NONE: Self = Self {
        transform: false,
        description: false,
        text: false,
    };

    const TRANSFORM: Self = Self {
        transform: true,
        description: false,
        text: false,
    };

    const DESCRIPTION: Self = Self {
        transform: false,
        description: true,
        text: false,
    };

    fn shape(tree: &TreeItem, key: ShapeKey) -> Self {
        match key {
            ShapeKey::Name
            | ShapeKey::PhType
            | ShapeKey::PhIndex
            | ShapeKey::Id
            | ShapeKey::IsTitle
            | ShapeKey::Zorder
            | ShapeKey::Link
            | ShapeKey::Tooltip => Self::NONE,
            ShapeKey::Alt => Self::DESCRIPTION,
            ShapeKey::X | ShapeKey::Y | ShapeKey::Width | ShapeKey::Height => Self::geometry(tree),
            ShapeKey::FlipH | ShapeKey::FlipV | ShapeKey::Rotation => Self::TRANSFORM,
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
            | ShapeKey::Direction => Self {
                transform: false,
                description: false,
                text: true,
            },
        }
    }

    fn frame(tree: &TreeItem, key: FrameKey) -> Self {
        match key {
            FrameKey::Name | FrameKey::Id | FrameKey::Zorder => Self::NONE,
            FrameKey::Alt => Self::DESCRIPTION,
            FrameKey::X | FrameKey::Y | FrameKey::Width | FrameKey::Height => Self::geometry(tree),
            FrameKey::Rotation | FrameKey::FlipH | FrameKey::FlipV => Self::TRANSFORM,
        }
    }

    fn smartart(tree: &TreeItem, key: SmartArtKey) -> Self {
        match key {
            SmartArtKey::Name | SmartArtKey::Id | SmartArtKey::Zorder => Self::NONE,
            SmartArtKey::X | SmartArtKey::Y | SmartArtKey::Width | SmartArtKey::Height => {
                Self::geometry(tree)
            },
            SmartArtKey::Layout | SmartArtKey::LayoutId | SmartArtKey::NodeCount => Self::NONE,
        }
    }
}

impl<'a> TreeProps<'a> {
    fn read(tree: &'a TreeItem, include_text: bool) -> Result<Self> {
        Self::read_for(tree, Needs::all(include_text))
    }

    /// The values whose inputs `needs` names; the others read as absent.
    fn read_for(tree: &'a TreeItem, needs: Needs) -> Result<Self> {
        let shape = tree.handle.view();
        let transform = if needs.transform {
            shape.transform().map_err(Error::from)?
        } else {
            None
        };
        let description = if needs.description {
            shape.description().map_err(Error::from)?
        } else {
            None
        };
        let text = if needs.text {
            tree.body()?.map(|body| text::Facts::of(body.as_ref()))
        } else {
            None
        };
        Ok(Self {
            tree,
            transform,
            description,
            text,
            link: None,
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
            ShapeKey::Link => super::links::click_value(self.link.as_deref(), false),
            ShapeKey::Tooltip => super::links::click_value(self.link.as_deref(), true),
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

    /// The value at `index` of [`LINKED_FRAME_PROPS`]: a frame key, then the
    /// click action.
    fn linked_frame(&self, index: usize) -> Option<Value<'a>> {
        match LinkedFrameKey::at(index)? {
            LinkedFrameKey::Link => super::links::click_value(self.link.as_deref(), false),
            LinkedFrameKey::Tooltip => super::links::click_value(self.link.as_deref(), true),
            _ => self.frame(FrameKey::at(index)?),
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
            SmartArtKey::Layout | SmartArtKey::LayoutId | SmartArtKey::NodeCount => None,
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
                    RootKey::Keywords => metadata.and_then(|m| Value::owned_list(&m.keywords)),
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
            PropKind::LinkedFrame(props) => props.linked_frame(index),
            PropKind::SmartArt(props) => props.value(SmartArtKey::at(index)?),
            PropKind::DiagramNode(item) => match NodeKey::at(index)? {
                NodeKey::Level => item
                    .entry()
                    .map(|entry| Value::Int(i64::from(entry.depth) + 1)),
            },
            PropKind::Chart(props) => props.value_at(index),
            PropKind::Series(item) => series_value(item, SeriesKey::at(index)?),
            PropKind::Axis(item) => axis_value(item, AxisKey::at(index)?),
            PropKind::Table(props) => props.value(TableKey::at(index)?),
            PropKind::Paragraph(paragraph) => {
                text::paragraph_value(paragraph, ParagraphKey::at(index)?)
            },
            PropKind::Run(properties, link) => match RunKey::at(index)? {
                RunKey::Link => super::links::click_value(link.as_deref(), false),
                RunKey::Tooltip => super::links::click_value(link.as_deref(), true),
                key => text::run_value(properties, key),
            },
            PropKind::Cell(properties) => text::cell_value(properties, CellKey::at(index)?),
            PropKind::Row(height) => match RowKey::at(index)? {
                RowKey::Height => height.map(length),
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
            PropKind::Reply(reply, slide, comment) => reply_value(reply, *slide, *comment, index),
            PropKind::Theme(theme) => super::theme::value(theme, index),
            PropKind::Column(item) => super::columns::value(item, index),
            PropKind::Hyperlink(item) => super::links::value(item, index),
            PropKind::Equation(item) => super::equations::value(item, index),
            PropKind::Media(props) => match MediaKey::at(index)? {
                MediaKey::TrimStart => props.trim.0.clone().map(|ms| Value::Text(Cow::Owned(ms))),
                MediaKey::TrimEnd => props.trim.1.clone().map(|ms| Value::Text(Cow::Owned(ms))),
                _ => props.tree.frame(FrameKey::at(index)?),
            },
            PropKind::Empty => None,
        }
    }
}

/// A SmartArt frame's values: its frame properties, then its data model's.
struct SmartArtProps<'a> {
    tree: TreeProps<'a>,
    outline: Option<Arc<Outline>>,
}

impl<'a> SmartArtProps<'a> {
    fn value(&self, key: SmartArtKey) -> Option<Value<'a>> {
        match key {
            SmartArtKey::Layout => self
                .outline
                .as_ref()
                .and_then(|outline| super::smartart::layout_name(outline))
                .map(|name| Value::Text(Cow::Owned(name))),
            SmartArtKey::LayoutId => self
                .outline
                .as_ref()
                .and_then(|outline| outline.layout_type_id.clone())
                .map(|id| Value::Text(Cow::Owned(id))),
            SmartArtKey::NodeCount => self.outline.as_ref().map(|outline| {
                Value::Int(i64::try_from(outline.entries.len()).unwrap_or(i64::MAX))
            }),
            _ => self.tree.smartart(key),
        }
    }
}

/// A chart frame's values: its frame properties, then its chart's.
struct ChartProps<'a> {
    tree: TreeProps<'a>,
    chart: Option<Arc<Summary>>,
}

impl<'a> ChartProps<'a> {
    /// The value of the write table's entry at `index`: a chart key, or a
    /// formatting key read back from the chart's spec.
    fn value_at(&self, index: usize) -> Option<Value<'a>> {
        if let Some(key) = ChartKey::at(index) {
            return self.value(key);
        }
        let format = index.checked_sub(super::spec::CHART_FORMAT_INDEX)?;
        let (key, _) = crate::format::chart::FORMAT_PROPS.get(format)?;
        self.chart.as_ref()?.format(*key)
    }

    fn value(&self, key: ChartKey) -> Option<Value<'a>> {
        if let Some(frame) = FrameKey::at(key as usize) {
            return self.tree.frame(frame);
        }
        let chart = self.chart.as_ref()?;
        match key {
            ChartKey::ChartType => chart.chart_type().map(|name| Value::Text(Cow::Owned(name))),
            ChartKey::Title => chart
                .title
                .clone()
                .map(|title| Value::Text(Cow::Owned(title))),
            ChartKey::Legend => Some(Value::Text(Cow::Borrowed(
                crate::format::chart::legend_name(chart.legend),
            ))),
            ChartKey::Categories => chart
                .categories()
                .map(|categories| Value::Text(Cow::Owned(categories.join(",")))),
            ChartKey::CatTitle => chart
                .axis(crate::format::chart::Role::Category)
                .and_then(|axis| axis.title.clone())
                .map(|title| Value::Text(Cow::Owned(title))),
            _ => None,
        }
    }
}

fn series_value(item: &ChartItem, key: SeriesKey) -> Option<Value<'_>> {
    let series = item.summary.series.get(item.index as usize)?;
    match key {
        SeriesKey::Name => series
            .name
            .as_deref()
            .map(|name| Value::Text(Cow::Borrowed(name))),
        SeriesKey::Values => Some(Value::Text(Cow::Owned(crate::format::chart::join_values(
            &series.values,
        )))),
        SeriesKey::Categories => (!series.categories.is_empty())
            .then(|| Value::Text(Cow::Owned(series.categories.join(",")))),
        SeriesKey::Color => series
            .color
            .as_deref()
            .map(|color| Value::Text(Cow::Borrowed(color))),
    }
}

fn axis_value(item: &ChartItem, key: AxisKey) -> Option<Value<'_>> {
    let axis = item.summary.axes.get(item.index as usize)?;
    match key {
        AxisKey::Role => Some(Value::Text(Cow::Borrowed(axis.role.name()))),
        AxisKey::Title => axis
            .title
            .as_deref()
            .map(|title| Value::Text(Cow::Borrowed(title))),
        AxisKey::Min => axis.min.and_then(Value::number),
        AxisKey::Max => axis.max.and_then(Value::number),
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
    use super::spec::ModernCommentKey as Key;
    match Key::at(index)? {
        Key::Text => comment.text().map(|text| Value::Text(Cow::Borrowed(text))),
        Key::Author => comment
            .author_name()
            .map(|name| Value::Text(Cow::Borrowed(name))),
        Key::Initials => comment
            .author_initials()
            .map(|initials| Value::Text(Cow::Borrowed(initials))),
        Key::Created => Some(Value::Text(Cow::Borrowed(comment.created()))),
        Key::StartDate => comment
            .start_date()
            .map(|date| Value::Text(Cow::Borrowed(date))),
        Key::DueDate => comment
            .due_date()
            .map(|date| Value::Text(Cow::Borrowed(date))),
        Key::Title => comment
            .title()
            .map(|title| Value::Text(Cow::Borrowed(title))),
        Key::Status => Some(Value::Token(status_token(comment.status()))),
        Key::Resolved => Some(Value::Bool(comment.status() == ModernStatus::Resolved)),
        Key::Progress => Value::number(f64::from(comment.progress().thousandths()) / 1_000.0),
        Key::X => comment.location().map(|(x, _)| length(x)),
        Key::Y => comment.location().map(|(_, y)| length(y)),
        // A thread replies to nothing: OfficeCLI's `parent` is null.
        Key::Parent => None,
        // The native GUID, for raw-admitted queries only.
        Key::Id => Some(Value::Text(Cow::Borrowed(
            litchi_pptx::presentation::modern_comments::raw::comment_id(comment),
        ))),
    }
}

fn reply_value(reply: &ModernReply, slide: u32, comment: u32, index: usize) -> Option<Value<'_>> {
    use super::spec::ReplyKey as Key;
    match Key::at(index)? {
        Key::Text => reply.text().map(|text| Value::Text(Cow::Borrowed(text))),
        Key::Author => reply
            .author_name()
            .map(|name| Value::Text(Cow::Borrowed(name))),
        Key::Initials => reply
            .author_initials()
            .map(|initials| Value::Text(Cow::Borrowed(initials))),
        Key::Created => Some(Value::Text(Cow::Borrowed(reply.created()))),
        Key::Status => Some(Value::Token(status_token(reply.status()))),
        // The thread's path, as OfficeCLI's `parent` (record 0346).
        Key::Parent => Some(Value::Text(Cow::Owned(format!(
            "/slide[{}]/moderncomment[{}]",
            u64::from(slide) + 1,
            u64::from(comment) + 1
        )))),
        Key::Id => Some(Value::Text(Cow::Borrowed(
            litchi_pptx::presentation::modern_comments::raw::reply_id(reply),
        ))),
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
