//! The DOCX node: one leaf value and its compact address (record 0010 §3,
//! §2.6).
//!
//! A [`Node`] carries the leaf's own value for its element, moved in from
//! the leaf's indexes: a paragraph, a run, a table row or cell, a hyperlink,
//! or the entry of a block container's index (the body, a cell or a
//! content control). Those values share the visible main-document payload
//! and the block index through `Arc`s, so a node pins exactly what the
//! leaf's memos hold and nothing else (§3.3).
//!
//! Addresses retain kind-specific positions. Canonical paths are positional;
//! native identity selection is available only through explicit raw paths.

use std::fmt;
use std::sync::{Arc, OnceLock};

use litchi_core::Position;
use litchi_docx::{
    BlockContentControl, BlockKind, Blocks, Cell, IndexedBlock, Inline, InlineHyperlink, Paragraph,
    Row, Run, Table,
};

use super::Docx;
use super::spec::Kind;
use crate::error::{Code, Error, Result};
use crate::node::{NodePath, PathStep};

/// A node of a DOCX document: an element of the main story, with its leaf
/// value and its address.
///
/// Reach one through `File::get`, `File::query` or a parent's children
/// (record 0010 §8.1). The typed accessors return the `litchi_docx` values
/// the node carries, so callers that need more than the node view read the
/// leaf directly.
///
/// ```
/// use litchi_ops::node::Node;
///
/// // The text of a DOCX paragraph node, read from the leaf value.
/// fn paragraph_text(node: &Node<'_>) -> Option<String> {
///     let paragraph = node.as_docx()?.as_paragraph().ok()??;
///     paragraph.text().ok()
/// }
/// # let _ = paragraph_text;
/// ```
#[derive(Clone)]
pub struct Node<'f> {
    pub(super) docx: &'f Docx,
    pub(super) kind: Kind,
    pub(super) value: Value,
    pub(super) address: Address,
    /// The rendered positional path, built on first use; a style's step
    /// borrows its identifier from the handle's styles part.
    paths: OnceLock<Box<[PathStep<'f>]>>,
}

/// The leaf value of a node.
#[derive(Clone)]
pub(super) enum Value {
    /// The document root reads through the handle.
    Document,
    /// The body's block index, read on first use: a body listed under the
    /// root decodes nothing until a request needs its blocks.
    Body(OnceLock<Blocks>),
    /// A paragraph and the container entry it was read from.
    Paragraph { block: Block, paragraph: Paragraph },
    /// A table, as a container entry; its `litchi_docx::Table` is built on
    /// use, because the leaf value carries an inline cache of 17 KiB.
    Table(Block),
    /// A block content control, as a container entry.
    Sdt(Block),
    /// Display math, as a container entry.
    Equation(Block),
    /// A body-level section break, as a container entry.
    Section(Block),
    /// A table row, shared: the leaf value carries an inline cell cache.
    Row(Arc<Row>),
    /// A table cell.
    Cell(Cell),
    /// A run.
    Run(Run),
    /// A hyperlink with its relationship resolved.
    Hyperlink(Arc<InlineHyperlink>),
    /// The comments part.
    Comments,
    /// An inline picture of a paragraph.
    Picture(Arc<litchi_docx::image::InlineImage>),
    /// One comment of the comments part.
    Comment(Arc<litchi_docx::comment::Comment>),
    /// A chart of a paragraph, with its paragraph.
    Chart(Arc<super::chart::ChartValue>),
    /// A SmartArt graphic (record 0295).
    SmartArt(Arc<super::smartart::SmartArtValue>),
    /// A series or an axis of a chart, by its position among the chart's
    /// series or axes (record 0335); the node's kind says which.
    ChartChild(Arc<super::chart::ChartValue>, u32),
    /// A content node of a SmartArt graphic (record 0335).
    DiagramNode(Arc<super::smartart::DiagramNodeValue>),
    /// A styles, numbering, story or note element (record 0319).
    Part(super::parts::PartValue),
    /// A tab stop, break, bookmark, field, table of contents, watermark or
    /// text box (record 0319).
    Mark(Arc<super::marks::MarkValue>),
    /// A grid column: its table's container entry and its position in the
    /// grid (record 0345).
    Column(Block, u32),
}

/// One direct child of a block container: the container's index and the
/// child's position in it.
#[derive(Clone)]
pub(super) struct Block {
    pub(super) container: Blocks,
    pub(super) position: Position,
}

impl Block {
    /// The index entry of the block.
    pub(super) fn indexed(&self) -> Result<IndexedBlock> {
        self.container
            .get(self.position)
            .ok_or_else(|| internal("a DOCX block position is outside its container"))
    }
}

/// One step of an address: a kind and the zero-based ordinal among the
/// parent's children of that kind.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(super) struct Step {
    pub(super) kind: Kind,
    pub(super) ordinal: u32,
}

impl Step {
    pub(super) const fn new(kind: Kind, ordinal: u32) -> Self {
        Self { kind, ordinal }
    }

    /// The positional path step: `p[3]`, `oMathPara[1]`, or `body`.
    fn positional(self) -> PathStep<'static> {
        if self.kind.is_singleton() {
            PathStep::new(self.kind.step_name())
        } else {
            PathStep::at(self.kind.step_name(), Position::new(self.ordinal as usize))
        }
    }
}

/// Steps kept inline: `/body/tbl/tr/tc/p/r` fits with one to spare;
/// deeper addresses (nested tables) are shared on the heap.
const INLINE_STEPS: usize = 7;

/// Compact positional address with shared storage for deeply nested paths.
#[derive(Clone)]
pub(super) struct Address {
    steps: Steps,
}

#[derive(Clone)]
enum Steps {
    Inline {
        len: u8,
        steps: [Step; INLINE_STEPS],
    },
    Heap(Arc<[Step]>),
}

impl Address {
    /// The root's address, `/`.
    pub(super) fn root() -> Self {
        Self {
            steps: Steps::Inline {
                len: 0,
                steps: [Step::new(Kind::Document, 0); INLINE_STEPS],
            },
        }
    }

    /// The steps, from the root.
    pub(super) fn steps(&self) -> &[Step] {
        match &self.steps {
            Steps::Inline { len, steps } => &steps[..usize::from(*len)],
            Steps::Heap(steps) => steps,
        }
    }

    /// The same address with `step` appended.
    pub(super) fn with(&self, step: Step) -> Self {
        let steps = match &self.steps {
            Steps::Inline { len, steps } if usize::from(*len) < INLINE_STEPS => {
                let mut steps = *steps;
                steps[usize::from(*len)] = step;
                Steps::Inline {
                    len: *len + 1,
                    steps,
                }
            },
            _ => {
                let mut all = Vec::with_capacity(self.steps().len() + 1);
                all.extend_from_slice(self.steps());
                all.push(step);
                Steps::Heap(all.into())
            },
        };
        Self { steps }
    }
}

impl<'f> Node<'f> {
    pub(super) fn new(docx: &'f Docx, kind: Kind, value: Value, address: Address) -> Self {
        Self {
            docx,
            kind,
            value,
            address,
            paths: OnceLock::new(),
        }
    }

    /// The document root, `/`.
    pub(super) fn root(docx: &'f Docx) -> Self {
        Self::new(docx, Kind::Document, Value::Document, Address::root())
    }

    /// The element kind.
    ///
    /// ```
    /// use litchi_ops::format::docx::Kind;
    /// use litchi_ops::node::Node;
    ///
    /// fn is_paragraph(node: &Node<'_>) -> bool {
    ///     node.as_docx().is_some_and(|node| node.element_kind() == Kind::Paragraph)
    /// }
    /// # let _ = is_paragraph;
    /// ```
    #[must_use]
    pub fn element_kind(&self) -> Kind {
        self.kind
    }

    /// The paragraph, for a paragraph node.
    ///
    /// # Errors
    /// Returns an error if a retained source revision changed.
    pub fn as_paragraph(&self) -> Result<Option<&Paragraph>> {
        self.docx.owner().check_source()?;
        let value: Option<&Paragraph> = match &self.value {
            Value::Paragraph { paragraph, .. } => Some(paragraph),
            _ => None,
        };
        self.docx.owner().finish(Ok(value))
    }

    /// The run, for a run node.
    ///
    /// # Errors
    /// Returns an error if a retained source revision changed.
    pub fn as_run(&self) -> Result<Option<&Run>> {
        self.docx.owner().check_source()?;
        let value: Option<&Run> = match &self.value {
            Value::Run(run) => Some(run),
            _ => None,
        };
        self.docx.owner().finish(Ok(value))
    }

    /// The table row, for a row node.
    ///
    /// # Errors
    /// Returns an error if a retained source revision changed.
    pub fn as_row(&self) -> Result<Option<&Row>> {
        self.docx.owner().check_source()?;
        let value: Option<&Row> = match &self.value {
            Value::Row(row) => Some(row),
            _ => None,
        };
        self.docx.owner().finish(Ok(value))
    }

    /// The table cell, for a cell node.
    ///
    /// # Errors
    /// Returns an error if a retained source revision changed.
    pub fn as_cell(&self) -> Result<Option<&Cell>> {
        self.docx.owner().check_source()?;
        let value: Option<&Cell> = match &self.value {
            Value::Cell(cell) => Some(cell),
            _ => None,
        };
        self.docx.owner().finish(Ok(value))
    }

    /// The hyperlink, with its target resolved, for a hyperlink node.
    ///
    /// # Errors
    /// Returns an error if a retained source revision changed.
    pub fn as_hyperlink(&self) -> Result<Option<&InlineHyperlink>> {
        self.docx.owner().check_source()?;
        let value: Option<&InlineHyperlink> = match &self.value {
            Value::Hyperlink(link) => Some(link),
            _ => None,
        };
        self.docx.owner().finish(Ok(value))
    }

    /// The table, for a table node, built from the block index on each call.
    ///
    /// # Errors
    /// Returns source revision and leaf block-read errors.
    pub fn table(&self) -> Result<Option<Table>> {
        self.docx.owner().check_source()?;
        let result = match &self.value {
            Value::Table(block) => block.indexed().map(|entry| entry.table()),
            _ => Ok(None),
        };
        self.docx.owner().finish(result)
    }

    /// The content control, for a content-control node, built from the
    /// block index on each call.
    ///
    /// # Errors
    /// Returns source revision and leaf block-read errors.
    pub fn content_control(&self) -> Result<Option<BlockContentControl>> {
        self.docx.owner().check_source()?;
        let result = match &self.value {
            Value::Sdt(block) => block.indexed().map(|entry| entry.content_control()),
            _ => Ok(None),
        };
        self.docx.owner().finish(result)
    }

    /// The block index entry of a paragraph, table, content control,
    /// equation or section node: its kind, its position in its container,
    /// and the transaction address M2 edits it by.
    ///
    /// # Errors
    /// Returns source revision and leaf block-read errors.
    pub fn block(&self) -> Result<Option<IndexedBlock>> {
        self.docx.owner().check_source()?;
        self.docx
            .owner()
            .finish(self.own_block().map(Block::indexed).transpose())
    }

    /// The position the document transaction addresses this block by, the
    /// raw direct-child ordinal of record 0010 §4.5.1, for a paragraph,
    /// table or content control of the main story.
    ///
    /// # Errors
    ///
    /// `unsupported_capability` when the element is not a block, sits
    /// inside a `w:customXml` wrapper, or when the raw and visible views of
    /// its container are not proven to agree (record 0011): the position is
    /// refused rather than guessed. Blocks read from a cell or a content
    /// control carry no proof of their own; their positions are refused
    /// until M2 checks them against the body index.
    pub fn transaction_position(&self) -> Result<Position> {
        self.docx.owner().check_source()?;
        let block = self.own_block().ok_or_else(|| {
            Error::new(
                Code::UnsupportedCapability,
                format!(
                    "a DOCX {} has no document transaction position",
                    self.kind.type_name()
                ),
            )
            .with_format(litchi_core::FileFormat::Docx)
        })?;
        let result = block
            .indexed()
            .and_then(|entry| Ok(entry.transaction_position()?));
        self.docx.owner().finish(result)
    }

    /// The container entry of a block node.
    pub(super) fn own_block(&self) -> Option<&Block> {
        match &self.value {
            Value::Paragraph { block, .. }
            | Value::Table(block)
            | Value::Sdt(block)
            | Value::Equation(block)
            | Value::Section(block) => Some(block),
            _ => None,
        }
    }

    /// The positional address its children extend.
    pub(super) fn child_address(&self) -> Address {
        self.address.clone()
    }

    /// The body's block index, read on first use.
    pub(super) fn body_blocks(&self, lock: &OnceLock<Blocks>) -> Result<Blocks> {
        if let Some(blocks) = lock.get() {
            return Ok(blocks.clone());
        }
        let blocks = super::nav::load_body(self.docx)?;
        let _ = lock.set(blocks.clone());
        Ok(blocks)
    }

    /// The only DOCX selections for which the M2 editor has a proven raw
    /// transaction address.
    ///
    /// The visible navigator flattens custom XML and markup-compatibility
    /// wrappers.  A write cannot use those visible ordinals unless the body
    /// index proves that they are the same direct positions in the raw main
    /// document.  Runs need one more proof: the addressable run set may walk
    /// through hyperlinks and other wrappers, while the transaction edits
    /// only direct `w:r` children of a paragraph.  Such mixed paragraphs are
    /// refused here instead of silently applying a stale or shifted ordinal.
    pub(super) fn write_selection(&self) -> Result<WriteSelection> {
        self.write_selection_cached(&mut None)
    }

    /// Reuse the body index only across nodes resolved from the same snapshot.
    /// Keep each node's physical-source fence and ordinal proof intact.
    pub(super) fn write_selection_cached(
        &self,
        body: &mut Option<Blocks>,
    ) -> Result<WriteSelection> {
        self.docx.owner().check_source()?;
        let result = self.write_selection_inner(body);
        self.docx.owner().finish(result)
    }

    fn write_selection_inner(&self, cached: &mut Option<Blocks>) -> Result<WriteSelection> {
        let steps = self.address.steps();
        let Some(body_step) = steps.first() else {
            return Err(write_refusal("the DOCX root has no editable selection"));
        };
        if body_step.kind != Kind::Body || body_step.ordinal != 0 {
            return Err(write_refusal(
                "only a direct main-document body selection is editable",
            ));
        }
        if cached.is_none() {
            // A direct body paragraph was resolved from the body index of
            // this owner's generation; reuse it rather than reading it again
            // (record 0251f).
            *cached = Some(match (&self.value, steps) {
                (Value::Paragraph { block, .. }, [_, _]) => block.container.clone(),
                _ => super::nav::load_body(self.docx)?,
            });
        }
        let body = cached
            .as_ref()
            .ok_or_else(|| Error::new(Code::InternalError, "the write body index is missing"))?;

        match (&self.value, self.kind, steps) {
            (Value::Paragraph { block, .. }, Kind::Paragraph, [_, paragraph_step]) => {
                let indexed = body
                    .nth(
                        BlockKind::Paragraph,
                        Position::new(paragraph_step.ordinal as usize),
                    )
                    .ok_or_else(|| write_refusal("the paragraph is outside the current body"))?;
                if indexed.position() != block.position {
                    return Err(write_refusal(
                        "the paragraph position is not stable in the main document",
                    ));
                }
                let position = indexed.transaction_position().map_err(Error::from)?;
                Ok(WriteSelection::Paragraph(position))
            },
            (Value::Run(_), Kind::Run, [_, paragraph_step, run_step]) => {
                let indexed = body
                    .nth(
                        BlockKind::Paragraph,
                        Position::new(paragraph_step.ordinal as usize),
                    )
                    .ok_or_else(|| write_refusal("the run's paragraph is outside the body"))?;
                let position = indexed.transaction_position().map_err(Error::from)?;
                let paragraph = indexed
                    .paragraph()
                    .ok_or_else(|| write_refusal("the run's paragraph has no transaction value"))?;
                let inlines = paragraph.inlines().map_err(Error::from)?;
                let direct_runs = inlines
                    .iter()
                    .filter(|inline| matches!(inline, Inline::Run(_)))
                    .count();
                // Opaque direct children are safe to retain around a run
                // when the leaf proves that they contain no additional
                // addressable runs.  A hyperlink, field, content control or
                // foreign subtree carrying a Word run therefore still fails
                // the equality check below, while inert extension children
                // keep their physical slots for exact fragment swaps.
                let addressable = paragraph.addressable_runs().map_err(Error::from)?;
                if direct_runs != addressable.len() || run_step.ordinal as usize >= direct_runs {
                    return Err(write_refusal(
                        "the run is not proven to be a direct paragraph child",
                    ));
                }
                Ok(WriteSelection::Run {
                    paragraph: position,
                    run: Position::new(run_step.ordinal as usize),
                })
            },
            _ => Err(write_refusal(
                "only direct body paragraphs and their direct runs are editable",
            )),
        }
    }

    /// The canonical positional path.
    pub(super) fn canonical_path(&self) -> NodePath<'_> {
        let steps = self.address.steps().len();
        let paths = self.paths();
        NodePath::new(&paths[..steps.min(paths.len())])
    }

    /// The positional path.
    pub(super) fn positional(&self) -> NodePath<'_> {
        self.canonical_path()
    }

    /// The last positional step, without allocating a complete path.
    pub(super) fn last_step(&self) -> Option<PathStep<'f>> {
        self.address.steps().last().map(|step| self.step(*step))
    }

    fn paths(&self) -> &[PathStep<'f>] {
        self.paths.get_or_init(|| self.build_paths())
    }

    /// Canonical semantic paths use positions; native identities require
    /// raw input. A style is addressed by its identifier (`/styles/Normal`).
    fn build_paths(&self) -> Box<[PathStep<'f>]> {
        self.address
            .steps()
            .iter()
            .map(|step| self.step(*step))
            .collect()
    }

    fn step(&self, step: Step) -> PathStep<'f> {
        match step.kind {
            Kind::Style => super::parts::style_step(self.docx, step.ordinal)
                .unwrap_or_else(|| step.positional()),
            _ => step.positional(),
        }
    }
}

/// A checked address accepted by the DOCX transaction editor.
#[derive(Clone, Copy, PartialEq, Eq)]
pub(super) enum WriteSelection {
    Paragraph(Position),
    Run { paragraph: Position, run: Position },
}

/// Read every relevant identity strictly; malformed values do not become misses.
pub(super) fn first_control_with_id(container: &Blocks, id: i32) -> Result<Option<Position>> {
    for block in container
        .iter()
        .filter(|block| block.kind() == BlockKind::ContentControl)
    {
        if let Some(control) = block.content_control()
            && litchi_docx::document::raw::content_control_id(&control)? == Some(id)
        {
            return Ok(Some(block.position()));
        }
    }
    Ok(None)
}

fn internal(message: &'static str) -> Error {
    Error::new(Code::InternalError, message).with_format(litchi_core::FileFormat::Docx)
}

fn write_refusal(message: &'static str) -> Error {
    Error::new(Code::UnsupportedCapability, message).with_format(litchi_core::FileFormat::Docx)
}

impl fmt::Debug for Node<'_> {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("docx::Node")
            .field("kind", &self.kind)
            .finish_non_exhaustive()
    }
}

const _: () = {
    const fn assert_send_sync<T: Send + Sync>() {}
    assert_send_sync::<Node<'static>>();
    // A node moves through every walk and view: keep it small. The leaf
    // values dominate (a paragraph and its container entry).
    assert!(size_of::<Node<'static>>() <= 200);
};
