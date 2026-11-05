//! The DOCX adapter's read access for path resolution and queries (record
//! 0010 §4.1, §4.5, §5.3): [`Nav`] implements `Navigate` and `Inspect` over
//! a [`Docx`] handle.
//!
//! Each step reads only its own axis:
//!
//! | Step | Leaf read | Cost |
//! |---|---|---|
//! | `/body` | the visible main part and its body block index (L-DOCX-R1/R2) | once per part generation, then shared |
//! | `p[N]`, `tbl[N]`, `sdt[N]`, `oMathPara[N]`, `section[N]` | the container's block index by kind and ordinal | O(1) |
//! | `p[@paraId=X]` | the index's identifier table (L-DOCX-R7) | one start-tag pass per index, then O(log n) |
//! | `sdt[@sdtId=N]` | each content control's `w:sdtPr` | O(content controls) |
//! | `tr[N]`, `tc[N]` | the table's rows, the row's cells | the table's range |
//! | `tc[K]/p[J]`, `sdt[N]/p[M]` | the cell's or control's own block index | the cell or control |
//! | `r[N]` | the paragraph's addressable runs (L-DOCX-R8) | one pass over the paragraph |
//! | `hyperlink[N]` | the paragraph's direct hyperlinks and the main part's relationships | one pass over the paragraph |
//!
//! Nothing else is decoded: a `get` reads the main document part only.

use std::borrow::Cow;
use std::fmt;
use std::sync::{Arc, OnceLock};

use litchi_core::FileFormat;
use litchi_core::selector::Position;
use litchi_docx::{BlockKind, Blocks, IndexedBlock, Inline, InlineHyperlink, Paragraph};

use super::Docx;
use super::DocxOwner;
use super::children::Children;
use super::node::{Address, Block, Node, Step, Value, first_control_with_id};
use super::props;
use super::spec::{Kind, SPEC};
use crate::error::{Error, Result};
use crate::node::spec::{FormatSpec, GapSpec};
use crate::node::{Format, PathStep};
use crate::path::navigate::{Available, KindSet, Lookup, Navigate, Resolved, Select};
use crate::path::vocabulary::Identity;
use crate::query::inspect::Inspect;

/// The DOCX adapter over one handle.
#[derive(Debug, Clone, Copy)]
pub(crate) struct Nav<'f> {
    docx: &'f Docx,
}

impl<'f> Nav<'f> {
    pub(crate) const fn new(docx: &'f Docx) -> Self {
        Self { docx }
    }
}

#[cfg(test)]
thread_local! {
    pub(super) static BODY_LOADS: std::cell::Cell<usize> = const { std::cell::Cell::new(0) };
}

/// The body's block index of `docx`, shared with every other view of the
/// same main-part generation (L-DOCX-R1/R2).
pub(super) fn load_body(docx: &Docx) -> Result<Blocks> {
    #[cfg(test)]
    BODY_LOADS.with(|loads| loads.set(loads.get() + 1));
    if let Some(blocks) = docx.cached_body() {
        return Ok(blocks.clone());
    }
    // The handle's package is immutable, so its body index is too; a
    // source-backed owner's source fence is checked where the operation
    // finishes (record 0355a: one body index per batch base).
    let blocks = match docx.owner() {
        DocxOwner::Eager(package) => package.document()?.body_blocks()?,
        DocxOwner::Source(package) => package.document()?.body_blocks()?,
    };
    Ok(docx.cache_body(blocks).clone())
}

/// The body's node, with its block index read.
pub(super) fn body_node(docx: &Docx) -> Result<Node<'_>> {
    let blocks = load_body(docx)?;
    Ok(Node::new(
        docx,
        Kind::Body,
        Value::Body(OnceLock::from(blocks)),
        Address::root().with(Step::new(Kind::Body, 0)),
    ))
}

/// The direct hyperlinks of a main-story paragraph, with their targets
/// resolved through the main part's relationships.
pub(super) fn hyperlinks(docx: &Docx, paragraph: &Paragraph) -> Result<Vec<InlineHyperlink>> {
    let inlines = inlines(docx, paragraph)?;
    Ok(inlines
        .into_iter()
        .filter_map(|inline| match inline {
            Inline::Hyperlink(link) => Some(*link),
            _ => None,
        })
        .collect())
}

/// The inline children of a main-story paragraph, with hyperlinks resolved
/// through the main part's relationships.
pub(super) fn inlines(docx: &Docx, paragraph: &Paragraph) -> Result<Vec<Inline>> {
    Ok(match docx.owner() {
        DocxOwner::Eager(package) => {
            let document = package.document()?;
            paragraph.inlines_with_relationships(document.part().part().rels())?
        },
        DocxOwner::Source(package) => package.paragraph_inlines(paragraph)?,
    })
}

/// Whether the main document relates a comments part.
pub(super) fn has_comments(docx: &Docx) -> Result<bool> {
    let comments = |relationship: &litchi_opc::rel::Relationship| {
        !relationship.is_external()
            && matches!(
                relationship.reltype(),
                litchi_opc::constants::relationship_type::COMMENTS
                    | litchi_opc::constants::relationship_type::STRICT_COMMENTS
            )
    };
    let result = match docx.owner() {
        DocxOwner::Eager(package) => Ok(package
            .opc_package()
            .main_document_part()?
            .rels()
            .iter()
            .any(comments)),
        DocxOwner::Source(package) => package.has_comments().map_err(Error::from),
    };
    docx.owner().finish(result)
}

/// The comments of the main document, in part order.
pub(super) fn comments(docx: &Docx) -> Result<Vec<litchi_docx::comment::Comment>> {
    let result = match docx.owner() {
        DocxOwner::Eager(package) => package.document()?.comments().map_err(Error::from),
        DocxOwner::Source(package) => package.comments().map_err(Error::from),
    };
    docx.owner().finish(result)
}

/// The `commentsExtended` entries of the main document (record 0345);
/// empty when the part is absent. A part the leaf refuses as malformed
/// reads as empty, so a comment's other properties stay readable.
pub(super) fn comments_extended(
    docx: &Docx,
) -> Result<Arc<[litchi_docx::modern_comments::Extended]>> {
    if let Some(entries) = docx.parts().comment_threads.get() {
        return Ok(Arc::clone(entries));
    }
    let result = match docx.owner() {
        DocxOwner::Eager(package) => {
            Ok(package.document()?.comments_extended().unwrap_or_default())
        },
        DocxOwner::Source(package) => {
            package.check_source().map_err(Error::from)?;
            Ok(package.comments_extended().unwrap_or_default())
        },
    };
    let entries: Arc<[_]> = docx.owner().finish(result)?.into();
    Ok(Arc::clone(
        docx.parts().comment_threads.get_or_init(|| entries),
    ))
}

/// What a container selection picked.
enum Pick {
    /// The block, from the container's index.
    Found(IndexedBlock),
    Missing,
    NotApplicable,
}

/// The child of a block container (the body, a cell, a content control)
/// that `step` selects.
fn pick_block(container: &Blocks, kind: Kind, select: &Select<'_>) -> Result<Pick> {
    let Some(block_kind) = kind.block_kind() else {
        return Ok(Pick::NotApplicable);
    };
    let nth = |ordinal: usize| {
        container
            .nth(block_kind, Position::new(ordinal))
            .map_or(Pick::Missing, Pick::Found)
    };
    Ok(match *select {
        Select::None => nth(0),
        Select::Ordinal(position) => nth(position.get()),
        Select::Last => match container.count(block_kind).checked_sub(1) {
            Some(last) => nth(last),
            None => Pick::Missing,
        },
        Select::Identity(identities) => match kind {
            Kind::Paragraph => paragraph_by_identity(container, identities.iter())?,
            Kind::Sdt => sdt_by_identity(container, identities.iter())?,
            _ => Pick::Missing,
        },
        _ => Pick::NotApplicable,
    })
}

/// The first paragraph whose identities all match: `paraId` through the
/// index's identifier table, `textId` by reading start tags.
fn paragraph_by_identity<'s>(
    container: &Blocks,
    identities: impl Iterator<Item = (Identity, &'s str)>,
) -> Result<Pick> {
    let mut para_id = None;
    let mut text_id = None;
    for (identity, value) in identities {
        let id = litchi_docx::paragraph::Id::parse(value).map_err(|_| {
            Error::new(
                crate::Code::InvalidPath,
                "invalid native paragraph identifier",
            )
        })?;
        let slot = if identity.key.eq_ignore_ascii_case("paraId") {
            &mut para_id
        } else {
            &mut text_id
        };
        if slot.is_some_and(|held| held != id) {
            return Ok(Pick::Missing);
        }
        *slot = Some(id);
    }
    if let Some(id) = para_id {
        if let Some(block) = litchi_docx::document::raw::by_para_id(container, id)?
            && block.kind() == BlockKind::Paragraph
        {
            let paragraph = block.paragraph().ok_or_else(|| {
                Error::new(
                    crate::Code::InternalError,
                    "paragraph index has no paragraph",
                )
            })?;
            if text_id.is_none()
                || litchi_docx::document::raw::paragraph_ids(&paragraph)?.text_id() == text_id
            {
                return Ok(Pick::Found(block));
            }
        }
        return Ok(Pick::Missing);
    }
    for block in container
        .iter()
        .filter(|block| block.kind() == BlockKind::Paragraph)
    {
        if let Some(paragraph) = block.paragraph()
            && litchi_docx::document::raw::paragraph_ids(&paragraph)?.text_id() == text_id
        {
            return Ok(Pick::Found(block));
        }
    }
    Ok(Pick::Missing)
}

fn sdt_by_identity<'s>(
    container: &Blocks,
    identities: impl Iterator<Item = (Identity, &'s str)>,
) -> Result<Pick> {
    let mut wanted = None;
    for (_, value) in identities {
        let id = value.trim().parse::<i32>().map_err(|_| {
            Error::new(
                crate::Code::InvalidPath,
                "invalid native content-control identifier",
            )
        })?;
        if wanted.is_some_and(|held| held != id) {
            return Ok(Pick::Missing);
        }
        wanted = Some(id);
    }
    let Some(id) = wanted else {
        return Ok(Pick::Missing);
    };
    Ok(first_control_with_id(container, id)?
        .and_then(|position| container.get(position))
        .map_or(Pick::Missing, Pick::Found))
}

/// The block node of `kind` for `indexed`, a block of `container`.
pub(super) fn block_node<'f>(
    docx: &'f Docx,
    kind: Kind,
    container: &Blocks,
    indexed: &IndexedBlock,
    parent: &Address,
) -> Result<Node<'f>> {
    let ordinal = u32::try_from(indexed.ordinal().get()).unwrap_or(u32::MAX);
    let block = Block {
        container: container.clone(),
        position: indexed.position(),
    };
    let value = match kind {
        Kind::Paragraph => Value::Paragraph {
            paragraph: indexed.paragraph().ok_or_else(|| {
                Error::new(
                    crate::Code::InternalError,
                    "a DOCX paragraph block holds no paragraph",
                )
            })?,
            block,
        },
        Kind::Table => Value::Table(block),
        Kind::Sdt => Value::Sdt(block),
        Kind::Equation => Value::Equation(block),
        _ => Value::Section(block),
    };
    Ok(Node::new(
        docx,
        kind,
        value,
        parent.with(Step::new(kind, ordinal)),
    ))
}

/// The block container a node holds its block children in, if it is one.
pub(super) fn container_of(node: &Node<'_>) -> Result<Option<Blocks>> {
    Ok(match &node.value {
        Value::Body(lock) => Some(node.body_blocks(lock)?),
        Value::Cell(cell) => Some(cell.blocks()?),
        Value::Sdt(block) => match block.indexed()?.content_control() {
            Some(control) => Some(control.blocks()?),
            None => None,
        },
        // Header, footer and note stories (record 0325).
        Value::Part(part) => super::parts::story_blocks(node.docx, *part)?,
        _ => None,
    })
}

impl<'f> Nav<'f> {
    /// The rows of a table node, in order.
    pub(super) fn rows(node: &Node<'_>) -> Result<Vec<litchi_docx::Row>> {
        match &node.value {
            Value::Table(block) => match block.indexed()?.table() {
                Some(table) => Ok(table.rows()?.into_vec()),
                None => Ok(Vec::new()),
            },
            _ => Ok(Vec::new()),
        }
    }
}

impl<'f> Navigate<'f> for Nav<'f> {
    type Kind = Kind;
    type Node = Node<'f>;
    type Gap = GapSpec;
    type Children = Children<'f>;

    const FORMAT: FileFormat = FileFormat::Docx;
    /// `w:customXml` wrappers are flattened (record 0010 D4): their content
    /// is addressed as the container's own children, and the wrappers are
    /// not elements of the visible view.
    const ABSENT: &'static [&'static str] = &["customXml"];

    fn check_source(&self) -> Result<()> {
        self.docx.owner().ensure_current()
    }

    fn fence_key(&self) -> usize {
        self.docx.owner().fence_key()
    }

    fn root(&self) -> Result<Node<'f>> {
        Ok(Node::root(self.docx))
    }

    fn child(
        &self,
        parent: &Node<'f>,
        step: &Resolved<'_, Kind>,
    ) -> Result<Lookup<Node<'f>, GapSpec>> {
        let kind = step.kind;
        if let Some(gap) = kind.spec().gap {
            return Ok(Lookup::Gap(gap));
        }
        let address = parent.child_address();
        match (&parent.value, kind) {
            (Value::Document, Kind::Body) => {
                let found = match step.select {
                    Select::None | Select::Last => true,
                    Select::Ordinal(position) => position.get() == 0,
                    _ => return Ok(Lookup::NotApplicable),
                };
                if !found {
                    return Ok(Lookup::Missing);
                }
                Ok(Lookup::Found(body_node(self.docx)?))
            },
            (
                Value::Body(_)
                | Value::Cell(_)
                | Value::Sdt(_)
                | Value::Part(super::parts::PartValue::Story(_) | super::parts::PartValue::Note(..)),
                _,
            ) if kind.block_kind().is_some() => {
                let Some(container) = container_of(parent)? else {
                    return Ok(Lookup::Missing);
                };
                Ok(match pick_block(&container, kind, &step.select)? {
                    Pick::Found(indexed) => {
                        Lookup::Found(block_node(self.docx, kind, &container, &indexed, &address)?)
                    },
                    Pick::Missing => Lookup::Missing,
                    Pick::NotApplicable => Lookup::NotApplicable,
                })
            },
            (Value::Document, Kind::Comments) => {
                let found = match step.select {
                    Select::None | Select::Last => true,
                    Select::Ordinal(position) => position.get() == 0,
                    _ => return Ok(Lookup::NotApplicable),
                };
                if !found || !has_comments(self.docx)? {
                    return Ok(Lookup::Missing);
                }
                Ok(Lookup::Found(Node::new(
                    self.docx,
                    Kind::Comments,
                    Value::Comments,
                    address.with(Step::new(Kind::Comments, 0)),
                )))
            },
            (Value::Comments, Kind::Comment) => {
                let comments = comments(self.docx)?;
                let picked = if let Select::Identity(identities) = &step.select {
                    let mut wanted = None;
                    for (_, value) in identities.iter() {
                        let id = value.trim().parse::<u32>().map_err(|_| {
                            Error::new(crate::Code::InvalidPath, "invalid comment identifier")
                        })?;
                        if wanted.is_some_and(|held| held != id) {
                            wanted = None;
                            break;
                        }
                        wanted = Some(id);
                    }
                    Some(wanted.and_then(|id| comments.iter().position(|c| c.id() == id)))
                } else {
                    pick_index(comments.len(), &step.select)
                };
                Ok(match picked {
                    Some(Some(index)) => {
                        let comment = comments.into_iter().nth(index).ok_or_else(missing_child)?;
                        Lookup::Found(Node::new(
                            self.docx,
                            Kind::Comment,
                            Value::Comment(Arc::new(comment)),
                            address.with(Step::new(Kind::Comment, ordinal(index))),
                        ))
                    },
                    Some(None) => Lookup::Missing,
                    None => Lookup::NotApplicable,
                })
            },
            (Value::Table(block), Kind::Row) => {
                if let Select::Ordinal(position) = step.select {
                    // Read the table only up to the addressed row (record 0281).
                    let row = match block.indexed()?.table() {
                        Some(table) => table.row(position.get())?,
                        None => None,
                    };
                    return Ok(row.map_or(Lookup::Missing, |row| {
                        Lookup::Found(Node::new(
                            self.docx,
                            Kind::Row,
                            Value::Row(Arc::new(row)),
                            address.with(Step::new(Kind::Row, ordinal(position.get()))),
                        ))
                    }));
                }
                let rows = Self::rows(parent)?;
                Ok(match pick_index(rows.len(), &step.select) {
                    Some(Some(index)) => {
                        let row = rows.into_iter().nth(index).ok_or_else(missing_child)?;
                        Lookup::Found(Node::new(
                            self.docx,
                            Kind::Row,
                            Value::Row(Arc::new(row)),
                            address.with(Step::new(Kind::Row, ordinal(index))),
                        ))
                    },
                    Some(None) => Lookup::Missing,
                    None => Lookup::NotApplicable,
                })
            },
            // Grid columns (record 0345).
            (Value::Table(block), Kind::Column) => {
                super::column::pick_column(self.docx, parent, block, &step.select)
            },
            (Value::Row(row), Kind::Cell) => {
                if let Select::Ordinal(position) = step.select {
                    return Ok(row.cell(position.get())?.map_or(Lookup::Missing, |cell| {
                        Lookup::Found(Node::new(
                            self.docx,
                            Kind::Cell,
                            Value::Cell(cell),
                            address.with(Step::new(Kind::Cell, ordinal(position.get()))),
                        ))
                    }));
                }
                let cells = row.cells()?;
                Ok(match pick_index(cells.len(), &step.select) {
                    Some(Some(index)) => {
                        let cell = cells.into_iter().nth(index).ok_or_else(missing_child)?;
                        Lookup::Found(Node::new(
                            self.docx,
                            Kind::Cell,
                            Value::Cell(cell),
                            address.with(Step::new(Kind::Cell, ordinal(index))),
                        ))
                    },
                    Some(None) => Lookup::Missing,
                    None => Lookup::NotApplicable,
                })
            },
            (Value::Paragraph { paragraph, .. }, Kind::Run) => {
                let runs = paragraph.addressable_runs()?;
                Ok(match pick_index(runs.len(), &step.select) {
                    Some(Some(index)) => {
                        let run = runs.into_iter().nth(index).ok_or_else(missing_child)?;
                        Lookup::Found(Node::new(
                            self.docx,
                            Kind::Run,
                            Value::Run(run),
                            address.with(Step::new(Kind::Run, ordinal(index))),
                        ))
                    },
                    Some(None) => Lookup::Missing,
                    None => Lookup::NotApplicable,
                })
            },
            (Value::Paragraph { paragraph, .. }, Kind::Picture) => {
                let pictures = paragraph.images()?;
                Ok(match pick_index(pictures.len(), &step.select) {
                    Some(Some(index)) => {
                        let picture = pictures.into_iter().nth(index).ok_or_else(missing_child)?;
                        Lookup::Found(Node::new(
                            self.docx,
                            Kind::Picture,
                            Value::Picture(Arc::new(picture)),
                            address.with(Step::new(Kind::Picture, ordinal(index))),
                        ))
                    },
                    Some(None) => Lookup::Missing,
                    None => Lookup::NotApplicable,
                })
            },
            (Value::Paragraph { paragraph, .. }, Kind::Chart) => {
                let charts = paragraph.charts()?;
                Ok(match pick_index(charts.len(), &step.select) {
                    Some(Some(index)) => {
                        let chart = charts.into_iter().nth(index).ok_or_else(missing_child)?;
                        Lookup::Found(super::chart::chart_node(
                            self.docx,
                            &parent.value,
                            &address,
                            ordinal(index),
                            chart,
                        ))
                    },
                    Some(None) => Lookup::Missing,
                    None => Lookup::NotApplicable,
                })
            },
            // OfficeCLI addresses charts document-wide (`/chart[N]`).
            (Value::Document | Value::Body(_), Kind::Chart) => {
                super::chart::pick_document_chart(self.docx, &step.select)
            },
            // SmartArt is addressed document-wide too (record 0295).
            (Value::Document | Value::Body(_), Kind::SmartArt) => {
                super::smartart::pick_document_smartart(self.docx, &step.select)
            },
            // A chart's series and axes, a graphic's nodes (record 0335).
            (Value::Chart(chart), Kind::Series | Kind::Axis) => {
                super::chart::pick_chart_child(self.docx, parent, chart, kind, &step.select)
            },
            (Value::SmartArt(_) | Value::DiagramNode(_), Kind::Node) => {
                super::smartart::pick_node(self.docx, parent, &step.select)
            },
            (Value::Paragraph { paragraph, .. }, Kind::Hyperlink) => {
                let links = hyperlinks(self.docx, paragraph)?;
                Ok(match pick_index(links.len(), &step.select) {
                    Some(Some(index)) => {
                        let link = links.into_iter().nth(index).ok_or_else(missing_child)?;
                        Lookup::Found(Node::new(
                            self.docx,
                            Kind::Hyperlink,
                            Value::Hyperlink(Arc::new(link)),
                            address.with(Step::new(Kind::Hyperlink, ordinal(index))),
                        ))
                    },
                    Some(None) => Lookup::Missing,
                    None => Lookup::NotApplicable,
                })
            },
            // Styles, numbering, headers, footers and notes (record 0319).
            (Value::Document | Value::Part(_), kind) if super::parts::is_part_kind(kind) => {
                super::parts::child(self.docx, parent, step)
            },
            // Tab stops, breaks, bookmarks, fields, tables of contents, the
            // watermark and text boxes (record 0319).
            (_, kind) if super::marks::is_mark_kind(kind) => {
                super::marks::child(self.docx, parent, step)
            },
            _ => Ok(Lookup::NotApplicable),
        }
    }

    fn children(&self, parent: &Node<'f>, kinds: KindSet<Kind>) -> Result<Children<'f>> {
        Children::of(parent, kinds)
    }

    fn available(
        &self,
        parent: &Node<'f>,
        step: &Resolved<'_, Kind>,
        out: &mut Available,
    ) -> Result<()> {
        let _ = step;
        match &parent.value {
            Value::Document => out.name("body"),
            Value::Body(_) | Value::Cell(_) | Value::Sdt(_) => {
                if let Some(container) = container_of(parent)? {
                    for kind in [
                        Kind::Paragraph,
                        Kind::Table,
                        Kind::Sdt,
                        Kind::Equation,
                        Kind::Section,
                    ] {
                        if let Some(block_kind) = kind.block_kind() {
                            out.kind(kind.step_name(), container.count(block_kind));
                        }
                    }
                }
            },
            Value::Table(_) => out.kind("tr", Self::rows(parent)?.len()),
            Value::Row(row) => out.kind("tc", row.cell_count()?),
            Value::Paragraph { paragraph, .. } => {
                out.kind("r", paragraph.addressable_runs()?.len());
                out.kind("hyperlink", hyperlinks(self.docx, paragraph)?.len());
                out.kind("picture", paragraph.images()?.len());
            },
            _ => {},
        }
        Ok(())
    }
}

/// The index `select` picks among `len` children: `Some(None)` when there
/// is nothing at that index, `None` when the selection does not apply.
fn pick_index(len: usize, select: &Select<'_>) -> Option<Option<usize>> {
    Some(match *select {
        Select::None => (len > 0).then_some(0),
        Select::Ordinal(position) => (position.get() < len).then(|| position.get()),
        Select::Last => len.checked_sub(1),
        _ => return None,
    })
}

fn ordinal(index: usize) -> u32 {
    u32::try_from(index).unwrap_or(u32::MAX)
}

fn missing_child() -> Error {
    Error::new(
        crate::Code::InternalError,
        "a DOCX child vanished between counting and reading",
    )
}

impl<'f> Inspect<'f> for Nav<'f> {
    const SPEC: &'static FormatSpec<Kind> = &SPEC;
    const TOP: &'static [Kind] = &[Kind::Body];

    fn kind(&self, node: &Node<'f>) -> Kind {
        node.kind
    }

    fn type_name(&self, node: &Node<'f>) -> &'static str {
        node.kind.type_name()
    }

    fn text<'n>(&'n self, node: &'n Node<'f>) -> Result<Option<Cow<'n, str>>> {
        props::text(node)
    }

    fn style<'n>(&'n self, node: &'n Node<'f>) -> Result<Option<Cow<'n, str>>> {
        props::style(node)
    }

    fn format<'n>(&'n self, node: &'n Node<'f>) -> Result<Format<'n>> {
        props::format(node)
    }

    fn write_step(
        &self,
        node: &Node<'f>,
        kind: Kind,
        ordinal: Position,
        out: &mut fmt::Formatter<'_>,
    ) -> fmt::Result {
        let step = node.last_step().unwrap_or_else(|| {
            if kind.is_singleton() {
                PathStep::new(kind.step_name())
            } else {
                PathStep::at(kind.step_name(), ordinal)
            }
        });
        fmt::Display::fmt(&step, out)
    }
}
