//! Physical element paths of DOCX nodes beyond the direct body (record 0251a).
//!
//! The read model addresses elements by their visible positions: blocks in a
//! container's visible index (`w:customXml` flattened, display equations
//! counted as `oMathPara`), rows and cells in the table's and row's visible
//! sets, runs in the paragraph's addressable set. The leaf's
//! [`ElementPath`] counts raw direct children. A node's path is converted
//! only when every step is proven to address the same element:
//!
//! - blocks use the leaf's [`IndexedBlock::transaction_position`], which
//!   refuses wrapped or unproven blocks;
//! - rows and cells use their visible ordinals, which the leaf refuses when
//!   a row- or cell-level wrapper would make them differ from raw positions;
//! - a run must be a direct run of its paragraph or a direct run of a direct
//!   hyperlink, with the addressable set equal to exactly those runs;
//! - a hyperlink must be resolved, with no unresolved direct hyperlink in its
//!   paragraph.
//!
//! Anything else is `unsupported_capability`, never a guess.

use litchi_core::selector::{Path, Predicate, Step as PathStep};
use litchi_core::{FileFormat, Position};
use litchi_docx::document::{ElementPath, ElementStep};
use litchi_docx::{Blocks, IndexedBlock, Inline, InlineKind, Paragraph};

use super::Docx;
use super::nav;
use super::node::Node;
use super::spec::Kind;
use crate::error::{Code, Error, Result};

/// Where a step of the walk stands. Table structure is materialized only
/// when a later step needs a cell's blocks.
enum State {
    Blocks(Blocks),
    Paragraph(Paragraph),
    Table(IndexedBlock),
    Row(IndexedBlock, usize),
    Cell(IndexedBlock, usize, usize),
    Done,
}

impl State {
    /// The blocks of a container state, materializing a cell's.
    fn into_blocks(self) -> Result<Option<Blocks>> {
        Ok(match self {
            Self::Blocks(blocks) => Some(blocks),
            Self::Cell(table, row, cell) => {
                let table = table
                    .table()
                    .ok_or_else(|| refusal("the DOCX table block holds no table"))?;
                let row = table
                    .rows()?
                    .into_iter()
                    .nth(row)
                    .ok_or_else(|| refusal("the DOCX row is outside its table"))?;
                let cell = row
                    .cells()?
                    .into_iter()
                    .nth(cell)
                    .ok_or_else(|| refusal("the DOCX cell is outside its row"))?;
                Some(cell.blocks()?)
            },
            _ => None,
        })
    }
}

/// The leaf element path of `node`, proven against the visible view.
///
/// Row and cell positions are the visible ordinals; the leaf refuses a row-
/// or cell-level wrapper that would make them differ from raw positions.
///
/// # Errors
///
/// Returns `unsupported_capability` for a node whose visible position is not
/// proven to be its raw position, or whose kind has no element path.
pub(super) fn element_path(node: &Node<'_>) -> Result<ElementPath> {
    let docx = node.docx;
    docx.owner().check_source()?;
    let result = element_path_inner(node);
    docx.owner().finish(result)
}

fn element_path_inner(node: &Node<'_>) -> Result<ElementPath> {
    let steps = node.address.steps();
    let Some((body, rest)) = steps.split_first() else {
        return Err(refusal("the DOCX root has no element path"));
    };
    if body.kind != Kind::Body || rest.is_empty() {
        return Err(refusal(
            "only elements below the main-document body have element paths",
        ));
    }
    walk(node, State::Blocks(nav::load_body(node.docx)?), rest, false)
}

/// Where a story element lives: its header or footer part, or its note.
#[derive(Debug, Clone, PartialEq, Eq)]
pub(super) enum StoryKey {
    Part(String),
    Note(litchi_docx::document::NoteKind, u32),
}

/// The story of a node under `/header[N]`, `/footer[N]`, `/footnote[N]`
/// or `/endnote[N]`, and the node's element path inside the story's
/// synthetic body (record 0325); `None` for any other node.
pub(super) fn story_path(node: &Node<'_>) -> Result<Option<(StoryKey, ElementPath)>> {
    let steps = node.address.steps();
    let skip = usize::from(
        steps
            .first()
            .is_some_and(|step| matches!(step.kind, Kind::Footnotes | Kind::Endnotes)),
    );
    let Some((story, rest)) = steps.get(skip..).and_then(<[_]>::split_first) else {
        return Ok(None);
    };
    if !matches!(
        story.kind,
        Kind::Header | Kind::Footer | Kind::Footnote | Kind::Endnote
    ) || rest.is_empty()
    {
        return Ok(None);
    }
    let docx = node.docx;
    let (key, value) = super::parts::story_of(docx, story.kind, story.ordinal as usize)?;
    let blocks = super::parts::story_blocks(docx, value)?
        .ok_or_else(|| refusal("the DOCX story holds no blocks"))?;
    let path = walk(node, State::Blocks(blocks), rest, true)?;
    docx.owner().check_source()?;
    Ok(Some((key, path)))
}

/// The element path of `rest`, the address steps below the container whose
/// blocks `state` holds. A story's block positions are its direct
/// positions.
fn walk(
    node: &Node<'_>,
    mut state: State,
    rest: &[super::node::Step],
    story: bool,
) -> Result<ElementPath> {
    let mut path = Vec::with_capacity(rest.len() + 1);
    let last = rest.len() - 1;
    for (index, step) in rest.iter().enumerate() {
        let ordinal = step.ordinal as usize;
        state = match (state, step.kind) {
            (state, Kind::Paragraph | Kind::Table | Kind::Sdt)
                if matches!(state, State::Blocks(_) | State::Cell(..)) =>
            {
                // The final block's own index entry is already on the node.
                let indexed = match node.own_block() {
                    Some(block) if index == last => block.indexed()?,
                    _ => {
                        let blocks = state
                            .into_blocks()?
                            .ok_or_else(|| refusal("the DOCX parent holds no blocks"))?;
                        let block_kind = step
                            .kind
                            .block_kind()
                            .ok_or_else(|| refusal("the DOCX block kind has no element path"))?;
                        blocks
                            .nth(block_kind, Position::new(ordinal))
                            .ok_or_else(|| refusal("the DOCX block is outside its container"))?
                    },
                };
                let position = if story {
                    indexed
                        .direct_position()
                        .ok_or_else(|| refusal("the DOCX story block has no direct position"))?
                } else {
                    raw_position(node.docx, &indexed)?
                };
                match step.kind {
                    Kind::Paragraph => {
                        path.push(ElementStep::Paragraph(position));
                        if index == last {
                            State::Done
                        } else {
                            State::Paragraph(indexed.paragraph().ok_or_else(|| {
                                refusal("the DOCX paragraph block holds no paragraph")
                            })?)
                        }
                    },
                    Kind::Table => {
                        path.push(ElementStep::Table(position));
                        State::Table(indexed)
                    },
                    _ => {
                        path.push(ElementStep::ContentControl(position));
                        if index == last {
                            State::Done
                        } else {
                            let control = indexed.content_control().ok_or_else(|| {
                                refusal("the DOCX content-control block holds no control")
                            })?;
                            State::Blocks(control.blocks()?)
                        }
                    },
                }
            },
            (State::Table(table), Kind::Row) => {
                path.push(ElementStep::Row(Position::new(ordinal)));
                State::Row(table, ordinal)
            },
            (State::Row(table, row), Kind::Cell) => {
                path.push(ElementStep::Cell(Position::new(ordinal)));
                State::Cell(table, row, ordinal)
            },
            (State::Paragraph(paragraph), Kind::Run) => {
                path.extend(run_location(node.docx, &paragraph, ordinal)?);
                State::Done
            },
            (State::Paragraph(paragraph), Kind::Hyperlink) => {
                hyperlinks_proven(node.docx, &paragraph)?;
                path.push(ElementStep::Hyperlink(Position::new(ordinal)));
                State::Done
            },
            _ => return Err(refusal("the DOCX element has no element path")),
        };
    }
    Ok(ElementPath::new(path))
}

/// The raw position of a visible block: the leaf's transaction position for
/// a body block, or its direct position in a nested container when the
/// body's index proves raw positions for the whole main document.
pub(super) fn raw_position(docx: &Docx, indexed: &IndexedBlock) -> Result<Position> {
    if let Ok(position) = indexed.transaction_position() {
        return Ok(position);
    }
    if nav::load_body(docx)?.raw_positions_proven()
        && let Some(position) = indexed.direct_position()
    {
        return Ok(position);
    }
    Err(refusal(
        "the DOCX block position is not proven in the raw part",
    ))
}

/// The raw location of addressable run `ordinal` of `paragraph`: a direct
/// run, or a direct run of a direct hyperlink.
pub(super) fn run_location(
    docx: &Docx,
    paragraph: &Paragraph,
    ordinal: usize,
) -> Result<Vec<ElementStep>> {
    let locations = run_locations(docx, paragraph)?;
    locations
        .into_iter()
        .nth(ordinal)
        .ok_or_else(|| refusal("the DOCX run is outside its paragraph"))
}

/// The raw locations of every addressable run of `paragraph`, in order.
pub(super) fn run_locations(docx: &Docx, paragraph: &Paragraph) -> Result<Vec<Vec<ElementStep>>> {
    let inlines = nav::inlines(docx, paragraph)?;
    let mut locations = Vec::new();
    let mut direct = 0usize;
    let mut links = 0usize;
    for inline in &inlines {
        match inline {
            Inline::Run(_) => {
                locations.push(vec![ElementStep::Run(Position::new(direct))]);
                direct += 1;
            },
            Inline::Hyperlink(link) => {
                for run in 0..link.runs().len() {
                    locations.push(vec![
                        ElementStep::Hyperlink(Position::new(links)),
                        ElementStep::Run(Position::new(run)),
                    ]);
                }
                links += 1;
            },
            Inline::Unknown(opaque) if opaque.kind() == InlineKind::Hyperlink => {
                return Err(refusal(
                    "the paragraph holds an unresolved hyperlink, so its run positions are not proven",
                ));
            },
            _ => {},
        }
    }
    if locations.len() != paragraph.addressable_runs()?.len() {
        return Err(refusal(
            "the paragraph's addressable runs are not all direct or hyperlink runs",
        ));
    }
    Ok(locations)
}

/// Prove that the paragraph's resolved hyperlinks are its direct
/// `w:hyperlink` children, in order.
fn hyperlinks_proven(docx: &Docx, paragraph: &Paragraph) -> Result<()> {
    let unresolved = nav::inlines(docx, paragraph)?.iter().any(|inline| {
        matches!(inline, Inline::Unknown(opaque) if opaque.kind() == InlineKind::Hyperlink)
    });
    if unresolved {
        return Err(refusal(
            "the paragraph holds an unresolved hyperlink, so hyperlink positions are not proven",
        ));
    }
    Ok(())
}

/// The canonical positional path of `node`, with `child` appended.
pub(super) fn path_with(node: &Node<'_>, child: Option<(Kind, Position)>) -> Path<'static> {
    let mut steps: Vec<PathStep<'static>> = node
        .address
        .steps()
        .iter()
        .map(|step| step_of(step.kind, Position::new(step.ordinal as usize)))
        .collect();
    if let Some((kind, position)) = child {
        steps.push(step_of(kind, position));
    }
    Path::from_steps(steps)
}

fn step_of(kind: Kind, position: Position) -> PathStep<'static> {
    if kind.is_singleton() {
        PathStep::new(kind.step_name())
    } else {
        PathStep::new(kind.step_name()).with(Predicate::Index(position))
    }
}

/// The visible ordinal, among the blocks of `kind` in `blocks`, of the block
/// whose raw position is `raw`.
pub(super) fn visible_ordinal(
    docx: &Docx,
    blocks: &Blocks,
    kind: Kind,
    raw: Position,
) -> Option<usize> {
    blocks
        .iter()
        .filter(|block| Kind::of_block(block.kind()) == Some(kind))
        .position(|block| raw_position(docx, &block).ok() == Some(raw))
}

pub(super) fn refusal(message: &'static str) -> Error {
    Error::new(Code::UnsupportedCapability, message).with_format(FileFormat::Docx)
}
