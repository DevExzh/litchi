//! Paragraph and run `add` and `remove` in shape, table-cell and notes text
//! bodies (record 0333).
//!
//! `add(parent, paragraph, {text, align, bold, italic, size, color, font,
//! font.latin})` inserts a paragraph into a shape or placeholder, a table
//! cell (`/slide[N]/table[M]/tr[R]/tc[C]`) or the speaker notes
//! (`/slide[N]/notes`); `add(paragraph, run, {text, …})` inserts a text run.
//! Placement follows DOCX paragraph insertion: start, end (the default), a
//! zero-based position (past the end appends), or before or after a sibling
//! of the same body. `remove` deletes a paragraph (a body keeps one), a run
//! or a line break.
//!
//! Slide bodies stage through the native slide edit of either owner; notes
//! bodies live in the notes part, so a source owner is promoted to an eager
//! package for them, as notes text writes are (record 0252).

use std::sync::Arc;

use litchi_core::Position;
use litchi_core::selector::Path;
use litchi_pptx::opened::{TextBody, TextProperties};

use super::node::{CellRef, Item, TextItem, TreeContainer, TreeItem};
use super::spec::{Kind, PARAGRAPH_WRITE_TABLE, RUN_WRITE_TABLE};
use super::structure::{bound, error};
use super::write::{
    Commit, Proof, native_snapshot_budget, opened_edit_options, source_edit_options,
};
use super::{Editable, Nav, Pptx, text};
use crate::path::PathArg;
use crate::registry::Operation;
use crate::target::TargetArg;
use crate::transaction::{LimitKind, SemanticLimits};
use crate::{Code, NewNode, Placement, Result};

/// The notes page's body placeholder of a slide, as a tree item of the
/// notes container; `None` when the slide has no notes page or the page
/// has no body placeholder.
pub(super) fn notes_body(file: &Pptx, slide: u32) -> Result<Option<TreeItem>> {
    let Some(scene) = file.slide_notes_scene(slide as usize)? else {
        return Ok(None);
    };
    Ok(scene
        .roots()
        .find(|handle| {
            handle
                .view()
                .common()
                .placeholder()
                .is_some_and(|placeholder| placeholder.kind() == Some("body"))
        })
        .map(|handle| TreeItem::in_container(TreeContainer::Notes(slide), Kind::Shape, handle, 0)))
}

/// A text body located for a native edit: its slide and its address.
#[derive(Clone, Copy, PartialEq, Eq)]
struct Located {
    slide: usize,
    body: TextBody,
}

impl Located {
    const fn notes(self) -> bool {
        matches!(self.body, TextBody::Notes { .. })
    }
}

/// The body of a shape tree item: a slide shape or a notes body.
fn tree_body(tree: &TreeItem) -> Result<Located> {
    if !matches!(tree.kind, Kind::Shape | Kind::Placeholder) {
        return Err(error(
            Code::UnsupportedType,
            "paragraphs go into a shape, a placeholder, a table cell or the notes",
        ));
    }
    let shape = tree.handle.position();
    match tree.container {
        TreeContainer::Slide(slide) => Ok(Located {
            slide: slide as usize,
            body: TextBody::Shape { shape },
        }),
        TreeContainer::Notes(slide) => Ok(Located {
            slide: slide as usize,
            body: TextBody::Notes { shape },
        }),
        TreeContainer::Master(_) | TreeContainer::Layout { .. } => Err(error(
            Code::UnsupportedCapability,
            "master and layout text structure writes are not integrated",
        )),
    }
}

/// The body that holds a paragraph, run or line break.
fn text_body(text: &TextItem) -> Result<Located> {
    match text.cell.get() {
        Some((row, column)) => {
            let Some(slide) = text.shape.container.slide() else {
                return Err(error(
                    Code::UnsupportedCapability,
                    "master and layout tables are not editable",
                ));
            };
            Ok(Located {
                slide: slide as usize,
                body: TextBody::Cell {
                    shape: text.shape.handle.position(),
                    row: Position::new(row as usize),
                    column: Position::new(column as usize),
                },
            })
        },
        None => tree_body(&text.shape),
    }
}

/// The body and the paragraphs of an `add` parent.
fn parent_body(
    file: &Pptx,
    item: &Item,
) -> Result<(Located, Arc<litchi_drawingml::text::body::Body>)> {
    match item {
        Item::Tree(tree) => {
            let located = tree_body(tree)?;
            let body = tree
                .body()?
                .cloned()
                .ok_or_else(|| error(Code::UnsupportedCapability, "the shape has no text body"))?;
            Ok((located, body))
        },
        Item::Cell(cell) => {
            let Some(slide) = cell.table.container.slide() else {
                return Err(error(
                    Code::UnsupportedCapability,
                    "master and layout tables are not editable",
                ));
            };
            if CellRef::new(cell.row, cell.column).is_none() {
                return Err(error(
                    Code::OutOfRange,
                    "the cell is past the addressable rows and columns",
                ));
            }
            let body = text::cell_body(cell)?.ok_or_else(|| {
                error(
                    Code::UnsupportedCapability,
                    "the table cell has no text body",
                )
            })?;
            Ok((
                Located {
                    slide: slide as usize,
                    body: TextBody::Cell {
                        shape: cell.table.handle.position(),
                        row: Position::new(cell.row as usize),
                        column: Position::new(cell.column as usize),
                    },
                },
                Arc::new(body),
            ))
        },
        Item::Notes(notes) => {
            let shape = notes_body(file, notes.slide)?.ok_or_else(|| {
                error(
                    Code::UnsupportedCapability,
                    "the notes page has no body placeholder",
                )
            })?;
            let located = tree_body(&shape)?;
            let body = shape.body()?.cloned().ok_or_else(|| {
                error(
                    Code::UnsupportedCapability,
                    "the notes body has no text body",
                )
            })?;
            Ok((located, body))
        },
        _ => Err(error(
            Code::UnsupportedType,
            "paragraphs go into a shape, a placeholder, a table cell or the notes",
        )),
    }
}

/// One native text structure change.
#[derive(Clone, Copy)]
enum Change<'a> {
    InsertParagraph(usize, TextProperties<'a>),
    RemoveParagraph(usize),
    InsertRun(usize, usize, TextProperties<'a>),
    RemoveRun(usize, usize),
}

impl Pptx {
    /// Publish one text structure change through the native edit of the
    /// body's owner.
    fn commit_text_structure(
        &self,
        located: Located,
        change: Change<'_>,
        limits: &SemanticLimits,
    ) -> Result<Commit> {
        let source = self.editable_owner()?;
        let slide = Position::new(located.slide);
        let body = located.body;
        let owner = match (&source, located.notes()) {
            (Editable::Source(_), true) => {
                Editable::Eager(super::smartart::eager(&source, self.0.ooxml_limits)?)
            },
            _ => source.clone(),
        };
        let native_limit = native_snapshot_budget(limits);
        let (target, proof) = match &owner {
            Editable::Source(package) => {
                let mut edit = package
                    .edit_with_options(source_edit_options(native_limit))
                    .map_err(super::write::native_error)?;
                match change {
                    Change::InsertParagraph(position, properties) => {
                        edit.insert_paragraph(slide, body, Position::new(position), properties)?;
                    },
                    Change::RemoveParagraph(paragraph) => {
                        edit.remove_paragraph(slide, body, Position::new(paragraph))?;
                    },
                    Change::InsertRun(paragraph, position, properties) => edit.insert_run(
                        slide,
                        body,
                        Position::new(paragraph),
                        Position::new(position),
                        properties,
                    )?,
                    Change::RemoveRun(paragraph, run) => {
                        edit.remove_run(slide, body, Position::new(paragraph), Position::new(run))?
                    },
                }
                let (target, patch) = edit.commit()?.into_parts();
                (Editable::Source(target), Proof::Source(patch))
            },
            Editable::Eager(package) => {
                let mut edit = package
                    .edit_with_options(opened_edit_options(native_limit))
                    .map_err(super::write::native_error)?;
                match change {
                    Change::InsertParagraph(position, properties) => {
                        edit.insert_paragraph(slide, body, Position::new(position), properties)?;
                    },
                    Change::RemoveParagraph(paragraph) => {
                        edit.remove_paragraph(slide, body, Position::new(paragraph))?;
                    },
                    Change::InsertRun(paragraph, position, properties) => edit.insert_run(
                        slide,
                        body,
                        Position::new(paragraph),
                        Position::new(position),
                        properties,
                    )?,
                    Change::RemoveRun(paragraph, run) => {
                        edit.remove_run(slide, body, Position::new(paragraph), Position::new(run))?
                    },
                }
                let (target, patch) = edit.commit()?.into_parts();
                (Editable::Eager(Arc::new(target)), Proof::Eager(patch))
            },
        };
        let after = match &target {
            Editable::Source(package) => self.adopt_source(package.clone())?,
            Editable::Eager(package) => self.adopt_eager(Arc::clone(package)),
        };
        Ok(Commit::new(after, source, target, proof, 1))
    }

    /// `add(parent, paragraph, {…})`.
    pub(super) fn add_paragraph(
        &self,
        parent: PathArg<'_>,
        node: NewNode<'_>,
        limits: &SemanticLimits,
    ) -> Result<(Commit, Path<'static>)> {
        let nav = Nav::new(self);
        let resolved = crate::path::resolve::get(&nav, &parent)?;
        let (located, body) = parent_body(self, resolved.item())?;
        let parent_path = crate::node::Project::path(&resolved).to_string();
        drop(resolved);
        bound(LimitKind::ExpandedTargets, 1, limits.max_expanded_targets())?;
        let values = super::write::text::admit_for(
            &PARAGRAPH_WRITE_TABLE,
            node.props(),
            Operation::Add,
            limits,
        )?;
        let count = body.paragraphs.len();
        let position = match node.placement() {
            None | Some(Placement::End) => count,
            Some(Placement::Start) => 0,
            Some(Placement::Position(position)) => position.get().min(count),
            Some(Placement::Before(path) | Placement::After(path)) => {
                let sibling = crate::path::resolve::get(&nav, path)?;
                let Item::Paragraph(text) = sibling.item() else {
                    return Err(error(
                        Code::InvalidArgument,
                        "a paragraph goes beside a paragraph of the same body",
                    ));
                };
                if text_body(text)? != located {
                    return Err(error(
                        Code::InvalidArgument,
                        "a paragraph goes beside a paragraph of the same body",
                    ));
                }
                let index = text.paragraph as usize;
                if matches!(node.placement(), Some(Placement::After(_))) {
                    index + 1
                } else {
                    index
                }
            },
            Some(Placement::Into(..)) => {
                return Err(error(
                    Code::UnsupportedCapability,
                    "a paragraph is placed at the start, the end, a position or beside a sibling",
                ));
            },
        };
        let commit = self.commit_text_structure(
            located,
            Change::InsertParagraph(position, values.properties()),
            limits,
        )?;
        let path = child_path(&parent_path, "paragraph", position)?;
        Ok((commit, path))
    }

    /// `add(paragraph, run, {…})`.
    pub(super) fn add_run(
        &self,
        parent: PathArg<'_>,
        node: NewNode<'_>,
        limits: &SemanticLimits,
    ) -> Result<(Commit, Path<'static>)> {
        let nav = Nav::new(self);
        let resolved = crate::path::resolve::get(&nav, &parent)?;
        let Item::Paragraph(paragraph) = resolved.item() else {
            return Err(error(Code::UnsupportedType, "runs go into a paragraph"));
        };
        let paragraph = paragraph.clone();
        let parent_path = crate::node::Project::path(&resolved).to_string();
        drop(resolved);
        let located = text_body(&paragraph)?;
        bound(LimitKind::ExpandedTargets, 1, limits.max_expanded_targets())?;
        let values =
            super::write::text::admit_for(&RUN_WRITE_TABLE, node.props(), Operation::Add, limits)?;
        if values
            .text()
            .is_some_and(|text| text.contains(['\n', '\r']))
        {
            return Err(
                error(Code::InvalidValue, "run text cannot contain a line break")
                    .with_property("text"),
            );
        }
        let Some(runs) = paragraph.paragraph().map(|paragraph| &paragraph.runs) else {
            return Err(error(Code::PathNotFound, "the paragraph is gone"));
        };
        // Text runs in order: their index among the paragraph's run
        // children, as the native edit counts them.
        let listed: Vec<usize> = runs
            .iter()
            .enumerate()
            .filter(|(_, run)| text::listed_run(run) == Some(Kind::Run))
            .map(|(index, _)| index)
            .collect();
        let (at, ordinal) = match node.placement() {
            None | Some(Placement::End) => (runs.len(), listed.len()),
            Some(Placement::Start) => (0, 0),
            Some(Placement::Position(position)) => match listed.get(position.get()) {
                Some(index) => (*index, position.get()),
                None => (runs.len(), listed.len()),
            },
            Some(Placement::Before(path) | Placement::After(path)) => {
                let sibling = crate::path::resolve::get(&nav, path)?;
                let (Item::Run(text) | Item::LineBreak(text)) = sibling.item() else {
                    return Err(error(
                        Code::InvalidArgument,
                        "a run goes beside a run of the same paragraph",
                    ));
                };
                if text_body(text)? != located || text.paragraph != paragraph.paragraph {
                    return Err(error(
                        Code::InvalidArgument,
                        "a run goes beside a run of the same paragraph",
                    ));
                }
                let after = matches!(node.placement(), Some(Placement::After(_)));
                let at = text.run as usize + usize::from(after);
                (at, listed.iter().filter(|index| **index < at).count())
            },
            Some(Placement::Into(..)) => {
                return Err(error(
                    Code::UnsupportedCapability,
                    "a run is placed at the start, the end, a position or beside a sibling",
                ));
            },
        };
        let commit = self.commit_text_structure(
            located,
            Change::InsertRun(paragraph.paragraph as usize, at, values.properties()),
            limits,
        )?;
        let path = child_path(&parent_path, "run", ordinal)?;
        Ok((commit, path))
    }

    /// A `remove` of a paragraph, run or line break path; `None` for other
    /// targets.
    pub(super) fn try_remove_text(
        &self,
        target: &TargetArg<'_>,
        limits: &SemanticLimits,
    ) -> Result<Option<Commit>> {
        let Some(path) = target.path() else {
            return Ok(None);
        };
        if !path.path().steps().last().is_some_and(|step| {
            ["paragraph", "p", "para", "run", "r", "linebreak", "br"]
                .iter()
                .any(|name| step.name().eq_ignore_ascii_case(name))
        }) {
            return Ok(None);
        }
        let node = crate::path::resolve::get(&Nav::new(self), path)?;
        let (text, change) = match node.item() {
            Item::Paragraph(text) => (
                text.clone(),
                Change::RemoveParagraph(text.paragraph as usize),
            ),
            Item::Run(text) | Item::LineBreak(text) => (
                text.clone(),
                Change::RemoveRun(text.paragraph as usize, text.run as usize),
            ),
            _ => return Ok(None),
        };
        drop(node);
        bound(LimitKind::ExpandedTargets, 1, limits.max_expanded_targets())?;
        let located = text_body(&text)?;
        self.commit_text_structure(located, change, limits)
            .map(Some)
    }
}

/// `parent/name[position + 1]` as a path.
fn child_path(parent: &str, name: &str, position: usize) -> Result<Path<'static>> {
    let text = format!("{parent}/{name}[{}]", position + 1);
    Path::parse(&text)
        .map(Path::into_owned)
        .map_err(|_| error(Code::InternalError, "the new element's path does not parse"))
}
