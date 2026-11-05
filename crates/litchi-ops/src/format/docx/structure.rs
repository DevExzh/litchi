//! Structural DOCX verbs over the direct main story (M2 W5a).
//!
//! This module deliberately keeps the supported surface small and proved:
//! direct main-body paragraphs and their direct runs.  The leaf transaction
//! owns the XML splice, dependency checks and inverse patch, so these verbs
//! never rebuild a paragraph or run from the read model.  Visible nodes that
//! came through a wrapper, a table cell, a content control, or an unmodelled
//! inline context are refused before an edit is staged.

use litchi_core::selector::{Path, Predicate, Step};
use litchi_core::{FileFormat, Position};
use litchi_docx::document::{BodyAnchor, BodyBlock, NewParagraph, NewRun, RunAnchor};
use litchi_docx::{BlockKind, IndexedBlock};

use super::Docx;
use super::Nav;
use super::node::{Node, Value, WriteSelection};
use super::setters::Setter;
use super::spec::{Kind, PARAGRAPH_REGISTRY, RUN_REGISTRY};
use super::write::{Commit, WorkEdit};
use crate::error::{Code, Error, Result};
use crate::new_node::NewNode;
use crate::path::PathArg;
use crate::placement::Placement;
use crate::props::{Key, PropValue, Props};
use crate::registry::{Access, Entry, Operation};
use crate::target::TargetArg;
use crate::transaction::SemanticLimits;

/// A checked structural target.  The selection is computed from the source
/// snapshot before the editor is staged, so later removals cannot shift an
/// unproved ordinal.
#[derive(Clone, Copy)]
enum Selection {
    Paragraph(Position),
    Run { paragraph: Position, run: Position },
}

impl Selection {
    fn from_node(node: &Node<'_>) -> Result<Self> {
        match node.write_selection()? {
            WriteSelection::Paragraph(position) => Ok(Self::Paragraph(position)),
            WriteSelection::Run { paragraph, run } => Ok(Self::Run { paragraph, run }),
        }
    }

    fn key(self) -> (usize, usize, usize) {
        match self {
            // Paragraphs sort after runs at the same paragraph position.  A
            // run selected under a paragraph is pruned before sorting.
            Self::Paragraph(position) => (position.get(), 1, 0),
            Self::Run { paragraph, run } => (paragraph.get(), 0, run.get()),
        }
    }
}

impl Docx {
    /// Add a direct main-body paragraph or run and return its canonical path.
    ///
    /// The returned path is owned because it outlives the input node and the
    /// borrowed path arguments.  Every failure, including an edit admission
    /// failure, is passed through the final source fence.
    pub(crate) fn add<'p, 'n>(
        &self,
        parent: PathArg<'p>,
        node: NewNode<'n>,
    ) -> Result<(Commit, Path<'static>)> {
        self.add_with_limits(parent, node, &SemanticLimits::default())
    }

    pub(crate) fn add_with_limits<'p, 'n>(
        &self,
        parent: PathArg<'p>,
        node: NewNode<'n>,
        limits: &SemanticLimits,
    ) -> Result<(Commit, Path<'static>)> {
        let result = self.add_inner(parent, node, false, limits.max_snapshot_bytes());
        self.owner().finish(result)
    }

    pub(crate) fn add_command_with_limits<'p, 'n>(
        &self,
        parent: PathArg<'p>,
        node: NewNode<'n>,
        limits: &SemanticLimits,
    ) -> Result<(Commit, Path<'static>)> {
        self.owner()
            .finish(self.add_inner(parent, node, true, limits.max_snapshot_bytes()))
    }

    fn add_inner<'p, 'n>(
        &self,
        parent: PathArg<'p>,
        node: NewNode<'n>,
        command_text: bool,
        snapshot_budget: usize,
    ) -> Result<(Commit, Path<'static>)> {
        // Security and retention policy are checked before any property
        // conversion or source read.  Dropping this editor rolls back every
        // admission failure below.
        let editor = self.edit_with_snapshot_budget(snapshot_budget)?;
        let parent = resolve_path(self, &parent)?;
        // Header, footer and note paragraphs and runs (record 0325).
        let editor = match self.try_add_story(editor, &parent, &node)? {
            Ok(added) => return Ok(added),
            Err(editor) => editor,
        };
        let placement = node.placement().cloned();
        let kind = new_kind(node.kind())?;
        let (planned, anchor) = match (kind, parent.element_kind()) {
            (NewKind::Paragraph, Kind::Body) => {
                let (paragraph, list) = new_paragraph_with_list(node.props())?;
                let anchor = body_anchor(self, placement.as_ref(), false)?;
                (
                    PlannedNode::Paragraph(Box::new(paragraph), list),
                    PlannedAnchor::Body(anchor),
                )
            },
            (NewKind::Run, Kind::Paragraph) => {
                let selection = match parent.write_selection() {
                    Ok(selection) => selection,
                    // A nested paragraph, or a body paragraph whose
                    // addressable runs include hyperlink runs, takes the
                    // element-path route.
                    Err(error) if error.code() == Code::UnsupportedCapability => {
                        return self.add_nested(editor, &parent, kind, &node);
                    },
                    Err(error) => return Err(error),
                };
                let paragraph = match selection {
                    WriteSelection::Paragraph(position) => position,
                    WriteSelection::Run { .. } => {
                        return Err(refusal(
                            "a run can only be added under a direct body paragraph",
                        ));
                    },
                };
                // Hyperlink runs make the addressable run axis differ from
                // the direct runs the body route counts.
                if let Value::Paragraph {
                    paragraph: value, ..
                } = &parent.value
                    && super::element::run_locations(self, value)
                        .is_ok_and(|locations| locations.iter().any(|location| location.len() != 1))
                {
                    return self.add_nested(editor, &parent, kind, &node);
                }
                let run = new_run(node.props())?;
                let at = run_position(self, &parent, placement.as_ref())?;
                (
                    PlannedNode::Run(Box::new(run)),
                    PlannedAnchor::Run { paragraph, at },
                )
            },
            (NewKind::Run, _) => {
                return Err(refusal("a DOCX run can only be added under a paragraph"));
            },
            (
                NewKind::Paragraph
                | NewKind::Table
                | NewKind::Row
                | NewKind::Cell
                | NewKind::Column
                | NewKind::Hyperlink
                | NewKind::Comment
                | NewKind::Picture
                | NewKind::Chart
                | NewKind::SmartArt
                | NewKind::Node,
                _,
            ) => {
                return self.add_nested(editor, &parent, kind, &node);
            },
            // Styles, numbering, headers, footers and notes (record 0319).
            (NewKind::Part(part), _) => return self.add_part(editor, &parent, part, &node),
            // Tab stops and breaks (record 0319).
            (NewKind::Mark(mark), _) => return self.add_mark(editor, &parent, mark, &node),
        };
        self.owner().check_source()?;
        // A new list definition is a new part (record 0342), which only the
        // owning package publishes.
        let editor = match &planned {
            PlannedNode::Paragraph(_, Some(_)) => self.eager_editor(editor)?,
            _ => editor,
        };
        let mut edit = WorkEdit::from(editor)?;
        let path = match (planned, anchor) {
            (PlannedNode::Paragraph(paragraph, list), PlannedAnchor::Body(anchor)) => {
                let position = if command_text {
                    insert_command_paragraphs(&mut edit, anchor, &paragraph)?
                } else {
                    edit.insert_paragraph(anchor, &paragraph)?
                };
                if let Some(list) = list {
                    let WorkEdit::Eager(eager) = &mut edit else {
                        return Err(internal("a DOCX list edit needs the owning package editor"));
                    };
                    super::lists::number_added(
                        eager,
                        &litchi_docx::document::ElementPath::body(),
                        position,
                        list,
                    )?;
                }
                paragraph_path(position)
            },
            (PlannedNode::Run(run), PlannedAnchor::Run { paragraph, at }) => {
                let position = edit.insert_run(paragraph, at, &run)?;
                run_path(paragraph, position)
            },
            _ => return Err(internal("DOCX add plan does not match its parent")),
        };
        Ok((edit.commit(self)?, path))
    }

    /// Remove direct body paragraphs and direct runs selected by a path or a
    /// scoped query.  Query matches are deduplicated, paragraph descendants
    /// are subsumed by their selected paragraph, and removals are staged in
    /// reverse document order so no earlier operation shifts a later proof.
    pub(crate) fn remove<'p>(&self, target: TargetArg<'p>) -> Result<Commit> {
        self.remove_with_limits(target, &SemanticLimits::default())
    }

    pub(crate) fn remove_with_limits<'p>(
        &self,
        target: TargetArg<'p>,
        limits: &SemanticLimits,
    ) -> Result<Commit> {
        let result = self.remove_inner(target, limits.max_snapshot_bytes());
        self.owner().finish(result)
    }

    fn remove_inner<'p>(&self, target: TargetArg<'p>, snapshot_budget: usize) -> Result<Commit> {
        let editor = self.edit_with_snapshot_budget(snapshot_budget)?;
        let nodes = self.resolve_targets(&target)?;
        // Header, footer and note paragraphs and runs (record 0325).
        let editor = match self.try_remove_story(editor, &nodes)? {
            Ok(commit) => return Ok(commit),
            Err(editor) => editor,
        };
        let selections = match admit_remove(&nodes) {
            Ok(selections) => selections,
            // Anything beyond direct-body paragraphs and runs takes the
            // element-path route.
            Err(error)
                if matches!(
                    error.code(),
                    Code::UnsupportedCapability | Code::UnsupportedElement
                ) =>
            {
                return self.remove_nested(editor, &nodes);
            },
            Err(error) => return Err(error),
        };
        self.owner().check_source()?;
        if selections.is_empty() {
            return WorkEdit::from(editor)?.commit(self);
        }
        let mut edit = WorkEdit::from(editor)?;
        for selection in selections {
            match selection {
                Selection::Paragraph(position) => edit.remove_paragraph(position)?,
                Selection::Run { paragraph, run } => edit.remove_run(paragraph, run)?,
            }
        }
        edit.commit(self)
    }

    /// Move a direct body paragraph or direct run while retaining its exact
    /// source fragment.  The leaf transaction performs the namespace,
    /// dependency and field-boundary checks before splicing the fragment.
    pub(crate) fn move_to<'p, 'q>(
        &self,
        path: PathArg<'p>,
        placement: Placement<'q>,
    ) -> Result<Commit> {
        self.move_to_with_limits(path, placement, &SemanticLimits::default())
    }

    pub(crate) fn move_to_with_limits<'p, 'q>(
        &self,
        path: PathArg<'p>,
        placement: Placement<'q>,
        limits: &SemanticLimits,
    ) -> Result<Commit> {
        let result = self.move_inner(path, placement, limits.max_snapshot_bytes());
        self.owner().finish(result)
    }

    fn move_inner<'p, 'q>(
        &self,
        path: PathArg<'p>,
        placement: Placement<'q>,
        snapshot_budget: usize,
    ) -> Result<Commit> {
        let editor = self.edit_with_snapshot_budget(snapshot_budget)?;
        let node = resolve_path(self, &path)?;
        let selection = match Selection::from_node(&node) {
            Ok(selection) => selection,
            Err(error) if error.code() == Code::UnsupportedCapability => {
                return self.move_nested(editor, &node, &placement);
            },
            Err(error) => return Err(error),
        };
        let destination = match selection {
            Selection::Paragraph(position) => {
                MoveDestination::Paragraph(move_anchor(self, &node, position, &placement)?)
            },
            Selection::Run { paragraph, run } => {
                let (target_paragraph, anchor) = run_destination(self, paragraph, run, &placement)?;
                MoveDestination::Run {
                    source: (paragraph, run),
                    target: (target_paragraph, anchor),
                }
            },
        };
        self.owner().check_source()?;
        let mut edit = WorkEdit::from(editor)?;
        match destination {
            MoveDestination::Paragraph(anchor) => {
                let position = match selection {
                    Selection::Paragraph(position) => position,
                    Selection::Run { .. } => {
                        return Err(internal("DOCX paragraph move selection changed"));
                    },
                };
                edit.move_paragraph(position, anchor)?;
            },
            MoveDestination::Run {
                source: (paragraph, run),
                target: (target_paragraph, anchor),
            } => {
                if paragraph == target_paragraph {
                    edit.move_run(paragraph, run, anchor)?;
                } else {
                    edit.move_run_to(paragraph, run, target_paragraph, anchor)?;
                }
            },
        }
        edit.commit(self)
    }

    /// Swap two direct body paragraphs or direct runs while retaining each
    /// selected fragment's exact source bytes.
    pub(crate) fn swap<'p, 'q>(&self, a: PathArg<'p>, b: PathArg<'q>) -> Result<Commit> {
        self.swap_with_limits(a, b, &SemanticLimits::default())
    }

    pub(crate) fn swap_with_limits<'p, 'q>(
        &self,
        a: PathArg<'p>,
        b: PathArg<'q>,
        limits: &SemanticLimits,
    ) -> Result<Commit> {
        let result = self.swap_inner(a, b, limits.max_snapshot_bytes());
        self.owner().finish(result)
    }

    fn swap_inner<'p, 'q>(
        &self,
        a: PathArg<'p>,
        b: PathArg<'q>,
        snapshot_budget: usize,
    ) -> Result<Commit> {
        let editor = self.edit_with_snapshot_budget(snapshot_budget)?;
        let left = resolve_path(self, &a)?;
        let right = resolve_path(self, &b)?;
        let left = Selection::from_node(&left)?;
        let right = Selection::from_node(&right)?;
        if left.key() == right.key() {
            return WorkEdit::from(editor)?.commit(self);
        }
        self.owner().check_source()?;
        let mut edit = WorkEdit::from(editor)?;
        match (left, right) {
            (Selection::Paragraph(left), Selection::Paragraph(right)) => {
                edit.swap_paragraphs(left, right)?;
            },
            (
                Selection::Run {
                    paragraph: left_paragraph,
                    run: left_run,
                },
                Selection::Run {
                    paragraph: right_paragraph,
                    run: right_run,
                },
            ) if left_paragraph == right_paragraph => {
                edit.swap_runs(left_paragraph, left_run, right_run)?;
            },
            (
                Selection::Run {
                    paragraph: left_paragraph,
                    run: left_run,
                },
                Selection::Run {
                    paragraph: right_paragraph,
                    run: right_run,
                },
            ) => {
                edit.swap_runs_to(left_paragraph, left_run, right_paragraph, right_run)?;
            },
            _ => {
                return Err(refusal(
                    "DOCX swap requires two paragraphs or two direct runs",
                ));
            },
        }
        edit.commit(self)
    }
}

#[derive(Clone, Copy)]
pub(super) enum NewKind {
    Paragraph,
    Run,
    Table,
    Row,
    /// A table cell (record 0345).
    Cell,
    /// A table grid column (record 0345).
    Column,
    Hyperlink,
    Comment,
    Picture,
    Chart,
    SmartArt,
    /// A SmartArt content node (record 0335).
    Node,
    Part(super::parts::write::NewPart),
    Mark(super::marks::write::NewMark),
}

enum PlannedNode {
    Paragraph(Box<NewParagraph>, Option<super::lists::Request>),
    Run(Box<NewRun>),
}

enum PlannedAnchor {
    Body(BodyAnchor),
    Run {
        paragraph: Position,
        at: Option<Position>,
    },
}

enum MoveDestination {
    Paragraph(BodyAnchor),
    Run {
        source: (Position, Position),
        target: (Position, RunAnchor),
    },
}

fn new_kind(name: &str) -> Result<NewKind> {
    if name.eq_ignore_ascii_case("p") || name.eq_ignore_ascii_case("paragraph") {
        Ok(NewKind::Paragraph)
    } else if name.eq_ignore_ascii_case("r") || name.eq_ignore_ascii_case("run") {
        Ok(NewKind::Run)
    } else if name.eq_ignore_ascii_case("tbl") || name.eq_ignore_ascii_case("table") {
        Ok(NewKind::Table)
    } else if name.eq_ignore_ascii_case("tr") || name.eq_ignore_ascii_case("row") {
        Ok(NewKind::Row)
    } else if name.eq_ignore_ascii_case("tc") || name.eq_ignore_ascii_case("cell") {
        Ok(NewKind::Cell)
    } else if name.eq_ignore_ascii_case("col") || name.eq_ignore_ascii_case("column") {
        Ok(NewKind::Column)
    } else if name.eq_ignore_ascii_case("hyperlink") || name.eq_ignore_ascii_case("link") {
        Ok(NewKind::Hyperlink)
    } else if name.eq_ignore_ascii_case("comment") {
        Ok(NewKind::Comment)
    } else if ["picture", "image", "img"]
        .iter()
        .any(|kind| name.eq_ignore_ascii_case(kind))
    {
        Ok(NewKind::Picture)
    } else if name.eq_ignore_ascii_case("chart") {
        Ok(NewKind::Chart)
    } else if name.eq_ignore_ascii_case("smartart") {
        Ok(NewKind::SmartArt)
    } else if name.eq_ignore_ascii_case("node") {
        Ok(NewKind::Node)
    } else if let Some(part) = super::parts::write::NewPart::of(name) {
        Ok(NewKind::Part(part))
    } else if let Some(mark) = super::marks::write::NewMark::of(name) {
        Ok(NewKind::Mark(mark))
    } else {
        Err(Error::new(
            Code::UnsupportedElement,
            "the DOCX structural verb does not support this element type",
        )
        .with_format(FileFormat::Docx))
    }
}

pub(super) fn resolve_path<'f>(docx: &'f Docx, path: &PathArg<'_>) -> Result<Node<'f>> {
    crate::path::resolve::get(&Nav::new(docx), path)
}

/// A command is one transaction even when its authored text creates several
/// paragraphs. Native text semantics are retained by the ordinary add path.
fn insert_command_paragraphs(
    edit: &mut WorkEdit,
    mut anchor: BodyAnchor,
    template: &NewParagraph,
) -> Result<Position> {
    let [run] = template.runs() else {
        return edit.insert_paragraph(anchor, template);
    };
    let text = run.text();
    if !text.contains(['\r', '\n', '\u{b}']) {
        return edit.insert_paragraph(anchor, template);
    }
    let normalized = if text.contains('\r') {
        std::borrow::Cow::Owned(text.replace("\r\n", "\n").replace('\r', "\n"))
    } else {
        std::borrow::Cow::Borrowed(text)
    };
    let count = normalized.split('\n').count();
    let maximum = match edit {
        WorkEdit::Eager(edit) => edit.remaining_operations(),
        WorkEdit::Source(edit) => edit.remaining_operations(),
    };
    if count > maximum {
        return Err(litchi_docx::document::Error::Limit {
            kind: litchi_docx::document::LimitKind::Operations,
            observed: count,
            maximum,
        }
        .into());
    }
    let mut first = None;
    for line in normalized.split('\n') {
        let paragraph = NewParagraph::new()
            .with_properties(template.properties().clone())
            .with_run(
                NewRun::new(line.replace('\u{b}', "\n")).with_properties(run.properties().clone()),
            );
        let position = edit.insert_paragraph(anchor, &paragraph)?;
        first.get_or_insert(position);
        anchor = BodyAnchor::After(BodyBlock::Paragraph(position));
    }
    first.ok_or_else(|| internal("paragraph splitting produced no paragraph"))
}

/// A new paragraph that takes no list: `listStyle` and `start` are refused
/// (headers, footers and notes, as on `set`).
pub(super) fn new_paragraph(props: &Props<'_>) -> Result<NewParagraph> {
    let (paragraph, list) = new_paragraph_with_list(props)?;
    if list.is_some() {
        return Err(Error::new(
            Code::UnsupportedCapability,
            "list styles are not written in headers, footers and notes",
        )
        .with_property("listStyle")
        .with_format(FileFormat::Docx));
    }
    Ok(paragraph)
}

/// A new paragraph and the list its `listStyle` and `start` ask for (record
/// 0342), which [`super::lists::number_added`] applies once it is inserted.
pub(super) fn new_paragraph_with_list(
    props: &Props<'_>,
) -> Result<(NewParagraph, Option<super::lists::Request>)> {
    let mut paragraph = NewParagraph::new();
    let mut run = NewRun::new("");
    let mut has_run = false;
    let mut list = None::<super::lists::Request>;
    for (key, value) in props.iter() {
        if let Some((_, entry)) = PARAGRAPH_REGISTRY.lookup(key.as_str()) {
            let result = match (entry.access, super::lists::list_setter(entry.access)) {
                (_, Some(setter)) => {
                    let request = list.get_or_insert(super::lists::Request {
                        style: None,
                        start: None,
                    });
                    entry.prepare(Operation::Add, value, |_, value| {
                        if setter == Setter::ParagraphListStyle {
                            let text = match value {
                                Some(
                                    crate::value::Scalar::Text(text)
                                    | crate::value::Scalar::Token(text),
                                ) => text,
                                _ => {
                                    return Err(Error::new(
                                        Code::InvalidValue,
                                        "a list style is text",
                                    ));
                                },
                            };
                            request.style = Some(super::lists::style(text));
                        } else {
                            request.start = Some(super::lists::start(value)?);
                        }
                        Ok(())
                    })
                },
                (Access::Add(_), None) => prepare_run_entry(entry, value, &mut run, &mut has_run),
                _ => prepare_paragraph_entry(entry, value, &mut paragraph, &mut run, &mut has_run),
            };
            result.map_err(|error| property_error(error, key))?;
            continue;
        }
        let (_, entry) = RUN_REGISTRY
            .resolve(key.as_str())
            .map_err(|error| property_error(error, key))?;
        prepare_run_entry(entry, value, &mut run, &mut has_run)
            .map_err(|error| property_error(error, key))?;
    }
    if has_run {
        paragraph.push_run(run);
    }
    Ok((paragraph, list))
}

pub(super) fn new_run(props: &Props<'_>) -> Result<NewRun> {
    let mut run = NewRun::new("");
    let mut has_text = false;
    for (key, value) in props.iter() {
        let (_, entry) = RUN_REGISTRY
            .resolve(key.as_str())
            .map_err(|error| property_error(error, key))?;
        prepare_run_entry(entry, value, &mut run, &mut has_text)
            .map_err(|error| property_error(error, key))?;
    }
    Ok(run)
}

fn prepare_paragraph_entry(
    entry: &Entry<Setter>,
    value: &PropValue<'_>,
    paragraph: &mut NewParagraph,
    run: &mut NewRun,
    has_run: &mut bool,
) -> Result<()> {
    entry.prepare(Operation::Add, value, |setter, value| {
        if setter.is_text() {
            run_text(run, value)?;
            *has_run = true;
        } else {
            setter.apply_paragraph(paragraph.properties_mut(), value)?;
        }
        Ok(())
    })
}

fn prepare_run_entry(
    entry: &Entry<Setter>,
    value: &PropValue<'_>,
    run: &mut NewRun,
    has_run: &mut bool,
) -> Result<()> {
    entry.prepare(Operation::Add, value, |setter, value| {
        if setter.is_text() {
            run_text(run, value)?;
        } else {
            setter.apply_run(run.properties_mut(), value)?;
        }
        *has_run = true;
        Ok(())
    })
}

fn run_text(run: &mut NewRun, value: Option<crate::value::Scalar<'_>>) -> Result<()> {
    let text = match value {
        None => String::new(),
        Some(crate::value::Scalar::Text(text) | crate::value::Scalar::Token(text)) => {
            text.to_owned()
        },
        Some(_) => return Err(Error::new(Code::InvalidValue, "the text value is invalid")),
    };
    // NewRun intentionally exposes text only through its constructor.  Keep
    // the current typed properties while replacing its authored text.
    let properties = run.properties().clone();
    *run = NewRun::new(text).with_properties(properties);
    Ok(())
}

pub(super) fn body_anchor(
    docx: &Docx,
    placement: Option<&Placement<'_>>,
    _moving: bool,
) -> Result<BodyAnchor> {
    let body = super::nav::load_body(docx)?;
    let insertion_limit = body_insertion_limit(&body);
    let at = |position: Position| {
        let index = position.get();
        if index == insertion_limit {
            Ok(BodyAnchor::End)
        } else if index > insertion_limit {
            Err(out_of_range(
                "the DOCX body insertion position is outside the body",
            ))
        } else {
            before_block(
                body.get(position)
                    .ok_or_else(|| refusal("the body position is unavailable"))?,
            )
        }
    };
    match placement {
        None | Some(Placement::End) => Ok(BodyAnchor::End),
        Some(Placement::Start) => at(Position::new(0)),
        Some(Placement::Position(position)) => at(*position),
        Some(Placement::Before(path)) => {
            let node = resolve_path(docx, path)?;
            direct_body_anchor(&node, false)
        },
        Some(Placement::After(path)) => {
            let node = resolve_path(docx, path)?;
            direct_body_anchor(&node, true)
        },
        Some(Placement::Into(path, index)) => {
            let target = resolve_path(docx, path)?;
            if target.element_kind() != Kind::Body {
                return Err(refusal(
                    "a DOCX paragraph can only be added into the main body",
                ));
            }
            index.map_or(Ok(BodyAnchor::End), at)
        },
    }
}

fn run_position(
    docx: &Docx,
    parent: &Node<'_>,
    placement: Option<&Placement<'_>>,
) -> Result<Option<Position>> {
    let WriteSelection::Paragraph(paragraph) = parent.write_selection()? else {
        return Err(refusal("the run parent is not a direct body paragraph"));
    };
    match placement {
        None | Some(Placement::End) => Ok(None),
        Some(Placement::Start) => Ok(Some(Position::new(0))),
        Some(Placement::Position(position)) => Ok(Some(*position)),
        Some(Placement::Before(path)) => {
            let target = resolve_path(docx, path)?;
            match Selection::from_node(&target)? {
                Selection::Run {
                    paragraph: owner,
                    run,
                } if owner == paragraph => Ok(Some(run)),
                _ => Err(refusal(
                    "a run placement target must be a run in the same paragraph",
                )),
            }
        },
        Some(Placement::After(path)) => {
            let target = resolve_path(docx, path)?;
            match Selection::from_node(&target)? {
                Selection::Run {
                    paragraph: owner,
                    run,
                } if owner == paragraph => run
                    .get()
                    .checked_add(1)
                    .map(Position::new)
                    .ok_or_else(|| {
                        Error::new(
                            Code::LimitExceeded,
                            "the run position exceeds the document limit",
                        )
                        .with_format(FileFormat::Docx)
                    })
                    .map(Some),
                _ => Err(refusal(
                    "a run placement target must be a run in the same paragraph",
                )),
            }
        },
        Some(Placement::Into(path, index)) => {
            let target = resolve_path(docx, path)?;
            if target.element_kind() != Kind::Paragraph {
                return Err(refusal("a run can only be placed into a paragraph"));
            }
            let WriteSelection::Paragraph(owner) = target.write_selection()? else {
                return Err(refusal("the run parent is not a direct body paragraph"));
            };
            if owner != paragraph {
                return Err(refusal("a run placement parent must match the add parent"));
            }
            Ok(*index)
        },
    }
}

/// Resolve a run move destination while the source snapshot is still
/// unchanged.  Positional destinations are intentionally represented as
/// `Before(position)`: the leaf interprets that position after detaching the
/// source run, exactly as the public placement contract specifies.
fn run_destination(
    docx: &Docx,
    source_paragraph: Position,
    source_run: Position,
    placement: &Placement<'_>,
) -> Result<(Position, RunAnchor)> {
    match placement {
        Placement::Start => Ok((
            source_paragraph,
            run_position_anchor(source_run, Position::new(0))?,
        )),
        Placement::End => Ok((source_paragraph, RunAnchor::End)),
        Placement::Position(position) => Ok((
            source_paragraph,
            run_position_anchor(source_run, *position)?,
        )),
        Placement::Before(path) => {
            let target = resolve_path(docx, path)?;
            match Selection::from_node(&target)? {
                Selection::Run { paragraph, run } => Ok((paragraph, RunAnchor::Before(run))),
                _ => Err(refusal("a run move target must be a direct run")),
            }
        },
        Placement::After(path) => {
            let target = resolve_path(docx, path)?;
            match Selection::from_node(&target)? {
                Selection::Run { paragraph, run } => Ok((paragraph, RunAnchor::After(run))),
                _ => Err(refusal("a run move target must be a direct run")),
            }
        },
        Placement::Into(path, index) => {
            let target = resolve_path(docx, path)?;
            if target.element_kind() != Kind::Paragraph {
                return Err(refusal("a run can only move into a paragraph"));
            }
            let WriteSelection::Paragraph(paragraph) = target.write_selection()? else {
                return Err(refusal("a run move parent is not a direct body paragraph"));
            };
            let anchor = match index {
                None => RunAnchor::End,
                Some(position) if paragraph == source_paragraph => {
                    run_position_anchor(source_run, *position)?
                },
                Some(position) => RunAnchor::Before(*position),
            };
            Ok((paragraph, anchor))
        },
    }
}

/// Convert the public position-after-detach contract to the leaf's
/// pre-detach anchor.  A destination after the source must skip the source's
/// old slot when expressed against the current paragraph; an equal position is
/// already the source's unchanged slot.
fn run_position_anchor(source: Position, destination: Position) -> Result<RunAnchor> {
    if destination < source {
        return Ok(RunAnchor::Before(destination));
    }
    if destination == source {
        return Ok(RunAnchor::Before(source));
    }
    let anchor = destination
        .get()
        .checked_add(1)
        .map(Position::new)
        .ok_or_else(|| out_of_range("the DOCX run destination exceeds the position limit"))?;
    Ok(RunAnchor::Before(anchor))
}

fn move_anchor(
    docx: &Docx,
    node: &Node<'_>,
    source: Position,
    placement: &Placement<'_>,
) -> Result<BodyAnchor> {
    match placement {
        Placement::Start => {
            let body = super::nav::load_body(docx)?;
            let remaining = body_without_source(&body, source)?;
            if body_insertion_limit_from(&remaining) == 0 {
                Ok(BodyAnchor::End)
            } else {
                before_block(
                    remaining
                        .first()
                        .ok_or_else(|| refusal("the DOCX move destination is unavailable"))?
                        .clone(),
                )
            }
        },
        Placement::Position(position) => {
            let body = super::nav::load_body(docx)?;
            let remaining = body_without_source(&body, source)?;
            let index = position.get();
            let insertion_limit = body_insertion_limit_from(&remaining);
            if index == insertion_limit {
                Ok(BodyAnchor::End)
            } else if index > insertion_limit {
                Err(out_of_range(
                    "the DOCX move destination is outside the body",
                ))
            } else {
                before_block(
                    remaining
                        .get(index)
                        .ok_or_else(|| refusal("the DOCX move destination is unavailable"))?
                        .clone(),
                )
            }
        },
        Placement::End => Ok(BodyAnchor::End),
        Placement::Before(path) => {
            let target = resolve_path(docx, path)?;
            direct_body_anchor(&target, false)
        },
        Placement::After(path) => {
            let target = resolve_path(docx, path)?;
            direct_body_anchor(&target, true)
        },
        Placement::Into(path, index) => {
            let parent = resolve_path(docx, path)?;
            if parent.element_kind() != Kind::Body {
                return Err(refusal(
                    "a DOCX paragraph can only move within the main body",
                ));
            }
            match index {
                None => Ok(BodyAnchor::End),
                Some(index) => {
                    let body = super::nav::load_body(docx)?;
                    let remaining = body_without_source(&body, source)?;
                    let insertion_limit = body_insertion_limit_from(&remaining);
                    if index.get() == insertion_limit {
                        Ok(BodyAnchor::End)
                    } else if index.get() > insertion_limit {
                        Err(out_of_range(
                            "the DOCX move destination is outside the body",
                        ))
                    } else {
                        before_block(
                            remaining
                                .get(index.get())
                                .ok_or_else(|| refusal("the DOCX move destination is unavailable"))?
                                .clone(),
                        )
                    }
                },
            }
        },
    }
    .and_then(|anchor| {
        // A placement target must remain in the same direct body as the
        // source.  `node` is intentionally consumed only for this proof.
        if !is_direct_body(node) {
            Err(refusal("the DOCX source is not a direct body paragraph"))
        } else {
            Ok(anchor)
        }
    })
}

fn body_without_source(body: &litchi_docx::Blocks, source: Position) -> Result<Vec<IndexedBlock>> {
    let mut remaining = Vec::with_capacity(body.len().saturating_sub(1));
    for block in body.iter() {
        if matches!(block.kind(), BlockKind::Paragraph) && block.transaction_position()? == source {
            continue;
        }
        remaining.push(block);
    }
    Ok(remaining)
}

fn body_insertion_limit(body: &litchi_docx::Blocks) -> usize {
    body.iter()
        .position(|block| block.kind() == BlockKind::SectionProperties)
        .unwrap_or_else(|| body.len())
}

fn body_insertion_limit_from(blocks: &[IndexedBlock]) -> usize {
    blocks
        .iter()
        .position(|block| block.kind() == BlockKind::SectionProperties)
        .unwrap_or(blocks.len())
}

fn direct_body_anchor(node: &Node<'_>, after: bool) -> Result<BodyAnchor> {
    if !is_direct_body(node) {
        return Err(refusal(
            "the DOCX placement target is not a direct body block",
        ));
    }
    let position = node.transaction_position()?;
    let block = match node.element_kind() {
        Kind::Paragraph | Kind::Equation => BodyBlock::Paragraph(position),
        Kind::Table => BodyBlock::Table(position),
        Kind::Sdt => BodyBlock::ContentControl(position),
        _ => {
            return Err(refusal(
                "the DOCX placement target is not a movable body block",
            ));
        },
    };
    Ok(if after {
        BodyAnchor::After(block)
    } else {
        BodyAnchor::Before(block)
    })
}

fn before_block(block: IndexedBlock) -> Result<BodyAnchor> {
    let position = block.transaction_position().map_err(Error::from)?;
    Ok(match block.kind() {
        BlockKind::Paragraph | BlockKind::Equation => {
            BodyAnchor::Before(BodyBlock::Paragraph(position))
        },
        BlockKind::Table => BodyAnchor::Before(BodyBlock::Table(position)),
        BlockKind::ContentControl => BodyAnchor::Before(BodyBlock::ContentControl(position)),
        _ => {
            return Err(refusal(
                "the DOCX body position is not an addressable block",
            ));
        },
    })
}

fn is_direct_body(node: &Node<'_>) -> bool {
    let path = node.canonical_path();
    let steps = path.steps();
    steps.len() == 2 && steps.first().is_some_and(|step| step.name() == "body")
}

fn admit_remove(nodes: &[Node<'_>]) -> Result<Vec<Selection>> {
    let mut selections = Vec::with_capacity(nodes.len());
    for node in nodes {
        match node.element_kind() {
            Kind::Paragraph | Kind::Run => selections.push(Selection::from_node(node)?),
            _ => {
                return Err(refusal(
                    "the DOCX structural remove target must be a paragraph or run",
                ));
            },
        }
    }
    let paragraphs: Vec<Position> = selections
        .iter()
        .filter_map(|selection| match selection {
            Selection::Paragraph(position) => Some(*position),
            Selection::Run { .. } => None,
        })
        .collect();
    selections.retain(|selection| match selection {
        Selection::Paragraph(_) => true,
        Selection::Run { paragraph, .. } => !paragraphs.contains(paragraph),
    });
    selections.sort_by_key(|selection| std::cmp::Reverse(selection.key()));
    for pair in selections.windows(2) {
        if pair[0].key() == pair[1].key() {
            return Err(Error::new(
                Code::InvalidSelector,
                "the DOCX remove selector addressed one element more than once",
            )
            .with_format(FileFormat::Docx));
        }
    }
    Ok(selections)
}

fn paragraph_path(position: Position) -> Path<'static> {
    Path::from_steps([
        Step::new("body"),
        Step::new("p").with(Predicate::Index(position)),
    ])
}

fn run_path(paragraph: Position, run: Position) -> Path<'static> {
    Path::from_steps([
        Step::new("body"),
        Step::new("p").with(Predicate::Index(paragraph)),
        Step::new("r").with(Predicate::Index(run)),
    ])
}

fn property_error(error: Error, key: &Key<'_>) -> Error {
    error
        .with_property(key.as_str().to_owned())
        .with_format(FileFormat::Docx)
}

fn refusal(message: &'static str) -> Error {
    Error::new(Code::UnsupportedCapability, message).with_format(FileFormat::Docx)
}

fn out_of_range(message: &'static str) -> Error {
    Error::new(Code::OutOfRange, message).with_format(FileFormat::Docx)
}

fn internal(message: &'static str) -> Error {
    Error::new(Code::InternalError, message).with_format(FileFormat::Docx)
}
