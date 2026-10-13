//! Structural verbs beyond the direct body (record 0251a): paragraphs and
//! runs inside cells, content controls and hyperlinks, and table, row and
//! cell structure.
//!
//! Each verb converts its nodes to proven leaf element paths
//! ([`super::element`]) and stages the leaf's path-addressed edits. Visible
//! positions of new elements are read back from the committed handle, so a
//! returned path always addresses what was added.

use litchi_core::selector::Path;
use litchi_core::{FileFormat, Position};
use litchi_docx::document::{ElementAnchor, ElementPath, ElementStep, NewTable};

use super::element::{
    element_path, path_with, raw_position, refusal, run_locations, visible_ordinal,
};
use super::nav::{Nav, container_of};
use super::node::{Node, Value};
use super::spec::{Kind, ROW_REGISTRY, TABLE_REGISTRY};
use super::structure::{NewKind, new_paragraph, new_run, resolve_path};
use super::write::{Commit, WorkEdit};
use super::{Docx, DocxEditor};
use crate::error::{Code, Error, Result};
use crate::new_node::NewNode;
use crate::path::PathArg;
use crate::placement::Placement;
use crate::props::Props;
use crate::registry::Operation;

impl Docx {
    /// Add a paragraph, run, table or row whose parent is not on the direct
    /// body route.
    pub(super) fn add_nested(
        &self,
        editor: DocxEditor,
        parent: &Node<'_>,
        kind: NewKind,
        node: &NewNode<'_>,
    ) -> Result<(Commit, Path<'static>)> {
        let placement = node.placement();
        match (kind, parent.element_kind()) {
            (NewKind::Paragraph, Kind::Body | Kind::Cell | Kind::Sdt) => {
                let container = container_path(parent)?;
                let anchor = block_anchor(self, parent, &container, placement, None)?;
                let paragraph = new_paragraph(node.props())?;
                self.owner().check_source()?;
                let mut edit = WorkEdit::from(editor)?;
                let raw = edit.insert_paragraph_at(&container, anchor, &paragraph)?;
                let commit = edit.commit(self)?;
                let ordinal = visible_child(self, commit.file(), parent, Kind::Paragraph, raw)?;
                Ok((commit, path_with(parent, Some((Kind::Paragraph, ordinal)))))
            },
            (NewKind::Table, Kind::Body | Kind::Cell | Kind::Sdt) => {
                let container = container_path(parent)?;
                let anchor = block_anchor(self, parent, &container, placement, None)?;
                let table = new_table(node.props())?;
                self.owner().check_source()?;
                let mut edit = WorkEdit::from(editor)?;
                let raw = edit.insert_table_at(&container, anchor, &table)?;
                let commit = edit.commit(self)?;
                let ordinal = visible_child(self, commit.file(), parent, Kind::Table, raw)?;
                Ok((commit, path_with(parent, Some((Kind::Table, ordinal)))))
            },
            (NewKind::Row, Kind::Table) => {
                let table = element_path(parent)?;
                let rows = Nav::rows(parent)?.len();
                let at = sibling_position(self, parent, Kind::Row, rows, placement)?;
                let cells = row_cells(node.props())?;
                self.owner().check_source()?;
                let mut edit = WorkEdit::from(editor)?;
                let raw = edit.insert_row_at(&table, at.map(Position::new), cells)?;
                let commit = edit.commit(self)?;
                Ok((commit, path_with(parent, Some((Kind::Row, raw)))))
            },
            (NewKind::Hyperlink, Kind::Paragraph) => {
                let owner = element_path(parent)?;
                let Value::Paragraph { paragraph, .. } = &parent.value else {
                    return Err(refusal("the hyperlink parent holds no paragraph"));
                };
                let locations = run_locations(self, paragraph)?;
                let at = sibling_position(self, parent, Kind::Run, locations.len(), placement)?;
                let direct = match at {
                    None => None,
                    Some(index) if index == locations.len() => None,
                    Some(index) => match locations.get(index).map(Vec::as_slice) {
                        Some([ElementStep::Run(run)]) => Some(*run),
                        _ => {
                            return Err(refusal(
                                "a hyperlink can only be added beside a direct run of the paragraph",
                            ));
                        },
                    },
                };
                let (target, run) = new_hyperlink(node.props())?;
                // The hyperlink's position is its ordinal among the
                // paragraph's direct hyperlinks.
                self.owner().check_source()?;
                let mut edit = WorkEdit::from(self.eager_editor(editor)?)?;
                let WorkEdit::Eager(eager) = &mut edit else {
                    return Err(refusal("hyperlinks need the owning package editor"));
                };
                let position = eager
                    .insert_hyperlink_at(&owner, direct, &target, &run)
                    .map_err(Error::from)?;
                let commit = edit.commit(self)?;
                Ok((commit, path_with(parent, Some((Kind::Hyperlink, position)))))
            },
            (NewKind::Picture, Kind::Paragraph) => {
                let owner = element_path(parent)?;
                let Value::Paragraph { paragraph, .. } = &parent.value else {
                    return Err(refusal("the picture parent holds no paragraph"));
                };
                let locations = run_locations(self, paragraph)?;
                let at = sibling_position(self, parent, Kind::Run, locations.len(), placement)?;
                let direct = match at {
                    None => None,
                    Some(index) if index == locations.len() => None,
                    Some(index) => match locations.get(index).map(Vec::as_slice) {
                        Some([ElementStep::Run(run)]) => Some(*run),
                        _ => {
                            return Err(refusal(
                                "a picture can only be added beside a direct run of the paragraph",
                            ));
                        },
                    },
                };
                let picture = super::picture::new_picture(node.props())?;
                self.owner().check_source()?;
                let mut edit = WorkEdit::from(self.eager_editor(editor)?)?;
                let WorkEdit::Eager(eager) = &mut edit else {
                    return Err(refusal("pictures need the owning package editor"));
                };
                let ordinal = eager
                    .insert_picture_at(&owner, direct, &picture)
                    .map_err(Error::from)?;
                let commit = edit.commit(self)?;
                Ok((commit, path_with(parent, Some((Kind::Picture, ordinal)))))
            },
            (NewKind::Comment, Kind::Paragraph | Kind::Run) => {
                // A paragraph parent comments the whole paragraph; a run
                // parent comments that direct run.
                let path = element_path(parent)?;
                let (paragraph, range) = if parent.element_kind() == Kind::Paragraph {
                    (path, None)
                } else {
                    match path.split_last() {
                        Some((paragraph, ElementStep::Run(run))) => (paragraph, Some((run, run))),
                        _ => {
                            return Err(refusal(
                                "a comment can only be anchored on a direct run of a paragraph",
                            ));
                        },
                    }
                };
                let comment = new_comment(node.props())?;
                self.owner().check_source()?;
                let mut edit = WorkEdit::from(self.eager_editor(editor)?)?;
                let WorkEdit::Eager(eager) = &mut edit else {
                    return Err(refusal("comments need the owning package editor"));
                };
                let id = eager
                    .add_comment_at(&paragraph, range, &comment)
                    .map_err(Error::from)?;
                let commit = edit.commit(self)?;
                let ordinal = super::nav::comments(commit.file())?
                    .iter()
                    .position(|comment| comment.id() == id)
                    .ok_or_else(|| refusal("the new comment is not in the comments part"))?;
                Ok((commit, comment_path(ordinal)))
            },
            (NewKind::Run, Kind::Paragraph) => {
                let owner = element_path(parent)?;
                let Value::Paragraph { paragraph, .. } = &parent.value else {
                    return Err(refusal("the run parent holds no paragraph"));
                };
                let locations = run_locations(self, paragraph)?;
                let at = sibling_position(self, parent, Kind::Run, locations.len(), placement)?;
                let direct = match at {
                    None => None,
                    Some(index) if index == locations.len() => None,
                    Some(index) => match locations.get(index).map(Vec::as_slice) {
                        Some([ElementStep::Run(run)]) => Some(*run),
                        _ => {
                            return Err(refusal(
                                "a run can only be added beside a direct run of the paragraph",
                            ));
                        },
                    },
                };
                let ordinal = at.unwrap_or(locations.len()).min(locations.len());
                let run = new_run(node.props())?;
                self.owner().check_source()?;
                let mut edit = WorkEdit::from(editor)?;
                edit.insert_run_at(&owner, direct, &run)?;
                let commit = edit.commit(self)?;
                Ok((
                    commit,
                    path_with(parent, Some((Kind::Run, Position::new(ordinal)))),
                ))
            },
            _ => Err(Error::new(
                Code::UnsupportedCapability,
                "the DOCX element cannot be added under this parent",
            )
            .with_format(FileFormat::Docx)),
        }
    }

    /// Remove nodes beyond the direct-body paragraph/run route: nested
    /// paragraphs and runs, hyperlinks, tables, rows, cells and content
    /// controls. Descendants of a removed node are subsumed, and removals
    /// run in reverse raw order so no removal shifts a later one.
    pub(super) fn remove_nested(&self, editor: DocxEditor, nodes: &[Node<'_>]) -> Result<Commit> {
        let mut paths = Vec::with_capacity(nodes.len());
        let mut comments = Vec::new();
        for node in nodes {
            match node.element_kind() {
                Kind::Comment => {
                    let Value::Comment(comment) = &node.value else {
                        return Err(refusal("the comment node holds no comment"));
                    };
                    comments.push(comment.id());
                },
                Kind::Paragraph
                | Kind::Run
                | Kind::Hyperlink
                | Kind::Table
                | Kind::Row
                | Kind::Cell
                | Kind::Sdt => paths.push(element_path(node)?),
                _ => {
                    return Err(Error::new(
                        Code::UnsupportedCapability,
                        "the DOCX element cannot be removed",
                    )
                    .with_format(FileFormat::Docx));
                },
            }
        }
        let mut keys: Vec<(Vec<(u8, usize)>, ElementPath)> = paths
            .into_iter()
            .map(|path| (order_key(&path), path))
            .collect();
        keys.sort_by(|left, right| right.0.cmp(&left.0));
        keys.dedup_by(|left, right| left.0 == right.0);
        let removable: Vec<ElementPath> = keys
            .iter()
            .filter(|(key, _)| {
                !keys
                    .iter()
                    .any(|(other, _)| other.len() < key.len() && key.starts_with(other))
            })
            .map(|(_, path)| path.clone())
            .collect();
        self.owner().check_source()?;
        let editor = if comments.is_empty() {
            editor
        } else {
            self.eager_editor(editor)?
        };
        let mut edit = WorkEdit::from(editor)?;
        comments.sort_unstable();
        comments.dedup();
        for id in comments {
            let WorkEdit::Eager(eager) = &mut edit else {
                return Err(refusal("comments need the owning package editor"));
            };
            eager.remove_comment(id).map_err(Error::from)?;
        }
        for path in &removable {
            edit.remove_element_at(path)?;
        }
        edit.commit(self)
    }

    /// Move a table, row, nested paragraph or content control.
    pub(super) fn move_nested(
        &self,
        editor: DocxEditor,
        node: &Node<'_>,
        placement: &Placement<'_>,
    ) -> Result<Commit> {
        let path = element_path(node)?;
        let (parent, step) = path
            .split_last()
            .ok_or_else(|| refusal("the DOCX body cannot move"))?;
        let (container, anchor) = match placement {
            Placement::Before(target) | Placement::After(target) => {
                let target = element_path(&resolve_path(self, target)?)?;
                let (container, sibling) = target
                    .split_last()
                    .ok_or_else(|| refusal("the move target is the body"))?;
                let anchor = if matches!(placement, Placement::Before(_)) {
                    ElementAnchor::Before(sibling)
                } else {
                    ElementAnchor::After(sibling)
                };
                (container, anchor)
            },
            Placement::Start => (parent.clone(), ElementAnchor::Start),
            Placement::End => (parent.clone(), ElementAnchor::End),
            Placement::Position(position) => {
                let parent_node = parent_node(self, node)?;
                let anchor = moved_anchor(self, &parent_node, node, step, *position)?;
                (parent.clone(), anchor)
            },
            Placement::Into(target, index) => {
                let target_node = resolve_path(self, target)?;
                let container = container_path(&target_node)?;
                let anchor = match index {
                    None => ElementAnchor::End,
                    Some(position) if container == parent => {
                        moved_anchor(self, &target_node, node, step, *position)?
                    },
                    Some(position) => {
                        block_anchor(self, &target_node, &container, None, Some(*position))?
                    },
                };
                (container, anchor)
            },
        };
        self.owner().check_source()?;
        let mut edit = WorkEdit::from(editor)?;
        edit.move_element_at(&path, &container, anchor)?;
        edit.commit(self)
    }
}

/// The element path of a block container node; the body is the empty path.
fn container_path(node: &Node<'_>) -> Result<ElementPath> {
    if node.element_kind() == Kind::Body {
        Ok(ElementPath::body())
    } else {
        element_path(node)
    }
}

/// The parent node of `node`, resolved from its canonical path.
fn parent_node<'f>(docx: &'f Docx, node: &Node<'_>) -> Result<Node<'f>> {
    let path = path_with(node, None);
    let steps = path.steps();
    let parent = Path::from_steps(steps[..steps.len().saturating_sub(1)].to_vec());
    resolve_path(docx, &PathArg::from(parent))
}

/// The visible ordinal the `kind` block inserted at raw position `raw` into
/// `parent` takes: the number of visible blocks of `kind` before it, read
/// from the container's index before the insertion. When a visible block of
/// `kind` has no proven raw position the committed handle is read instead.
fn visible_child(
    docx: &Docx,
    file: &Docx,
    parent: &Node<'_>,
    kind: Kind,
    raw: Position,
) -> Result<Position> {
    let blocks =
        container_of(parent)?.ok_or_else(|| refusal("the DOCX parent is not a block container"))?;
    let mut before = 0usize;
    let mut proven = true;
    for block in blocks
        .iter()
        .filter(|block| Kind::of_block(block.kind()) == Some(kind))
    {
        match raw_position(docx, &block) {
            Ok(position) if position < raw => before += 1,
            Ok(_) => {},
            Err(_) => {
                proven = false;
                break;
            },
        }
    }
    if proven {
        return Ok(Position::new(before));
    }
    let parent = resolve_path(file, &PathArg::from(path_with(parent, None)))?;
    let blocks = container_of(&parent)?
        .ok_or_else(|| refusal("the DOCX parent is not a block container"))?;
    visible_ordinal(file, &blocks, kind, raw)
        .map(Position::new)
        .ok_or_else(|| {
            Error::new(
                Code::InternalError,
                "the added DOCX block is missing from its container's visible view",
            )
            .with_format(FileFormat::Docx)
        })
}

/// The leaf anchor of a block placement in a block container.
fn block_anchor(
    docx: &Docx,
    parent: &Node<'_>,
    container: &ElementPath,
    placement: Option<&Placement<'_>>,
    index: Option<Position>,
) -> Result<ElementAnchor> {
    let at_visible = |position: Position| -> Result<ElementAnchor> {
        let blocks =
            container_of(parent)?.ok_or_else(|| refusal("the DOCX parent holds no blocks"))?;
        let addressable: Vec<_> = blocks
            .iter()
            .filter(|block| {
                matches!(
                    Kind::of_block(block.kind()),
                    Some(Kind::Paragraph | Kind::Table | Kind::Sdt | Kind::Equation)
                )
            })
            .collect();
        match addressable.get(position.get()) {
            None if position.get() == addressable.len() => Ok(ElementAnchor::End),
            None => Err(Error::new(
                Code::OutOfRange,
                "the DOCX insertion position is outside its container",
            )
            .with_format(FileFormat::Docx)),
            Some(block) => Ok(ElementAnchor::Before(raw_step(docx, block)?)),
        }
    };
    if let Some(position) = index {
        return at_visible(position);
    }
    match placement {
        None | Some(Placement::End) => Ok(ElementAnchor::End),
        Some(Placement::Start) => Ok(ElementAnchor::Start),
        Some(Placement::Position(position)) => at_visible(*position),
        Some(Placement::Before(path) | Placement::After(path)) => {
            let sibling = element_path(&resolve_path(docx, path)?)?;
            let (owner, step) = sibling
                .split_last()
                .ok_or_else(|| refusal("the placement target is the body"))?;
            if owner != *container {
                return Err(refusal(
                    "a placement target must be a child of the same container",
                ));
            }
            Ok(if matches!(placement, Some(Placement::Before(_))) {
                ElementAnchor::Before(step)
            } else {
                ElementAnchor::After(step)
            })
        },
        Some(Placement::Into(path, index)) => {
            let target = resolve_path(docx, path)?;
            if container_path(&target)? != *container {
                return Err(refusal("the placement parent must match the add parent"));
            }
            index.map_or(Ok(ElementAnchor::End), at_visible)
        },
    }
}

/// The raw leaf step of a visible block.
fn raw_step(docx: &Docx, block: &litchi_docx::IndexedBlock) -> Result<ElementStep> {
    let position = raw_position(docx, block)?;
    Ok(match block.kind() {
        litchi_docx::BlockKind::Paragraph | litchi_docx::BlockKind::Equation => {
            ElementStep::Paragraph(position)
        },
        litchi_docx::BlockKind::Table => ElementStep::Table(position),
        litchi_docx::BlockKind::ContentControl => ElementStep::ContentControl(position),
        _ => return Err(refusal("the DOCX block is not an anchor")),
    })
}

/// The anchor that moves `node` (whose leaf step is `step`) to visible
/// position `position` among its parent's children after it is taken out.
fn moved_anchor(
    docx: &Docx,
    parent: &Node<'_>,
    node: &Node<'_>,
    step: ElementStep,
    position: Position,
) -> Result<ElementAnchor> {
    if let ElementStep::Row(source) = step {
        let rows = Nav::rows(parent)?.len();
        let target = position.get();
        if target >= rows {
            return Err(Error::new(
                Code::OutOfRange,
                "the DOCX row destination is outside its table",
            )
            .with_format(FileFormat::Docx));
        }
        return Ok(if target < source.get() {
            ElementAnchor::Before(ElementStep::Row(Position::new(target)))
        } else if target + 1 == rows {
            ElementAnchor::End
        } else {
            ElementAnchor::Before(ElementStep::Row(Position::new(target + 1)))
        });
    }
    let blocks = container_of(parent)?.ok_or_else(|| refusal("the DOCX parent holds no blocks"))?;
    let source = element_path(node)?;
    let mut remaining = Vec::new();
    for block in blocks.iter().filter(|block| {
        matches!(
            Kind::of_block(block.kind()),
            Some(Kind::Paragraph | Kind::Table | Kind::Sdt | Kind::Equation)
        )
    }) {
        let step = raw_step(docx, &block)?;
        if source.steps().last() != Some(&step) {
            remaining.push(step);
        }
    }
    Ok(match remaining.get(position.get()) {
        Some(step) => ElementAnchor::Before(*step),
        None if position.get() == remaining.len() => ElementAnchor::End,
        None => {
            return Err(Error::new(
                Code::OutOfRange,
                "the DOCX move destination is outside its container",
            )
            .with_format(FileFormat::Docx));
        },
    })
}

/// The visible sibling position a run or row placement names, as an
/// insertion index among `count` existing siblings of `kind` (`None` is the
/// end).
fn sibling_position(
    docx: &Docx,
    parent: &Node<'_>,
    kind: Kind,
    count: usize,
    placement: Option<&Placement<'_>>,
) -> Result<Option<usize>> {
    let sibling = |path: &PathArg<'_>| -> Result<usize> {
        let target = resolve_path(docx, path)?;
        let parent_path = path_with(parent, None);
        let target_path = path_with(&target, None);
        let steps = target_path.steps();
        if target.element_kind() != kind
            || steps.len() != parent_path.steps().len() + 1
            || steps[..steps.len() - 1] != *parent_path.steps()
        {
            return Err(refusal(
                "a placement target must be a sibling of the same kind in the same parent",
            ));
        }
        Ok(target
            .address
            .steps()
            .last()
            .map_or(0, |step| step.ordinal as usize))
    };
    let index = match placement {
        None | Some(Placement::End) => None,
        Some(Placement::Start) => Some(0),
        Some(Placement::Position(position)) => Some(position.get()),
        Some(Placement::Before(path)) => Some(sibling(path)?),
        Some(Placement::After(path)) => Some(sibling(path)? + 1),
        Some(Placement::Into(path, index)) => {
            let target = resolve_path(docx, path)?;
            if path_with(&target, None) != path_with(parent, None) {
                return Err(refusal("the placement parent must match the add parent"));
            }
            index.map(Position::get)
        },
    };
    if index.is_some_and(|index| index > count) {
        return Err(Error::new(
            Code::OutOfRange,
            "the DOCX insertion position is outside its parent",
        )
        .with_format(FileFormat::Docx));
    }
    Ok(index)
}

/// The document-order sort key of an element path: the raw position of
/// each step, ranked by element name.
fn order_key(path: &ElementPath) -> Vec<(u8, usize)> {
    path.steps()
        .iter()
        .map(|step| {
            let rank = match step {
                ElementStep::Paragraph(_) => 0,
                ElementStep::Table(_) => 1,
                ElementStep::ContentControl(_) => 2,
                ElementStep::Row(_) => 3,
                ElementStep::Cell(_) => 4,
                ElementStep::Hyperlink(_) => 5,
                ElementStep::Run(_) => 6,
                _ => 7,
            };
            (rank, step.position().get())
        })
        .collect()
}

/// A new table from the add properties: `rows`, `cols`, `data`, `style` and
/// `colWidths`.
fn new_table(props: &Props<'_>) -> Result<NewTable> {
    let mut rows = None::<usize>;
    let mut columns = None::<usize>;
    let mut data = None::<Vec<Vec<String>>>;
    let mut style = None::<String>;
    let mut widths = None::<Vec<u32>>;
    for (key, value) in props.iter() {
        let (_, entry) = TABLE_REGISTRY
            .resolve(key.as_str())
            .map_err(|error| error.with_property(key.as_str().to_owned()))?;
        entry
            .prepare(Operation::Add, value, |setter, value| {
                let text = match value {
                    Some(crate::value::Scalar::Text(text) | crate::value::Scalar::Token(text)) => {
                        text.to_owned()
                    },
                    Some(scalar) => scalar.to_string(),
                    None => String::new(),
                };
                use super::setters::Setter;
                match setter {
                    Setter::TableRows => rows = Some(count(&text)?),
                    Setter::TableColumns => columns = Some(count(&text)?),
                    Setter::TableStyle => style = Some(text),
                    Setter::TableColumnWidths => {
                        widths = Some(
                            text.split(',')
                                .map(|width| {
                                    width.trim().parse::<u32>().map_err(|_| {
                                        Error::new(
                                            Code::InvalidValue,
                                            "table column widths are comma-separated twips",
                                        )
                                    })
                                })
                                .collect::<Result<_>>()?,
                        );
                    },
                    Setter::TableData => {
                        data = Some(
                            text.split(';')
                                .map(|row| row.split(',').map(str::to_owned).collect())
                                .collect(),
                        );
                    },
                    _ => {
                        return Err(Error::new(
                            Code::UnsupportedProperty,
                            "the property cannot seed a new table",
                        ));
                    },
                }
                Ok(())
            })
            .map_err(|error| {
                error
                    .with_property(key.as_str().to_owned())
                    .with_format(FileFormat::Docx)
            })?;
    }
    let data_rows = data.as_ref().map(Vec::len);
    let data_columns = data
        .as_ref()
        .and_then(|data| data.iter().map(Vec::len).max());
    let rows = rows.or(data_rows).ok_or_else(|| missing("rows"))?;
    let columns = columns
        .or(widths.as_ref().map(Vec::len))
        .or(data_columns)
        .ok_or_else(|| missing("cols"))?;
    let mut table = NewTable::new(rows, columns)?;
    if let Some(widths) = widths {
        table = table.with_column_widths(widths)?;
    }
    if let Some(style) = style {
        table = table.with_style(style);
    }
    for (row, cells) in data.into_iter().flatten().enumerate() {
        for (column, text) in cells.into_iter().enumerate() {
            table = table.with_cell_text(row, column, text);
        }
    }
    Ok(table)
}

/// A new hyperlink's target and run from its add properties: exactly one of
/// `url` and `anchor`, its `text`, and its run formatting. Without explicit
/// formatting the run gets OfficeCLI's fallback hyperlink look (color
/// `0563C1`, single underline).
fn new_hyperlink(
    props: &Props<'_>,
) -> Result<(
    litchi_docx::document::HyperlinkTarget,
    litchi_docx::document::NewRun,
)> {
    use super::setters::Setter;
    let mut url = None::<String>;
    let mut anchor = None::<String>;
    let mut text = String::new();
    let mut properties = litchi_docx::properties::Run::default();
    let mut styled = false;
    for (key, value) in props.iter() {
        let (_, entry) = super::spec::HYPERLINK_REGISTRY
            .resolve(key.as_str())
            .map_err(|error| error.with_property(key.as_str().to_owned()))?;
        entry
            .prepare(Operation::Add, value, |setter, value| {
                let as_text = |value: Option<crate::value::Scalar<'_>>| match value {
                    Some(crate::value::Scalar::Text(text) | crate::value::Scalar::Token(text)) => {
                        text.to_owned()
                    },
                    Some(scalar) => scalar.to_string(),
                    None => String::new(),
                };
                match setter {
                    Setter::HyperlinkUrl => url = Some(as_text(value)),
                    Setter::HyperlinkAnchor => anchor = Some(as_text(value)),
                    Setter::HyperlinkText => text = as_text(value),
                    other => {
                        styled |= matches!(
                            other,
                            Setter::RunColor | Setter::RunUnderline | Setter::RunStyle
                        );
                        other.apply_run(&mut properties, value)?;
                    },
                }
                Ok(())
            })
            .map_err(|error| {
                error
                    .with_property(key.as_str().to_owned())
                    .with_format(FileFormat::Docx)
            })?;
    }
    let target = match (url, anchor) {
        (Some(url), None) => litchi_docx::document::HyperlinkTarget::Url(url),
        (None, Some(anchor)) => litchi_docx::document::HyperlinkTarget::Anchor(anchor),
        _ => {
            return Err(Error::new(
                Code::MissingProperty,
                "a new DOCX hyperlink needs exactly one of url and anchor",
            )
            .with_property("url")
            .with_format(FileFormat::Docx));
        },
    };
    if !styled {
        Setter::RunColor.apply_run(
            &mut properties,
            Some(crate::value::Scalar::Color(crate::value::Color::rgb(
                0x05, 0x63, 0xC1,
            ))),
        )?;
        Setter::RunUnderline
            .apply_run(&mut properties, Some(crate::value::Scalar::Token("single")))?;
    }
    Ok((
        target,
        litchi_docx::document::NewRun::new(text).with_properties(properties),
    ))
}

/// The unified path of the comment at `ordinal` in the comments part.
fn comment_path(ordinal: usize) -> Path<'static> {
    use litchi_core::selector::{Predicate, Step};
    Path::from_steps([
        Step::new("comments"),
        Step::new("comment").with(Predicate::Index(Position::new(ordinal))),
    ])
}

/// A new comment from its add properties: `author` is required, `text`,
/// `initials` and `date` are optional.
fn new_comment(props: &Props<'_>) -> Result<litchi_docx::document::NewComment> {
    use super::setters::Setter;
    let mut author = None::<String>;
    let mut initials = None::<String>;
    let mut date = None::<String>;
    let mut text = String::new();
    for (key, value) in props.iter() {
        let (_, entry) = super::spec::COMMENT_REGISTRY
            .resolve(key.as_str())
            .map_err(|error| error.with_property(key.as_str().to_owned()))?;
        entry
            .prepare(Operation::Add, value, |setter, value| {
                let value = match value {
                    Some(crate::value::Scalar::Text(text) | crate::value::Scalar::Token(text)) => {
                        text.to_owned()
                    },
                    Some(scalar) => scalar.to_string(),
                    None => String::new(),
                };
                match setter {
                    Setter::CommentAuthor => author = Some(value),
                    Setter::CommentInitials => initials = Some(value),
                    Setter::CommentDate => date = Some(value),
                    Setter::CommentText => text = value,
                    _ => {
                        return Err(Error::new(
                            Code::UnsupportedProperty,
                            "the property does not seed a DOCX comment",
                        ));
                    },
                }
                Ok(())
            })
            .map_err(|error| {
                error
                    .with_property(key.as_str().to_owned())
                    .with_format(FileFormat::Docx)
            })?;
    }
    let Some(author) = author else {
        return Err(
            Error::new(Code::MissingProperty, "a new DOCX comment needs its author")
                .with_property("author")
                .with_format(FileFormat::Docx),
        );
    };
    let mut comment = litchi_docx::document::NewComment::new(author, text);
    if let Some(initials) = initials {
        comment = comment.with_initials(initials);
    }
    if let Some(date) = date {
        comment = comment.with_date(date);
    }
    Ok(comment)
}

/// A new row's cell count from its add properties.
fn row_cells(props: &Props<'_>) -> Result<Option<usize>> {
    let mut cells = None;
    for (key, value) in props.iter() {
        let (_, entry) = ROW_REGISTRY
            .resolve(key.as_str())
            .map_err(|error| error.with_property(key.as_str().to_owned()))?;
        entry
            .prepare(Operation::Add, value, |_, value| {
                if let Some(crate::value::Scalar::Int(count)) = value {
                    cells = usize::try_from(count).ok();
                }
                Ok(())
            })
            .map_err(|error| {
                error
                    .with_property(key.as_str().to_owned())
                    .with_format(FileFormat::Docx)
            })?;
    }
    Ok(cells)
}

fn count(text: &str) -> Result<usize> {
    text.trim()
        .parse::<usize>()
        .ok()
        .filter(|count| *count > 0)
        .ok_or_else(|| {
            Error::new(
                Code::InvalidValue,
                "a table dimension is a positive integer",
            )
        })
}

fn missing(property: &'static str) -> Error {
    Error::new(
        Code::MissingProperty,
        "a new DOCX table needs its dimension",
    )
    .with_property(property)
    .with_format(FileFormat::Docx)
}
