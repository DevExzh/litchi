//! Writes of the DOCX marks (record 0319): paragraph tab stops are set,
//! added and removed through the paragraph's typed properties; breaks are
//! added as a run holding one `w:br` and removed with that run.
//!
//! Both owners take these edits: they change only the main document.

use litchi_core::selector::Path;
use litchi_core::{FileFormat, Position};
use litchi_docx::RunBreakType;
use litchi_docx::document::ElementStep;
use litchi_docx::properties::{TabKind, TabStop, Twips};

use super::super::element::{element_path, path_with, refusal, run_locations};
use super::super::node::{Node, Value};
use super::super::spec::Kind;
use super::super::write::{Commit, WorkEdit};
use super::super::{Docx, DocxEditor};
use super::{BREAK_REGISTRY, MarkSetter, MarkValue, TAB_REGISTRY, tab_kind, tab_leader, with_stop};
use crate::error::{Code, Error, Result};
use crate::new_node::NewNode;
use crate::props::Props;
use crate::registry::Operation;
use crate::value::Scalar;

fn error(code: Code, message: &'static str) -> Error {
    Error::new(code, message).with_format(FileFormat::Docx)
}

/// A new mark element, by the type name `add` was given.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(in super::super) enum NewMark {
    Tab,
    Break,
}

impl NewMark {
    /// The new mark element `name` names, if it is one.
    pub(in super::super) fn of(name: &str) -> Option<Self> {
        if name.eq_ignore_ascii_case("tab") {
            Some(Self::Tab)
        } else if ["break", "br", "pagebreak", "columnbreak"]
            .iter()
            .any(|known| name.eq_ignore_ascii_case(known))
        {
            Some(Self::Break)
        } else {
            None
        }
    }
}

/// The paragraph node a mark belongs to.
fn paragraph_node<'f>(
    docx: &'f Docx,
    paragraph: &Value,
    address: &super::super::node::Address,
) -> Node<'f> {
    Node::new(docx, Kind::Paragraph, paragraph.clone(), address.clone())
}

/// Apply `props` to `stop` through the tab registry.
fn apply_tab(stop: &mut TabStop, props: &Props<'_>, operation: Operation) -> Result<()> {
    for (key, value) in props.iter() {
        let (_, entry) = TAB_REGISTRY.resolve(key.as_str()).map_err(|error| {
            error
                .with_property(key.as_str().to_owned())
                .with_format(FileFormat::Docx)
        })?;
        entry
            .prepare(operation, value, |setter, value| {
                match (setter, value) {
                    (MarkSetter::TabPosition, Some(Scalar::Int(position))) => {
                        stop.position = Twips(position)
                    },
                    (MarkSetter::TabKind, Some(Scalar::Token(token) | Scalar::Text(token))) => {
                        stop.kind = tab_kind(token).ok_or_else(|| {
                            error(Code::InvalidValue, "unknown tab stop alignment")
                        })?;
                    },
                    (MarkSetter::TabLeader, Some(Scalar::Token(token) | Scalar::Text(token))) => {
                        stop.leader = Some(
                            tab_leader(token)
                                .ok_or_else(|| error(Code::InvalidValue, "unknown tab leader"))?,
                        );
                    },
                    _ => {},
                }
                Ok(())
            })
            .map_err(|error| {
                error
                    .with_property(key.as_str().to_owned())
                    .with_format(FileFormat::Docx)
            })?;
    }
    Ok(())
}

/// Format the paragraph at `path` with `change`, on either owner.
fn format_paragraph(
    edit: &mut WorkEdit,
    path: &litchi_docx::document::ElementPath,
    change: impl FnOnce(&mut litchi_docx::properties::Paragraph) -> Result<()>,
) -> Result<()> {
    let mut failure = None;
    let mut change = Some(change);
    let mut apply = |paragraph: &mut litchi_docx::properties::Paragraph| {
        if let Some(change) = change.take()
            && let Err(error) = change(paragraph)
        {
            failure = Some(error);
        }
    };
    match edit {
        WorkEdit::Eager(edit) => {
            edit.format_paragraph_at(path, &mut apply)
                .map_err(Error::from)?;
        },
        WorkEdit::Source(edit) => {
            edit.format_paragraph_at(path, &mut apply)
                .map_err(Error::from)?;
        },
    }
    failure.map_or(Ok(()), Err)
}

impl Docx {
    /// `set` on marks: `Err(editor)` hands the editor back when the targets
    /// are not marks.
    pub(in super::super) fn try_set_marks(
        &self,
        editor: DocxEditor,
        nodes: &[Node<'_>],
        props: &Props<'_>,
    ) -> Result<std::result::Result<Commit, DocxEditor>> {
        if nodes.is_empty()
            || nodes
                .iter()
                .any(|node| !super::is_mark_kind(node.element_kind()))
        {
            return Ok(Err(editor));
        }
        if nodes.iter().any(|node| node.element_kind() != Kind::Tab) {
            return Err(error(
                Code::UnsupportedCapability,
                "only paragraph tab stops are set; breaks, bookmarks, fields, tables of contents, watermarks and text boxes are read-only here",
            ));
        }
        // Admit on a scratch stop first, so a refused key stages nothing.
        let mut probe = TabStop::new(TabKind::Left, Twips(0));
        apply_tab(&mut probe, props, Operation::Set)?;
        self.owner().check_source()?;
        let mut edit = WorkEdit::from(editor)?;
        for node in nodes {
            let Value::Mark(mark) = &node.value else {
                return Err(error(Code::InternalError, "a mark node holds no mark"));
            };
            let MarkValue::Tab {
                anchor,
                index,
                stop,
            } = mark.as_ref()
            else {
                return Err(error(Code::InternalError, "a tab node holds no tab stop"));
            };
            let path = element_path(&paragraph_node(self, &anchor.paragraph, &anchor.address))?;
            let mut next = *stop;
            apply_tab(&mut next, props, Operation::Set)?;
            let index = *index;
            format_paragraph(&mut edit, &path, |properties| {
                with_stop(properties, index, Some(next))
            })?;
        }
        Ok(Ok(edit.commit(self)?))
    }

    /// `add` of a mark under a paragraph.
    pub(in super::super) fn add_mark(
        &self,
        editor: DocxEditor,
        parent: &Node<'_>,
        kind: NewMark,
        node: &NewNode<'_>,
    ) -> Result<(Commit, Path<'static>)> {
        if parent.element_kind() != Kind::Paragraph {
            return Err(error(
                Code::UnsupportedCapability,
                "tab stops and breaks are added under a paragraph",
            ));
        }
        let path = element_path(parent)?;
        match kind {
            NewMark::Tab => {
                let mut stop = TabStop::new(TabKind::Left, Twips(0));
                apply_tab(&mut stop, node.props(), Operation::Add)?;
                let count = super::tab_stops(&parent.value)?.len();
                self.owner().check_source()?;
                let mut edit = WorkEdit::from(editor)?;
                format_paragraph(&mut edit, &path, |properties| {
                    properties.tabs.get_or_insert_with(Vec::new).push(stop);
                    Ok(())
                })?;
                let commit = edit.commit(self)?;
                Ok((
                    commit,
                    path_with(parent, Some((Kind::Tab, Position::new(count)))),
                ))
            },
            NewMark::Break => {
                let mut kind = RunBreakType::Page;
                for (key, value) in node.props().iter() {
                    let (_, entry) = BREAK_REGISTRY.resolve(key.as_str()).map_err(|error| {
                        error
                            .with_property(key.as_str().to_owned())
                            .with_format(FileFormat::Docx)
                    })?;
                    entry
                        .prepare(Operation::Add, value, |_, value| {
                            kind = match value {
                                Some(Scalar::Token("column")) => RunBreakType::Column,
                                Some(Scalar::Token("textWrapping" | "line")) => {
                                    RunBreakType::TextWrapping
                                },
                                _ => RunBreakType::Page,
                            };
                            Ok(())
                        })
                        .map_err(|error| {
                            error
                                .with_property(key.as_str().to_owned())
                                .with_format(FileFormat::Docx)
                        })?;
                }
                let Value::Paragraph { paragraph, .. } = &parent.value else {
                    return Err(refusal("the break parent holds no paragraph"));
                };
                let locations = run_locations(self, paragraph)?;
                let at = super::super::nested::sibling_position(
                    self,
                    parent,
                    Kind::Run,
                    locations.len(),
                    node.placement(),
                )?;
                let direct = match at {
                    None => None,
                    Some(index) if index >= locations.len() => None,
                    Some(index) => match locations.get(index).map(Vec::as_slice) {
                        Some([ElementStep::Run(run)]) => Some(*run),
                        _ => {
                            return Err(refusal(
                                "a break goes beside a direct run of the paragraph",
                            ));
                        },
                    },
                };
                // The new break's ordinal: the breaks of the runs before it.
                let before = super::breaks(&parent.value)?
                    .iter()
                    .filter(|(_, run)| at.is_none_or(|at| *run < at))
                    .count();
                self.owner().check_source()?;
                let mut edit = WorkEdit::from(editor)?;
                match &mut edit {
                    WorkEdit::Eager(edit) => edit.insert_break_at(&path, direct, kind),
                    WorkEdit::Source(edit) => edit.insert_break_at(&path, direct, kind),
                }
                .map_err(Error::from)?;
                let commit = edit.commit(self)?;
                Ok((
                    commit,
                    path_with(parent, Some((Kind::Break, Position::new(before)))),
                ))
            },
        }
    }

    /// `remove` of marks: tab stops, and breaks whose run holds nothing
    /// else.
    pub(in super::super) fn remove_marks(
        &self,
        editor: DocxEditor,
        nodes: &[Node<'_>],
    ) -> Result<Commit> {
        enum Removal {
            Tab(litchi_docx::document::ElementPath, usize),
            Run(litchi_docx::document::ElementPath),
        }
        let mut removals = Vec::with_capacity(nodes.len());
        for node in nodes {
            let Value::Mark(mark) = &node.value else {
                return Err(error(Code::InternalError, "a mark node holds no mark"));
            };
            match mark.as_ref() {
                MarkValue::Tab { anchor, index, .. } => removals.push(Removal::Tab(
                    element_path(&paragraph_node(self, &anchor.paragraph, &anchor.address))?,
                    *index,
                )),
                MarkValue::Break { anchor, run, .. } => {
                    let Value::Paragraph {
                        paragraph: value, ..
                    } = &anchor.paragraph
                    else {
                        return Err(error(
                            Code::InternalError,
                            "a break node holds no paragraph",
                        ));
                    };
                    let runs = value.addressable_runs().map_err(Error::from)?;
                    let holds_only_breaks = runs.get(*run).is_some_and(|candidate| {
                        candidate.contents().is_ok_and(|contents| {
                            contents
                                .iter()
                                .all(|content| matches!(content, litchi_docx::RunContent::Break(_)))
                        })
                    });
                    if !holds_only_breaks {
                        return Err(error(
                            Code::UnsupportedCapability,
                            "a break is removed with its run, which here holds other content",
                        ));
                    }
                    let parent =
                        element_path(&paragraph_node(self, &anchor.paragraph, &anchor.address))?;
                    let locations = run_locations(self, value)?;
                    let location = locations
                        .get(*run)
                        .ok_or_else(|| error(Code::PathNotFound, "the break's run is gone"))?;
                    let mut path = parent;
                    for step in location {
                        path = path.child(*step);
                    }
                    removals.push(Removal::Run(path));
                },
                _ => {
                    return Err(error(
                        Code::UnsupportedCapability,
                        "only tab stops and breaks are removed; bookmarks, fields, tables of contents, watermarks and text boxes are read-only here",
                    ));
                },
            }
        }
        // Later positions first, so no removal shifts another.
        removals.sort_by(|left, right| {
            let key = |removal: &Removal| match removal {
                Removal::Tab(path, index) => (super::super::nested::order_key(path), *index),
                Removal::Run(path) => (super::super::nested::order_key(path), 0),
            };
            key(right).cmp(&key(left))
        });
        self.owner().check_source()?;
        let mut edit = WorkEdit::from(editor)?;
        for removal in removals {
            match removal {
                Removal::Tab(path, index) => {
                    format_paragraph(&mut edit, &path, |properties| {
                        with_stop(properties, index, None)
                    })?;
                },
                Removal::Run(path) => edit.remove_element_at(&path)?,
            }
        }
        edit.commit(self)
    }
}
