//! Paragraphs and runs of headers, footers, footnotes and endnotes
//! (record 0325): `/header[N]/p[M]`, `/footer[N]/p[M]`, `/footnote[N]/p[M]`
//! and `/endnote[N]/p[M]`, with their runs.
//!
//! A story's paragraphs are read through the story's block index, as the
//! body's are, so every paragraph and run read applies. Writes reuse the
//! main document's element-path verbs: the leaf's `Edit::edit_story` runs
//! them on the story's content as a document body and splices the result
//! back into the part. Unscoped queries stay in the main story, and the
//! story nodes' views do not list their paragraphs (record 0319).

use litchi_core::selector::Path;
use litchi_core::{FileFormat, Position};
use litchi_docx::document::{ElementAnchor, ElementPath, ElementStep, StoryTarget};

use super::element::{StoryKey, path_with, run_locations, story_path};
use super::node::{Node, Value};
use super::spec::Kind;
use super::write::{Commit, WorkEdit};
use super::{Docx, DocxEditor};
use crate::error::{Code, Error, Result};
use crate::new_node::NewNode;
use crate::placement::Placement;

fn error(code: Code, message: &'static str) -> Error {
    Error::new(code, message).with_format(FileFormat::Docx)
}

/// The leaf target of a story key.
pub(super) fn target(key: &StoryKey) -> StoryTarget<'_> {
    match key {
        StoryKey::Part(name) => StoryTarget::Part(name),
        StoryKey::Note(kind, id) => StoryTarget::Note(*kind, *id),
    }
}

/// Whether `kind` is a story container a paragraph is added under.
const fn is_story(kind: Kind) -> bool {
    matches!(
        kind,
        Kind::Header | Kind::Footer | Kind::Footnote | Kind::Endnote
    )
}

/// The story key of a story node itself (`/header[N]`, `/footnote[N]`).
fn story_key(node: &Node<'_>) -> Result<StoryKey> {
    let step = node
        .address
        .steps()
        .last()
        .ok_or_else(|| error(Code::InternalError, "a story node has no address"))?;
    Ok(super::parts::story_of(node.docx, step.kind, step.ordinal as usize)?.0)
}

impl Docx {
    fn story_edit(&self, editor: DocxEditor) -> Result<WorkEdit> {
        self.owner().check_source()?;
        WorkEdit::from(self.eager_editor(editor)?)
    }

    /// `add` of a paragraph under a story, or of a run under a story
    /// paragraph; `Err(editor)` hands the editor back for any other parent.
    pub(super) fn try_add_story(
        &self,
        editor: DocxEditor,
        parent: &Node<'_>,
        node: &NewNode<'_>,
    ) -> Result<std::result::Result<(Commit, Path<'static>), DocxEditor>> {
        let kind = node.kind();
        let paragraph = kind.eq_ignore_ascii_case("p") || kind.eq_ignore_ascii_case("paragraph");
        let run = kind.eq_ignore_ascii_case("r") || kind.eq_ignore_ascii_case("run");
        if paragraph && is_story(parent.element_kind()) {
            return self.add_story_paragraph(editor, parent, node).map(Ok);
        }
        // Fields and bookmarks in a story paragraph take the mark route.
        if run
            && parent.element_kind() == Kind::Paragraph
            && let Some((key, path)) = story_path(parent)?
        {
            return self
                .add_story_run(editor, parent, &key, &path, node)
                .map(Ok);
        }
        if is_story(parent.element_kind()) && run {
            return Err(error(
                Code::UnsupportedCapability,
                "runs are added under a paragraph of the story",
            ));
        }
        Ok(Err(editor))
    }

    fn add_story_paragraph(
        &self,
        editor: DocxEditor,
        parent: &Node<'_>,
        node: &NewNode<'_>,
    ) -> Result<(Commit, Path<'static>)> {
        let key = story_key(parent)?;
        let paragraph = super::structure::new_paragraph(node.props())?;
        let anchor = self.story_anchor(parent, node.placement())?;
        let mut edit = self.story_edit(editor)?;
        let WorkEdit::Eager(eager) = &mut edit else {
            return Err(error(
                Code::InternalError,
                "a story edit needs the package editor",
            ));
        };
        let position = eager.edit_story(target(&key), |story| -> Result<Position> {
            Ok(story.insert_paragraph_at(&ElementPath::new(Vec::new()), anchor, &paragraph)?)
        })?;
        let commit = edit.commit(self)?;
        Ok((commit, path_with(parent, Some((Kind::Paragraph, position)))))
    }

    /// Where a new story paragraph goes among the story's blocks.
    fn story_anchor(
        &self,
        parent: &Node<'_>,
        placement: Option<&Placement<'_>>,
    ) -> Result<ElementAnchor> {
        let blocks = super::nav::container_of(parent)?
            .ok_or_else(|| error(Code::InternalError, "a story holds no blocks"))?;
        let paragraph_at = |ordinal: usize| -> Result<Option<ElementStep>> {
            blocks
                .nth(litchi_docx::BlockKind::Paragraph, Position::new(ordinal))
                .map(|block| {
                    block
                        .direct_position()
                        .map(ElementStep::Paragraph)
                        .ok_or_else(|| error(Code::InternalError, "a story block has no position"))
                })
                .transpose()
        };
        Ok(match placement {
            None | Some(Placement::End) => ElementAnchor::End,
            Some(Placement::Start) => ElementAnchor::Start,
            Some(Placement::Position(position)) => match paragraph_at(position.get())? {
                Some(step) => ElementAnchor::Before(step),
                None => ElementAnchor::End,
            },
            Some(Placement::Before(path) | Placement::After(path)) => {
                let sibling = crate::path::resolve::get(&super::Nav::new(self), path)?;
                let same_story = sibling.element_kind() == Kind::Paragraph
                    && sibling.address.steps().len() == parent.address.steps().len() + 1
                    && sibling.address.steps().starts_with(parent.address.steps());
                if !same_story {
                    return Err(error(
                        Code::InvalidArgument,
                        "a story paragraph goes beside a paragraph of the same story",
                    ));
                }
                let ordinal = sibling
                    .address
                    .steps()
                    .last()
                    .map_or(0, |step| step.ordinal as usize);
                let step = paragraph_at(ordinal)?
                    .ok_or_else(|| error(Code::PathNotFound, "the sibling paragraph is gone"))?;
                if matches!(placement, Some(Placement::Before(_))) {
                    ElementAnchor::Before(step)
                } else {
                    ElementAnchor::After(step)
                }
            },
            Some(_) => {
                return Err(error(
                    Code::UnsupportedCapability,
                    "a story paragraph is placed at the start, the end, a position or beside a sibling",
                ));
            },
        })
    }

    fn add_story_run(
        &self,
        editor: DocxEditor,
        parent: &Node<'_>,
        key: &StoryKey,
        path: &ElementPath,
        node: &NewNode<'_>,
    ) -> Result<(Commit, Path<'static>)> {
        let Value::Paragraph { paragraph, .. } = &parent.value else {
            return Err(error(
                Code::InternalError,
                "a paragraph node holds no paragraph",
            ));
        };
        let run = super::structure::new_run(node.props())?;
        let locations = run_locations(self, paragraph)?;
        let index = super::nested::sibling_position(
            self,
            parent,
            Kind::Run,
            locations.len(),
            node.placement(),
        )?;
        let (at, ordinal) = match index {
            None => (None, locations.len()),
            Some(index) if index >= locations.len() => (None, locations.len()),
            Some(index) => match locations.get(index).map(Vec::as_slice) {
                Some([ElementStep::Run(run)]) => (Some(*run), index),
                _ => {
                    return Err(error(
                        Code::UnsupportedCapability,
                        "a run goes beside a direct run of the paragraph",
                    ));
                },
            },
        };
        let mut edit = self.story_edit(editor)?;
        let WorkEdit::Eager(eager) = &mut edit else {
            return Err(error(
                Code::InternalError,
                "a story edit needs the package editor",
            ));
        };
        eager.edit_story(target(key), |story| -> Result<Position> {
            Ok(story.insert_run_at(path, at, &run)?)
        })?;
        let commit = edit.commit(self)?;
        Ok((
            commit,
            path_with(parent, Some((Kind::Run, Position::new(ordinal)))),
        ))
    }

    /// `remove` of story paragraphs and runs; `Err(editor)` when no target
    /// is in a story. A story keeps at least one paragraph.
    pub(super) fn try_remove_story(
        &self,
        editor: DocxEditor,
        nodes: &[Node<'_>],
    ) -> Result<std::result::Result<Commit, DocxEditor>> {
        let mut found = Vec::with_capacity(nodes.len());
        for node in nodes {
            if let Some(story) = story_path(node)? {
                found.push((story, node));
            }
        }
        if found.is_empty() {
            return Ok(Err(editor));
        }
        if found.len() != nodes.len() {
            return Err(error(
                Code::UnsupportedCapability,
                "one remove takes story elements or main-story elements, not both",
            ));
        }
        if found
            .iter()
            .any(|(_, node)| !matches!(node.element_kind(), Kind::Paragraph | Kind::Run))
        {
            return Err(error(
                Code::UnsupportedCapability,
                "story paragraphs and runs are removed; other story elements are not",
            ));
        }
        // A story's last paragraph stays: a header, footer or note needs one.
        for (key, _) in found.iter().map(|((key, _), node)| (key, node)) {
            let paragraphs = found
                .iter()
                .filter(|((other, _), node)| other == key && node.element_kind() == Kind::Paragraph)
                .count();
            let parent = found
                .iter()
                .find(|((other, _), _)| other == key)
                .map(|(_, node)| *node)
                .ok_or_else(|| error(Code::InternalError, "a story target vanished"))?;
            let total = super::parts::story_blocks(self, story_value(parent)?)?
                .map_or(0, |blocks| blocks.count(litchi_docx::BlockKind::Paragraph));
            if paragraphs > 0 && paragraphs >= total {
                return Err(error(
                    Code::InvalidArgument,
                    "a header, footer or note keeps at least one paragraph",
                ));
            }
        }
        let mut edit = self.story_edit(editor)?;
        let WorkEdit::Eager(eager) = &mut edit else {
            return Err(error(
                Code::InternalError,
                "a story edit needs the package editor",
            ));
        };
        let mut index = 0;
        while index < found.len() {
            let key = found[index].0.0.clone();
            let end = index
                + found[index..]
                    .iter()
                    .take_while(|((other, _), _)| *other == key)
                    .count();
            let mut paths: Vec<&ElementPath> = found[index..end]
                .iter()
                .map(|((_, path), _)| path)
                .collect();
            // Later positions first, so no removal shifts another.
            paths.sort_by(|left, right| {
                super::nested::order_key(right).cmp(&super::nested::order_key(left))
            });
            eager.edit_story(target(&key), |story| -> Result<()> {
                for path in paths {
                    story.remove_element_at(path)?;
                }
                Ok(())
            })?;
            index = end;
        }
        Ok(Ok(edit.commit(self)?))
    }
}

/// The part value of the story a story element belongs to.
fn story_value(node: &Node<'_>) -> Result<super::parts::PartValue> {
    let steps = node.address.steps();
    let skip = usize::from(
        steps
            .first()
            .is_some_and(|step| matches!(step.kind, Kind::Footnotes | Kind::Endnotes)),
    );
    let step = steps
        .get(skip)
        .ok_or_else(|| error(Code::InternalError, "a story element has no story step"))?;
    Ok(super::parts::story_of(node.docx, step.kind, step.ordinal as usize)?.1)
}
