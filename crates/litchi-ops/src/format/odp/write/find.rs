//! ODP find/replace and template merge (record 0295) over the shared text
//! surface (`crate::find::surface`).
//!
//! The leaf scene lists every story: a drawing object's text (group members
//! included), each table cell's text and each notes page's objects. A
//! paragraph's runs are one group; a hyperlink is a group of its own and a
//! field (page number, date, bookmark) a boundary that is not searched. A run whose markup holds
//! only text, spaces, tabs and breaks is rewritten whole by one leaf
//! `SetRunText` change, which keeps the run's own style; a run holding other
//! markup (a nested span or a field inside a span) is fixed: a find that
//! touches it is refused and merge leaves it. Format properties are
//! refused. Every change is one scene change of one checked splice and one
//! commit.
//!
//! Scope: the presentation searches every slide and notes page (OfficeCLI's
//! root fan-out); a slide its objects; an object, table, row, cell,
//! paragraph or run what lies inside it; a notes node its notes page.

use std::ops::Range;

use litchi_core::FileFormat;
use litchi_odp::scene::{Change, Object, ObjectRef, RunKind, Scene, Scope, Story, StoryRef};

use super::super::node::{Item, Node};
use super::super::{Nav, Odp};
use super::{WriteCommit, commit_changes};
use crate::error::{Code, Error, Result};
use crate::find::Spec;
use crate::find::surface::{self, Run};
use crate::merge::{Options, Planner, Record, Summary};
use crate::target::TargetArg;
use crate::transaction::SemanticLimits;

/// Where a surface run lives.
struct Key {
    story: StoryRef,
    paragraph: usize,
    run: usize,
}

/// Find and replace text below `target`.
pub(crate) fn find_replace(
    file: &Odp,
    target: &TargetArg<'_>,
    spec: &Spec<'_>,
    limits: &SemanticLimits,
) -> Result<(WriteCommit, usize)> {
    file.check_source()?;
    let result = (|| {
        let replacement = surface::text_only(
            spec,
            FileFormat::Odp,
            "the leaf rewrites whole runs and cannot split one at a match",
        )?;
        file.admit_edit()?;
        let scene = file.scene()?;
        let mut every = false;
        let mut ranges: Vec<Range<usize>> = Vec::new();
        surface::each_target(Nav::new(file), target, |node: Node<'_>| {
            match scope_range(scene, node.item())? {
                None => every = true,
                Some(range) => ranges.push(range),
            }
            Ok(())
        })?;
        let content = file.content_xml()?;
        let (runs, keys) = surface_of(scene, content, (!every).then_some(ranges.as_slice()));
        let (changes, matches) = surface::find(&runs, spec, replacement, FileFormat::Odp, limits)?;
        let changes: Vec<Change> = changes
            .into_iter()
            .map(|change| {
                let key = &keys[change.run];
                Change::SetRunText {
                    story: key.story.clone(),
                    paragraph: key.paragraph,
                    run: key.run,
                    text: change.text,
                }
            })
            .collect();
        Ok((commit_changes(file, &changes, limits)?, matches))
    })()
    .map_err(|error: Error| error.with_format(FileFormat::Odp));
    file.check_source()?;
    result
}

/// Fill `record` into every story. `None` means nothing matched.
pub(crate) fn merge(
    file: &Odp,
    record: &Record,
    options: &Options,
    limits: &SemanticLimits,
) -> Result<(Option<WriteCommit>, Summary)> {
    file.check_source()?;
    let result = (|| {
        file.admit_edit()?;
        let scene = file.scene()?;
        let content = file.content_xml()?;
        let (runs, keys) = surface_of(scene, content, None);
        let mut planner = Planner::new(record, options, limits);
        let changes = surface::merge(&runs, &mut planner, true, FileFormat::Odp, limits)?;
        let summary = planner.finish();
        if changes.is_empty() {
            return Ok((None, summary));
        }
        let changes: Vec<Change> = changes
            .into_iter()
            .map(|change| {
                let key = &keys[change.run];
                Change::SetRunText {
                    story: key.story.clone(),
                    paragraph: key.paragraph,
                    run: key.run,
                    text: change.text,
                }
            })
            .collect();
        Ok((Some(commit_changes(file, &changes, limits)?), summary))
    })()
    .map_err(|error: Error| error.with_format(FileFormat::Odp));
    file.check_source()?;
    result
}

/// The `content.xml` bytes a target covers; `None` for the presentation.
fn scope_range(scene: &Scene, item: &Item) -> Result<Option<Range<usize>>> {
    let missing = || Error::new(Code::PathNotFound, "the ODP find target vanished");
    let story = |story: &StoryRef| -> Result<&Story> { scene.story(story).ok_or_else(missing) };
    Ok(Some(match item {
        Item::Presentation => return Ok(None),
        Item::Slide(page) => scene.page(*page).ok_or_else(missing)?.range(),
        Item::Object(object, _) => scene.object(object).ok_or_else(missing)?.range(),
        Item::Notes(page) => scene
            .page(*page)
            .and_then(|page| page.notes())
            .ok_or_else(missing)?
            .range(),
        Item::Paragraph(at, paragraph) => story(at)?
            .paragraphs()
            .get(*paragraph)
            .ok_or_else(missing)?
            .range(),
        Item::Run(at, paragraph, run) => story(at)?
            .paragraphs()
            .get(*paragraph)
            .and_then(|paragraph| paragraph.runs().get(*run))
            .ok_or_else(missing)?
            .range(),
        Item::Row(object, row) => scene
            .object(object)
            .and_then(Object::table)
            .and_then(|table| table.rows().get(*row))
            .ok_or_else(missing)?
            .range(),
        Item::Cell(object, row, cell) => scene
            .object(object)
            .and_then(Object::table)
            .and_then(|table| table.rows().get(*row))
            .and_then(|row| row.cells().get(*cell))
            .ok_or_else(missing)?
            .range(),
        Item::RunLink(at, paragraph, run) => story(at)?
            .paragraphs()
            .get(*paragraph)
            .and_then(|paragraph| paragraph.runs().get(*run))
            .ok_or_else(missing)?
            .range(),
        Item::ShapeLink(object) => scene.object(object).ok_or_else(missing)?.range(),
        Item::Comment(_)
        | Item::Master(_)
        | Item::Layout(_)
        | Item::LineBreak(..)
        | Item::Column(..)
        | Item::Series(..)
        | Item::Axis(..) => {
            return Err(Error::new(
                Code::UnsupportedType,
                "an ODP find searches slides, their objects and notes",
            )
            .with_suggestion("target the presentation, a slide, an object or a notes page"));
        },
    }))
}

/// Every story's runs in document order (slides, then each slide's notes
/// page), scoped by `ranges` (`None`: everything).
fn surface_of<'s>(
    scene: &'s Scene,
    content: &str,
    ranges: Option<&[Range<usize>]>,
) -> (Vec<Run<'s>>, Vec<Key>) {
    let mut out = Surface {
        content,
        ranges,
        runs: Vec::new(),
        keys: Vec::new(),
        group: 0,
    };
    for (index, page) in scene.pages().iter().enumerate() {
        for (position, object) in page.objects().iter().enumerate() {
            out.object(object, ObjectRef::new(index, position));
        }
        if let Some(notes) = page.notes() {
            for (position, object) in notes.objects().iter().enumerate() {
                out.object(
                    object,
                    ObjectRef {
                        page: index,
                        scope: Scope::Notes,
                        path: vec![position],
                    },
                );
            }
        }
    }
    (out.runs, out.keys)
}

struct Surface<'s, 'c, 'r> {
    content: &'c str,
    ranges: Option<&'r [Range<usize>]>,
    runs: Vec<Run<'s>>,
    keys: Vec<Key>,
    group: usize,
}

impl<'s> Surface<'s, '_, '_> {
    fn object(&mut self, object: &'s Object, at: ObjectRef) {
        if let Some(story) = object.story() {
            self.story(story, &StoryRef::object(at.clone()));
        }
        if let Some(table) = object.table() {
            for (row_index, row) in table.rows().iter().enumerate() {
                for (cell_index, cell) in row.cells().iter().enumerate() {
                    if cell.is_covered() {
                        continue;
                    }
                    self.story(
                        cell.story(),
                        &StoryRef {
                            object: at.clone(),
                            cell: Some((row_index, cell_index)),
                        },
                    );
                }
            }
        }
        for (position, child) in object.children().iter().enumerate() {
            let mut path = at.path.clone();
            path.push(position);
            self.object(
                child,
                ObjectRef {
                    page: at.page,
                    scope: at.scope,
                    path,
                },
            );
        }
    }

    fn story(&mut self, story: &'s Story, at: &StoryRef) {
        for (index, paragraph) in story.paragraphs().iter().enumerate() {
            let range = paragraph.range();
            if let Some(ranges) = self.ranges
                && !ranges
                    .iter()
                    .any(|scope| scope.start < range.end && range.start < scope.end)
            {
                continue;
            }
            self.group += 1;
            for (position, run) in paragraph.runs().iter().enumerate() {
                let own = run.range();
                let (editable, alone) = match run.kind() {
                    RunKind::Text => (true, false),
                    RunKind::Span => (
                        inner(self.content, &own).is_some_and(surface::odf_text_only),
                        false,
                    ),
                    RunKind::Link => (
                        inner(self.content, &own).is_some_and(surface::odf_text_only),
                        true,
                    ),
                    // Fields (page numbers, dates) are computed, as PPTX
                    // fields: a boundary, never searched.
                    _ => {
                        self.group += 1;
                        continue;
                    },
                };
                if alone {
                    self.group += 1;
                }
                let scoped = self.ranges.is_none_or(|ranges| {
                    ranges
                        .iter()
                        .any(|scope| scope.start <= own.start && own.end <= scope.end)
                });
                self.runs.push(
                    Run::new(run.text(), self.group)
                        .scoped(scoped)
                        .editable(editable),
                );
                self.keys.push(Key {
                    story: at.clone(),
                    paragraph: index,
                    run: position,
                });
                if alone {
                    self.group += 1;
                }
            }
        }
    }
}

/// The markup between an element's start and end tags.
fn inner<'c>(content: &'c str, range: &Range<usize>) -> Option<&'c str> {
    let element = content.get(range.clone())?;
    if element.ends_with("/>") && element.matches('<').count() == 1 {
        return Some("");
    }
    let start = element.find('>')? + 1;
    let end = element.rfind('<')?;
    element.get(start..end.max(start))
}
