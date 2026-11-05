//! PPT find/replace and template merge (record 0295) over the shared text
//! surface (`crate::find::surface`).
//!
//! The binary text atom of a shape holds its paragraphs separated by
//! paragraph marks, and the native editor replaces a shape's text whole
//! (`ShapeTextReplacement`), keeping the formatting of a single-run body and
//! refusing a length change where per-character formatting ranges would
//! shift. Each paragraph is therefore one run of the surface (one group);
//! every changed shape is one native text replacement, every changed notes
//! page one notes replacement, in one native edit and one commit.
//!
//! Scope: the presentation searches every slide's text shapes and notes
//! (OfficeCLI's root fan-out), a slide its shapes, a shape its paragraphs,
//! a paragraph itself and a notes node its notes. Shapes inside groups are
//! searched and refused when they would change (the native editor
//! addresses direct slide shapes only). Format properties are refused.

use std::collections::BTreeMap;
use std::sync::Arc;

use litchi_core::{FileFormat, Position};

use super::super::node::{Item, Node, has_text_body, paragraphs};
use super::super::spec::Kind;
use super::super::{Nav, Ppt};
use super::{NotesPlan, Plan, ShapePlan, WriteCommit, finish_commit, stage_plans};
use crate::error::{Code, Error, Result};
use crate::find::Spec;
use crate::find::surface::{self, Change, Run};
use crate::merge::{Options, Planner, Record, Summary};
use crate::path::navigate::{KindSet, Navigate};
use crate::target::TargetArg;
use crate::transaction::SemanticLimits;

/// What a find covers.
#[derive(Clone)]
enum Scope {
    Every,
    Slide(usize),
    Shape(usize, Arc<[usize]>),
    Paragraph(usize, Arc<[usize]>, usize),
    Notes(usize),
}

/// Where a surface run lives: one paragraph of a shape or of a notes page.
#[derive(Clone)]
struct Key {
    holder: Holder,
    paragraph: usize,
}

#[derive(Clone, PartialEq, Eq, PartialOrd, Ord)]
enum Holder {
    /// A direct slide shape (`None` target: nested in a group).
    Shape {
        slide: usize,
        path: Arc<[usize]>,
        target: Option<(usize, usize)>,
    },
    Notes(usize),
}

/// One holder's paragraphs, as read.
struct Text {
    holder: Holder,
    paragraphs: Vec<String>,
}

/// Find and replace text below `target`.
pub(crate) fn find_replace(
    file: &Ppt,
    target: &TargetArg<'_>,
    spec: &Spec<'_>,
    limits: &SemanticLimits,
) -> Result<(WriteCommit, usize)> {
    file.check_source()?;
    let result = (|| {
        let replacement = surface::text_only(
            spec,
            FileFormat::Ppt,
            "the native PPT editor replaces a shape's text whole",
        )?;
        file.admit_edit_for_write()?;
        let mut scopes = Vec::new();
        surface::each_target(Nav::new(file), target, |node: Node<'_>| {
            scopes.push(match node.item() {
                Item::Presentation => Scope::Every,
                Item::Slide(slide) => Scope::Slide(slide.position),
                Item::Shape(shape) => Scope::Shape(shape.slide, shape.path.clone()),
                Item::Paragraph(paragraph) => Scope::Paragraph(
                    paragraph.shape.slide,
                    paragraph.shape.path.clone(),
                    paragraph.index,
                ),
                Item::Notes(notes) => Scope::Notes(notes.slide),
                Item::Element(_) => {
                    return Err(Error::new(
                        Code::UnsupportedCapability,
                        "PPT find and replace scopes are slides, shapes, paragraphs and notes",
                    )
                    .with_format(FileFormat::Ppt));
                },
            });
            Ok(())
        })?;
        let texts = read(file)?;
        let (runs, keys) = runs(&texts, &scopes);
        let (changes, matches) = surface::find(&runs, spec, replacement, FileFormat::Ppt, limits)?;
        Ok((commit(file, &texts, &keys, changes, limits)?, matches))
    })()
    .map_err(|error: Error| error.with_format(FileFormat::Ppt));
    file.check_source()?;
    result
}

/// Fill `record` into every text shape and notes page. `None` means
/// nothing matched.
pub(crate) fn merge(
    file: &Ppt,
    record: &Record,
    options: &Options,
    limits: &SemanticLimits,
) -> Result<(Option<WriteCommit>, Summary)> {
    file.check_source()?;
    let result = (|| {
        let texts = read(file)?;
        let (runs, keys) = runs(&texts, &[Scope::Every]);
        let mut planner = Planner::new(record, options, limits);
        let changes = surface::merge(&runs, &mut planner, false, FileFormat::Ppt, limits)?;
        let summary = planner.finish();
        if changes.is_empty() {
            return Ok((None, summary));
        }
        Ok((Some(commit(file, &texts, &keys, changes, limits)?), summary))
    })()
    .map_err(|error: Error| error.with_format(FileFormat::Ppt));
    file.check_source()?;
    result
}

/// Every text shape (group members included) and notes page, in order.
fn read(file: &Ppt) -> Result<Vec<Text>> {
    let nav = Nav::new(file);
    let root = nav.root()?;
    let mut texts = Vec::new();
    for slide in nav.children(&root, KindSet::of(Kind::Slide))? {
        let slide = slide?;
        let mut pending: Vec<Node<'_>> =
            nav.children(&slide, KindSet::ALL)?.collect::<Result<_>>()?;
        pending.reverse();
        while let Some(node) = pending.pop() {
            match node.item() {
                Item::Shape(item) => {
                    let Some(shape) = node.shape() else {
                        continue;
                    };
                    if has_text_body(shape) {
                        let text = shape.text().map_err(Error::from)?;
                        let target = node
                            .native_shape_target()
                            .map(|target| (target.slide().get(), target.shape().get()));
                        texts.push(Text {
                            holder: Holder::Shape {
                                slide: item.slide,
                                path: item.path.clone(),
                                target,
                            },
                            paragraphs: paragraphs(&text).map(str::to_owned).collect(),
                        });
                    } else if item.kind == Kind::Group {
                        let mut children: Vec<Node<'_>> =
                            nav.children(&node, KindSet::ALL)?.collect::<Result<_>>()?;
                        children.reverse();
                        pending.extend(children);
                    }
                },
                Item::Notes(notes) => texts.push(Text {
                    holder: Holder::Notes(notes.slide),
                    paragraphs: paragraphs(&notes.text).map(str::to_owned).collect(),
                }),
                _ => {},
            }
        }
    }
    Ok(texts)
}

/// The surface of `texts` under `scopes`.
fn runs<'t>(texts: &'t [Text], scopes: &[Scope]) -> (Vec<Run<'t>>, Vec<Key>) {
    let mut runs = Vec::new();
    let mut keys = Vec::new();
    for text in texts {
        for (index, paragraph) in text.paragraphs.iter().enumerate() {
            let scoped = scopes
                .iter()
                .any(|scope| covers(scope, &text.holder, index));
            let editable = match &text.holder {
                Holder::Shape { target, .. } => target.is_some(),
                Holder::Notes(_) => true,
            };
            runs.push(
                Run::new(paragraph.as_str(), runs.len())
                    .scoped(scoped)
                    .editable(editable),
            );
            keys.push(Key {
                holder: text.holder.clone(),
                paragraph: index,
            });
        }
    }
    (runs, keys)
}

fn covers(scope: &Scope, holder: &Holder, paragraph: usize) -> bool {
    match (scope, holder) {
        (Scope::Every, _) => true,
        (Scope::Slide(at), Holder::Shape { slide, .. }) => at == slide,
        (
            Scope::Shape(at, path),
            Holder::Shape {
                slide, path: own, ..
            },
        ) => at == slide && own.starts_with(path),
        (
            Scope::Paragraph(at, path, index),
            Holder::Shape {
                slide, path: own, ..
            },
        ) => at == slide && own == path && *index == paragraph,
        (Scope::Notes(at), Holder::Notes(slide)) => at == slide,
        _ => false,
    }
}

/// One native replacement per changed shape or notes page, in one edit.
fn commit(
    file: &Ppt,
    texts: &[Text],
    keys: &[Key],
    changes: Vec<Change>,
    limits: &SemanticLimits,
) -> Result<WriteCommit> {
    let mut changed: BTreeMap<Holder, Vec<String>> = BTreeMap::new();
    for change in changes {
        let key = &keys[change.run];
        let paragraphs = changed.entry(key.holder.clone()).or_insert_with(|| {
            texts
                .iter()
                .find(|text| text.holder == key.holder)
                .map(|text| text.paragraphs.clone())
                .unwrap_or_default()
        });
        let slot = paragraphs.get_mut(key.paragraph).ok_or_else(|| {
            Error::new(
                Code::InternalError,
                "a PPT paragraph vanished from its shape",
            )
        })?;
        *slot = change.text;
    }
    let mut plans = Vec::with_capacity(changed.len());
    for (holder, paragraphs) in changed {
        let text = paragraphs.join("\r");
        plans.push(match holder {
            Holder::Shape {
                target: Some((slide, shape)),
                ..
            } => Plan::Shape(ShapePlan {
                target: litchi_ppt::text_edit::Target::new(
                    Position::new(slide),
                    Position::new(shape),
                ),
                text: Some(text),
                anchor: None,
            }),
            Holder::Shape { target: None, .. } => {
                return Err(Error::new(
                    Code::InternalError,
                    "a nested PPT shape was planned for a change",
                ));
            },
            Holder::Notes(slide) => Plan::Notes(NotesPlan {
                position: Position::new(slide),
                text,
            }),
        });
    }
    let native = stage_plans(file, plans, limits, limits.max_snapshot_bytes())?;
    finish_commit(file, native, limits)
}
