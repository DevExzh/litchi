//! ODT (and OTH/ODM, which share this adapter) find/replace and template
//! merge (record 0295) over the shared text surface
//! (`crate::find::surface`).
//!
//! The leaf body index lists every paragraph and heading of `office:text`
//! (tables, sections, lists, frames, text boxes and notes included) and their
//! runs: `text:span` elements and character data directly inside a
//! paragraph, heading or hyperlink. A paragraph's runs are one group; a
//! hyperlink's edges and anything else inside the paragraph (a note, a
//! frame) start a new group. Bookmarks join the text around them.
//!
//! A run whose content is only text, spaces (`text:s`), tabs and line
//! breaks is rewritten whole through the leaf's `set_text`, which keeps the
//! run's own style and re-encodes spaces, tabs and breaks; a run holding a
//! field, a nested span, a note or a bookmark is fixed: a find that touches
//! it is refused and merge does not replace placeholders in it. Format
//! properties are refused (the leaf formats whole runs and cannot split one
//! at a match). The root searches every paragraph except comments
//! (OfficeCLI's root find skips them); headers and footers live in
//! `styles.xml` and are not searched.

use std::ops::Range;

use litchi_odt::body::{Body, Kind as LeafKind, Node as LeafNode, NodeId};

use super::super::spec::Kind;
use super::super::{Nav, Node, Odt};
use super::{Commit, finish};
use crate::error::{Code, Error, Result};
use crate::find::Spec;
use crate::find::surface::{self, Change, Run};
use crate::merge::{Options, Planner, Record, Summary};
use crate::target::TargetArg;
use crate::transaction::SemanticLimits;

/// What a find covers: everything but comments, or explicit byte ranges
/// of `content.xml` (the nodes the target selects).
enum Scope {
    Every,
    Ranges(Vec<Range<usize>>),
}

/// Find and replace text below `target`.
pub(crate) fn find_replace(
    file: &Odt,
    target: &TargetArg<'_>,
    spec: &Spec<'_>,
    limits: &SemanticLimits,
) -> Result<(Commit, usize)> {
    file.check_source()?;
    let format = file.format();
    let result = (|| {
        let replacement = surface::text_only(
            spec,
            format,
            "the leaf formats whole runs and cannot split one at a match",
        )?;
        file.admit_edit()?;
        let mut every = false;
        let mut ranges = Vec::new();
        surface::each_target(Nav::new(file), target, |node: Node<'_>| {
            if node.element_kind() == Kind::Document {
                every = true;
                return Ok(());
            }
            match node.leaf() {
                Some(leaf) if leaf.kind() == LeafKind::Body => every = true,
                Some(leaf) => ranges.push(leaf.range()),
                None => {
                    return Err(Error::new(
                        Code::UnsupportedType,
                        "an ODT find searches the text body",
                    )
                    .with_suggestion("target the document, the body, or a body element"));
                },
            }
            Ok(())
        })?;
        let scope = if every {
            Scope::Every
        } else {
            Scope::Ranges(ranges)
        };
        let body = file.body()?;
        let (runs, ids) = surface_of(body, &scope)?;
        let (changes, matches) = surface::find(&runs, spec, replacement, format, limits)?;
        Ok((commit(file, &ids, changes, limits)?, matches))
    })()
    .map_err(|error: Error| error.with_format(format));
    file.check_source()?;
    result
}

/// Fill `record` into every paragraph of the body. `None` means nothing
/// matched.
pub(crate) fn merge(
    file: &Odt,
    record: &Record,
    options: &Options,
    limits: &SemanticLimits,
) -> Result<(Option<Commit>, Summary)> {
    file.check_source()?;
    let format = file.format();
    let result = (|| {
        file.admit_edit()?;
        let body = file.body()?;
        let (runs, ids) = surface_of(body, &Scope::Every)?;
        let mut planner = Planner::new(record, options, limits);
        let changes = surface::merge(&runs, &mut planner, true, format, limits)?;
        let summary = planner.finish();
        if changes.is_empty() {
            return Ok((None, summary));
        }
        Ok((Some(commit(file, &ids, changes, limits)?), summary))
    })()
    .map_err(|error: Error| error.with_format(format));
    file.check_source()?;
    result
}

/// The runs of every paragraph and heading in `scope`, and their ids.
fn surface_of<'b>(body: Body<'b>, scope: &Scope) -> Result<(Vec<Run<'b>>, Vec<NodeId>)> {
    let mut runs = Vec::new();
    let mut ids = Vec::new();
    let mut group = 0usize;
    for paragraph in body.nodes() {
        if !matches!(paragraph.kind(), LeafKind::Paragraph | LeafKind::Heading) {
            continue;
        }
        let range = paragraph.range();
        let wanted = match scope {
            Scope::Every => !in_comment(paragraph),
            // A paragraph is searched when a scope holds it, or lies in it
            // (a run scope).
            Scope::Ranges(ranges) => ranges.iter().any(|scope| overlaps(scope, &range)),
        };
        if !wanted {
            continue;
        }
        group += 1;
        let mut push = |run: LeafNode<'b>, group: usize| -> Result<()> {
            let scoped = match scope {
                Scope::Every => true,
                Scope::Ranges(ranges) => {
                    let own = run.range();
                    ranges
                        .iter()
                        .any(|scope| scope.start <= own.start && own.end <= scope.end)
                },
            };
            let text = run.text().map_err(super::super::core_error)?;
            runs.push(
                Run::new(text, group)
                    .scoped(scoped)
                    .editable(rewritable(run)),
            );
            ids.push(run.id());
            Ok(())
        };
        for child in paragraph.children() {
            match child.kind() {
                LeafKind::Span | LeafKind::Text => push(child, group)?,
                LeafKind::Hyperlink => {
                    group += 1;
                    for inner in child.children() {
                        match inner.kind() {
                            LeafKind::Span | LeafKind::Text => push(inner, group)?,
                            LeafKind::Bookmark => {},
                            _ => group += 1,
                        }
                    }
                    group += 1;
                },
                LeafKind::Bookmark => {},
                _ => group += 1,
            }
        }
    }
    Ok((runs, ids))
}

fn overlaps(left: &Range<usize>, right: &Range<usize>) -> bool {
    left.start < right.end && right.start < left.end
}

/// Whether `node` lies in a comment.
fn in_comment(node: LeafNode<'_>) -> bool {
    let mut current = node.parent();
    while let Some(parent) = current {
        if parent.kind() == LeafKind::Annotation {
            return true;
        }
        current = parent.parent();
    }
    false
}

/// Whether the leaf can rewrite `run` as text without losing content: no
/// indexed node inside it, and no markup but spaces, tabs and breaks.
fn rewritable(run: LeafNode<'_>) -> bool {
    if run.encloses_indexed() {
        return false;
    }
    let (range, inner) = (run.range(), run.content_range());
    let Some(content) = run
        .xml()
        .get(inner.start.saturating_sub(range.start)..inner.end.saturating_sub(range.start))
    else {
        return false;
    };
    let mut rest = content;
    while let Some(start) = rest.find('<') {
        let tag = &rest[start + 1..];
        let end = tag.find('>').unwrap_or(tag.len());
        let name = tag[..end]
            .trim_start_matches('/')
            .split(|character: char| character.is_whitespace() || character == '/')
            .next()
            .unwrap_or("");
        let local = name.rsplit(':').next().unwrap_or(name);
        if !matches!(local, "s" | "tab" | "line-break" | "soft-page-break") {
            return false;
        }
        rest = &tag[end..];
    }
    true
}

/// One leaf `set_text` per changed run, in one edit and one commit.
fn commit(
    file: &Odt,
    ids: &[NodeId],
    changes: Vec<Change>,
    limits: &SemanticLimits,
) -> Result<Commit> {
    let mut edit = file.package()?.edit();
    let operations = changes.len();
    for change in changes {
        edit.set_text(ids[change.run], change.text)
            .map_err(super::staging_error)?;
    }
    finish(file, edit, operations, limits)
}
