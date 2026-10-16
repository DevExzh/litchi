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
        // A paragraph without a candidate holds no match (every match lies
        // in one paragraph, and `may_match` is exact when it says no), so
        // only candidates are listed (record 0302).
        let matcher = spec.matcher();
        let (runs, ids) = surface_of(body, &scope, |text| matcher.may_match(text))?;
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
        // `surface::merge` looks for placeholders only in groups holding a
        // `{`: a paragraph without one is not listed.
        let (runs, ids) = surface_of(body, &Scope::Every, |text| text.contains('{'))?;
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

/// The runs of every paragraph and heading in `scope` whose text `wanted`
/// accepts, and their ids. Every run's text is decoded once; a paragraph
/// `wanted` rejects (its runs' text, joined) lists nothing, and its runs'
/// scope and rewritability are not computed.
fn surface_of<'b>(
    body: Body<'b>,
    scope: &Scope,
    wanted: impl Fn(&str) -> bool,
) -> Result<(Vec<Run<'b>>, Vec<NodeId>)> {
    let mut runs = Vec::new();
    let mut ids = Vec::new();
    let mut group = 0usize;
    // The current paragraph's runs, with their groups, before the filter.
    let mut pending: Vec<(LeafNode<'b>, std::borrow::Cow<'b, str>, usize)> = Vec::new();
    let mut joined = String::new();
    for paragraph in body.nodes() {
        if !matches!(paragraph.kind(), LeafKind::Paragraph | LeafKind::Heading) {
            continue;
        }
        let range = paragraph.range();
        let wanted_paragraph = match scope {
            Scope::Every => !in_comment(paragraph),
            // A paragraph is searched when a scope holds it, or lies in it
            // (a run scope).
            Scope::Ranges(ranges) => ranges.iter().any(|scope| overlaps(scope, &range)),
        };
        if !wanted_paragraph {
            continue;
        }
        group += 1;
        pending.clear();
        let mut push = |run: LeafNode<'b>, group: usize| -> Result<()> {
            let text = run.text().map_err(super::super::core_error)?;
            pending.push((run, text, group));
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
        // A match or placeholder lies inside one group, so inside the
        // paragraph's joined text.
        let candidate = match pending.as_slice() {
            [] => false,
            [(_, text, _)] => wanted(text),
            several => {
                joined.clear();
                for (_, text, _) in several {
                    joined.push_str(text);
                }
                wanted(&joined)
            },
        };
        if !candidate {
            continue;
        }
        for (run, text, group) in pending.drain(..) {
            let scoped = match scope {
                Scope::Every => true,
                Scope::Ranges(ranges) => {
                    let own = run.range();
                    ranges
                        .iter()
                        .any(|scope| scope.start <= own.start && own.end <= scope.end)
                },
            };
            runs.push(
                Run::new(text, group)
                    .scoped(scoped)
                    .editable(rewritable(run)),
            );
            ids.push(run.id());
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
    surface::odf_text_only(content)
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
