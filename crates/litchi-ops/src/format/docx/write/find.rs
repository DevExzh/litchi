//! DOCX find/replace (records 0273, 0285): `set` with `find=`.
//!
//! The scope is the target's subtree. The document root searches the main
//! story and, as OfficeCLI's root find does, every header, footer, footnote
//! and endnote; the body searches the main story; a table, row, cell or
//! content control searches its paragraphs (text boxes included, except
//! those inside alternate content); a
//! paragraph searches itself; a hyperlink or a run searches only its own
//! text, as OfficeCLI's run scope does; the comments part and one comment
//! search their paragraphs.
//!
//! Every story is read and rewritten by the leaf's one substitution
//! primitive ([`litchi_docx::document::substitute`], shared with template
//! merge): the leaf lists each paragraph's text segments in groups, the
//! shared matcher finds the matches of each group, and one leaf call per
//! story replaces them. Format properties are applied by the leaf to the
//! runs the replacement splits out, through the ordinary DOCX run setters.
//!
//! `mc:AlternateContent` wrappers in runs (drawings and text boxes with a
//! fallback, record 0307) are opaque: their text, text boxes included, is
//! not searched, and a match that would span one is refused.

use litchi_core::FileFormat;
use litchi_docx::document::ElementPath;
use litchi_docx::document::substitute::{OBJECT, Paragraph, Replacement, Story, StoryKind};

use super::super::element;
use super::super::merge::owning_editor;
use super::super::node::{Node, WriteSelection};
use super::super::spec::Kind;
use super::super::{Docx, DocxOwner};
use super::{Assignment, Commit, TargetKind, WorkEdit, admit_assignments, property_error};
use crate::error::{Code, Error, Result};
use crate::find::{Matcher, Piece, Scratch, Spec, matches_with};
use crate::registry::Operation;
use crate::target::TargetArg;
use crate::transaction::SemanticLimits;

/// What a find reads.
#[derive(Default)]
struct Reach {
    /// The whole main story.
    body: bool,
    /// Elements of the main story, when not all of it.
    paths: Vec<ElementPath>,
    /// Headers, footers, footnotes and endnotes.
    parts: bool,
    /// The comments part: `Some(None)` every comment, else the items.
    comments: Option<Option<Vec<usize>>>,
}

impl Reach {
    fn main(&self) -> bool {
        self.body || !self.paths.is_empty()
    }

    fn stories(&self) -> bool {
        self.parts || self.comments.is_some()
    }

    /// Whether a story part of `kind` is in reach.
    fn wants(&self, kind: StoryKind) -> bool {
        match kind {
            StoryKind::Comments => self.comments.is_some(),
            _ => self.parts,
        }
    }
}

/// The running totals of one find.
struct Totals {
    matches: usize,
    staged: usize,
    max_bytes: usize,
}

impl Docx {
    /// Find (and replace or format) text in the subtree of `target`,
    /// returning the commit and the match count.
    pub(crate) fn find_replace(
        &self,
        target: &TargetArg<'_>,
        spec: &Spec<'_>,
        limits: &SemanticLimits,
    ) -> Result<(Commit, usize)> {
        let result = self
            .find_inner(target, spec, limits)
            .map_err(|error| error.with_format(FileFormat::Docx));
        self.owner().finish(result)
    }

    fn find_inner(
        &self,
        target: &TargetArg<'_>,
        spec: &Spec<'_>,
        limits: &SemanticLimits,
    ) -> Result<(Commit, usize)> {
        if spec.format().get("text").is_some() {
            return Err(Error::new(
                Code::InvalidArgument,
                "`text` replaces a whole run; with `find`, use `replace`",
            )
            .with_property("text"));
        }
        // Admit the format properties as run properties, values included,
        // before reading.
        let assignments = if spec.formats() {
            let assignments = admit_assignments(TargetKind::Run, spec.format())?;
            let mut failure = None;
            format_run(
                &assignments,
                &mut litchi_docx::properties::Run::default(),
                &mut failure,
            );
            if let Some(error) = failure {
                return Err(error);
            }
            assignments
        } else {
            super::Few::new()
        };
        let nodes = self.resolve_targets(target)?;
        let mut reach = reach(&nodes)?;
        // A source-backed document without the story parts the find reaches
        // keeps its source editor: only the main part's relationships are
        // read to know.
        if reach.stories()
            && let DocxOwner::Source(package) = self.owner()
            && !package
                .story_parts()?
                .iter()
                .any(|part| reach.wants(part.kind()))
        {
            reach.parts = false;
            reach.comments = None;
        }
        // Header, footer, note and comment stories publish only through
        // the owning package's editor.
        let options = litchi_docx::document::EditOptions::default()
            .max_snapshot_bytes(limits.max_snapshot_bytes());
        let editor = if reach.stories() {
            owning_editor(self, options)?
        } else {
            self.edit_with_options(options)?
        };
        let mut work = WorkEdit::from(editor)?;
        let matcher = spec.matcher();
        let mut totals = Totals {
            matches: 0,
            staged: 0,
            max_bytes: limits.max_staged_bytes(),
        };
        let mut scratch = Scratch::default();
        let mut failure = None;
        let mut format = |run: &mut litchi_docx::properties::Run| {
            format_run(&assignments, run, &mut failure);
        };
        if reach.main() {
            let story = both!(&work, |edit| if reach.body {
                edit.story()
            } else {
                edit.story_within(&reach.paths)
            }
            .map_err(Error::from))?;
            let replacements = plan(&story, None, &matcher, spec, &mut scratch, &mut totals)?;
            if !replacements.is_empty() {
                both!(&mut work, |edit| edit
                    .substitute_formatted(&story, &replacements, &mut format)
                    .map_err(Error::from))?;
            }
        }
        if reach.stories() {
            let WorkEdit::Eager(edit) = &mut work else {
                return Err(Error::new(
                    Code::InternalError,
                    "a DOCX story find needs the owning package editor",
                ));
            };
            for part in edit.story_parts()? {
                if !reach.wants(part.kind()) {
                    continue;
                }
                let items = match (part.kind(), &reach.comments) {
                    (StoryKind::Comments, Some(items)) => items.as_deref(),
                    _ => None,
                };
                let story = edit.part_story(part.name())?;
                let replacements = plan(&story, items, &matcher, spec, &mut scratch, &mut totals)?;
                if !replacements.is_empty() {
                    edit.substitute_part_formatted(&story, &replacements, &mut format)?;
                }
            }
        }
        if let Some(error) = failure {
            return Err(error);
        }
        Ok((work.commit(self)?, totals.matches))
    }
}

/// Apply the format properties to one run's direct properties, keeping the
/// first failure.
fn format_run(
    assignments: &[Assignment<'_>],
    run: &mut litchi_docx::properties::Run,
    failure: &mut Option<Error>,
) {
    for assignment in assignments {
        let applied =
            assignment
                .entry
                .prepare(Operation::Set, &assignment.value, |setter, value| {
                    setter.apply_run(run, value)
                });
        if let Err(error) = applied {
            failure.get_or_insert_with(|| property_error(error, &assignment.key));
        }
    }
}

/// The replacements of a find over `story`: the paragraphs of `items` only,
/// when given.
fn plan(
    story: &Story,
    items: Option<&[usize]>,
    matcher: &Matcher<'_>,
    spec: &Spec<'_>,
    scratch: &mut Scratch,
    totals: &mut Totals,
) -> Result<Vec<Replacement>> {
    let mut replacements = Vec::new();
    let mut pieces: Vec<Piece<'_>> = Vec::new();
    let mut offsets: Vec<usize> = Vec::new();
    for (index, paragraph) in story.paragraphs().iter().enumerate() {
        if items.is_some_and(|items| paragraph.item().is_none_or(|item| !items.contains(&item))) {
            continue;
        }
        let text = paragraph.text();
        refuse_alternate_spans(paragraph, matcher, scratch)?;
        if !matcher.may_match(text) {
            continue;
        }
        pieces.clear();
        offsets.clear();
        for segment in paragraph.segments() {
            let range = segment.range();
            offsets.push(range.start);
            pieces.push(Piece {
                text: &text[range],
                group: segment.group(),
                scoped: segment.is_scoped(),
            });
        }
        matches_with(scratch, &pieces, matcher, |_, _, _, found, first| {
            let base = offsets[first];
            totals.matches = totals.matches.saturating_add(found.len());
            for range in found {
                let range = base + range.start..base + range.end;
                let replacement = match spec.replacement() {
                    Some(text) => {
                        totals.staged = totals.staged.saturating_add(text.len());
                        let replacement = Replacement::new(index, range, text);
                        if spec.formats() {
                            replacement.formatted()
                        } else {
                            replacement
                        }
                    },
                    None => Replacement::keep(index, range),
                };
                if totals.staged > totals.max_bytes {
                    return Err(Error::new(
                        Code::LimitExceeded,
                        format!(
                            "find/replace would stage more than {} bytes of text",
                            totals.max_bytes
                        ),
                    ));
                }
                replacements.try_reserve(1).map_err(|source| {
                    Error::from(crate::transaction::Error::Allocation { source })
                })?;
                replacements.push(replacement);
            }
            Ok(())
        })?;
    }
    Ok(replacements)
}

/// Refuse a find with a match that would cross the object of an
/// `mc:AlternateContent` wrapper (a drawing or text box with a fallback;
/// record 0307).
///
/// The leaf lists a wrapper as one opaque object, which ends a text group as
/// a picture does, so no match ever spans it and the wrapper is never split
/// or dropped. Readers, though, show the wrapper's selected branch inline,
/// and one without text not at all: a find that would match across the
/// object if it were transparent is refused rather than reported as no
/// match.
fn refuse_alternate_spans(
    paragraph: &Paragraph,
    matcher: &Matcher<'_>,
    scratch: &mut Scratch,
) -> Result<()> {
    let alternate = paragraph.alternate_content();
    if alternate.is_empty() {
        return Ok(());
    }
    let text = paragraph.text();
    let mut pieces: Vec<Piece<'_>> = Vec::new();
    // Indexes of the pieces joined to the previous one across a wrapper.
    let mut joints: Vec<usize> = Vec::new();
    let mut group = 0usize;
    let mut previous: Option<(usize, usize)> = None;
    for segment in paragraph.segments() {
        let range = segment.range();
        if let Some((original, end)) = previous
            && segment.group() != original
        {
            let gap = text.get(end..range.start).unwrap_or_default();
            let only_wrappers = !gap.is_empty()
                && gap
                    .char_indices()
                    .all(|(at, character)| character == OBJECT && alternate.contains(&(end + at)));
            if only_wrappers {
                joints.push(pieces.len());
            } else {
                group += 1;
            }
        }
        previous = Some((segment.group(), range.end));
        pieces.push(Piece {
            text: &text[range],
            group,
            scoped: segment.is_scoped(),
        });
    }
    if joints.is_empty() {
        return Ok(());
    }
    matches_with(
        scratch,
        &pieces,
        matcher,
        |members, starts, _, found, first| {
            let spans = joints
                .iter()
                .filter(|joint| **joint > first && **joint < first + members.len())
                .map(|joint| starts[joint - first])
                .any(|at| found.iter().any(|range| range.start < at && at < range.end));
            if spans {
                return Err(Error::new(
                Code::UnsupportedCapability,
                "a match would span markup-compatibility alternate content (a drawing or text box with a fallback), which find does not edit",
            )
            .with_property(crate::find::FIND)
            .with_suggestion(
                "find text on one side of the object; its own text is not searched (record 0307)",
            ));
            }
            Ok(())
        },
    )
}

/// What the target nodes search.
fn reach(nodes: &[Node<'_>]) -> Result<Reach> {
    let mut reach = Reach::default();
    let mut body_index = None;
    for node in nodes {
        match node.element_kind() {
            Kind::Document => {
                reach.body = true;
                reach.parts = true;
            },
            Kind::Body => reach.body = true,
            Kind::Table | Kind::Row | Kind::Cell | Kind::Sdt | Kind::Run | Kind::Hyperlink => {
                reach.paths.push(element::element_path(node)?);
            },
            Kind::Paragraph => reach.paths.push(paragraph_path(node, &mut body_index)?),
            Kind::Comments => reach.comments = Some(None),
            Kind::Comment => {
                let item = node
                    .address
                    .steps()
                    .last()
                    .map(|step| step.ordinal as usize)
                    .ok_or_else(|| Error::new(Code::InternalError, "a comment has no address"))?;
                match &mut reach.comments {
                    Some(None) => {},
                    Some(Some(items)) => items.push(item),
                    None => reach.comments = Some(Some(vec![item])),
                }
            },
            _ => {
                return Err(Error::new(
                    Code::UnsupportedCapability,
                    "find searches the document, the body, tables, rows, cells, content controls, paragraphs, hyperlinks, runs and comments",
                )
                .with_property(crate::find::FIND)
                .with_suggestion(
                    "the document root also searches headers, footers, footnotes and endnotes (record 0285)",
                ));
            },
        }
    }
    if reach.body {
        reach.paths.clear();
    }
    Ok(reach)
}

/// The leaf path of a paragraph node: its body position when it is a direct
/// body paragraph, else its element path.
fn paragraph_path(node: &Node<'_>, body: &mut Option<litchi_docx::Blocks>) -> Result<ElementPath> {
    match node.write_selection_cached(body) {
        Ok(WriteSelection::Paragraph(position)) => Ok(ElementPath::new([
            litchi_docx::document::ElementStep::Paragraph(position),
        ])),
        Ok(WriteSelection::Run { .. }) => Err(Error::new(
            Code::InternalError,
            "a DOCX paragraph resolved to a run",
        )),
        Err(error) if error.code() == Code::UnsupportedCapability => element::element_path(node),
        Err(error) => Err(error),
    }
}
