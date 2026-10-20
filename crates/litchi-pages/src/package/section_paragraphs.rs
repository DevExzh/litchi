//! Paragraph-level transactions over one section's body text
//! (IWORK_API_REQUESTS.md request 12).
//!
//! A section's paragraphs are its body text split at the paragraph break
//! (`\n`). An [`Edit`] stages paragraph operations — set a paragraph's text,
//! insert, remove, relocate, swap — against the section's immutable text and
//! commits them through the exact-source section-text transaction
//! ([`crate::section::text`]), which owns every native preservation and
//! verification rule:
//!
//! - a commit publishes exactly what committing each operation on its own,
//!   in order, publishes (record 0314);
//! - consecutive text sets share **one** section-text transaction (one
//!   splice from the text before them to the text after them) when that is
//!   provably the same: the splice consumes no object anchor (an inline
//!   table, chart, image or footnote mark), no attribute entry of the body
//!   storage lies inside or between the paragraphs they touch, no set
//!   empties the storage before a later one, and they do not cancel out;
//! - otherwise each operation runs as its own minimal splice, in order, each
//!   over the package the previous one published.
//!
//! Setting a paragraph's text keeps its object anchors: anchors at the start
//! or end of the paragraph stay there, around the new text. A paragraph whose
//! anchors sit between text runs is refused ([`Error::InterleavedAnchors`]),
//! and removing, relocating or swapping a paragraph that holds an anchor is
//! refused ([`Error::AnchoredParagraph`]): deleting or moving an anchored
//! object is the object's own capability.

use std::fmt;

use litchi_core::Position;
use litchi_iwa_text::TextSpan;
use thiserror::Error as ThisError;

use super::Package;
use crate::SectionSelector;
use crate::section::text as section_text;

/// The paragraph break of Pages section text.
const BREAK: char = '\n';
/// The most operations one edit stages.
const MAX_OPERATIONS: usize = 1 << 20;

/// A finite resource governed while paragraphs are edited.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum LimitKind {
    /// Operations staged on one edit.
    Operations,
}

impl fmt::Display for LimitKind {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter.write_str(match self {
            Self::Operations => "operations",
        })
    }
}

/// An error raised while staging or publishing paragraph operations.
///
/// Diagnostics never carry authored text.
#[derive(Debug, Clone, Copy, PartialEq, Eq, ThisError)]
#[non_exhaustive]
pub enum Error {
    /// A well-formed request this route refuses; [`Refusal::kind`] classifies it
    /// (record 0351 §2.4, record 0361).
    #[error(transparent)]
    Refused(Refusal),
    /// The underlying section-text transaction refused the change.
    #[error(transparent)]
    Text(#[from] section_text::Error),
    /// A finite transaction resource ceiling was exceeded.
    #[error("Pages paragraph {kind} limit exceeded: observed {observed}, maximum {maximum}")]
    Limit {
        /// Resource category that exceeded its limit.
        kind: LimitKind,
        /// Observed or requested amount.
        observed: u64,
        /// Configured maximum.
        maximum: u64,
    },
    /// The patch does not belong to this exact immutable package artifact.
    #[error("the Pages paragraph patch does not match the exact source package")]
    PatchConflict,
}

/// A well-formed request this family's route refuses, classified by
/// [`Refusal::kind`] (record 0351 §2.4, record 0361).
#[derive(Debug, Clone, Copy, PartialEq, Eq, ThisError)]
#[non_exhaustive]
pub enum Refusal {
    /// No paragraph exists at the position.
    #[error("the Pages paragraph position {position:?} does not exist (the section has {count})")]
    PositionNotFound {
        /// Requested paragraph position.
        position: Position,
        /// Paragraphs in the staged section.
        count: usize,
    },
    /// Paragraph text cannot hold a paragraph break.
    #[error("a Pages paragraph's text cannot contain a paragraph break")]
    ParagraphBreak,
    /// The paragraph's object anchors sit between text runs, so new text has
    /// no unambiguous place.
    #[error("the Pages paragraph interleaves object anchors with text")]
    InterleavedAnchors,
    /// Removing, moving or swapping a paragraph that anchors an object
    /// (inline table, chart, image or footnote mark) would delete or move
    /// the object, which is the object's own capability.
    #[error("the Pages paragraph anchors an object; remove or move the object itself")]
    AnchoredParagraph,
    /// A section keeps at least one paragraph.
    #[error("a Pages section keeps at least one paragraph")]
    LastParagraph,
}

impl Refusal {
    /// The shared refusal kind litchi-ops classifies this reason by.
    #[must_use]
    pub const fn kind(&self) -> litchi_core::edit::RefusalKind {
        use litchi_core::edit::RefusalKind;
        match self {
            Self::PositionNotFound { .. } => RefusalKind::NotFound,
            Self::ParagraphBreak { .. } | Self::InterleavedAnchors { .. } => {
                RefusalKind::UnsupportedContent
            },
            Self::AnchoredParagraph { .. } | Self::LastParagraph { .. } => {
                RefusalKind::UnsupportedTarget
            },
        }
    }
}

impl From<Refusal> for Error {
    fn from(refusal: Refusal) -> Self {
        Self::Refused(refusal)
    }
}

/// Whether `character` is a native anchor: an inline object marker or a
/// footnote anchor.
const fn is_anchor(character: char) -> bool {
    matches!(character, '\u{fffc}' | '\u{000e}')
}

/// The paragraphs of a section text.
fn paragraphs(text: &str) -> Vec<String> {
    text.split(BREAK).map(str::to_owned).collect()
}

/// `paragraph` with its text replaced by `text`, keeping its anchors.
fn with_text(paragraph: &str, text: &str) -> Result<String, Error> {
    if !paragraph.contains(is_anchor) {
        return Ok(text.to_owned());
    }
    let leading: String = paragraph.chars().take_while(|c| is_anchor(*c)).collect();
    let rest = &paragraph[leading.len()..];
    let body_end = rest.trim_end_matches(is_anchor).len();
    let (body, trailing) = rest.split_at(body_end);
    if body.contains(is_anchor) {
        return Err(Error::Refused(Refusal::InterleavedAnchors));
    }
    Ok(format!("{leading}{text}{trailing}"))
}

/// One staged paragraph operation.
#[derive(Debug, Clone, PartialEq, Eq)]
enum Operation {
    Set(usize, String),
    Insert(usize, String),
    Remove(usize),
    Relocate(usize, usize),
    Swap(usize, usize),
}

impl Operation {
    /// Apply this operation to `paragraphs`.
    fn apply(&self, paragraphs: &mut Vec<String>) -> Result<(), Error> {
        let count = paragraphs.len();
        let check = |index: usize, bound: usize| {
            if index >= bound {
                Err(Error::Refused(Refusal::PositionNotFound {
                    position: Position::new(index),
                    count,
                }))
            } else {
                Ok(())
            }
        };
        match self {
            Self::Set(index, text) => {
                check(*index, count)?;
                paragraphs[*index] = with_text(&paragraphs[*index], text)?;
            },
            Self::Insert(index, text) => {
                check(*index, count + 1)?;
                paragraphs.insert(*index, text.clone());
            },
            Self::Remove(index) => {
                check(*index, count)?;
                if count == 1 {
                    return Err(Error::Refused(Refusal::LastParagraph));
                }
                paragraphs.remove(*index);
            },
            Self::Relocate(from, to) => {
                check(*from, count)?;
                check(*to, count)?;
                let paragraph = paragraphs.remove(*from);
                paragraphs.insert(*to, paragraph);
            },
            Self::Swap(left, right) => {
                check(*left, count)?;
                check(*right, count)?;
                paragraphs.swap(*left, *right);
            },
        }
        Ok(())
    }
}

/// The minimal splice turning `before` into `after`: the UTF-16 span of
/// `before` between their common prefix and suffix, and its replacement.
fn splice<'a>(before: &str, after: &'a str) -> Option<(TextSpan, &'a str)> {
    if before == after {
        return None;
    }
    let prefix: usize = before
        .chars()
        .zip(after.chars())
        .take_while(|(left, right)| left == right)
        .map(|(character, _)| character.len_utf8())
        .sum();
    let suffix: usize = before[prefix..]
        .chars()
        .rev()
        .zip(after[prefix..].chars().rev())
        .take_while(|(left, right)| left == right)
        .map(|(character, _)| character.len_utf8())
        .sum();
    let start = before[..prefix].encode_utf16().count();
    let end = start + before[prefix..before.len() - suffix].encode_utf16().count();
    let replacement = &after[prefix..after.len() - suffix];
    TextSpan::from_utf16_indexes(start, end)
        .ok()
        .map(|span| (span, replacement))
}

/// Whether the splice of `before` into `after` consumes or inserts an
/// anchor.
fn touches_anchor(before: &str, span: TextSpan, replacement: &str) -> bool {
    let start = u64::from(span.start().utf16_index());
    let end = u64::from(span.end().utf16_index());
    let mut units = 0_u64;
    for character in before.chars() {
        let width = u64::try_from(character.len_utf16()).unwrap_or(2);
        if units >= start && units + width <= end && is_anchor(character) {
            return true;
        }
        units += width;
    }
    replacement.contains(is_anchor)
}

/// A staged paragraph edit of one section.
pub struct Edit<'a> {
    source: &'a Package,
    position: Position,
    before: &'a str,
    operations: Vec<Operation>,
    /// The paragraphs with every staged operation applied, kept current so
    /// staging is linear in the operation count.
    current: Vec<String>,
}

impl fmt::Debug for Edit<'_> {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("Edit")
            .field("position", &self.position)
            .field("operations", &self.operations.len())
            .finish_non_exhaustive()
    }
}

impl<'a> Edit<'a> {
    fn stage(&mut self, operation: Operation) -> Result<&mut Self, Error> {
        if self.operations.len() >= MAX_OPERATIONS {
            return Err(Error::Limit {
                kind: LimitKind::Operations,
                observed: u64::try_from(self.operations.len() + 1).unwrap_or(u64::MAX),
                maximum: u64::try_from(MAX_OPERATIONS).unwrap_or(u64::MAX),
            });
        }
        // Apply to the running paragraphs first: a refusal names the
        // operation that caused it, and `apply` validates before it mutates,
        // so a refused operation leaves the edit as it was.
        operation.apply(&mut self.current)?;
        self.operations.push(operation);
        Ok(self)
    }

    fn checked_text(text: &str) -> Result<&str, Error> {
        if text.contains(BREAK) {
            return Err(Error::Refused(Refusal::ParagraphBreak));
        }
        if text.contains(is_anchor) || text.contains('\u{0004}') {
            return Err(Error::Text(section_text::Error::Refused(
                crate::section::text::Refusal::ObjectMarkerReplacement,
            )));
        }
        Ok(text)
    }

    /// The section's paragraphs with every staged operation applied.
    ///
    /// # Errors
    ///
    /// Infallible for operations that staged successfully.
    pub fn paragraphs(&self) -> Result<Vec<String>, Error> {
        Ok(self.current.clone())
    }

    /// The selected section's source position.
    #[must_use]
    pub const fn position(&self) -> Position {
        self.position
    }

    /// Stage replacing the text of the paragraph at `index`, keeping its
    /// object anchors.
    ///
    /// # Errors
    ///
    /// [`Refusal::PositionNotFound`], [`Refusal::ParagraphBreak`],
    /// [`Refusal::InterleavedAnchors`], or a marker refusal for text holding a
    /// native marker character.
    pub fn set(&mut self, index: Position, text: &str) -> Result<&mut Self, Error> {
        let text = Self::checked_text(text)?;
        self.stage(Operation::Set(index.get(), text.to_owned()))
    }

    /// Stage inserting a paragraph so it lands at `index` (`index` equal to
    /// the paragraph count appends).
    ///
    /// # Errors
    ///
    /// As [`Self::set`].
    pub fn insert(&mut self, index: Position, text: &str) -> Result<&mut Self, Error> {
        let text = Self::checked_text(text)?;
        self.stage(Operation::Insert(index.get(), text.to_owned()))
    }

    /// Stage removing the paragraph at `index`.
    ///
    /// # Errors
    ///
    /// [`Refusal::PositionNotFound`]; [`Refusal::LastParagraph`] for a section's
    /// only paragraph.
    pub fn remove(&mut self, index: Position) -> Result<&mut Self, Error> {
        self.stage(Operation::Remove(index.get()))
    }

    /// Stage moving the paragraph at `from` so it lands at `to`.
    ///
    /// # Errors
    ///
    /// [`Refusal::PositionNotFound`].
    pub fn relocate(&mut self, from: Position, to: Position) -> Result<&mut Self, Error> {
        self.stage(Operation::Relocate(from.get(), to.get()))
    }

    /// Stage exchanging two paragraphs.
    ///
    /// # Errors
    ///
    /// [`Refusal::PositionNotFound`].
    pub fn swap(&mut self, left: Position, right: Position) -> Result<&mut Self, Error> {
        self.stage(Operation::Swap(left.get(), right.get()))
    }

    /// Publish the staged operations: exactly what committing each one on
    /// its own, in order, publishes (record 0314), in as few section-text
    /// transactions as that allows.
    ///
    /// Consecutive text sets share one transaction (one splice from the
    /// text before them to the text after them) while that splice is
    /// provably what their own splices publish one after another: no
    /// attribute entry of the body storage lies inside the paragraphs they
    /// touch or between them, the splice consumes no object anchor, no set
    /// empties the storage before a later one, and the sets do not cancel
    /// out. Any other run of sets, and every insertion, removal, relocation
    /// and swap, publishes one transaction per operation.
    ///
    /// # Errors
    ///
    /// The first refusal, leaving the source package untouched.
    pub fn commit(self) -> Result<Commit, Error> {
        let mut publisher = Publisher {
            position: self.position,
            package: self.source.clone(),
            steps: Vec::new(),
        };
        if let Some(single) = self.single_splice()? {
            if let Some((span, replacement)) = single {
                publisher.replace(span, &replacement)?;
            }
            return Ok(publisher.into_commit());
        }
        let operations = &self.operations;
        let mut paragraphs = paragraphs(self.before);
        let mut group = Group::new(self.before.to_owned(), 0);
        for (index, operation) in operations.iter().enumerate() {
            let Operation::Set(line, _) = operation else {
                group.publish(
                    &mut publisher,
                    &operations[group.start..index],
                    &mut paragraphs,
                )?;
                publisher.each(std::slice::from_ref(operation), &mut paragraphs)?;
                group = Group::new(paragraphs.join("\n"), index + 1);
                continue;
            };
            if index > group.start && !group.admits(&publisher, &paragraphs, *line)? {
                group.publish(
                    &mut publisher,
                    &operations[group.start..index],
                    &mut paragraphs,
                )?;
                group = Group::new(paragraphs.join("\n"), index);
            }
            group.touch(*line);
            let before = paragraphs.get(*line).cloned();
            operation.apply(&mut paragraphs)?;
            group.changed |= before.as_ref() != paragraphs.get(*line);
        }
        group.publish(&mut publisher, &operations[group.start..], &mut paragraphs)?;
        Ok(publisher.into_commit())
    }

    /// The text after every staged operation, when they are all text sets
    /// that one splice from the text before them publishes exactly (the
    /// conditions of [`Self::commit`] for the whole sequence, checked
    /// without replaying it); `None` sends the commit to the per-group
    /// path.
    #[allow(
        clippy::option_option,
        reason = "`None` sends the commit to the per-group path; `Some(None)` publishes nothing"
    )]
    fn single_splice(&self) -> Result<Option<Option<(TextSpan, String)>>, Error> {
        let mut touched: Option<(usize, usize)> = None;
        for operation in &self.operations {
            let Operation::Set(line, _) = operation else {
                return Ok(None);
            };
            touched = Some(touched.map_or((*line, *line), |(first, last)| {
                (first.min(*line), last.max(*line))
            }));
        }
        let Some((first, last)) = touched else {
            return Ok(Some(None));
        };
        // One set is its own transaction; an emptied single paragraph, a
        // splice that touches an anchor, or a sequence that ends where it
        // began takes the per-group path (which refuses or replays them).
        let after = self.current.join("\n");
        let Some((span, replacement)) = splice(self.before, &after) else {
            return Ok((self.operations.len() == 1).then_some(None));
        };
        if touches_anchor(self.before, span, replacement) {
            return Ok(None);
        }
        let single = Some(Some((span, replacement.to_owned())));
        if self.operations.len() == 1 {
            return Ok(single);
        }
        if self.current.len() == 1 && !self.before.is_empty() {
            return Ok(None);
        }
        let mut unit = 0_usize;
        let mut hull = None;
        for (line, paragraph) in self.before.split(BREAK).enumerate() {
            let units = if paragraph.is_ascii() {
                paragraph.len()
            } else {
                paragraph.encode_utf16().count()
            };
            let end = unit.saturating_add(units);
            if line == first {
                hull = Some(unit..end);
            }
            if line == last {
                hull = hull.map(|hull| hull.start..end);
                break;
            }
            unit = end.saturating_add(1);
        }
        let Some(hull) = hull else {
            return Ok(None);
        };
        let boundaries = super::section_text::section_boundaries(self.source, self.position)?;
        Ok(boundaries.admits(hull).then_some(single).flatten())
    }
}

/// The package a paragraph commit publishes onto, step by step.
struct Publisher {
    position: Position,
    package: Package,
    steps: Vec<section_text::Patch>,
}

impl Publisher {
    fn into_commit(self) -> Commit {
        let transactions = self.steps.len();
        Commit {
            package: self.package,
            patch: Patch { steps: self.steps },
            diagnostics: Diagnostics {
                changed: transactions != 0,
                transactions,
            },
        }
    }

    /// Publish one section-text splice from `before` to `after`, if they
    /// differ. A splice that consumes or inserts an anchor is refused.
    fn splice(&mut self, before: &str, after: &str) -> Result<(), Error> {
        let Some((span, replacement)) = splice(before, after) else {
            return Ok(());
        };
        if touches_anchor(before, span, replacement) {
            return Err(Error::Refused(Refusal::AnchoredParagraph));
        }
        self.replace(span, replacement)
    }

    /// Publish one checked section-text splice.
    fn replace(&mut self, span: TextSpan, replacement: &str) -> Result<(), Error> {
        let mut edit = self.package.edit_section_text(self.position)?;
        edit.replace(span, replacement)?;
        let commit = edit.commit()?;
        self.steps.push(commit.patch().clone());
        self.package = commit.into_package();
        Ok(())
    }

    /// Publish `operations` one transaction each, over `paragraphs`.
    fn each(
        &mut self,
        operations: &[Operation],
        paragraphs: &mut Vec<String>,
    ) -> Result<(), Error> {
        for operation in operations {
            let before = paragraphs.join("\n");
            operation.apply(paragraphs)?;
            self.splice(&before, &paragraphs.join("\n"))?;
        }
        Ok(())
    }
}

/// Consecutive text sets that may publish as one splice.
struct Group {
    /// The section text before the group's first set.
    base: String,
    /// The group's first operation.
    start: usize,
    /// The first and last paragraph the group's sets touch.
    touched: Option<(usize, usize)>,
    /// Whether any set changed its paragraph.
    changed: bool,
    /// The UTF-16 offset of each paragraph of `base`, then its length plus
    /// one, computed when a second set asks to join.
    offsets: Option<Vec<usize>>,
    /// The storage boundaries of the group's base package, read when a
    /// second set asks to join.
    boundaries: Option<super::section_text::SectionBoundaries>,
}

impl Group {
    const fn new(base: String, start: usize) -> Self {
        Self {
            base,
            start,
            touched: None,
            changed: false,
            offsets: None,
            boundaries: None,
        }
    }

    fn touch(&mut self, line: usize) {
        self.touched = Some(match self.touched {
            Some((first, last)) => (first.min(line), last.max(line)),
            None => (line, line),
        });
    }

    /// Whether a set of paragraph `line` may join the group, whose sets so
    /// far produced `paragraphs`.
    fn admits(
        &mut self,
        publisher: &Publisher,
        paragraphs: &[String],
        line: usize,
    ) -> Result<bool, Error> {
        // An emptied storage loses its text field; a later splice adds it
        // back last, which one splice would not.
        if publisher.package.sections().len() == 1
            && !self.base.is_empty()
            && paragraphs.iter().all(String::is_empty)
        {
            return Ok(false);
        }
        let Some((first, last)) = self.touched else {
            return Ok(true);
        };
        let (first, last) = (first.min(line), last.max(line));
        let offsets = self.offsets.get_or_insert_with(|| {
            let mut offsets = Vec::new();
            let mut unit = 0_usize;
            for paragraph in self.base.split(BREAK) {
                offsets.push(unit);
                unit = unit
                    .saturating_add(paragraph.encode_utf16().count())
                    .saturating_add(1);
            }
            offsets.push(unit);
            offsets
        });
        // Paragraph `last` ends just before the break that follows it.
        let (Some(&start), Some(&next)) = (offsets.get(first), offsets.get(last + 1)) else {
            return Ok(false);
        };
        let hull = start..next.saturating_sub(1);
        if self.boundaries.is_none() {
            self.boundaries = Some(super::section_text::section_boundaries(
                &publisher.package,
                publisher.position,
            )?);
        }
        Ok(self
            .boundaries
            .as_ref()
            .is_some_and(|boundaries| boundaries.admits(hull)))
    }

    /// Publish the group's `sets`, which produced `paragraphs`: one splice
    /// for one set, or for several that consume no anchor and do not cancel
    /// out; otherwise each on its own, replayed from the group's base.
    fn publish(
        &self,
        publisher: &mut Publisher,
        sets: &[Operation],
        paragraphs: &mut Vec<String>,
    ) -> Result<(), Error> {
        if sets.is_empty() {
            return Ok(());
        }
        let after = paragraphs.join("\n");
        let anchored = splice(&self.base, &after)
            .is_some_and(|(span, replacement)| touches_anchor(&self.base, span, replacement));
        let cancelled = after == self.base && self.changed;
        if sets.len() == 1 || (!anchored && !cancelled) {
            return publisher.splice(&self.base, &after);
        }
        *paragraphs = self::paragraphs(&self.base);
        publisher.each(sets, paragraphs)
    }
}

/// The reversible proof of one paragraph edit: its section-text patches,
/// in order.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Patch {
    steps: Vec<section_text::Patch>,
}

impl Patch {
    /// Whether the patch changes nothing.
    #[must_use]
    pub fn is_noop(&self) -> bool {
        self.steps.iter().all(section_text::Patch::is_noop)
    }

    /// The patch from the target back to the source: each step inverted, in
    /// reverse order.
    #[must_use]
    pub fn inverse(&self) -> Self {
        Self {
            steps: self
                .steps
                .iter()
                .rev()
                .map(section_text::Patch::inverse)
                .collect(),
        }
    }

    /// The section-text transactions, in order.
    #[must_use]
    pub fn steps(&self) -> &[section_text::Patch] {
        &self.steps
    }
}

/// Compact evidence describing one paragraph commit.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct Diagnostics {
    changed: bool,
    transactions: usize,
}

impl Diagnostics {
    /// Whether the package differs from its source.
    #[must_use]
    pub const fn changed(self) -> bool {
        self.changed
    }

    /// The section-text transactions the commit ran (zero for a no-op, one
    /// when the change coalesced into a single splice).
    #[must_use]
    pub const fn transactions(self) -> usize {
        self.transactions
    }
}

/// The verified result of one paragraph transaction.
#[must_use = "a Pages paragraph commit contains the validated package snapshot"]
#[derive(Debug)]
pub struct Commit {
    package: Package,
    patch: Patch,
    diagnostics: Diagnostics,
}

impl Commit {
    /// Borrow the published package.
    #[must_use]
    pub const fn package(&self) -> &Package {
        &self.package
    }

    /// Consume the commit and return its package.
    #[must_use]
    pub fn into_package(self) -> Package {
        self.package
    }

    /// Consume the commit and return its package and patch.
    #[must_use]
    pub fn into_parts(self) -> (Package, Patch) {
        (self.package, self.patch)
    }

    /// Borrow the reversible patch.
    #[must_use]
    pub const fn patch(&self) -> &Patch {
        &self.patch
    }

    /// Borrow compact publication diagnostics.
    #[must_use]
    pub const fn diagnostics(&self) -> &Diagnostics {
        &self.diagnostics
    }
}

impl Package {
    /// The paragraphs of one section: its body text split at paragraph
    /// breaks. Anchored objects appear as their marker characters.
    ///
    /// # Errors
    ///
    /// The section-text selection errors.
    pub fn section_paragraphs<'selector>(
        &self,
        selector: impl Into<SectionSelector<'selector>>,
    ) -> Result<Vec<&str>, Error> {
        Ok(self.section_text(selector)?.split(BREAK).collect())
    }

    /// The body tables each paragraph of one section anchors: for every
    /// paragraph, in order, the catalog positions
    /// ([`crate::table::catalog::Snapshot::position`]) of the tables whose
    /// object anchor sits in it (record 0301).
    ///
    /// A paragraph whose anchor count is larger than its table count also
    /// anchors other objects (charts, images, footnotes).
    ///
    /// # Errors
    ///
    /// The section-text selection errors; [`section_text::Error::InvalidSource`]
    /// when the rooted body tables cannot be resolved.
    pub fn section_paragraph_tables<'selector>(
        &self,
        selector: impl Into<SectionSelector<'selector>>,
    ) -> Result<Vec<Vec<Position>>, Error> {
        let edit = self.edit_section_text(selector)?;
        let position = edit.position();
        let text = edit.text();
        let start = super::section_text::section_start(self, position)?;
        let invalid = || Error::Text(section_text::Error::InvalidSource);
        let mut budget = super::table_lock::WireBudget::new(self.state.source.limits())
            .map_err(|_error| invalid())?;
        let targets = super::table_lock::native_body_table_targets_with_budget(self, &mut budget)
            .unwrap_or_default();
        let mut ranges = Vec::new();
        let mut offset = start;
        for paragraph in text.split(BREAK) {
            let length = paragraph.encode_utf16().count();
            ranges.push(offset..offset + length);
            offset += length + 1;
        }
        let mut tables = vec![Vec::new(); ranges.len()];
        for target in targets {
            if let Some(index) = ranges
                .iter()
                .position(|range| range.contains(&target.anchor_character_index))
            {
                tables[index].push(Position::new(target.table_position));
            }
        }
        Ok(tables)
    }

    /// Start a paragraph edit of one section.
    ///
    /// # Errors
    ///
    /// The section-text selection errors.
    pub fn edit_section_paragraphs<'selector>(
        &self,
        selector: impl Into<SectionSelector<'selector>>,
    ) -> Result<Edit<'_>, Error> {
        let edit = self.edit_section_text(selector)?;
        let before = edit.text();
        Ok(Edit {
            source: self,
            position: edit.position(),
            before,
            operations: Vec::new(),
            current: paragraphs(before),
        })
    }

    /// Apply a paragraph patch to the exact package it was made from.
    ///
    /// # Errors
    ///
    /// [`Error::PatchConflict`] (through the first step) unless this package
    /// is the patch's source.
    pub fn apply_section_paragraphs(&self, patch: &Patch) -> Result<Commit, Error> {
        let mut package = self.clone();
        let mut changed = false;
        for step in &patch.steps {
            let commit = package
                .apply_section_text(step)
                .map_err(|error| match error {
                    section_text::Error::PatchConflict => Error::PatchConflict,
                    other => Error::Text(other),
                })?;
            changed |= commit.diagnostics().changed();
            package = commit.into_package();
        }
        Ok(Commit {
            package,
            patch: patch.clone(),
            diagnostics: Diagnostics {
                changed,
                transactions: patch.steps.len(),
            },
        })
    }
}

#[cfg(test)]
mod tests {
    use super::{Error, Refusal, splice, touches_anchor, with_text};

    #[test]
    fn set_keeps_leading_and_trailing_anchors() {
        assert_eq!(
            with_text("Old\u{fffc}", "New").ok().as_deref(),
            Some("New\u{fffc}")
        );
        assert_eq!(
            with_text("\u{fffc}Old", "New").ok().as_deref(),
            Some("\u{fffc}New")
        );
        assert_eq!(with_text("Plain", "New").ok().as_deref(), Some("New"));
        assert_eq!(
            with_text("A\u{fffc}B", "New"),
            Err(Error::Refused(Refusal::InterleavedAnchors))
        );
    }

    #[test]
    fn splices_are_minimal_and_anchor_aware() {
        let (span, replacement) = splice("one\ntwo\u{fffc}", "one\nTWO\u{fffc}").unwrap();
        assert_eq!(
            (span.start().utf16_index(), span.end().utf16_index()),
            (4, 7)
        );
        assert_eq!(replacement, "TWO");
        assert!(!touches_anchor("one\ntwo\u{fffc}", span, replacement));
        let (span, replacement) = splice("a\nb\u{fffc}\nc", "b\u{fffc}\na\nc").unwrap();
        assert!(touches_anchor("a\nb\u{fffc}\nc", span, replacement));
        assert!(splice("same", "same").is_none());
        let (span, _) = splice("😀x", "😀y").unwrap();
        assert_eq!(span.start().utf16_index(), 2);
    }
}
