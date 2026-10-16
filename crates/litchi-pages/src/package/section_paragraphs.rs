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
//! - when the whole change is one contiguous splice that consumes no object
//!   anchor (an inline table, chart, image or footnote mark), the commit is a
//!   **single** section-text transaction, however many operations were
//!   staged;
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
    /// The underlying section-text transaction refused the change.
    #[error(transparent)]
    Text(#[from] section_text::Error),
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
    /// A finite transaction resource ceiling was exceeded.
    #[error("Pages paragraph {kind} limit exceeded: observed {observed}, maximum {maximum}")]
    LimitExceeded {
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
        return Err(Error::InterleavedAnchors);
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
                Err(Error::PositionNotFound {
                    position: Position::new(index),
                    count,
                })
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
                    return Err(Error::LastParagraph);
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
            return Err(Error::LimitExceeded {
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
            return Err(Error::ParagraphBreak);
        }
        if text.contains(is_anchor) || text.contains('\u{0004}') {
            return Err(Error::Text(section_text::Error::ObjectMarkerReplacement));
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
    /// [`Error::PositionNotFound`], [`Error::ParagraphBreak`],
    /// [`Error::InterleavedAnchors`], or a marker refusal for text holding a
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
    /// [`Error::PositionNotFound`]; [`Error::LastParagraph`] for a section's
    /// only paragraph.
    pub fn remove(&mut self, index: Position) -> Result<&mut Self, Error> {
        self.stage(Operation::Remove(index.get()))
    }

    /// Stage moving the paragraph at `from` so it lands at `to`.
    ///
    /// # Errors
    ///
    /// [`Error::PositionNotFound`].
    pub fn relocate(&mut self, from: Position, to: Position) -> Result<&mut Self, Error> {
        self.stage(Operation::Relocate(from.get(), to.get()))
    }

    /// Stage exchanging two paragraphs.
    ///
    /// # Errors
    ///
    /// [`Error::PositionNotFound`].
    pub fn swap(&mut self, left: Position, right: Position) -> Result<&mut Self, Error> {
        self.stage(Operation::Swap(left.get(), right.get()))
    }

    /// Publish the staged operations: one section-text transaction when the
    /// whole change is one anchor-free splice, else one per operation.
    ///
    /// # Errors
    ///
    /// The first refusal, leaving the source package untouched.
    pub fn commit(self) -> Result<Commit, Error> {
        let after = self.current.join("\n");
        let Some((span, replacement)) = splice(self.before, &after) else {
            return Ok(Commit {
                package: self.source.clone(),
                patch: Patch { steps: Vec::new() },
                diagnostics: Diagnostics {
                    changed: false,
                    transactions: 0,
                },
            });
        };
        if !touches_anchor(self.before, span, replacement) {
            let mut edit = self.source.edit_section_text(self.position)?;
            edit.replace(span, replacement)?;
            let commit = edit.commit()?;
            let changed = commit.diagnostics().changed();
            let (package, patch) = (commit.package().clone(), commit.patch().clone());
            return Ok(Commit {
                package,
                patch: Patch { steps: vec![patch] },
                diagnostics: Diagnostics {
                    changed,
                    transactions: 1,
                },
            });
        }
        let mut package = self.source.clone();
        let mut steps = Vec::new();
        let mut paragraphs = paragraphs(self.before);
        for operation in &self.operations {
            let before = paragraphs.join("\n");
            operation.apply(&mut paragraphs)?;
            let after = paragraphs.join("\n");
            let Some((span, replacement)) = splice(&before, &after) else {
                continue;
            };
            if touches_anchor(&before, span, replacement) {
                return Err(Error::AnchoredParagraph);
            }
            let mut edit = package.edit_section_text(self.position)?;
            edit.replace(span, replacement)?;
            let commit = edit.commit()?;
            steps.push(commit.patch().clone());
            package = commit.into_package();
        }
        let transactions = steps.len();
        Ok(Commit {
            package,
            patch: Patch { steps },
            diagnostics: Diagnostics {
                changed: transactions != 0,
                transactions,
            },
        })
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
    use super::{Error, splice, touches_anchor, with_text};

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
            Err(Error::InterleavedAnchors)
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
