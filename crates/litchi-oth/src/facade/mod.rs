//! Concise family entry points.

use litchi_core::{Error, HistoryLimits, Metadata, PatchError, Position, Result};
use std::fmt;
use std::{path::Path, sync::Arc};

pub use crate::authoring::Builder;

const MAX_PARAGRAPH_BYTES: usize = 16 * 1024 * 1024;

/// A read-only semantic text-web body projection.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct TextBody {
    bookmarks: Vec<crate::bookmark::Bookmark>,
    forms: Vec<crate::form::Form>,
    headings: Vec<crate::heading::Heading>,
    lists: Vec<crate::list::List>,
    order: Vec<crate::codec::BlockOrder>,
    paragraphs: Vec<crate::paragraph::Paragraph>,
    resources: Vec<crate::resource::Resource>,
}

impl TextBody {
    /// Bookmarks in source-close order.
    #[must_use]
    pub fn bookmarks(&self) -> &[crate::bookmark::Bookmark] {
        &self.bookmarks
    }

    /// Lists in source-close order. Nested lists carry their explicit level.
    #[must_use]
    pub fn lists(&self) -> &[crate::list::List] {
        &self.lists
    }

    /// Inert image and object references.
    #[must_use]
    pub fn resources(&self) -> &[crate::resource::Resource] {
        &self.resources
    }

    /// Inert forms and their controls.
    #[must_use]
    pub fn forms(&self) -> &[crate::form::Form] {
        &self.forms
    }
    /// Iterates paragraphs and headings in source document order.
    #[must_use]
    pub fn blocks(&self) -> impl ExactSizeIterator<Item = Block<'_>> + '_ {
        self.order.iter().map(|block| match *block {
            crate::codec::BlockOrder::Heading(index) => Block::Heading(&self.headings[index]),
            crate::codec::BlockOrder::Paragraph(index) => Block::Paragraph(&self.paragraphs[index]),
        })
    }

    /// Returns projected headings in source order among headings.
    #[must_use]
    pub fn headings(&self) -> &[crate::heading::Heading] {
        &self.headings
    }

    /// Returns projected paragraph character data in document order.
    #[must_use]
    pub fn paragraphs(&self) -> &[crate::paragraph::Paragraph] {
        &self.paragraphs
    }
}

/// A borrowed semantic text block.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
#[non_exhaustive]
pub enum Block<'a> {
    /// A `text:h` heading.
    Heading(&'a crate::heading::Heading),
    /// A `text:p` paragraph.
    Paragraph(&'a crate::paragraph::Paragraph),
}

impl Block<'_> {
    /// Returns the block's projected character data.
    #[must_use]
    pub fn text(&self) -> &str {
        match self {
            Self::Heading(heading) => heading.text(),
            Self::Paragraph(paragraph) => paragraph.text(),
        }
    }
}

/// Immutable document snapshot.
#[derive(Clone)]
pub struct Template {
    package: crate::package::Snapshot,
}

impl Template {
    /// Opens a web-template package from a file path.
    ///
    /// # Errors
    ///
    /// Returns an error if the file cannot be read or is not a valid package.
    pub fn open(path: impl AsRef<Path>) -> Result<Self> {
        crate::package::Snapshot::open(path).map(|package| Self { package })
    }

    /// Opens a web-template package from in-memory bytes.
    ///
    /// # Errors
    ///
    /// Returns an error if the bytes are not a valid package.
    pub fn from_bytes(bytes: Vec<u8>) -> Result<Self> {
        crate::package::Snapshot::from_bytes(bytes).map(|package| Self { package })
    }

    /// Opens a web-template package from shared in-memory bytes without
    /// copying the archive buffer.
    ///
    /// # Errors
    ///
    /// Returns an error if the bytes are not a valid package.
    pub fn from_shared_bytes(bytes: Arc<Vec<u8>>) -> Result<Self> {
        crate::package::Snapshot::from_shared_bytes(bytes).map(|package| Self { package })
    }

    /// Returns the `content.xml` document.
    #[must_use]
    pub fn content_xml(&self) -> &str {
        self.package.content_xml()
    }

    /// Returns the `styles.xml` document, if present.
    #[must_use]
    pub fn styles_xml(&self) -> Option<&str> {
        self.package.styles_xml()
    }

    /// Returns the document metadata, if present.
    #[must_use]
    pub fn metadata(&self) -> Option<&Metadata> {
        self.package.metadata()
    }

    /// Returns named style declarations from `content.xml` and `styles.xml`.
    #[must_use]
    pub fn styles(&self) -> &[crate::style::Style] {
        self.package.styles()
    }

    /// Returns the raw package bytes.
    #[must_use]
    pub fn as_bytes(&self) -> &[u8] {
        self.package.as_bytes()
    }

    /// Lists the file entries stored in the package.
    ///
    /// # Errors
    ///
    /// Returns an error if the package entries cannot be enumerated.
    pub fn files(&self) -> Result<Vec<String>> {
        self.package.files()
    }

    /// Projects inert paragraph character data from the validated text body.
    ///
    /// Fields, links, scripts, forms, resources, and embedded objects are not
    /// evaluated, followed, activated, or otherwise executed.
    ///
    /// # Errors
    ///
    /// The current eager snapshot projection is infallible. The result wrapper
    /// is retained so future lazy projections can report bounded read errors
    /// without changing this public entry point.
    pub fn text_body(&self) -> Result<TextBody> {
        Ok(TextBody {
            bookmarks: self.package.bookmarks().to_vec(),
            forms: self.package.forms().to_vec(),
            headings: self.package.headings().to_vec(),
            lists: self.package.lists().to_vec(),
            order: self.package.order().to_vec(),
            paragraphs: self.package.paragraphs().to_vec(),
            resources: self.package.resources().to_vec(),
        })
    }

    /// Starts a source-bound text-body transaction.
    #[must_use]
    pub fn edit(&self) -> Edit<'_> {
        Edit {
            appended: Vec::new(),
            changes: Vec::new(),
            source: self,
        }
    }

    /// Consumes the snapshot and returns the raw package bytes.
    #[must_use]
    pub fn into_bytes(self) -> Vec<u8> {
        self.package.into_bytes()
    }
}

/// A source-bound web-template text transaction.
pub struct Edit<'a> {
    appended: Vec<crate::ContentBlock>,
    changes: Vec<ParagraphChange>,
    source: &'a Template,
}

impl Edit<'_> {
    /// Appends one typed block at the end of `office:text`.
    ///
    /// The block is rendered as compact XML and the complete package is
    /// reopened before publication.
    ///
    /// # Errors
    ///
    /// Returns an allocation error if the bounded staging collection cannot grow.
    pub fn append_block(&mut self, block: impl Into<crate::ContentBlock>) -> Result<()> {
        self.appended
            .try_reserve(1)
            .map_err(|source| Error::Allocation {
                resource: "OTH appended blocks",
                source,
            })?;
        self.appended.push(block.into());
        Ok(())
    }

    /// Replaces one paragraph's sole direct character-data XML span.
    ///
    /// Nested, split, CDATA, and empty paragraphs remain readable but are not
    /// rewritten, because recreating their markup could lose unknown content.
    ///
    /// # Errors
    ///
    /// Returns an error when the selector is invalid or the paragraph has no
    /// lossless replacement span. Multiple distinct paragraphs may be staged
    /// atomically; staging one selector again replaces its pending value.
    pub fn set_paragraph_text(
        &mut self,
        paragraph: Position,
        text: impl Into<String>,
    ) -> Result<()> {
        let after = text.into();
        if after.len() > MAX_PARAGRAPH_BYTES {
            return Err(Error::InvalidFormat(
                "OTH replacement paragraph text exceeds the limit".to_string(),
            ));
        }
        let before = self
            .source
            .package
            .paragraphs()
            .get(paragraph.get())
            .ok_or_else(|| {
                Error::InvalidFormat("OTH paragraph selector is out of bounds".to_string())
            })?
            .text()
            .to_owned();
        if before == after {
            self.changes.retain(|change| change.paragraph != paragraph);
            return Ok(());
        }
        if self
            .source
            .package
            .replacement_site(paragraph.get())
            .is_none()
        {
            return Err(Error::InvalidFormat(
                "OTH paragraph is not one losslessly replaceable XML text span".to_string(),
            ));
        }
        if let Some(change) = self
            .changes
            .iter_mut()
            .find(|change| change.paragraph == paragraph)
        {
            change.after = after;
        } else {
            self.changes
                .try_reserve(1)
                .map_err(|source| Error::Allocation {
                    resource: "OTH staged paragraph changes",
                    source,
                })?;
            self.changes.push(ParagraphChange {
                paragraph,
                before,
                after,
            });
        }
        Ok(())
    }

    /// Atomically validates, publishes, and records this text edit.
    ///
    /// # Errors
    ///
    /// Returns an error if the source cannot be losslessly rewritten or the
    /// fully reopened candidate fails semantic readback.
    pub fn commit(self) -> Result<Commit> {
        if self.changes.is_empty() && self.appended.is_empty() {
            return Ok(Commit::unchanged(self.source.clone()));
        }
        let content = crate::codec::compact_for_publication(&replace_texts(
            self.source,
            &self.changes,
            &self.appended,
        )?)?;
        let snapshot = Template {
            package: self.source.package.rebuild_with_content(&content)?,
        };
        for change in &self.changes {
            let actual = snapshot
                .package
                .paragraphs()
                .get(change.paragraph.get())
                .ok_or_else(|| {
                    Error::InvalidFormat("OTH edited paragraph disappeared".to_string())
                })?;
            if actual.text() != change.after {
                return Err(Error::InvalidFormat(
                    "OTH package edit failed semantic readback".to_string(),
                ));
            }
        }
        let appended_block_count = self.appended.iter().fold(0_usize, |count, block| {
            count.saturating_add(match block {
                crate::ContentBlock::Heading(_) | crate::ContentBlock::Paragraph(_) => 1,
                crate::ContentBlock::List(list) => list
                    .items()
                    .iter()
                    .map(|item| item.paragraphs().len())
                    .fold(0_usize, usize::saturating_add),
            })
        });
        if snapshot.package.order().len()
            != self
                .source
                .package
                .order()
                .len()
                .saturating_add(appended_block_count)
        {
            return Err(Error::InvalidFormat(
                "OTH appended blocks failed semantic readback".to_string(),
            ));
        }
        Ok(Commit {
            snapshot: snapshot.clone(),
            patch: Patch {
                source: self.source.clone(),
                target: snapshot,
                appended: self.appended,
                changes: self.changes,
            },
            changed: true,
        })
    }
}

impl<'a> Edit<'a> {
    /// Joins an independently prepared transaction when its effects are disjoint.
    ///
    /// Two append sequences conflict because their relative order was not
    /// established by either author. On failure this edit is unchanged and the
    /// rejected edit remains recoverable from [`JoinError`].
    ///
    /// # Errors
    ///
    /// Returns the rejected transaction with a typed source or overlap reason.
    pub fn join(&mut self, other: Self) -> std::result::Result<&mut Self, JoinError<'a>> {
        let failure = if !self.source.package.is_same(&other.source.package) {
            Some(JoinFailure::DifferentSnapshot)
        } else if !self.appended.is_empty() && !other.appended.is_empty() {
            Some(JoinFailure::Append)
        } else {
            self.changes.iter().find_map(|accepted| {
                other
                    .changes
                    .iter()
                    .any(|incoming| incoming.paragraph == accepted.paragraph)
                    .then_some(JoinFailure::Paragraph(accepted.paragraph))
            })
        };
        if let Some(reason) = failure {
            return Err(JoinError {
                failure: reason,
                rejected: Box::new(other),
            });
        }
        self.changes.extend(other.changes);
        self.appended.extend(other.appended);
        Ok(self)
    }
}

/// Deterministic edit-composition refusal.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
#[non_exhaustive]
pub enum JoinFailure {
    /// Transactions originated from distinct immutable snapshots.
    DifferentSnapshot,
    /// Both transactions append at the same structural tail.
    Append,
    /// Both transactions replace the same paragraph.
    Paragraph(Position),
}

/// A join refusal that retains the rejected transaction.
pub struct JoinError<'a> {
    failure: JoinFailure,
    rejected: Box<Edit<'a>>,
}

impl fmt::Debug for JoinError<'_> {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("JoinError")
            .field("failure", &self.failure)
            .finish_non_exhaustive()
    }
}

impl<'a> JoinError<'a> {
    /// Structured refusal reason.
    #[must_use]
    pub const fn failure(&self) -> JoinFailure {
        self.failure
    }

    /// Recovers the rejected work.
    #[must_use]
    pub fn into_rejected(self) -> Edit<'a> {
        *self.rejected
    }
}

/// One reversible semantic paragraph-text operation.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct ParagraphChange {
    paragraph: Position,
    before: String,
    after: String,
}

impl ParagraphChange {
    /// The zero-based source-order paragraph position.
    #[must_use]
    pub const fn paragraph(&self) -> Position {
        self.paragraph
    }

    /// The text expected before applying the patch.
    #[must_use]
    pub fn before(&self) -> &str {
        &self.before
    }

    /// The replacement text.
    #[must_use]
    pub fn after(&self) -> &str {
        &self.after
    }
}

/// A committed immutable template and its exact-source patch.
pub struct Commit {
    snapshot: Template,
    patch: Patch,
    changed: bool,
}

impl Commit {
    fn unchanged(snapshot: Template) -> Self {
        Self {
            patch: Patch {
                appended: Vec::new(),
                changes: Vec::new(),
                source: snapshot.clone(),
                target: snapshot.clone(),
            },
            snapshot,
            changed: false,
        }
    }

    /// Whether the committed package differs from its source snapshot.
    #[must_use]
    pub const fn changed(&self) -> bool {
        self.changed
    }

    /// Returns the committed immutable template snapshot.
    #[must_use]
    pub fn template(&self) -> &Template {
        &self.snapshot
    }

    /// Returns the source-checked reversible patch.
    #[must_use]
    pub fn patch(&self) -> &Patch {
        &self.patch
    }

    /// Consumes the commit and returns the published template.
    #[must_use]
    pub fn into_template(self) -> Template {
        self.snapshot
    }
}

/// A source-checked reversible OTH paragraph-text patch.
#[derive(Clone)]
pub struct Patch {
    appended: Vec<crate::ContentBlock>,
    changes: Vec<ParagraphChange>,
    source: Template,
    target: Template,
}

impl Patch {
    /// Returns whether this patch authorizes the supplied exact source bytes.
    #[must_use]
    pub fn is_applicable_to(&self, source: &Template) -> bool {
        self.source.as_bytes() == source.as_bytes()
    }

    /// Applies this patch only to its exact immutable source.
    ///
    /// # Errors
    ///
    /// Returns an error when the supplied source differs byte-for-byte.
    pub fn apply(&self, source: &Template) -> Result<Template> {
        if !self.is_applicable_to(source) {
            return Err(Error::InvalidFormat(
                "OTH patch source does not match its expected snapshot".to_string(),
            ));
        }
        Ok(self.target.clone())
    }

    /// Returns the semantic change, if this is not an exact no-op patch.
    ///
    /// For a multi-paragraph transaction, use [`Self::changes`] to inspect
    /// the complete operation list.
    #[must_use]
    pub fn change(&self) -> Option<&ParagraphChange> {
        self.changes.first()
    }

    /// Returns all semantic changes in staging order.
    #[must_use]
    pub fn changes(&self) -> &[ParagraphChange] {
        &self.changes
    }

    /// Typed blocks appended by the transaction.
    #[must_use]
    pub fn appended(&self) -> &[crate::ContentBlock] {
        &self.appended
    }

    /// Returns the patch that restores the exact source snapshot.
    #[must_use]
    pub fn inverse(&self) -> Self {
        Self {
            appended: Vec::new(),
            changes: self
                .changes
                .iter()
                .map(|change| ParagraphChange {
                    paragraph: change.paragraph,
                    before: change.after.clone(),
                    after: change.before.clone(),
                })
                .collect(),
            source: self.target.clone(),
            target: self.source.clone(),
        }
    }
}

/// Explicit bounded undo/redo history for immutable OTH snapshots.
pub struct History {
    inner: litchi_core::History<Template>,
}

impl History {
    /// Starts history at one immutable template.
    #[must_use]
    pub fn new(current: Template, limits: HistoryLimits) -> Self {
        Self {
            inner: litchi_core::History::new(current, limits),
        }
    }

    /// Current immutable snapshot.
    #[must_use]
    pub const fn current(&self) -> &Template {
        self.inner.current()
    }

    /// Records a commit using its exact target package size as finite weight.
    ///
    /// # Errors
    ///
    /// Returns a history budget error without changing the selected snapshot.
    pub fn record(&mut self, commit: Commit) -> std::result::Result<Vec<Template>, PatchError> {
        let weight = u64::try_from(commit.template().as_bytes().len()).unwrap_or(u64::MAX);
        self.inner.record(commit.into_template(), weight)
    }

    /// Selects the previous retained snapshot.
    pub fn undo(&mut self) -> bool {
        self.inner.undo()
    }

    /// Selects the next retained snapshot.
    pub fn redo(&mut self) -> bool {
        self.inner.redo()
    }

    /// Whether undo is available.
    #[must_use]
    pub fn can_undo(&self) -> bool {
        self.inner.can_undo()
    }

    /// Whether redo is available.
    #[must_use]
    pub fn can_redo(&self) -> bool {
        self.inner.can_redo()
    }
}

fn replace_texts(
    source: &Template,
    changes: &[ParagraphChange],
    appended: &[crate::ContentBlock],
) -> Result<String> {
    let mut replacements = Vec::new();
    replacements
        .try_reserve_exact(
            changes
                .len()
                .saturating_add(usize::from(!appended.is_empty())),
        )
        .map_err(|allocation_error| Error::Allocation {
            resource: "OTH paragraph replacements",
            source: allocation_error,
        })?;
    for change in changes {
        let site = source
            .package
            .replacement_site(change.paragraph.get())
            .ok_or_else(|| {
                Error::InvalidFormat("OTH paragraph edit site disappeared".to_string())
            })?;
        replacements.push((site.clone(), quick_xml::escape::escape(&change.after)));
    }
    if !appended.is_empty() {
        replacements.push((
            crate::codec::ReplacementSite {
                prefix: String::new(),
                range: source.package.text_close()..source.package.text_close(),
                suffix: String::new(),
            },
            std::borrow::Cow::Owned(crate::authoring::render_fragment(appended)?),
        ));
    }
    replacements.sort_unstable_by_key(|(site, _replacement)| site.range.start);

    let input = source.content_xml();
    let mut capacity = input.len();
    let mut previous_end = 0;
    for (site, replacement) in &replacements {
        if site.range.start < previous_end
            || site.range.start > site.range.end
            || site.range.end > input.len()
        {
            return Err(Error::InvalidFormat(
                "OTH paragraph source spans overlap or are invalid".to_string(),
            ));
        }
        previous_end = site.range.end;
        capacity = capacity
            .checked_sub(site.range.end - site.range.start)
            .and_then(|size| size.checked_add(site.prefix.len()))
            .and_then(|size| size.checked_add(replacement.len()))
            .and_then(|size| size.checked_add(site.suffix.len()))
            .ok_or_else(|| Error::InvalidFormat("OTH edited content size overflow".to_string()))?;
    }
    if capacity > MAX_PARAGRAPH_BYTES.saturating_mul(16) {
        return Err(Error::InvalidFormat(
            "OTH edited content exceeds the output limit".to_string(),
        ));
    }
    let mut output = String::new();
    output
        .try_reserve_exact(capacity)
        .map_err(|allocation_error| Error::Allocation {
            resource: "OTH edited content",
            source: allocation_error,
        })?;
    let mut cursor = 0;
    for (site, replacement) in replacements {
        output.push_str(&input[cursor..site.range.start]);
        output.push_str(&site.prefix);
        output.push_str(&replacement);
        output.push_str(&site.suffix);
        cursor = site.range.end;
    }
    output.push_str(&input[cursor..]);
    Ok(output)
}
