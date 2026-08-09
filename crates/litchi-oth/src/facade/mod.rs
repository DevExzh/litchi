//! Concise family entry points.

use litchi_core::{Error, HistoryLimits, Metadata, PatchError, Position, Result};
use std::fmt;
use std::{path::Path, sync::Arc};

pub use crate::authoring::Builder;

const MAX_PARAGRAPH_BYTES: usize = 16 * 1024 * 1024;
const MAX_DURABLE_PATCH_BYTES: usize = 512 * 1024 * 1024;
const PATCH_MAGIC: &[u8; 8] = b"LOTHP001";

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

    /// Returns the exact `meta.xml` source, if present.
    #[must_use]
    pub fn meta_xml(&self) -> Option<&str> {
        self.package.meta_xml()
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

    /// Inventories inert active-content surfaces and enforces an explicit policy.
    ///
    /// # Errors
    ///
    /// Returns an error if package enumeration fails or a denied surface is present.
    pub fn check_security(&self, policy: SecurityPolicy) -> Result<SecurityReport> {
        let files = self.files()?;
        let report = SecurityReport {
            embedded_objects: self
                .package
                .resources()
                .iter()
                .filter(|resource| resource.is_embedded())
                .count(),
            external_resources: self
                .package
                .resources()
                .iter()
                .filter(|resource| !resource.is_embedded())
                .count(),
            forms: self.package.forms().len(),
            scripts: files
                .iter()
                .filter(|path| path.starts_with("Basic/") || path.starts_with("Scripts/"))
                .count(),
            signed: files.iter().any(|path| {
                matches!(
                    path.as_str(),
                    "META-INF/documentsignatures.xml" | "META-INF/macrosignatures.xml"
                )
            }),
        };
        if (!policy.allow_embedded_objects && report.embedded_objects > 0)
            || (!policy.allow_external_resources && report.external_resources > 0)
            || (!policy.allow_forms && report.forms > 0)
            || (!policy.allow_scripts && report.scripts > 0)
            || (!policy.allow_signatures && report.signed)
        {
            return Err(Error::InvalidFormat(
                "OTH package is refused by the active-content security policy".to_string(),
            ));
        }
        Ok(report)
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
            heading_changes: Vec::new(),
            list_changes: Vec::new(),
            metadata_xml: None,
            source: self,
            styles_xml: None,
        }
    }

    /// Consumes the snapshot and returns the raw package bytes.
    #[must_use]
    pub fn into_bytes(self) -> Vec<u8> {
        self.package.into_bytes()
    }
}

/// Explicit policy for inert active-content surfaces.
#[derive(Clone, Copy, Debug, Default, PartialEq, Eq)]
#[expect(
    clippy::struct_excessive_bools,
    reason = "the policy deliberately exposes one independent decision per inert surface"
)]
pub struct SecurityPolicy {
    /// Permit package-relative image and object references.
    pub allow_embedded_objects: bool,
    /// Permit external resource references without resolving them.
    pub allow_external_resources: bool,
    /// Permit inert forms and controls.
    pub allow_forms: bool,
    /// Permit script package members without executing them.
    pub allow_scripts: bool,
    /// Permit signed snapshots for read-only inspection.
    pub allow_signatures: bool,
}

/// Counts of inert security-relevant package surfaces.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub struct SecurityReport {
    /// Package-relative image and object references.
    pub embedded_objects: usize,
    /// External resource references.
    pub external_resources: usize,
    /// Forms.
    pub forms: usize,
    /// Script package members.
    pub scripts: usize,
    /// Whether a recognized signature member exists.
    pub signed: bool,
}

/// A source-bound web-template text transaction.
pub struct Edit<'a> {
    appended: Vec<crate::ContentBlock>,
    changes: Vec<ParagraphChange>,
    heading_changes: Vec<HeadingChange>,
    list_changes: Vec<ListChange>,
    metadata_xml: Option<String>,
    source: &'a Template,
    styles_xml: Option<String>,
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
        if self.list_changes.iter().any(|change| {
            change
                .before
                .as_ref()
                .is_some_and(|list| list_contains_paragraph(list, paragraph))
        }) {
            return Err(Error::InvalidFormat(
                "OTH paragraph edit overlaps a staged list structural edit".to_string(),
            ));
        }
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

    /// Replaces one heading's sole direct character-data XML span.
    ///
    /// # Errors
    ///
    /// Returns an error for an invalid selector, oversized value, or heading
    /// whose nested markup cannot be rewritten losslessly.
    pub fn set_heading_text(&mut self, heading: Position, text: impl Into<String>) -> Result<()> {
        let after = text.into();
        if after.len() > MAX_PARAGRAPH_BYTES {
            return Err(Error::InvalidFormat(
                "OTH replacement heading text exceeds the limit".to_string(),
            ));
        }
        let before = self
            .source
            .package
            .headings()
            .get(heading.get())
            .ok_or_else(|| {
                Error::InvalidFormat("OTH heading selector is out of bounds".to_string())
            })?
            .text()
            .to_owned();
        if before == after {
            self.heading_changes
                .retain(|change| change.heading != heading);
            return Ok(());
        }
        if self
            .source
            .package
            .heading_replacement_site(heading.get())
            .is_none()
        {
            return Err(Error::InvalidFormat(
                "OTH heading is not one losslessly replaceable XML text span".to_string(),
            ));
        }
        if let Some(change) = self
            .heading_changes
            .iter_mut()
            .find(|change| change.heading == heading)
        {
            change.after = after;
        } else {
            self.heading_changes
                .try_reserve(1)
                .map_err(|source| Error::Allocation {
                    resource: "OTH staged heading changes",
                    source,
                })?;
            self.heading_changes.push(HeadingChange {
                heading,
                before,
                after,
            });
        }
        Ok(())
    }

    /// Replaces one isolated list with a detached typed list.
    ///
    /// # Errors
    ///
    /// Returns an error for an invalid selector or an overlapping nested list.
    pub fn set_list(&mut self, list: Position, replacement: crate::list::List) -> Result<()> {
        self.stage_list(list, Some(replacement))
    }

    /// Removes one isolated list.
    ///
    /// # Errors
    ///
    /// Returns an error for an invalid selector or an overlapping nested list.
    pub fn remove_list(&mut self, list: Position) -> Result<()> {
        self.stage_list(list, None)
    }

    /// Replaces document metadata through a compact typed `meta.xml` projection.
    ///
    /// # Errors
    ///
    /// Returns an error if the metadata cannot be rendered within XML limits.
    pub fn set_metadata(&mut self, metadata: &Metadata) -> Result<()> {
        self.metadata_xml = Some(crate::authoring::render_metadata(metadata)?);
        Ok(())
    }

    /// Replaces the named common-style catalog.
    ///
    /// # Errors
    ///
    /// Returns an error if the style catalog cannot be rendered compactly.
    pub fn set_styles(&mut self, styles: &[crate::style::Style]) -> Result<()> {
        self.styles_xml = Some(crate::authoring::render_styles(styles)?);
        Ok(())
    }

    fn stage_list(&mut self, list: Position, after: Option<crate::list::List>) -> Result<()> {
        let before = self
            .source
            .package
            .lists()
            .get(list.get())
            .ok_or_else(|| Error::InvalidFormat("OTH list selector is out of bounds".to_string()))?
            .clone();
        let site = self.source.package.list_site(list.get()).ok_or_else(|| {
            Error::InvalidFormat("OTH list structural site is missing".to_string())
        })?;
        let overlaps =
            self.source
                .package
                .list_sites()
                .iter()
                .enumerate()
                .any(|(index, candidate)| {
                    index != list.get()
                        && candidate.range.start >= site.range.start
                        && candidate.range.end <= site.range.end
                });
        if overlaps || before.level() > 1 {
            return Err(Error::InvalidFormat(
                "OTH nested list structural edits are refused".to_string(),
            ));
        }
        if self
            .changes
            .iter()
            .any(|change| list_contains_paragraph(&before, change.paragraph))
        {
            return Err(Error::InvalidFormat(
                "OTH list structural edit overlaps a staged paragraph edit".to_string(),
            ));
        }
        if after.as_ref() == Some(&before) {
            self.list_changes.retain(|change| change.list != list);
            return Ok(());
        }
        if let Some(change) = self
            .list_changes
            .iter_mut()
            .find(|change| change.list == list)
        {
            change.after = after;
        } else {
            self.list_changes
                .try_reserve(1)
                .map_err(|source| Error::Allocation {
                    resource: "OTH staged list changes",
                    source,
                })?;
            self.list_changes.push(ListChange {
                after,
                before: Some(before),
                list,
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
        if self.changes.is_empty()
            && self.heading_changes.is_empty()
            && self.list_changes.is_empty()
            && self.metadata_xml.is_none()
            && self.styles_xml.is_none()
            && self.appended.is_empty()
        {
            return Ok(Commit::unchanged(self.source.clone()));
        }
        let content = crate::codec::compact_for_publication(&replace_texts(
            self.source,
            &self.changes,
            &self.heading_changes,
            &self.list_changes,
            &self.appended,
            None,
        )?)?;
        let snapshot = Template {
            package: self.source.package.rebuild_with_parts(
                &content,
                self.metadata_xml.as_deref(),
                self.styles_xml.as_deref(),
            )?,
        };
        validate_edit_readback(&self, &snapshot)?;
        Ok(Commit {
            snapshot: snapshot.clone(),
            patch: Patch {
                source: self.source.clone(),
                target: snapshot,
                appended: self.appended,
                changes: self.changes,
                durable_fragment: None,
                heading_changes: self.heading_changes,
                list_changes: self.list_changes,
                metadata_xml: self.metadata_xml,
                styles_xml: self.styles_xml,
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
        } else if self.metadata_xml.is_some() && other.metadata_xml.is_some() {
            Some(JoinFailure::Metadata)
        } else if self.styles_xml.is_some() && other.styles_xml.is_some() {
            Some(JoinFailure::Styles)
        } else {
            self.changes
                .iter()
                .find_map(|accepted| {
                    other
                        .changes
                        .iter()
                        .any(|incoming| incoming.paragraph == accepted.paragraph)
                        .then_some(JoinFailure::Paragraph(accepted.paragraph))
                })
                .or_else(|| {
                    self.heading_changes.iter().find_map(|accepted| {
                        other
                            .heading_changes
                            .iter()
                            .any(|incoming| incoming.heading == accepted.heading)
                            .then_some(JoinFailure::Heading(accepted.heading))
                    })
                })
                .or_else(|| {
                    self.list_changes.iter().find_map(|accepted| {
                        other
                            .list_changes
                            .iter()
                            .any(|incoming| incoming.list == accepted.list)
                            .then_some(JoinFailure::List(accepted.list))
                    })
                })
                .or_else(|| {
                    self.list_changes.iter().find_map(|list_change| {
                        let list = list_change.before.as_ref()?;
                        other.changes.iter().find_map(|paragraph_change| {
                            list_contains_paragraph(list, paragraph_change.paragraph).then_some(
                                JoinFailure::ListParagraph {
                                    list: list_change.list,
                                    paragraph: paragraph_change.paragraph,
                                },
                            )
                        })
                    })
                })
                .or_else(|| {
                    other.list_changes.iter().find_map(|list_change| {
                        let list = list_change.before.as_ref()?;
                        self.changes.iter().find_map(|paragraph_change| {
                            list_contains_paragraph(list, paragraph_change.paragraph).then_some(
                                JoinFailure::ListParagraph {
                                    list: list_change.list,
                                    paragraph: paragraph_change.paragraph,
                                },
                            )
                        })
                    })
                })
        };
        if let Some(reason) = failure {
            return Err(JoinError {
                failure: reason,
                rejected: Box::new(other),
            });
        }
        self.changes.extend(other.changes);
        self.heading_changes.extend(other.heading_changes);
        self.list_changes.extend(other.list_changes);
        self.appended.extend(other.appended);
        if self.metadata_xml.is_none() {
            self.metadata_xml = other.metadata_xml;
        }
        if self.styles_xml.is_none() {
            self.styles_xml = other.styles_xml;
        }
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
    /// Both transactions replace the same heading.
    Heading(Position),
    /// Both transactions structurally edit the same list.
    List(Position),
    /// A paragraph edit overlaps its containing list structural edit.
    ListParagraph {
        /// List selector.
        list: Position,
        /// Paragraph selector.
        paragraph: Position,
    },
    /// Both transactions replace metadata.
    Metadata,
    /// Both transactions replace the style catalog.
    Styles,
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

/// One reversible semantic heading-text operation.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct HeadingChange {
    heading: Position,
    before: String,
    after: String,
}

impl HeadingChange {
    /// Zero-based heading position.
    #[must_use]
    pub const fn heading(&self) -> Position {
        self.heading
    }

    /// Expected source text.
    #[must_use]
    pub fn before(&self) -> &str {
        &self.before
    }

    /// Replacement text.
    #[must_use]
    pub fn after(&self) -> &str {
        &self.after
    }
}

/// One reversible typed list replacement or removal.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct ListChange {
    after: Option<crate::list::List>,
    before: Option<crate::list::List>,
    list: Position,
}

impl ListChange {
    /// Zero-based projected list position.
    #[must_use]
    pub const fn list(&self) -> Position {
        self.list
    }

    /// Source list, or `None` for an inverse insertion.
    #[must_use]
    pub const fn before(&self) -> Option<&crate::list::List> {
        self.before.as_ref()
    }

    /// Replacement list, or `None` for removal.
    #[must_use]
    pub const fn after(&self) -> Option<&crate::list::List> {
        self.after.as_ref()
    }
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
                durable_fragment: None,
                heading_changes: Vec::new(),
                list_changes: Vec::new(),
                metadata_xml: None,
                source: snapshot.clone(),
                styles_xml: None,
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
    durable_fragment: Option<String>,
    heading_changes: Vec<HeadingChange>,
    list_changes: Vec<ListChange>,
    metadata_xml: Option<String>,
    source: Template,
    styles_xml: Option<String>,
    target: Template,
}

impl Patch {
    /// Creates a non-mutating deterministic three-way plan.
    ///
    /// # Errors
    ///
    /// Returns an error unless both patches authorize the exact supplied base.
    pub fn plan_three_way(base: &Template, left: &Self, right: &Self) -> Result<MergePlan> {
        if !left.is_applicable_to(base) || !right.is_applicable_to(base) {
            return Err(Error::InvalidFormat(
                "OTH three-way patch base does not match".to_string(),
            ));
        }
        let mut conflicts = Vec::new();
        for change in &left.changes {
            if right
                .changes
                .iter()
                .any(|candidate| candidate.paragraph == change.paragraph)
            {
                conflicts.push(MergeConflict::Paragraph(change.paragraph));
            }
        }
        for change in &left.heading_changes {
            if right
                .heading_changes
                .iter()
                .any(|candidate| candidate.heading == change.heading)
            {
                conflicts.push(MergeConflict::Heading(change.heading));
            }
        }
        for change in &left.list_changes {
            if right
                .list_changes
                .iter()
                .any(|candidate| candidate.list == change.list)
            {
                conflicts.push(MergeConflict::List(change.list));
            }
            if let Some(list) = &change.before {
                for paragraph in &right.changes {
                    if list_contains_paragraph(list, paragraph.paragraph) {
                        conflicts.push(MergeConflict::ListParagraph {
                            list: change.list,
                            paragraph: paragraph.paragraph,
                        });
                    }
                }
            }
        }
        for change in &right.list_changes {
            if let Some(list) = &change.before {
                for paragraph in &left.changes {
                    if list_contains_paragraph(list, paragraph.paragraph) {
                        conflicts.push(MergeConflict::ListParagraph {
                            list: change.list,
                            paragraph: paragraph.paragraph,
                        });
                    }
                }
            }
        }
        if left.has_append() && right.has_append() {
            conflicts.push(MergeConflict::Append);
        }
        if left.metadata_xml.is_some() && right.metadata_xml.is_some() {
            conflicts.push(MergeConflict::Metadata);
        }
        if left.styles_xml.is_some() && right.styles_xml.is_some() {
            conflicts.push(MergeConflict::Styles);
        }
        Ok(MergePlan {
            base: base.clone(),
            conflicts,
            left: left.clone(),
            right: right.clone(),
        })
    }

    fn has_append(&self) -> bool {
        !self.appended.is_empty()
            || self
                .durable_fragment
                .as_deref()
                .is_some_and(|fragment| !fragment.is_empty())
    }
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

    /// Heading-text changes in staging order.
    #[must_use]
    pub fn heading_changes(&self) -> &[HeadingChange] {
        &self.heading_changes
    }

    /// Typed list changes in staging order.
    #[must_use]
    pub fn list_changes(&self) -> &[ListChange] {
        &self.list_changes
    }

    /// Replacement metadata XML retained by this semantic patch.
    #[must_use]
    pub fn metadata_xml(&self) -> Option<&str> {
        self.metadata_xml.as_deref()
    }

    /// Replacement styles XML retained by this semantic patch.
    #[must_use]
    pub fn styles_xml(&self) -> Option<&str> {
        self.styles_xml.as_deref()
    }

    /// Typed blocks appended by the transaction.
    #[must_use]
    pub fn appended(&self) -> &[crate::ContentBlock] {
        &self.appended
    }

    /// Compact appended XML retained by a decoded durable patch.
    ///
    /// In-memory commits expose typed values through [`Self::appended`]; a
    /// durable decode exposes their validated compact representation here.
    #[must_use]
    pub fn durable_appended_xml(&self) -> Option<&str> {
        self.durable_fragment.as_deref()
    }

    /// Serializes a deterministic, exact-source semantic patch envelope.
    ///
    /// The bounded envelope contains source and target packages so application
    /// after process restart remains byte-exact and requires no ambient files.
    ///
    /// # Errors
    ///
    /// Returns an error if a size cannot be represented or the finite envelope
    /// limit would be exceeded.
    pub fn to_bytes(&self) -> Result<Vec<u8>> {
        let appended_xml = match &self.durable_fragment {
            Some(fragment) => fragment.clone(),
            None => crate::authoring::render_fragment(&self.appended)?,
        };
        let mut output = Vec::new();
        output
            .try_reserve_exact(
                self.source
                    .as_bytes()
                    .len()
                    .saturating_add(self.target.as_bytes().len())
                    .saturating_add(appended_xml.len())
                    .saturating_add(128),
            )
            .map_err(|source| Error::Allocation {
                resource: "OTH durable patch",
                source,
            })?;
        output.extend_from_slice(PATCH_MAGIC);
        push_wire_bytes(&mut output, self.source.as_bytes())?;
        push_wire_bytes(&mut output, self.target.as_bytes())?;
        push_wire_bytes(&mut output, appended_xml.as_bytes())?;
        push_wire_usize(&mut output, self.changes.len())?;
        for change in &self.changes {
            push_wire_usize(&mut output, change.paragraph.get())?;
            push_wire_bytes(&mut output, change.before.as_bytes())?;
            push_wire_bytes(&mut output, change.after.as_bytes())?;
        }
        push_wire_usize(&mut output, self.heading_changes.len())?;
        for change in &self.heading_changes {
            push_wire_usize(&mut output, change.heading.get())?;
            push_wire_bytes(&mut output, change.before.as_bytes())?;
            push_wire_bytes(&mut output, change.after.as_bytes())?;
        }
        if output.len() > MAX_DURABLE_PATCH_BYTES {
            return Err(Error::InvalidFormat(
                "OTH durable patch exceeds the byte limit".to_string(),
            ));
        }
        Ok(output)
    }

    /// Decodes and fully reopens a deterministic durable patch envelope.
    ///
    /// # Errors
    ///
    /// Returns an error for malformed, over-limit, stale-semantic, or invalid
    /// embedded source/target packages.
    pub fn from_bytes(bytes: &[u8]) -> Result<Self> {
        if bytes.len() > MAX_DURABLE_PATCH_BYTES || !bytes.starts_with(PATCH_MAGIC) {
            return Err(Error::InvalidFormat(
                "invalid or oversized OTH durable patch".to_string(),
            ));
        }
        let mut cursor = PATCH_MAGIC.len();
        let source = Template::from_bytes(read_wire_bytes(bytes, &mut cursor)?.to_vec())?;
        let target = Template::from_bytes(read_wire_bytes(bytes, &mut cursor)?.to_vec())?;
        let appended_xml = read_wire_string(bytes, &mut cursor)?;
        if !appended_xml.is_empty() {
            let wrapped = format!(
                "<?xml version=\"1.0\" encoding=\"UTF-8\"?><office:document-content xmlns:office=\"urn:oasis:names:tc:opendocument:xmlns:office:1.0\" xmlns:text=\"urn:oasis:names:tc:opendocument:xmlns:text:1.0\"><office:body><office:text>{appended_xml}</office:text></office:body></office:document-content>"
            );
            crate::codec::validate_authored(&wrapped)?;
        }
        let changes = read_paragraph_changes(bytes, &mut cursor, &source, &target)?;
        let heading_changes = read_heading_changes(bytes, &mut cursor, &source, &target)?;
        if cursor != bytes.len() {
            return Err(Error::InvalidFormat(
                "OTH durable patch has trailing bytes".to_string(),
            ));
        }
        if !appended_xml.is_empty() && !target.content_xml().contains(&appended_xml) {
            return Err(Error::InvalidFormat(
                "OTH durable patch append failed target readback".to_string(),
            ));
        }
        Ok(Self {
            appended: Vec::new(),
            changes,
            durable_fragment: (!appended_xml.is_empty()).then_some(appended_xml),
            heading_changes,
            list_changes: Vec::new(),
            metadata_xml: (source.meta_xml() != target.meta_xml())
                .then(|| target.meta_xml().unwrap_or_default().to_owned()),
            styles_xml: (source.styles_xml() != target.styles_xml())
                .then(|| target.styles_xml().unwrap_or_default().to_owned()),
            source,
            target,
        })
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
            durable_fragment: None,
            heading_changes: self
                .heading_changes
                .iter()
                .map(|change| HeadingChange {
                    heading: change.heading,
                    before: change.after.clone(),
                    after: change.before.clone(),
                })
                .collect(),
            list_changes: self
                .list_changes
                .iter()
                .map(|change| ListChange {
                    after: change.before.clone(),
                    before: change.after.clone(),
                    list: change.list,
                })
                .collect(),
            metadata_xml: self.source.meta_xml().map(str::to_owned),
            source: self.target.clone(),
            styles_xml: self.source.styles_xml().map(str::to_owned),
            target: self.source.clone(),
        }
    }
}

/// Deterministic semantic three-way conflict.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
#[non_exhaustive]
pub enum MergeConflict {
    /// Both patches edit one paragraph.
    Paragraph(Position),
    /// Both patches edit one heading.
    Heading(Position),
    /// Both patches structurally edit one list.
    List(Position),
    /// A paragraph edit overlaps its containing list structural edit.
    ListParagraph {
        /// List selector.
        list: Position,
        /// Paragraph selector.
        paragraph: Position,
    },
    /// Both patches replace metadata.
    Metadata,
    /// Both patches replace styles.
    Styles,
    /// Both patches append at the structural tail.
    Append,
}

/// Non-mutating three-way composition plan.
pub struct MergePlan {
    base: Template,
    conflicts: Vec<MergeConflict>,
    left: Patch,
    right: Patch,
}

impl MergePlan {
    /// Deterministically ordered conflicts.
    #[must_use]
    pub fn conflicts(&self) -> &[MergeConflict] {
        &self.conflicts
    }

    /// Publishes the disjoint plan after a complete rebuild and reopen.
    ///
    /// # Errors
    ///
    /// Returns an error while conflicts remain or publication/readback fails.
    pub fn publish(&self) -> Result<Template> {
        if !self.conflicts.is_empty() {
            return Err(Error::InvalidFormat(
                "OTH three-way plan has unresolved conflicts".to_string(),
            ));
        }
        let mut paragraph_changes = self.left.changes.clone();
        paragraph_changes.extend(self.right.changes.clone());
        let mut heading_changes = self.left.heading_changes.clone();
        heading_changes.extend(self.right.heading_changes.clone());
        let mut list_changes = self.left.list_changes.clone();
        list_changes.extend(self.right.list_changes.clone());
        let mut appended = self.left.appended.clone();
        appended.extend(self.right.appended.clone());
        let durable_fragment = self
            .left
            .durable_fragment
            .as_deref()
            .or(self.right.durable_fragment.as_deref());
        let content = crate::codec::compact_for_publication(&replace_texts(
            &self.base,
            &paragraph_changes,
            &heading_changes,
            &list_changes,
            &appended,
            durable_fragment,
        )?)?;
        let metadata_xml = self
            .left
            .metadata_xml
            .as_deref()
            .or(self.right.metadata_xml.as_deref());
        let styles_xml = self
            .left
            .styles_xml
            .as_deref()
            .or(self.right.styles_xml.as_deref());
        let candidate = Template {
            package: self
                .base
                .package
                .rebuild_with_parts(&content, metadata_xml, styles_xml)?,
        };
        for change in paragraph_changes {
            if candidate
                .package
                .paragraphs()
                .get(change.paragraph.get())
                .map(crate::paragraph::Paragraph::text)
                != Some(change.after.as_str())
            {
                return Err(Error::InvalidFormat(
                    "OTH merged paragraph failed readback".to_string(),
                ));
            }
        }
        for change in heading_changes {
            if candidate
                .package
                .headings()
                .get(change.heading.get())
                .map(crate::heading::Heading::text)
                != Some(change.after.as_str())
            {
                return Err(Error::InvalidFormat(
                    "OTH merged heading failed readback".to_string(),
                ));
            }
        }
        for change in &list_changes {
            let Some(expected) = change.after.as_ref() else {
                continue;
            };
            let target_index = list_target_index(&list_changes, change)?;
            let actual = candidate.package.lists().get(target_index).ok_or_else(|| {
                Error::InvalidFormat("OTH merged replacement list disappeared".to_string())
            })?;
            if !lists_semantically_equal(actual, expected) {
                return Err(Error::InvalidFormat(
                    "OTH merged replacement list failed semantic readback".to_string(),
                ));
            }
        }
        Ok(candidate)
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

fn list_paragraph_count(list: &crate::list::List) -> usize {
    list.items()
        .iter()
        .map(|item| item.paragraphs().len())
        .fold(0_usize, usize::saturating_add)
}

fn validate_edit_readback(edit: &Edit<'_>, snapshot: &Template) -> Result<()> {
    for change in &edit.changes {
        let actual = snapshot
            .package
            .paragraphs()
            .get(change.paragraph.get())
            .ok_or_else(|| Error::InvalidFormat("OTH edited paragraph disappeared".to_string()))?;
        if actual.text() != change.after {
            return Err(Error::InvalidFormat(
                "OTH package paragraph edit failed semantic readback".to_string(),
            ));
        }
    }
    for change in &edit.heading_changes {
        let actual = snapshot
            .package
            .headings()
            .get(change.heading.get())
            .ok_or_else(|| Error::InvalidFormat("OTH edited heading disappeared".to_string()))?;
        if actual.text() != change.after {
            return Err(Error::InvalidFormat(
                "OTH package heading edit failed semantic readback".to_string(),
            ));
        }
    }
    let removed_lists = edit
        .list_changes
        .iter()
        .filter(|change| change.after.is_none())
        .count();
    let appended_lists = edit
        .appended
        .iter()
        .filter(|block| matches!(block, crate::ContentBlock::List(_)))
        .count();
    if snapshot.package.lists().len()
        != edit
            .source
            .package
            .lists()
            .len()
            .saturating_sub(removed_lists)
            .saturating_add(appended_lists)
    {
        return Err(Error::InvalidFormat(
            "OTH list edit failed structural readback".to_string(),
        ));
    }
    for change in &edit.list_changes {
        let Some(expected) = change.after.as_ref() else {
            continue;
        };
        let target_index = list_target_index(&edit.list_changes, change)?;
        let actual =
            snapshot.package.lists().get(target_index).ok_or_else(|| {
                Error::InvalidFormat("OTH replacement list disappeared".to_string())
            })?;
        if !lists_semantically_equal(actual, expected) {
            return Err(Error::InvalidFormat(
                "OTH replacement list failed semantic readback".to_string(),
            ));
        }
    }
    if let Some(expected) = edit.metadata_xml.as_deref()
        && snapshot.package.meta_xml() != Some(expected)
    {
        return Err(Error::InvalidFormat(
            "OTH metadata replacement failed exact readback".to_string(),
        ));
    }
    if let Some(expected) = edit.styles_xml.as_deref()
        && snapshot.package.styles_xml() != Some(expected)
    {
        return Err(Error::InvalidFormat(
            "OTH styles replacement failed exact readback".to_string(),
        ));
    }
    let appended_block_count = edit.appended.iter().fold(0_usize, |count, block| {
        count.saturating_add(match block {
            crate::ContentBlock::Heading(_) | crate::ContentBlock::Paragraph(_) => 1,
            crate::ContentBlock::List(list) => list_paragraph_count(list),
        })
    });
    let replaced_block_count = edit
        .list_changes
        .iter()
        .filter_map(|change| change.before.as_ref())
        .map(list_paragraph_count)
        .fold(0_usize, usize::saturating_add);
    let replacement_block_count = edit
        .list_changes
        .iter()
        .filter_map(|change| change.after.as_ref())
        .map(list_paragraph_count)
        .fold(0_usize, usize::saturating_add);
    let expected_order_len = edit
        .source
        .package
        .order()
        .len()
        .saturating_sub(replaced_block_count)
        .saturating_add(replacement_block_count)
        .saturating_add(appended_block_count);
    if snapshot.package.order().len() != expected_order_len {
        return Err(Error::InvalidFormat(
            "OTH structural edit failed block-order readback".to_string(),
        ));
    }
    Ok(())
}

fn list_target_index(changes: &[ListChange], change: &ListChange) -> Result<usize> {
    let removed_before = changes
        .iter()
        .filter(|candidate| candidate.after.is_none() && candidate.list.get() < change.list.get())
        .count();
    change
        .list
        .get()
        .checked_sub(removed_before)
        .ok_or_else(|| {
            Error::InvalidFormat("OTH replacement list target position underflow".to_string())
        })
}

fn lists_semantically_equal(actual: &crate::list::List, expected: &crate::list::List) -> bool {
    actual.level() == expected.level()
        && actual.style_name() == expected.style_name()
        && actual.items().len() == expected.items().len()
        && actual
            .items()
            .iter()
            .zip(expected.items())
            .all(|(actual_item, expected_item)| {
                actual_item.start_value() == expected_item.start_value()
                    && actual_item.paragraphs() == expected_item.paragraphs()
            })
}

fn list_contains_paragraph(list: &crate::list::List, paragraph: Position) -> bool {
    list.items()
        .iter()
        .flat_map(crate::list::Item::paragraph_positions)
        .any(|position| *position == paragraph)
}

fn push_wire_usize(output: &mut Vec<u8>, value: usize) -> Result<()> {
    let wire_value = u64::try_from(value)
        .map_err(|error| Error::InvalidFormat(format!("OTH patch value is too large: {error}")))?;
    output.extend_from_slice(&wire_value.to_le_bytes());
    Ok(())
}

fn push_wire_bytes(output: &mut Vec<u8>, value: &[u8]) -> Result<()> {
    push_wire_usize(output, value.len())?;
    if output.len().saturating_add(value.len()) > MAX_DURABLE_PATCH_BYTES {
        return Err(Error::InvalidFormat(
            "OTH durable patch exceeds the byte limit".to_string(),
        ));
    }
    output.extend_from_slice(value);
    Ok(())
}

fn read_wire_usize(bytes: &[u8], cursor: &mut usize) -> Result<usize> {
    let end = cursor
        .checked_add(8)
        .ok_or_else(|| Error::InvalidFormat("OTH patch cursor overflow".to_string()))?;
    let raw: [u8; 8] = bytes
        .get(*cursor..end)
        .ok_or_else(|| Error::InvalidFormat("truncated OTH durable patch".to_string()))?
        .try_into()
        .map_err(|error| Error::InvalidFormat(format!("invalid OTH patch integer: {error}")))?;
    *cursor = end;
    usize::try_from(u64::from_le_bytes(raw))
        .map_err(|error| Error::InvalidFormat(format!("OTH patch integer overflow: {error}")))
}

fn read_wire_bytes<'a>(bytes: &'a [u8], cursor: &mut usize) -> Result<&'a [u8]> {
    let length = read_wire_usize(bytes, cursor)?;
    let end = cursor
        .checked_add(length)
        .ok_or_else(|| Error::InvalidFormat("OTH patch span overflow".to_string()))?;
    let value = bytes
        .get(*cursor..end)
        .ok_or_else(|| Error::InvalidFormat("truncated OTH durable patch value".to_string()))?;
    *cursor = end;
    Ok(value)
}

fn read_wire_string(bytes: &[u8], cursor: &mut usize) -> Result<String> {
    String::from_utf8(read_wire_bytes(bytes, cursor)?.to_vec())
        .map_err(|error| Error::InvalidFormat(format!("invalid OTH patch UTF-8: {error}")))
}

fn read_paragraph_changes(
    bytes: &[u8],
    cursor: &mut usize,
    source: &Template,
    target: &Template,
) -> Result<Vec<ParagraphChange>> {
    let count = read_wire_usize(bytes, cursor)?;
    if count > source.package.paragraphs().len() || count > target.package.paragraphs().len() {
        return Err(Error::InvalidFormat(
            "OTH durable patch paragraph count is invalid".to_string(),
        ));
    }
    let mut changes = Vec::new();
    changes
        .try_reserve_exact(count)
        .map_err(|allocation_error| Error::Allocation {
            resource: "OTH durable paragraph changes",
            source: allocation_error,
        })?;
    for _ in 0..count {
        let paragraph = Position::new(read_wire_usize(bytes, cursor)?);
        let before = read_wire_string(bytes, cursor)?;
        let after = read_wire_string(bytes, cursor)?;
        let source_text = source
            .package
            .paragraphs()
            .get(paragraph.get())
            .map(crate::paragraph::Paragraph::text);
        let target_text = target
            .package
            .paragraphs()
            .get(paragraph.get())
            .map(crate::paragraph::Paragraph::text);
        if source_text != Some(before.as_str()) || target_text != Some(after.as_str()) {
            return Err(Error::InvalidFormat(
                "OTH durable paragraph change failed semantic readback".to_string(),
            ));
        }
        changes.push(ParagraphChange {
            paragraph,
            before,
            after,
        });
    }
    Ok(changes)
}

fn read_heading_changes(
    bytes: &[u8],
    cursor: &mut usize,
    source: &Template,
    target: &Template,
) -> Result<Vec<HeadingChange>> {
    let count = read_wire_usize(bytes, cursor)?;
    if count > source.package.headings().len() || count > target.package.headings().len() {
        return Err(Error::InvalidFormat(
            "OTH durable patch heading count is invalid".to_string(),
        ));
    }
    let mut changes = Vec::new();
    changes
        .try_reserve_exact(count)
        .map_err(|allocation_error| Error::Allocation {
            resource: "OTH durable heading changes",
            source: allocation_error,
        })?;
    for _ in 0..count {
        let heading = Position::new(read_wire_usize(bytes, cursor)?);
        let before = read_wire_string(bytes, cursor)?;
        let after = read_wire_string(bytes, cursor)?;
        let source_text = source
            .package
            .headings()
            .get(heading.get())
            .map(crate::heading::Heading::text);
        let target_text = target
            .package
            .headings()
            .get(heading.get())
            .map(crate::heading::Heading::text);
        if source_text != Some(before.as_str()) || target_text != Some(after.as_str()) {
            return Err(Error::InvalidFormat(
                "OTH durable heading change failed semantic readback".to_string(),
            ));
        }
        changes.push(HeadingChange {
            heading,
            before,
            after,
        });
    }
    Ok(changes)
}

fn replace_texts(
    source: &Template,
    changes: &[ParagraphChange],
    heading_changes: &[HeadingChange],
    list_changes: &[ListChange],
    appended: &[crate::ContentBlock],
    durable_fragment: Option<&str>,
) -> Result<String> {
    if !appended.is_empty() && durable_fragment.is_some() {
        return Err(Error::InvalidFormat(
            "OTH edit cannot combine typed and durable appended fragments".to_string(),
        ));
    }
    let has_append =
        !appended.is_empty() || durable_fragment.is_some_and(|value| !value.is_empty());
    let mut replacements = Vec::new();
    replacements
        .try_reserve_exact(
            changes
                .len()
                .saturating_add(heading_changes.len())
                .saturating_add(list_changes.len())
                .saturating_add(usize::from(has_append)),
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
    for change in heading_changes {
        let site = source
            .package
            .heading_replacement_site(change.heading.get())
            .ok_or_else(|| Error::InvalidFormat("OTH heading edit site disappeared".to_string()))?;
        replacements.push((site.clone(), quick_xml::escape::escape(&change.after)));
    }
    for change in list_changes {
        if change.before.is_none() {
            return Err(Error::InvalidFormat(
                "OTH inverse list insertion cannot be recomposed".to_string(),
            ));
        }
        let site = source
            .package
            .list_site(change.list.get())
            .ok_or_else(|| Error::InvalidFormat("OTH list edit site disappeared".to_string()))?;
        let replacement = match &change.after {
            Some(list) => {
                crate::authoring::render_fragment(&[crate::ContentBlock::List(list.clone())])?
            },
            None => String::new(),
        };
        replacements.push((site.clone(), std::borrow::Cow::Owned(replacement)));
    }
    if has_append {
        let fragment = match durable_fragment {
            Some(fragment) => fragment.to_owned(),
            None => crate::authoring::render_fragment(appended)?,
        };
        replacements.push((
            crate::codec::ReplacementSite {
                prefix: String::new(),
                range: source.package.text_close()..source.package.text_close(),
                suffix: String::new(),
            },
            std::borrow::Cow::Owned(fragment),
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
