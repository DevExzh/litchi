//! Concise family entry points.

use litchi_core::{Error, HistoryLimits, Metadata, PatchError, Position, Result};
use std::fmt;
use std::{path::Path, sync::Arc};

pub use crate::authoring::Builder;

const MAX_PARAGRAPH_BYTES: usize = 16 * 1024 * 1024;
const MAX_DURABLE_PATCH_BYTES: usize = 512 * 1024 * 1024;
const PATCH_MAGIC: &[u8; 8] = b"LOTHP002";

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

    /// Plans a dependency-checked block transfer from another immutable template.
    ///
    /// # Errors
    ///
    /// Returns an error when either package violates the policy, the selector is
    /// invalid, rich inline markup cannot be preserved by structural authoring,
    /// or a required style is missing or collides with the destination catalog.
    pub fn plan_transfer_from(
        &self,
        source: &Self,
        selector: TransferSelector,
        policy: TransferPolicy,
    ) -> Result<TransferPlan> {
        source.check_security(policy.security)?;
        self.check_security(policy.security)?;
        let block = transfer_block(source, selector)?;
        let style_names = block_style_names(&block);
        let imported_styles = resolve_transfer_styles(
            source.styles(),
            self.styles(),
            &style_names,
            policy.include_styles,
        )?;
        Ok(TransferPlan {
            block,
            destination: self.clone(),
            imported_styles,
        })
    }

    /// Starts a source-bound text-body transaction.
    #[must_use]
    pub fn edit(&self) -> Edit<'_> {
        Edit {
            appended: Vec::new(),
            changes: Vec::new(),
            forms_change: None,
            heading_changes: Vec::new(),
            inline_changes: Vec::new(),
            list_changes: Vec::new(),
            metadata: PartChange::Keep,
            source: self,
            styles: PartChange::Keep,
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

/// One source block selected for cross-template transfer.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
#[non_exhaustive]
pub enum TransferSelector {
    /// A heading by source-order heading position.
    Heading(Position),
    /// An isolated list by source-close list position.
    List(Position),
    /// A paragraph by source-order paragraph position.
    Paragraph(Position),
}

/// Explicit dependency and active-content policy for block transfer.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub struct TransferPolicy {
    /// Permit copying missing named style dependencies and their parents.
    pub include_styles: bool,
    /// Policy applied independently to both source and destination packages.
    pub security: SecurityPolicy,
}

impl Default for TransferPolicy {
    fn default() -> Self {
        Self {
            include_styles: true,
            security: SecurityPolicy::default(),
        }
    }
}

/// A validated, non-mutating cross-template publication plan.
pub struct TransferPlan {
    block: crate::ContentBlock,
    destination: Template,
    imported_styles: Vec<crate::style::Style>,
}

impl TransferPlan {
    /// Missing styles selected for import in deterministic dependency order.
    #[must_use]
    pub fn imported_styles(&self) -> &[crate::style::Style] {
        &self.imported_styles
    }

    /// Publishes the transfer as a normal source-checked commit and reopens it.
    ///
    /// # Errors
    ///
    /// Returns an error if style or block publication fails semantic readback.
    pub fn publish(self) -> Result<Commit> {
        let mut edit = self.destination.edit();
        edit.append_block(self.block)?;
        if !self.imported_styles.is_empty() {
            let mut styles = self
                .destination
                .styles()
                .iter()
                .filter(|style| style.origin() == crate::style::Origin::Styles)
                .cloned()
                .collect::<Vec<_>>();
            styles.extend(self.imported_styles);
            edit.set_styles(&styles)?;
        }
        edit.commit()
    }
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub(crate) enum PartChange {
    Keep,
    Remove,
    Set(String),
}

/// A text block selected for a rich inline-content replacement.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
#[non_exhaustive]
pub enum InlineBlock {
    /// A heading by source-order heading position.
    Heading(Position),
    /// A paragraph by source-order paragraph position.
    Paragraph(Position),
}

/// One reversible rich inline-content replacement.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct InlineChange {
    after_text: String,
    after_xml: String,
    before_text: String,
    before_xml: String,
    block: InlineBlock,
}

/// One reversible replacement of the complete inert form catalog.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct FormsChange {
    after: Vec<crate::form::Form>,
    after_xml: String,
    before: Vec<crate::form::Form>,
    before_xml: String,
}

impl FormsChange {
    /// Source forms.
    #[must_use]
    pub fn before(&self) -> &[crate::form::Form] {
        &self.before
    }

    /// Replacement forms.
    #[must_use]
    pub fn after(&self) -> &[crate::form::Form] {
        &self.after
    }
}

impl InlineChange {
    /// Selected block.
    #[must_use]
    pub const fn block(&self) -> InlineBlock {
        self.block
    }

    /// Exact source inline XML.
    #[must_use]
    pub fn before_xml(&self) -> &str {
        &self.before_xml
    }

    /// Compact replacement inline XML.
    #[must_use]
    pub fn after_xml(&self) -> &str {
        &self.after_xml
    }

    /// Projected source text.
    #[must_use]
    pub fn before_text(&self) -> &str {
        &self.before_text
    }

    /// Projected replacement text.
    #[must_use]
    pub fn after_text(&self) -> &str {
        &self.after_text
    }
}

impl PartChange {
    const fn is_keep(&self) -> bool {
        matches!(self, Self::Keep)
    }

    fn replacement(&self) -> Option<&str> {
        match self {
            Self::Set(xml) => Some(xml),
            Self::Keep | Self::Remove => None,
        }
    }

    const fn removes(&self) -> bool {
        matches!(self, Self::Remove)
    }
}

/// A source-bound web-template text transaction.
pub struct Edit<'a> {
    appended: Vec<crate::ContentBlock>,
    changes: Vec<ParagraphChange>,
    forms_change: Option<FormsChange>,
    heading_changes: Vec<HeadingChange>,
    inline_changes: Vec<InlineChange>,
    list_changes: Vec<ListChange>,
    metadata: PartChange,
    source: &'a Template,
    styles: PartChange,
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
        if self
            .inline_changes
            .iter()
            .any(|change| change.block == InlineBlock::Paragraph(paragraph))
        {
            return Err(Error::InvalidFormat(
                "OTH paragraph text edit overlaps a staged rich inline edit".to_string(),
            ));
        }
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
        if self
            .inline_changes
            .iter()
            .any(|change| change.block == InlineBlock::Heading(heading))
        {
            return Err(Error::InvalidFormat(
                "OTH heading text edit overlaps a staged rich inline edit".to_string(),
            ));
        }
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

    /// Replaces a paragraph's complete inline content with typed rich items.
    ///
    /// This is the structural CRUD boundary for common links, formatting
    /// spans, inert fields, and point or range bookmark markers.
    ///
    /// # Errors
    ///
    /// Returns an error for an invalid selector, overlapping staged work, or
    /// inline content that cannot be compactly authored and fully reopened.
    pub fn set_paragraph_inline(
        &mut self,
        paragraph: Position,
        content: &[crate::inline::Content],
    ) -> Result<()> {
        self.stage_inline(InlineBlock::Paragraph(paragraph), content)
    }

    /// Replaces a heading's complete inline content with typed rich items.
    ///
    /// # Errors
    ///
    /// Returns an error for an invalid selector, overlapping staged work, or
    /// inline content that cannot be compactly authored and fully reopened.
    pub fn set_heading_inline(
        &mut self,
        heading: Position,
        content: &[crate::inline::Content],
    ) -> Result<()> {
        self.stage_inline(InlineBlock::Heading(heading), content)
    }

    fn stage_inline(
        &mut self,
        block: InlineBlock,
        content: &[crate::inline::Content],
    ) -> Result<()> {
        let (site, before_text) = match block {
            InlineBlock::Paragraph(position) => {
                if self
                    .changes
                    .iter()
                    .any(|change| change.paragraph == position)
                    || self.list_changes.iter().any(|change| {
                        change
                            .before
                            .as_ref()
                            .is_some_and(|list| list_contains_paragraph(list, position))
                    })
                {
                    return Err(Error::InvalidFormat(
                        "OTH rich paragraph edit overlaps staged work".to_string(),
                    ));
                }
                let value = self
                    .source
                    .package
                    .paragraphs()
                    .get(position.get())
                    .ok_or_else(|| {
                        Error::InvalidFormat(
                            "OTH rich paragraph selector is out of bounds".to_string(),
                        )
                    })?;
                (
                    self.source
                        .package
                        .paragraph_content_site(position.get())
                        .ok_or_else(|| {
                            Error::InvalidFormat(
                                "OTH rich paragraph content site is missing".to_string(),
                            )
                        })?,
                    value.text(),
                )
            },
            InlineBlock::Heading(position) => {
                if self
                    .heading_changes
                    .iter()
                    .any(|change| change.heading == position)
                {
                    return Err(Error::InvalidFormat(
                        "OTH rich heading edit overlaps staged work".to_string(),
                    ));
                }
                let value = self
                    .source
                    .package
                    .headings()
                    .get(position.get())
                    .ok_or_else(|| {
                        Error::InvalidFormat(
                            "OTH rich heading selector is out of bounds".to_string(),
                        )
                    })?;
                (
                    self.source
                        .package
                        .heading_content_site(position.get())
                        .ok_or_else(|| {
                            Error::InvalidFormat(
                                "OTH rich heading content site is missing".to_string(),
                            )
                        })?,
                    value.text(),
                )
            },
        };
        let before_xml = self
            .source
            .content_xml()
            .get(site.range.clone())
            .ok_or_else(|| {
                Error::InvalidFormat("OTH rich inline source span is invalid".to_string())
            })?;
        let (after_xml, after_text) = crate::authoring::render_inline(content)?;
        if before_xml == after_xml {
            self.inline_changes.retain(|change| change.block != block);
            return Ok(());
        }
        let staged = InlineChange {
            after_text,
            after_xml,
            before_text: before_text.to_owned(),
            before_xml: before_xml.to_owned(),
            block,
        };
        if let Some(change) = self
            .inline_changes
            .iter_mut()
            .find(|change| change.block == block)
        {
            *change = staged;
        } else {
            self.inline_changes
                .try_reserve(1)
                .map_err(|source| Error::Allocation {
                    resource: "OTH staged rich inline changes",
                    source,
                })?;
            self.inline_changes.push(staged);
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
        let xml = crate::authoring::render_metadata(metadata)?;
        self.metadata = if self.source.meta_xml() == Some(xml.as_str()) {
            PartChange::Keep
        } else {
            PartChange::Set(xml)
        };
        Ok(())
    }

    /// Removes the optional `meta.xml` package member.
    pub fn remove_metadata(&mut self) {
        self.metadata = if self.source.meta_xml().is_some() {
            PartChange::Remove
        } else {
            PartChange::Keep
        };
    }

    /// Replaces the named common-style catalog.
    ///
    /// # Errors
    ///
    /// Returns an error if the style catalog cannot be rendered compactly.
    pub fn set_styles(&mut self, styles: &[crate::style::Style]) -> Result<()> {
        let xml = crate::authoring::render_styles(styles)?;
        self.styles = if self.source.styles_xml() == Some(xml.as_str()) {
            PartChange::Keep
        } else {
            PartChange::Set(xml)
        };
        Ok(())
    }

    /// Removes the optional `styles.xml` package member.
    pub fn remove_styles(&mut self) {
        self.styles = if self.source.styles_xml().is_some() {
            PartChange::Remove
        } else {
            PartChange::Keep
        };
    }

    /// Replaces the complete inert form/control catalog.
    ///
    /// Passing an empty slice removes `office:forms`; passing a non-empty
    /// slice creates or replaces it. Controls remain inert after publication.
    ///
    /// # Errors
    ///
    /// Returns an error when a control kind is not safely authorable XML.
    pub fn set_forms(&mut self, forms: &[crate::form::Form]) -> Result<()> {
        let before = self.source.package.forms();
        if before == forms {
            self.forms_change = None;
            return Ok(());
        }
        let before_xml = match self.source.package.forms_site() {
            Some(site) => self
                .source
                .content_xml()
                .get(site.range.clone())
                .ok_or_else(|| {
                    Error::InvalidFormat("OTH forms source span is invalid".to_string())
                })?
                .to_owned(),
            None => String::new(),
        };
        self.forms_change = Some(FormsChange {
            after: forms.to_vec(),
            after_xml: crate::authoring::render_forms(forms)?,
            before: before.to_vec(),
            before_xml,
        });
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
        let normalized_after = after.map(|value| {
            crate::list::List::projected(
                value.items().to_vec(),
                before.level(),
                value.style_name().map(str::to_owned),
            )
        });
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
        if overlaps {
            return Err(Error::InvalidFormat(
                "OTH list edit would implicitly discard nested lists".to_string(),
            ));
        }
        if self
            .changes
            .iter()
            .any(|change| list_contains_paragraph(&before, change.paragraph))
            || self.inline_changes.iter().any(|change| {
                matches!(change.block, InlineBlock::Paragraph(position) if list_contains_paragraph(&before, position))
            })
        {
            return Err(Error::InvalidFormat(
                "OTH list structural edit overlaps a staged paragraph edit".to_string(),
            ));
        }
        if normalized_after.as_ref() == Some(&before) {
            self.list_changes.retain(|change| change.list != list);
            return Ok(());
        }
        if let Some(change) = self
            .list_changes
            .iter_mut()
            .find(|change| change.list == list)
        {
            change.after = normalized_after;
        } else {
            self.list_changes
                .try_reserve(1)
                .map_err(|source| Error::Allocation {
                    resource: "OTH staged list changes",
                    source,
                })?;
            self.list_changes.push(ListChange {
                after: normalized_after,
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
            && self.forms_change.is_none()
            && self.heading_changes.is_empty()
            && self.inline_changes.is_empty()
            && self.list_changes.is_empty()
            && self.metadata.is_keep()
            && self.styles.is_keep()
            && self.appended.is_empty()
        {
            return Ok(Commit::unchanged(self.source.clone()));
        }
        let content = crate::codec::compact_for_publication(&replace_texts(
            self.source,
            &self.changes,
            &self.heading_changes,
            &self.inline_changes,
            self.forms_change.as_ref(),
            &self.list_changes,
            &self.appended,
            None,
        )?)?;
        let snapshot = Template {
            package: self.source.package.rebuild_with_parts(
                &content,
                &self.metadata,
                &self.styles,
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
                forms_change: self.forms_change,
                heading_changes: self.heading_changes,
                inline_changes: self.inline_changes,
                list_changes: self.list_changes,
                metadata: self.metadata,
                styles: self.styles,
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
        } else if !self.metadata.is_keep() && !other.metadata.is_keep() {
            Some(JoinFailure::Metadata)
        } else if !self.styles.is_keep() && !other.styles.is_keep() {
            Some(JoinFailure::Styles)
        } else if self.forms_change.is_some() && other.forms_change.is_some() {
            Some(JoinFailure::Forms)
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
                .or_else(|| inline_join_conflict(&self.inline_changes, &other))
                .or_else(|| inline_join_conflict(&other.inline_changes, self))
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
        if self.forms_change.is_none() {
            self.forms_change = other.forms_change;
        }
        self.heading_changes.extend(other.heading_changes);
        self.inline_changes.extend(other.inline_changes);
        self.list_changes.extend(other.list_changes);
        self.appended.extend(other.appended);
        if self.metadata.is_keep() {
            self.metadata = other.metadata;
        }
        if self.styles.is_keep() {
            self.styles = other.styles;
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
    /// Both transactions replace the form catalog.
    Forms,
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
                forms_change: None,
                heading_changes: Vec::new(),
                inline_changes: Vec::new(),
                list_changes: Vec::new(),
                metadata: PartChange::Keep,
                source: snapshot.clone(),
                styles: PartChange::Keep,
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
    forms_change: Option<FormsChange>,
    heading_changes: Vec<HeadingChange>,
    inline_changes: Vec<InlineChange>,
    list_changes: Vec<ListChange>,
    metadata: PartChange,
    source: Template,
    styles: PartChange,
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
        collect_inline_merge_conflicts(&left.inline_changes, right, &mut conflicts);
        collect_inline_merge_conflicts(&right.inline_changes, left, &mut conflicts);
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
        if !left.metadata.is_keep() && !right.metadata.is_keep() {
            conflicts.push(MergeConflict::Metadata);
        }
        if !left.styles.is_keep() && !right.styles.is_keep() {
            conflicts.push(MergeConflict::Styles);
        }
        if left.forms_change.is_some() && right.forms_change.is_some() {
            conflicts.push(MergeConflict::Forms);
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

    /// Rich inline-content changes in staging order.
    #[must_use]
    pub fn inline_changes(&self) -> &[InlineChange] {
        &self.inline_changes
    }

    /// Complete form-catalog change, if present.
    #[must_use]
    pub const fn forms_change(&self) -> Option<&FormsChange> {
        self.forms_change.as_ref()
    }

    /// Typed list changes in staging order.
    #[must_use]
    pub fn list_changes(&self) -> &[ListChange] {
        &self.list_changes
    }

    /// Replacement metadata XML retained by this semantic patch.
    #[must_use]
    pub fn metadata_xml(&self) -> Option<&str> {
        self.metadata.replacement()
    }

    /// Whether this patch removes `meta.xml`.
    #[must_use]
    pub const fn removes_metadata(&self) -> bool {
        self.metadata.removes()
    }

    /// Replacement styles XML retained by this semantic patch.
    #[must_use]
    pub fn styles_xml(&self) -> Option<&str> {
        self.styles.replacement()
    }

    /// Whether this patch removes `styles.xml`.
    #[must_use]
    pub const fn removes_styles(&self) -> bool {
        self.styles.removes()
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
        push_wire_usize(&mut output, self.inline_changes.len())?;
        for change in &self.inline_changes {
            match change.block {
                InlineBlock::Paragraph(position) => {
                    push_wire_usize(&mut output, 0)?;
                    push_wire_usize(&mut output, position.get())?;
                },
                InlineBlock::Heading(position) => {
                    push_wire_usize(&mut output, 1)?;
                    push_wire_usize(&mut output, position.get())?;
                },
            }
            push_wire_bytes(&mut output, change.before_xml.as_bytes())?;
            push_wire_bytes(&mut output, change.after_xml.as_bytes())?;
            push_wire_bytes(&mut output, change.before_text.as_bytes())?;
            push_wire_bytes(&mut output, change.after_text.as_bytes())?;
        }
        match &self.forms_change {
            None => push_wire_usize(&mut output, 0)?,
            Some(change) => {
                push_wire_usize(&mut output, 1)?;
                push_wire_bytes(&mut output, change.before_xml.as_bytes())?;
                push_wire_bytes(&mut output, change.after_xml.as_bytes())?;
            },
        }
        push_wire_usize(&mut output, self.list_changes.len())?;
        for change in &self.list_changes {
            push_wire_usize(&mut output, change.list.get())?;
            push_wire_list(&mut output, change.before.as_ref())?;
            push_wire_list(&mut output, change.after.as_ref())?;
        }
        push_wire_part_change(&mut output, &self.metadata)?;
        push_wire_part_change(&mut output, &self.styles)?;
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
        let inline_changes = read_inline_changes(bytes, &mut cursor, &source, &target)?;
        let forms_change = read_forms_change(bytes, &mut cursor, &source, &target)?;
        let list_changes = read_list_changes(bytes, &mut cursor, &source, &target)?;
        let metadata = read_wire_part_change(bytes, &mut cursor)?;
        let styles = read_wire_part_change(bytes, &mut cursor)?;
        validate_part_change(&metadata, source.meta_xml(), target.meta_xml(), "metadata")?;
        validate_part_change(&styles, source.styles_xml(), target.styles_xml(), "styles")?;
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
            forms_change,
            heading_changes,
            inline_changes,
            list_changes,
            metadata,
            styles,
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
            forms_change: self.forms_change.as_ref().map(|change| FormsChange {
                after: change.before.clone(),
                after_xml: change.before_xml.clone(),
                before: change.after.clone(),
                before_xml: change.after_xml.clone(),
            }),
            heading_changes: self
                .heading_changes
                .iter()
                .map(|change| HeadingChange {
                    heading: change.heading,
                    before: change.after.clone(),
                    after: change.before.clone(),
                })
                .collect(),
            inline_changes: self
                .inline_changes
                .iter()
                .map(|change| InlineChange {
                    after_text: change.before_text.clone(),
                    after_xml: change.before_xml.clone(),
                    before_text: change.after_text.clone(),
                    before_xml: change.after_xml.clone(),
                    block: change.block,
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
            metadata: part_change_between(self.target.meta_xml(), self.source.meta_xml()),
            source: self.target.clone(),
            styles: part_change_between(self.target.styles_xml(), self.source.styles_xml()),
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
    /// Both patches replace the form catalog.
    Forms,
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
        let mut inline_changes = self.left.inline_changes.clone();
        inline_changes.extend(self.right.inline_changes.clone());
        let forms_change = self
            .left
            .forms_change
            .as_ref()
            .or(self.right.forms_change.as_ref());
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
            &inline_changes,
            forms_change,
            &list_changes,
            &appended,
            durable_fragment,
        )?)?;
        let metadata = if self.left.metadata.is_keep() {
            &self.right.metadata
        } else {
            &self.left.metadata
        };
        let styles = if self.left.styles.is_keep() {
            &self.right.styles
        } else {
            &self.left.styles
        };
        let candidate = Template {
            package: self
                .base
                .package
                .rebuild_with_parts(&content, metadata, styles)?,
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
        validate_inline_readback(&inline_changes, &candidate)?;
        validate_forms_readback(forms_change, &candidate)?;
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

fn transfer_block(source: &Template, selector: TransferSelector) -> Result<crate::ContentBlock> {
    match selector {
        TransferSelector::Heading(position) => {
            let heading = source
                .package
                .headings()
                .get(position.get())
                .ok_or_else(|| {
                    Error::InvalidFormat("OTH transfer heading is out of bounds".to_string())
                })?;
            if !heading.links().is_empty()
                || !heading.fields().is_empty()
                || !heading.formatting_runs().is_empty()
            {
                return Err(Error::InvalidFormat(
                    "OTH transfer refuses a heading with rich inline markup".to_string(),
                ));
            }
            Ok(crate::ContentBlock::Heading(heading.clone()))
        },
        TransferSelector::Paragraph(position) => {
            let paragraph = source
                .package
                .paragraphs()
                .get(position.get())
                .ok_or_else(|| {
                    Error::InvalidFormat("OTH transfer paragraph is out of bounds".to_string())
                })?;
            validate_transfer_paragraph(paragraph)?;
            Ok(crate::ContentBlock::Paragraph(paragraph.clone()))
        },
        TransferSelector::List(position) => {
            let list = source.package.lists().get(position.get()).ok_or_else(|| {
                Error::InvalidFormat("OTH transfer list is out of bounds".to_string())
            })?;
            let site = source.package.list_site(position.get()).ok_or_else(|| {
                Error::InvalidFormat("OTH transfer list site is missing".to_string())
            })?;
            let contains_nested =
                source
                    .package
                    .list_sites()
                    .iter()
                    .enumerate()
                    .any(|(index, candidate)| {
                        index != position.get()
                            && candidate.range.start >= site.range.start
                            && candidate.range.end <= site.range.end
                    });
            if list.level() != 1 || contains_nested {
                return Err(Error::InvalidFormat(
                    "OTH transfer refuses nested list structure".to_string(),
                ));
            }
            for paragraph in list.items().iter().flat_map(crate::list::Item::paragraphs) {
                validate_transfer_paragraph(paragraph)?;
            }
            Ok(crate::ContentBlock::List(list.clone()))
        },
    }
}

fn validate_transfer_paragraph(paragraph: &crate::paragraph::Paragraph) -> Result<()> {
    if paragraph.links().is_empty()
        && paragraph.fields().is_empty()
        && paragraph.formatting_runs().is_empty()
    {
        Ok(())
    } else {
        Err(Error::InvalidFormat(
            "OTH transfer refuses a paragraph with rich inline markup".to_string(),
        ))
    }
}

fn block_style_names(block: &crate::ContentBlock) -> Vec<String> {
    match block {
        crate::ContentBlock::Heading(heading) => heading
            .style_name()
            .map(str::to_owned)
            .into_iter()
            .collect(),
        crate::ContentBlock::Paragraph(paragraph) => paragraph
            .style_name()
            .map(str::to_owned)
            .into_iter()
            .collect(),
        crate::ContentBlock::List(list) => list
            .items()
            .iter()
            .flat_map(crate::list::Item::paragraphs)
            .filter_map(crate::paragraph::Paragraph::style_name)
            .map(str::to_owned)
            .collect(),
    }
}

fn resolve_transfer_styles(
    source: &[crate::style::Style],
    destination: &[crate::style::Style],
    roots: &[String],
    include_styles: bool,
) -> Result<Vec<crate::style::Style>> {
    let mut pending = roots.to_vec();
    let mut visited = Vec::<String>::new();
    let mut imported = Vec::new();
    let mut index = 0;
    while let Some(name) = pending.get(index) {
        index = index.saturating_add(1);
        if visited.iter().any(|candidate| candidate == name) {
            continue;
        }
        visited.push(name.clone());
        let source_matches = source
            .iter()
            .filter(|style| style.name() == name)
            .collect::<Vec<_>>();
        let destination_matches = destination
            .iter()
            .filter(|style| style.name() == name)
            .collect::<Vec<_>>();
        if source_matches.len() > 1 || destination_matches.len() > 1 {
            return Err(Error::InvalidFormat(format!(
                "OTH transfer style {name:?} is ambiguous"
            )));
        }
        let source_style = source_matches.first().copied();
        let destination_style = destination_matches.first().copied();
        match (source_style, destination_style) {
            (Some(left), Some(right)) if !styles_semantically_equal(left, right) => {
                return Err(Error::InvalidFormat(format!(
                    "OTH transfer style {name:?} collides with the destination"
                )));
            },
            (Some(style), None) if include_styles => imported.push(style.clone()),
            (Some(_) | None, None) => {
                return Err(Error::InvalidFormat(format!(
                    "OTH transfer style dependency {name:?} is unavailable"
                )));
            },
            (Some(_) | None, Some(_)) => {},
        }
        if let Some(parent) = source_style.and_then(crate::style::Style::parent_name) {
            pending.push(parent.to_owned());
        }
    }
    Ok(imported)
}

fn styles_semantically_equal(left: &crate::style::Style, right: &crate::style::Style) -> bool {
    left.name() == right.name()
        && left.family() == right.family()
        && left.parent_name() == right.parent_name()
        && left.text_properties() == right.text_properties()
}

fn inline_join_conflict(changes: &[InlineChange], other: &Edit<'_>) -> Option<JoinFailure> {
    changes.iter().find_map(|change| match change.block {
        InlineBlock::Paragraph(position) => (other
            .inline_changes
            .iter()
            .any(|incoming| incoming.block == change.block)
            || other
                .changes
                .iter()
                .any(|incoming| incoming.paragraph == position))
        .then_some(JoinFailure::Paragraph(position))
        .or_else(|| {
            other.list_changes.iter().find_map(|list_change| {
                list_change.before.as_ref().and_then(|list| {
                    list_contains_paragraph(list, position).then_some(JoinFailure::ListParagraph {
                        list: list_change.list,
                        paragraph: position,
                    })
                })
            })
        }),
        InlineBlock::Heading(position) => (other
            .inline_changes
            .iter()
            .any(|incoming| incoming.block == change.block)
            || other
                .heading_changes
                .iter()
                .any(|incoming| incoming.heading == position))
        .then_some(JoinFailure::Heading(position)),
    })
}

fn collect_inline_merge_conflicts(
    changes: &[InlineChange],
    other: &Patch,
    conflicts: &mut Vec<MergeConflict>,
) {
    for change in changes {
        match change.block {
            InlineBlock::Paragraph(position) => {
                if other
                    .inline_changes
                    .iter()
                    .any(|incoming| incoming.block == change.block)
                    || other
                        .changes
                        .iter()
                        .any(|incoming| incoming.paragraph == position)
                {
                    conflicts.push(MergeConflict::Paragraph(position));
                }
                for list_change in &other.list_changes {
                    if list_change
                        .before
                        .as_ref()
                        .is_some_and(|list| list_contains_paragraph(list, position))
                    {
                        conflicts.push(MergeConflict::ListParagraph {
                            list: list_change.list,
                            paragraph: position,
                        });
                    }
                }
            },
            InlineBlock::Heading(position) => {
                if other
                    .inline_changes
                    .iter()
                    .any(|incoming| incoming.block == change.block)
                    || other
                        .heading_changes
                        .iter()
                        .any(|incoming| incoming.heading == position)
                {
                    conflicts.push(MergeConflict::Heading(position));
                }
            },
        }
    }
}

fn validate_inline_readback(changes: &[InlineChange], snapshot: &Template) -> Result<()> {
    for change in changes {
        validate_one_inline_change(change, snapshot, true)?;
    }
    Ok(())
}

fn validate_forms_readback(
    optional_change: Option<&FormsChange>,
    snapshot: &Template,
) -> Result<()> {
    if let Some(forms_change) = optional_change {
        validate_forms_source(forms_change, snapshot, true)?;
    }
    Ok(())
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
    validate_inline_readback(&edit.inline_changes, snapshot)?;
    validate_forms_readback(edit.forms_change.as_ref(), snapshot)?;
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
    validate_part_change(
        &edit.metadata,
        edit.source.meta_xml(),
        snapshot.meta_xml(),
        "metadata",
    )?;
    validate_part_change(
        &edit.styles,
        edit.source.styles_xml(),
        snapshot.styles_xml(),
        "styles",
    )?;
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

fn push_wire_list(output: &mut Vec<u8>, optional_list: Option<&crate::list::List>) -> Result<()> {
    match optional_list {
        None => push_wire_usize(output, 0),
        Some(value) => {
            push_wire_usize(output, 1)?;
            push_wire_usize(output, value.level())?;
            let xml =
                crate::authoring::render_fragment(&[crate::ContentBlock::List(value.clone())])?;
            push_wire_bytes(output, xml.as_bytes())
        },
    }
}

fn push_wire_part_change(output: &mut Vec<u8>, change: &PartChange) -> Result<()> {
    match change {
        PartChange::Keep => push_wire_usize(output, 0),
        PartChange::Remove => push_wire_usize(output, 1),
        PartChange::Set(xml) => {
            push_wire_usize(output, 2)?;
            push_wire_bytes(output, xml.as_bytes())
        },
    }
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

fn read_wire_list(bytes: &[u8], cursor: &mut usize) -> Result<Option<crate::list::List>> {
    match read_wire_usize(bytes, cursor)? {
        0 => Ok(None),
        1 => {
            let level = read_wire_usize(bytes, cursor)?;
            if level == 0 {
                return Err(Error::InvalidFormat(
                    "OTH durable patch list level is invalid".to_string(),
                ));
            }
            let fragment = read_wire_string(bytes, cursor)?;
            let wrapped = format!(
                "<?xml version=\"1.0\" encoding=\"UTF-8\"?><office:document-content xmlns:office=\"urn:oasis:names:tc:opendocument:xmlns:office:1.0\" xmlns:text=\"urn:oasis:names:tc:opendocument:xmlns:text:1.0\"><office:body><office:text>{fragment}</office:text></office:body></office:document-content>"
            );
            let mut projection = crate::codec::project(&wrapped)?;
            if projection.lists.len() != 1 || projection.lists[0].level() != 1 {
                return Err(Error::InvalidFormat(
                    "OTH durable patch list fragment is not one top-level list".to_string(),
                ));
            }
            let parsed = projection.lists.pop().ok_or_else(|| {
                Error::InvalidFormat("OTH durable patch list disappeared".to_string())
            })?;
            Ok(Some(crate::list::List::projected(
                parsed.items().to_vec(),
                level,
                parsed.style_name().map(str::to_owned),
            )))
        },
        _ => Err(Error::InvalidFormat(
            "OTH durable patch list marker is invalid".to_string(),
        )),
    }
}

fn read_wire_part_change(bytes: &[u8], cursor: &mut usize) -> Result<PartChange> {
    match read_wire_usize(bytes, cursor)? {
        0 => Ok(PartChange::Keep),
        1 => Ok(PartChange::Remove),
        2 => Ok(PartChange::Set(read_wire_string(bytes, cursor)?)),
        _ => Err(Error::InvalidFormat(
            "OTH durable patch part-change marker is invalid".to_string(),
        )),
    }
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

fn read_inline_changes(
    bytes: &[u8],
    cursor: &mut usize,
    source: &Template,
    target: &Template,
) -> Result<Vec<InlineChange>> {
    let count = read_wire_usize(bytes, cursor)?;
    let maximum = source
        .package
        .paragraphs()
        .len()
        .saturating_add(source.package.headings().len());
    if count > maximum {
        return Err(Error::InvalidFormat(
            "OTH durable rich inline count is invalid".to_string(),
        ));
    }
    let mut changes = Vec::new();
    changes
        .try_reserve_exact(count)
        .map_err(|allocation_error| Error::Allocation {
            resource: "OTH durable rich inline changes",
            source: allocation_error,
        })?;
    for _ in 0..count {
        let kind = read_wire_usize(bytes, cursor)?;
        let position = Position::new(read_wire_usize(bytes, cursor)?);
        let block = match kind {
            0 => InlineBlock::Paragraph(position),
            1 => InlineBlock::Heading(position),
            _ => {
                return Err(Error::InvalidFormat(
                    "OTH durable rich inline block marker is invalid".to_string(),
                ));
            },
        };
        if changes
            .iter()
            .any(|change: &InlineChange| change.block == block)
        {
            return Err(Error::InvalidFormat(
                "OTH durable patch repeats a rich inline selector".to_string(),
            ));
        }
        let change = InlineChange {
            before_xml: read_wire_string(bytes, cursor)?,
            after_xml: read_wire_string(bytes, cursor)?,
            before_text: read_wire_string(bytes, cursor)?,
            after_text: read_wire_string(bytes, cursor)?,
            block,
        };
        validate_one_inline_change(&change, source, false)?;
        validate_one_inline_change(&change, target, true)?;
        changes.push(change);
    }
    Ok(changes)
}

fn validate_one_inline_change(
    change: &InlineChange,
    template: &Template,
    after: bool,
) -> Result<()> {
    let (optional_site, text) = match change.block {
        InlineBlock::Paragraph(position) => (
            template.package.paragraph_content_site(position.get()),
            template
                .package
                .paragraphs()
                .get(position.get())
                .map(crate::paragraph::Paragraph::text),
        ),
        InlineBlock::Heading(position) => (
            template.package.heading_content_site(position.get()),
            template
                .package
                .headings()
                .get(position.get())
                .map(crate::heading::Heading::text),
        ),
    };
    let content_site = optional_site.ok_or_else(|| {
        Error::InvalidFormat("OTH durable rich inline site is invalid".to_string())
    })?;
    let actual_xml = template
        .content_xml()
        .get(content_site.range.clone())
        .ok_or_else(|| {
            Error::InvalidFormat("OTH durable rich inline span is invalid".to_string())
        })?;
    let (expected_xml, expected_text) = if after {
        (change.after_xml.as_str(), change.after_text.as_str())
    } else {
        (change.before_xml.as_str(), change.before_text.as_str())
    };
    if actual_xml != expected_xml || text != Some(expected_text) {
        return Err(Error::InvalidFormat(
            "OTH durable rich inline semantic readback failed".to_string(),
        ));
    }
    Ok(())
}

fn read_forms_change(
    bytes: &[u8],
    cursor: &mut usize,
    source: &Template,
    target: &Template,
) -> Result<Option<FormsChange>> {
    match read_wire_usize(bytes, cursor)? {
        0 => Ok(None),
        1 => {
            let change = FormsChange {
                before: source.package.forms().to_vec(),
                before_xml: read_wire_string(bytes, cursor)?,
                after: target.package.forms().to_vec(),
                after_xml: read_wire_string(bytes, cursor)?,
            };
            validate_forms_source(&change, source, false)?;
            validate_forms_source(&change, target, true)?;
            Ok(Some(change))
        },
        _ => Err(Error::InvalidFormat(
            "OTH durable forms marker is invalid".to_string(),
        )),
    }
}

fn validate_forms_source(change: &FormsChange, template: &Template, after: bool) -> Result<()> {
    let actual_xml = match template.package.forms_site() {
        Some(site) => template
            .content_xml()
            .get(site.range.clone())
            .ok_or_else(|| Error::InvalidFormat("OTH forms source span is invalid".to_string()))?,
        None => "",
    };
    let (expected_xml, expected_forms) = if after {
        (change.after_xml.as_str(), change.after.as_slice())
    } else {
        (change.before_xml.as_str(), change.before.as_slice())
    };
    if actual_xml != expected_xml || template.package.forms() != expected_forms {
        return Err(Error::InvalidFormat(
            "OTH durable forms semantic readback failed".to_string(),
        ));
    }
    Ok(())
}

fn read_list_changes(
    bytes: &[u8],
    cursor: &mut usize,
    source: &Template,
    target: &Template,
) -> Result<Vec<ListChange>> {
    let count = read_wire_usize(bytes, cursor)?;
    if count
        > source
            .package
            .lists()
            .len()
            .saturating_add(target.package.lists().len())
    {
        return Err(Error::InvalidFormat(
            "OTH durable patch list count is invalid".to_string(),
        ));
    }
    let mut changes = Vec::new();
    changes
        .try_reserve_exact(count)
        .map_err(|allocation_error| Error::Allocation {
            resource: "OTH durable list changes",
            source: allocation_error,
        })?;
    for _ in 0..count {
        let list = Position::new(read_wire_usize(bytes, cursor)?);
        if changes
            .iter()
            .any(|candidate: &ListChange| candidate.list == list)
        {
            return Err(Error::InvalidFormat(
                "OTH durable patch repeats a list selector".to_string(),
            ));
        }
        let before = read_wire_list(bytes, cursor)?;
        let after = read_wire_list(bytes, cursor)?;
        if before.is_none() && after.is_none() {
            return Err(Error::InvalidFormat(
                "OTH durable patch contains an empty list change".to_string(),
            ));
        }
        if let Some(expected) = before.as_ref()
            && !source
                .package
                .lists()
                .get(list.get())
                .is_some_and(|actual| lists_semantically_equal(actual, expected))
        {
            return Err(Error::InvalidFormat(
                "OTH durable list change failed source readback".to_string(),
            ));
        }
        changes.push(ListChange {
            after,
            before,
            list,
        });
    }
    let removed = changes
        .iter()
        .filter(|change| change.before.is_some() && change.after.is_none())
        .count();
    let inserted = changes
        .iter()
        .filter(|change| change.before.is_none() && change.after.is_some())
        .count();
    if target.package.lists().len()
        != source
            .package
            .lists()
            .len()
            .saturating_sub(removed)
            .saturating_add(inserted)
    {
        return Err(Error::InvalidFormat(
            "OTH durable list changes fail structural target readback".to_string(),
        ));
    }
    for change in &changes {
        let Some(expected) = change.after.as_ref() else {
            continue;
        };
        let target_index = list_target_index(&changes, change)?;
        if !target
            .package
            .lists()
            .get(target_index)
            .is_some_and(|actual| lists_semantically_equal(actual, expected))
        {
            return Err(Error::InvalidFormat(
                "OTH durable list change failed target readback".to_string(),
            ));
        }
    }
    Ok(changes)
}

fn part_change_between(before: Option<&str>, after: Option<&str>) -> PartChange {
    match (before, after) {
        (left, right) if left == right => PartChange::Keep,
        (_, Some(xml)) => PartChange::Set(xml.to_owned()),
        (_, None) => PartChange::Remove,
    }
}

fn validate_part_change(
    change: &PartChange,
    before: Option<&str>,
    after: Option<&str>,
    label: &str,
) -> Result<()> {
    let valid = match change {
        PartChange::Keep => before == after,
        PartChange::Remove => after.is_none(),
        PartChange::Set(xml) => after == Some(xml.as_str()),
    };
    if valid {
        Ok(())
    } else {
        Err(Error::InvalidFormat(format!(
            "OTH {label} part change failed exact readback"
        )))
    }
}

fn replace_texts(
    source: &Template,
    changes: &[ParagraphChange],
    heading_changes: &[HeadingChange],
    inline_changes: &[InlineChange],
    forms_change: Option<&FormsChange>,
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
                .saturating_add(inline_changes.len())
                .saturating_add(usize::from(forms_change.is_some()))
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
    for change in inline_changes {
        let site = match change.block {
            InlineBlock::Paragraph(position) => source
                .package
                .paragraph_content_site(position.get())
                .ok_or_else(|| {
                    Error::InvalidFormat("OTH rich paragraph edit site disappeared".to_string())
                })?,
            InlineBlock::Heading(position) => source
                .package
                .heading_content_site(position.get())
                .ok_or_else(|| {
                Error::InvalidFormat("OTH rich heading edit site disappeared".to_string())
            })?,
        };
        let actual = source
            .content_xml()
            .get(site.range.clone())
            .ok_or_else(|| {
                Error::InvalidFormat("OTH rich inline edit source span is invalid".to_string())
            })?;
        if actual != change.before_xml {
            return Err(Error::InvalidFormat(
                "OTH rich inline edit source precondition failed".to_string(),
            ));
        }
        replacements.push((
            site.clone(),
            std::borrow::Cow::Borrowed(change.after_xml.as_str()),
        ));
    }
    if let Some(change) = forms_change {
        let site =
            source
                .package
                .forms_site()
                .cloned()
                .unwrap_or_else(|| crate::codec::ReplacementSite {
                    prefix: String::new(),
                    range: source.package.text_close()..source.package.text_close(),
                    suffix: String::new(),
                });
        let actual = source
            .content_xml()
            .get(site.range.clone())
            .ok_or_else(|| Error::InvalidFormat("OTH forms source span is invalid".to_string()))?;
        if actual != change.before_xml {
            return Err(Error::InvalidFormat(
                "OTH forms source precondition failed".to_string(),
            ));
        }
        replacements.push((site, std::borrow::Cow::Borrowed(change.after_xml.as_str())));
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
