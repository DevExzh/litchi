//! Concise family entry points.

use litchi_core::{Error, HistoryLimits, Metadata, PatchError, Position, Result};
use litchi_odf_common::compact_xml;
use std::fmt;
use std::{path::Path, sync::Arc};

pub use crate::authoring::{Builder, ResourceMember};

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

    /// Returns the stable XML validation capability contract.
    #[must_use]
    pub const fn validation_capabilities(&self) -> ValidationCapabilities {
        ValidationCapabilities::oth()
    }

    /// Returns the stable security lifecycle capability contract.
    #[must_use]
    pub const fn security_capabilities(&self) -> SecurityCapabilities {
        SecurityCapabilities::oth()
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
        if let TransferSelector::Resource(position) = selector {
            let resource = source
                .package
                .resources()
                .get(position.get())
                .ok_or_else(|| {
                    Error::InvalidFormat("OTH transfer resource is out of bounds".to_string())
                })?
                .clone();
            let payloads = transfer_resource_payloads(source, self, &resource)?;
            return Ok(TransferPlan {
                block: None,
                destination: self.clone(),
                fragment: None,
                imported_styles: Vec::new(),
                payloads,
                resource: Some(resource),
            });
        }
        let (block, site) = transfer_block(source, selector)?;
        let fragment = source
            .content_xml()
            .get(site.range.clone())
            .ok_or_else(|| Error::InvalidFormat("OTH transfer source span is invalid".to_string()))?
            .to_owned();
        let payloads = transfer_site_payloads(source, self, &site)?;
        let style_names = block_style_names(&block);
        let imported_styles = resolve_transfer_styles(
            source.styles(),
            self.styles(),
            &style_names,
            policy.include_styles,
        )?;
        Ok(TransferPlan {
            block: Some(block),
            destination: self.clone(),
            fragment: Some(fragment),
            imported_styles,
            payloads,
            resource: None,
        })
    }

    /// Starts a source-bound text-body transaction.
    #[must_use]
    pub fn edit(&self) -> Edit<'_> {
        Edit {
            appended: Vec::new(),
            changes: Vec::new(),
            durable_fragment: None,
            forms_change: None,
            heading_changes: Vec::new(),
            inline_changes: Vec::new(),
            list_changes: Vec::new(),
            metadata: PartChange::Keep,
            payload_changes: Vec::new(),
            resource_changes: Vec::new(),
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

/// Explicit support state for validation and security capabilities.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum CapabilityState {
    /// The capability is implemented at the documented OTH boundary.
    Supported,
    /// The capability is deliberately unavailable and fails closed.
    Refused,
}

/// Stable validation contract for OTH XML.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub struct ValidationCapabilities {
    compact_publication: CapabilityState,
    namespace_family_envelope: CapabilityState,
    odf_relax_ng: CapabilityState,
    semantic_subset: CapabilityState,
}

impl ValidationCapabilities {
    const fn oth() -> Self {
        Self {
            compact_publication: CapabilityState::Supported,
            namespace_family_envelope: CapabilityState::Supported,
            odf_relax_ng: CapabilityState::Refused,
            semantic_subset: CapabilityState::Supported,
        }
    }

    /// Namespace-aware root/body/text family-envelope validation.
    #[must_use]
    pub const fn namespace_family_envelope(self) -> CapabilityState {
        self.namespace_family_envelope
    }

    /// Bounded validation for the projected text/list/form/resource subset.
    #[must_use]
    pub const fn semantic_subset(self) -> CapabilityState {
        self.semantic_subset
    }

    /// Full OASIS ODF Relax NG validation.
    #[must_use]
    pub const fn odf_relax_ng(self) -> CapabilityState {
        self.odf_relax_ng
    }

    /// Compact authored and changed-part publication.
    #[must_use]
    pub const fn compact_publication(self) -> CapabilityState {
        self.compact_publication
    }
}

/// Stable security and protected-package lifecycle contract.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub struct SecurityCapabilities {
    active_execution: CapabilityState,
    changed_signed_publication: CapabilityState,
    external_resolution: CapabilityState,
    inert_inventory: CapabilityState,
    password_open: CapabilityState,
    policy_gate: CapabilityState,
    signature_verification: CapabilityState,
}

impl SecurityCapabilities {
    const fn oth() -> Self {
        Self {
            active_execution: CapabilityState::Refused,
            changed_signed_publication: CapabilityState::Refused,
            external_resolution: CapabilityState::Refused,
            inert_inventory: CapabilityState::Supported,
            password_open: CapabilityState::Refused,
            policy_gate: CapabilityState::Supported,
            signature_verification: CapabilityState::Refused,
        }
    }

    /// Inert inventory of embedded/external resources, forms, scripts, and signatures.
    #[must_use]
    pub const fn inert_inventory(self) -> CapabilityState {
        self.inert_inventory
    }

    /// Explicit default-deny policy enforcement over inventoried surfaces.
    #[must_use]
    pub const fn policy_gate(self) -> CapabilityState {
        self.policy_gate
    }

    /// Network or external-link resolution.
    #[must_use]
    pub const fn external_resolution(self) -> CapabilityState {
        self.external_resolution
    }

    /// Script, macro, form, object, field, or action execution.
    #[must_use]
    pub const fn active_execution(self) -> CapabilityState {
        self.active_execution
    }

    /// Password-encrypted package opening.
    #[must_use]
    pub const fn password_open(self) -> CapabilityState {
        self.password_open
    }

    /// Cryptographic signature verification.
    #[must_use]
    pub const fn signature_verification(self) -> CapabilityState {
        self.signature_verification
    }

    /// Changed publication of a signed source.
    #[must_use]
    pub const fn changed_signed_publication(self) -> CapabilityState {
        self.changed_signed_publication
    }
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
    /// An inert resource/object reference by projection position.
    Resource(Position),
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
    block: Option<crate::ContentBlock>,
    destination: Template,
    fragment: Option<String>,
    imported_styles: Vec<crate::style::Style>,
    payloads: Vec<ResourcePayloadChange>,
    resource: Option<crate::resource::Resource>,
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
        for payload in self.payloads {
            edit.stage_payload(payload.path, payload.media_type, payload.after)?;
        }
        if let Some(resource) = self.resource {
            edit.append_resource(resource)?;
        }
        if let Some(block) = self.block {
            edit.append_block(block)?;
        }
        edit.durable_fragment = self.fragment;
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

#[derive(Clone, Copy)]
enum InlinePlacement {
    Append,
    Prepend,
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

/// One reversible resource/object reference replacement, insertion, or removal.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct ResourceChange {
    after: Option<crate::resource::Resource>,
    after_xml: String,
    before: Option<crate::resource::Resource>,
    before_xml: String,
    resource: Position,
}

impl ResourceChange {
    /// Projected resource position.
    #[must_use]
    pub const fn resource(&self) -> Position {
        self.resource
    }

    /// Source reference, absent for insertion.
    #[must_use]
    pub const fn before(&self) -> Option<&crate::resource::Resource> {
        self.before.as_ref()
    }

    /// Replacement reference, absent for removal.
    #[must_use]
    pub const fn after(&self) -> Option<&crate::resource::Resource> {
        self.after.as_ref()
    }
}

/// One reversible embedded package-member payload change.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct ResourcePayloadChange {
    pub(crate) after: Option<Vec<u8>>,
    pub(crate) before: Option<Vec<u8>>,
    pub(crate) media_type: String,
    pub(crate) path: String,
}

impl ResourcePayloadChange {
    /// Package-relative member path.
    #[must_use]
    pub fn path(&self) -> &str {
        &self.path
    }

    /// Replacement media type.
    #[must_use]
    pub fn media_type(&self) -> &str {
        &self.media_type
    }

    /// Exact source bytes, if the member existed.
    #[must_use]
    pub fn before(&self) -> Option<&[u8]> {
        self.before.as_deref()
    }

    /// Replacement bytes, or `None` for deletion.
    #[must_use]
    pub fn after(&self) -> Option<&[u8]> {
        self.after.as_deref()
    }
}

pub(crate) type MemberChange = ResourcePayloadChange;

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
    durable_fragment: Option<String>,
    forms_change: Option<FormsChange>,
    heading_changes: Vec<HeadingChange>,
    inline_changes: Vec<InlineChange>,
    list_changes: Vec<ListChange>,
    metadata: PartChange,
    payload_changes: Vec<ResourcePayloadChange>,
    resource_changes: Vec<ResourceChange>,
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

    /// Prepends typed rich content without rewriting existing paragraph markup.
    ///
    /// This exact-boundary splice preserves projected and unknown producer
    /// inline elements already present in the paragraph.
    ///
    /// # Errors
    ///
    /// Returns an error for an invalid selector, overlapping staged work, or
    /// content that cannot be compactly authored and fully reopened.
    pub fn prepend_paragraph_inline(
        &mut self,
        paragraph: Position,
        content: &[crate::inline::Content],
    ) -> Result<()> {
        self.stage_inline_boundary(
            InlineBlock::Paragraph(paragraph),
            content,
            InlinePlacement::Prepend,
        )
    }

    /// Appends typed rich content without rewriting existing paragraph markup.
    ///
    /// # Errors
    ///
    /// Returns an error for an invalid selector, overlapping staged work, or
    /// content that cannot be compactly authored and fully reopened.
    pub fn append_paragraph_inline(
        &mut self,
        paragraph: Position,
        content: &[crate::inline::Content],
    ) -> Result<()> {
        self.stage_inline_boundary(
            InlineBlock::Paragraph(paragraph),
            content,
            InlinePlacement::Append,
        )
    }

    /// Prepends typed rich content without rewriting existing heading markup.
    ///
    /// # Errors
    ///
    /// Returns an error for an invalid selector, overlapping staged work, or
    /// content that cannot be compactly authored and fully reopened.
    pub fn prepend_heading_inline(
        &mut self,
        heading: Position,
        content: &[crate::inline::Content],
    ) -> Result<()> {
        self.stage_inline_boundary(
            InlineBlock::Heading(heading),
            content,
            InlinePlacement::Prepend,
        )
    }

    /// Appends typed rich content without rewriting existing heading markup.
    ///
    /// # Errors
    ///
    /// Returns an error for an invalid selector, overlapping staged work, or
    /// content that cannot be compactly authored and fully reopened.
    pub fn append_heading_inline(
        &mut self,
        heading: Position,
        content: &[crate::inline::Content],
    ) -> Result<()> {
        self.stage_inline_boundary(
            InlineBlock::Heading(heading),
            content,
            InlinePlacement::Append,
        )
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
                if !self
                    .source
                    .package
                    .paragraph_inline_replaceable(position.get())
                {
                    return Err(Error::InvalidFormat(
                        "OTH rich paragraph replacement refuses unknown inline content".to_string(),
                    ));
                }
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
                if !self
                    .source
                    .package
                    .heading_inline_replaceable(position.get())
                {
                    return Err(Error::InvalidFormat(
                        "OTH rich heading replacement refuses unknown inline content".to_string(),
                    ));
                }
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

    fn stage_inline_boundary(
        &mut self,
        block: InlineBlock,
        content: &[crate::inline::Content],
        placement: InlinePlacement,
    ) -> Result<()> {
        let (site, source_text) = self.inline_boundary_source(block)?;
        let source_xml = self
            .source
            .content_xml()
            .get(site.range.clone())
            .ok_or_else(|| {
                Error::InvalidFormat("OTH inline boundary source span is invalid".to_string())
            })?;
        let existing = self
            .inline_changes
            .iter()
            .find(|change| change.block == block)
            .cloned();
        let (before_xml, before_text, current_xml, current_text) = match existing {
            Some(change) => (
                change.before_xml,
                change.before_text,
                change.after_xml,
                change.after_text,
            ),
            None => (
                source_xml.to_owned(),
                source_text.to_owned(),
                if site.prefix.is_empty() && site.suffix.is_empty() {
                    source_xml.to_owned()
                } else {
                    String::new()
                },
                source_text.to_owned(),
            ),
        };
        let (addition_xml, addition_text) = crate::authoring::render_inline(content)?;
        if addition_xml.is_empty() {
            return Ok(());
        }
        let target_xml_bytes = current_xml
            .len()
            .checked_add(addition_xml.len())
            .ok_or_else(|| Error::InvalidFormat("OTH inline boundary size overflow".to_string()))?;
        let target_text_bytes = current_text
            .len()
            .checked_add(addition_text.len())
            .ok_or_else(|| Error::InvalidFormat("OTH inline boundary size overflow".to_string()))?;
        if target_xml_bytes > compact_xml::DEFAULT_MAX_BYTES
            || target_text_bytes > MAX_PARAGRAPH_BYTES
        {
            return Err(Error::InvalidFormat(
                "OTH inline boundary edit exceeds the limit".to_string(),
            ));
        }
        let mut after_xml = String::new();
        after_xml
            .try_reserve_exact(target_xml_bytes)
            .map_err(|source| Error::Allocation {
                resource: "OTH inline boundary XML",
                source,
            })?;
        let mut after_text = String::new();
        after_text
            .try_reserve_exact(target_text_bytes)
            .map_err(|source| Error::Allocation {
                resource: "OTH inline boundary text",
                source,
            })?;
        match placement {
            InlinePlacement::Prepend => {
                after_xml.push_str(&addition_xml);
                after_xml.push_str(&current_xml);
                after_text.push_str(&addition_text);
                after_text.push_str(&current_text);
            },
            InlinePlacement::Append => {
                after_xml.push_str(&current_xml);
                after_xml.push_str(&addition_xml);
                after_text.push_str(&current_text);
                after_text.push_str(&addition_text);
            },
        }
        let staged = InlineChange {
            after_text,
            after_xml,
            before_text,
            before_xml,
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
                    resource: "OTH staged inline boundary changes",
                    source,
                })?;
            self.inline_changes.push(staged);
        }
        Ok(())
    }

    fn inline_boundary_source(
        &self,
        block: InlineBlock,
    ) -> Result<(&crate::codec::ReplacementSite, &str)> {
        match block {
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
                        "OTH paragraph boundary edit overlaps staged work".to_string(),
                    ));
                }
                let value = self
                    .source
                    .package
                    .paragraphs()
                    .get(position.get())
                    .ok_or_else(|| {
                        Error::InvalidFormat(
                            "OTH paragraph boundary selector is out of bounds".to_string(),
                        )
                    })?;
                let site = self
                    .source
                    .package
                    .paragraph_content_site(position.get())
                    .ok_or_else(|| {
                        Error::InvalidFormat(
                            "OTH paragraph boundary content site is missing".to_string(),
                        )
                    })?;
                Ok((site, value.text()))
            },
            InlineBlock::Heading(position) => {
                if self
                    .heading_changes
                    .iter()
                    .any(|change| change.heading == position)
                {
                    return Err(Error::InvalidFormat(
                        "OTH heading boundary edit overlaps staged work".to_string(),
                    ));
                }
                let value = self
                    .source
                    .package
                    .headings()
                    .get(position.get())
                    .ok_or_else(|| {
                        Error::InvalidFormat(
                            "OTH heading boundary selector is out of bounds".to_string(),
                        )
                    })?;
                let site = self
                    .source
                    .package
                    .heading_content_site(position.get())
                    .ok_or_else(|| {
                        Error::InvalidFormat(
                            "OTH heading boundary content site is missing".to_string(),
                        )
                    })?;
                Ok((site, value.text()))
            },
        }
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

    /// Replaces one inert resource or object reference.
    ///
    /// # Errors
    ///
    /// Returns an error for an invalid selector or source span.
    pub fn set_resource(
        &mut self,
        resource: Position,
        replacement: crate::resource::Resource,
    ) -> Result<()> {
        self.stage_resource(resource, Some(replacement), false)
    }

    /// Removes one inert resource or object reference while preserving payload bytes.
    ///
    /// # Errors
    ///
    /// Returns an error for an invalid selector or source span.
    pub fn remove_resource(&mut self, resource: Position) -> Result<()> {
        self.stage_resource(resource, None, false)
    }

    /// Appends one inert reference whose embedded payload already exists, or
    /// whose target is external.
    ///
    /// # Errors
    ///
    /// Returns an error when an embedded dependency is absent.
    pub fn append_resource(&mut self, resource: crate::resource::Resource) -> Result<()> {
        if resource.is_embedded() {
            let path = embedded_path(resource.href())?;
            let prefix = format!("{path}/");
            let staged = self.payload_changes.iter().any(|change| {
                (change.path == path || change.path.starts_with(&prefix)) && change.after.is_some()
            });
            let packaged = self
                .source
                .files()?
                .iter()
                .any(|member| member == path || member.starts_with(&prefix));
            if !staged && !packaged {
                return Err(Error::InvalidFormat(
                    "OTH appended embedded resource payload is absent".to_string(),
                ));
            }
        }
        let position = self.source.package.resources().len().saturating_add(
            self.resource_changes
                .iter()
                .filter(|change| change.before.is_none())
                .count(),
        );
        self.stage_resource(Position::new(position), Some(resource), true)
    }

    /// Creates an embedded payload and appends its inert reference atomically.
    ///
    /// # Errors
    ///
    /// Returns an error for an external reference, unsafe path, or invalid payload.
    pub fn append_resource_with_payload(
        &mut self,
        resource: crate::resource::Resource,
        media_type: impl Into<String>,
        bytes: Vec<u8>,
    ) -> Result<()> {
        let path = embedded_path(resource.href())?.to_owned();
        self.stage_payload(path, media_type.into(), Some(bytes))?;
        self.append_resource(resource)
    }

    /// Replaces or creates the payload referenced by one embedded resource.
    ///
    /// # Errors
    ///
    /// Returns an error for an invalid selector, external reference, or unsafe path.
    pub fn set_resource_payload(
        &mut self,
        resource: Position,
        media_type: impl Into<String>,
        bytes: Vec<u8>,
    ) -> Result<()> {
        let reference = self.resource_after(resource)?;
        let path = embedded_path(reference.href())?.to_owned();
        let prefix = format!("{path}/");
        let is_directory = reference.kind() != crate::resource::Kind::Image
            && (self
                .source
                .files()?
                .iter()
                .any(|member| member.starts_with(&prefix))
                || self
                    .payload_changes
                    .iter()
                    .any(|change| change.path.starts_with(&prefix) && change.after.is_some()));
        if is_directory {
            return Err(Error::InvalidFormat(
                "OTH directory-backed object payload requires member mutation".to_string(),
            ));
        }
        self.stage_payload(path, media_type.into(), Some(bytes))
    }

    /// Replaces or creates one member below a directory-backed embedded object.
    ///
    /// # Errors
    ///
    /// Returns an error for an image, external reference, unsafe member path,
    /// invalid selector, or invalid payload.
    pub fn set_resource_payload_member(
        &mut self,
        resource: Position,
        member: &str,
        media_type: impl Into<String>,
        bytes: Vec<u8>,
    ) -> Result<()> {
        let reference = self.resource_after(resource)?;
        if reference.kind() == crate::resource::Kind::Image {
            return Err(Error::InvalidFormat(
                "OTH image payloads do not have nested members".to_string(),
            ));
        }
        let path = resource_member_path(reference, member)?;
        self.stage_payload(path, media_type.into(), Some(bytes))
    }

    /// Removes the payload referenced by one embedded resource.
    ///
    /// # Errors
    ///
    /// Returns an error for an invalid selector, external reference, or unsafe path.
    pub fn remove_resource_payload(&mut self, resource: Position) -> Result<()> {
        let reference = self.resource_after(resource)?;
        let root = embedded_path(reference.href())?.to_owned();
        let prefix = format!("{root}/");
        let mut members = self
            .source
            .files()?
            .into_iter()
            .filter(|path| path == &root || path.starts_with(&prefix))
            .map(|path| {
                let media_type = self
                    .source
                    .package
                    .member_media_type(&path)?
                    .unwrap_or_else(|| "application/octet-stream".to_string());
                Ok((path, media_type))
            })
            .collect::<Result<Vec<_>>>()?;
        for change in &self.payload_changes {
            if change.after.is_some()
                && (change.path == root || change.path.starts_with(&prefix))
                && !members
                    .iter()
                    .any(|(path, _media_type)| path == &change.path)
            {
                members.push((change.path.clone(), change.media_type.clone()));
            }
        }
        if members.is_empty() {
            return Err(Error::InvalidFormat(
                "OTH embedded resource payload is absent".to_string(),
            ));
        }
        for (path, media_type) in members {
            self.stage_payload(path, media_type, None)?;
        }
        Ok(())
    }

    /// Removes one member below a directory-backed embedded object.
    ///
    /// # Errors
    ///
    /// Returns an error for an image, external reference, unsafe member path,
    /// or invalid selector.
    pub fn remove_resource_payload_member(
        &mut self,
        resource: Position,
        member: &str,
    ) -> Result<()> {
        let reference = self.resource_after(resource)?;
        if reference.kind() == crate::resource::Kind::Image {
            return Err(Error::InvalidFormat(
                "OTH image payloads do not have nested members".to_string(),
            ));
        }
        let path = resource_member_path(reference, member)?;
        let staged = self
            .payload_changes
            .iter()
            .any(|change| change.path == path && change.after.is_some());
        if !staged && self.source.package.member(&path)?.is_none() {
            return Err(Error::InvalidFormat(
                "OTH embedded object payload member is absent".to_string(),
            ));
        }
        let media_type = self
            .source
            .package
            .member_media_type(&path)?
            .unwrap_or_else(|| "application/octet-stream".to_string());
        self.stage_payload(path, media_type, None)
    }

    fn resource_after(&self, resource: Position) -> Result<&crate::resource::Resource> {
        if let Some(change) = self
            .resource_changes
            .iter()
            .find(|change| change.resource == resource)
        {
            return change.after.as_ref().ok_or_else(|| {
                Error::InvalidFormat("OTH removed resource has no payload target".to_string())
            });
        }
        self.source
            .package
            .resources()
            .get(resource.get())
            .ok_or_else(|| {
                Error::InvalidFormat("OTH resource selector is out of bounds".to_string())
            })
    }

    fn stage_resource(
        &mut self,
        resource: Position,
        after: Option<crate::resource::Resource>,
        insertion: bool,
    ) -> Result<()> {
        let before = (!insertion)
            .then(|| self.source.package.resources().get(resource.get()).cloned())
            .flatten();
        if !insertion && before.is_none() {
            return Err(Error::InvalidFormat(
                "OTH resource selector is out of bounds".to_string(),
            ));
        }
        if before == after {
            self.resource_changes
                .retain(|change| change.resource != resource);
            return Ok(());
        }
        let before_xml = if insertion {
            String::new()
        } else {
            let site = self
                .source
                .package
                .resource_site(resource.get())
                .ok_or_else(|| {
                    Error::InvalidFormat("OTH resource source site is missing".to_string())
                })?;
            self.source
                .content_xml()
                .get(site.range.clone())
                .ok_or_else(|| {
                    Error::InvalidFormat("OTH resource source span is invalid".to_string())
                })?
                .to_owned()
        };
        let after_xml = after.as_ref().map_or_else(String::new, |value| {
            let element = crate::authoring::render_resource(value);
            if insertion {
                format!("<text:p xmlns:text=\"urn:oasis:names:tc:opendocument:xmlns:text:1.0\"><draw:frame xmlns:draw=\"urn:oasis:names:tc:opendocument:xmlns:drawing:1.0\">{element}</draw:frame></text:p>")
            } else {
                element
            }
        });
        let staged = ResourceChange {
            after,
            after_xml,
            before,
            before_xml,
            resource,
        };
        if let Some(change) = self
            .resource_changes
            .iter_mut()
            .find(|change| change.resource == resource)
        {
            *change = staged;
        } else {
            self.resource_changes
                .try_reserve(1)
                .map_err(|allocation_error| Error::Allocation {
                    resource: "OTH staged resource changes",
                    source: allocation_error,
                })?;
            self.resource_changes.push(staged);
        }
        Ok(())
    }

    fn stage_payload(
        &mut self,
        path: String,
        media_type: String,
        after: Option<Vec<u8>>,
    ) -> Result<()> {
        if media_type.is_empty() {
            return Err(Error::InvalidFormat(
                "OTH resource media type cannot be empty".to_string(),
            ));
        }
        let before = self.source.package.member(&path)?;
        if before == after {
            self.payload_changes.retain(|change| change.path != path);
            return Ok(());
        }
        let staged = ResourcePayloadChange {
            after,
            before,
            media_type,
            path,
        };
        if let Some(change) = self
            .payload_changes
            .iter_mut()
            .find(|change| change.path == staged.path)
        {
            *change = staged;
        } else {
            self.payload_changes
                .try_reserve(1)
                .map_err(|allocation_error| Error::Allocation {
                    resource: "OTH staged resource payload changes",
                    source: allocation_error,
                })?;
            self.payload_changes.push(staged);
        }
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
        let normalized_after = after.map(|value| normalize_list_level(&value, before.level()));
        let site = self.source.package.list_site(list.get()).ok_or_else(|| {
            Error::InvalidFormat("OTH list structural site is missing".to_string())
        })?;
        if self.list_changes.iter().any(|change| {
            change.list != list
                && self
                    .source
                    .package
                    .list_site(change.list.get())
                    .is_some_and(|candidate| replacement_sites_overlap(site, candidate))
        }) {
            return Err(Error::InvalidFormat(
                "OTH list structural edits overlap".to_string(),
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
            && self.durable_fragment.is_none()
            && self.forms_change.is_none()
            && self.heading_changes.is_empty()
            && self.inline_changes.is_empty()
            && self.list_changes.is_empty()
            && self.metadata.is_keep()
            && self.payload_changes.is_empty()
            && self.resource_changes.is_empty()
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
            &self.resource_changes,
            &self.list_changes,
            &self.appended,
            self.durable_fragment.as_deref(),
        )?)?;
        let snapshot = Template {
            package: self.source.package.rebuild_with_parts(
                &content,
                &self.metadata,
                &self.styles,
                &self.payload_changes,
            )?,
        };
        validate_edit_readback(&self, &snapshot)?;
        let durable_fragment = published_durable_fragment(&self, &snapshot)?;
        Ok(Commit {
            snapshot: snapshot.clone(),
            patch: Patch {
                source: self.source.clone(),
                target: snapshot,
                appended: self.appended,
                changes: self.changes,
                durable_fragment,
                forms_change: self.forms_change,
                heading_changes: self.heading_changes,
                inline_changes: self.inline_changes,
                list_changes: self.list_changes,
                metadata: self.metadata,
                payload_changes: self.payload_changes,
                resource_changes: self.resource_changes,
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
        } else if self.payload_changes.iter().any(|accepted| {
            other
                .payload_changes
                .iter()
                .any(|incoming| incoming.path == accepted.path)
        }) {
            Some(JoinFailure::ResourcePayload)
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
                    self.resource_changes.iter().find_map(|accepted| {
                        other
                            .resource_changes
                            .iter()
                            .any(|incoming| incoming.resource == accepted.resource)
                            .then_some(JoinFailure::Resource(accepted.resource))
                    })
                })
                .or_else(|| {
                    self.list_changes.iter().find_map(|accepted| {
                        other
                            .list_changes
                            .iter()
                            .any(|incoming| list_changes_overlap(self.source, accepted, incoming))
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
        self.payload_changes.extend(other.payload_changes);
        self.resource_changes.extend(other.resource_changes);
        self.appended.extend(other.appended);
        if self.durable_fragment.is_none() {
            self.durable_fragment = other.durable_fragment;
        }
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
    /// Both transactions edit one resource reference.
    Resource(Position),
    /// Both transactions edit one embedded payload path.
    ResourcePayload,
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
                payload_changes: Vec::new(),
                resource_changes: Vec::new(),
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
    payload_changes: Vec<ResourcePayloadChange>,
    resource_changes: Vec<ResourceChange>,
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
                .any(|candidate| list_changes_overlap(base, change, candidate))
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
        for change in &left.resource_changes {
            if right
                .resource_changes
                .iter()
                .any(|candidate| candidate.resource == change.resource)
            {
                conflicts.push(MergeConflict::Resource(change.resource));
            }
        }
        if left.payload_changes.iter().any(|accepted| {
            right
                .payload_changes
                .iter()
                .any(|incoming| incoming.path == accepted.path)
        }) {
            conflicts.push(MergeConflict::ResourcePayload);
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

    /// Resource reference changes in staging order.
    #[must_use]
    pub fn resource_changes(&self) -> &[ResourceChange] {
        &self.resource_changes
    }

    /// Embedded payload changes in staging order.
    #[must_use]
    pub fn resource_payload_changes(&self) -> &[ResourcePayloadChange] {
        &self.payload_changes
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
        push_wire_usize(&mut output, self.resource_changes.len())?;
        for change in &self.resource_changes {
            push_wire_usize(&mut output, change.resource.get())?;
            push_wire_resource(&mut output, change.before.as_ref())?;
            push_wire_resource(&mut output, change.after.as_ref())?;
            push_wire_bytes(&mut output, change.before_xml.as_bytes())?;
            push_wire_bytes(&mut output, change.after_xml.as_bytes())?;
        }
        push_wire_usize(&mut output, self.payload_changes.len())?;
        for change in &self.payload_changes {
            push_wire_bytes(&mut output, change.path.as_bytes())?;
            push_wire_bytes(&mut output, change.media_type.as_bytes())?;
            push_wire_optional_bytes(&mut output, change.before.as_deref())?;
            push_wire_optional_bytes(&mut output, change.after.as_deref())?;
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
        let appended_list_count = if appended_xml.is_empty() {
            0
        } else {
            let wrapped = format!(
                "<?xml version=\"1.0\" encoding=\"UTF-8\"?><office:document-content xmlns:office=\"urn:oasis:names:tc:opendocument:xmlns:office:1.0\" xmlns:text=\"urn:oasis:names:tc:opendocument:xmlns:text:1.0\" xmlns:form=\"urn:oasis:names:tc:opendocument:xmlns:form:1.0\" xmlns:xlink=\"http://www.w3.org/1999/xlink\"><office:body><office:text>{appended_xml}</office:text></office:body></office:document-content>"
            );
            crate::codec::validate_authored(&wrapped)?;
            crate::codec::project(&wrapped)?.lists.len()
        };
        let changes = read_paragraph_changes(bytes, &mut cursor, &source, &target)?;
        let heading_changes = read_heading_changes(bytes, &mut cursor, &source, &target)?;
        let inline_changes = read_inline_changes(bytes, &mut cursor, &source, &target)?;
        let forms_change = read_forms_change(bytes, &mut cursor, &source, &target)?;
        let resource_changes = read_resource_changes(bytes, &mut cursor, &source, &target)?;
        let payload_changes = read_payload_changes(bytes, &mut cursor, &source, &target)?;
        let list_changes =
            read_list_changes(bytes, &mut cursor, &source, &target, appended_list_count)?;
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
            payload_changes,
            resource_changes,
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
            payload_changes: self
                .payload_changes
                .iter()
                .map(|change| ResourcePayloadChange {
                    after: change.before.clone(),
                    before: change.after.clone(),
                    media_type: change.media_type.clone(),
                    path: change.path.clone(),
                })
                .collect(),
            resource_changes: self
                .resource_changes
                .iter()
                .map(|change| ResourceChange {
                    after: change.before.clone(),
                    after_xml: change.before_xml.clone(),
                    before: change.after.clone(),
                    before_xml: change.after_xml.clone(),
                    resource: change.resource,
                })
                .collect(),
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
    /// Both patches edit one resource reference.
    Resource(Position),
    /// Both patches edit one embedded payload path.
    ResourcePayload,
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
        let mut resource_changes = self.left.resource_changes.clone();
        resource_changes.extend(self.right.resource_changes.clone());
        let mut payload_changes = self.left.payload_changes.clone();
        payload_changes.extend(self.right.payload_changes.clone());
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
            &resource_changes,
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
            package: self.base.package.rebuild_with_parts(
                &content,
                metadata,
                styles,
                &payload_changes,
            )?,
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
        validate_resource_readback(&resource_changes, &payload_changes, &candidate)?;
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

fn transfer_block(
    source: &Template,
    selector: TransferSelector,
) -> Result<(crate::ContentBlock, crate::codec::ReplacementSite)> {
    match selector {
        TransferSelector::Heading(position) => {
            let heading = source
                .package
                .headings()
                .get(position.get())
                .ok_or_else(|| {
                    Error::InvalidFormat("OTH transfer heading is out of bounds".to_string())
                })?;
            if !source.package.heading_inline_replaceable(position.get()) {
                return Err(Error::InvalidFormat(
                    "OTH transfer refuses unknown heading inline content".to_string(),
                ));
            }
            let site = source
                .package
                .heading_full_site(position.get())
                .ok_or_else(|| {
                    Error::InvalidFormat("OTH transfer heading site is missing".to_string())
                })?
                .clone();
            Ok((crate::ContentBlock::Heading(heading.clone()), site))
        },
        TransferSelector::Paragraph(position) => {
            let paragraph = source
                .package
                .paragraphs()
                .get(position.get())
                .ok_or_else(|| {
                    Error::InvalidFormat("OTH transfer paragraph is out of bounds".to_string())
                })?;
            if !source.package.paragraph_inline_replaceable(position.get()) {
                return Err(Error::InvalidFormat(
                    "OTH transfer refuses unknown paragraph inline content".to_string(),
                ));
            }
            let site = source
                .package
                .paragraph_full_site(position.get())
                .ok_or_else(|| {
                    Error::InvalidFormat("OTH transfer paragraph site is missing".to_string())
                })?
                .clone();
            Ok((crate::ContentBlock::Paragraph(paragraph.clone()), site))
        },
        TransferSelector::List(position) => {
            let list = source.package.lists().get(position.get()).ok_or_else(|| {
                Error::InvalidFormat("OTH transfer list is out of bounds".to_string())
            })?;
            let site = source.package.list_site(position.get()).ok_or_else(|| {
                Error::InvalidFormat("OTH transfer list site is missing".to_string())
            })?;
            for paragraph in list_paragraph_positions(list) {
                if !source.package.paragraph_inline_replaceable(paragraph.get()) {
                    return Err(Error::InvalidFormat(
                        "OTH transfer refuses unknown list inline content".to_string(),
                    ));
                }
            }
            Ok((
                crate::ContentBlock::List(normalize_list_level(list, 1)),
                site.clone(),
            ))
        },
        TransferSelector::Resource(_) => Err(Error::InvalidFormat(
            "OTH resource transfer must use dependency planning".to_string(),
        )),
    }
}

fn transfer_resource_payloads(
    source: &Template,
    destination: &Template,
    resource: &crate::resource::Resource,
) -> Result<Vec<ResourcePayloadChange>> {
    if !resource.is_embedded() {
        return Ok(Vec::new());
    }
    let root = embedded_path(resource.href())?;
    let prefix = format!("{root}/");
    let files = source.files()?;
    let mut paths = files
        .into_iter()
        .filter(|path| path == root || path.starts_with(&prefix))
        .collect::<Vec<_>>();
    paths.sort();
    if paths.is_empty() {
        return Err(Error::InvalidFormat(
            "OTH transfer embedded resource dependency is absent".to_string(),
        ));
    }
    let mut changes = Vec::new();
    changes
        .try_reserve_exact(paths.len())
        .map_err(|allocation_error| Error::Allocation {
            resource: "OTH transfer resource dependencies",
            source: allocation_error,
        })?;
    for path in paths {
        let after = source.package.member(&path)?.ok_or_else(|| {
            Error::InvalidFormat("OTH transfer resource dependency disappeared".to_string())
        })?;
        let before = destination.package.member(&path)?;
        if before.as_ref().is_some_and(|bytes| bytes != &after) {
            return Err(Error::InvalidFormat(format!(
                "OTH transfer resource dependency {path:?} collides"
            )));
        }
        if before.is_none() {
            changes.push(ResourcePayloadChange {
                after: Some(after),
                before,
                media_type: source
                    .package
                    .member_media_type(&path)?
                    .unwrap_or_else(|| "application/octet-stream".to_string()),
                path,
            });
        }
    }
    Ok(changes)
}

fn transfer_site_payloads(
    source: &Template,
    destination: &Template,
    site: &crate::codec::ReplacementSite,
) -> Result<Vec<ResourcePayloadChange>> {
    let mut changes = Vec::new();
    for (index, candidate) in source.package.resource_sites().iter().enumerate() {
        if candidate.range.start < site.range.start || candidate.range.end > site.range.end {
            continue;
        }
        let resource = source.package.resources().get(index).ok_or_else(|| {
            Error::InvalidFormat("OTH transfer resource projection disappeared".to_string())
        })?;
        for change in transfer_resource_payloads(source, destination, resource)? {
            if !changes
                .iter()
                .any(|accepted: &ResourcePayloadChange| accepted.path == change.path)
            {
                changes.push(change);
            }
        }
    }
    Ok(changes)
}

fn block_style_names(block: &crate::ContentBlock) -> Vec<String> {
    let mut names = Vec::new();
    match block {
        crate::ContentBlock::Heading(heading) => {
            names.extend(heading.style_name().map(str::to_owned));
            names.extend(
                heading
                    .formatting_runs()
                    .iter()
                    .map(crate::formatting::Run::style_name)
                    .map(str::to_owned),
            );
        },
        crate::ContentBlock::Paragraph(paragraph) => {
            extend_paragraph_style_names(&mut names, paragraph);
        },
        crate::ContentBlock::List(list) => extend_list_style_names(&mut names, list),
    }
    names
}

fn extend_paragraph_style_names(names: &mut Vec<String>, paragraph: &crate::paragraph::Paragraph) {
    names.extend(paragraph.style_name().map(str::to_owned));
    names.extend(
        paragraph
            .formatting_runs()
            .iter()
            .map(crate::formatting::Run::style_name)
            .map(str::to_owned),
    );
}

fn extend_list_style_names(names: &mut Vec<String>, list: &crate::list::List) {
    names.extend(list.style_name().map(str::to_owned));
    for item in list.items() {
        for paragraph in item.paragraphs() {
            extend_paragraph_style_names(names, paragraph);
        }
        for nested in item.nested_lists() {
            extend_list_style_names(names, nested);
        }
    }
}

fn list_paragraph_positions(list: &crate::list::List) -> Vec<Position> {
    let mut positions = Vec::new();
    for item in list.items() {
        positions.extend_from_slice(item.paragraph_positions());
        for nested in item.nested_lists() {
            positions.extend(list_paragraph_positions(nested));
        }
    }
    positions
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

fn validate_resource_readback(
    references: &[ResourceChange],
    payloads: &[ResourcePayloadChange],
    snapshot: &Template,
) -> Result<()> {
    validate_resource_references(references, snapshot, true)?;
    for change in payloads {
        if snapshot.package.member(&change.path)? != change.after {
            return Err(Error::InvalidFormat(
                "OTH resource payload publication failed readback".to_string(),
            ));
        }
        if change.after.is_some()
            && snapshot.package.member_media_type(&change.path)?.as_deref()
                != Some(change.media_type.as_str())
        {
            return Err(Error::InvalidFormat(
                "OTH resource payload media type failed readback".to_string(),
            ));
        }
    }
    if !payloads.is_empty() {
        validate_embedded_dependencies(snapshot)?;
    }
    Ok(())
}

fn validate_embedded_dependencies(snapshot: &Template) -> Result<()> {
    let files = snapshot.files()?;
    for resource in snapshot
        .package
        .resources()
        .iter()
        .filter(|resource| resource.is_embedded())
    {
        let root = embedded_path(resource.href())?;
        let prefix = format!("{root}/");
        if !files
            .iter()
            .any(|path| path == root || path.starts_with(&prefix))
        {
            return Err(Error::InvalidFormat(
                "OTH resource payload edit would leave a dangling embedded reference".to_string(),
            ));
        }
    }
    Ok(())
}

fn embedded_path(href: &str) -> Result<&str> {
    let path = href.strip_prefix("./").unwrap_or(href);
    if path.is_empty()
        || path.starts_with('/')
        || path.contains("://")
        || path.starts_with("data:")
        || path
            .split('/')
            .any(|segment| matches!(segment, "" | "." | ".."))
    {
        return Err(Error::InvalidFormat(
            "OTH resource payload path is not a safe package member".to_string(),
        ));
    }
    Ok(path)
}

fn resource_member_path(resource: &crate::resource::Resource, member: &str) -> Result<String> {
    let root = embedded_path(resource.href())?;
    let member_path = embedded_path(member)?;
    Ok(format!("{root}/{member_path}"))
}

fn list_paragraph_count(list: &crate::list::List) -> usize {
    list.items()
        .iter()
        .map(|item| {
            item.nested_lists()
                .iter()
                .map(list_paragraph_count)
                .fold(item.paragraphs().len(), usize::saturating_add)
        })
        .fold(0_usize, usize::saturating_add)
}

fn list_tree_count(list: &crate::list::List) -> usize {
    list.items()
        .iter()
        .flat_map(crate::list::Item::nested_lists)
        .map(list_tree_count)
        .fold(1_usize, usize::saturating_add)
}

fn normalize_list_level(list: &crate::list::List, level: usize) -> crate::list::List {
    let items = list
        .items()
        .iter()
        .map(|item| {
            crate::list::Item::projected(
                item.nested_lists()
                    .iter()
                    .map(|nested| normalize_list_level(nested, level.saturating_add(1)))
                    .collect(),
                item.paragraphs().to_vec(),
                item.paragraph_positions().to_vec(),
                item.start_value(),
            )
        })
        .collect();
    crate::list::List::projected(items, level, list.style_name().map(str::to_owned))
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
    validate_resource_readback(&edit.resource_changes, &edit.payload_changes, snapshot)?;
    let replaced_lists = edit
        .list_changes
        .iter()
        .filter_map(|change| change.before.as_ref())
        .map(list_tree_count)
        .fold(0_usize, usize::saturating_add);
    let replacement_lists = edit
        .list_changes
        .iter()
        .filter_map(|change| change.after.as_ref())
        .map(list_tree_count)
        .fold(0_usize, usize::saturating_add);
    let appended_lists = edit
        .appended
        .iter()
        .filter_map(|block| match block {
            crate::ContentBlock::List(list) => Some(list_tree_count(list)),
            crate::ContentBlock::Heading(_) | crate::ContentBlock::Paragraph(_) => None,
        })
        .fold(0_usize, usize::saturating_add);
    if snapshot.package.lists().len()
        != edit
            .source
            .package
            .lists()
            .len()
            .saturating_sub(replaced_lists)
            .saturating_add(replacement_lists)
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
    let inserted_resource_blocks = edit
        .resource_changes
        .iter()
        .filter(|change| change.before.is_none() && change.after.is_some())
        .count();
    let expected_order_len = edit
        .source
        .package
        .order()
        .len()
        .saturating_sub(replaced_block_count)
        .saturating_add(replacement_block_count)
        .saturating_add(appended_block_count)
        .saturating_add(inserted_resource_blocks);
    if snapshot.package.order().len() != expected_order_len {
        return Err(Error::InvalidFormat(
            "OTH structural edit failed block-order readback".to_string(),
        ));
    }
    Ok(())
}

fn published_durable_fragment(edit: &Edit<'_>, snapshot: &Template) -> Result<Option<String>> {
    if edit.durable_fragment.is_none() {
        return Ok(None);
    }
    let [block] = edit.appended.as_slice() else {
        return Err(Error::InvalidFormat(
            "OTH exact transfer requires one semantic appended block".to_string(),
        ));
    };
    let site = match block {
        crate::ContentBlock::Heading(_) => snapshot
            .package
            .heading_full_site(snapshot.package.headings().len().saturating_sub(1)),
        crate::ContentBlock::Paragraph(_) => snapshot
            .package
            .paragraph_full_site(snapshot.package.paragraphs().len().saturating_sub(1)),
        crate::ContentBlock::List(_) => snapshot
            .package
            .list_site(snapshot.package.lists().len().saturating_sub(1)),
    }
    .ok_or_else(|| Error::InvalidFormat("OTH exact transfer target site is missing".to_string()))?;
    snapshot
        .content_xml()
        .get(site.range.clone())
        .map(str::to_owned)
        .map(Some)
        .ok_or_else(|| {
            Error::InvalidFormat("OTH exact transfer target span is invalid".to_string())
        })
}

fn list_target_index(changes: &[ListChange], change: &ListChange) -> Result<usize> {
    let mut target = change.list.get();
    for candidate in changes
        .iter()
        .filter(|candidate| candidate.list.get() < change.list.get())
    {
        target = shift_list_index(target, candidate)?;
    }
    if change.before.is_some() {
        target = shift_list_index(target, change)?;
    }
    Ok(target)
}

fn shift_list_index(index: usize, change: &ListChange) -> Result<usize> {
    let before = change.before.as_ref().map_or(0, list_tree_count);
    let after = change.after.as_ref().map_or(0, list_tree_count);
    if after >= before {
        Ok(index.saturating_add(after - before))
    } else {
        index.checked_sub(before - after).ok_or_else(|| {
            Error::InvalidFormat("OTH replacement list target position underflow".to_string())
        })
    }
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
                    && actual_item.nested_lists().len() == expected_item.nested_lists().len()
                    && actual_item
                        .nested_lists()
                        .iter()
                        .zip(expected_item.nested_lists())
                        .all(|(actual_nested, expected_nested)| {
                            lists_semantically_equal(actual_nested, expected_nested)
                        })
            })
}

fn list_contains_paragraph(list: &crate::list::List, paragraph: Position) -> bool {
    list.items().iter().any(|item| {
        item.paragraph_positions().contains(&paragraph)
            || item
                .nested_lists()
                .iter()
                .any(|nested| list_contains_paragraph(nested, paragraph))
    })
}

fn replacement_sites_overlap(
    left: &crate::codec::ReplacementSite,
    right: &crate::codec::ReplacementSite,
) -> bool {
    left.range.start < right.range.end && right.range.start < left.range.end
}

fn list_changes_overlap(template: &Template, left: &ListChange, right: &ListChange) -> bool {
    match (
        template.package.list_site(left.list.get()),
        template.package.list_site(right.list.get()),
    ) {
        (Some(left_site), Some(right_site)) => replacement_sites_overlap(left_site, right_site),
        _ => left.list == right.list,
    }
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

fn push_wire_optional_bytes(output: &mut Vec<u8>, value: Option<&[u8]>) -> Result<()> {
    match value {
        None => push_wire_usize(output, 0),
        Some(bytes) => {
            push_wire_usize(output, 1)?;
            push_wire_bytes(output, bytes)
        },
    }
}

fn push_wire_resource(
    output: &mut Vec<u8>,
    resource: Option<&crate::resource::Resource>,
) -> Result<()> {
    let Some(resource_entry) = resource else {
        return push_wire_usize(output, 0);
    };
    push_wire_usize(output, 1)?;
    let kind = match resource_entry.kind() {
        crate::resource::Kind::Image => 0,
        crate::resource::Kind::Object => 1,
        crate::resource::Kind::OleObject => 2,
        crate::resource::Kind::Plugin => 3,
        crate::resource::Kind::FloatingFrame => 4,
    };
    push_wire_usize(output, kind)?;
    push_wire_bytes(output, resource_entry.href().as_bytes())
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
            let parsed = projection.lists.pop().ok_or_else(|| {
                Error::InvalidFormat("OTH durable patch list disappeared".to_string())
            })?;
            if parsed.level() != 1 || list_tree_count(&parsed) != projection.lists.len() + 1 {
                return Err(Error::InvalidFormat(
                    "OTH durable patch list fragment is not one top-level list".to_string(),
                ));
            }
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

fn read_wire_optional_bytes(bytes: &[u8], cursor: &mut usize) -> Result<Option<Vec<u8>>> {
    match read_wire_usize(bytes, cursor)? {
        0 => Ok(None),
        1 => Ok(Some(read_wire_bytes(bytes, cursor)?.to_vec())),
        _ => Err(Error::InvalidFormat(
            "OTH durable optional-byte marker is invalid".to_string(),
        )),
    }
}

fn read_wire_resource(
    bytes: &[u8],
    cursor: &mut usize,
) -> Result<Option<crate::resource::Resource>> {
    match read_wire_usize(bytes, cursor)? {
        0 => Ok(None),
        1 => {
            let kind = match read_wire_usize(bytes, cursor)? {
                0 => crate::resource::Kind::Image,
                1 => crate::resource::Kind::Object,
                2 => crate::resource::Kind::OleObject,
                3 => crate::resource::Kind::Plugin,
                4 => crate::resource::Kind::FloatingFrame,
                _ => {
                    return Err(Error::InvalidFormat(
                        "OTH durable resource kind is invalid".to_string(),
                    ));
                },
            };
            Ok(Some(crate::resource::Resource::new(
                kind,
                read_wire_string(bytes, cursor)?,
            )?))
        },
        _ => Err(Error::InvalidFormat(
            "OTH durable resource marker is invalid".to_string(),
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
        return Err(Error::InvalidFormat(format!(
            "OTH durable rich inline semantic readback failed for {:?}: XML matches={}, text matches={}",
            change.block,
            actual_xml == expected_xml,
            text == Some(expected_text)
        )));
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

fn read_resource_changes(
    bytes: &[u8],
    cursor: &mut usize,
    source: &Template,
    target: &Template,
) -> Result<Vec<ResourceChange>> {
    let count = read_wire_usize(bytes, cursor)?;
    if count
        > source
            .package
            .resources()
            .len()
            .saturating_add(target.package.resources().len())
    {
        return Err(Error::InvalidFormat(
            "OTH durable resource count is invalid".to_string(),
        ));
    }
    let mut changes = Vec::new();
    changes
        .try_reserve_exact(count)
        .map_err(|allocation_error| Error::Allocation {
            resource: "OTH durable resource changes",
            source: allocation_error,
        })?;
    for _ in 0..count {
        let change = ResourceChange {
            resource: Position::new(read_wire_usize(bytes, cursor)?),
            before: read_wire_resource(bytes, cursor)?,
            after: read_wire_resource(bytes, cursor)?,
            before_xml: read_wire_string(bytes, cursor)?,
            after_xml: read_wire_string(bytes, cursor)?,
        };
        if change.before.is_none() && change.after.is_none() {
            return Err(Error::InvalidFormat(
                "OTH durable resource change is empty".to_string(),
            ));
        }
        if changes
            .iter()
            .any(|candidate: &ResourceChange| candidate.resource == change.resource)
        {
            return Err(Error::InvalidFormat(
                "OTH durable resource selector is duplicated".to_string(),
            ));
        }
        changes.push(change);
    }
    validate_resource_references(&changes, source, false)?;
    validate_resource_references(&changes, target, true)?;
    Ok(changes)
}

fn read_payload_changes(
    bytes: &[u8],
    cursor: &mut usize,
    source: &Template,
    target: &Template,
) -> Result<Vec<ResourcePayloadChange>> {
    let count = read_wire_usize(bytes, cursor)?;
    let mut changes = Vec::new();
    changes
        .try_reserve_exact(count)
        .map_err(|allocation_error| Error::Allocation {
            resource: "OTH durable payload changes",
            source: allocation_error,
        })?;
    for _ in 0..count {
        let change = ResourcePayloadChange {
            path: read_wire_string(bytes, cursor)?,
            media_type: read_wire_string(bytes, cursor)?,
            before: read_wire_optional_bytes(bytes, cursor)?,
            after: read_wire_optional_bytes(bytes, cursor)?,
        };
        embedded_path(&change.path)?;
        if change.media_type.is_empty() || change.before == change.after {
            return Err(Error::InvalidFormat(
                "OTH durable payload change is invalid".to_string(),
            ));
        }
        if changes
            .iter()
            .any(|candidate: &ResourcePayloadChange| candidate.path == change.path)
        {
            return Err(Error::InvalidFormat(
                "OTH durable payload path is duplicated".to_string(),
            ));
        }
        if source.package.member(&change.path)? != change.before
            || target.package.member(&change.path)? != change.after
        {
            return Err(Error::InvalidFormat(
                "OTH durable payload readback failed".to_string(),
            ));
        }
        if change.after.is_some()
            && target.package.member_media_type(&change.path)?.as_deref()
                != Some(change.media_type.as_str())
        {
            return Err(Error::InvalidFormat(
                "OTH durable payload media type readback failed".to_string(),
            ));
        }
        changes.push(change);
    }
    if !changes.is_empty() {
        validate_embedded_dependencies(target)?;
    }
    Ok(changes)
}

fn validate_resource_references(
    changes: &[ResourceChange],
    template: &Template,
    after: bool,
) -> Result<()> {
    let removed = changes
        .iter()
        .filter(|change| change.before.is_some() && change.after.is_none())
        .count();
    let inserted = changes
        .iter()
        .filter(|change| change.before.is_none() && change.after.is_some())
        .count();
    let source_count = if after {
        template
            .package
            .resources()
            .len()
            .saturating_add(removed)
            .saturating_sub(inserted)
    } else {
        template.package.resources().len()
    };
    let expected_count = if after {
        source_count
            .saturating_sub(removed)
            .saturating_add(inserted)
    } else {
        source_count
    };
    if template.package.resources().len() != expected_count {
        return Err(Error::InvalidFormat(
            "OTH durable resource structural readback failed".to_string(),
        ));
    }
    for change in changes {
        let expected = if after {
            change.after.as_ref()
        } else {
            change.before.as_ref()
        };
        let Some(expected_resource) = expected else {
            continue;
        };
        let index = if after {
            resource_target_index(changes, change)
        } else {
            change.resource.get()
        };
        if template.package.resources().get(index) != Some(expected_resource) {
            return Err(Error::InvalidFormat(
                "OTH durable resource semantic readback failed".to_string(),
            ));
        }
        if after && expected_resource.is_embedded() {
            let root = embedded_path(expected_resource.href())?;
            let prefix = format!("{root}/");
            if !template
                .files()?
                .iter()
                .any(|path| path == root || path.starts_with(&prefix))
            {
                return Err(Error::InvalidFormat(
                    "OTH edited embedded resource dependency is absent".to_string(),
                ));
            }
        }
    }
    Ok(())
}

fn resource_target_index(changes: &[ResourceChange], change: &ResourceChange) -> usize {
    if change.before.is_none() {
        let removed = changes
            .iter()
            .filter(|candidate| candidate.before.is_some() && candidate.after.is_none())
            .count();
        return change.resource.get().saturating_sub(removed);
    }
    let removed_before = changes
        .iter()
        .filter(|candidate| {
            candidate.before.is_some()
                && candidate.after.is_none()
                && candidate.resource.get() < change.resource.get()
        })
        .count();
    change.resource.get().saturating_sub(removed_before)
}

fn read_list_changes(
    bytes: &[u8],
    cursor: &mut usize,
    source: &Template,
    target: &Template,
    appended_list_count: usize,
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
        let pending = ListChange {
            after,
            before,
            list,
        };
        if changes
            .iter()
            .any(|candidate| list_changes_overlap(source, candidate, &pending))
        {
            return Err(Error::InvalidFormat(
                "OTH durable list changes overlap".to_string(),
            ));
        }
        if let Some(expected) = pending.before.as_ref()
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
        changes.push(pending);
    }
    let removed = changes
        .iter()
        .filter_map(|change| change.before.as_ref())
        .map(list_tree_count)
        .fold(0_usize, usize::saturating_add);
    let inserted = changes
        .iter()
        .filter_map(|change| change.after.as_ref())
        .map(list_tree_count)
        .fold(0_usize, usize::saturating_add);
    if target.package.lists().len()
        != source
            .package
            .lists()
            .len()
            .saturating_sub(removed)
            .saturating_add(inserted)
            .saturating_add(appended_list_count)
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
    resource_changes: &[ResourceChange],
    list_changes: &[ListChange],
    appended: &[crate::ContentBlock],
    durable_fragment: Option<&str>,
) -> Result<String> {
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
                .saturating_add(resource_changes.len())
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
    let mut resource_tail = String::new();
    for change in resource_changes {
        if change.before.is_none() {
            if !change.before_xml.is_empty() {
                return Err(Error::InvalidFormat(
                    "OTH resource insertion precondition is invalid".to_string(),
                ));
            }
            resource_tail.push_str(&change.after_xml);
            continue;
        }
        let site = source
            .package
            .resource_site(change.resource.get())
            .cloned()
            .ok_or_else(|| {
                Error::InvalidFormat("OTH resource edit site disappeared".to_string())
            })?;
        let actual = source
            .content_xml()
            .get(site.range.clone())
            .ok_or_else(|| {
                Error::InvalidFormat("OTH resource edit source span is invalid".to_string())
            })?;
        if actual != change.before_xml {
            return Err(Error::InvalidFormat(
                "OTH resource edit source precondition failed".to_string(),
            ));
        }
        replacements.push((site, std::borrow::Cow::Borrowed(change.after_xml.as_str())));
    }
    if has_append || !resource_tail.is_empty() {
        let mut fragment = match durable_fragment {
            Some(fragment) => fragment.to_owned(),
            None => crate::authoring::render_fragment(appended)?,
        };
        fragment.push_str(&resource_tail);
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
