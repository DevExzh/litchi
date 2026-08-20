//! Unified source-checked ODP slide, rich-content, media, chart, design, annotation, and RDF edits.

use super::mutable::MutablePresentation;
use crate::core::OwnedPackage;
use crate::{Presentation, Reference, Shape, Slide};
use litchi_core::{Error, Result};
use std::fs::File;
use std::io::Read;
use std::path::Path;
use std::sync::Arc;
use xml_minifier::audit;

const MAX_PACKAGE_BYTES: usize = 128 * 1024 * 1024;
const MAX_DRAFT_BYTES: usize = 128 * 1024 * 1024;
const MAX_SLIDES: usize = 65_536;
const MAX_TEXT_BYTES: usize = 16 * 1024 * 1024;
const MAX_XML_PARTS: usize = 65_536;
const DURABLE_PATCH_MAGIC: &[u8; 16] = b"LITCHI-ODP-PATCH";
const DURABLE_PATCH_VERSION: u16 = 1;
const DURABLE_HISTORY_MAGIC: &[u8; 16] = b"LITCHI-ODP-HIST\0";
const DURABLE_HISTORY_VERSION: u16 = 1;
const MAX_DURABLE_HISTORY_BYTES: usize = 512 * 1024 * 1024;

/// Semantic dependency domain touched by a root package patch.
#[derive(Clone, Copy, Debug, PartialEq, Eq, Hash, PartialOrd, Ord)]
#[non_exhaustive]
pub enum Domain {
    /// Slide, shape, or embedded-media projection.
    Slides,
    /// RDF metadata graphs.
    Rdf,
    /// Embedded chart occurrences or parts.
    Charts,
    /// Presentation layouts, master pages, or their slide assignments.
    Design,
    /// Slide- or shape-anchored annotations.
    Annotations,
    /// Rich text boxes, tables, and inert form controls.
    Content,
}

/// Conservative non-mutating merge assessment for two package patches.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct MergePlan {
    conflicts: Vec<Domain>,
}

impl MergePlan {
    /// Return semantic domains requiring an explicit merge decision.
    #[must_use]
    pub fn conflicts(&self) -> &[Domain] {
        &self.conflicts
    }

    /// Return whether the two patches are provably independent at this API layer.
    #[must_use]
    pub fn is_independent(&self) -> bool {
        self.conflicts.is_empty()
    }
}

/// Mutation policy derived from package encryption and signature state.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
#[non_exhaustive]
pub enum SecurityPolicy {
    /// No encrypted entries or package signature parts were detected.
    Editable,
    /// Package entries are encrypted and must remain read-only in this editor.
    EncryptedReadOnly,
    /// Package signatures would be invalidated by mutation.
    SignedReadOnly,
    /// Both encrypted entries and package signatures are present.
    SignedAndEncryptedReadOnly,
}

impl SecurityPolicy {
    /// Return whether source-checked mutation and patch application are allowed.
    #[must_use]
    pub const fn is_editable(self) -> bool {
        matches!(self, Self::Editable)
    }
}

/// Explicit cryptographic package lifecycle request.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
#[non_exhaustive]
pub enum CryptoOperation {
    /// Cryptographically verify retained document or macro signatures.
    VerifySignatures,
    /// Add a new package signature.
    AddSignature,
    /// Remove retained signature owners.
    ClearSignatures,
    /// Encrypt package entries or change their password.
    Encrypt,
    /// Decrypt an editing snapshot in place.
    DecryptForEditing,
}

/// Stable reason a crypto lifecycle request is outside the ODP editor boundary.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
#[non_exhaustive]
pub enum CryptoRefusal {
    /// No cryptographic signature verifier is connected to this facade.
    SignatureVerificationUnavailable,
    /// No signing key/certificate owner is connected to this facade.
    SignatureAuthoringUnavailable,
    /// Removing signatures requires an explicit external unsigned-copy workflow.
    SignatureRemovalUnavailable,
    /// Package encryption writing and password changes are unavailable.
    EncryptionAuthoringUnavailable,
    /// Password opening must occur through the ordinary password-aware read facade.
    UsePasswordOpening,
}

/// Typed availability result for one cryptographic lifecycle operation.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
#[non_exhaustive]
pub enum CryptoCapability {
    /// The requested operation is already satisfied and publication is an exact no-op.
    AvailableNoOp,
    /// The operation is explicitly outside this facade for the retained package.
    Refused(CryptoRefusal),
}

/// Maximum package-media changes accepted by one atomic transaction batch.
pub const MAX_MEDIA_CHANGES: usize = 256;

/// One bounded change to an inert package-contained ODP media member.
///
/// Media is never opened, decoded, played, or fetched.  Replacements retain
/// every existing `xlink:href` owner and only change the exact package member;
/// removals are accepted only when no retained XML owner can reference the
/// member.  All changes in one slice are preflighted before the transaction is
/// modified and publish through the transaction's single package commit.
#[derive(Clone, Debug, PartialEq, Eq)]
#[non_exhaustive]
pub enum MediaChange {
    /// Add a new package member at an unused path.
    Add {
        /// Safe package-relative member path.
        path: String,
        /// Opaque payload bytes.
        payload: Vec<u8>,
        /// Manifest media type.
        media_type: String,
    },
    /// Replace an existing package member without changing its path.
    Replace {
        /// Existing package-relative member path.
        path: String,
        /// Opaque replacement payload bytes.
        payload: Vec<u8>,
        /// Manifest media type.  Existing members retain their declared type.
        media_type: String,
    },
    /// Remove an unreferenced package member.
    Remove {
        /// Existing package-relative member path.
        path: String,
    },
}

impl MediaChange {
    /// Construct an add operation from owned or borrowed-compatible values.
    #[must_use]
    pub fn add(
        path: impl Into<String>,
        payload: impl Into<Vec<u8>>,
        media_type: impl Into<String>,
    ) -> Self {
        Self::Add {
            path: path.into(),
            payload: payload.into(),
            media_type: media_type.into(),
        }
    }

    /// Construct a replacement operation from owned or borrowed-compatible values.
    #[must_use]
    pub fn replace(
        path: impl Into<String>,
        payload: impl Into<Vec<u8>>,
        media_type: impl Into<String>,
    ) -> Self {
        Self::Replace {
            path: path.into(),
            payload: payload.into(),
            media_type: media_type.into(),
        }
    }

    /// Construct a removal operation.
    #[must_use]
    pub fn remove(path: impl Into<String>) -> Self {
        Self::Remove { path: path.into() }
    }

    /// Borrow the path targeted by this change.
    #[must_use]
    pub fn path(&self) -> &str {
        match self {
            Self::Add { path, .. } | Self::Replace { path, .. } | Self::Remove { path } => path,
        }
    }
}

impl CryptoCapability {
    /// Borrow the refusal reason, if any.
    #[must_use]
    pub const fn refusal(self) -> Option<CryptoRefusal> {
        match self {
            Self::AvailableNoOp => None,
            Self::Refused(reason) => Some(reason),
        }
    }
}

/// An immutable presentation package and its parsed slide projection.
#[derive(Clone)]
pub struct Snapshot {
    bytes: Arc<Vec<u8>>,
    /// The validated physical package retained with this immutable snapshot.
    ///
    /// `OwnedPackage` keeps its ZIP index behind an `Arc`, so semantic
    /// reopenings of the same immutable artifact can share the validated
    /// archive state instead of rebuilding the central directory from the
    /// same bytes. The separate byte owner remains authoritative for exact
    /// lineage and patch source checks.
    package: OwnedPackage,
    resource_bytes: usize,
    slides: Arc<[Slide]>,
}

impl Snapshot {
    /// Open an ODP editing snapshot from a path.
    ///
    /// # Errors
    ///
    /// Returns an error when the file cannot be read or is not a bounded valid ODP package.
    pub fn open(path: impl AsRef<Path>) -> Result<Self> {
        Self::from_bytes(read_bounded(path.as_ref(), MAX_PACKAGE_BYTES)?)
    }

    /// Parse owned ODP package bytes into a source-bound snapshot.
    ///
    /// # Errors
    ///
    /// Returns an error for an oversized, malformed, or non-ODP package.
    /// Signed and encrypted packages can be inspected but have a read-only
    /// [`SecurityPolicy`] and cannot start a transaction.
    pub fn from_bytes(source_bytes: Vec<u8>) -> Result<Self> {
        if source_bytes.len() > MAX_PACKAGE_BYTES {
            return invalid("ODP editing package exceeds the 128 MiB limit");
        }
        Self::from_shared_bytes(Arc::new(source_bytes))
    }

    pub(crate) fn from_shared_bytes(bytes: Arc<Vec<u8>>) -> Result<Self> {
        if bytes.len() > MAX_PACKAGE_BYTES {
            return invalid("ODP editing package exceeds the 128 MiB limit");
        }
        let package = OwnedPackage::from_shared_bytes(Arc::clone(&bytes))?;
        Self::from_owned_package(package)
    }

    pub(crate) fn from_owned_package(package: OwnedPackage) -> Result<Self> {
        validate_package_size(package.as_bytes().len())?;
        let presentation = Presentation::from_owned_package(package)?;
        let slides = presentation.slides()?;
        if slides.len() > MAX_SLIDES {
            return invalid("ODP editing snapshot exceeds the slide-count limit");
        }
        let resource_bytes = slides_resource(&slides)?;
        if resource_bytes > MAX_DRAFT_BYTES {
            return invalid("ODP editing snapshot exceeds the aggregate draft limit");
        }
        let package = presentation.owned_package().clone_without_password();
        let bytes = package.shared_bytes();
        Ok(Self {
            bytes,
            package,
            resource_bytes,
            slides: Arc::from(slides),
        })
    }

    /// Borrow the exact source package bytes.
    #[must_use]
    pub fn bytes(&self) -> &[u8] {
        self.bytes.as_slice()
    }

    /// Borrow the immutable parsed slide projection.
    #[must_use]
    pub fn slides(&self) -> &[Slide] {
        &self.slides
    }

    /// Inspect the package mutation policy without staging an edit.
    ///
    /// # Errors
    ///
    /// Returns an error when the retained package cannot be reopened.
    pub fn security_policy(&self) -> Result<SecurityPolicy> {
        package_security_policy(&self.package)
    }

    /// Resolve one crypto lifecycle request without mutating the package.
    ///
    /// Signing, signature verification/removal, and encryption writing are
    /// deliberately final-scoped as typed refusals until a key/certificate and
    /// cryptographic verification owner is connected to the public ODP facade.
    ///
    /// # Errors
    ///
    /// Returns an error when the retained package cannot be reopened.
    pub fn crypto_capability(&self, operation: CryptoOperation) -> Result<CryptoCapability> {
        let policy = self.security_policy()?;
        Ok(match operation {
            CryptoOperation::VerifySignatures => {
                CryptoCapability::Refused(CryptoRefusal::SignatureVerificationUnavailable)
            },
            CryptoOperation::AddSignature => {
                CryptoCapability::Refused(CryptoRefusal::SignatureAuthoringUnavailable)
            },
            CryptoOperation::ClearSignatures => match policy {
                SecurityPolicy::SignedReadOnly | SecurityPolicy::SignedAndEncryptedReadOnly => {
                    CryptoCapability::Refused(CryptoRefusal::SignatureRemovalUnavailable)
                },
                SecurityPolicy::Editable | SecurityPolicy::EncryptedReadOnly => {
                    CryptoCapability::AvailableNoOp
                },
            },
            CryptoOperation::Encrypt => {
                CryptoCapability::Refused(CryptoRefusal::EncryptionAuthoringUnavailable)
            },
            CryptoOperation::DecryptForEditing => match policy {
                SecurityPolicy::EncryptedReadOnly | SecurityPolicy::SignedAndEncryptedReadOnly => {
                    CryptoCapability::Refused(CryptoRefusal::UsePasswordOpening)
                },
                SecurityPolicy::Editable | SecurityPolicy::SignedReadOnly => {
                    CryptoCapability::AvailableNoOp
                },
            },
        })
    }

    /// Read arbitrary source-backed story/list, table, and inert form owners.
    ///
    /// This remains available for signed read-only snapshots because it never
    /// stages or publishes a mutation.
    ///
    /// # Errors
    ///
    /// Returns an error for malformed, ambiguous, or oversized content owners.
    pub fn rich_content(&self) -> Result<crate::content::Inventory> {
        crate::content::inventory(&self.package)
    }

    /// Select a slide by checked zero-based position or exact title.
    ///
    /// # Errors
    ///
    /// Returns an error when a title matches more than one slide.
    pub fn slide<'a, S>(&self, selector: S) -> Result<Option<&Slide>>
    where
        S: Into<Selector<'a>>,
    {
        select(&self.slides, selector.into())
            .map(|selected| selected.map(|position| &self.slides[position]))
    }

    /// Start an isolated transaction over a detached staging engine.
    ///
    /// # Errors
    ///
    /// Returns an error when the exact source package cannot be reparsed for staging or
    /// its signature/encryption policy makes it read-only.
    pub fn transaction(&self) -> Result<Transaction> {
        let presentation = Presentation::from_owned_package(self.package.clone())?;
        ensure_editable_source(presentation.owned_package())?;
        Ok(Transaction {
            source: self.clone(),
            draft: MutablePresentation::from_presentation_with_validated_slides(
                &presentation,
                self.slides(),
            )?,
            changed: false,
            rdf: None,
            charts: None,
            design: None,
            annotations: None,
            content: None,
            media_bytes: 0,
            resource_bytes: self.resource_bytes,
            source_resource_bytes: self.resource_bytes,
            slide_order_changed: false,
            dependency_free_slide_copy_changed: false,
            dependency_free_slide_removal_changed: false,
        })
    }

    /// Materialize this snapshot as the ordinary read facade.
    ///
    /// # Errors
    ///
    /// Returns an error when the retained package can no longer be parsed.
    pub fn to_presentation(&self) -> Result<Presentation> {
        Presentation::from_owned_package(self.package.clone())
    }

    /// Return the identity of the validated archive index retained by this
    /// immutable snapshot.
    #[doc(hidden)]
    #[must_use]
    pub fn prepared_index_identity(&self) -> usize {
        self.package.prepared_index_identity()
    }
}

/// A semantic slide selector.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum Selector<'a> {
    /// Checked zero-based position.
    Index(usize),
    /// Exact slide title. Duplicate titles are an ambiguity error.
    Title(&'a str),
}

/// Final position of a slide after a checked move.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[non_exhaustive]
pub enum SlidePosition {
    /// Move the selected slide to the first position.
    First,
    /// Move the selected slide to the last position.
    Last,
    /// Move the selected slide to this final zero-based position.
    Index(usize),
}

impl From<usize> for Selector<'_> {
    fn from(value: usize) -> Self {
        Self::Index(value)
    }
}

impl<'a> From<&'a str> for Selector<'a> {
    fn from(value: &'a str) -> Self {
        Self::Title(value)
    }
}

/// An isolated mutable draft that cannot alter its source snapshot.
pub struct Transaction {
    source: Snapshot,
    draft: MutablePresentation,
    changed: bool,
    rdf: Option<RdfDraft>,
    charts: Option<ChartDraft>,
    design: Option<DesignDraft>,
    annotations: Option<AnnotationDraft>,
    content: Option<ContentDraft>,
    media_bytes: usize,
    resource_bytes: usize,
    source_resource_bytes: usize,
    slide_order_changed: bool,
    dependency_free_slide_copy_changed: bool,
    dependency_free_slide_removal_changed: bool,
}

#[derive(Clone)]
enum RdfOperation {
    AddGraph {
        path: String,
        triples: Vec<crate::rdf::Triple>,
    },
    ReplaceGraph {
        path: String,
        triples: Vec<crate::rdf::Triple>,
    },
    RemoveGraph {
        path: String,
    },
    AddTriple {
        path: String,
        triple: crate::rdf::Triple,
    },
    ReplaceTriple {
        path: String,
        index: usize,
        triple: crate::rdf::Triple,
    },
    RemoveTriple {
        path: String,
        index: usize,
    },
    MoveTriple {
        path: String,
        from: usize,
        to: usize,
    },
}

struct RdfDraft {
    package: OwnedPackage,
    original_graphs: Vec<crate::rdf::Graph>,
    graphs: Vec<crate::rdf::Graph>,
    operations: Vec<RdfOperation>,
}

#[derive(Clone)]
enum ChartSelector {
    Index(usize),
    Name(String),
}

#[derive(Clone)]
enum ChartPage {
    Index(usize),
    Name(String),
}

impl ChartSelector {
    fn from_borrowed(selector: crate::charts::Selector<'_>) -> Self {
        match selector {
            crate::charts::Selector::Index(index) => Self::Index(index),
            crate::charts::Selector::Name(name) => Self::Name(name.to_string()),
        }
    }

    fn borrowed(&self) -> crate::charts::Selector<'_> {
        match self {
            Self::Index(index) => crate::charts::Selector::Index(*index),
            Self::Name(name) => crate::charts::Selector::Name(name),
        }
    }
}

impl ChartPage {
    fn from_borrowed(page: crate::charts::Page<'_>) -> Self {
        match page {
            crate::charts::Page::Index(index) => Self::Index(index),
            crate::charts::Page::Name(name) => Self::Name(name.to_string()),
        }
    }

    fn borrowed(&self) -> crate::charts::Page<'_> {
        match self {
            Self::Index(index) => crate::charts::Page::Index(*index),
            Self::Name(name) => crate::charts::Page::Name(name),
        }
    }
}

#[derive(Clone)]
enum ChartOperation {
    Replace {
        selector: ChartSelector,
        part: crate::charts::Part,
    },
    Remove {
        selector: ChartSelector,
    },
    Add {
        page: ChartPage,
        name: String,
        storage: crate::charts::Storage,
        part: crate::charts::Part,
    },
}

struct ChartDraft {
    package: OwnedPackage,
    original: Vec<crate::charts::Chart>,
    charts: Vec<crate::charts::Chart>,
    operations: Vec<ChartOperation>,
    limits: crate::charts::Limits,
}

#[derive(Clone)]
enum DesignOperation {
    AddLayout(crate::layout::Layout),
    ReplaceLayout(crate::layout::Layout),
    RemoveLayout {
        name: String,
        replacement: Option<String>,
    },
    ReorderLayouts(Vec<String>),
    AddMaster(crate::MasterPage),
    ReplaceMaster(crate::MasterPage),
    RemoveMaster {
        name: String,
        replacement: Option<String>,
    },
    ReorderMasters(Vec<String>),
    AssignSlideMaster {
        slide_index: usize,
        name: Option<String>,
    },
    AssignSlideLayout {
        slide_index: usize,
        name: Option<String>,
    },
}

struct DesignDraft {
    package: OwnedPackage,
    original_layouts: crate::layout::Collection,
    layouts: crate::layout::Collection,
    original_masters: Vec<crate::MasterPage>,
    masters: Vec<crate::MasterPage>,
    original_pages: crate::page::Collection,
    pages: crate::page::Collection,
    operations: Vec<DesignOperation>,
}

#[derive(Clone)]
enum AnnotationOperation {
    Add {
        anchor: crate::annotation::Anchor,
        annotation: crate::annotation::Annotation,
    },
    Replace {
        index: usize,
        annotation: crate::annotation::Annotation,
    },
    Remove {
        index: usize,
    },
}

struct AnnotationDraft {
    package: OwnedPackage,
    original: Vec<crate::annotation::Info>,
    annotations: Vec<crate::annotation::Info>,
    operations: Vec<AnnotationOperation>,
}

struct ContentDraft {
    bytes: Arc<Vec<u8>>,
    package: OwnedPackage,
    operations: Vec<crate::content::Operation>,
}

impl Transaction {
    /// Borrow the current staged slide projection.
    #[must_use]
    pub fn slides(&self) -> &[Slide] {
        self.draft.slides()
    }

    /// Return the identity of the validated archive index retained by the
    /// transaction's immutable source snapshot.
    #[doc(hidden)]
    #[must_use]
    pub fn prepared_index_identity(&self) -> usize {
        self.source.prepared_index_identity()
    }

    /// Append a compact title/body slide.
    ///
    /// # Errors
    ///
    /// Returns an error for invalid text, exhausted limits, or ambiguous retained bindings.
    pub fn add(&mut self, title: &str, text: &str) -> Result<()> {
        Self::check_text(title, text)?;
        self.check_no_slide_order_change("slide insertion")?;
        self.check_structure_edit()?;
        if self.draft.slides().len() == MAX_SLIDES {
            return invalid("ODP transaction exceeds the slide-count limit");
        }
        let candidate = self.resource_candidate(0, text_resource(title, text)?)?;
        self.draft
            .insert_slide(self.draft.slides().len(), title, text)?;
        self.resource_bytes = candidate;
        self.changed = true;
        Ok(())
    }

    /// Insert a slide before a selected source-snapshot anchor.
    ///
    /// # Errors
    ///
    /// Returns an error for an ambiguous selector, invalid text, exhausted limits, or retained
    /// declarations whose identity closure cannot be updated losslessly.
    pub fn add_before<'a, S>(&mut self, selector: S, title: &str, text: &str) -> Result<Option<()>>
    where
        S: Into<Selector<'a>>,
    {
        Self::check_text(title, text)?;
        self.check_no_slide_order_change("slide insertion")?;
        self.check_structure_edit()?;
        if self.draft.slides().len() == MAX_SLIDES {
            return invalid("ODP transaction exceeds the slide-count limit");
        }
        let Some(index) = select(self.draft.slides(), selector.into())? else {
            return Ok(None);
        };
        let candidate = self.resource_candidate(0, text_resource(title, text)?)?;
        self.draft.insert_slide(index, title, text)?;
        self.resource_bytes = candidate;
        self.changed = true;
        Ok(Some(()))
    }

    /// Replace one supported slide's title and body text.
    ///
    /// A pristine page retained from an opened package is refused because its
    /// unmodelled children cannot yet be proven lossless under regeneration.
    ///
    /// # Errors
    ///
    /// Returns an error for an ambiguous selector, invalid text, or a preservation-only page.
    pub fn replace<'a, S>(&mut self, selector: S, title: &str, text: &str) -> Result<Option<()>>
    where
        S: Into<Selector<'a>>,
    {
        self.check_no_slide_order_change("slide replacement")?;
        Self::check_text(title, text)?;
        let Some(index) = select(self.draft.slides(), selector.into())? else {
            return Ok(None);
        };
        let slide = &self.draft.slides()[index];
        if slide.title.as_deref() == Some(title) && slide.text == text {
            return Ok(Some(()));
        }
        self.check_slide_rewrite(index)?;
        let removed = slide_primary_text_resource(slide)?;
        let candidate = self.resource_candidate(removed, text_resource(title, text)?)?;
        self.draft.update_slide(index, title, text)?;
        self.resource_bytes = candidate;
        self.changed = true;
        Ok(Some(()))
    }

    /// Remove one selected slide and return its staged semantic value.
    ///
    /// # Errors
    ///
    /// Returns an error for an ambiguous selector or an unresolved page/declaration reference.
    pub fn remove<'a, S>(&mut self, selector: S) -> Result<Option<Slide>>
    where
        S: Into<Selector<'a>>,
    {
        self.check_no_slide_order_change("slide removal")?;
        self.check_structure_edit()?;
        let Some(index) = select(self.draft.slides(), selector.into())? else {
            return Ok(None);
        };
        let removed_bytes = slide_resource(&self.draft.slides()[index])?;
        let candidate = self.resource_candidate(removed_bytes, 0)?;
        let removed = self.draft.remove_slide(index)?;
        self.resource_bytes = candidate;
        self.changed = true;
        Ok(Some(removed))
    }

    /// Move one selected slide to a checked final position in the staged deck.
    ///
    /// Both the selector and [`SlidePosition::Index`] use the transaction's
    /// current staged sequence. The destination is the slide's final index
    /// after the move, so moving index `0` to index `2` in a three-slide deck
    /// produces `[1, 2, 0]`. Page names, raw producer page fragments, notes,
    /// transitions, animations, shapes, annotations, and resource references
    /// travel with the selected slide. Sources with retained declarations or
    /// slide-show settings are conservatively refused until those producer
    /// fragments can be preserved verbatim.
    ///
    /// A move cannot be mixed with already-staged page-indexed chart, design,
    /// annotation, or rich-content operations: their coordinate replay would
    /// otherwise be ambiguous. The refusal leaves the transaction unchanged.
    /// RDF and package-media operations are position-independent and may be
    /// composed with a move.
    ///
    /// # Errors
    ///
    /// Returns an error for an ambiguous selector, an out-of-range destination,
    /// or a conflicting page-indexed operation. A missing selector returns
    /// `Ok(None)`. Moving to the current position is an exact no-op.
    pub fn move_slide<'a, S>(&mut self, selector: S, position: SlidePosition) -> Result<Option<()>>
    where
        S: Into<Selector<'a>>,
    {
        if self.dependency_free_slide_copy_changed || self.dependency_free_slide_removal_changed {
            return unsupported(
                "ODP slide move cannot be staged after a dependency-free blank-slide copy or removal",
            );
        }
        let Some(from) = select(self.draft.slides(), selector.into())? else {
            return Ok(None);
        };
        let to = match position {
            SlidePosition::First => 0,
            SlidePosition::Last => self.draft.slides().len() - 1,
            SlidePosition::Index(index) if index < self.draft.slides().len() => index,
            SlidePosition::Index(index) => {
                return invalid(format!(
                    "ODP slide move destination {index} exceeds slide count {}",
                    self.draft.slides().len()
                ));
            },
        };
        if from == to {
            return Ok(Some(()));
        }
        self.check_page_indexed_move_conflict()?;
        self.draft.check_slide_move_supported()?;
        self.draft.move_slide(from, to)?;
        self.slide_order_changed = true;
        self.changed = true;
        Ok(Some(()))
    }

    /// Append an exact-fragment copy of one dependency-free blank source slide.
    ///
    /// This is intentionally not a general slide-copy API. It accepts only a
    /// compact, self-closing `draw:page` whose sole non-namespace attribute is
    /// `draw:name`. Consequently the source has no slide-local content, style,
    /// master, layout, identifier/navigation, hyperlink, event, script, MCE,
    /// protection, or opaque dependency closure to duplicate. The source page
    /// fragment is preserved byte-for-byte except for a deterministic unique
    /// name (`" Copy"`, then `" Copy 2"`, and so on).
    ///
    /// The operation is append-only. A transaction may contain at most one
    /// dependency-free blank-slide copy and cannot combine it with other slide
    /// or page-indexed operations. RDF edits remain position-independent and
    /// may be composed. A missing selector returns `Ok(None)`.
    ///
    /// # Errors
    ///
    /// Returns an error for an ambiguous selector, a non-retained/non-compact
    /// page, any dependency-bearing page markup, a read-only source, or an
    /// exhausted slide, name, fragment, draft, or package bound. Refusal leaves
    /// the transaction unchanged.
    pub fn copy_dependency_free_blank_slide<'a, S>(&mut self, selector: S) -> Result<Option<usize>>
    where
        S: Into<Selector<'a>>,
    {
        if self.dependency_free_slide_copy_changed {
            return unsupported(
                "ODP transaction already contains a dependency-free blank-slide copy",
            );
        }
        if self.slide_order_changed {
            return unsupported(
                "ODP dependency-free blank-slide copy cannot be staged after a slide move",
            );
        }
        let Some(index) = select(self.draft.slides(), selector.into())? else {
            return Ok(None);
        };
        if self.changed || self.has_page_indexed_operations() {
            return unsupported(
                "ODP dependency-free blank-slide copy cannot combine with slide or page-indexed operations",
            );
        }
        if self.draft.slides().len() == MAX_SLIDES {
            return invalid("ODP transaction exceeds the slide-count limit");
        }
        let copy = self.draft.prepare_dependency_free_blank_slide_copy(index)?;
        let candidate = self.resource_candidate(0, copy.resource_bytes())?;
        let copied_index = self.draft.apply_dependency_free_blank_slide_copy(copy)?;
        self.resource_bytes = candidate;
        self.dependency_free_slide_copy_changed = true;
        self.changed = true;
        Ok(Some(copied_index))
    }

    /// Append one dependency-free blank slide from an independent ODP
    /// snapshot.
    ///
    /// This deliberately admits only a self-closing donor `draw:page` whose
    /// sole semantic attribute is `draw:name`. The donor page has no body,
    /// layout, master, identifier, navigation, hyperlink, animation, macro,
    /// protection, or package-resource closure to remap. Its name is remapped
    /// deterministically against the destination page names and the exact
    /// donor fragment is retained as a destination-owned page override.
    ///
    /// The destination must already contain a retained source presentation
    /// body and cannot have staged slide/page-indexed operations, declarations,
    /// settings, macros, signatures, or encryption. The donor is never
    /// modified. A missing selector is an exact no-op.
    ///
    /// # Errors
    ///
    /// Returns an error for an unsupported donor or destination package,
    /// ambiguous selector, dependency-bearing page, exhausted limits, or a
    /// read-only source. Refusal leaves the transaction unchanged.
    pub fn transfer_dependency_free_blank_slide_from<'a, S>(
        &mut self,
        source: &Snapshot,
        selector: S,
    ) -> Result<Option<usize>>
    where
        S: Into<Selector<'a>>,
    {
        if self.dependency_free_slide_copy_changed {
            return unsupported(
                "ODP transaction already contains a dependency-free blank-slide copy",
            );
        }
        if self.slide_order_changed {
            return unsupported(
                "ODP foreign blank-slide transfer cannot be staged after a slide move",
            );
        }
        let Some(source_index) = select(source.slides(), selector.into())? else {
            return Ok(None);
        };
        if self.dependency_free_slide_removal_changed {
            return unsupported(
                "ODP foreign blank-slide transfer cannot be staged after a slide removal",
            );
        }
        if self.changed || self.has_page_indexed_operations() {
            return unsupported(
                "ODP foreign blank-slide transfer cannot combine with slide or page-indexed operations",
            );
        }
        if self.draft.slides().len() == MAX_SLIDES {
            return invalid("ODP transaction exceeds the slide-count limit");
        }

        let source_package = source.package.clone();
        ensure_editable_source(&source_package)?;
        super::mutable::ensure_foreign_transfer_package_safety(&source_package, "donor")?;
        let source_presentation = Presentation::from_owned_package(source.package.clone())?;
        audit::verify_authored(
            source_presentation.content_xml().as_bytes(),
            audit::Limits::default(),
        )
        .map(|_report| ())
        .map_err(|error| {
            Error::Unsupported(format!(
                "ODP foreign blank-slide transfer refuses unaudited donor XML: {error}"
            ))
        })?;
        if source_presentation.settings()?.is_some()
            || !source_presentation.declarations()?.is_empty()
        {
            return unsupported(
                "ODP foreign blank-slide transfer refuses donor settings or declarations",
            );
        }
        let source_pages = source_presentation.pages()?;
        let source_page = source_pages.page(source_index).ok_or_else(|| {
            invalid_error("ODP foreign blank-slide transfer donor page metadata is missing")
        })?;
        let source_name = source_page.name.as_deref().ok_or_else(|| {
            Error::Unsupported(
                "ODP foreign blank-slide transfer donor requires draw:name".to_string(),
            )
        })?;
        let source_content =
            crate::codec::content_source::ContentSource::parse(source_presentation.content_xml())?
                .ok_or_else(|| {
                    Error::Unsupported(
                "ODP foreign blank-slide transfer requires retained donor content fragments"
                    .to_string(),
            )
                })?;
        let donor_page = source_content.page(source_index).ok_or_else(|| {
            invalid_error("ODP foreign blank-slide transfer donor page is outside content coverage")
        })?;
        let copy = self
            .draft
            .prepare_foreign_dependency_free_blank_slide_copy(
                donor_page,
                &source.slides()[source_index],
                source_name,
            )?;
        let candidate = self.resource_candidate(0, copy.resource_bytes())?;
        let copied_index = self
            .draft
            .apply_foreign_dependency_free_blank_slide_copy(copy)?;
        self.resource_bytes = candidate;
        self.dependency_free_slide_copy_changed = true;
        self.changed = true;
        Ok(Some(copied_index))
    }

    /// Remove one exact retained dependency-free blank source slide.
    ///
    /// This is intentionally not a general slide-removal API. It accepts only
    /// a compact self-closing `draw:page` whose sole non-namespace attribute is
    /// `draw:name`. It refuses the final slide, retained declarations/settings,
    /// package or content macro owners, copied pages, inbound name-bearing XML
    /// attributes or fragment hyperlinks, and every dependency-bearing
    /// selected-page construct.
    /// Other package members and unselected page fragments are retained exactly
    /// through the ordinary source-backed writer.
    ///
    /// A transaction may contain at most one dependency-free removal and cannot
    /// combine it with other slide or page-indexed operations. A missing
    /// selector returns `Ok(None)` without staging.
    ///
    /// # Errors
    ///
    /// Returns an error for an ambiguous selector, the final slide, any unsafe
    /// dependency owner, a non-retained/non-compact page, or a resource bound.
    /// Refusal leaves the transaction unchanged.
    pub fn remove_dependency_free_blank_slide<'a, S>(
        &mut self,
        selector: S,
    ) -> Result<Option<Slide>>
    where
        S: Into<Selector<'a>>,
    {
        if self.dependency_free_slide_removal_changed {
            return unsupported(
                "ODP transaction already contains a dependency-free blank-slide removal",
            );
        }
        if self.slide_order_changed || self.dependency_free_slide_copy_changed {
            return unsupported(
                "ODP dependency-free blank-slide removal cannot follow a slide move or copy",
            );
        }
        let Some(index) = select(self.draft.slides(), selector.into())? else {
            return Ok(None);
        };
        if self.changed || self.has_page_indexed_operations() {
            return unsupported(
                "ODP dependency-free blank-slide removal cannot combine with slide or page-indexed operations",
            );
        }
        let removed_bytes = slide_resource(&self.draft.slides()[index])?;
        let candidate = self.resource_candidate(removed_bytes, 0)?;
        let removal = self
            .draft
            .prepare_dependency_free_blank_slide_removal(index)?;
        let removed = self
            .draft
            .apply_dependency_free_blank_slide_removal(removal);
        self.resource_bytes = candidate;
        self.dependency_free_slide_removal_changed = true;
        self.changed = true;
        Ok(Some(removed))
    }

    /// Append a typed shape to one supported slide.
    ///
    /// Hyperlinks, actions, event bindings, and media references remain inert
    /// metadata; this operation never follows or executes them.
    ///
    /// # Errors
    ///
    /// Returns an error for an ambiguous selector, invalid shape, or preservation-only page.
    pub fn add_shape<'a, S>(&mut self, selector: S, shape: Shape) -> Result<Option<()>>
    where
        S: Into<Selector<'a>>,
    {
        self.check_no_slide_order_change("shape insertion")?;
        let Some(index) = select(self.draft.slides(), selector.into())? else {
            return Ok(None);
        };
        self.check_slide_rewrite(index)?;
        let candidate = self.resource_candidate(0, shape_resource(&shape)?)?;
        super::builder::Builder::generate_shape_xml(
            &shape,
            self.draft.slides()[index].shapes.len(),
        )?;
        self.draft.add_shape(index, shape)?;
        self.resource_bytes = candidate;
        self.changed = true;
        Ok(Some(()))
    }

    /// Remove a shape by checked zero-based position from a selected slide.
    ///
    /// # Errors
    ///
    /// Returns an error for an ambiguous selector or preservation-only page.
    pub fn remove_shape<'a, S>(&mut self, selector: S, shape_index: usize) -> Result<Option<Shape>>
    where
        S: Into<Selector<'a>>,
    {
        self.check_no_slide_order_change("shape removal")?;
        let Some(slide_index) = select(self.draft.slides(), selector.into())? else {
            return Ok(None);
        };
        if shape_index >= self.draft.slides()[slide_index].shapes.len() {
            return Ok(None);
        }
        self.check_slide_rewrite(slide_index)?;
        let removed_bytes = shape_resource(&self.draft.slides()[slide_index].shapes[shape_index])?;
        let candidate = self.resource_candidate(removed_bytes, 0)?;
        let removed = self.draft.remove_shape(slide_index, shape_index)?;
        self.resource_bytes = candidate;
        self.changed = true;
        Ok(Some(removed))
    }

    /// Clear the title, body, and shapes of one supported slide.
    ///
    /// # Errors
    ///
    /// Returns an error for an ambiguous selector or preservation-only page.
    pub fn clear<'a, S>(&mut self, selector: S) -> Result<Option<()>>
    where
        S: Into<Selector<'a>>,
    {
        self.check_no_slide_order_change("slide clear")?;
        let Some(index) = select(self.draft.slides(), selector.into())? else {
            return Ok(None);
        };
        let slide = &self.draft.slides()[index];
        if slide.title.is_none() && slide.text.is_empty() && slide.shapes.is_empty() {
            return Ok(Some(()));
        }
        self.check_slide_rewrite(index)?;
        let removed = slide_primary_resource(slide)?;
        let candidate = self.resource_candidate(removed, 0)?;
        self.draft.clear_slide(index)?;
        self.resource_bytes = candidate;
        self.changed = true;
        Ok(Some(()))
    }

    /// Add bounded package-contained media and return its inert reference.
    ///
    /// # Errors
    ///
    /// Returns an error for an invalid or colliding path, media type, or oversized payload.
    pub fn embed_media(
        &mut self,
        path: &str,
        payload: &[u8],
        media_type: &str,
    ) -> Result<Reference> {
        crate::model::media::validate_package_media_path(path)?;
        let addition = path
            .len()
            .checked_add(media_type.len())
            .and_then(|value| value.checked_add(payload.len()))
            .ok_or_else(|| invalid_error("ODP embedded media resource size overflow"))?;
        let media_bytes = self
            .media_bytes
            .checked_add(addition)
            .ok_or_else(|| invalid_error("ODP aggregate media size overflow"))?;
        self.check_projected(self.resource_bytes, media_bytes)?;
        let reference = self.draft.embed_media(
            try_owned_str(path, "ODP media path")?,
            try_owned_bytes(payload, "ODP media payload")?,
            try_owned_str(media_type, "ODP media type")?,
        )?;
        self.media_bytes = media_bytes;
        self.changed = true;
        Ok(reference)
    }

    /// Apply a bounded atomic batch of inert package-media changes.
    ///
    /// Additions use new paths, replacements preserve all existing XML
    /// references, and removals require an unreferenced source member.  The
    /// batch rejects duplicate paths and validates every payload before any
    /// staged state changes.  A batch containing only exact replacement
    /// no-ops leaves the transaction unchanged.
    ///
    /// # Errors
    ///
    /// Returns an error for an oversized batch or payload, unsafe/colliding
    /// paths, media-type changes, referenced removals, signed/encrypted
    /// sources, or any invalid operation.  A failed batch leaves this
    /// transaction unchanged.
    pub fn apply_media_changes(&mut self, changes: &[MediaChange]) -> Result<usize> {
        if changes.len() > MAX_MEDIA_CHANGES {
            return invalid(format!(
                "ODP media change batch exceeds {MAX_MEDIA_CHANGES} operations"
            ));
        }
        let added = changes.iter().try_fold(0usize, |total, change| {
            let bytes = match change {
                MediaChange::Add {
                    path,
                    payload,
                    media_type,
                }
                | MediaChange::Replace {
                    path,
                    payload,
                    media_type,
                } => path
                    .len()
                    .checked_add(media_type.len())
                    .and_then(|value| value.checked_add(payload.len()))
                    .ok_or_else(|| invalid_error("ODP media change size overflow"))?,
                MediaChange::Remove { .. } => 0,
            };
            total
                .checked_add(bytes)
                .ok_or_else(|| invalid_error("ODP media change aggregate size overflow"))
        })?;
        if added > 0 {
            let projected_media = self
                .media_bytes
                .checked_add(added)
                .ok_or_else(|| invalid_error("ODP media change projected size overflow"))?;
            self.check_projected(self.resource_bytes, projected_media)?;
        }
        let changed = self.draft.apply_media_changes(changes)?;
        if changed == 0 {
            return Ok(0);
        }
        self.media_bytes = self.draft.staged_media_bytes()?;
        self.check_projected(self.resource_bytes, self.media_bytes)?;
        self.changed = true;
        Ok(changed)
    }

    /// Read arbitrary source-backed text-box, list, table, and form owners.
    ///
    /// Each model retains a compact namespace-complete XML fragment so producer
    /// attributes and children can be inspected or replaced without flattening.
    ///
    /// # Errors
    ///
    /// Returns an error for malformed, ambiguous, or oversized content owners.
    pub fn rich_content(&mut self) -> Result<crate::content::Inventory> {
        let package = self.content_package()?;
        crate::content::inventory(&package)
    }

    /// Add a named common rich-text box to an exact presentation page.
    ///
    /// The object is inserted as a compact source-backed fragment, so opened
    /// pages do not need to be regenerated and unrelated markup is preserved.
    ///
    /// # Errors
    ///
    /// Returns an error for an invalid page, duplicate name, malformed package, or limit.
    pub fn add_text_box<'a, P>(&mut self, page: P, text_box: &crate::content::TextBox) -> Result<()>
    where
        P: Into<crate::charts::Page<'a>>,
    {
        let page_index = self.content_page_index(page.into())?;
        self.stage_content(crate::content::Operation::AddObject {
            page: page_index,
            kind: crate::content::ObjectKind::TextBox,
            name: text_box.name().to_string(),
            xml: text_box.xml()?,
        })
    }

    /// Replace a named rich-text box, optionally changing its stable name.
    ///
    /// # Errors
    ///
    /// Returns an error for a missing/ambiguous name, collision, malformed package, or limit.
    pub fn replace_text_box(
        &mut self,
        name: &str,
        text_box: &crate::content::TextBox,
    ) -> Result<()> {
        self.stage_content(crate::content::Operation::ReplaceObject {
            kind: crate::content::ObjectKind::TextBox,
            name: name.to_string(),
            new_name: text_box.name().to_string(),
            xml: text_box.xml()?,
        })
    }

    /// Replace an arbitrary source-backed text-box/list story.
    ///
    /// # Errors
    ///
    /// Returns an error for a missing/ambiguous source name, collision, or malformed model.
    pub fn replace_text_box_model(
        &mut self,
        name: &str,
        text_box: &crate::content::TextBoxModel,
    ) -> Result<()> {
        self.stage_content(crate::content::Operation::ReplaceObject {
            kind: crate::content::ObjectKind::TextBox,
            name: name.to_string(),
            new_name: text_box.name().to_string(),
            xml: text_box.xml().to_string(),
        })
    }

    /// Atomically replace a bounded set of existing source-backed text boxes.
    ///
    /// Every page/object selector is resolved against the same immutable
    /// staged `content.xml`. Duplicate, overlapping, noncanonical, protected,
    /// or colliding owners are refused before any package bytes change. Caller
    /// order does not affect output bytes, and an all-no-op batch retains the
    /// exact source snapshot.
    ///
    /// Returns the number of owners whose complete models changed.
    ///
    /// # Errors
    ///
    /// Returns an error for more than 256 replacements, invalid or ambiguous
    /// selectors, duplicate selections/destination names, unsafe opaque owner
    /// shapes, malformed models, exhausted limits, or failed complete readback.
    pub fn replace_text_box_models(
        &mut self,
        replacements: &[crate::content::TextBoxModelReplacement<'_>],
    ) -> Result<usize> {
        if replacements.len() > crate::content::MAX_TEXT_BOX_MODEL_REPLACEMENTS {
            return invalid(format!(
                "ODP text-box replacement count exceeds {}",
                crate::content::MAX_TEXT_BOX_MODEL_REPLACEMENTS
            ));
        }
        let mut owned = Vec::new();
        owned
            .try_reserve_exact(replacements.len())
            .map_err(|source| Error::Allocation {
                resource: "ODP text-box batch replacements",
                source,
            })?;
        owned.extend(
            replacements
                .iter()
                .copied()
                .map(crate::content::OwnedTextBoxModelReplacement::from_borrowed),
        );
        self.stage_text_box_models(owned)
    }

    /// Remove a named rich-text box.
    ///
    /// # Errors
    ///
    /// Returns an error for a missing/ambiguous name or malformed package.
    pub fn remove_text_box(&mut self, name: &str) -> Result<()> {
        self.stage_content(crate::content::Operation::RemoveObject {
            kind: crate::content::ObjectKind::TextBox,
            name: name.to_string(),
        })
    }

    /// Add a typed rectangular table to an exact presentation page.
    ///
    /// # Errors
    ///
    /// Returns an error for an invalid page, duplicate name, malformed package, or limit.
    pub fn add_table<'a, P>(&mut self, page: P, table: &crate::content::Table) -> Result<()>
    where
        P: Into<crate::charts::Page<'a>>,
    {
        let page_index = self.content_page_index(page.into())?;
        self.stage_content(crate::content::Operation::AddObject {
            page: page_index,
            kind: crate::content::ObjectKind::Table,
            name: table.name().to_string(),
            xml: table.xml()?,
        })
    }

    /// Replace a named typed table, optionally changing its stable name.
    ///
    /// # Errors
    ///
    /// Returns an error for a missing/ambiguous name, collision, malformed package, or limit.
    pub fn replace_table(&mut self, name: &str, table: &crate::content::Table) -> Result<()> {
        self.stage_content(crate::content::Operation::ReplaceObject {
            kind: crate::content::ObjectKind::Table,
            name: name.to_string(),
            new_name: table.name().to_string(),
            xml: table.xml()?,
        })
    }

    /// Replace an arbitrary source-backed table, including spans, repeats, and formulas.
    ///
    /// # Errors
    ///
    /// Returns an error for a missing/ambiguous source name, collision, or malformed model.
    pub fn replace_table_model(
        &mut self,
        name: &str,
        table: &crate::content::TableModel,
    ) -> Result<()> {
        self.stage_content(crate::content::Operation::ReplaceObject {
            kind: crate::content::ObjectKind::Table,
            name: name.to_string(),
            new_name: table.name().to_string(),
            xml: table.xml().to_string(),
        })
    }

    /// Remove a named typed table.
    ///
    /// # Errors
    ///
    /// Returns an error for a missing/ambiguous name or malformed package.
    pub fn remove_table(&mut self, name: &str) -> Result<()> {
        self.stage_content(crate::content::Operation::RemoveObject {
            kind: crate::content::ObjectKind::Table,
            name: name.to_string(),
        })
    }

    /// Add an inert typed form declaration and its visual control atomically.
    ///
    /// # Errors
    ///
    /// Returns an error for an invalid page, duplicate name, malformed package, or limit.
    pub fn add_form_control<'a, P>(
        &mut self,
        page: P,
        control: &crate::content::FormControl,
    ) -> Result<()>
    where
        P: Into<crate::charts::Page<'a>>,
    {
        let page_index = self.content_page_index(page.into())?;
        self.stage_content(crate::content::Operation::AddControl {
            page: page_index,
            control: control.clone(),
        })
    }

    /// Replace an inert form declaration and visual control as one operation.
    ///
    /// # Errors
    ///
    /// Returns an error for a missing/ambiguous name, collision, malformed package, or limit.
    pub fn replace_form_control(
        &mut self,
        name: &str,
        control: &crate::content::FormControl,
    ) -> Result<()> {
        self.stage_content(crate::content::Operation::ReplaceControl {
            name: name.to_string(),
            control: control.clone(),
        })
    }

    /// Replace an arbitrary source-backed form declaration/control pair atomically.
    ///
    /// # Errors
    ///
    /// Returns an error for a missing/ambiguous source identity, collision, or malformed model.
    pub fn replace_form_control_model(
        &mut self,
        name: &str,
        control: &crate::content::FormControlModel,
    ) -> Result<()> {
        self.stage_content(crate::content::Operation::ReplaceControlModel {
            name: name.to_string(),
            new_name: control.name().to_string(),
            declaration_xml: control.declaration_xml().to_string(),
            visual_xml: control.visual_xml().to_string(),
        })
    }

    /// Remove an inert form declaration and its visual control atomically.
    ///
    /// # Errors
    ///
    /// Returns an error for a missing/ambiguous name or malformed package.
    pub fn remove_form_control(&mut self, name: &str) -> Result<()> {
        self.stage_content(crate::content::Operation::RemoveControl {
            name: name.to_string(),
        })
    }

    /// Transfer a dependency-closed rich-text box from another immutable deck.
    ///
    /// Named styles and package resources are copied, and destination collisions
    /// are remapped without changing the source snapshot.
    ///
    /// # Errors
    ///
    /// Returns an error for a missing object/page, dangling dependency, unsafe resource, or limit.
    pub fn transfer_text_box_from<'a, P>(
        &mut self,
        source: &Snapshot,
        source_name: &str,
        destination_page: P,
        destination_name: impl Into<String>,
    ) -> Result<()>
    where
        P: Into<crate::charts::Page<'a>>,
    {
        self.transfer_content_object_from(
            source,
            source_name,
            destination_page.into(),
            destination_name.into(),
            crate::content::ObjectKind::TextBox,
        )
    }

    /// Transfer a dependency-closed rich-cell table from another immutable deck.
    ///
    /// # Errors
    ///
    /// Returns an error for a missing object/page, dangling dependency, unsafe resource, or limit.
    pub fn transfer_table_from<'a, P>(
        &mut self,
        source: &Snapshot,
        source_name: &str,
        destination_page: P,
        destination_name: impl Into<String>,
    ) -> Result<()>
    where
        P: Into<crate::charts::Page<'a>>,
    {
        self.transfer_content_object_from(
            source,
            source_name,
            destination_page.into(),
            destination_name.into(),
            crate::content::ObjectKind::Table,
        )
    }

    /// Transfer an inert form declaration and its visual control from another deck.
    ///
    /// # Errors
    ///
    /// Returns an error for a missing owner/page, dangling dependency, collision, or limit.
    pub fn transfer_form_control_from<'a, P>(
        &mut self,
        source: &Snapshot,
        source_name: &str,
        destination_page: P,
        destination_name: impl Into<String>,
    ) -> Result<()>
    where
        P: Into<crate::charts::Page<'a>>,
    {
        let page_index = self.content_page_index(destination_page.into())?;
        let source_presentation = source.to_presentation()?;
        let operation = crate::content::prepare_control_transfer(
            source_presentation.owned_package(),
            page_index,
            source_name,
            destination_name.into(),
        )?;
        self.stage_content(operation)
    }

    /// Inspect the RDF metadata graphs in the current transaction draft.
    ///
    /// The inventory is loaded lazily so slide-only transactions do not reject
    /// unrelated malformed metadata that they never touch.
    ///
    /// # Errors
    ///
    /// Returns an error when a declared RDF part is malformed, dangling, or over budget.
    pub fn rdf_graphs(&mut self) -> Result<&[crate::rdf::Graph]> {
        self.ensure_rdf()?;
        self.rdf
            .as_ref()
            .map(|draft| draft.graphs.as_slice())
            .ok_or_else(|| invalid_error("ODP RDF draft initialization failed"))
    }

    /// Add one RDF metadata graph to this package transaction.
    ///
    /// A missing preferred path is resolved immediately to a collision-free,
    /// deterministic package path which is retained by the transaction.
    ///
    /// # Errors
    ///
    /// Returns an error for an unsafe or colliding path, invalid triples, or a package limit.
    pub fn add_rdf_graph(
        &mut self,
        preferred_path: Option<&str>,
        triples: &[crate::rdf::Triple],
    ) -> Result<String> {
        let package = self.rdf_package()?;
        let (bytes, path) = crate::rdf::add_graph(&package, preferred_path, triples)?;
        self.stage_rdf(
            bytes,
            RdfOperation::AddGraph {
                path: path.clone(),
                triples: triples.to_vec(),
            },
        )?;
        Ok(path)
    }

    /// Replace all triples in one RDF metadata graph.
    ///
    /// # Errors
    ///
    /// Returns an error for a missing graph, invalid triples, or a package limit.
    pub fn replace_rdf_graph(&mut self, path: &str, triples: &[crate::rdf::Triple]) -> Result<()> {
        let package = self.rdf_package()?;
        let bytes = crate::rdf::replace_graph(&package, path, triples)?;
        self.stage_rdf(
            bytes,
            RdfOperation::ReplaceGraph {
                path: path.to_string(),
                triples: triples.to_vec(),
            },
        )
    }

    /// Remove one RDF metadata graph after dependency validation.
    ///
    /// # Errors
    ///
    /// Returns an error for a missing graph or an incoming graph reference.
    pub fn remove_rdf_graph(&mut self, path: &str) -> Result<()> {
        let package = self.rdf_package()?;
        let bytes = crate::rdf::remove_graph(&package, path)?;
        self.stage_rdf(
            bytes,
            RdfOperation::RemoveGraph {
                path: path.to_string(),
            },
        )
    }

    /// Append one RDF triple and return its checked projected index.
    ///
    /// # Errors
    ///
    /// Returns an error for a missing graph, invalid triple, or graph limit.
    pub fn add_rdf_triple(&mut self, path: &str, triple: &crate::rdf::Triple) -> Result<usize> {
        let package = self.rdf_package()?;
        let (bytes, index) = crate::rdf::add_triple(&package, path, triple)?;
        self.stage_rdf(
            bytes,
            RdfOperation::AddTriple {
                path: path.to_string(),
                triple: triple.clone(),
            },
        )?;
        Ok(index)
    }

    /// Replace one RDF triple selected by checked zero-based position.
    ///
    /// # Errors
    ///
    /// Returns an error for a missing graph, out-of-range position, changed subject, or invalid
    /// triple.
    pub fn replace_rdf_triple(
        &mut self,
        path: &str,
        index: usize,
        triple: &crate::rdf::Triple,
    ) -> Result<()> {
        let package = self.rdf_package()?;
        let bytes = crate::rdf::replace_triple(&package, path, index, triple)?;
        self.stage_rdf(
            bytes,
            RdfOperation::ReplaceTriple {
                path: path.to_string(),
                index,
                triple: triple.clone(),
            },
        )
    }

    /// Remove one RDF triple selected by checked zero-based position.
    ///
    /// # Errors
    ///
    /// Returns an error for a missing graph or out-of-range position.
    pub fn remove_rdf_triple(&mut self, path: &str, index: usize) -> Result<()> {
        let package = self.rdf_package()?;
        let bytes = crate::rdf::remove_triple(&package, path, index)?;
        self.stage_rdf(
            bytes,
            RdfOperation::RemoveTriple {
                path: path.to_string(),
                index,
            },
        )
    }

    /// Move one RDF triple to another checked zero-based position within its subject.
    ///
    /// # Errors
    ///
    /// Returns an error for a missing graph, out-of-range position, or subject mismatch.
    pub fn move_rdf_triple(&mut self, path: &str, from: usize, to: usize) -> Result<()> {
        let package = self.rdf_package()?;
        let bytes = crate::rdf::move_triple(&package, path, from, to)?;
        self.stage_rdf(
            bytes,
            RdfOperation::MoveTriple {
                path: path.to_string(),
                from,
                to,
            },
        )
    }

    /// Inspect embedded charts in the current package transaction draft.
    ///
    /// # Errors
    ///
    /// Returns an error when chart discovery encounters malformed or over-budget content.
    pub fn charts(&mut self) -> Result<&[crate::charts::Chart]> {
        self.ensure_charts()?;
        self.charts
            .as_ref()
            .map(|draft| draft.charts.as_slice())
            .ok_or_else(|| invalid_error("ODP chart draft initialization failed"))
    }

    /// Replace one embedded chart part selected by exact name or checked position.
    ///
    /// Every occurrence sharing the selected package part is updated together.
    ///
    /// # Errors
    ///
    /// Returns an error for a missing or ambiguous selector, malformed chart XML, or a limit.
    pub fn replace_chart<'a, S>(&mut self, selector: S, part: crate::charts::Part) -> Result<()>
    where
        S: Into<crate::charts::Selector<'a>>,
    {
        let owned_selector = ChartSelector::from_borrowed(selector.into());
        let snapshot = self.chart_snapshot()?;
        let mut edit = snapshot.edit();
        edit.replace(owned_selector.borrowed(), part.clone())?;
        let commit = edit.commit()?;
        self.stage_chart(
            commit.snapshot().owned_package().clone(),
            commit.snapshot().charts().to_vec(),
            ChartOperation::Replace {
                selector: owned_selector,
                part,
            },
        )
    }

    /// Replace one chart from the complete typed ODF chart authoring model.
    ///
    /// This is the unified-root entry point for cached tables, typed series,
    /// axes, legends, plot-area details, and chart-local styles.
    ///
    /// # Errors
    ///
    /// Returns an error for an invalid definition, missing or ambiguous selector, or limit.
    pub fn replace_chart_definition<'a, S>(
        &mut self,
        selector: S,
        definition: &crate::charts::Definition,
    ) -> Result<()>
    where
        S: Into<crate::charts::Selector<'a>>,
    {
        self.replace_chart(selector, crate::charts::Part::from_definition(definition)?)
    }

    /// Append one typed chart series without replacing the rest of the chart part.
    ///
    /// # Errors
    ///
    /// Returns an error for a missing/ambiguous chart, missing plot area, invalid series, or limit.
    pub fn add_chart_series<'a, S>(
        &mut self,
        selector: S,
        series: &crate::charts::SeriesSpec,
    ) -> Result<()>
    where
        S: Into<crate::charts::Selector<'a>>,
    {
        let owned_selector = ChartSelector::from_borrowed(selector.into());
        let snapshot = self.chart_snapshot()?;
        let chart = snapshot
            .get(owned_selector.borrowed())?
            .ok_or_else(|| invalid_error("ODP chart selector did not match"))?;
        let part = chart.part().with_series_added(series)?;
        self.replace_chart(owned_selector.borrowed(), part)
    }

    /// Replace one physical chart series by checked zero-based position.
    ///
    /// # Errors
    ///
    /// Returns an error for a missing/ambiguous chart, out-of-range series, invalid value, or limit.
    pub fn replace_chart_series<'a, S>(
        &mut self,
        selector: S,
        series_index: usize,
        series: &crate::charts::SeriesSpec,
    ) -> Result<()>
    where
        S: Into<crate::charts::Selector<'a>>,
    {
        let owned_selector = ChartSelector::from_borrowed(selector.into());
        let snapshot = self.chart_snapshot()?;
        let chart = snapshot
            .get(owned_selector.borrowed())?
            .ok_or_else(|| invalid_error("ODP chart selector did not match"))?;
        let part = chart.part().with_series_replaced(series_index, series)?;
        self.replace_chart(owned_selector.borrowed(), part)
    }

    /// Remove one physical chart series by checked zero-based position.
    ///
    /// # Errors
    ///
    /// Returns an error for a missing/ambiguous chart, out-of-range series, malformed XML, or limit.
    pub fn remove_chart_series<'a, S>(&mut self, selector: S, series_index: usize) -> Result<()>
    where
        S: Into<crate::charts::Selector<'a>>,
    {
        let owned_selector = ChartSelector::from_borrowed(selector.into());
        let snapshot = self.chart_snapshot()?;
        let chart = snapshot
            .get(owned_selector.borrowed())?
            .ok_or_else(|| invalid_error("ODP chart selector did not match"))?;
        let part = chart.part().with_series_removed(series_index)?;
        self.replace_chart(owned_selector.borrowed(), part)
    }

    /// Replace one physical cached-table cell by checked row and column positions.
    ///
    /// Header rows are included in row indexing. Repeated XML runs remain
    /// physical entries and are not expanded implicitly.
    ///
    /// # Errors
    ///
    /// Returns an error for a missing chart/table, out-of-range coordinate, invalid cell, or limit.
    pub fn replace_chart_cached_cell<'a, S>(
        &mut self,
        selector: S,
        row: usize,
        column: usize,
        cell: &crate::charts::CachedCell,
    ) -> Result<()>
    where
        S: Into<crate::charts::Selector<'a>>,
    {
        let owned_selector = ChartSelector::from_borrowed(selector.into());
        let snapshot = self.chart_snapshot()?;
        let chart = snapshot
            .get(owned_selector.borrowed())?
            .ok_or_else(|| invalid_error("ODP chart selector did not match"))?;
        let part = chart.part().with_cached_cell_replaced(row, column, cell)?;
        self.replace_chart(owned_selector.borrowed(), part)
    }

    /// Remove one embedded chart selected by exact name or checked position.
    ///
    /// # Errors
    ///
    /// Returns an error for a missing or ambiguous selector or malformed chart package content.
    pub fn remove_chart<'a, S>(&mut self, selector: S) -> Result<crate::charts::Chart>
    where
        S: Into<crate::charts::Selector<'a>>,
    {
        let owned_selector = ChartSelector::from_borrowed(selector.into());
        let snapshot = self.chart_snapshot()?;
        let mut edit = snapshot.edit();
        let removed = edit.remove(owned_selector.borrowed())?;
        let commit = edit.commit()?;
        self.stage_chart(
            commit.snapshot().owned_package().clone(),
            commit.snapshot().charts().to_vec(),
            ChartOperation::Remove {
                selector: owned_selector,
            },
        )?;
        Ok(removed)
    }

    /// Add a named embedded chart to an exact-name or checked-position page.
    ///
    /// # Errors
    ///
    /// Returns an error for a missing or ambiguous page, duplicate chart name, malformed part,
    /// or resource-limit breach.
    pub fn add_chart<'a, P>(
        &mut self,
        page: P,
        name: impl Into<String>,
        storage: crate::charts::Storage,
        part: crate::charts::Part,
    ) -> Result<usize>
    where
        P: Into<crate::charts::Page<'a>>,
    {
        let owned_page = ChartPage::from_borrowed(page.into());
        let chart_name = name.into();
        let snapshot = self.chart_snapshot()?;
        let mut edit = snapshot.edit();
        let index = edit.add(
            owned_page.borrowed(),
            chart_name.clone(),
            storage,
            part.clone(),
        )?;
        let commit = edit.commit()?;
        self.stage_chart(
            commit.snapshot().owned_package().clone(),
            commit.snapshot().charts().to_vec(),
            ChartOperation::Add {
                page: owned_page,
                name: chart_name,
                storage,
                part,
            },
        )?;
        Ok(index)
    }

    /// Add a chart from the complete typed ODF chart authoring model.
    ///
    /// # Errors
    ///
    /// Returns an error for an invalid definition, page selector, name, or resource limit.
    pub fn add_chart_definition<'a, P>(
        &mut self,
        page: P,
        name: impl Into<String>,
        storage: crate::charts::Storage,
        definition: &crate::charts::Definition,
    ) -> Result<usize>
    where
        P: Into<crate::charts::Page<'a>>,
    {
        self.add_chart(
            page,
            name,
            storage,
            crate::charts::Part::from_definition(definition)?,
        )
    }

    /// Copy one dependency-closed chart from another immutable presentation snapshot.
    ///
    /// The chart's complete typed part, including chart-local styles, cached data, and
    /// package-contained `xlink:href` resources, is detached from the source. Resource-path
    /// collisions are remapped deterministically; external references remain inert and are
    /// never fetched. The destination always receives a fresh occurrence and storage root.
    ///
    /// # Errors
    ///
    /// Returns an error for a missing or ambiguous source chart, dangling/unsafe package
    /// dependency, invalid destination selector, identity collision, or resource limit.
    pub fn transfer_chart_from<'a, 'b, S, P>(
        &mut self,
        source: &Snapshot,
        source_chart: S,
        destination_page: P,
        destination_name: impl Into<String>,
        storage: crate::charts::Storage,
    ) -> Result<usize>
    where
        S: Into<crate::charts::Selector<'a>>,
        P: Into<crate::charts::Page<'b>>,
    {
        let inventory = crate::charts::Snapshot::from_owned_package(
            source.package.clone(),
            crate::charts::Limits::default(),
        )?;
        let selected = inventory
            .get(source_chart)?
            .ok_or_else(|| invalid_error("ODP source chart selector did not match"))?;
        let source_package = source.package.clone();
        let destination_package = self.content_package()?;
        let source_base = selected.content_path().unwrap_or("content.xml");
        let destination_base = match storage {
            crate::charts::Storage::InlineXml => "content.xml",
            crate::charts::Storage::PackageSubdocument => "Object/content.xml",
        };
        let (chart_xml, resources) = crate::content::prepare_resource_transfer(
            &source_package,
            &destination_package,
            selected.part().xml(),
            source_base,
            destination_base,
        )?;
        let part = crate::charts::Part::from_xml(chart_xml)?;
        let index = self.add_chart(destination_page, destination_name, storage, part)?;
        if !crate::content::resource_operation_is_empty(&resources) {
            self.stage_content(resources)?;
        }
        Ok(index)
    }

    /// Inspect named presentation page layouts in the current package draft.
    ///
    /// # Errors
    ///
    /// Returns an error when `styles.xml` is missing, malformed, or over budget.
    pub fn layouts(&mut self) -> Result<&crate::layout::Collection> {
        self.ensure_design()?;
        self.design
            .as_ref()
            .map(|draft| &draft.layouts)
            .ok_or_else(|| invalid_error("ODP design draft initialization failed"))
    }

    /// Inspect master pages in the current package draft.
    ///
    /// # Errors
    ///
    /// Returns an error when `styles.xml` is missing, malformed, or over budget.
    pub fn master_pages(&mut self) -> Result<&[crate::MasterPage]> {
        self.ensure_design()?;
        self.design
            .as_ref()
            .map(|draft| draft.masters.as_slice())
            .ok_or_else(|| invalid_error("ODP design draft initialization failed"))
    }

    /// Add one named presentation page layout.
    ///
    /// # Errors
    ///
    /// Returns an error for an invalid or duplicate layout or malformed package dependencies.
    pub fn add_layout(&mut self, layout: &crate::layout::Layout) -> Result<()> {
        self.stage_design_operation(DesignOperation::AddLayout(layout.clone()))
    }

    /// Replace one named presentation page layout.
    ///
    /// # Errors
    ///
    /// Returns an error when the layout is invalid or does not exist.
    pub fn replace_page_layout(&mut self, layout: &crate::layout::Layout) -> Result<()> {
        self.stage_design_operation(DesignOperation::ReplaceLayout(layout.clone()))
    }

    /// Remove one layout and optionally retarget all modeled incoming references.
    ///
    /// # Errors
    ///
    /// Returns an error when either name is invalid or the replacement does not exist.
    pub fn remove_page_layout(&mut self, name: &str, replacement: Option<&str>) -> Result<()> {
        self.stage_design_operation(DesignOperation::RemoveLayout {
            name: name.to_string(),
            replacement: replacement.map(str::to_string),
        })
    }

    /// Reorder every named presentation layout using an exact dependency-checked name list.
    ///
    /// # Errors
    ///
    /// Returns an error when the list is incomplete, duplicated, or contains an unknown name.
    pub fn reorder_layouts(&mut self, names: &[String]) -> Result<()> {
        self.stage_design_operation(DesignOperation::ReorderLayouts(names.to_vec()))
    }

    /// Add one named master page.
    ///
    /// # Errors
    ///
    /// Returns an error for invalid XML, duplicate identity, or dangling dependencies.
    pub fn add_master_page(&mut self, master: &crate::MasterPage) -> Result<()> {
        self.stage_design_operation(DesignOperation::AddMaster(master.clone()))
    }

    /// Replace one named master page.
    ///
    /// # Errors
    ///
    /// Returns an error for invalid XML, missing identity, or dangling dependencies.
    pub fn replace_master_page(&mut self, master: &crate::MasterPage) -> Result<()> {
        self.stage_design_operation(DesignOperation::ReplaceMaster(master.clone()))
    }

    /// Remove one master page and optionally retarget modeled incoming references.
    ///
    /// # Errors
    ///
    /// Returns an error when either name is invalid or the replacement does not exist.
    pub fn remove_master_page(&mut self, name: &str, replacement: Option<&str>) -> Result<()> {
        self.stage_design_operation(DesignOperation::RemoveMaster {
            name: name.to_string(),
            replacement: replacement.map(str::to_string),
        })
    }

    /// Reorder every master page using an exact dependency-checked name list.
    ///
    /// # Errors
    ///
    /// Returns an error when the list is incomplete, duplicated, or contains an unknown name.
    pub fn reorder_master_pages(&mut self, names: &[String]) -> Result<()> {
        self.stage_design_operation(DesignOperation::ReorderMasters(names.to_vec()))
    }

    /// Assign or clear a slide's master-page dependency by checked zero-based position.
    ///
    /// # Errors
    ///
    /// Returns an error for an out-of-range slide or missing master.
    pub fn assign_slide_master_page(
        &mut self,
        slide_index: usize,
        master_name: Option<&str>,
    ) -> Result<()> {
        self.stage_design_operation(DesignOperation::AssignSlideMaster {
            slide_index,
            name: master_name.map(str::to_string),
        })
    }

    /// Assign or clear a slide's presentation-layout dependency by checked position.
    ///
    /// # Errors
    ///
    /// Returns an error for an out-of-range slide or missing layout.
    pub fn assign_slide_page_layout(
        &mut self,
        slide_index: usize,
        layout_name: Option<&str>,
    ) -> Result<()> {
        self.stage_design_operation(DesignOperation::AssignSlideLayout {
            slide_index,
            name: layout_name.map(str::to_string),
        })
    }

    /// Inspect slide- and shape-anchored annotations in the current package draft.
    ///
    /// # Errors
    ///
    /// Returns an error for malformed annotation XML, anchors, or resource limits.
    pub fn annotations(&mut self) -> Result<&[crate::annotation::Info]> {
        self.ensure_annotations()?;
        self.annotations
            .as_ref()
            .map(|draft| draft.annotations.as_slice())
            .ok_or_else(|| invalid_error("ODP annotation draft initialization failed"))
    }

    /// Add an annotation at a checked page or uniquely named shape anchor.
    ///
    /// # Errors
    ///
    /// Returns an error for a missing or ambiguous anchor, duplicate name, or invalid body.
    pub fn add_annotation(
        &mut self,
        anchor: &crate::annotation::Anchor,
        annotation: &crate::annotation::Annotation,
    ) -> Result<usize> {
        self.ensure_annotations()?;
        let current = self.annotation_package()?;
        let mut presentation = Presentation::from_owned_package(current)?;
        let index = presentation.add_annotation(anchor, annotation)?;
        self.stage_annotation(
            OwnedPackage::from_bytes(presentation.to_bytes()?)?,
            AnnotationOperation::Add {
                anchor: anchor.clone(),
                annotation: annotation.clone(),
            },
        )?;
        Ok(index)
    }

    /// Replace one annotation selected by checked zero-based document order.
    ///
    /// # Errors
    ///
    /// Returns an error for an out-of-range position, duplicate name, or invalid body.
    pub fn replace_annotation(
        &mut self,
        index: usize,
        annotation: &crate::annotation::Annotation,
    ) -> Result<()> {
        self.ensure_annotations()?;
        let current = self.annotation_package()?;
        let mut presentation = Presentation::from_owned_package(current)?;
        presentation.replace_annotation(index, annotation)?;
        self.stage_annotation(
            OwnedPackage::from_bytes(presentation.to_bytes()?)?,
            AnnotationOperation::Replace {
                index,
                annotation: annotation.clone(),
            },
        )
    }

    /// Remove one annotation selected by checked zero-based document order.
    ///
    /// # Errors
    ///
    /// Returns an error when the position is out of range or package content is malformed.
    pub fn remove_annotation(&mut self, index: usize) -> Result<()> {
        self.ensure_annotations()?;
        let current = self.annotation_package()?;
        let mut presentation = Presentation::from_owned_package(current)?;
        presentation.remove_annotation(index)?;
        self.stage_annotation(
            OwnedPackage::from_bytes(presentation.to_bytes()?)?,
            AnnotationOperation::Remove { index },
        )
    }

    /// Validate, serialize, reparse, and atomically publish the staged draft.
    ///
    /// # Errors
    ///
    /// Returns an error if validation, bounded serialization, package parsing, or semantic
    /// readback fails. The source snapshot is never changed.
    pub fn commit(self) -> Result<Commit> {
        let rdf_changed = self
            .rdf
            .as_ref()
            .is_some_and(|draft| !draft.operations.is_empty());
        let charts_changed = self
            .charts
            .as_ref()
            .is_some_and(|draft| !draft.operations.is_empty());
        let design_changed = self
            .design
            .as_ref()
            .is_some_and(|draft| !draft.operations.is_empty());
        let annotations_changed = self
            .annotations
            .as_ref()
            .is_some_and(|draft| !draft.operations.is_empty());
        let content_changed = self
            .content
            .as_ref()
            .is_some_and(|draft| !draft.operations.is_empty());
        let slide_only = self.changed
            && !rdf_changed
            && !charts_changed
            && !design_changed
            && !annotations_changed
            && !content_changed;
        let mut domains = Vec::new();
        if self.changed {
            domains.push(Domain::Slides);
        }
        if rdf_changed {
            domains.push(Domain::Rdf);
        }
        if charts_changed {
            domains.push(Domain::Charts);
        }
        if design_changed {
            domains.push(Domain::Design);
        }
        if annotations_changed {
            domains.push(Domain::Annotations);
        }
        if content_changed {
            domains.push(Domain::Content);
        }
        if !self.changed
            && !rdf_changed
            && !charts_changed
            && !design_changed
            && !annotations_changed
            && !content_changed
        {
            return Ok(Commit::unchanged(self.source));
        }
        let (mut bytes, mut package, slide_candidate) = if self.changed {
            let package =
                OwnedPackage::from_bytes(self.draft.to_bytes_bounded(MAX_PACKAGE_BYTES)?)?;
            let bytes = package.shared_bytes();
            let candidate = Snapshot::from_owned_package(package.clone())?;
            if candidate.slides() != self.draft.slides() {
                return invalid("ODP transaction readback differs from the staged slide model");
            }
            (bytes, package, Some(candidate))
        } else {
            let package = self.source.package.clone();
            (package.shared_bytes(), package, None)
        };
        if let Some(design) = &self.design
            && !design.operations.is_empty()
        {
            if self.changed {
                for operation in &design.operations {
                    let output = apply_design_operation(&package, operation)?;
                    if output.len() > MAX_PACKAGE_BYTES {
                        return invalid("ODP design transaction exceeds the 128 MiB package limit");
                    }
                    package = OwnedPackage::from_bytes(output)?;
                    bytes = package.shared_bytes();
                }
            } else {
                package = design.package.clone();
                bytes = package.shared_bytes();
            }
        }
        if let Some(annotations) = &self.annotations
            && !annotations.operations.is_empty()
        {
            if self.changed || design_changed {
                for operation in &annotations.operations {
                    let output = apply_annotation_operation(&package, operation)?;
                    if output.len() > MAX_PACKAGE_BYTES {
                        return invalid(
                            "ODP annotation transaction exceeds the 128 MiB package limit",
                        );
                    }
                    package = OwnedPackage::from_bytes(output)?;
                    bytes = package.shared_bytes();
                }
            } else {
                package = annotations.package.clone();
                bytes = package.shared_bytes();
            }
        }
        if let Some(rdf) = &self.rdf
            && !rdf.operations.is_empty()
        {
            if self.changed || design_changed || annotations_changed {
                for operation in &rdf.operations {
                    let output = apply_rdf_operation(&package, operation)?;
                    if output.len() > MAX_PACKAGE_BYTES {
                        return invalid("ODP RDF transaction exceeds the 128 MiB package limit");
                    }
                    package = OwnedPackage::from_bytes(output)?;
                    bytes = package.shared_bytes();
                }
            } else {
                package = rdf.package.clone();
                bytes = package.shared_bytes();
            }
        }
        if let Some(charts) = &self.charts
            && !charts.operations.is_empty()
        {
            if self.changed || design_changed || annotations_changed || rdf_changed {
                for operation in &charts.operations {
                    let output = apply_chart_operation(&package, charts.limits, operation)?;
                    if output.len() > MAX_PACKAGE_BYTES {
                        return invalid("ODP chart transaction exceeds the 128 MiB package limit");
                    }
                    package = OwnedPackage::from_bytes(output)?;
                    bytes = package.shared_bytes();
                }
            } else {
                package = charts.package.clone();
                bytes = package.shared_bytes();
            }
        }
        if let Some(content) = &self.content
            && !content.operations.is_empty()
        {
            if self.changed
                || design_changed
                || annotations_changed
                || rdf_changed
                || charts_changed
            {
                for operation in &content.operations {
                    let output = crate::content::apply(&package, operation)?;
                    if output.len() > MAX_PACKAGE_BYTES {
                        return invalid(
                            "ODP semantic-content transaction exceeds the 128 MiB package limit",
                        );
                    }
                    package = OwnedPackage::from_bytes(output)?;
                    bytes = package.shared_bytes();
                }
            } else {
                package = content.package.clone();
                bytes = package.shared_bytes();
            }
        }
        let reopened = package;
        let source_package = self.source.package.clone();
        if !content_changed {
            validate_raw_preserved_referenced_xml_parts(&source_package)?;
        }
        validate_compact_xml_parts(&reopened, &source_package)?;
        self.draft.verify_embedded_media(&reopened)?;
        self.draft.verify_removed_media(&reopened)?;
        if let Some(rdf) = &self.rdf
            && crate::rdf::graphs(&reopened)? != rdf.graphs
        {
            return invalid("ODP transaction RDF readback differs from the staged graph model");
        }
        if let Some(charts) = &self.charts {
            let reopened_charts =
                crate::charts::Snapshot::from_owned_package(reopened.clone(), charts.limits)?;
            if !root_charts_equal(reopened_charts.charts(), &charts.charts) {
                return invalid("ODP transaction chart readback differs from the staged model");
            }
        }
        let presentation = if design_changed || annotations_changed {
            Some(Presentation::from_owned_package(reopened.clone())?)
        } else {
            None
        };
        if let Some(design) = &self.design {
            let presentation = presentation
                .as_ref()
                .ok_or_else(|| invalid_error("ODP design readback presentation missing"))?;
            if presentation.layouts()? != design.layouts
                || !root_masters_equal(&presentation.master_pages()?, &design.masters)
                || !root_design_pages_equal(&presentation.pages()?, &design.pages)
            {
                return invalid("ODP transaction design readback differs from the staged model");
            }
        }
        if let Some(annotations) = &self.annotations {
            let presentation = presentation
                .as_ref()
                .ok_or_else(|| invalid_error("ODP annotation readback presentation missing"))?;
            if !root_annotations_equal(&presentation.annotations()?, &annotations.annotations) {
                return invalid(
                    "ODP transaction annotation readback differs from the staged model",
                );
            }
        }
        let snapshot = match (slide_only, slide_candidate) {
            (true, Some(candidate)) => {
                debug_assert!(Arc::ptr_eq(&candidate.bytes, &bytes));
                candidate
            },
            (true, None) => unreachable!("a slide-only ODP commit always has a slide candidate"),
            (false, _) => Snapshot::from_owned_package(reopened)?,
        };
        let patch = Patch {
            before: self.source,
            after: snapshot.clone(),
            domains: Arc::from(domains),
        };
        Ok(Commit {
            snapshot,
            patch,
            changed: true,
        })
    }

    fn ensure_rdf(&mut self) -> Result<()> {
        if self.rdf.is_none() {
            let package = self.source.package.clone();
            let graphs = crate::rdf::graphs(&package)?;
            self.rdf = Some(RdfDraft {
                package,
                original_graphs: graphs.clone(),
                graphs,
                operations: Vec::new(),
            });
        }
        Ok(())
    }

    fn rdf_package(&mut self) -> Result<OwnedPackage> {
        self.ensure_rdf()?;
        self.rdf
            .as_ref()
            .map(|draft| draft.package.clone())
            .ok_or_else(|| invalid_error("ODP RDF draft initialization failed"))
    }

    fn stage_rdf(&mut self, bytes: Vec<u8>, operation: RdfOperation) -> Result<()> {
        if bytes.len() > MAX_PACKAGE_BYTES {
            return invalid("ODP RDF transaction exceeds the 128 MiB package limit");
        }
        let package = OwnedPackage::from_bytes(bytes)?;
        let presentation = Presentation::from_owned_package(package.clone())?;
        let graphs = crate::rdf::graphs(presentation.owned_package())?;
        let draft = self
            .rdf
            .as_mut()
            .ok_or_else(|| invalid_error("ODP RDF draft initialization failed"))?;
        if graphs == draft.graphs {
            return Ok(());
        }
        if graphs == draft.original_graphs {
            draft.package = self.source.package.clone();
            draft.graphs = graphs;
            draft.operations.clear();
            return Ok(());
        }
        draft.package = package;
        draft.graphs = graphs;
        draft.operations.push(operation);
        Ok(())
    }

    fn ensure_charts(&mut self) -> Result<()> {
        if self.charts.is_none() {
            let limits = crate::charts::Limits::default();
            let snapshot =
                crate::charts::Snapshot::from_owned_package(self.source.package.clone(), limits)?;
            let charts = snapshot.charts().to_vec();
            self.charts = Some(ChartDraft {
                package: self.source.package.clone(),
                original: charts.clone(),
                charts,
                operations: Vec::new(),
                limits,
            });
        }
        Ok(())
    }

    fn chart_snapshot(&mut self) -> Result<crate::charts::Snapshot> {
        self.ensure_charts()?;
        let (current_package, limits, operations) = self
            .charts
            .as_ref()
            .map(|draft| {
                (
                    draft.package.clone(),
                    draft.limits,
                    draft.operations.clone(),
                )
            })
            .ok_or_else(|| invalid_error("ODP chart draft initialization failed"))?;
        if !self.changed {
            return crate::charts::Snapshot::from_owned_package(current_package, limits);
        }
        let mut package =
            OwnedPackage::from_bytes(self.draft.to_bytes_bounded(MAX_PACKAGE_BYTES)?)?;
        for operation in &operations {
            package =
                OwnedPackage::from_bytes(apply_chart_operation(&package, limits, operation)?)?;
        }
        crate::charts::Snapshot::from_owned_package(package, limits)
    }

    fn stage_chart(
        &mut self,
        package: OwnedPackage,
        charts: Vec<crate::charts::Chart>,
        operation: ChartOperation,
    ) -> Result<()> {
        self.check_no_slide_order_change("chart")?;
        if package.as_bytes().len() > MAX_PACKAGE_BYTES {
            return invalid("ODP chart transaction exceeds the 128 MiB package limit");
        }
        let draft = self
            .charts
            .as_mut()
            .ok_or_else(|| invalid_error("ODP chart draft initialization failed"))?;
        if root_charts_equal(&charts, &draft.charts) {
            return Ok(());
        }
        if root_charts_equal(&charts, &draft.original) {
            draft.package = self.source.package.clone();
            draft.charts = charts;
            draft.operations.clear();
            return Ok(());
        }
        draft.package = package;
        draft.charts = charts;
        draft.operations.push(operation);
        Ok(())
    }

    fn ensure_design(&mut self) -> Result<()> {
        if self.design.is_none() {
            let package = self.source.package.clone();
            let presentation = Presentation::from_owned_package(package.clone())?;
            let layouts = presentation.layouts()?;
            let masters = presentation.master_pages()?;
            let pages = presentation.pages()?;
            self.design = Some(DesignDraft {
                package,
                original_layouts: layouts.clone(),
                layouts,
                original_masters: masters.clone(),
                masters,
                original_pages: pages.clone(),
                pages,
                operations: Vec::new(),
            });
        }
        Ok(())
    }

    fn stage_design_operation(&mut self, operation: DesignOperation) -> Result<()> {
        self.check_no_slide_order_change("design")?;
        self.ensure_design()?;
        let current = self
            .design
            .as_ref()
            .map(|draft| draft.package.clone())
            .ok_or_else(|| invalid_error("ODP design draft initialization failed"))?;
        let package = OwnedPackage::from_bytes(apply_design_operation(&current, &operation)?)?;
        if package.as_bytes().len() > MAX_PACKAGE_BYTES {
            return invalid("ODP design transaction exceeds the 128 MiB package limit");
        }
        let presentation = Presentation::from_owned_package(package.clone())?;
        let layouts = presentation.layouts()?;
        let masters = presentation.master_pages()?;
        let pages = presentation.pages()?;
        let draft = self
            .design
            .as_mut()
            .ok_or_else(|| invalid_error("ODP design draft initialization failed"))?;
        if layouts == draft.layouts
            && root_masters_equal(&masters, &draft.masters)
            && pages == draft.pages
        {
            return Ok(());
        }
        if layouts == draft.original_layouts
            && root_masters_equal(&masters, &draft.original_masters)
            && pages == draft.original_pages
        {
            draft.package = self.source.package.clone();
            draft.layouts = layouts;
            draft.masters = masters;
            draft.pages = pages;
            draft.operations.clear();
            return Ok(());
        }
        draft.package = package;
        draft.layouts = layouts;
        draft.masters = masters;
        draft.pages = pages;
        draft.operations.push(operation);
        Ok(())
    }

    fn ensure_annotations(&mut self) -> Result<()> {
        if self.annotations.is_none() {
            let package = self.source.package.clone();
            let presentation = Presentation::from_owned_package(package.clone())?;
            let annotations = presentation.annotations()?;
            self.annotations = Some(AnnotationDraft {
                package,
                original: annotations.clone(),
                annotations,
                operations: Vec::new(),
            });
        }
        Ok(())
    }

    fn content_page_index(&mut self, page: crate::charts::Page<'_>) -> Result<usize> {
        let package = self.content_package()?;
        let content = String::from_utf8(package.get_file("content.xml")?)
            .map_err(|source| invalid_error(format!("ODP content.xml is not UTF-8: {source}")))?;
        crate::charts::page_index(&content, page)
    }

    fn transfer_content_object_from(
        &mut self,
        source: &Snapshot,
        source_name: &str,
        destination_page: crate::charts::Page<'_>,
        destination_name: String,
        kind: crate::content::ObjectKind,
    ) -> Result<()> {
        let page_index = self.content_page_index(destination_page)?;
        let source_presentation = source.to_presentation()?;
        let operation = crate::content::prepare_object_transfer(
            source_presentation.owned_package(),
            page_index,
            kind,
            source_name,
            destination_name,
        )?;
        self.stage_content(operation)
    }

    fn content_bytes(&mut self) -> Result<Arc<Vec<u8>>> {
        if let Some(content) = &self.content {
            return Ok(Arc::clone(&content.bytes));
        }
        let package = if self.changed {
            OwnedPackage::from_bytes(self.draft.to_bytes_bounded(MAX_PACKAGE_BYTES)?)?
        } else {
            self.source.package.clone()
        };
        let bytes = package.shared_bytes();
        self.content = Some(ContentDraft {
            bytes: Arc::clone(&bytes),
            package,
            operations: Vec::new(),
        });
        Ok(bytes)
    }

    fn content_package(&mut self) -> Result<OwnedPackage> {
        self.content_bytes()?;
        self.content
            .as_ref()
            .map(|draft| draft.package.clone())
            .ok_or_else(|| invalid_error("ODP semantic-content draft initialization failed"))
    }

    fn stage_content(&mut self, operation: crate::content::Operation) -> Result<()> {
        self.check_no_slide_order_change("rich-content")?;
        let package = self.content_package()?;
        let bytes = crate::content::apply(&package, &operation)?;
        if bytes.len() > MAX_PACKAGE_BYTES {
            return invalid("ODP semantic-content transaction exceeds the 128 MiB package limit");
        }
        let package = OwnedPackage::from_bytes(bytes)?;
        Presentation::from_owned_package(package.clone())?;
        let draft = self
            .content
            .as_mut()
            .ok_or_else(|| invalid_error("ODP semantic-content draft initialization failed"))?;
        draft.bytes = package.shared_bytes();
        draft.package = package;
        draft.operations.push(operation);
        Ok(())
    }

    fn stage_text_box_models(
        &mut self,
        replacements: Vec<crate::content::OwnedTextBoxModelReplacement>,
    ) -> Result<usize> {
        self.check_no_slide_order_change("rich-content")?;
        let package = self.content_package()?;
        let (bytes, changed) =
            crate::content::apply_text_box_model_replacements(&package, &replacements)?;
        let Some(bytes) = bytes else {
            return Ok(0);
        };
        if bytes.len() > MAX_PACKAGE_BYTES {
            return invalid("ODP semantic-content transaction exceeds the 128 MiB package limit");
        }
        let package = OwnedPackage::from_bytes(bytes)?;
        let presentation = Presentation::from_owned_package(package.clone())?;
        crate::content::verify_text_box_model_replacements(
            presentation.owned_package(),
            &replacements,
        )?;
        let draft = self
            .content
            .as_mut()
            .ok_or_else(|| invalid_error("ODP semantic-content draft initialization failed"))?;
        draft.bytes = package.shared_bytes();
        draft.package = package;
        draft
            .operations
            .push(crate::content::Operation::ReplaceTextBoxModels { replacements });
        Ok(changed)
    }

    fn annotation_package(&self) -> Result<OwnedPackage> {
        self.annotations
            .as_ref()
            .map(|draft| draft.package.clone())
            .ok_or_else(|| invalid_error("ODP annotation draft initialization failed"))
    }

    fn stage_annotation(
        &mut self,
        package: OwnedPackage,
        operation: AnnotationOperation,
    ) -> Result<()> {
        self.check_no_slide_order_change("annotation")?;
        let bytes = package.shared_bytes();
        if bytes.len() > MAX_PACKAGE_BYTES {
            return invalid("ODP annotation transaction exceeds the 128 MiB package limit");
        }
        let presentation = Presentation::from_owned_package(package.clone())?;
        let annotations = presentation.annotations()?;
        let draft = self
            .annotations
            .as_mut()
            .ok_or_else(|| invalid_error("ODP annotation draft initialization failed"))?;
        if annotations == draft.annotations {
            return Ok(());
        }
        if annotations == draft.original {
            draft.package = self.source.package.clone();
            draft.annotations = annotations;
            draft.operations.clear();
            return Ok(());
        }
        draft.package = package;
        draft.annotations = annotations;
        draft.operations.push(operation);
        Ok(())
    }

    fn check_text(title: &str, text: &str) -> Result<()> {
        let size = title
            .len()
            .checked_add(text.len())
            .ok_or_else(|| invalid_error("ODP slide text size overflow"))?;
        if size > MAX_TEXT_BYTES {
            return invalid("ODP slide text exceeds the 16 MiB limit");
        }
        if title.chars().chain(text.chars()).any(|value| {
            !matches!(
                value,
                '\u{9}'
                    | '\u{A}'
                    | '\u{D}'
                    | '\u{20}'..='\u{D7FF}'
                    | '\u{E000}'..='\u{FFFD}'
                    | '\u{10000}'..='\u{10FFFF}'
            )
        }) {
            return invalid("ODP slide text contains a character forbidden by XML 1.0");
        }
        Ok(())
    }

    fn check_structure_edit(&self) -> Result<()> {
        if self.draft.has_source_declarations() {
            return unsupported(
                "ODP structural editing with retained header/footer declarations is not lossless",
            );
        }
        Ok(())
    }

    fn check_page_indexed_move_conflict(&self) -> Result<()> {
        let conflict = self.has_page_indexed_operations();
        if conflict {
            return unsupported(
                "ODP slide move conflicts with already-staged page-indexed operations",
            );
        }
        Ok(())
    }

    fn has_page_indexed_operations(&self) -> bool {
        self.charts
            .as_ref()
            .is_some_and(|draft| !draft.operations.is_empty())
            || self
                .design
                .as_ref()
                .is_some_and(|draft| !draft.operations.is_empty())
            || self
                .annotations
                .as_ref()
                .is_some_and(|draft| !draft.operations.is_empty())
            || self
                .content
                .as_ref()
                .is_some_and(|draft| !draft.operations.is_empty())
    }

    fn check_no_slide_order_change(&self, operation: &str) -> Result<()> {
        if self.slide_order_changed {
            return unsupported(format!(
                "ODP {operation} operation cannot be staged after a slide move"
            ));
        }
        if self.dependency_free_slide_copy_changed {
            return unsupported(format!(
                "ODP {operation} operation cannot be staged after a dependency-free blank-slide copy"
            ));
        }
        if self.dependency_free_slide_removal_changed {
            return unsupported(format!(
                "ODP {operation} operation cannot be staged after a dependency-free blank-slide removal"
            ));
        }
        Ok(())
    }

    fn check_slide_rewrite(&self, index: usize) -> Result<()> {
        if self.draft.retains_source_slide(index) {
            return unsupported(
                "ODP retained slide contains XML that cannot yet be proven lossless under rewrite",
            );
        }
        Ok(())
    }

    fn resource_candidate(&self, removed: usize, added: usize) -> Result<usize> {
        let candidate = bounded_candidate(self.resource_bytes, removed, added, MAX_DRAFT_BYTES)?;
        self.check_projected(candidate, self.media_bytes)?;
        Ok(candidate)
    }

    fn check_projected(&self, resource_bytes: usize, media_bytes: usize) -> Result<()> {
        let _projected = projected_size(
            self.source.bytes().len(),
            self.source_resource_bytes,
            resource_bytes,
            media_bytes,
            MAX_PACKAGE_BYTES,
        )?;
        Ok(())
    }
}

/// A validated publication result containing a snapshot and reversible patch.
pub struct Commit {
    snapshot: Snapshot,
    patch: Patch,
    changed: bool,
}

impl Commit {
    fn unchanged(snapshot: Snapshot) -> Self {
        let patch = Patch {
            before: snapshot.clone(),
            after: snapshot.clone(),
            domains: Arc::from([]),
        };
        Self {
            snapshot,
            patch,
            changed: false,
        }
    }

    /// Return whether publication rebuilt the package.
    #[must_use]
    pub const fn changed(&self) -> bool {
        self.changed
    }

    /// Borrow the published immutable snapshot.
    #[must_use]
    pub const fn snapshot(&self) -> &Snapshot {
        &self.snapshot
    }

    /// Borrow the exact-source-checked reversible patch.
    #[must_use]
    pub const fn patch(&self) -> &Patch {
        &self.patch
    }

    /// Consume the commit into its published snapshot.
    #[must_use]
    pub fn into_snapshot(self) -> Snapshot {
        self.snapshot
    }
}

/// An exact-byte-source-checked reversible ODP package patch.
#[derive(Clone)]
pub struct Patch {
    before: Snapshot,
    after: Snapshot,
    domains: Arc<[Domain]>,
}

impl Patch {
    /// Apply this patch only to its exact source package.
    ///
    /// # Errors
    ///
    /// Returns an error when `source` is not byte-for-byte identical to the patch source or
    /// its signature/encryption policy makes it read-only.
    pub fn apply(&self, source: &Snapshot) -> Result<Snapshot> {
        if !same_source(&self.before, source) {
            return invalid("stale ODP presentation patch source");
        }
        ensure_editable_source(&source.package)?;
        Ok(self.after.clone())
    }

    /// Return the patch that restores the exact pre-commit package.
    #[must_use]
    pub fn inverse(&self) -> Self {
        Self {
            before: self.after.clone(),
            after: self.before.clone(),
            domains: Arc::clone(&self.domains),
        }
    }

    /// Return whether this patch preserves the exact package bytes.
    #[must_use]
    pub fn is_noop(&self) -> bool {
        same_source(&self.before, &self.after)
    }

    /// Borrow the sorted semantic dependency domains changed by this patch.
    #[must_use]
    pub fn domains(&self) -> &[Domain] {
        &self.domains
    }

    /// Produce a conservative, non-mutating join plan against another patch.
    ///
    /// RDF-only work is independent from other modeled domains. All edits that
    /// can rewrite `content.xml` are conservatively reported as conflicts until
    /// a semantic operation compositor is available.
    ///
    /// # Errors
    ///
    /// Returns an error unless both patches accept the exact same source package.
    pub fn plan_join(&self, other: &Self) -> Result<MergePlan> {
        if !same_source(&self.before, &other.before) {
            return invalid("ODP patch join requires an exact common source");
        }
        let mut conflicts = Vec::new();
        for domain in self.domains.iter().copied() {
            if other.domains.contains(&domain)
                || (domain != Domain::Rdf
                    && other
                        .domains
                        .iter()
                        .any(|other_domain| *other_domain != Domain::Rdf))
            {
                conflicts.push(domain);
            }
        }
        conflicts.sort_unstable();
        conflicts.dedup();
        Ok(MergePlan { conflicts })
    }

    /// Join two patch intents into a non-mutating merge plan.
    ///
    /// This deliberately does not publish a package: even an independent plan
    /// must be replayed from semantic operations so neither target archive wins
    /// by accident.
    ///
    /// # Errors
    ///
    /// Returns an error unless both patches accept the exact same source package.
    pub fn join(&self, other: &Self) -> Result<MergePlan> {
        self.plan_join(other)
    }

    /// Materialize two patch intents when the semantic planner proves them independent.
    ///
    /// The current compositor has one deliberately narrow safe case: an RDF-only patch can
    /// be replayed over a patch that leaves RDF untouched. Content-coupled domains continue
    /// to be refused rather than selecting one complete target archive as an accidental winner.
    /// Durable patches retain enough source and target state for this operation after reload.
    ///
    /// # Errors
    ///
    /// Returns an error for different sources, a reported conflict, or an RDF delta that cannot
    /// be replayed and verified over the other target.
    pub fn join_snapshot(&self, other: &Self) -> Result<Snapshot> {
        let plan = self.plan_join(other)?;
        if !plan.is_independent() {
            return unsupported("ODP patch join requires resolution of semantic conflicts");
        }
        if self.is_noop() {
            return Ok(other.after.clone());
        }
        if other.is_noop() {
            return Ok(self.after.clone());
        }
        let (rdf_patch, target_patch) = if self.domains.as_ref() == [Domain::Rdf]
            && !other.domains.contains(&Domain::Rdf)
        {
            (self, other)
        } else if other.domains.as_ref() == [Domain::Rdf] && !self.domains.contains(&Domain::Rdf) {
            (other, self)
        } else {
            return unsupported("ODP independent join has no bounded semantic compositor");
        };
        materialize_rdf_join(rdf_patch, target_patch)
    }

    /// Plan a conservative three-way merge rooted at an exact base snapshot.
    ///
    /// # Errors
    ///
    /// Returns an error when either side was not authored against `base`.
    pub fn plan_three_way(base: &Snapshot, left: &Self, right: &Self) -> Result<MergePlan> {
        if !same_source(base, &left.before) || !same_source(base, &right.before) {
            return invalid("ODP three-way merge patches do not share the supplied base");
        }
        left.plan_join(right)
    }

    /// Build a non-mutating three-way merge plan from an exact common base.
    ///
    /// # Errors
    ///
    /// Returns an error when either side was not authored against `base`.
    pub fn three_way(base: &Snapshot, left: &Self, right: &Self) -> Result<MergePlan> {
        Self::plan_three_way(base, left, right)
    }

    /// Materialize a checked three-way merge for the independent cases supported by
    /// [`Self::join_snapshot`].
    ///
    /// # Errors
    ///
    /// Returns an error when either patch is not rooted at `base`, the planner reports a
    /// conflict, or the bounded semantic compositor cannot verify the merged package.
    pub fn three_way_snapshot(base: &Snapshot, left: &Self, right: &Self) -> Result<Snapshot> {
        Self::plan_three_way(base, left, right)?;
        left.join_snapshot(right)
    }

    /// Serialize this exact reversible patch into a deterministic bounded binary envelope.
    ///
    /// The envelope retains both complete package artifacts so stale-source authorization and
    /// byte-exact inversion remain available after process boundaries.
    ///
    /// # Errors
    ///
    /// Returns an allocation or size error when the bounded envelope cannot be materialized.
    pub fn to_durable_bytes(&self) -> Result<Vec<u8>> {
        let before_len = self.before.bytes().len();
        let after_len = self.after.bytes().len();
        let capacity = DURABLE_PATCH_MAGIC
            .len()
            .checked_add(2 + 1 + 8 + 8)
            .and_then(|size| size.checked_add(before_len))
            .and_then(|size| size.checked_add(after_len))
            .ok_or_else(|| invalid_error("ODP durable patch size overflow"))?;
        let maximum = MAX_PACKAGE_BYTES
            .checked_mul(2)
            .and_then(|size| size.checked_add(64))
            .ok_or_else(|| invalid_error("ODP durable patch limit overflow"))?;
        if capacity > maximum {
            return invalid("ODP durable patch exceeds its package-derived size limit");
        }
        let mut output = Vec::new();
        output
            .try_reserve_exact(capacity)
            .map_err(|source| Error::Allocation {
                resource: "ODP durable patch envelope",
                source,
            })?;
        output.extend_from_slice(DURABLE_PATCH_MAGIC);
        output.extend_from_slice(&DURABLE_PATCH_VERSION.to_le_bytes());
        output.push(domain_bits(&self.domains));
        output.extend_from_slice(&u64::try_from(before_len).unwrap_or(u64::MAX).to_le_bytes());
        output.extend_from_slice(&u64::try_from(after_len).unwrap_or(u64::MAX).to_le_bytes());
        output.extend_from_slice(self.before.bytes());
        output.extend_from_slice(self.after.bytes());
        Ok(output)
    }

    /// Rehydrate a deterministic durable patch with full package validation.
    ///
    /// # Errors
    ///
    /// Returns an error for a malformed version, unknown domain, oversized artifact, trailing
    /// bytes, or invalid ODP source/target package.
    pub fn from_durable_bytes(bytes: &[u8]) -> Result<Self> {
        let header_len = DURABLE_PATCH_MAGIC.len() + 2 + 1 + 8 + 8;
        if bytes.len() < header_len || &bytes[..DURABLE_PATCH_MAGIC.len()] != DURABLE_PATCH_MAGIC {
            return invalid("invalid ODP durable patch magic or truncated header");
        }
        let mut offset = DURABLE_PATCH_MAGIC.len();
        let version = read_u16(bytes, &mut offset)?;
        if version != DURABLE_PATCH_VERSION {
            return invalid(format!("unsupported ODP durable patch version {version}"));
        }
        let bits = *bytes
            .get(offset)
            .ok_or_else(|| invalid_error("truncated ODP durable patch domains"))?;
        offset += 1;
        let domains = domains_from_bits(bits)?;
        let before_len = read_len(bytes, &mut offset)?;
        let after_len = read_len(bytes, &mut offset)?;
        if before_len > MAX_PACKAGE_BYTES || after_len > MAX_PACKAGE_BYTES {
            return invalid("ODP durable patch contains an oversized package");
        }
        let expected = offset
            .checked_add(before_len)
            .and_then(|size| size.checked_add(after_len))
            .ok_or_else(|| invalid_error("ODP durable patch length overflow"))?;
        if expected != bytes.len() {
            return invalid("ODP durable patch length does not match its envelope");
        }
        let before_end = offset + before_len;
        let before = Snapshot::from_bytes(bytes[offset..before_end].to_vec())?;
        let after = Snapshot::from_bytes(bytes[before_end..].to_vec())?;
        Ok(Self {
            before,
            after,
            domains: Arc::from(domains),
        })
    }
}

/// Entry- and byte-bounded undo/redo history for immutable ODP snapshots.
pub struct History {
    entries: Vec<Snapshot>,
    cursor: usize,
    max_entries: usize,
    max_bytes: usize,
    retained_bytes: usize,
}

impl History {
    /// Create history rooted at one immutable snapshot.
    ///
    /// # Errors
    ///
    /// Returns an error for zero limits or a byte budget smaller than the root package.
    pub fn new(initial: Snapshot, max_entries: usize, max_bytes: usize) -> Result<Self> {
        if max_entries == 0 || max_bytes == 0 {
            return invalid("ODP history limits must be positive");
        }
        if initial.bytes().len() > max_bytes {
            return invalid("ODP history byte budget cannot retain its initial snapshot");
        }
        let retained_bytes = initial.bytes().len();
        Ok(Self {
            entries: vec![initial],
            cursor: 0,
            max_entries,
            max_bytes,
            retained_bytes,
        })
    }

    /// Borrow the current immutable snapshot.
    #[must_use]
    pub fn current(&self) -> &Snapshot {
        &self.entries[self.cursor]
    }

    /// Record a commit whose exact source is the current history snapshot.
    ///
    /// Redo entries are discarded only after source validation succeeds. Oldest
    /// entries are evicted deterministically to enforce both configured limits.
    ///
    /// # Errors
    ///
    /// Returns an error for a stale commit or a target package larger than the byte budget.
    pub fn record(&mut self, commit: &Commit) -> Result<()> {
        if !same_source(self.current(), &commit.patch.before) {
            return invalid("ODP history commit source is not current");
        }
        if !commit.changed {
            return Ok(());
        }
        let target_bytes = commit.snapshot.bytes().len();
        if target_bytes > self.max_bytes {
            return invalid("ODP history target exceeds the byte budget");
        }
        if self.cursor + 1 < self.entries.len() {
            for removed in self.entries.drain(self.cursor + 1..) {
                self.retained_bytes = self.retained_bytes.saturating_sub(removed.bytes().len());
            }
        }
        self.entries.push(commit.snapshot.clone());
        self.retained_bytes = self
            .retained_bytes
            .checked_add(target_bytes)
            .ok_or_else(|| invalid_error("ODP history byte count overflow"))?;
        self.cursor = self.entries.len() - 1;
        while self.entries.len() > self.max_entries || self.retained_bytes > self.max_bytes {
            let removed = self.entries.remove(0);
            self.retained_bytes = self.retained_bytes.saturating_sub(removed.bytes().len());
            self.cursor = self.cursor.saturating_sub(1);
        }
        Ok(())
    }

    /// Move to the previous retained snapshot.
    #[must_use]
    pub fn undo(&mut self) -> Option<&Snapshot> {
        if self.cursor == 0 {
            return None;
        }
        self.cursor -= 1;
        Some(self.current())
    }

    /// Move to the next retained snapshot.
    #[must_use]
    pub fn redo(&mut self) -> Option<&Snapshot> {
        if self.cursor + 1 >= self.entries.len() {
            return None;
        }
        self.cursor += 1;
        Some(self.current())
    }

    /// Return the number of retained immutable snapshots.
    #[must_use]
    pub fn len(&self) -> usize {
        self.entries.len()
    }

    /// Return whether no snapshots are retained.
    #[must_use]
    pub fn is_empty(&self) -> bool {
        self.entries.is_empty()
    }

    /// Return the exact package-byte accounting used by the history budget.
    #[must_use]
    pub const fn retained_bytes(&self) -> usize {
        self.retained_bytes
    }

    /// Serialize the complete bounded undo/redo timeline into a deterministic envelope.
    ///
    /// Every retained package is included so cursor position, redo state, and exact package
    /// bytes survive a process boundary. The envelope is independently capped at 512 MiB.
    ///
    /// # Errors
    ///
    /// Returns an error when the retained history exceeds the durable bound or a native limit
    /// cannot be represented in the envelope.
    pub fn to_durable_bytes(&self) -> Result<Vec<u8>> {
        let header_len = DURABLE_HISTORY_MAGIC.len() + 2 + (8 * 4);
        let capacity = self.entries.iter().try_fold(header_len, |size, snapshot| {
            size.checked_add(8)
                .and_then(|value| value.checked_add(snapshot.bytes().len()))
                .ok_or_else(|| invalid_error("ODP durable history size overflow"))
        })?;
        if capacity > MAX_DURABLE_HISTORY_BYTES {
            return invalid("ODP durable history exceeds the 512 MiB envelope limit");
        }
        let mut output = Vec::new();
        output
            .try_reserve_exact(capacity)
            .map_err(|source| Error::Allocation {
                resource: "ODP durable history envelope",
                source,
            })?;
        output.extend_from_slice(DURABLE_HISTORY_MAGIC);
        output.extend_from_slice(&DURABLE_HISTORY_VERSION.to_le_bytes());
        write_len(&mut output, self.max_entries)?;
        write_len(&mut output, self.max_bytes)?;
        write_len(&mut output, self.cursor)?;
        write_len(&mut output, self.entries.len())?;
        for snapshot in &self.entries {
            write_len(&mut output, snapshot.bytes().len())?;
            output.extend_from_slice(snapshot.bytes());
        }
        Ok(output)
    }

    /// Rehydrate a bounded undo/redo timeline with full package validation.
    ///
    /// # Errors
    ///
    /// Returns an error for malformed bounds, cursor/count inconsistencies, trailing bytes,
    /// oversized packages, budget violations, or invalid retained ODP artifacts.
    pub fn from_durable_bytes(bytes: &[u8]) -> Result<Self> {
        let header_len = DURABLE_HISTORY_MAGIC.len() + 2 + (8 * 4);
        if bytes.len() < header_len
            || bytes.len() > MAX_DURABLE_HISTORY_BYTES
            || &bytes[..DURABLE_HISTORY_MAGIC.len()] != DURABLE_HISTORY_MAGIC
        {
            return invalid("invalid ODP durable history magic, size, or truncated header");
        }
        let mut offset = DURABLE_HISTORY_MAGIC.len();
        let version = read_u16(bytes, &mut offset)?;
        if version != DURABLE_HISTORY_VERSION {
            return invalid(format!("unsupported ODP durable history version {version}"));
        }
        let max_entries = read_len(bytes, &mut offset)?;
        let max_bytes = read_len(bytes, &mut offset)?;
        let cursor = read_len(bytes, &mut offset)?;
        let count = read_len(bytes, &mut offset)?;
        if max_entries == 0
            || max_bytes == 0
            || count == 0
            || count > max_entries
            || cursor >= count
        {
            return invalid("ODP durable history contains inconsistent bounds or cursor state");
        }
        let mut entries = Vec::new();
        entries
            .try_reserve_exact(count)
            .map_err(|source| Error::Allocation {
                resource: "ODP durable history entries",
                source,
            })?;
        let mut retained_bytes = 0usize;
        for _ in 0..count {
            let length = read_len(bytes, &mut offset)?;
            if length > MAX_PACKAGE_BYTES {
                return invalid("ODP durable history contains an oversized package");
            }
            let end = offset
                .checked_add(length)
                .ok_or_else(|| invalid_error("ODP durable history package offset overflow"))?;
            let package = bytes
                .get(offset..end)
                .ok_or_else(|| invalid_error("truncated ODP durable history package"))?;
            entries.push(Snapshot::from_bytes(package.to_vec())?);
            retained_bytes = retained_bytes
                .checked_add(length)
                .ok_or_else(|| invalid_error("ODP durable history byte count overflow"))?;
            offset = end;
        }
        if offset != bytes.len() || retained_bytes > max_bytes {
            return invalid(
                "ODP durable history length or byte budget does not match its envelope",
            );
        }
        Ok(Self {
            entries,
            cursor,
            max_entries,
            max_bytes,
            retained_bytes,
        })
    }
}

fn select(slides: &[Slide], selector: Selector<'_>) -> Result<Option<usize>> {
    match selector {
        Selector::Index(index) => Ok((index < slides.len()).then_some(index)),
        Selector::Title(title) => {
            let mut selected = None;
            for (index, slide) in slides.iter().enumerate() {
                if slide.title.as_deref() != Some(title) {
                    continue;
                }
                if selected.is_some() {
                    return invalid("ODP slide title selector is ambiguous");
                }
                selected = Some(index);
            }
            Ok(selected)
        },
    }
}

fn same_source(left: &Snapshot, right: &Snapshot) -> bool {
    Arc::ptr_eq(&left.bytes, &right.bytes) || left.bytes == right.bytes
}

fn domain_bits(domains: &[Domain]) -> u8 {
    domains.iter().fold(0u8, |bits, domain| {
        bits | match domain {
            Domain::Slides => 1 << 0,
            Domain::Rdf => 1 << 1,
            Domain::Charts => 1 << 2,
            Domain::Design => 1 << 3,
            Domain::Annotations => 1 << 4,
            Domain::Content => 1 << 5,
        }
    })
}

fn domains_from_bits(bits: u8) -> Result<Vec<Domain>> {
    if bits & !0b11_1111 != 0 {
        return invalid("ODP durable patch contains an unknown semantic domain");
    }
    let mut domains = Vec::new();
    for (mask, domain) in [
        (1 << 0, Domain::Slides),
        (1 << 1, Domain::Rdf),
        (1 << 2, Domain::Charts),
        (1 << 3, Domain::Design),
        (1 << 4, Domain::Annotations),
        (1 << 5, Domain::Content),
    ] {
        if bits & mask != 0 {
            domains.push(domain);
        }
    }
    Ok(domains)
}

fn read_u16(bytes: &[u8], offset: &mut usize) -> Result<u16> {
    let end = offset
        .checked_add(2)
        .ok_or_else(|| invalid_error("ODP durable patch offset overflow"))?;
    let value = bytes
        .get(*offset..end)
        .ok_or_else(|| invalid_error("truncated ODP durable patch version"))?;
    *offset = end;
    Ok(u16::from_le_bytes([value[0], value[1]]))
}

fn read_len(bytes: &[u8], offset: &mut usize) -> Result<usize> {
    let end = offset
        .checked_add(8)
        .ok_or_else(|| invalid_error("ODP durable patch offset overflow"))?;
    let value = bytes
        .get(*offset..end)
        .ok_or_else(|| invalid_error("truncated ODP durable patch length"))?;
    *offset = end;
    let decoded = u64::from_le_bytes([
        value[0], value[1], value[2], value[3], value[4], value[5], value[6], value[7],
    ]);
    usize::try_from(decoded)
        .map_err(|error| invalid_error(format!("ODP durable patch length is not native: {error}")))
}

fn write_len(output: &mut Vec<u8>, value: usize) -> Result<()> {
    let encoded = u64::try_from(value).map_err(|source| {
        invalid_error(format!(
            "ODP durable envelope length is not representable: {source}"
        ))
    })?;
    output.extend_from_slice(&encoded.to_le_bytes());
    Ok(())
}

fn ensure_editable_source(package: &OwnedPackage) -> Result<()> {
    match package_security_policy(package)? {
        SecurityPolicy::Editable => Ok(()),
        SecurityPolicy::EncryptedReadOnly => unsupported(
            "ODP mutation refuses encrypted package entries; decrypt to a new unsigned package first",
        ),
        SecurityPolicy::SignedReadOnly => unsupported(
            "ODP mutation refuses signed packages because publication would invalidate their signatures",
        ),
        SecurityPolicy::SignedAndEncryptedReadOnly => unsupported(
            "ODP mutation refuses signed packages with encrypted entries; decrypt and remove signatures in a new package first",
        ),
    }
}

fn package_security_policy(package: &OwnedPackage) -> Result<SecurityPolicy> {
    let archive = package.package()?;
    let encrypted = archive.manifest().has_encrypted_entries();
    let signed = archive.has_file("META-INF/documentsignatures.xml")
        || archive.has_file("META-INF/macrosignatures.xml");
    Ok(match (signed, encrypted) {
        (false, false) => SecurityPolicy::Editable,
        (false, true) => SecurityPolicy::EncryptedReadOnly,
        (true, false) => SecurityPolicy::SignedReadOnly,
        (true, true) => SecurityPolicy::SignedAndEncryptedReadOnly,
    })
}

fn materialize_rdf_join(rdf_patch: &Patch, target_patch: &Patch) -> Result<Snapshot> {
    let base_package = rdf_patch.before.package.clone();
    let rdf_target_package = rdf_patch.after.package.clone();
    let base_graphs = crate::rdf::graphs(&base_package)?;
    let expected_graphs = crate::rdf::graphs(&rdf_target_package)?;
    let mut package = target_patch.after.package.clone();

    for expected in &expected_graphs {
        let Some(before) = base_graphs.iter().find(|graph| graph.path == expected.path) else {
            continue;
        };
        if before != expected {
            package = OwnedPackage::from_bytes(crate::rdf::replace_graph(
                &package,
                &expected.path,
                &expected.triples,
            )?)?;
        }
    }

    let mut removals = base_graphs
        .iter()
        .filter(|graph| {
            !expected_graphs
                .iter()
                .any(|expected| expected.path == graph.path)
        })
        .map(|graph| graph.path.clone())
        .collect::<Vec<_>>();
    while !removals.is_empty() {
        let mut progress = false;
        let mut retained = Vec::new();
        for path in removals {
            match crate::rdf::remove_graph(&package, &path) {
                Ok(updated) => {
                    package = OwnedPackage::from_bytes(updated)?;
                    progress = true;
                },
                Err(_error) => retained.push(path),
            }
        }
        if !progress {
            return unsupported("ODP RDF join cannot close graph-removal dependencies");
        }
        removals = retained;
    }

    for expected in &expected_graphs {
        if base_graphs.iter().any(|graph| graph.path == expected.path) {
            continue;
        }
        let (updated, actual_path) =
            crate::rdf::add_graph(&package, Some(&expected.path), &expected.triples)?;
        if actual_path != expected.path {
            return invalid("ODP RDF join resolved a different metadata graph path");
        }
        package = OwnedPackage::from_bytes(updated)?;
    }

    if package.as_bytes().len() > MAX_PACKAGE_BYTES {
        return invalid("ODP joined package exceeds the 128 MiB package limit");
    }
    let joined = Snapshot::from_owned_package(package)?;
    if crate::rdf::graphs(&joined.package)? != expected_graphs {
        return invalid("ODP joined package RDF readback differs from the expected graph model");
    }
    Ok(joined)
}

fn apply_rdf_operation(package: &OwnedPackage, operation: &RdfOperation) -> Result<Vec<u8>> {
    match operation {
        RdfOperation::AddGraph { path, triples } => {
            let (bytes, actual_path) = crate::rdf::add_graph(package, Some(path), triples)?;
            if actual_path != *path {
                return invalid("ODP RDF replay resolved a different graph path");
            }
            Ok(bytes)
        },
        RdfOperation::ReplaceGraph { path, triples } => {
            crate::rdf::replace_graph(package, path, triples)
        },
        RdfOperation::RemoveGraph { path } => crate::rdf::remove_graph(package, path),
        RdfOperation::AddTriple { path, triple } => {
            crate::rdf::add_triple(package, path, triple).map(|(bytes, _)| bytes)
        },
        RdfOperation::ReplaceTriple {
            path,
            index,
            triple,
        } => crate::rdf::replace_triple(package, path, *index, triple),
        RdfOperation::RemoveTriple { path, index } => {
            crate::rdf::remove_triple(package, path, *index)
        },
        RdfOperation::MoveTriple { path, from, to } => {
            crate::rdf::move_triple(package, path, *from, *to)
        },
    }
}

fn apply_chart_operation(
    package: &OwnedPackage,
    limits: crate::charts::Limits,
    operation: &ChartOperation,
) -> Result<Vec<u8>> {
    let snapshot = crate::charts::Snapshot::from_owned_package(package.clone(), limits)?;
    let mut edit = snapshot.edit();
    match operation {
        ChartOperation::Replace { selector, part } => {
            edit.replace(selector.borrowed(), part.clone())?;
        },
        ChartOperation::Remove { selector } => {
            let _removed = edit.remove(selector.borrowed())?;
        },
        ChartOperation::Add {
            page,
            name,
            storage,
            part,
        } => {
            let _index = edit.add(page.borrowed(), name.clone(), *storage, part.clone())?;
        },
    }
    edit.commit()
        .map(|commit| commit.snapshot().bytes().to_vec())
}

fn root_charts_equal(left: &[crate::charts::Chart], right: &[crate::charts::Chart]) -> bool {
    left.len() == right.len()
        && left.iter().zip(right).all(|(left_chart, right_chart)| {
            left_chart.frame() == right_chart.frame()
                && left_chart.storage() == right_chart.storage()
                && left_chart.part() == right_chart.part()
        })
}

fn apply_design_operation(package: &OwnedPackage, operation: &DesignOperation) -> Result<Vec<u8>> {
    let mut presentation = Presentation::from_owned_package(package.clone())?;
    match operation {
        DesignOperation::AddLayout(layout) => presentation.add_layout(layout)?,
        DesignOperation::ReplaceLayout(layout) => presentation.replace_page_layout(layout)?,
        DesignOperation::RemoveLayout { name, replacement } => {
            presentation.remove_page_layout(name, replacement.as_deref())?;
        },
        DesignOperation::ReorderLayouts(names) => presentation.reorder_layouts(names)?,
        DesignOperation::AddMaster(master) => presentation.add_master_page(master)?,
        DesignOperation::ReplaceMaster(master) => presentation.replace_master_page(master)?,
        DesignOperation::RemoveMaster { name, replacement } => {
            presentation.remove_master_page(name, replacement.as_deref())?;
        },
        DesignOperation::ReorderMasters(names) => presentation.reorder_master_pages(names)?,
        DesignOperation::AssignSlideMaster { slide_index, name } => {
            presentation.assign_slide_master_page(*slide_index, name.as_deref())?;
        },
        DesignOperation::AssignSlideLayout { slide_index, name } => {
            presentation.assign_slide_page_layout(*slide_index, name.as_deref())?;
        },
    }
    presentation.to_bytes()
}

fn root_masters_equal(left: &[crate::MasterPage], right: &[crate::MasterPage]) -> bool {
    left.len() == right.len()
        && left.iter().zip(right).all(|(left_master, right_master)| {
            left_master.master_page == right_master.master_page
                && left_master.page_layout_name == right_master.page_layout_name
                && left_master.header_name == right_master.header_name
                && left_master.footer_name == right_master.footer_name
                && left_master.date_time_name == right_master.date_time_name
        })
}

fn root_design_pages_equal(
    actual: &crate::page::Collection,
    expected: &crate::page::Collection,
) -> bool {
    expected.pages().iter().all(|expected_page| {
        actual
            .page(expected_page.slide_index)
            .is_some_and(|actual_page| actual_page == expected_page)
    })
}

fn apply_annotation_operation(
    package: &OwnedPackage,
    operation: &AnnotationOperation,
) -> Result<Vec<u8>> {
    let mut presentation = Presentation::from_owned_package(package.clone())?;
    match operation {
        AnnotationOperation::Add { anchor, annotation } => {
            let _index = presentation.add_annotation(anchor, annotation)?;
        },
        AnnotationOperation::Replace { index, annotation } => {
            presentation.replace_annotation(*index, annotation)?;
        },
        AnnotationOperation::Remove { index } => presentation.remove_annotation(*index)?,
    }
    presentation.to_bytes()
}

fn root_annotations_equal(
    left: &[crate::annotation::Info],
    right: &[crate::annotation::Info],
) -> bool {
    left.len() == right.len()
        && left.iter().zip(right).all(|(left_info, right_info)| {
            left_info.index == right_info.index
                && left_info.anchor == right_info.anchor
                && left_info.annotation.attributes() == right_info.annotation.attributes()
                && left_info.annotation.children() == right_info.annotation.children()
        })
}

fn bounded_candidate(current: usize, removed: usize, added: usize, limit: usize) -> Result<usize> {
    let candidate = current
        .checked_sub(removed)
        .and_then(|value| value.checked_add(added))
        .ok_or_else(|| invalid_error("ODP aggregate draft accounting overflow"))?;
    if candidate > limit {
        return invalid("ODP transaction exceeds the aggregate draft limit");
    }
    Ok(candidate)
}

fn projected_size(
    source_bytes: usize,
    source_resource_bytes: usize,
    resource_bytes: usize,
    media_bytes: usize,
    limit: usize,
) -> Result<usize> {
    let growth = resource_bytes.saturating_sub(source_resource_bytes);
    let projected = source_bytes
        .checked_add(growth)
        .and_then(|value| value.checked_add(media_bytes))
        .ok_or_else(|| invalid_error("ODP projected package size overflow"))?;
    if projected > limit {
        return invalid("ODP transaction exceeds the projected package limit");
    }
    Ok(projected)
}

fn try_owned_bytes(value: &[u8], resource: &'static str) -> Result<Vec<u8>> {
    let mut owned = Vec::new();
    owned
        .try_reserve_exact(value.len())
        .map_err(|source| Error::Allocation { resource, source })?;
    owned.extend_from_slice(value);
    Ok(owned)
}

fn try_owned_str(value: &str, resource: &'static str) -> Result<String> {
    let mut owned = String::new();
    owned
        .try_reserve_exact(value.len())
        .map_err(|source| Error::Allocation { resource, source })?;
    owned.push_str(value);
    Ok(owned)
}

fn read_bounded(path: &Path, limit: usize) -> Result<Vec<u8>> {
    let mut file = File::open(path)?;
    let declared = file.metadata()?.len();
    let limit_u64 = u64::try_from(limit).unwrap_or(u64::MAX);
    if declared > limit_u64 {
        return invalid(format!(
            "ODP editing package exceeds the {limit}-byte limit"
        ));
    }
    let capacity = usize::try_from(declared).map_err(|error| {
        invalid_error(format!(
            "ODP package length does not fit this platform: {error}"
        ))
    })?;
    let mut bytes = Vec::new();
    bytes
        .try_reserve_exact(capacity)
        .map_err(|source| Error::Allocation {
            resource: "ODP bounded package input",
            source,
        })?;
    let mut buffer = [0u8; 8 * 1024];
    loop {
        let read = file.read(&mut buffer)?;
        if read == 0 {
            break;
        }
        let candidate = bytes
            .len()
            .checked_add(read)
            .ok_or_else(|| invalid_error("ODP bounded package input size overflow"))?;
        if candidate > limit {
            return invalid(format!(
                "ODP editing package exceeds the {limit}-byte limit"
            ));
        }
        bytes
            .try_reserve_exact(read)
            .map_err(|source| Error::Allocation {
                resource: "ODP bounded package input",
                source,
            })?;
        bytes.extend_from_slice(&buffer[..read]);
    }
    Ok(bytes)
}

fn slides_resource(slides: &[Slide]) -> Result<usize> {
    slides.iter().try_fold(0usize, |total, slide| {
        total
            .checked_add(slide_resource(slide)?)
            .ok_or_else(|| invalid_error("ODP aggregate slide resource size overflow"))
    })
}

fn slide_resource(slide: &Slide) -> Result<usize> {
    let primary = slide_primary_resource(slide)?;
    primary
        .checked_add(slide.notes.as_ref().map_or(0, String::len))
        .ok_or_else(|| invalid_error("ODP slide resource size overflow"))
}

fn slide_primary_resource(slide: &Slide) -> Result<usize> {
    slide_primary_text_resource(slide)?
        .checked_add(slide.shapes.iter().try_fold(0usize, |total, shape| {
            total
                .checked_add(shape_resource(shape)?)
                .ok_or_else(|| invalid_error("ODP shape resource size overflow"))
        })?)
        .ok_or_else(|| invalid_error("ODP slide resource size overflow"))
}

fn slide_primary_text_resource(slide: &Slide) -> Result<usize> {
    text_resource(slide.title.as_deref().unwrap_or_default(), &slide.text)
}

fn text_resource(title: &str, text: &str) -> Result<usize> {
    title
        .len()
        .checked_add(text.len())
        .ok_or_else(|| invalid_error("ODP text resource size overflow"))
}

fn shape_resource(root: &Shape) -> Result<usize> {
    let mut stack = Vec::new();
    stack.try_reserve(1).map_err(|source| Error::Allocation {
        resource: "ODP shape resource stack",
        source,
    })?;
    stack.push(root);
    let mut total = 0usize;
    while let Some(shape) = stack.pop() {
        total = total
            .checked_add(shape.text.len())
            .and_then(|value| value.checked_add(shape.name().map_or(0, str::len)))
            .ok_or_else(|| invalid_error("ODP shape resource size overflow"))?;
        if stack.len().saturating_add(shape.children().len()) > MAX_SLIDES {
            return invalid("ODP shape resource traversal exceeds the node limit");
        }
        stack
            .try_reserve(shape.children().len())
            .map_err(|source| Error::Allocation {
                resource: "ODP shape resource stack",
                source,
            })?;
        stack.extend(shape.children());
    }
    Ok(total)
}

fn validate_compact_xml_parts(package: &OwnedPackage, source: &OwnedPackage) -> Result<()> {
    validate_compact_xml_parts_against(package, Some(source), true)
}

pub(crate) fn validate_raw_preserved_xml_parts(package: &OwnedPackage) -> Result<()> {
    validate_compact_xml_parts_against(package, None, true)
}

fn validate_raw_preserved_referenced_xml_parts(package: &OwnedPackage) -> Result<()> {
    validate_compact_xml_parts_against(package, None, false)
}

fn validate_compact_xml_parts_against(
    package: &OwnedPackage,
    source: Option<&OwnedPackage>,
    audit_core_parts: bool,
) -> Result<()> {
    let mut part_count = 0usize;
    let mut aggregate_bytes = 0usize;
    for path in package.files()? {
        if !path.rsplit_once('.').is_some_and(|(_, extension)| {
            extension.eq_ignore_ascii_case("xml") || extension.eq_ignore_ascii_case("rdf")
        }) {
            continue;
        }
        if !audit_core_parts
            && matches!(
                path.as_str(),
                "content.xml"
                    | "styles.xml"
                    | "meta.xml"
                    | "settings.xml"
                    | "META-INF/manifest.xml"
            )
        {
            continue;
        }
        let payload = package.get_file(&path)?;
        if let Some(source) = source {
            if source.has_file(&path)?
                && source
                    .get_file(&path)
                    .is_ok_and(|source_payload| source_payload == payload)
            {
                continue;
            }
            if let Ok(candidate) = std::str::from_utf8(&payload)
                && litchi_odf_common::package::xml_splice_publication(source, &path, candidate)
                    .is_ok()
            {
                continue;
            }
        }
        part_count = part_count
            .checked_add(1)
            .ok_or_else(|| invalid_error("ODP XML part count overflow"))?;
        aggregate_bytes = aggregate_bytes
            .checked_add(payload.len())
            .ok_or_else(|| invalid_error("ODP aggregate XML size overflow"))?;
        if part_count > MAX_XML_PARTS || aggregate_bytes > MAX_PACKAGE_BYTES {
            return invalid("ODP XML package audit exceeds its aggregate limit");
        }
        let limits = audit::Limits::new(
            MAX_PACKAGE_BYTES,
            512,
            1_000_000,
            250_000,
            MAX_TEXT_BYTES,
            MAX_PACKAGE_BYTES,
        )
        .map_err(|config_error| {
            invalid_error(format!("invalid ODP XML audit limits: {config_error}"))
        })?;
        let audit_payload = payload.strip_prefix(b"\xef\xbb\xbf").unwrap_or(&payload);
        let _report =
            audit::verify(audit_payload, limits).map_err(|audit_error| match audit_error {
                audit::Error::NotCompact(_) => Error::Unsupported(format!(
                    "ODP XML part '{path}' is not compact: {audit_error}"
                )),
                audit::Error::Limit { .. }
                | audit::Error::Encoding { .. }
                | audit::Error::Malformed { .. }
                | audit::Error::Doctype { .. }
                | audit::Error::Allocation
                | _ => Error::InvalidFormat(format!(
                    "ODP XML part '{path}' failed audit: {audit_error}"
                )),
            })?;
    }
    Ok(())
}

fn invalid<T>(message: impl Into<String>) -> Result<T> {
    Err(invalid_error(message))
}

fn validate_package_size(length: usize) -> Result<()> {
    if length > MAX_PACKAGE_BYTES {
        return invalid("ODP editing package exceeds the 128 MiB limit");
    }
    Ok(())
}

fn invalid_error(message: impl Into<String>) -> Error {
    Error::InvalidFormat(message.into())
}

fn unsupported<T>(message: impl Into<String>) -> Result<T> {
    Err(Error::Unsupported(message.into()))
}

#[cfg(test)]
mod tests {
    use super::{
        MAX_PACKAGE_BYTES, bounded_candidate, projected_size, read_bounded, validate_package_size,
    };
    use litchi_core::Result;

    #[test]
    fn bounded_file_reader_accepts_n_and_rejects_n_plus_one() -> Result<()> {
        let file = tempfile::NamedTempFile::new()?;
        std::fs::write(file.path(), b"1234")?;
        assert_eq!(read_bounded(file.path(), 4)?, b"1234");
        std::fs::write(file.path(), b"12345")?;
        assert!(read_bounded(file.path(), 4).is_err());
        Ok(())
    }

    #[test]
    fn aggregate_accounting_accepts_n_and_rejects_n_plus_one() -> Result<()> {
        assert_eq!(bounded_candidate(3, 0, 1, 4)?, 4);
        assert!(bounded_candidate(3, 0, 2, 4).is_err());
        assert_eq!(projected_size(3, 2, 2, 1, 4)?, 4);
        assert!(projected_size(3, 2, 2, 2, 4).is_err());
        Ok(())
    }

    #[test]
    fn owned_package_size_gate_accepts_n_and_rejects_n_plus_one() -> Result<()> {
        validate_package_size(MAX_PACKAGE_BYTES)?;
        assert!(validate_package_size(MAX_PACKAGE_BYTES + 1).is_err());
        Ok(())
    }
}
