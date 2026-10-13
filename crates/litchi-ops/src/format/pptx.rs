//! Handle for PowerPoint OOXML (`.pptx`) inputs.
//!
//! [`Pptx`] is one `Arc` over the opened presentation (record 0003 §2.2). The
//! state holds one of two owners (§2.3):
//!
//! - the source-backed [`SourceBackedPresentation`](litchi_pptx::SourceBackedPresentation),
//!   which every path, byte and prepared route builds today. It keeps the
//!   positional source, the indexed OPC catalog and the validated slide graph,
//!   and defers slide and media payloads until a read selects them;
//! - the eager [`Package`](litchi_pptx::Package), which the eager detection
//!   route ([`DetectedFormat::Pptx`](crate::detect::DetectedFormat)) builds
//!   from an already parsed OPC package.
//!
//! Core-properties metadata is a lazily filled sidecar for the source owner
//! and is read at open for the eager owner, as the facade does today (§1.6).
//!
//! # Routes
//!
//! The constructors take what detection already produced, so an open does
//! one detection pass and one container index:
//!
//! | Open route | Detection result | Constructor |
//! |---|---|---|
//! | `open` (path) on `unix`/`windows`, `open_bytes`, `open_prepared` | source-backed OPC package | `Pptx::from_source_package` |
//! | [`SourceMode::Eager`](crate::SourceMode::Eager), and paths elsewhere | eager `OpcPackage` | `Pptx::from_opc_package` |
//! | `Locked::unlock` | decrypted package | `Pptx::from_package` |
//!
//! # Source freshness
//!
//! Every source-backed read first checks the exact source revision and the
//! caller's execution policy (`PptxOwner::check_source`, the facade's
//! up-front `check_source` calls), and the leaf checks the revision again
//! after each payload read. A source changed after open is therefore reported
//! as a typed [`Code::IoError`](crate::Code::IoError) whose leaf error is the
//! OPC `SourceChanged` (finding F10: the facade flattened these to a string
//! in `text`, `slide_count`, `slides`, `slide`, `metadata` and
//! `tolerance_report`).

use std::fmt;
use std::sync::{Arc, Mutex, OnceLock};

use litchi_core::Metadata;
use litchi_core::detection::FileFormat;
use litchi_opc::PackURI;

use super::signature::{self, Opened, Signature};
use crate::error::{Error, Result};
use crate::options::OpenOptions;
use crate::{opc, pptx};

mod children;
mod comments;
pub(crate) mod create;
mod nav;
mod node;
mod order;
mod owner;
mod picture;
mod publish;
mod remove;
mod shape_add;
mod spec;
mod steps;
mod structure;
mod text;
mod views;
pub(crate) mod write;

pub(crate) use children::Children;
pub(crate) use nav::Nav;
pub use node::Node;
pub(crate) use node::Props;
pub(crate) use spec::{
    CELL_CAPABILITY_PROPS, COMMENT_CAPABILITY_PROPS, NOTES_CAPABILITY_PROPS,
    PARAGRAPH_CAPABILITY_PROPS, PICTURE_CAPABILITY_PROPS, PLACEHOLDER_CAPABILITY_PROPS,
    RUN_CAPABILITY_PROPS, SHAPE_CAPABILITY_PROPS, SLIDE_CAPABILITY_PROPS, SPEC,
};
pub(crate) use views::{Json as ViewJson, view_json, view_to};

/// An opened PowerPoint OOXML presentation.
///
/// Cloning is one reference-count increment; clones share the opened source,
/// its indexed catalog and the cached metadata.
#[derive(Clone)]
pub struct Pptx(Arc<Opened<PptxState>>);

const _: () = {
    const fn assert_handle<T: Send + Sync + Clone>() {}
    assert_handle::<Pptx>();
    assert!(size_of::<Pptx>() == size_of::<usize>());
};

struct PptxState {
    owner: PptxOwner,
    source_guard: Option<pptx::SourceBackedPresentation>,
    /// Core-properties metadata. Filled at open for the eager owner and on
    /// the first successful [`Pptx::metadata`] call for the source owner; a
    /// failed read leaves it unset so a later call can retry.
    metadata: OnceLock<Option<Metadata>>,
    /// Whether the opened source stays available for an editor
    /// ([`OpenOptions::retain_source`]).
    retain_source: bool,
    /// Physical OPC limits retained with immutable facade revisions.
    /// Source promotion reuses the native package's admitted limits.
    ooxml_limits: opc::SourceLimits,
    /// The validated eager presentation catalog, retaining each ordered slide
    /// ID and its resolved physical target.  Selection can therefore build a
    /// borrowed leaf slide directly from the OPC graph without reparsing the
    /// presentation root for every node read.
    eager_catalog: OnceLock<Box<[EagerSlide]>>,
    /// Successful native source promotion, shared by readers and editors.
    promotion: OnceLock<Mutex<Option<pptx::source::Package>>>,
    /// Current native revision, including slide topology, order and text overlays.
    revision: Option<pptx::source::Package>,
    /// Explicit signature removal never replaces the original source cache.
    unsigned_promotion: OnceLock<Mutex<Option<Arc<pptx::Package>>>>,
}

struct EagerSlide {
    id: u32,
    target: PackURI,
}

/// The two states of the PPTX owner (record 0003 §2.3).
enum PptxOwner {
    /// Source-backed, read-only. Its package is the retained source an editor
    /// is built from (L-PPTX-1).
    Source(pptx::SourceBackedPresentation),
    /// Eager, editable through `edit`.
    Eager(Arc<pptx::Package>),
}

impl PptxOwner {
    /// The TOCTOU guard: verifies the caller's execution policy and that the
    /// positional source still has the revision captured at open. Eager
    /// package checks are supplemented by `Pptx::check_owner` when promotion
    /// retains an original source guard.
    fn check_source(&self) -> Result<()> {
        match self {
            Self::Source(presentation) => presentation.check_source().map_err(Error::from),
            Self::Eager(_) => Ok(()),
        }
    }
}

/// A point-in-time descriptor of one slide in the presentation's catalog.
///
/// The position is zero-based and the slide ID is the producer-visible
/// `p:sldId@id` value. Building a descriptor reads no slide XML.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct SlideDescriptor {
    position: usize,
    slide_id: u32,
}

impl SlideDescriptor {
    const fn new(position: usize, slide_id: u32) -> Self {
        Self { position, slide_id }
    }

    /// Zero-based position in the presentation's ordered slide list.
    #[must_use]
    pub const fn position(&self) -> usize {
        self.position
    }

    /// Stable producer-visible `p:sldId@id` value.
    #[must_use]
    pub const fn slide_id(&self) -> u32 {
        self.slide_id
    }
}

/// The text and name of one slide, as the facade's `Slide::Pptx` carries them.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct PptxSlide {
    text: String,
    name: String,
}

impl PptxSlide {
    /// The slide's text, in reading order.
    #[must_use]
    pub fn text(&self) -> &str {
        &self.text
    }

    /// The slide's name (`p:cSld@name`), empty when the slide has none.
    #[must_use]
    pub fn name(&self) -> &str {
        &self.name
    }

    /// The slide's text and name, moved out without a copy.
    pub(crate) fn into_parts(self) -> (String, String) {
        (self.text, self.name)
    }
}

/// An immutable native owner shared by facade edit operations.
#[derive(Clone)]
pub(crate) enum Editable {
    Source(pptx::source::Package),
    Eager(Arc<pptx::Package>),
}

impl Pptx {
    pub(crate) fn origin(&self) -> Option<Arc<std::path::Path>> {
        self.0.origin()
    }
    pub(crate) fn with_origin(&self, origin: Option<Arc<std::path::Path>>) -> Self {
        Self(self.0.with_origin(origin))
    }
    #[cfg(feature = "encryption")]
    pub(crate) fn requires_reencryption(&self) -> bool {
        matches!(&self.0.owner, PptxOwner::Eager(package) if package.encryption().is_some())
    }
    pub(crate) fn check_source(&self) -> Result<()> {
        if self.signature() == Signature::Signed {
            return Err(signature::refusal(FileFormat::Pptx));
        }
        #[cfg(feature = "encryption")]
        if self.requires_reencryption() {
            return Err(Error::new(
                crate::Code::DocumentProtected,
                "the encrypted PPTX owner refuses native edits",
            )
            .with_format(FileFormat::Pptx));
        }
        self.check_owner()
    }
    pub(crate) fn same_snapshot(&self, other: &Self) -> bool {
        Arc::ptr_eq(&self.0, &other.0)
    }

    fn check_owner(&self) -> Result<()> {
        if let Some(source) = &self.0.source_guard {
            source.check_source()?;
        }
        self.0.owner.check_source()
    }

    fn adopt_eager(&self, package: Arc<pptx::Package>) -> Self {
        let source_guard = match &self.0.owner {
            PptxOwner::Source(source) => Some(source.clone()),
            PptxOwner::Eager(_) => self.0.source_guard.clone(),
        };
        Self(Opened::new(PptxState {
            owner: PptxOwner::Eager(package),
            source_guard,
            metadata: self.0.metadata.clone(),
            retain_source: self.0.retain_source,
            ooxml_limits: self.0.ooxml_limits,
            eager_catalog: OnceLock::new(),
            promotion: OnceLock::new(),
            revision: None,
            unsigned_promotion: OnceLock::new(),
        }))
        .with_origin(self.origin())
    }

    pub(crate) fn write_to(&self, writer: impl std::io::Write) -> Result<()> {
        self.finish(|| {
            if self.signature() == Signature::RemovalAuthorized {
                return Ok(self.promote_unsigned()?.write_to(writer)?);
            }
            if let Some(revision) = &self.0.revision {
                return Ok(revision.write_to(writer)?);
            }
            match &self.0.owner {
                PptxOwner::Source(source) => Ok(source
                    .source_backed_package()
                    .source_artifact()
                    .write_to_stream(writer)?),
                PptxOwner::Eager(package) => Ok(package.write_to(writer)?),
            }
        })
    }

    /// Whether this handle retains an edit source, without reading any payload.
    pub(crate) fn retains_source(&self) -> bool {
        self.0.retain_source
    }

    fn new(owner: PptxOwner, metadata: OnceLock<Option<Metadata>>, options: &OpenOptions) -> Self {
        Self(Opened::new(PptxState {
            owner,
            source_guard: None,
            metadata,
            retain_source: options.retain_source,
            ooxml_limits: options.ooxml_limits.unwrap_or_default(),
            eager_catalog: OnceLock::new(),
            promotion: OnceLock::new(),
            revision: None,
            unsigned_promotion: OnceLock::new(),
        }))
    }

    /// The same presentation under another signature disposition (record
    /// 0013).
    pub(crate) fn with_signature(&self, signature: Signature) -> Self {
        Self(self.0.with_signature(signature))
    }

    /// Adopt a native source commit without serializing or reopening.
    /// Callers must supply a committed revision of this handle's native owner.
    pub(crate) fn adopt_source(&self, revision: pptx::source::Package) -> Result<Self> {
        self.check_owner()?;
        let PptxOwner::Source(presentation) = &self.0.owner else {
            return Err(Error::from(pptx::Error::Invalid(
                "source revision requires a source owner".into(),
            )));
        };
        let Editable::Source(current) = self.editable()? else {
            return Err(Error::from(pptx::Error::Invalid(
                "source revision requires a retained source editor".into(),
            )));
        };
        if !current.same_source_owner(&revision) {
            return Err(Error::from(pptx::Error::PatchConflict));
        }
        // Check the candidate's own freshness and execution fence before adoption.
        revision.presentation()?;
        let handle = Self(Opened::new(PptxState {
            owner: PptxOwner::Source(presentation.clone()),
            source_guard: self.0.source_guard.clone(),
            metadata: self.0.metadata.clone(),
            retain_source: self.0.retain_source,
            ooxml_limits: self.0.ooxml_limits,
            eager_catalog: OnceLock::new(),
            promotion: OnceLock::new(),
            revision: Some(revision),
            unsigned_promotion: OnceLock::new(),
        }))
        .with_signature(self.signature())
        .with_origin(self.origin());
        self.check_owner()?;
        Ok(handle)
    }

    /// This handle's signature disposition (record 0013).
    pub(crate) fn signature(&self) -> Signature {
        self.0.signature()
    }

    /// Adopts the source-backed presentation built by the path route
    /// (`PptxSourcePathDetection::Pptx`) or the byte route
    /// (`PresentationSourceBytesDetection::Pptx`). Nothing is read: metadata
    /// stays deferred until [`Self::metadata`].
    pub(crate) fn from_source(
        presentation: pptx::SourceBackedPresentation,
        options: &OpenOptions,
    ) -> Self {
        Self::new(PptxOwner::Source(presentation), OnceLock::new(), options)
    }

    /// Builds the source-backed presentation over the OPC package indexed by
    /// the unified detector, without a second catalog pass.
    pub(crate) fn from_source_package(
        package: opc::SourceBackedPackage,
        options: &OpenOptions,
    ) -> Result<Self> {
        let presentation = pptx::SourceBackedPresentation::from_source_backed_package(package)?;
        Ok(Self::from_source(presentation, options))
    }

    /// Adopts the OPC package the eager detector already parsed
    /// (`DetectedFormat::Pptx`). The core properties are read now, as the
    /// facade does, and a properties part the reader rejects fails the open.
    pub(crate) fn from_opc_package(
        package: opc::OpcPackage,
        options: &OpenOptions,
    ) -> Result<Self> {
        // `Package::from_opc_package` reads the same core properties and
        // folds their repairs into its own report, so this read's report is
        // not kept.
        let metadata = litchi_ooxml_common::properties::read_with_report(
            &package,
            &mut litchi_ooxml_common::leniency::ToleranceReport::default(),
        )
        .map_err(pptx::Error::from)?
        .map(Metadata::from)
        .filter(Metadata::has_data);
        let package = pptx::Package::from_opc_package(package)?;
        Ok(Self::new(
            PptxOwner::Eager(Arc::new(package)),
            OnceLock::from(metadata),
            options,
        ))
    }

    /// Adopts an owning package another route built, reading its core
    /// properties now as [`Self::from_opc_package`] does.
    /// Creation adopts its canonical native graph here. `Locked::unlock`
    /// also uses this route, retaining the leaf's encryption provenance.
    pub(crate) fn from_package(package: pptx::Package, options: &OpenOptions) -> Result<Self> {
        let metadata = litchi_ooxml_common::properties::read_with_report(
            package.opc()?,
            &mut litchi_ooxml_common::leniency::ToleranceReport::default(),
        )
        .map_err(pptx::Error::from)?
        .map(Metadata::from)
        .filter(Metadata::has_data);
        Ok(Self::new(
            PptxOwner::Eager(Arc::new(package)),
            OnceLock::from(metadata),
            options,
        ))
    }
}

impl Pptx {
    /// Whether the presentation is source-backed (`true`) or was read
    /// eagerly into memory (`false`).
    #[must_use]
    pub fn is_source_backed(&self) -> bool {
        matches!(self.0.owner, PptxOwner::Source(_))
    }

    /// The source owner's cold payload loads so far, for the tests that
    /// prove a read stays catalog-only; `None` for the eager owner.
    #[cfg(test)]
    pub(crate) fn cold_loads(&self) -> Option<u64> {
        match &self.0.owner {
            PptxOwner::Source(presentation) => Some(presentation.cache_diagnostics().cold_loads),
            PptxOwner::Eager(_) => None,
        }
    }

    /// All slide text, slides in presentation order, separated by a blank
    /// line. Slides without text are skipped.
    ///
    /// # Errors
    ///
    /// A source change since open, cancellation, or a slide the reader
    /// rejects.
    pub fn text(&self) -> Result<String> {
        self.finish(|| {
            if let Some(revision) = &self.0.revision {
                return self.finish(|| {
                    let presentation = revision.presentation()?;
                    let mut texts = Vec::new();
                    for position in 0..presentation.slide_count()? {
                        let text = presentation.slide(position)?.text()?;
                        if !text.is_empty() {
                            texts.push(text);
                        }
                    }
                    Ok(texts.join("\n\n"))
                });
            }
            let mut texts = Vec::new();
            match &self.0.owner {
                PptxOwner::Source(presentation) => {
                    self.check_owner()?;
                    for slide in presentation.slides() {
                        let text = slide.text()?;
                        if !text.is_empty() {
                            texts.push(text);
                        }
                    }
                },
                PptxOwner::Eager(package) => {
                    let count = self.eager_catalog(package)?.len();
                    for position in 0..count {
                        let slide = eager_slide(self, package, position)?.ok_or_else(|| {
                            Error::from(pptx::Error::Invalid(
                                "validated eager slide catalog lost its selected part".into(),
                            ))
                        })?;
                        let text = slide.text()?;
                        if !text.is_empty() {
                            texts.push(text);
                        }
                    }
                },
            }
            Ok(texts.join("\n\n"))
        })
    }

    /// The number of slides.
    ///
    /// # Errors
    ///
    /// A source change since open, cancellation, or (eager owner) a slide
    /// list the reader rejects.
    pub fn slide_count(&self) -> Result<usize> {
        self.finish(|| {
            if let Some(revision) = &self.0.revision {
                return Ok(revision.presentation()?.slide_count()?);
            }
            match &self.0.owner {
                PptxOwner::Source(presentation) => Ok(presentation.slide_count()),
                PptxOwner::Eager(package) => Ok(self.eager_catalog(package)?.len()),
            }
        })
    }

    /// The text and name of every slide, in presentation order.
    ///
    /// # Errors
    ///
    /// As [`Self::text`].
    pub fn slides(&self) -> Result<Vec<PptxSlide>> {
        self.finish(|| {
            if self.0.revision.is_some() {
                return (0..self.slide_count()?)
                    .map(|position| {
                        self.slide(position)?
                            .ok_or_else(|| Error::from(pptx::Error::PatchConflict))
                    })
                    .collect();
            }
            match &self.0.owner {
                PptxOwner::Source(presentation) => {
                    self.check_owner()?;
                    presentation
                        .slides()
                        .map(|slide| {
                            let (text, name) = slide.text_and_name()?;
                            Ok(PptxSlide { text, name })
                        })
                        .collect()
                },
                PptxOwner::Eager(package) => {
                    let count = self.eager_catalog(package)?.len();
                    (0..count)
                        .map(|position| {
                            let slide = eager_slide(self, package, position)?.ok_or_else(|| {
                                Error::from(pptx::Error::Invalid(
                                    "validated eager slide catalog lost its selected part".into(),
                                ))
                            })?;
                            Ok(PptxSlide {
                                text: slide.text()?,
                                name: slide.name()?,
                            })
                        })
                        .collect()
                },
            }
        })
    }

    /// The text and name of the slide at zero-based `position`, or `None`
    /// past the last slide.
    ///
    /// # Errors
    ///
    /// As [`Self::text`], for the selected slide only.
    pub fn slide(&self, position: usize) -> Result<Option<PptxSlide>> {
        self.finish(|| {
            if let Some(revision) = &self.0.revision {
                return self.finish(|| {
                    let presentation = revision.presentation()?;
                    if position >= presentation.slide_count()? {
                        return Ok(None);
                    }
                    let (text, name) = presentation.slide(position)?.text_and_name()?;
                    Ok(Some(PptxSlide { text, name }))
                });
            }
            match &self.0.owner {
                PptxOwner::Source(presentation) => {
                    self.check_owner()?;
                    let Some(slide) = presentation.slide(position) else {
                        return Ok(None);
                    };
                    let (text, name) = slide.text_and_name()?;
                    Ok(Some(PptxSlide { text, name }))
                },
                PptxOwner::Eager(package) => {
                    let Some(slide) = eager_slide(self, package, position)? else {
                        return Ok(None);
                    };
                    Ok(Some(PptxSlide {
                        text: slide.text()?,
                        name: slide.name()?,
                    }))
                },
            }
        })
    }

    /// The ordered slide catalog. No slide XML is read.
    ///
    /// # Errors
    ///
    /// A source change or cancellation observed before, during or after the
    /// walk; an allocation failure.
    pub fn slide_catalog(&self) -> Result<Vec<SlideDescriptor>> {
        self.finish(|| match &self.0.owner {
            PptxOwner::Source(presentation) => {
                self.check_owner()?;
                let count = self.slide_count()?;
                let mut descriptors = reserve_slide_catalog(count)?;
                let revised = self
                    .0
                    .revision
                    .as_ref()
                    .map(|revision| revision.presentation())
                    .transpose()?;
                for position in 0..count {
                    self.check_owner()?;
                    let id = if let Some(revised) = &revised {
                        revised.slide_id(position)?
                    } else {
                        presentation
                            .slide(position)
                            .ok_or_else(|| {
                                Error::from(pptx::Error::Invalid(
                                    "source slide binding is missing".into(),
                                ))
                            })?
                            .slide_id()
                    };
                    descriptors.push(SlideDescriptor::new(position, id));
                }
                self.check_owner()?;
                Ok(descriptors)
            },
            PptxOwner::Eager(package) => {
                let catalog = self.eager_catalog(package)?;
                let mut descriptors = reserve_slide_catalog(catalog.len())?;
                for (position, slide) in catalog.iter().enumerate() {
                    descriptors.push(SlideDescriptor::new(position, slide.id));
                }
                Ok(descriptors)
            },
        })
    }

    /// The slide size `(width, height)` in EMUs, from the presentation root.
    ///
    /// # Errors
    ///
    /// A source change or cancellation, or a malformed `p:sldSz`.
    pub fn slide_size(&self) -> Result<(i64, i64)> {
        self.finish(|| match &self.0.owner {
            PptxOwner::Source(presentation) => Ok(presentation.slide_size()?),
            PptxOwner::Eager(package) => Ok(package.presentation()?.slide_size()?),
        })
    }

    /// The core-properties metadata, or `None` when the package has no core
    /// properties or they carry no data.
    ///
    /// For a source-backed presentation the first successful call reads the
    /// core-properties part and caches the result; every call, cached or not,
    /// first checks that the source is unchanged.
    ///
    /// # Errors
    ///
    /// A source change since open, cancellation, or a core-properties part
    /// the reader rejects.
    pub fn metadata(&self) -> Result<Option<Metadata>> {
        self.finish(|| {
            self.check_owner()?;
            if let Some(cached) = self.0.metadata.get() {
                return Ok(cached.clone());
            }
            let PptxOwner::Source(presentation) = &self.0.owner else {
                // The eager owner filled the cell at open.
                return Ok(None);
            };
            // A failed read is not cached, so a caller can retry. The repairs
            // this read applies are reported by `tolerance_report`.
            let metadata = presentation
                .properties_with_tolerance(&mut pptx::ToleranceReport::default())?
                .map(Metadata::from)
                .filter(Metadata::has_data);
            let _ = self.0.metadata.set(metadata.clone());
            Ok(metadata)
        })
    }

    /// Every defect the reader repaired: the open-time report, plus (source
    /// owner) the core-properties repairs, plus the repairs of reading every
    /// slide's text.
    ///
    /// # Errors
    ///
    /// A source change since open, cancellation, or a part the reader
    /// rejects.
    pub fn tolerance_report(&self) -> Result<pptx::ToleranceReport> {
        self.finish(|| match &self.0.owner {
            PptxOwner::Source(presentation) => {
                self.check_owner()?;
                let mut report = presentation.tolerance_report().clone();
                presentation.properties_with_tolerance(&mut report)?;
                if let Some(revision) = &self.0.revision {
                    let revised = revision.presentation()?;
                    for position in 0..revised.slide_count()? {
                        revised.slide(position)?.text_with_tolerance(&mut report)?;
                    }
                } else {
                    for slide in presentation.slides() {
                        slide.text_with_tolerance(&mut report)?;
                    }
                }
                self.check_owner()?;
                Ok(report)
            },
            PptxOwner::Eager(package) => {
                let mut report = package.tolerance_report().clone();
                for slide in package.presentation()?.slides()? {
                    slide.text_with_tolerance(&mut report)?;
                }
                Ok(report)
            },
        })
    }

    /// Share an immutable native owner for editing without consuming readers.
    /// Successful source promotion reuses the admitted OPC catalog. Failures
    /// are retryable; cached results still enforce freshness and execution.
    pub(crate) fn editable(&self) -> Result<Editable> {
        let owner = self.editable_owner()?;
        if let Editable::Eager(package) = &owner {
            // Admission-only callers still require the native strict checks.
            drop(package.edit()?);
        }
        Ok(owner)
    }

    /// Acquire the retained owner for an immediately following native edit.
    /// The caller must construct that edit even for an empty/no-op request:
    /// it performs encryption and strict graph admission exactly once.
    fn editable_owner(&self) -> Result<Editable> {
        if self.0.signature() == Signature::Signed {
            return Err(signature::refusal(FileFormat::Pptx));
        }
        if !self.0.retain_source {
            return Err(source_not_retained());
        }
        self.check_source()?;
        if self.0.signature() == Signature::RemovalAuthorized {
            return self.promote_unsigned().map(Editable::Eager);
        }
        if let Some(revision) = &self.0.revision {
            return self.finish(|| Ok(Editable::Source(revision.clone())));
        }
        match &self.0.owner {
            PptxOwner::Source(presentation) => {
                let cache = self.0.promotion.get_or_init(|| Mutex::new(None));
                let mut cached = cache
                    .lock()
                    .unwrap_or_else(std::sync::PoisonError::into_inner);
                if let Some(package) = cached.as_ref() {
                    return self.finish(|| Ok(Editable::Source(package.clone())));
                }
                let package = self.finish(|| {
                    pptx::source::Package::from_presentation(presentation).map_err(Error::from)
                })?;
                *cached = Some(package.clone());
                if let Err(error) = self.check_owner() {
                    *cached = None;
                    return Err(error);
                }
                Ok(Editable::Source(package))
            },
            PptxOwner::Eager(package) => Ok(Editable::Eager(Arc::clone(package))),
        }
    }

    fn promote_unsigned(&self) -> Result<Arc<pptx::Package>> {
        let cache = self.0.unsigned_promotion.get_or_init(|| Mutex::new(None));
        let mut cached = cache
            .lock()
            .unwrap_or_else(std::sync::PoisonError::into_inner);
        if let Some(package) = cached.as_ref() {
            return self.finish(|| Ok(Arc::clone(package)));
        }
        let package = self.finish(|| self.unsigned_package()).map(Arc::new)?;
        *cached = Some(Arc::clone(&package));
        if let Err(error) = self.check_owner() {
            *cached = None;
            return Err(error);
        }
        Ok(package)
    }

    /// Return the validated eager catalog, filling the shared cache once.
    fn eager_catalog(&self, package: &pptx::Package) -> Result<&[EagerSlide]> {
        if let Some(catalog) = self.0.eager_catalog.get() {
            return Ok(catalog);
        }
        let presentation = package.presentation()?;
        let references = presentation.slide_references()?;
        let relationships = presentation.part().part().rels();
        let mut catalog = Vec::new();
        catalog
            .try_reserve_exact(references.len())
            .map_err(|source| {
                Error::from(litchi_core::Error::Allocation {
                    resource: "PPTX eager slide catalog",
                    source,
                })
                .with_format(FileFormat::Pptx)
            })?;
        for reference in references {
            let relationship = relationships
                .get(reference.relationship_id())
                .ok_or_else(|| {
                    Error::from(pptx::Error::Relationship(
                        "validated eager slide catalog lost its relationship".into(),
                    ))
                })?;
            let target = relationship
                .target_partname()
                .map_err(pptx::Error::from)
                .map_err(Error::from)?;
            catalog.push(EagerSlide {
                id: reference.id(),
                target,
            });
        }
        Ok(self.0.eager_catalog.get_or_init(|| catalog.into()))
    }

    /// An owning copy of the presentation whose signature the leaf removed
    /// explicitly ([`pptx::Package::unsign`]).
    fn unsigned_package(&self) -> Result<pptx::Package> {
        let mut package = match &self.0.owner {
            PptxOwner::Source(presentation) => pptx::Package::from_opc_package(
                presentation
                    .source_backed_package()
                    .to_opc_package_preserving_source()?,
            )?,
            PptxOwner::Eager(package) => {
                // A copy of the graph would drop the encryption profile the
                // eager owner keeps, and the leaf refuses to mutate a package
                // it decrypted; refuse rather than publish it in the clear.
                #[cfg(feature = "encryption")]
                if package.encryption().is_some() {
                    return Err(Error::new(
                        crate::Code::UnsupportedCapability,
                        "an unlocked encrypted PPTX presentation cannot be edited, signed or not",
                    )
                    .with_format(FileFormat::Pptx));
                }
                pptx::Package::from_opc_package(package.opc()?.clone())?
            },
        };
        package.unsign()?;
        Ok(package)
    }
}

impl fmt::Debug for Pptx {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("Pptx")
            .field("source_backed", &self.is_source_backed())
            .field("retain_source", &self.0.retain_source)
            .finish_non_exhaustive()
    }
}

fn eager_slide<'a>(
    owner: &Pptx,
    package: &'a pptx::Package,
    position: usize,
) -> Result<Option<pptx::slide::Slide<'a>>> {
    let Some(target) = owner
        .eager_catalog(package)?
        .get(position)
        .map(|slide| &slide.target)
    else {
        return Ok(None);
    };
    let opc = package.opc()?;
    let part = opc.get_part(target).map_err(pptx::Error::from)?;
    let part = pptx::parts::SlidePart::from_part(part)?;
    Ok(Some(pptx::slide::Slide::new(opc, part)))
}

fn reserve_slide_catalog(capacity: usize) -> Result<Vec<SlideDescriptor>> {
    let mut descriptors = Vec::new();
    descriptors.try_reserve_exact(capacity).map_err(|source| {
        Error::from(litchi_core::Error::Allocation {
            resource: "PPTX slide descriptors",
            source,
        })
        .with_format(FileFormat::Pptx)
    })?;
    Ok(descriptors)
}

fn source_not_retained() -> Error {
    Error::new(
        crate::Code::UnsupportedCapability,
        "the presentation was opened without retaining its source, so it cannot be edited",
    )
    .with_format(FileFormat::Pptx)
    .with_suggestion("open it again with `OpenOptions::retain_source(true)`")
}

#[cfg(test)]
mod tests;
