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
//! | Facade route today | Detection result | Constructor |
//! |---|---|---|
//! | `Presentation::open` on `unix`/`windows` | `PptxSourcePathDetection::Pptx` | `Pptx::from_source` |
//! | `Presentation::from_bytes` | `PresentationSourceBytesDetection::Pptx` | `Pptx::from_source` |
//! | `Presentation::from_prepared` | `PreparedOoxmlSource` | `Pptx::from_prepared` |
//! | `from_detected` (eager detector) | `DetectedFormat::Pptx` | `Pptx::from_opc_package` |
//! | unified detector (record 0003 §2.7) | `SourceBackedPackage` | `Pptx::from_source_package` |
//!
//! The facade has no reader route for presentations. `Pptx::path_detection_error`
//! maps the path detector's failure without flattening it.
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
use std::sync::{Arc, OnceLock};

use litchi_core::Metadata;
use litchi_core::detection::FileFormat;

use crate::error::{Error, Result};
use crate::options::OpenOptions;
use crate::{opc, pptx};

/// An opened PowerPoint OOXML presentation.
///
/// Cloning is one reference-count increment; clones share the opened source,
/// its indexed catalog and the cached metadata.
#[derive(Clone)]
pub struct Pptx(Arc<PptxState>);

const _: () = {
    const fn assert_handle<T: Send + Sync + Clone>() {}
    assert_handle::<Pptx>();
    assert!(size_of::<Pptx>() == size_of::<usize>());
};

struct PptxState {
    owner: PptxOwner,
    /// Core-properties metadata. Filled at open for the eager owner and on
    /// the first successful [`Pptx::metadata`] call for the source owner; a
    /// failed read leaves it unset so a later call can retry.
    metadata: OnceLock<Option<Metadata>>,
    /// Whether the opened source stays available for an editor
    /// ([`OpenOptions::retain_source`]).
    retain_source: bool,
}

/// The two states of the PPTX owner (record 0003 §2.3).
enum PptxOwner {
    /// Source-backed, read-only. Its package is the retained source an editor
    /// is built from (L-PPTX-1).
    Source(pptx::SourceBackedPresentation),
    /// Eager, editable through `opened_presentation_transaction`.
    Eager(Box<pptx::Package>),
}

impl PptxOwner {
    /// The TOCTOU guard: verifies the caller's execution policy and that the
    /// positional source still has the revision captured at open. An eager
    /// owner holds no live source, so the check always passes.
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
}

/// The editor [`Pptx::into_editor`] builds from the retained owner.
#[cfg_attr(
    not(test),
    expect(dead_code, reason = "edit verbs consume the editor from M2 on")
)]
pub(crate) enum PptxEditor {
    /// The source-backed slide editor over the package indexed at open.
    Source(Box<pptx::SourceBackedPresentationEditor>),
    /// A detached transaction over the eager package.
    Eager(Box<pptx::opened::Transaction>),
}

/// Why [`Pptx::into_editor`] built no editor. The handle is returned so the
/// caller keeps a readable snapshot.
#[cfg_attr(
    not(test),
    expect(dead_code, reason = "edit verbs consume the editor from M2 on")
)]
pub(crate) struct EditorRefused {
    /// The unchanged handle.
    pub(crate) handle: Pptx,
    /// The typed reason, or `None` when the opened source is still shared by
    /// another handle and so cannot be moved into an editor (record 0003 R4).
    pub(crate) error: Option<Error>,
}

#[cfg_attr(
    not(test),
    expect(
        dead_code,
        reason = "the unified opener (M0a step S6) is the first caller of the constructors"
    )
)]
impl Pptx {
    fn new(owner: PptxOwner, metadata: OnceLock<Option<Metadata>>, options: &OpenOptions) -> Self {
        Self(Arc::new(PptxState {
            owner,
            metadata,
            retain_source: options.retain_source,
        }))
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

    /// Adopts a package prepared by `detect_and_prepare`: no ZIP central
    /// directory is read and no OPC catalog is parsed a second time.
    ///
    /// A prepared package of another OOXML family is refused with
    /// [`Code::UnsupportedFormat`](crate::Code::UnsupportedFormat) naming
    /// the detected format. A source changed between preparation and this
    /// call is a typed source-change error.
    pub(crate) fn from_prepared(
        prepared: opc::PreparedOoxmlSource,
        options: &OpenOptions,
    ) -> Result<Self> {
        let detected =
            crate::detect::ooxml::format_for_main_content_type(prepared.main_part_content_type())
                .ok_or_else(|| Error::unsupported_format(None))?;
        if detected != FileFormat::Pptx {
            return Err(Error::unsupported_format(Some(detected)));
        }
        let package = prepared
            .into_package()
            .map_err(|error| Error::from(error).with_format(FileFormat::Pptx))?;
        Self::from_source_package(package, options)
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
            PptxOwner::Eager(Box::new(package)),
            OnceLock::from(metadata),
            options,
        ))
    }

    /// Maps the path route's detection failure to the unified error, keeping
    /// the typed leaf error.
    #[cfg(any(unix, windows))]
    pub(crate) fn path_detection_error(
        error: crate::detect::detected::PptxSourcePathError,
    ) -> Error {
        use crate::detect::detected::PptxSourcePathError;
        match error {
            PptxSourcePathError::Opc(error) => error.into(),
            PptxSourcePathError::Pptx(error) => error.into(),
            #[cfg(any(feature = "odp", feature = "ppt"))]
            PptxSourcePathError::Source(error) => error.into(),
        }
    }
}

impl Pptx {
    /// Whether the presentation is source-backed (`true`) or was read
    /// eagerly into memory (`false`).
    #[must_use]
    pub fn is_source_backed(&self) -> bool {
        matches!(self.0.owner, PptxOwner::Source(_))
    }

    /// All slide text, slides in presentation order, separated by a blank
    /// line. Slides without text are skipped.
    ///
    /// # Errors
    ///
    /// A source change since open, cancellation, or a slide the reader
    /// rejects.
    pub fn text(&self) -> Result<String> {
        let mut texts = Vec::new();
        match &self.0.owner {
            PptxOwner::Source(presentation) => {
                self.0.owner.check_source()?;
                for slide in presentation.slides() {
                    let text = slide.text()?;
                    if !text.is_empty() {
                        texts.push(text);
                    }
                }
            },
            PptxOwner::Eager(package) => {
                for slide in package.presentation()?.slides()? {
                    let text = slide.text()?;
                    if !text.is_empty() {
                        texts.push(text);
                    }
                }
            },
        }
        Ok(texts.join("\n\n"))
    }

    /// The number of slides.
    ///
    /// # Errors
    ///
    /// A source change since open, cancellation, or (eager owner) a slide
    /// list the reader rejects.
    pub fn slide_count(&self) -> Result<usize> {
        match &self.0.owner {
            PptxOwner::Source(presentation) => {
                self.0.owner.check_source()?;
                Ok(presentation.slide_count())
            },
            PptxOwner::Eager(package) => Ok(package.presentation()?.slide_count()?),
        }
    }

    /// The text and name of every slide, in presentation order.
    ///
    /// # Errors
    ///
    /// As [`Self::text`].
    pub fn slides(&self) -> Result<Vec<PptxSlide>> {
        match &self.0.owner {
            PptxOwner::Source(presentation) => {
                self.0.owner.check_source()?;
                presentation
                    .slides()
                    .map(|slide| {
                        let (text, name) = slide.text_and_name()?;
                        Ok(PptxSlide { text, name })
                    })
                    .collect()
            },
            PptxOwner::Eager(package) => package
                .presentation()?
                .slides()?
                .iter()
                .map(|slide| {
                    Ok(PptxSlide {
                        text: slide.text()?,
                        name: slide.name()?,
                    })
                })
                .collect(),
        }
    }

    /// The text and name of the slide at zero-based `position`, or `None`
    /// past the last slide.
    ///
    /// # Errors
    ///
    /// As [`Self::text`], for the selected slide only.
    pub fn slide(&self, position: usize) -> Result<Option<PptxSlide>> {
        match &self.0.owner {
            PptxOwner::Source(presentation) => {
                self.0.owner.check_source()?;
                let Some(slide) = presentation.slide(position) else {
                    return Ok(None);
                };
                let (text, name) = slide.text_and_name()?;
                Ok(Some(PptxSlide { text, name }))
            },
            PptxOwner::Eager(package) => {
                let presentation = package.presentation()?;
                let Some(slide) = presentation.slide(position)? else {
                    return Ok(None);
                };
                Ok(Some(PptxSlide {
                    text: slide.text()?,
                    name: slide.name()?,
                }))
            },
        }
    }

    /// The ordered slide catalog. No slide XML is read.
    ///
    /// # Errors
    ///
    /// A source change or cancellation observed before, during or after the
    /// walk; an allocation failure.
    pub fn slide_catalog(&self) -> Result<Vec<SlideDescriptor>> {
        match &self.0.owner {
            PptxOwner::Source(presentation) => {
                self.0.owner.check_source()?;
                let mut descriptors = reserve_slide_catalog(presentation.slide_count())?;
                for slide in presentation.slides() {
                    self.0.owner.check_source()?;
                    descriptors.push(SlideDescriptor::new(slide.position(), slide.slide_id()));
                }
                self.0.owner.check_source()?;
                Ok(descriptors)
            },
            PptxOwner::Eager(package) => {
                let references = package.presentation()?.slide_references()?;
                let mut descriptors = reserve_slide_catalog(references.len())?;
                for (position, reference) in references.into_iter().enumerate() {
                    descriptors.push(SlideDescriptor::new(position, reference.id()));
                }
                Ok(descriptors)
            },
        }
    }

    /// The slide size `(width, height)` in EMUs, from the presentation root.
    ///
    /// # Errors
    ///
    /// A source change or cancellation, or a malformed `p:sldSz`.
    pub fn slide_size(&self) -> Result<(i64, i64)> {
        match &self.0.owner {
            PptxOwner::Source(presentation) => Ok(presentation.slide_size()?),
            PptxOwner::Eager(package) => Ok(package.presentation()?.slide_size()?),
        }
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
        self.0.owner.check_source()?;
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
        match &self.0.owner {
            PptxOwner::Source(presentation) => {
                self.0.owner.check_source()?;
                let mut report = presentation.tolerance_report().clone();
                presentation.properties_with_tolerance(&mut report)?;
                for slide in presentation.slides() {
                    slide.text_with_tolerance(&mut report)?;
                }
                Ok(report)
            },
            PptxOwner::Eager(package) => {
                let mut report = package.tolerance_report().clone();
                for slide in package.presentation()?.slides()? {
                    slide.text_with_tolerance(&mut report)?;
                }
                Ok(report)
            },
        }
    }

    /// Builds the leaf editor from the retained owner.
    ///
    /// - Eager owner: a detached `opened` transaction over the shared package.
    /// - Source owner: when this handle is the only one, the
    ///   [`SourceBackedPresentationEditor`](pptx::SourceBackedPresentationEditor)
    ///   is built from the package indexed at open, with no re-indexing
    ///   (L-PPTX-1). When another handle still shares the source, the handle
    ///   comes back with no error (record 0003 R4: M2 decides the fallback).
    ///
    /// An open with `retain_source(false)` has no editable owner and is
    /// refused with [`Code::UnsupportedCapability`](crate::Code::UnsupportedCapability).
    #[cfg_attr(
        not(test),
        expect(dead_code, reason = "edit verbs build the editor from M2 on")
    )]
    pub(crate) fn into_editor(self) -> std::result::Result<PptxEditor, EditorRefused> {
        if !self.0.retain_source {
            return Err(EditorRefused {
                handle: self,
                error: Some(source_not_retained()),
            });
        }
        if let PptxOwner::Eager(package) = &self.0.owner {
            return match package.opened_presentation_transaction() {
                Ok(transaction) => Ok(PptxEditor::Eager(Box::new(transaction))),
                Err(error) => Err(EditorRefused {
                    handle: self,
                    error: Some(error.into()),
                }),
            };
        }
        let state = match Arc::try_unwrap(self.0) {
            Ok(state) => state,
            Err(shared) => {
                return Err(EditorRefused {
                    handle: Self(shared),
                    error: None,
                });
            },
        };
        let PptxState {
            owner,
            metadata,
            retain_source,
        } = state;
        let PptxOwner::Source(presentation) = owner else {
            unreachable!("the eager owner returned above");
        };
        presentation
            .try_into_editor()
            .map(|editor| PptxEditor::Source(Box::new(editor)))
            .map_err(|refused| {
                let (presentation, error) = refused.into_parts();
                EditorRefused {
                    handle: Self(Arc::new(PptxState {
                        owner: PptxOwner::Source(presentation),
                        metadata,
                        retain_source,
                    })),
                    error: error.map(Error::from),
                }
            })
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
