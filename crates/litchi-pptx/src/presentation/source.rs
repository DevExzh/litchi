//! Immutable PresentationML reads backed by a caller-provided positional source.
//!
//! This facade keeps ordinary slide payloads deferred. Opening validates the
//! OPC catalog and mandatory presentation root, then resolves only slide
//! metadata. A slide body is loaded when a selected [`SourceSlide`] is read.

use std::io::Write;
#[cfg(any(unix, windows))]
use std::path::Path;
use std::sync::{Arc, Mutex};

#[cfg(test)]
use std::cell::Cell;

#[cfg(any(unix, windows))]
use litchi_core::FileSource;
use litchi_core::{
    ExecutionContext, ExecutionError, ReadAt, SequentialTextWriter, SourceVersion, TextObjectKind,
    TextOutputError, TextOutputOptions, TextOutputReport,
};
use litchi_opc::constants::{content_type as ct, relationship_type as rt};
use litchi_opc::{
    PackURI, Part, PartData, PartView, ReadLimits, Relationships, SourceBackedPackage,
    SourceCacheLimits, SourceLineage, TargetMode,
};
use quick_xml::events::Event;
use quick_xml::name::{Namespace, ResolveResult};
use quick_xml::reader::NsReader;

use crate::parts::{PresentationPart, SlidePart, SlideReference};
use crate::shape::{Bounds, Shape};
use crate::{Error, Result};

#[cfg(test)]
thread_local! {
    static SOURCE_CATALOG_BUILDS: Cell<usize> = const { Cell::new(0) };
}

#[cfg(test)]
fn reset_source_catalog_builds() {
    SOURCE_CATALOG_BUILDS.set(0);
}

#[cfg(test)]
fn source_catalog_builds() -> usize {
    SOURCE_CATALOG_BUILDS.get()
}

/// Maximum number of existing slides in one source-backed batch edit.
pub const MAX_SOURCE_BACKED_SLIDE_BATCH: usize = 32;

pub(super) struct SourceSlideData {
    position: usize,
    pub(super) slide_id: u32,
    pub(super) part_uri: PackURI,
    // Every source-backed edit against this slide uses the same validated
    // relationship closure. Keep it behind an Arc so a broad batch or an
    // inverse patch does not clone all relationship IDs and targets again.
    pub(super) binding: Arc<SlideBinding>,
}

pub(super) struct SourceInner {
    pub(super) package: SourceBackedPackage,
    // Retain the pinned mandatory root without relying on the OPC payload cache.
    _presentation: SourcePart,
    pub(super) slides: Box<[Arc<SourceSlideData>]>,
}

struct SourceCatalog {
    presentation: SourcePart,
    package_relationships: Box<[RelationshipBinding]>,
    presentation_binding: PartBinding,
    slides: Box<[Arc<SourceSlideData>]>,
}

/// Read-only PresentationML part adapter that retains the source-backed
/// payload handle instead of escaping it as an unmanaged `Arc`.
#[derive(Debug, Clone)]
pub(super) struct SourcePart {
    partname: PackURI,
    content_type: String,
    data: PartData,
    replacement: Option<Vec<u8>>,
    rels: Relationships,
}

impl SourcePart {
    pub(super) fn from_view(view: &PartView<'_>, data: PartData) -> Self {
        Self {
            partname: view.partname().clone(),
            content_type: view.content_type().to_string(),
            data,
            replacement: None,
            rels: view.rels().clone(),
        }
    }
}

struct SourceCheckedTextSink<'a, W: ?Sized> {
    output: &'a mut W,
    package: &'a SourceBackedPackage,
    failure: Arc<Mutex<Option<Error>>>,
}

impl<'a, W: ?Sized> SourceCheckedTextSink<'a, W> {
    fn record_failure(&self, error: Error) {
        let mut failure = self
            .failure
            .lock()
            .unwrap_or_else(|poisoned| poisoned.into_inner());
        if failure.is_none() {
            *failure = Some(error);
        }
    }

    fn check(&self) -> std::io::Result<()> {
        let result = self
            .package
            .check_execution()
            .map_err(Error::from)
            .and_then(|_| {
                self.package
                    .source_version()
                    .map(|_| ())
                    .map_err(Error::from)
            });
        match result {
            Ok(()) => Ok(()),
            Err(error) => {
                let message = error.to_string();
                self.record_failure(error);
                Err(std::io::Error::other(message))
            },
        }
    }
}

impl<'a, W: Write + ?Sized> Write for SourceCheckedTextSink<'a, W> {
    fn write(&mut self, bytes: &[u8]) -> std::io::Result<usize> {
        self.check()?;
        let result = self.output.write(bytes);
        let _ = self.check();
        result
    }

    fn flush(&mut self) -> std::io::Result<()> {
        self.check()?;
        let result = self.output.flush();
        let _ = self.check();
        result
    }
}

fn take_source_text_failure(failure: &Arc<Mutex<Option<Error>>>) -> Option<Error> {
    failure
        .lock()
        .unwrap_or_else(|poisoned| poisoned.into_inner())
        .take()
}

impl Part for SourcePart {
    fn blob(&self) -> &[u8] {
        self.replacement
            .as_deref()
            .unwrap_or_else(|| self.data.as_bytes())
    }

    fn blob_arc(&self) -> Arc<Vec<u8>> {
        // Source catalog and image-query paths use `blob()` so this defensive
        // trait fallback never detaches a managed PartData reservation.
        Arc::new(self.blob().to_vec())
    }

    fn content_type(&self) -> &str {
        &self.content_type
    }

    fn partname(&self) -> &PackURI {
        &self.partname
    }

    fn rels(&self) -> &Relationships {
        &self.rels
    }

    fn rels_mut(&mut self) -> &mut Relationships {
        &mut self.rels
    }

    fn set_blob(&mut self, blob: Vec<u8>) {
        self.replacement = Some(blob);
    }
}

#[derive(Clone, PartialEq, Eq)]
struct RelationshipBinding {
    id: String,
    kind: String,
    target: String,
    mode: TargetMode,
}

#[derive(Clone, PartialEq, Eq)]
struct PartBinding {
    uri: PackURI,
    content_type: String,
    relationships: Box<[RelationshipBinding]>,
}

#[derive(Clone, PartialEq, Eq)]
pub(super) struct SlideBinding {
    pub(super) slide_reference_id: String,
    presentation_relationship: RelationshipBinding,
    slide: PartBinding,
}

#[derive(Clone)]
struct SlideClosure {
    package_relationships: Arc<[RelationshipBinding]>,
    presentation: Arc<PartBinding>,
    presentation_xml: PartData,
    slide: Arc<SlideBinding>,
}

/// A selected-slide payload either remains attached to the source-backed OPC
/// cache or is an explicit edit-owned candidate. Keeping the original as a
/// [`PartData`] handle is important for managed packages: converting it to a
/// bare `Arc<Vec<u8>>` would detach the caller's budget reservation.
#[derive(Clone)]
enum SourcePayload {
    Original(PartData),
    Edited(Arc<Vec<u8>>),
}

impl SourcePayload {
    fn as_bytes(&self) -> &[u8] {
        match self {
            Self::Original(data) => data.as_bytes(),
            Self::Edited(data) => data.as_slice(),
        }
    }

    /// Return the edit-owned payload allocation when this value is a changed
    /// candidate. Original source payloads intentionally do not escape their
    /// [`PartData`] handle: managed packages attach the caller's reservation
    /// to that handle and must be published through the empty-overlay exact
    /// source path for semantic no-ops.
    fn edited_bytes(&self) -> Option<Arc<Vec<u8>>> {
        match self {
            Self::Original(_) => None,
            Self::Edited(data) => Some(Arc::clone(data)),
        }
    }

    fn is_edited(&self) -> bool {
        matches!(self, Self::Edited(_))
    }
}

fn check_execution_context(context: Option<&ExecutionContext>) -> Result<()> {
    let Some(context) = context else {
        return Ok(());
    };
    context.check().map_err(|error| {
        Error::Opc(match error {
            ExecutionError::Cancelled => litchi_opc::OpcError::Cancelled,
            error => litchi_opc::OpcError::Execution(error),
        })
    })
}

#[cfg(any(unix, windows))]
fn file_source(path: impl AsRef<Path>) -> Result<Arc<dyn ReadAt>> {
    Ok(Arc::new(FileSource::open(path)?))
}

/// Read-only PPTX catalog and selected-slide access over a positional source.
///
/// Opening validates the OPC catalog, package relationships, presentation
/// part, and presentation-to-slide graph. Slide payloads remain deferred until
/// [`SourceSlide::text`] selects one. The type has no edit APIs; its bounded
/// semantic text sink is available through
/// [`SourceBackedPresentation::write_text_to`].
#[derive(Clone)]
pub struct SourceBackedPresentation {
    pub(super) inner: Arc<SourceInner>,
}

/// A lifetime-free read-only slide handle from [`SourceBackedPresentation`].
///
/// Creating or listing handles does not read slide XML.
#[derive(Clone)]
pub struct SourceSlide {
    owner: Arc<SourceInner>,
    data: Arc<SourceSlideData>,
}

/// The validated target of one direct `p:pic` image relationship.
///
/// Internal images are limited to `/ppt/media/` parts and retain only the
/// target metadata until [`SourceSlide::read_image`] is called. External
/// targets are reported as inert metadata; this crate never dereferences
/// them.
#[derive(Debug, Clone, PartialEq, Eq)]
#[non_exhaustive]
pub enum SourceImageTarget {
    /// An embedded image part in the package media directory.
    Internal {
        /// Absolute OPC part URI.
        part_uri: PackURI,
        /// Content type declared by `[Content_Types].xml`.
        content_type: String,
    },
    /// An external image relationship retained without network access.
    External {
        /// Producer-supplied external target URI.
        target: String,
    },
}

impl SourceImageTarget {
    /// Whether this target is external and therefore cannot be read by this
    /// crate.
    #[must_use]
    pub const fn is_external(&self) -> bool {
        matches!(self, Self::External { .. })
    }

    /// Borrow the embedded part URI, if this is an internal target.
    #[must_use]
    pub fn part_uri(&self) -> Option<&PackURI> {
        match self {
            Self::Internal { part_uri, .. } => Some(part_uri),
            Self::External { .. } => None,
        }
    }

    /// Borrow the embedded image content type, if this is an internal target.
    #[must_use]
    pub fn content_type(&self) -> Option<&str> {
        match self {
            Self::Internal { content_type, .. } => Some(content_type),
            Self::External { .. } => None,
        }
    }

    /// Borrow the external URI, if this is an external target.
    #[must_use]
    pub fn external_target(&self) -> Option<&str> {
        match self {
            Self::Internal { .. } => None,
            Self::External { target } => Some(target),
        }
    }
}

/// Metadata-only descriptor for one direct image in scene order.
///
/// Listing descriptors reads the selected slide XML and its relationship
/// metadata but never reads an embedded media payload. The `position` value
/// is the exact zero-based image position used by [`SourceSlide::read_image`].
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct SourceImageDescriptor {
    position: usize,
    shape_position: usize,
    id: Option<u32>,
    name: Option<String>,
    bounds: Option<Bounds>,
    relationship_id: String,
    target: SourceImageTarget,
}

impl SourceImageDescriptor {
    /// Exact zero-based image position in scene order.
    #[must_use]
    pub const fn position(&self) -> usize {
        self.position
    }

    /// Exact zero-based pre-order position of the owning shape.
    #[must_use]
    pub const fn shape_position(&self) -> usize {
        self.shape_position
    }

    /// Numeric non-visual shape ID, when safely available.
    #[must_use]
    pub const fn id(&self) -> Option<u32> {
        self.id
    }

    /// Producer-visible non-visual shape name, when safely available.
    #[must_use]
    pub fn name(&self) -> Option<&str> {
        self.name.as_deref()
    }

    /// Explicit local transform bounds, when safely available.
    #[must_use]
    pub const fn bounds(&self) -> Option<Bounds> {
        self.bounds
    }

    /// Relationship ID carried by the picture's `a:blip`.
    #[must_use]
    pub fn relationship_id(&self) -> &str {
        &self.relationship_id
    }

    /// Validated embedded or inert external target metadata.
    #[must_use]
    pub const fn target(&self) -> &SourceImageTarget {
        &self.target
    }

    /// Whether this descriptor points at an external, unreadable target.
    #[must_use]
    pub const fn is_external(&self) -> bool {
        self.target.is_external()
    }
}

/// One selected embedded image and its bounded source-backed payload handle.
///
/// The payload is held as [`litchi_opc::PartData`], so managed package memory
/// reservations and finite cache behavior remain active for the lifetime of
/// this value. External image descriptors cannot produce this type.
#[derive(Debug, Clone)]
pub struct SourceImage {
    descriptor: SourceImageDescriptor,
    data: PartData,
}

impl SourceImage {
    /// Metadata descriptor used to select this image.
    #[must_use]
    pub const fn descriptor(&self) -> &SourceImageDescriptor {
        &self.descriptor
    }

    /// Borrow the exact embedded image payload bytes.
    #[must_use]
    pub fn bytes(&self) -> &[u8] {
        self.data.as_bytes()
    }

    /// Borrow the underlying bounded OPC payload handle.
    #[must_use]
    pub const fn data(&self) -> &PartData {
        &self.data
    }
}

/// An owning, source-backed editor for one existing slide.
///
/// Unlike [`SourceBackedPresentation`], this type is intentionally not
/// cloneable: publishing consumes its deferred OPC source to ensure that the
/// exact source checked during editing is the source raw-copied to output.
/// It supports no package topology or relationship changes.
pub struct SourceBackedPresentationEditor {
    pub(super) package: SourceBackedPackage,
    // Retain the catalog validated at open so each selected slide does not
    // reparse the immutable presentation root and rebuild the complete slide
    // graph. Snapshot capture still checks execution, the selected Part, and
    // the exact source version before publishing anything.
    _presentation: SourcePart,
    pub(super) slides: Box<[Arc<SourceSlideData>]>,
    package_relationships: Arc<[RelationshipBinding]>,
    presentation_binding: Arc<PartBinding>,
    pub(super) limits: ReadLimits,
}

/// An immutable exact-source snapshot of one slide XML part.
#[derive(Clone)]
pub struct SourceBackedSlideSnapshot {
    position: usize,
    part_uri: PackURI,
    xml: SourcePayload,
    closure: SlideClosure,
    max_output_bytes: usize,
    source_version: SourceVersion,
    lineage: SourceLineage,
    context: Option<ExecutionContext>,
}

/// An isolated focused edit of one exact-source slide snapshot.
pub struct SourceBackedSlideEdit {
    source: SourceBackedSlideSnapshot,
    working: SourcePayload,
    operation_used: bool,
    context: Option<ExecutionContext>,
}

/// A reversible, exact-source-checked one-slide XML patch.
#[derive(Clone)]
pub struct SourceBackedSlidePatch {
    before: SourceBackedSlideSnapshot,
    after: SourceBackedSlideSnapshot,
}

/// A checked source-backed slide edit ready for publication.
pub struct SourceBackedSlideCommit {
    snapshot: SourceBackedSlideSnapshot,
    patch: SourceBackedSlidePatch,
}

/// A bounded multi-slide shape-text edit borrowing its deferred source.
pub struct SourceBackedSlideBatchEdit<'a> {
    editor: &'a SourceBackedPresentationEditor,
    commits: Vec<SourceBackedSlideCommit>,
    text_bytes: usize,
}

/// An immutable exact-source snapshot of a selected slide set.
#[derive(Clone)]
pub struct SourceBackedSlideBatchSnapshot {
    slides: Box<[SourceBackedSlideSnapshot]>,
}

/// A reversible, exact-source-checked multi-slide XML patch.
#[derive(Clone)]
pub struct SourceBackedSlideBatchPatch {
    before: SourceBackedSlideBatchSnapshot,
    after: SourceBackedSlideBatchSnapshot,
}

/// A checked source-backed multi-slide edit ready for publication.
pub struct SourceBackedSlideBatchCommit {
    snapshot: SourceBackedSlideBatchSnapshot,
    patch: SourceBackedSlideBatchPatch,
}

impl SourceBackedPresentation {
    /// Open a PPTX from a regular filesystem path without slurping its bytes.
    ///
    /// The path is adapted to the same immutable positional source used by
    /// [`Self::from_read_at`]. The open file handle, rather than the pathname,
    /// remains the source identity for the lifetime of this snapshot.
    #[cfg(any(unix, windows))]
    pub fn from_path(path: impl AsRef<Path>) -> Result<Self> {
        Self::from_path_with_limits(path, ReadLimits::default())
    }

    /// Open a filesystem-backed PPTX with explicit OPC resource limits.
    #[cfg(any(unix, windows))]
    pub fn from_path_with_limits(path: impl AsRef<Path>, limits: ReadLimits) -> Result<Self> {
        Self::from_read_at_with_limits(file_source(path)?, limits)
    }

    /// Open a filesystem-backed PPTX with an explicit finite payload cache.
    #[cfg(any(unix, windows))]
    pub fn from_path_with_cache_limits(
        path: impl AsRef<Path>,
        cache_limits: SourceCacheLimits,
    ) -> Result<Self> {
        Self::from_read_at_with_cache_limits(file_source(path)?, cache_limits)
    }

    /// Open a filesystem-backed PPTX with explicit read and cache policies.
    #[cfg(any(unix, windows))]
    pub fn from_path_with_limits_and_cache_limits(
        path: impl AsRef<Path>,
        limits: ReadLimits,
        cache_limits: SourceCacheLimits,
    ) -> Result<Self> {
        Self::from_read_at_with_limits_and_cache_limits(file_source(path)?, limits, cache_limits)
    }

    /// Open a filesystem-backed PPTX with explicit read and execution
    /// policies while retaining the default finite source cache.
    #[cfg(any(unix, windows))]
    pub fn from_path_with_execution_context(
        path: impl AsRef<Path>,
        limits: ReadLimits,
        context: ExecutionContext,
    ) -> Result<Self> {
        Self::from_read_at_with_execution_context(file_source(path)?, limits, context)
    }

    /// Open a filesystem-backed PPTX with explicit read and execution
    /// policies while retaining the default finite source cache.
    #[cfg(any(unix, windows))]
    pub fn from_path_with_limits_and_execution_context(
        path: impl AsRef<Path>,
        limits: ReadLimits,
        context: ExecutionContext,
    ) -> Result<Self> {
        Self::from_read_at_with_limits_and_execution_context(file_source(path)?, limits, context)
    }

    /// Open a filesystem-backed PPTX with explicit read, cache, and
    /// hierarchical execution policies.
    #[cfg(any(unix, windows))]
    pub fn from_path_with_limits_and_cache_limits_and_execution_context(
        path: impl AsRef<Path>,
        limits: ReadLimits,
        cache_limits: SourceCacheLimits,
        context: ExecutionContext,
    ) -> Result<Self> {
        Self::from_read_at_with_limits_and_cache_limits_and_execution_context(
            file_source(path)?,
            limits,
            cache_limits,
            context,
        )
    }

    /// Open a PPTX from a regular filesystem path.
    #[cfg(any(unix, windows))]
    pub fn open(path: impl AsRef<Path>) -> Result<Self> {
        Self::from_path(path)
    }

    /// Open a filesystem-backed PPTX with explicit OPC resource limits.
    #[cfg(any(unix, windows))]
    pub fn open_with_limits(path: impl AsRef<Path>, limits: ReadLimits) -> Result<Self> {
        Self::from_path_with_limits(path, limits)
    }

    /// Open a filesystem-backed PPTX with an explicit finite payload cache.
    #[cfg(any(unix, windows))]
    pub fn open_with_cache_limits(
        path: impl AsRef<Path>,
        cache_limits: SourceCacheLimits,
    ) -> Result<Self> {
        Self::from_path_with_cache_limits(path, cache_limits)
    }

    /// Open a filesystem-backed PPTX with explicit read and cache policies.
    #[cfg(any(unix, windows))]
    pub fn open_with_limits_and_cache_limits(
        path: impl AsRef<Path>,
        limits: ReadLimits,
        cache_limits: SourceCacheLimits,
    ) -> Result<Self> {
        Self::from_path_with_limits_and_cache_limits(path, limits, cache_limits)
    }

    /// Open a filesystem-backed PPTX with explicit read and execution
    /// policies while retaining the default finite source cache.
    #[cfg(any(unix, windows))]
    pub fn open_with_execution_context(
        path: impl AsRef<Path>,
        limits: ReadLimits,
        context: ExecutionContext,
    ) -> Result<Self> {
        Self::from_path_with_execution_context(path, limits, context)
    }

    /// Open a filesystem-backed PPTX with explicit read and execution
    /// policies while retaining the default finite source cache.
    #[cfg(any(unix, windows))]
    pub fn open_with_limits_and_execution_context(
        path: impl AsRef<Path>,
        limits: ReadLimits,
        context: ExecutionContext,
    ) -> Result<Self> {
        Self::from_path_with_limits_and_execution_context(path, limits, context)
    }

    /// Open a filesystem-backed PPTX with explicit read, cache, and
    /// hierarchical execution policies.
    #[cfg(any(unix, windows))]
    pub fn open_with_limits_and_cache_limits_and_execution_context(
        path: impl AsRef<Path>,
        limits: ReadLimits,
        cache_limits: SourceCacheLimits,
        context: ExecutionContext,
    ) -> Result<Self> {
        Self::from_path_with_limits_and_cache_limits_and_execution_context(
            path,
            limits,
            cache_limits,
            context,
        )
    }

    /// Open an ordinary PPTX package from a caller-provided positional source.
    ///
    /// # Errors
    ///
    /// Returns an error when the source, OPC catalog, presentation root, or
    /// presentation-to-slide graph is malformed.
    pub fn from_read_at(source: Arc<dyn ReadAt>) -> Result<Self> {
        Self::from_read_at_with_limits(source, ReadLimits::default())
    }

    /// Open from a positional source with explicit OPC resource limits.
    ///
    /// # Errors
    ///
    /// Returns an error when the source exceeds `limits`, changes while being
    /// read, or does not contain a coherent PresentationML catalog.
    pub fn from_read_at_with_limits(source: Arc<dyn ReadAt>, limits: ReadLimits) -> Result<Self> {
        Self::from_source_backed_package(SourceBackedPackage::from_read_at_with_limits(
            source, limits,
        )?)
    }

    /// Open from a positional source with an explicit finite payload-cache
    /// policy. This compatibility path remains unmanaged; use one of the
    /// `*_with_execution_context` constructors when cache payloads must be
    /// charged to a caller-owned hierarchical budget.
    pub fn from_read_at_with_cache_limits(
        source: Arc<dyn ReadAt>,
        cache_limits: SourceCacheLimits,
    ) -> Result<Self> {
        Self::from_source_backed_package(SourceBackedPackage::from_read_at_with_cache_limits(
            source,
            cache_limits,
        )?)
    }

    /// Open from a positional source with explicit read and cache policies.
    pub fn from_read_at_with_limits_and_cache_limits(
        source: Arc<dyn ReadAt>,
        limits: ReadLimits,
        cache_limits: SourceCacheLimits,
    ) -> Result<Self> {
        Self::from_source_backed_package(
            SourceBackedPackage::from_read_at_with_limits_and_cache_limits(
                source,
                limits,
                cache_limits,
            )?,
        )
    }

    /// Open from a positional source with an explicit caller execution
    /// context. Lazy slide payloads and cache retention remain attached to
    /// that context; no executor or global scheduler is created by this
    /// facade.
    pub fn from_read_at_with_execution_context(
        source: Arc<dyn ReadAt>,
        limits: ReadLimits,
        context: ExecutionContext,
    ) -> Result<Self> {
        Self::from_read_at_with_limits_and_execution_context(source, limits, context)
    }

    /// Open from a positional source with explicit read and execution
    /// policies while retaining the default finite source cache.
    pub fn from_read_at_with_limits_and_execution_context(
        source: Arc<dyn ReadAt>,
        limits: ReadLimits,
        context: ExecutionContext,
    ) -> Result<Self> {
        Self::from_source_backed_package(SourceBackedPackage::from_read_at_with_execution_context(
            source, limits, context,
        )?)
    }

    /// Open from a positional source with explicit read, cache, and
    /// hierarchical execution policies.
    pub fn from_read_at_with_limits_and_cache_limits_and_execution_context(
        source: Arc<dyn ReadAt>,
        limits: ReadLimits,
        cache_limits: SourceCacheLimits,
        context: ExecutionContext,
    ) -> Result<Self> {
        Self::from_source_backed_package(
            SourceBackedPackage::from_read_at_with_limits_and_cache_limits_and_execution_context(
                source,
                limits,
                cache_limits,
                context,
            )?,
        )
    }

    /// Build the read-only PPTX facade from a validated deferred OPC package.
    ///
    /// Only the mandatory presentation payload is read by this constructor.
    ///
    /// # Errors
    ///
    /// Returns an error when the main part or its ordered slide graph is not a
    /// coherent PresentationML presentation.
    pub fn from_source_backed_package(package: SourceBackedPackage) -> Result<Self> {
        package.check_execution()?;
        let catalog = source_catalog(&package)?;
        package.check_execution()?;

        Ok(Self {
            inner: Arc::new(SourceInner {
                package,
                _presentation: catalog.presentation,
                slides: catalog.slides,
            }),
        })
    }

    /// Number of logical slides in presentation order.
    #[must_use]
    pub fn slide_count(&self) -> usize {
        self.inner.slides.len()
    }

    /// Check the caller execution policy and exact source revision without
    /// reading or materializing any package-part payload.
    ///
    /// This is useful for cached root projections: callers can validate that
    /// the source and cancellation state are still current before returning a
    /// previously retained semantic value.
    ///
    /// # Errors
    ///
    /// Returns a typed OPC cancellation or source-change error when the
    /// caller policy is cancelled or the positional source has changed.
    pub fn check_source(&self) -> Result<()> {
        self.inner.package.check_execution()?;
        self.inner.package.source_version()?;
        Ok(())
    }

    /// Read the relationship-selected OOXML core properties without reading
    /// any slide or media payload.
    ///
    /// A missing core-properties relationship returns `None`; a present but
    /// empty properties part returns `Some(Props::new())`. The source version
    /// and caller execution policy are checked before and after the selected
    /// metadata payload read.
    pub fn properties(&self) -> Result<Option<litchi_ooxml_common::Props>> {
        self.inner.package.check_execution()?;
        let properties = litchi_ooxml_common::properties::read_source_backed(&self.inner.package)?;
        self.inner.package.check_execution()?;
        self.inner.package.source_version()?;
        Ok(properties)
    }

    /// Presentation slide size in EMUs.
    ///
    /// The mandatory presentation root was loaded and retained while the
    /// source catalog was opened; this method parses only that pinned root.
    pub fn slide_size(&self) -> Result<(i64, i64)> {
        self.inner.package.check_execution()?;
        let size = PresentationPart::from_part(&self.inner._presentation)?.slide_size()?;
        self.inner.package.check_execution()?;
        self.inner.package.source_version()?;
        Ok(size)
    }

    /// Iterate lightweight slide handles without reading slide payloads.
    #[must_use]
    pub fn slides(&self) -> impl ExactSizeIterator<Item = SourceSlide> + DoubleEndedIterator + '_ {
        self.inner.slides.iter().cloned().map(|data| SourceSlide {
            owner: Arc::clone(&self.inner),
            data,
        })
    }

    /// Stream one semantic text object per slide into a caller-owned sink.
    ///
    /// The retained lazy catalog supplies slide order and relationship
    /// metadata without collecting slide handles or payloads. One selected
    /// slide is parsed and emitted at a time under the source cache policy.
    /// The source and execution state is checked around every underlying sink
    /// write. For parity with [`SourceSlide::text`], use `"\n"` for the
    /// paragraph and slide separators, and exclude empty objects.
    ///
    /// # Errors
    ///
    /// Returns a typed document, resource-limit, or sink error with exact
    /// partial-output progress. A source revision or execution cancellation
    /// observed after accepted output takes precedence over another failure.
    pub fn write_text_to<W: Write + ?Sized>(
        &self,
        output: &mut W,
        options: TextOutputOptions<'_>,
    ) -> std::result::Result<TextOutputReport, TextOutputError<Error>> {
        let paragraph_separator = options.paragraph_separator();
        let source_failure = Arc::new(Mutex::new(None));
        let mut checked_output = SourceCheckedTextSink {
            output,
            package: &self.inner.package,
            failure: Arc::clone(&source_failure),
        };
        let mut writer = SequentialTextWriter::new(&mut checked_output, options);
        if let Err(source) = self.check_source() {
            return Err(writer.document_error(source));
        }

        for slide in self.slides() {
            if let Err(source) = self.check_source() {
                return Err(writer.document_error(source));
            }

            let parsed = slide.semantic_text(paragraph_separator);
            if let Err(source) = self.check_source() {
                return Err(writer.document_error(source));
            }
            let text = match parsed {
                Ok(text) => text,
                Err(source) => return Err(writer.document_error(source)),
            };

            let emitted = writer.write_object::<Error>(TextObjectKind::Slide, &text);
            let progress = writer.progress();
            let source = self
                .check_source()
                .err()
                .or_else(|| take_source_text_failure(&source_failure));
            if let Some(source) = source {
                return Err(TextOutputError::Document { source, progress });
            }
            emitted?;
        }

        let progress = writer.progress();
        if let Some(source) = self
            .check_source()
            .err()
            .or_else(|| take_source_text_failure(&source_failure))
        {
            Err(TextOutputError::Document { source, progress })
        } else {
            Ok(writer.finish())
        }
    }

    /// Select one slide by checked zero-based presentation position.
    #[must_use]
    pub fn slide(&self, position: usize) -> Option<SourceSlide> {
        let data = self.inner.slides.get(position)?.clone();
        Some(SourceSlide {
            owner: Arc::clone(&self.inner),
            data,
        })
    }

    /// Return content-free source-cache diagnostics without loading a part.
    #[must_use]
    pub fn cache_diagnostics(&self) -> litchi_opc::SourceCacheDiagnostics {
        self.inner.package.cache_diagnostics()
    }
}

impl SourceBackedPresentationEditor {
    /// Open a PPTX from a regular filesystem path without slurping its bytes.
    ///
    /// The path is adapted to the same immutable positional source used by
    /// [`Self::from_read_at`]. The open file handle remains the source
    /// identity for the lifetime of this owning editor.
    #[cfg(any(unix, windows))]
    pub fn from_path(path: impl AsRef<Path>) -> Result<Self> {
        Self::from_path_with_limits(path, ReadLimits::default())
    }

    /// Open a filesystem-backed PPTX with explicit OPC resource limits.
    #[cfg(any(unix, windows))]
    pub fn from_path_with_limits(path: impl AsRef<Path>, limits: ReadLimits) -> Result<Self> {
        Self::from_read_at_with_limits(file_source(path)?, limits)
    }

    /// Open a filesystem-backed PPTX with an explicit finite payload cache.
    #[cfg(any(unix, windows))]
    pub fn from_path_with_cache_limits(
        path: impl AsRef<Path>,
        cache_limits: SourceCacheLimits,
    ) -> Result<Self> {
        Self::from_read_at_with_cache_limits(file_source(path)?, cache_limits)
    }

    /// Open a filesystem-backed PPTX with explicit read and cache policies.
    #[cfg(any(unix, windows))]
    pub fn from_path_with_limits_and_cache_limits(
        path: impl AsRef<Path>,
        limits: ReadLimits,
        cache_limits: SourceCacheLimits,
    ) -> Result<Self> {
        Self::from_read_at_with_limits_and_cache_limits(file_source(path)?, limits, cache_limits)
    }

    /// Open a filesystem-backed PPTX with explicit read and execution
    /// policies while retaining the default finite source cache.
    #[cfg(any(unix, windows))]
    pub fn from_path_with_execution_context(
        path: impl AsRef<Path>,
        limits: ReadLimits,
        context: ExecutionContext,
    ) -> Result<Self> {
        Self::from_read_at_with_execution_context(file_source(path)?, limits, context)
    }

    /// Open a filesystem-backed PPTX with explicit read and execution
    /// policies while retaining the default finite source cache.
    #[cfg(any(unix, windows))]
    pub fn from_path_with_limits_and_execution_context(
        path: impl AsRef<Path>,
        limits: ReadLimits,
        context: ExecutionContext,
    ) -> Result<Self> {
        Self::from_read_at_with_limits_and_execution_context(file_source(path)?, limits, context)
    }

    /// Open a filesystem-backed PPTX with explicit read, cache, and
    /// hierarchical execution policies.
    #[cfg(any(unix, windows))]
    pub fn from_path_with_limits_and_cache_limits_and_execution_context(
        path: impl AsRef<Path>,
        limits: ReadLimits,
        cache_limits: SourceCacheLimits,
        context: ExecutionContext,
    ) -> Result<Self> {
        Self::from_read_at_with_limits_and_cache_limits_and_execution_context(
            file_source(path)?,
            limits,
            cache_limits,
            context,
        )
    }

    /// Open a PPTX from a regular filesystem path.
    #[cfg(any(unix, windows))]
    pub fn open(path: impl AsRef<Path>) -> Result<Self> {
        Self::from_path(path)
    }

    /// Open a filesystem-backed PPTX with explicit OPC resource limits.
    #[cfg(any(unix, windows))]
    pub fn open_with_limits(path: impl AsRef<Path>, limits: ReadLimits) -> Result<Self> {
        Self::from_path_with_limits(path, limits)
    }

    /// Open a filesystem-backed PPTX with an explicit finite payload cache.
    #[cfg(any(unix, windows))]
    pub fn open_with_cache_limits(
        path: impl AsRef<Path>,
        cache_limits: SourceCacheLimits,
    ) -> Result<Self> {
        Self::from_path_with_cache_limits(path, cache_limits)
    }

    /// Open a filesystem-backed PPTX with explicit read and cache policies.
    #[cfg(any(unix, windows))]
    pub fn open_with_limits_and_cache_limits(
        path: impl AsRef<Path>,
        limits: ReadLimits,
        cache_limits: SourceCacheLimits,
    ) -> Result<Self> {
        Self::from_path_with_limits_and_cache_limits(path, limits, cache_limits)
    }

    /// Open a filesystem-backed PPTX with explicit read and execution
    /// policies while retaining the default finite source cache.
    #[cfg(any(unix, windows))]
    pub fn open_with_execution_context(
        path: impl AsRef<Path>,
        limits: ReadLimits,
        context: ExecutionContext,
    ) -> Result<Self> {
        Self::from_path_with_execution_context(path, limits, context)
    }

    /// Open a filesystem-backed PPTX with explicit read and execution
    /// policies while retaining the default finite source cache.
    #[cfg(any(unix, windows))]
    pub fn open_with_limits_and_execution_context(
        path: impl AsRef<Path>,
        limits: ReadLimits,
        context: ExecutionContext,
    ) -> Result<Self> {
        Self::from_path_with_limits_and_execution_context(path, limits, context)
    }

    /// Open a filesystem-backed PPTX with explicit read, cache, and
    /// hierarchical execution policies.
    #[cfg(any(unix, windows))]
    pub fn open_with_limits_and_cache_limits_and_execution_context(
        path: impl AsRef<Path>,
        limits: ReadLimits,
        cache_limits: SourceCacheLimits,
        context: ExecutionContext,
    ) -> Result<Self> {
        Self::from_path_with_limits_and_cache_limits_and_execution_context(
            path,
            limits,
            cache_limits,
            context,
        )
    }

    /// Open an ordinary PPTX source for a bounded one-slide edit.
    ///
    /// Opening validates the OPC catalog, presentation root, and slide graph
    /// without materializing ordinary slide payloads.
    pub fn from_read_at(source: Arc<dyn ReadAt>) -> Result<Self> {
        Self::from_read_at_with_limits(source, ReadLimits::default())
    }

    /// Open an ordinary PPTX source with explicit OPC resource limits.
    pub fn from_read_at_with_limits(source: Arc<dyn ReadAt>, limits: ReadLimits) -> Result<Self> {
        Self::from_source_backed_package(
            SourceBackedPackage::from_read_at_with_limits(source, limits)?,
            limits,
        )
    }

    /// Open a source-backed editor with an explicit finite payload-cache
    /// policy. This compatibility path remains unmanaged.
    pub fn from_read_at_with_cache_limits(
        source: Arc<dyn ReadAt>,
        cache_limits: SourceCacheLimits,
    ) -> Result<Self> {
        Self::from_source_backed_package(
            SourceBackedPackage::from_read_at_with_cache_limits(source, cache_limits)?,
            ReadLimits::default(),
        )
    }

    /// Open a source-backed editor with explicit read and cache policies.
    pub fn from_read_at_with_limits_and_cache_limits(
        source: Arc<dyn ReadAt>,
        limits: ReadLimits,
        cache_limits: SourceCacheLimits,
    ) -> Result<Self> {
        Self::from_source_backed_package(
            SourceBackedPackage::from_read_at_with_limits_and_cache_limits(
                source,
                limits,
                cache_limits,
            )?,
            limits,
        )
    }

    /// Open a source-backed editor with an explicit caller execution context.
    /// Lazy payloads remain budget-managed for the lifetime of snapshots and
    /// edits, and no hidden executor is installed.
    pub fn from_read_at_with_execution_context(
        source: Arc<dyn ReadAt>,
        limits: ReadLimits,
        context: ExecutionContext,
    ) -> Result<Self> {
        Self::from_read_at_with_limits_and_execution_context(source, limits, context)
    }

    /// Open a source-backed editor with explicit read and execution policies.
    pub fn from_read_at_with_limits_and_execution_context(
        source: Arc<dyn ReadAt>,
        limits: ReadLimits,
        context: ExecutionContext,
    ) -> Result<Self> {
        Self::from_source_backed_package(
            SourceBackedPackage::from_read_at_with_execution_context(source, limits, context)?,
            limits,
        )
    }

    /// Open a source-backed editor with explicit read, cache, and hierarchical
    /// execution policies.
    pub fn from_read_at_with_limits_and_cache_limits_and_execution_context(
        source: Arc<dyn ReadAt>,
        limits: ReadLimits,
        cache_limits: SourceCacheLimits,
        context: ExecutionContext,
    ) -> Result<Self> {
        Self::from_source_backed_package(
            SourceBackedPackage::from_read_at_with_limits_and_cache_limits_and_execution_context(
                source,
                limits,
                cache_limits,
                context,
            )?,
            limits,
        )
    }

    /// Build an owning editor from a validated deferred OPC package.
    fn from_source_backed_package(
        package: SourceBackedPackage,
        limits: ReadLimits,
    ) -> Result<Self> {
        package.check_execution()?;
        let catalog = source_catalog(&package)?;
        package.check_execution()?;
        let package_relationships = Arc::from(catalog.package_relationships);
        let presentation_binding = Arc::new(catalog.presentation_binding);
        Ok(Self {
            package,
            _presentation: catalog.presentation,
            slides: catalog.slides,
            package_relationships,
            presentation_binding,
            limits,
        })
    }

    /// Number of logical slides in presentation order.
    #[must_use]
    pub fn slide_count(&self) -> usize {
        self.slides.len()
    }

    /// Return content-free payload-cache activity for this deferred editor.
    ///
    /// Opening loads the presentation root; capturing one slide loads that
    /// selected slide only. The diagnostic exposes no member identities.
    #[must_use]
    pub fn cache_diagnostics(&self) -> litchi_opc::SourceCacheDiagnostics {
        self.package.cache_diagnostics()
    }

    /// Capture the exact raw XML of one existing slide.
    ///
    /// Source-backed slide edits refuse markup-compatibility branch selection:
    /// semantic shape offsets must describe the same bytes that will later be
    /// published.
    pub fn slide_snapshot(&self, position: usize) -> Result<SourceBackedSlideSnapshot> {
        self.package.check_execution()?;
        self.slide_snapshot_for(position, "slide_snapshot")
    }

    /// Begin an isolated focused edit of one existing slide.
    pub fn edit_slide(&self, position: usize) -> Result<SourceBackedSlideEdit> {
        self.package.check_execution()?;
        self.slide_snapshot_for(position, "edit_slide")?
            .edit_checked()
    }

    /// Begin an atomic, bounded shape-text edit across existing slides.
    ///
    /// Each selected slide may receive one nonempty same-slide shape batch.
    /// The edit supports at most [`MAX_SOURCE_BACKED_SLIDE_BATCH`] distinct
    /// slides and the ordinary opened-presentation aggregate text budget.
    #[must_use]
    pub fn edit_slides(&self) -> SourceBackedSlideBatchEdit<'_> {
        SourceBackedSlideBatchEdit {
            editor: self,
            commits: Vec::new(),
            text_bytes: 0,
        }
    }

    /// Publish one exact-source-checked slide commit to a sequential stream.
    ///
    /// The publisher replaces only the selected existing slide XML part and
    /// raw-copies every other source ZIP member. A no-op reproduces the whole
    /// source artifact byte for byte; changed signed sources are refused by
    /// the underlying OPC publisher.
    pub fn publish_slide_commit_to_stream<W: Write>(
        self,
        writer: W,
        commit: &SourceBackedSlideCommit,
    ) -> Result<SourceBackedSlideSnapshot> {
        self.package.check_execution()?;
        let current = self.slide_snapshot_for(
            commit.patch.source().position,
            "publish_slide_commit_to_stream",
        )?;
        let target = commit.patch.apply(&current)?;
        self.package.check_execution()?;
        if let Some(replacement) = target.xml.edited_bytes() {
            self.package.write_part_overlay_shared_to_stream(
                writer,
                &current.part_uri,
                replacement,
            )?;
        } else {
            // An exact semantic no-op must retain the managed source PartData
            // reservation and use the byte-for-byte source publication path.
            self.package
                .write_part_overlays_shared_to_stream(writer, Vec::new())?;
        }
        Ok(target)
    }

    /// Publish one exact-source-checked multi-slide commit to a stream.
    ///
    /// Only changed selected slide XML members are regenerated. Selected
    /// no-ops and every unselected member retain their raw ZIP records. An
    /// all-no-op commit reproduces the complete source artifact byte for byte.
    pub fn publish_slide_batch_commit_to_stream<W: Write>(
        self,
        writer: W,
        commit: &SourceBackedSlideBatchCommit,
    ) -> Result<SourceBackedSlideBatchSnapshot> {
        self.package.check_execution()?;
        let mut slides = Vec::new();
        slides
            .try_reserve_exact(commit.patch.source().slide_count())
            .map_err(|source| Error::Allocation {
                resource: "source-backed slide batch publication snapshot",
                source,
            })?;
        for source in commit.patch.source().slides() {
            slides.push(self.slide_snapshot_from_retained_catalog(
                source.position,
                "publish_slide_batch_commit_to_stream",
            )?);
        }
        let current = SourceBackedSlideBatchSnapshot {
            slides: slides.into_boxed_slice(),
        };
        let target = commit.patch.apply(&current)?;
        self.package.check_execution()?;
        let mut replacements = Vec::new();
        let changed_slides = target
            .slides()
            .filter(|slide| slide.xml.is_edited())
            .count();
        if changed_slides != 0 {
            replacements
                .try_reserve_exact(changed_slides)
                .map_err(|source| Error::Allocation {
                    resource: "source-backed slide batch publication payloads",
                    source,
                })?;
        }
        for slide in target.slides() {
            if let Some(replacement) = slide.xml.edited_bytes() {
                replacements.push((slide.part_uri.clone(), replacement));
            }
        }
        self.package
            .write_part_overlays_shared_to_stream(writer, replacements)?;
        Ok(target)
    }

    fn slide_snapshot_for(
        &self,
        position: usize,
        operation: &'static str,
    ) -> Result<SourceBackedSlideSnapshot> {
        self.package.check_execution()?;
        let snapshot = self.slide_snapshot_from_retained_catalog(position, operation)?;
        self.package.check_execution()?;
        Ok(snapshot)
    }

    fn slide_snapshot_from_retained_catalog(
        &self,
        position: usize,
        operation: &'static str,
    ) -> Result<SourceBackedSlideSnapshot> {
        self.package.check_execution()?;
        let data = self
            .slides
            .get(position)
            .ok_or(Error::SlideIndexOutOfBounds {
                index: position,
                len: self.slides.len(),
            })?;
        let view = self.package.part(&data.part_uri)?;
        let raw = view.data()?;
        self.package.check_execution()?;
        let lineage = self.package.source_lineage();
        let context = self.package.execution_context();
        let part = SourcePart::from_view(&view, raw.clone());
        SlidePart::from_part(&part)?;
        let scene = crate::shape::Scene::read(raw.as_bytes())?;
        self.package.check_execution()?;
        let source_version = self.package.source_version()?;
        if scene.is_rewritten() {
            return Err(Error::UnsafeEdit {
                operation,
                reason: "source-backed slide edits do not support markup-compatibility branch selection",
            });
        }
        Ok(SourceBackedSlideSnapshot {
            position,
            part_uri: data.part_uri.clone(),
            xml: SourcePayload::Original(raw),
            closure: SlideClosure {
                package_relationships: self.package_relationships.clone(),
                presentation: self.presentation_binding.clone(),
                presentation_xml: self._presentation.data.clone(),
                slide: data.binding.clone(),
            },
            max_output_bytes: source_slide_output_limit(self.limits),
            source_version,
            lineage,
            context,
        })
    }
}

impl SourceBackedSlideSnapshot {
    /// Checked zero-based slide position in the presentation catalog.
    #[must_use]
    pub const fn position(&self) -> usize {
        self.position
    }

    /// Read this slide's direct standard transition, if present.
    ///
    /// Inherited layout/master transitions are intentionally outside this
    /// direct-slide capability. Markup-compatibility and extension transition
    /// forms return a typed refusal instead of being projected into a mutable
    /// approximation.
    pub fn transition(&self) -> Result<Option<crate::transition::Transition>> {
        self.check_execution()?;
        let transition = crate::presentation::transition::read_direct(self.xml.as_bytes())?;
        self.check_execution()?;
        Ok(transition)
    }

    /// Start an isolated focused edit from this exact source snapshot.
    #[must_use]
    pub fn edit(&self) -> SourceBackedSlideEdit {
        SourceBackedSlideEdit {
            source: self.clone(),
            working: self.xml.clone(),
            operation_used: false,
            context: self.context.clone(),
        }
    }

    /// Check the retained caller execution policy before creating an edit
    /// handle. The infallible [`Self::edit`] constructor remains available for
    /// compatibility because it performs no payload I/O.
    pub fn edit_checked(&self) -> Result<SourceBackedSlideEdit> {
        self.check_execution()?;
        Ok(self.edit())
    }

    fn check_execution(&self) -> Result<()> {
        check_execution_context(self.context.as_ref())
    }

    fn same_source(&self, other: &Self) -> bool {
        self.position == other.position
            && self.part_uri == other.part_uri
            && self.source_version == other.source_version
            && self.lineage == other.lineage
            && self.xml.as_bytes() == other.xml.as_bytes()
            && self.closure.package_relationships == other.closure.package_relationships
            && self.closure.presentation == other.closure.presentation
            && self.closure.presentation_xml.as_bytes() == other.closure.presentation_xml.as_bytes()
            && self.closure.slide == other.closure.slide
            && self.max_output_bytes == other.max_output_bytes
    }
}

impl SourceBackedSlideEdit {
    /// Exact immutable slide snapshot against which this edit was created.
    #[must_use]
    pub const fn source(&self) -> &SourceBackedSlideSnapshot {
        &self.source
    }

    /// Set or replace this existing slide's direct standard transition.
    ///
    /// This operation never changes relationships or package topology. The
    /// supplied value must use only the standard direct transition vocabulary;
    /// custom-duration, PowerPoint-extension, raw, and preserved-extension
    /// values are refused. An equal modeled value is an exact byte no-op.
    pub fn set_transition(&mut self, value: &crate::transition::Transition) -> Result<bool> {
        self.stage_transition(Some(value), "set_transition")
    }

    /// Clear this existing slide's direct standard transition.
    ///
    /// An absent direct transition is an exact byte no-op. Inherited
    /// layout/master state is not removed or rewritten.
    pub fn clear_transition(&mut self) -> Result<bool> {
        self.stage_transition(None, "clear_transition")
    }

    fn stage_transition(
        &mut self,
        value: Option<&crate::transition::Transition>,
        operation: &'static str,
    ) -> Result<bool> {
        self.check_execution()?;
        if self.operation_used {
            return Err(Error::UnsafeEdit {
                operation,
                reason: "source-backed slide edits support one atomic semantic operation",
            });
        }
        let (xml, changed) = crate::presentation::transition::stage(
            self.working.as_bytes(),
            self.source.closure.presentation_xml.as_bytes(),
            value,
            self.source.max_output_bytes,
            operation,
        )?;
        if let Some(xml) = xml {
            self.working = SourcePayload::Edited(Arc::new(xml));
        }
        self.operation_used = true;
        self.check_execution()?;
        Ok(changed)
    }

    /// Replace all visible DrawingML text runs in one existing shape.
    ///
    /// Structure, formatting, opaque XML, relationships, and every other ZIP
    /// member remain outside this edit's closure.
    pub fn set_shape_text<'k>(
        &mut self,
        shape: impl Into<crate::shape::Key<'k>>,
        text: impl AsRef<str>,
    ) -> Result<bool> {
        let text = text.as_ref();
        let replacement = crate::opened::ShapeTextReplacement::new(shape.into(), text);
        Ok(self.set_shape_texts(std::slice::from_ref(&replacement))? != 0)
    }

    /// Atomically replace visible text in a bounded set of shapes.
    ///
    /// All selectors resolve against the same immutable slide state. The
    /// batch is canonicalized by raw shape position, refuses duplicate and
    /// overlapping selections, and publishes at most this one existing slide
    /// part. The return value is the number of changed shapes.
    ///
    /// An empty batch is a no-op and does not consume the edit capability. A
    /// successful nonempty batch, including an all-equal batch, consumes it.
    pub fn set_shape_texts(
        &mut self,
        replacements: &[crate::opened::ShapeTextReplacement<'_>],
    ) -> Result<usize> {
        self.check_execution()?;
        if replacements.is_empty() {
            return Ok(0);
        }
        if self.operation_used {
            return Err(Error::UnsafeEdit {
                operation: "set_shape_texts",
                reason: "source-backed slide edits support one atomic shape text operation",
            });
        }
        let (xml, changed) = crate::opened::stage_shape_texts(
            self.working.as_bytes(),
            replacements,
            crate::opened::Limits::default().max_text_bytes(),
            self.source.max_output_bytes,
        )?;
        if let Some(xml) = xml {
            self.working = SourcePayload::Edited(Arc::new(xml));
        }
        self.operation_used = true;
        self.check_execution()?;
        Ok(changed)
    }

    fn check_execution(&self) -> Result<()> {
        check_execution_context(self.context.as_ref())
    }

    fn into_commit(self) -> SourceBackedSlideCommit {
        let snapshot = SourceBackedSlideSnapshot {
            position: self.source.position,
            part_uri: self.source.part_uri.clone(),
            xml: self.working,
            closure: self.source.closure.clone(),
            max_output_bytes: self.source.max_output_bytes,
            source_version: self.source.source_version,
            lineage: self.source.lineage.clone(),
            context: self.source.context.clone(),
        };
        let patch = SourceBackedSlidePatch {
            before: self.source,
            after: snapshot.clone(),
        };
        SourceBackedSlideCommit { snapshot, patch }
    }

    /// Validate and freeze this isolated edit for source-backed publication.
    #[must_use]
    pub fn commit(self) -> SourceBackedSlideCommit {
        self.into_commit()
    }

    /// Check cancellation and freeze this isolated edit for publication.
    pub fn commit_checked(self) -> Result<SourceBackedSlideCommit> {
        self.check_execution()?;
        let commit = self.into_commit();
        commit.snapshot.check_execution()?;
        Ok(commit)
    }
}

impl SourceBackedSlidePatch {
    /// Exact immutable source required by this patch.
    #[must_use]
    pub const fn source(&self) -> &SourceBackedSlideSnapshot {
        &self.before
    }

    /// Exact immutable target produced by this patch.
    #[must_use]
    pub const fn target(&self) -> &SourceBackedSlideSnapshot {
        &self.after
    }

    /// Whether this patch changes the selected slide XML bytes.
    #[must_use]
    pub fn is_changed(&self) -> bool {
        !self.before.same_source(&self.after)
    }

    /// Return the exact inverse slide patch.
    #[must_use]
    pub fn inverse(&self) -> Self {
        Self {
            before: self.after.clone(),
            after: self.before.clone(),
        }
    }

    /// Apply only to the exact raw slide source captured by this patch.
    pub fn apply(&self, source: &SourceBackedSlideSnapshot) -> Result<SourceBackedSlideSnapshot> {
        self.before.check_execution()?;
        source.check_execution()?;
        if !source.same_source(&self.before) {
            return Err(Error::StaleSource);
        }
        Ok(if self.is_changed() {
            self.after.clone()
        } else {
            source.clone()
        })
    }
}

impl SourceBackedSlideCommit {
    /// Candidate snapshot after this edit.
    #[must_use]
    pub const fn snapshot(&self) -> &SourceBackedSlideSnapshot {
        &self.snapshot
    }

    /// Exact-source slide patch for this edit.
    #[must_use]
    pub const fn patch(&self) -> &SourceBackedSlidePatch {
        &self.patch
    }

    /// Whether the selected slide XML changes.
    #[must_use]
    pub fn is_changed(&self) -> bool {
        self.patch.is_changed()
    }
}

impl SourceBackedSlideBatchEdit<'_> {
    /// Atomically replace visible text in one bounded shape set on one slide.
    ///
    /// A slide position may be selected only once in the outer batch. An empty
    /// shape batch is a no-op and does not select the slide. The return value
    /// is the number of shapes whose text bytes changed.
    pub fn set_shape_texts(
        &mut self,
        position: usize,
        replacements: &[crate::opened::ShapeTextReplacement<'_>],
    ) -> Result<usize> {
        self.editor.package.check_execution()?;
        if replacements.is_empty() {
            return Ok(0);
        }
        if position >= self.editor.slide_count() {
            return Err(Error::SlideIndexOutOfBounds {
                index: position,
                len: self.editor.slide_count(),
            });
        }
        if self
            .commits
            .iter()
            .any(|commit| commit.snapshot.position == position)
        {
            return Err(Error::UnsafeEdit {
                operation: "set_shape_texts",
                reason: "a source-backed slide batch may select each slide only once",
            });
        }
        if self.commits.len() >= MAX_SOURCE_BACKED_SLIDE_BATCH {
            return Err(Error::Limit {
                resource: "source-backed slide batch selections",
                limit: MAX_SOURCE_BACKED_SLIDE_BATCH,
            });
        }
        let added_text_bytes = replacements
            .iter()
            .try_fold(0_usize, |total, replacement| {
                total
                    .checked_add(replacement.text().len())
                    .ok_or(Error::Limit {
                        resource: "source-backed slide batch replacement text",
                        limit: crate::opened::Limits::default().max_text_bytes(),
                    })
            })?;
        let text_bytes = self
            .text_bytes
            .checked_add(added_text_bytes)
            .ok_or(Error::Limit {
                resource: "source-backed slide batch replacement text",
                limit: crate::opened::Limits::default().max_text_bytes(),
            })?;
        if text_bytes > crate::opened::Limits::default().max_text_bytes() {
            return Err(Error::Limit {
                resource: "source-backed slide batch replacement text",
                limit: crate::opened::Limits::default().max_text_bytes(),
            });
        }

        let mut edit = self.editor.edit_slide(position)?;
        let changed = edit.set_shape_texts(replacements)?;
        self.commits
            .try_reserve(1)
            .map_err(|source| Error::Allocation {
                resource: "source-backed slide batch edits",
                source,
            })?;
        self.commits.push(edit.commit_checked()?);
        self.editor.package.check_execution()?;
        self.text_bytes = text_bytes;
        Ok(changed)
    }

    /// Validate and freeze this selected slide set for publication.
    pub fn commit(mut self) -> Result<SourceBackedSlideBatchCommit> {
        self.editor.package.check_execution()?;
        if self.commits.is_empty() {
            return Err(Error::UnsafeEdit {
                operation: "commit_slide_batch",
                reason: "a source-backed slide batch requires at least one selected slide",
            });
        }
        self.commits
            .sort_unstable_by_key(|commit| commit.snapshot.position);
        let mut before = Vec::new();
        let mut after = Vec::new();
        before
            .try_reserve_exact(self.commits.len())
            .map_err(|source| Error::Allocation {
                resource: "source-backed slide batch source snapshot",
                source,
            })?;
        after
            .try_reserve_exact(self.commits.len())
            .map_err(|source| Error::Allocation {
                resource: "source-backed slide batch target snapshot",
                source,
            })?;
        for commit in self.commits {
            commit.patch.before.check_execution()?;
            before.push(commit.patch.before);
            after.push(commit.patch.after);
        }
        let before = SourceBackedSlideBatchSnapshot {
            slides: before.into_boxed_slice(),
        };
        let after = SourceBackedSlideBatchSnapshot {
            slides: after.into_boxed_slice(),
        };
        let patch = SourceBackedSlideBatchPatch {
            before,
            after: after.clone(),
        };
        Ok(SourceBackedSlideBatchCommit {
            snapshot: after,
            patch,
        })
    }
}

impl SourceBackedSlideBatchSnapshot {
    /// Number of selected slides in this exact-source snapshot.
    #[must_use]
    pub const fn slide_count(&self) -> usize {
        self.slides.len()
    }

    /// Selected slide snapshots in ascending presentation order.
    pub fn slides(&self) -> impl ExactSizeIterator<Item = &SourceBackedSlideSnapshot> {
        self.slides.iter()
    }

    fn same_source(&self, other: &Self) -> bool {
        self.slides.len() == other.slides.len()
            && self
                .slides
                .iter()
                .zip(other.slides.iter())
                .all(|(left, right)| left.same_source(right))
    }

    fn check_execution(&self) -> Result<()> {
        for slide in &self.slides {
            slide.check_execution()?;
        }
        Ok(())
    }
}

impl SourceBackedSlideBatchPatch {
    /// Exact immutable selected slide set required by this patch.
    #[must_use]
    pub const fn source(&self) -> &SourceBackedSlideBatchSnapshot {
        &self.before
    }

    /// Exact immutable selected slide set produced by this patch.
    #[must_use]
    pub const fn target(&self) -> &SourceBackedSlideBatchSnapshot {
        &self.after
    }

    /// Whether any selected slide XML bytes change.
    #[must_use]
    pub fn is_changed(&self) -> bool {
        !self.before.same_source(&self.after)
    }

    /// Return the exact inverse multi-slide patch.
    #[must_use]
    pub fn inverse(&self) -> Self {
        Self {
            before: self.after.clone(),
            after: self.before.clone(),
        }
    }

    /// Apply only to the exact selected slide set captured by this patch.
    pub fn apply(
        &self,
        source: &SourceBackedSlideBatchSnapshot,
    ) -> Result<SourceBackedSlideBatchSnapshot> {
        self.before.check_execution()?;
        source.check_execution()?;
        if !source.same_source(&self.before) {
            return Err(Error::StaleSource);
        }
        Ok(if self.is_changed() {
            self.after.clone()
        } else {
            source.clone()
        })
    }
}

impl SourceBackedSlideBatchCommit {
    /// Candidate selected slide set after this edit.
    #[must_use]
    pub const fn snapshot(&self) -> &SourceBackedSlideBatchSnapshot {
        &self.snapshot
    }

    /// Exact-source multi-slide patch for this edit.
    #[must_use]
    pub const fn patch(&self) -> &SourceBackedSlideBatchPatch {
        &self.patch
    }

    /// Whether any selected slide XML bytes change.
    #[must_use]
    pub fn is_changed(&self) -> bool {
        self.patch.is_changed()
    }
}

impl SourceSlide {
    /// Checked zero-based position in the presentation catalog.
    #[must_use]
    pub fn position(&self) -> usize {
        self.data.position
    }

    /// Stable `p:sldId@id` value retained from the presentation catalog.
    ///
    /// Reading the identifier never loads the selected slide XML or any
    /// related media payload. The value is captured while the source-backed
    /// presentation graph is opened.
    #[must_use]
    pub fn slide_id(&self) -> u32 {
        self.data.slide_id
    }

    /// Return the producer-visible slide name, falling back to the slide
    /// part URI when the root has no `name` attribute.
    ///
    /// Only this selected slide's XML is loaded. Source-version and caller
    /// execution checks surround the read, including when the payload is
    /// already resident in the bounded source cache.
    pub fn name(&self) -> Result<String> {
        self.owner.package.check_execution()?;
        let part = self.part()?;
        let name = SlidePart::from_part(&part)?.name()?;
        self.owner.package.check_execution()?;
        self.owner.package.source_version()?;
        Ok(name)
    }

    /// Read the selected slide's name and flattened text with one bounded
    /// source-part load.
    ///
    /// This is the efficient primitive for callers that need both values
    /// (such as a unified slide projection). The selected slide payload is
    /// loaded and validated once; unrelated slides and media remain cold.
    pub fn text_and_name(&self) -> Result<(String, String)> {
        self.owner.package.check_execution()?;
        let part = self.part()?;
        let slide = SlidePart::from_part(&part)?;
        let (text, name) = slide.text_and_name()?;
        self.owner.package.check_execution()?;
        self.owner.package.source_version()?;
        Ok((text, name))
    }

    fn semantic_text(&self, paragraph_separator: &str) -> Result<String> {
        self.owner.package.check_execution()?;
        let view = self.owner.package.part(&self.data.part_uri)?;
        let declared = view.declared_uncompressed_size()?;
        let raw_xml_limit = SlidePart::semantic_text_raw_xml_limit();
        if declared > raw_xml_limit as u64 {
            return Err(Error::Limit {
                resource: "semantic slide raw XML bytes",
                limit: raw_xml_limit,
            });
        }
        let part = self.load_part(&view)?;
        let text = SlidePart::semantic_text_from_part(&part, paragraph_separator)?;
        self.owner.package.check_execution()?;
        self.owner.package.source_version()?;
        Ok(text)
    }

    /// Flatten DrawingML text runs from this selected slide in source order.
    ///
    /// The slide payload is loaded only for this call and is released from the
    /// facade's view when the call returns. No other slide payload is read.
    ///
    /// # Errors
    ///
    /// Returns an error if the source changed, the selected slide exceeds the
    /// retained OPC limits, or its PresentationML is malformed.
    pub fn text(&self) -> Result<String> {
        self.owner.package.check_execution()?;
        let part = self.part()?;
        let text = SlidePart::from_part(&part)?.text()?;
        self.owner.package.check_execution()?;
        self.owner.package.source_version()?;
        Ok(text)
    }

    /// List direct `p:pic` image descriptors in scene order.
    ///
    /// This reads only the selected slide XML and its relationship metadata.
    /// It never reads an embedded `/ppt/media/` payload. Backgrounds,
    /// inherited layout/master images, notes, charts, OLE previews, and
    /// non-picture graphic frames are outside this direct-picture closure.
    /// Markup-compatibility branches and ambiguous or malformed picture
    /// grammar are refused rather than projected into a lossy descriptor.
    pub fn images(&self) -> Result<Vec<SourceImageDescriptor>> {
        self.owner.package.check_execution()?;
        let view = self.owner.package.part(&self.data.part_uri)?;
        let part = self.load_part(&view)?;
        validate_source_slide_root(&part)?;
        reject_picture_markup_compatibility(part.blob())?;
        // Keep the normal borrowed part validation on the safe path as well;
        // the raw-root check above prevents MCE preprocessing from masking a
        // malformed selected slide before our refusal scan runs.
        SlidePart::from_part(&part)?;
        let scene = crate::shape::Scene::read(part.blob())?;
        if scene.is_rewritten() {
            return Err(Error::UnsafeEdit {
                operation: "source-backed picture inventory",
                reason: "markup-compatibility picture branches are not selected",
            });
        }

        let mut descriptors = Vec::new();
        descriptors
            .try_reserve_exact(scene.len())
            .map_err(|source| Error::Allocation {
                resource: "source-backed picture descriptors",
                source,
            })?;
        for (shape_position, shape) in scene.iter().enumerate() {
            let Shape::Picture(picture) = shape else {
                continue;
            };
            let common = picture.common();
            let relationship = parse_picture_relationship(common.xml()?)?;
            let target = resolve_picture_target(&self.owner.package, &view, &relationship)?;
            descriptors.push(SourceImageDescriptor {
                position: descriptors.len(),
                shape_position,
                id: common.id(),
                name: common.name().map(str::to_owned),
                bounds: common.bounds(),
                relationship_id: relationship.id,
                target,
            });
        }
        self.owner.package.check_execution()?;
        self.owner.package.source_version()?;
        Ok(descriptors)
    }

    /// Select one direct image descriptor by exact zero-based image position.
    ///
    /// The selected embedded payload is still deferred; use
    /// [`Self::read_image`] when payload bytes are required.
    pub fn image(&self, position: usize) -> Result<SourceImageDescriptor> {
        let images = self.images()?;
        let len = images.len();
        images
            .into_iter()
            .nth(position)
            .ok_or(Error::IndexOutOfBounds {
                index: position,
                len,
            })
    }

    /// Read one embedded image payload by exact zero-based image position.
    ///
    /// Internal image bytes are loaded through the source-backed OPC PartView,
    /// retaining source-version checks, finite cache limits, cancellation,
    /// and managed budget reservations. External targets are described by
    /// [`Self::images`] but are refused here without any network or filesystem
    /// access.
    pub fn read_image(&self, position: usize) -> Result<SourceImage> {
        self.owner.package.check_execution()?;
        let descriptor = self.image(position)?;
        let SourceImageTarget::Internal { part_uri, .. } = &descriptor.target else {
            return Err(Error::Relationship(format!(
                "source-backed image {position} has an external target and is inert"
            )));
        };
        let view = self.owner.package.part(part_uri)?;
        if !view.rels().is_empty() {
            return Err(Error::Relationship(format!(
                "source-backed image part '{}' has outbound relationships",
                part_uri.as_str()
            )));
        }
        let data = view.data()?;
        self.owner.package.check_execution()?;
        self.owner.package.source_version()?;
        Ok(SourceImage { descriptor, data })
    }

    fn part(&self) -> Result<SourcePart> {
        self.owner.package.check_execution()?;
        // The metadata lookup keeps source-version checks active even after a
        // selected slide payload has entered the local cache.
        let view = self.owner.package.part(&self.data.part_uri)?;
        self.load_part(&view)
    }

    fn load_part(&self, view: &PartView<'_>) -> Result<SourcePart> {
        self.owner.package.check_execution()?;
        let part = SourcePart::from_view(view, view.data()?);
        self.owner.package.check_execution()?;
        Ok(part)
    }
}

struct PictureRelationship {
    id: String,
    external: bool,
}

#[derive(Clone, Copy, PartialEq, Eq)]
enum PictureNode {
    Picture,
    NonVisualProperties,
    BlipFill,
    ShapeProperties,
    Blip,
    SourceRectangle,
    Stretch,
    Tile,
    FillRectangle,
    Other,
}

const MCE_NAMESPACE: &[u8] = b"http://schemas.openxmlformats.org/markup-compatibility/2006";
const DRAWINGML_NAMESPACE: &[u8] = b"http://schemas.openxmlformats.org/drawingml/2006/main";
const STRICT_DRAWINGML_NAMESPACE: &[u8] = b"http://purl.oclc.org/ooxml/drawingml/main";

fn parse_picture_relationship(xml: &[u8]) -> Result<PictureRelationship> {
    let mut reader = NsReader::from_reader(xml);
    let mut stack: Vec<(PictureNode, Vec<u8>)> = Vec::new();
    let mut root_seen = false;
    let mut picture_stage = 0_u8;
    let mut blip_fill_stage = 0_u8;
    let mut saw_blip = false;
    let mut saw_stretch = false;
    let mut saw_tile = false;
    let mut saw_fill_rectangle = false;
    let mut relationship = None;
    loop {
        let (namespace, event) = reader
            .read_resolved_event()
            .map_err(|error| Error::Xml(error.to_string()))?;
        let is_start = matches!(&event, Event::Start(_));
        match event {
            Event::Start(element) | Event::Empty(element) => {
                let parent = stack.last().map(|(node, _)| *node);
                let node = picture_node(&namespace, element.name());
                if stack.is_empty() && node != PictureNode::Picture {
                    return Err(Error::Invalid(
                        "picture descriptor XML does not have a p:pic root".into(),
                    ));
                }
                if stack.is_empty() {
                    if root_seen {
                        return Err(Error::Invalid(
                            "picture descriptor contains more than one root element".into(),
                        ));
                    }
                    root_seen = true;
                }
                if node == PictureNode::Picture && !stack.is_empty() {
                    return Err(Error::Invalid(
                        "picture descriptor contains a nested p:pic".into(),
                    ));
                }

                match parent {
                    None => {},
                    Some(PictureNode::Picture) => {
                        let expected = match picture_stage {
                            0 => PictureNode::NonVisualProperties,
                            1 => PictureNode::BlipFill,
                            2 => PictureNode::ShapeProperties,
                            _ => PictureNode::Other,
                        };
                        if node != expected {
                            return Err(Error::Invalid(
                                "picture descriptor direct children must be nvPicPr, blipFill, and spPr in schema order".into(),
                            ));
                        }
                        picture_stage = picture_stage.saturating_add(1);
                    },
                    Some(PictureNode::BlipFill) => match node {
                        PictureNode::Blip if !saw_blip && blip_fill_stage == 0 => {
                            saw_blip = true;
                            blip_fill_stage = 1;
                            validate_blip_attributes(
                                &element,
                                reader.decoder(),
                                reader.resolver(),
                            )?;
                            let embed = crate::namespace::relationship_attribute_value(
                                &element,
                                b"embed",
                                reader.decoder(),
                                reader.resolver(),
                            )?;
                            let link = crate::namespace::relationship_attribute_value(
                                &element,
                                b"link",
                                reader.decoder(),
                                reader.resolver(),
                            )?;
                            let (id, external) = match (embed, link) {
                                (Some(id), None) => (id, false),
                                (None, Some(id)) => (id, true),
                                (Some(_), Some(_)) => {
                                    return Err(Error::Relationship(
                                        "picture a:blip cannot carry both r:embed and r:link"
                                            .into(),
                                    ));
                                },
                                (None, None) => {
                                    return Err(Error::Relationship(
                                        "picture a:blip is missing r:embed or r:link".into(),
                                    ));
                                },
                            };
                            if id.is_empty() {
                                return Err(Error::Relationship(
                                    "picture a:blip relationship ID is empty".into(),
                                ));
                            }
                            relationship = Some(PictureRelationship { id, external });
                        },
                        PictureNode::SourceRectangle if blip_fill_stage <= 1 => {
                            blip_fill_stage = 2;
                        },
                        PictureNode::Stretch
                            if blip_fill_stage <= 2 && !saw_stretch && !saw_tile =>
                        {
                            blip_fill_stage = 3;
                            saw_stretch = true;
                            if !is_start {
                                return Err(Error::Invalid(
                                    "picture blipFill stretch must contain one fillRect".into(),
                                ));
                            }
                        },
                        PictureNode::Tile if blip_fill_stage <= 2 && !saw_stretch && !saw_tile => {
                            blip_fill_stage = 3;
                            saw_tile = true;
                        },
                        _ => {
                            return Err(Error::Invalid(
                                "picture blipFill direct children are an ordered blip, srcRect, and stretch/tile set".into(),
                            ));
                        },
                    },
                    Some(PictureNode::Stretch) => {
                        if node != PictureNode::FillRectangle || saw_fill_rectangle {
                            return Err(Error::Invalid(
                                "picture stretch requires exactly one direct fillRect".into(),
                            ));
                        }
                        saw_fill_rectangle = true;
                    },
                    Some(
                        PictureNode::Blip
                        | PictureNode::SourceRectangle
                        | PictureNode::Tile
                        | PictureNode::FillRectangle,
                    ) => {
                        return Err(Error::Invalid(
                            "picture blipFill child contains an unsupported nested element".into(),
                        ));
                    },
                    Some(PictureNode::NonVisualProperties)
                    | Some(PictureNode::ShapeProperties)
                    | Some(PictureNode::Other) => {},
                }

                if parent == Some(PictureNode::BlipFill)
                    && node == PictureNode::Stretch
                    && !is_start
                {
                    return Err(Error::Invalid(
                        "picture blipFill stretch must contain one fillRect".into(),
                    ));
                }
                if node == PictureNode::Picture && !is_start {
                    return Err(Error::Invalid(
                        "picture descriptor requires nvPicPr, blipFill, and spPr children".into(),
                    ));
                }

                if is_start {
                    stack.push((node, element.name().as_ref().to_vec()));
                }
            },
            Event::End(element) => {
                let (_, start_name) = stack.pop().ok_or_else(|| {
                    Error::Invalid("picture descriptor has an unmatched end element".into())
                })?;
                if start_name.as_slice() != element.name().as_ref() {
                    return Err(Error::Invalid(
                        "picture descriptor has a mismatched end element".into(),
                    ));
                }
            },
            Event::Text(text) => {
                let parent = stack.last().map(|(node, _)| *node);
                if parent.is_none() && !text.as_ref().iter().all(u8::is_ascii_whitespace) {
                    return Err(Error::Invalid(
                        "picture descriptor contains text outside its p:pic root".into(),
                    ));
                }
                let finite_parent = matches!(
                    parent,
                    Some(
                        PictureNode::Picture
                            | PictureNode::BlipFill
                            | PictureNode::Blip
                            | PictureNode::SourceRectangle
                            | PictureNode::Stretch
                            | PictureNode::Tile
                            | PictureNode::FillRectangle
                    )
                );
                if finite_parent
                    && (parent == Some(PictureNode::Blip)
                        || !text.as_ref().iter().all(u8::is_ascii_whitespace))
                {
                    return Err(Error::Invalid(
                        "picture descriptor contains text in a finite picture grammar element"
                            .into(),
                    ));
                }
            },
            Event::CData(_) | Event::GeneralRef(_) | Event::Comment(_)
                if stack.last().is_some_and(|(node, _)| {
                    matches!(
                        node,
                        PictureNode::Picture
                            | PictureNode::BlipFill
                            | PictureNode::Blip
                            | PictureNode::SourceRectangle
                            | PictureNode::Stretch
                            | PictureNode::Tile
                            | PictureNode::FillRectangle
                    )
                }) =>
            {
                return Err(Error::Invalid(
                    "picture descriptor contains unsupported content in a finite grammar element"
                        .into(),
                ));
            },
            Event::Eof => break,
            Event::Decl(_) | Event::DocType(_) | Event::PI(_) => {
                return Err(Error::Invalid(
                    "picture descriptor contains forbidden XML declarations".into(),
                ));
            },
            _ => {},
        }
    }
    if !stack.is_empty() {
        return Err(Error::Invalid(
            "picture descriptor XML ended with unclosed elements".into(),
        ));
    }
    if !root_seen {
        return Err(Error::Invalid(
            "picture descriptor XML lacks a p:pic root".into(),
        ));
    }
    if picture_stage != 3 {
        return Err(Error::Invalid(
            "picture descriptor requires exactly nvPicPr, blipFill, and spPr children".into(),
        ));
    }
    if !saw_blip {
        return Err(Error::Relationship(
            "picture descriptor is missing its a:blip relationship".into(),
        ));
    }
    if saw_stretch && !saw_fill_rectangle {
        return Err(Error::Invalid(
            "picture stretch requires exactly one direct fillRect".into(),
        ));
    }
    relationship.ok_or_else(|| {
        Error::Relationship("picture descriptor is missing its a:blip relationship".into())
    })
}

fn picture_node(namespace: &ResolveResult<'_>, name: quick_xml::name::QName<'_>) -> PictureNode {
    if is_presentation_name(namespace, name, b"pic") {
        PictureNode::Picture
    } else if is_presentation_name(namespace, name, b"nvPicPr") {
        PictureNode::NonVisualProperties
    } else if is_presentation_name(namespace, name, b"blipFill") {
        PictureNode::BlipFill
    } else if is_presentation_name(namespace, name, b"spPr") {
        PictureNode::ShapeProperties
    } else if is_drawing_name(namespace, name, b"blip") {
        PictureNode::Blip
    } else if is_drawing_name(namespace, name, b"srcRect") {
        PictureNode::SourceRectangle
    } else if is_drawing_name(namespace, name, b"stretch") {
        PictureNode::Stretch
    } else if is_drawing_name(namespace, name, b"tile") {
        PictureNode::Tile
    } else if is_drawing_name(namespace, name, b"fillRect") {
        PictureNode::FillRectangle
    } else {
        PictureNode::Other
    }
}

fn validate_blip_attributes(
    element: &quick_xml::events::BytesStart<'_>,
    decoder: quick_xml::encoding::Decoder,
    resolver: &quick_xml::name::NamespaceResolver,
) -> Result<()> {
    for attribute in element.attributes().with_checks(true) {
        let attribute = attribute.map_err(|error| Error::Xml(error.to_string()))?;
        if attribute.key.as_namespace_binding().is_some() {
            continue;
        }
        let (namespace, local) = resolver.resolve_attribute(attribute.key);
        let is_relationship = matches!(
            namespace,
            ResolveResult::Bound(Namespace(value))
                if value == litchi_ooxml_common::relationships::TRANSITIONAL_NAMESPACE
                    || value == litchi_ooxml_common::relationships::STRICT_NAMESPACE
        ) || matches!(namespace, ResolveResult::Unknown(prefix) if prefix.as_slice() == b"r");
        if !is_relationship || !matches!(local.as_ref(), b"embed" | b"link") {
            return Err(Error::Invalid(
                "picture a:blip contains an unsupported attribute".into(),
            ));
        }
        let _ = attribute
            .decoded_and_normalized_value(quick_xml::XmlVersion::Implicit1_0, decoder)
            .map_err(|error| Error::Xml(error.to_string()))?;
    }
    Ok(())
}

fn reject_picture_markup_compatibility(xml: &[u8]) -> Result<()> {
    let mut reader = NsReader::from_reader(xml);
    loop {
        let (namespace, event) = reader
            .read_resolved_event()
            .map_err(|error| Error::Xml(error.to_string()))?;
        let namespace_is_mce = is_mce_namespace(&namespace);
        match event {
            Event::Start(element) | Event::Empty(element) => {
                let mut has_mce_attribute = false;
                for attribute in element.attributes().with_checks(true) {
                    let attribute = attribute.map_err(|error| Error::Xml(error.to_string()))?;
                    has_mce_attribute |=
                        is_mce_namespace(&reader.resolver().resolve_attribute(attribute.key).0);
                }
                if namespace_is_mce || has_mce_attribute {
                    return Err(Error::UnsafeEdit {
                        operation: "source-backed picture inventory",
                        reason: "markup-compatibility elements and attributes are refused",
                    });
                }
            },
            Event::DocType(_) | Event::PI(_) => {
                return Err(Error::Invalid(
                    "slide XML contains forbidden XML declarations".into(),
                ));
            },
            Event::Eof => return Ok(()),
            _ => {},
        }
    }
}

fn validate_source_slide_root(part: &dyn Part) -> Result<()> {
    if part.content_type() != ct::PML_SLIDE {
        return Err(Error::ContentType {
            expected: ct::PML_SLIDE.to_string(),
            actual: part.content_type().to_string(),
        });
    }
    let mut reader = NsReader::from_reader(part.blob());
    loop {
        let (namespace, event) = reader
            .read_resolved_event()
            .map_err(|error| Error::Xml(error.to_string()))?;
        match event {
            Event::Start(element) | Event::Empty(element) => {
                if !is_presentation_name(&namespace, element.name(), b"sld") {
                    return Err(Error::Invalid(
                        "slide part does not have a p:sld root".into(),
                    ));
                }
                return Ok(());
            },
            Event::Decl(_) | Event::Comment(_) => {},
            Event::Text(text) if text.as_ref().iter().all(u8::is_ascii_whitespace) => {},
            Event::DocType(_) | Event::PI(_) => {
                return Err(Error::Invalid(
                    "slide part contains a forbidden XML declaration".into(),
                ));
            },
            Event::Eof => {
                return Err(Error::Invalid("slide part lacks a p:sld root".into()));
            },
            _ => {
                return Err(Error::Invalid(
                    "slide part lacks a p:sld element root".into(),
                ));
            },
        }
    }
}

fn resolve_picture_target(
    package: &SourceBackedPackage,
    view: &PartView<'_>,
    picture: &PictureRelationship,
) -> Result<SourceImageTarget> {
    let relationship = view.rels().get(&picture.id).ok_or_else(|| {
        Error::Relationship(format!(
            "picture image relationship '{}' is missing",
            picture.id
        ))
    })?;
    if relationship.is_external() != picture.external {
        return Err(Error::Relationship(format!(
            "picture image relationship '{}' has the wrong target mode",
            picture.id
        )));
    }
    if relationship.reltype() != rt::IMAGE && relationship.reltype() != rt::STRICT_IMAGE {
        return Err(Error::Relationship(format!(
            "picture image relationship '{}' has an unexpected type",
            picture.id
        )));
    }
    if relationship.is_external() {
        if relationship.target_ref().is_empty() {
            return Err(Error::Relationship(
                "external picture image relationship has an empty target".into(),
            ));
        }
        return Ok(SourceImageTarget::External {
            target: relationship.target_ref().to_owned(),
        });
    }

    let part_uri = relationship.target_partname()?;
    if !part_uri.as_str().starts_with("/ppt/media/") {
        return Err(Error::Relationship(format!(
            "picture image target '{}' is outside /ppt/media",
            part_uri.as_str()
        )));
    }
    let media = package.part(&part_uri)?;
    if !is_image_content_type(media.content_type()) {
        return Err(Error::ContentType {
            expected: "image/*".into(),
            actual: media.content_type().to_owned(),
        });
    }
    if !media.rels().is_empty() {
        return Err(Error::Relationship(format!(
            "picture image part '{}' has outbound relationships",
            part_uri.as_str()
        )));
    }
    Ok(SourceImageTarget::Internal {
        part_uri,
        content_type: media.content_type().to_owned(),
    })
}

fn is_image_content_type(content_type: &str) -> bool {
    content_type.starts_with("image/")
}

fn is_mce_namespace(namespace: &ResolveResult<'_>) -> bool {
    matches!(namespace, ResolveResult::Bound(Namespace(value)) if *value == MCE_NAMESPACE)
}

fn is_presentation_name(
    namespace: &ResolveResult<'_>,
    name: quick_xml::name::QName<'_>,
    local: &[u8],
) -> bool {
    crate::namespace::is_presentationml_name(namespace, name, local)
}

fn is_drawing_name(
    namespace: &ResolveResult<'_>,
    name: quick_xml::name::QName<'_>,
    local: &[u8],
) -> bool {
    name.local_name().as_ref() == local
        && (matches!(
            namespace,
            ResolveResult::Bound(Namespace(value))
                if *value == DRAWINGML_NAMESPACE || *value == STRICT_DRAWINGML_NAMESPACE
        ) || matches!(namespace, ResolveResult::Unknown(prefix) if prefix.as_slice() == b"a"))
}

fn validate_slide_graph(
    package: &SourceBackedPackage,
    presentation: &PartView<'_>,
    references: &[SlideReference],
) -> Result<Box<[Arc<SourceSlideData>]>> {
    package.check_execution()?;
    let mut slides = Vec::new();
    slides
        .try_reserve_exact(references.len())
        .map_err(|source| Error::Allocation {
            resource: "source-backed presentation slide graph",
            source,
        })?;
    for (position, reference) in references.iter().enumerate() {
        package.check_execution()?;
        let (relationship, part_uri, slide) = crate::parts::validate_slide_relationship(
            presentation.rels().get(reference.relationship_id()),
            reference.relationship_id(),
            |target| Ok(package.part(target)?),
            |part| part.content_type(),
        )?;
        slides.push(Arc::new(SourceSlideData {
            position,
            slide_id: reference.id(),
            part_uri: part_uri.clone(),
            binding: Arc::new(SlideBinding {
                slide_reference_id: reference.relationship_id().to_string(),
                presentation_relationship: relationship_binding(relationship),
                slide: part_binding(&slide),
            }),
        }));
    }
    package.check_execution()?;
    Ok(slides.into_boxed_slice())
}

fn source_slide_output_limit(limits: ReadLimits) -> usize {
    let bounded = limits
        .max_part_bytes()
        .min(limits.max_archive_entry_bytes())
        .min(limits.max_total_part_bytes())
        .min(limits.max_archive_total_bytes());
    usize::try_from(bounded)
        .unwrap_or(usize::MAX)
        .min(crate::shape::Limits::DEFAULT.output_bytes())
}

fn source_catalog(package: &SourceBackedPackage) -> Result<SourceCatalog> {
    #[cfg(test)]
    SOURCE_CATALOG_BUILDS.set(SOURCE_CATALOG_BUILDS.get() + 1);
    package.check_execution()?;
    let view = package.main_document_part()?;
    let presentation = SourcePart::from_view(&view, view.data()?);
    package.check_execution()?;
    let presentation_binding = part_binding(&view);
    let references = PresentationPart::from_part(&presentation)?.slide_references()?;
    package.check_execution()?;
    let slides = validate_slide_graph(package, &view, &references)?;
    package.check_execution()?;
    Ok(SourceCatalog {
        presentation,
        package_relationships: relationship_bindings(package.rels()),
        presentation_binding,
        slides,
    })
}

fn part_binding(view: &PartView<'_>) -> PartBinding {
    PartBinding {
        uri: view.partname().clone(),
        content_type: view.content_type().to_string(),
        relationships: relationship_bindings(view.rels()),
    }
}

fn relationship_bindings(relationships: &Relationships) -> Box<[RelationshipBinding]> {
    let mut bindings = relationships
        .iter()
        .map(relationship_binding)
        .collect::<Vec<_>>();
    bindings.sort_unstable_by(|left, right| left.id.cmp(&right.id));
    bindings.into_boxed_slice()
}

fn relationship_binding(relationship: &litchi_opc::Relationship) -> RelationshipBinding {
    RelationshipBinding {
        id: relationship.r_id().to_string(),
        kind: relationship.reltype().to_string(),
        target: relationship.target_ref().to_string(),
        mode: relationship.target_mode(),
    }
}

#[cfg(test)]
mod tests {
    #[cfg(any(unix, windows))]
    use std::fs;
    use std::io;
    use std::num::{NonZeroU64, NonZeroUsize};
    use std::sync::Arc;
    use std::sync::atomic::{AtomicU64, AtomicUsize, Ordering};

    use litchi_core::{
        Budget, CancellationSource, ExecutionContext, ExecutionLimits, Limits, ReadAt, Resource,
        SourceVersion,
    };
    use litchi_opc::{PackURI, ReadLimits, ReadResource, SourceBackedPackage};
    use soapberry_zip::office::StreamingArchiveWriter;
    #[cfg(any(unix, windows))]
    use tempfile::NamedTempFile;

    use super::{
        SourceBackedPresentation, SourceBackedPresentationEditor, SourceImageTarget,
        reset_source_catalog_builds, source_catalog_builds,
    };
    use crate::Error;

    const SECOND_MARKER: &[u8] = b"source-backed-unrequested-second-slide";
    const MEDIA_MARKER: &[u8] = b"source-backed-picture-media-payload";
    const CORE_MARKER: &[u8] = b"source-backed-core-properties-payload";
    const COLD_SLIDE_MARKER: &[u8] = b"source-backed-core-cold-slide-payload";
    const COLD_MEDIA_MARKER: &[u8] = b"source-backed-core-cold-media-payload";

    struct CountingSource {
        bytes: Vec<u8>,
        marker_offset: usize,
        second_payload_reads: AtomicUsize,
        revision: AtomicU64,
    }

    impl CountingSource {
        fn new(bytes: Vec<u8>) -> Self {
            let marker_offset = bytes
                .windows(SECOND_MARKER.len())
                .position(|window| window == SECOND_MARKER)
                .expect("second slide marker is stored in archive");
            Self {
                bytes,
                marker_offset,
                second_payload_reads: AtomicUsize::new(0),
                revision: AtomicU64::new(0),
            }
        }

        fn changed(&self) {
            self.revision.fetch_add(1, Ordering::SeqCst);
        }
    }

    impl ReadAt for CountingSource {
        fn len(&self) -> io::Result<u64> {
            Ok(self.bytes.len() as u64)
        }

        fn read_at(&self, offset: u64, output: &mut [u8]) -> io::Result<usize> {
            let offset = usize::try_from(offset)
                .map_err(|_| io::Error::new(io::ErrorKind::InvalidInput, "offset too large"))?;
            if offset >= self.bytes.len() {
                return Ok(0);
            }
            let count = output.len().min(self.bytes.len() - offset);
            let end = offset + count;
            if offset < self.marker_offset + SECOND_MARKER.len() && self.marker_offset < end {
                self.second_payload_reads.fetch_add(1, Ordering::SeqCst);
            }
            output[..count].copy_from_slice(&self.bytes[offset..end]);
            Ok(count)
        }

        fn version(&self) -> io::Result<SourceVersion> {
            Ok(SourceVersion::new(92, self.revision.load(Ordering::SeqCst)))
        }
    }

    struct PictureCountingSource {
        bytes: Vec<u8>,
        marker_offset: usize,
        media_payload_reads: AtomicUsize,
        revision: AtomicU64,
    }

    impl PictureCountingSource {
        fn new(bytes: Vec<u8>) -> Self {
            let marker_offset = bytes
                .windows(MEDIA_MARKER.len())
                .position(|window| window == MEDIA_MARKER)
                .expect("picture media marker is stored in archive");
            Self {
                bytes,
                marker_offset,
                media_payload_reads: AtomicUsize::new(0),
                revision: AtomicU64::new(0),
            }
        }

        fn changed(&self) {
            self.revision.fetch_add(1, Ordering::SeqCst);
        }
    }

    impl ReadAt for PictureCountingSource {
        fn len(&self) -> io::Result<u64> {
            Ok(self.bytes.len() as u64)
        }

        fn read_at(&self, offset: u64, output: &mut [u8]) -> io::Result<usize> {
            let offset = usize::try_from(offset)
                .map_err(|_| io::Error::new(io::ErrorKind::InvalidInput, "offset too large"))?;
            if offset >= self.bytes.len() {
                return Ok(0);
            }
            let count = output.len().min(self.bytes.len() - offset);
            let end = offset + count;
            if offset < self.marker_offset + MEDIA_MARKER.len() && self.marker_offset < end {
                self.media_payload_reads.fetch_add(1, Ordering::SeqCst);
            }
            output[..count].copy_from_slice(&self.bytes[offset..end]);
            Ok(count)
        }

        fn version(&self) -> io::Result<SourceVersion> {
            Ok(SourceVersion::new(93, self.revision.load(Ordering::SeqCst)))
        }
    }

    struct CorePropertyCountingSource {
        bytes: Vec<u8>,
        core_offset: usize,
        slide_offset: usize,
        media_offset: usize,
        core_payload_reads: AtomicUsize,
        slide_payload_reads: AtomicUsize,
        media_payload_reads: AtomicUsize,
    }

    impl CorePropertyCountingSource {
        fn new(bytes: Vec<u8>) -> Self {
            let offset = |marker: &[u8]| {
                bytes
                    .windows(marker.len())
                    .position(|window| window == marker)
                    .expect("source-backed core counter marker is stored in archive")
            };
            let core_offset = offset(CORE_MARKER);
            let slide_offset = offset(COLD_SLIDE_MARKER);
            let media_offset = offset(COLD_MEDIA_MARKER);
            Self {
                bytes,
                core_offset,
                slide_offset,
                media_offset,
                core_payload_reads: AtomicUsize::new(0),
                slide_payload_reads: AtomicUsize::new(0),
                media_payload_reads: AtomicUsize::new(0),
            }
        }

        fn record(offset: usize, count: usize, marker_offset: usize, marker: &[u8]) -> bool {
            let Some(end) = offset.checked_add(count) else {
                return false;
            };
            offset < marker_offset.saturating_add(marker.len()) && marker_offset < end
        }
    }

    impl ReadAt for CorePropertyCountingSource {
        fn len(&self) -> io::Result<u64> {
            Ok(self.bytes.len() as u64)
        }

        fn read_at(&self, offset: u64, output: &mut [u8]) -> io::Result<usize> {
            let offset = usize::try_from(offset)
                .map_err(|_| io::Error::new(io::ErrorKind::InvalidInput, "offset too large"))?;
            if offset >= self.bytes.len() {
                return Ok(0);
            }
            let count = output.len().min(self.bytes.len() - offset);
            let end = offset + count;
            if Self::record(offset, count, self.core_offset, CORE_MARKER) {
                self.core_payload_reads.fetch_add(1, Ordering::SeqCst);
            }
            if Self::record(offset, count, self.slide_offset, COLD_SLIDE_MARKER) {
                self.slide_payload_reads.fetch_add(1, Ordering::SeqCst);
            }
            if Self::record(offset, count, self.media_offset, COLD_MEDIA_MARKER) {
                self.media_payload_reads.fetch_add(1, Ordering::SeqCst);
            }
            output[..count].copy_from_slice(&self.bytes[offset..end]);
            Ok(count)
        }

        fn version(&self) -> io::Result<SourceVersion> {
            Ok(SourceVersion::new(94, 0))
        }
    }

    fn source_backed_pptx() -> Vec<u8> {
        let mut writer = StreamingArchiveWriter::new();
        writer
            .write_stored(
                "[Content_Types].xml",
                br#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/><Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/><Override PartName="/ppt/slides/slide2.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/></Types>"#,
            )
            .unwrap();
        writer
            .write_stored(
                "_rels/.rels",
                br#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/></Relationships>"#,
            )
            .unwrap();
        writer
            .write_stored(
                "ppt/presentation.xml",
                br#"<p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><p:sldIdLst><p:sldId id="256" r:id="rId1"/><p:sldId id="257" r:id="rId2"/></p:sldIdLst><p:sldSz cx="9144000" cy="6858000"/></p:presentation>"#,
            )
            .unwrap();
        writer
            .write_stored(
                "ppt/_rels/presentation.xml.rels",
                br#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide2.xml"/></Relationships>"#,
            )
            .unwrap();
        writer
            .write_stored(
                "ppt/slides/slide1.xml",
                br#"<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><p:cSld><p:spTree><a:t>First slide</a:t></p:spTree></p:cSld></p:sld>"#,
            )
            .unwrap();
        let padding = "x".repeat(128 * 1024);
        let second = format!(
            r#"<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><!--{marker}{padding}--><p:cSld><p:spTree><a:t>Second slide</a:t></p:spTree></p:cSld></p:sld>"#,
            marker = std::str::from_utf8(SECOND_MARKER).unwrap(),
        );
        writer
            .write_stored("ppt/slides/slide2.xml", second.as_bytes())
            .unwrap();
        writer.finish_to_bytes().unwrap()
    }

    fn source_backed_pptx_with_core_properties() -> Vec<u8> {
        let mut writer = StreamingArchiveWriter::new();
        writer
            .write_stored(
                "[Content_Types].xml",
                br#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Default Extension="png" ContentType="image/png"/><Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/><Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/><Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/></Types>"#,
            )
            .unwrap();
        writer
            .write_stored(
                "_rels/.rels",
                br#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/></Relationships>"#,
            )
            .unwrap();
        writer
            .write_stored(
                "ppt/presentation.xml",
                br#"<p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><p:sldIdLst><p:sldId id="256" r:id="rId1"/></p:sldIdLst><p:sldSz cx="9144000" cy="6858000"/></p:presentation>"#,
            )
            .unwrap();
        writer
            .write_stored(
                "ppt/_rels/presentation.xml.rels",
                br#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide1.xml"/></Relationships>"#,
            )
            .unwrap();
        writer
            .write_stored(
                "ppt/slides/slide1.xml",
                format!(
                    r#"<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:cSld><p:spTree><p:sp>{}</p:sp></p:spTree></p:cSld></p:sld>"#,
                    std::str::from_utf8(COLD_SLIDE_MARKER).unwrap()
                )
                .as_bytes(),
            )
            .unwrap();
        writer
            .write_stored(
                "ppt/slides/_rels/slide1.xml.rels",
                br#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdImage" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image1.png"/></Relationships>"#,
            )
            .unwrap();
        writer
            .write_stored("ppt/media/image1.png", COLD_MEDIA_MARKER)
            .unwrap();
        writer
            .write_stored(
                "docProps/core.xml",
                format!(
                    r#"<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/"><dc:title>source-backed core</dc:title><dc:description>{}</dc:description></cp:coreProperties>"#,
                    std::str::from_utf8(CORE_MARKER).unwrap()
                )
                .as_bytes(),
            )
            .unwrap();
        writer.finish_to_bytes().unwrap()
    }

    fn picture_slide(blip: &str) -> Vec<u8> {
        let children = format!(
            r#"<p:nvPicPr><p:cNvPr id="42" name="Photo"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr><p:blipFill>{blip}<a:stretch><a:fillRect/></a:stretch></p:blipFill><p:spPr><a:xfrm><a:off x="1" y="2"/><a:ext cx="3" cy="4"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr>"#
        );
        picture_slide_with_pic_children(&children)
    }

    fn picture_slide_with_pic_children(children: &str) -> Vec<u8> {
        format!(
            r#"<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><p:cSld><p:spTree><p:nvGrpSpPr/><p:grpSpPr/><p:pic>{children}</p:pic></p:spTree></p:cSld></p:sld>"#
        )
        .into_bytes()
    }

    fn picture_pptx(
        slide: &[u8],
        slide_relationships: &[u8],
        media_content_type: Option<&str>,
        media_relationships: Option<&[u8]>,
    ) -> Vec<u8> {
        let mut writer = StreamingArchiveWriter::new();
        let media_default = media_content_type.map_or_else(String::new, |content_type| {
            format!(r#"<Default Extension="png" ContentType="{content_type}"/>"#)
        });
        let content_types = format!(
            r#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>{media_default}<Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/><Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/></Types>"#
        );
        writer
            .write_stored("[Content_Types].xml", content_types.as_bytes())
            .unwrap();
        writer
            .write_stored(
                "_rels/.rels",
                br#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/></Relationships>"#,
            )
            .unwrap();
        writer
            .write_stored(
                "ppt/presentation.xml",
                br#"<p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><p:sldIdLst><p:sldId id="256" r:id="rId1"/></p:sldIdLst><p:sldSz cx="9144000" cy="6858000"/></p:presentation>"#,
            )
            .unwrap();
        writer
            .write_stored(
                "ppt/_rels/presentation.xml.rels",
                br#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide1.xml"/></Relationships>"#,
            )
            .unwrap();
        writer.write_stored("ppt/slides/slide1.xml", slide).unwrap();
        writer
            .write_stored("ppt/slides/_rels/slide1.xml.rels", slide_relationships)
            .unwrap();
        if media_content_type.is_some() {
            writer
                .write_stored("ppt/media/image1.png", &media_payload())
                .unwrap();
            if let Some(relationships) = media_relationships {
                writer
                    .write_stored("ppt/media/_rels/image1.png.rels", relationships)
                    .unwrap();
            }
        }
        writer.finish_to_bytes().unwrap()
    }

    fn media_payload() -> Vec<u8> {
        let mut payload = MEDIA_MARKER.to_vec();
        payload.extend(std::iter::repeat_n(b'!', 2048));
        payload
    }

    fn embedded_picture_pptx() -> Vec<u8> {
        picture_pptx(
            &picture_slide(r#"<a:blip r:embed="rIdImage"/>"#),
            br#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdImage" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image1.png"/></Relationships>"#,
            Some("image/png"),
            None,
        )
    }

    fn source_backed_context(memory_limit: u64) -> (Budget, CancellationSource, ExecutionContext) {
        let budget = Budget::root(
            "pptx-source-managed-test",
            Limits::new(
                memory_limit,
                u64::MAX,
                u64::MAX,
                u64::MAX,
                u64::MAX,
                u64::MAX,
            ),
        );
        let (cancellation_source, cancellation) = CancellationSource::pair();
        let execution_limits = ExecutionLimits::new(
            NonZeroUsize::new(1).unwrap(),
            NonZeroUsize::new(1).unwrap(),
            NonZeroU64::new(memory_limit.max(1)).unwrap(),
            0,
        )
        .unwrap();
        let context = ExecutionContext::new(budget.clone(), cancellation, execution_limits);
        (budget, cancellation_source, context)
    }

    #[cfg(any(unix, windows))]
    fn temporary_source(bytes: &[u8]) -> NamedTempFile {
        let file = NamedTempFile::with_suffix(".pptx").unwrap();
        fs::write(file.path(), bytes).unwrap();
        file
    }

    fn source_backed_root_and_second_slide_bytes() -> (u64, u64) {
        let package =
            SourceBackedPackage::from_read_at(Arc::new(CountingSource::new(source_backed_pptx())))
                .unwrap();
        let presentation = package.main_document_part().unwrap();
        let root = presentation.data().unwrap().as_bytes().len() as u64;
        let second = package
            .part(&PackURI::new("/ppt/slides/slide2.xml").unwrap())
            .unwrap()
            .data()
            .unwrap()
            .as_bytes()
            .len() as u64;
        (root, second)
    }

    #[test]
    fn source_presentation_exposes_root_size_and_single_slide_text_name_projection() {
        let presentation = SourceBackedPresentation::from_read_at(Arc::new(CountingSource::new(
            source_backed_pptx(),
        )))
        .unwrap();
        assert_eq!(presentation.slide_size().unwrap(), (9_144_000, 6_858_000));
        let slide = presentation.slide(0).unwrap();
        assert_eq!(slide.name().unwrap(), "");
        assert_eq!(
            slide.text_and_name().unwrap(),
            ("First slide".to_owned(), String::new())
        );
    }

    #[test]
    fn source_properties_only_materialize_the_selected_core_payload() {
        let source = Arc::new(CorePropertyCountingSource::new(
            source_backed_pptx_with_core_properties(),
        ));
        let package = SourceBackedPackage::from_read_at(source.clone()).unwrap();
        let presentation = SourceBackedPresentation::from_source_backed_package(package).unwrap();
        assert_eq!(source.core_payload_reads.load(Ordering::SeqCst), 0);
        assert_eq!(source.slide_payload_reads.load(Ordering::SeqCst), 0);
        assert_eq!(source.media_payload_reads.load(Ordering::SeqCst), 0);

        presentation.check_source().unwrap();
        assert_eq!(source.core_payload_reads.load(Ordering::SeqCst), 0);
        assert_eq!(source.slide_payload_reads.load(Ordering::SeqCst), 0);
        assert_eq!(source.media_payload_reads.load(Ordering::SeqCst), 0);

        let properties = presentation.properties().unwrap().unwrap();
        assert_eq!(properties.title.as_deref(), Some("source-backed core"));
        assert!(source.core_payload_reads.load(Ordering::SeqCst) > 0);
        assert_eq!(source.slide_payload_reads.load(Ordering::SeqCst), 0);
        assert_eq!(source.media_payload_reads.load(Ordering::SeqCst), 0);
    }

    #[test]
    fn managed_selected_slide_materialization_is_exact_and_one_under_fails_before_io() {
        let archive = source_backed_pptx();
        let (root_bytes, second_bytes) = source_backed_root_and_second_slide_bytes();
        let exact_source = Arc::new(CountingSource::new(archive.clone()));
        let (exact_budget, _cancel, exact_context) =
            source_backed_context(root_bytes + second_bytes);
        let exact_editor =
            SourceBackedPresentationEditor::from_read_at_with_limits_and_execution_context(
                exact_source.clone(),
                ReadLimits::default(),
                exact_context,
            )
            .unwrap();
        assert_eq!(exact_editor.cache_diagnostics().retained_entries, 1);
        let snapshot = exact_editor.slide_snapshot(1).unwrap();
        assert_eq!(snapshot.position(), 1);
        assert!(exact_source.second_payload_reads.load(Ordering::SeqCst) > 0);
        assert_eq!(
            exact_budget.used(Resource::Memory),
            root_bytes + second_bytes
        );
        assert_eq!(exact_editor.cache_diagnostics().retained_entries, 2);
        drop(snapshot);
        drop(exact_editor);
        assert_eq!(exact_budget.used(Resource::Memory), 0);

        let under_source = Arc::new(CountingSource::new(archive));
        let (under_budget, _cancel, under_context) =
            source_backed_context(root_bytes + second_bytes - 1);
        let under_editor =
            SourceBackedPresentationEditor::from_read_at_with_limits_and_execution_context(
                under_source.clone(),
                ReadLimits::default(),
                under_context,
            )
            .unwrap();
        assert!(under_editor.slide_snapshot(1).is_err());
        assert_eq!(under_source.second_payload_reads.load(Ordering::SeqCst), 0);
        drop(under_editor);
        assert_eq!(under_budget.used(Resource::Memory), 0);
    }

    #[test]
    fn managed_editor_exact_noop_keeps_source_bytes_and_budgeted_handles() {
        let archive = source_backed_pptx();
        let source = Arc::new(CountingSource::new(archive.clone()));
        let (budget, _cancel, context) = source_backed_context(archive.len() as u64);
        let editor = SourceBackedPresentationEditor::from_read_at_with_execution_context(
            source,
            ReadLimits::default(),
            context,
        )
        .unwrap();
        let mut edit = editor.edit_slide(0).unwrap();
        assert!(!edit.clear_transition().unwrap());
        let commit = edit.commit();
        assert!(!commit.is_changed());
        let mut output = Vec::new();
        editor
            .publish_slide_commit_to_stream(&mut output, &commit)
            .unwrap();
        assert_eq!(output, archive);
        drop(commit);
        assert_eq!(budget.used(Resource::Memory), 0);
    }

    #[test]
    fn editor_reuses_one_catalog_across_capture_and_publication() {
        reset_source_catalog_builds();
        let archive = source_backed_pptx();
        let editor = SourceBackedPresentationEditor::from_read_at(Arc::new(CountingSource::new(
            archive.clone(),
        )))
        .unwrap();
        assert_eq!(source_catalog_builds(), 1);

        let mut edit = editor.edit_slide(0).unwrap();
        assert!(!edit.clear_transition().unwrap());
        let commit = edit.commit();
        assert_eq!(source_catalog_builds(), 1);

        let mut output = Vec::new();
        editor
            .publish_slide_commit_to_stream(&mut output, &commit)
            .unwrap();
        assert_eq!(source_catalog_builds(), 1);
        assert_eq!(output, archive);
    }

    #[test]
    fn editor_reuses_validated_relationship_closures_across_snapshots() {
        let editor = SourceBackedPresentationEditor::from_read_at(Arc::new(CountingSource::new(
            source_backed_pptx(),
        )))
        .unwrap();
        let first = editor.slide_snapshot(0).unwrap();
        let second = editor.slide_snapshot(0).unwrap();

        // Capturing a second snapshot must reuse the editor's immutable
        // relationship/dependency index rather than cloning its strings.
        assert!(Arc::ptr_eq(
            &first.closure.package_relationships,
            &second.closure.package_relationships
        ));
        assert!(Arc::ptr_eq(
            &first.closure.presentation,
            &second.closure.presentation
        ));
        assert!(Arc::ptr_eq(&first.closure.slide, &second.closure.slide));

        let mut edit = first.edit();
        assert!(!edit.clear_transition().unwrap());
        let commit = edit.commit();
        assert!(!commit.is_changed());
        assert!(commit.patch().apply(&second).is_ok());
    }

    #[test]
    fn managed_source_cancellation_refuses_cached_semantic_reads() {
        let source = Arc::new(CountingSource::new(source_backed_pptx()));
        let (budget, cancellation_source, context) = source_backed_context(u64::MAX);
        let presentation = SourceBackedPresentation::from_read_at_with_execution_context(
            source,
            ReadLimits::default(),
            context,
        )
        .unwrap();
        let slide = presentation.slide(0).unwrap();
        assert_eq!(slide.text().unwrap(), "First slide");
        presentation.check_source().unwrap();
        cancellation_source.cancel();
        assert!(matches!(
            presentation.check_source(),
            Err(Error::Opc(litchi_opc::OpcError::Cancelled))
        ));
        assert!(matches!(
            slide.text(),
            Err(Error::Opc(litchi_opc::OpcError::Cancelled))
        ));
        drop(slide);
        drop(presentation);
        assert_eq!(budget.used(Resource::Memory), 0);
    }

    #[test]
    fn managed_presentation_diagnostics_report_cache_mode_and_release_on_drop() {
        let source = Arc::new(CountingSource::new(source_backed_pptx()));
        let (budget, _cancel, context) = source_backed_context(u64::MAX);
        let presentation = SourceBackedPresentation::from_read_at_with_execution_context(
            source,
            ReadLimits::default(),
            context,
        )
        .unwrap();
        let opening = presentation.cache_diagnostics();
        assert!(opening.budget_managed);
        assert_eq!(opening.retained_entries, 1);
        assert!(opening.budget_cache_reserved_bytes > 0);
        let _ = presentation.slide(0).unwrap().text().unwrap();
        let selected = presentation.cache_diagnostics();
        assert!(selected.budget_managed);
        assert_eq!(selected.retained_entries, 2);
        assert!(selected.budget_cache_reserved_bytes > opening.budget_cache_reserved_bytes);
        drop(presentation);
        assert_eq!(budget.used(Resource::Memory), 0);
    }

    #[test]
    fn managed_selected_slide_payload_is_evictable_after_view_drop() {
        let source = Arc::new(CountingSource::new(source_backed_pptx()));
        let (_budget, _cancel, context) = source_backed_context(u64::MAX);
        let cache_limits = litchi_opc::SourceCacheLimits::new(usize::MAX, 2).unwrap();
        let presentation = SourceBackedPresentation::from_read_at_with_limits_and_cache_limits_and_execution_context(
            source,
            ReadLimits::default(),
            cache_limits,
            context,
        )
        .unwrap();

        presentation.slide(0).unwrap().text().unwrap();
        let after_first = presentation.cache_diagnostics();
        assert_eq!(after_first.retained_entries, 2);

        // The root remains intentionally pinned by the facade. The selected
        // read-only slide handle has already dropped its PartData, so loading
        // the next slide can evict that payload instead of bypassing cache
        // retention because a metadata owner still holds it forever.
        presentation.slide(1).unwrap().text().unwrap();
        let after_second = presentation.cache_diagnostics();
        assert_eq!(after_second.retained_entries, 2);
        assert!(after_second.evictions >= 1);
    }

    #[test]
    fn identical_foreign_source_closure_is_rejected_by_lineage() {
        let first = SourceBackedPresentationEditor::from_read_at(Arc::new(CountingSource::new(
            source_backed_pptx(),
        )))
        .unwrap();
        let second = SourceBackedPresentationEditor::from_read_at(Arc::new(CountingSource::new(
            source_backed_pptx(),
        )))
        .unwrap();
        let mut edit = first.edit_slide(0).unwrap();
        assert!(!edit.clear_transition().unwrap());
        let patch = edit.commit_checked().unwrap().patch().clone();
        let foreign = second.slide_snapshot(0).unwrap();

        assert!(matches!(patch.apply(&foreign), Err(Error::StaleSource)));
    }

    #[test]
    fn managed_snapshot_edit_and_patch_checks_are_cooperatively_cancelled() {
        let (_budget, cancellation_source, context) = source_backed_context(u64::MAX);
        let editor = SourceBackedPresentationEditor::from_read_at_with_execution_context(
            Arc::new(CountingSource::new(source_backed_pptx())),
            ReadLimits::default(),
            context,
        )
        .unwrap();
        let snapshot = editor.slide_snapshot(0).unwrap();
        let patch = snapshot.edit().commit_checked().unwrap().patch().clone();
        cancellation_source.cancel();

        assert!(matches!(
            snapshot.transition(),
            Err(Error::Opc(litchi_opc::OpcError::Cancelled))
        ));
        assert!(matches!(
            snapshot.edit_checked(),
            Err(Error::Opc(litchi_opc::OpcError::Cancelled))
        ));
        let mut edit = snapshot.edit();
        assert!(matches!(
            edit.clear_transition(),
            Err(Error::Opc(litchi_opc::OpcError::Cancelled))
        ));
        assert!(matches!(
            edit.commit_checked(),
            Err(Error::Opc(litchi_opc::OpcError::Cancelled))
        ));
        assert!(matches!(
            patch.apply(&snapshot),
            Err(Error::Opc(litchi_opc::OpcError::Cancelled))
        ));
    }

    #[test]
    fn catalog_and_selected_text_leave_unselected_slides_unread() {
        let source = Arc::new(CountingSource::new(source_backed_pptx()));
        let presentation = SourceBackedPresentation::from_read_at(source.clone()).unwrap();

        assert_eq!(presentation.slide_count(), 2);
        assert_eq!(
            presentation
                .slides()
                .map(|slide| (slide.position(), slide.slide_id()))
                .collect::<Vec<_>>(),
            [(0, 256), (1, 257)]
        );
        assert_eq!(source.second_payload_reads.load(Ordering::SeqCst), 0);

        let first = presentation.slide(0).unwrap();
        assert_eq!(first.text().unwrap(), "First slide");
        assert_eq!(source.second_payload_reads.load(Ordering::SeqCst), 0);

        let second = presentation.slide(1).unwrap();
        assert_eq!(second.text().unwrap(), "Second slide");
        assert!(source.second_payload_reads.load(Ordering::SeqCst) > 0);
    }

    #[test]
    fn picture_inventory_is_metadata_only_and_read_is_exact_and_cached() {
        let source = Arc::new(PictureCountingSource::new(embedded_picture_pptx()));
        let presentation = SourceBackedPresentation::from_read_at(source.clone()).unwrap();
        let slide = presentation.slide(0).unwrap();

        let images = slide.images().unwrap();
        assert_eq!(images.len(), 1);
        assert_eq!(images[0].position(), 0);
        assert_eq!(images[0].shape_position(), 0);
        assert_eq!(images[0].id(), Some(42));
        assert_eq!(images[0].name(), Some("Photo"));
        assert_eq!(
            images[0]
                .bounds()
                .map(|value| (value.x(), value.y(), value.width(), value.height())),
            Some((1, 2, 3, 4))
        );
        assert_eq!(images[0].relationship_id(), "rIdImage");
        assert_eq!(
            images[0].target().part_uri().unwrap().as_str(),
            "/ppt/media/image1.png"
        );
        assert_eq!(images[0].target().content_type(), Some("image/png"));
        assert!(!images[0].is_external());
        assert_eq!(source.media_payload_reads.load(Ordering::SeqCst), 0);

        let image = slide.read_image(0).unwrap();
        assert!(image.bytes().starts_with(MEDIA_MARKER));
        let reads_after_first = source.media_payload_reads.load(Ordering::SeqCst);
        assert!(reads_after_first > 0);
        let second = slide.read_image(0).unwrap();
        assert!(second.bytes().starts_with(MEDIA_MARKER));
        assert_eq!(
            source.media_payload_reads.load(Ordering::SeqCst),
            reads_after_first
        );
    }

    #[test]
    fn picture_selection_reports_checked_out_of_range() {
        let source = Arc::new(PictureCountingSource::new(embedded_picture_pptx()));
        let presentation = SourceBackedPresentation::from_read_at(source).unwrap();
        let slide = presentation.slide(0).unwrap();

        assert!(matches!(
            slide.image(1),
            Err(Error::IndexOutOfBounds { index: 1, len: 1 })
        ));
    }

    #[test]
    fn strict_picture_graph_is_supported_without_payload_reads() {
        let slide = String::from_utf8(picture_slide(r#"<a:blip r:embed="rIdImage"/>"#))
            .unwrap()
            .replace(
                "http://schemas.openxmlformats.org/presentationml/2006/main",
                "http://purl.oclc.org/ooxml/presentationml/main",
            )
            .replace(
                "http://schemas.openxmlformats.org/drawingml/2006/main",
                "http://purl.oclc.org/ooxml/drawingml/main",
            )
            .replace(
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
                "http://purl.oclc.org/ooxml/officeDocument/relationships",
            );
        let relationships = br#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdImage" Type="http://purl.oclc.org/ooxml/officeDocument/relationships/image" Target="../media/image1.png"/></Relationships>"#;
        let source = Arc::new(PictureCountingSource::new(picture_pptx(
            slide.as_bytes(),
            relationships,
            Some("image/png"),
            None,
        )));
        let presentation = SourceBackedPresentation::from_read_at(source.clone()).unwrap();
        let images = presentation.slide(0).unwrap().images().unwrap();
        assert_eq!(images.len(), 1);
        assert_eq!(images[0].target().content_type(), Some("image/png"));
        assert_eq!(source.media_payload_reads.load(Ordering::SeqCst), 0);
    }

    #[test]
    fn external_picture_is_listed_but_never_fetched() {
        let slide = picture_slide(r#"<a:blip r:link="rIdExternal"/>"#);
        let source = Arc::new(PictureCountingSource::new(picture_pptx(
            &slide,
            br#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdExternal" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="https://example.invalid/photo.png" TargetMode="External"/></Relationships>"#,
            Some("image/png"),
            None,
        )));
        let presentation = SourceBackedPresentation::from_read_at(source.clone()).unwrap();
        let image = presentation.slide(0).unwrap().image(0).unwrap();
        assert!(
            matches!(image.target(), SourceImageTarget::External { target } if target == "https://example.invalid/photo.png")
        );
        assert!(image.is_external());
        assert!(matches!(
            presentation.slide(0).unwrap().read_image(0),
            Err(Error::Relationship(message)) if message.contains("external target")
        ));
        assert_eq!(source.media_payload_reads.load(Ordering::SeqCst), 0);
    }

    #[test]
    fn malformed_picture_grammar_and_graphs_are_refused() {
        let cases = [
            (
                picture_slide("<a:stretch><a:fillRect/></a:stretch>"),
                br#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdImage" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image1.png"/></Relationships>"#.as_slice(),
                Some("image/png"),
            ),
            (
                picture_slide(r#"<a:blip r:embed="rIdImage"/><a:blip r:embed="rIdImage"/>"#),
                br#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdImage" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image1.png"/></Relationships>"#.as_slice(),
                Some("image/png"),
            ),
            (
                picture_slide(r#"<a:blip r:embed="rIdImage"/>"#),
                br#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdImage" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="../media/image1.png"/></Relationships>"#.as_slice(),
                Some("image/png"),
            ),
            (
                picture_slide(r#"<a:blip r:embed="rIdImage"/>"#),
                br#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdImage" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../custom/image1.png"/></Relationships>"#.as_slice(),
                Some("image/png"),
            ),
            (
                picture_slide(r#"<a:blip r:embed="rIdImage"/>"#),
                br#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdImage" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image1.png"/></Relationships>"#.as_slice(),
                Some("application/octet-stream"),
            ),
            (
                picture_slide(r#"<a:blip r:embed="rIdMissing"/>"#),
                br#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>"#.as_slice(),
                Some("image/png"),
            ),
            (
                picture_slide(r#"<a:blip r:embed="rIdImage"/>"#),
                br#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdImage" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/missing.png"/></Relationships>"#.as_slice(),
                Some("image/png"),
            ),
        ];
        for (slide, relationships, content_type) in cases {
            let source = Arc::new(PictureCountingSource::new(picture_pptx(
                &slide,
                relationships,
                content_type,
                None,
            )));
            let presentation = SourceBackedPresentation::from_read_at(source).unwrap();
            assert!(presentation.slide(0).unwrap().images().is_err());
        }
    }

    #[test]
    fn direct_picture_children_are_finite_ordered_and_complete() {
        let blip_fill = r#"<p:blipFill><a:blip r:embed="rIdImage"/><a:stretch><a:fillRect/></a:stretch></p:blipFill>"#;
        let cases = [
            // Missing nvPicPr.
            format!(r#"{blip_fill}<p:spPr/>"#),
            // Duplicate nvPicPr.
            format!(r#"<p:nvPicPr/><p:nvPicPr/>{blip_fill}<p:spPr/>"#),
            // Misordered required children.
            format!(r#"{blip_fill}<p:nvPicPr/><p:spPr/>"#),
            // Arbitrary direct child after the required sequence.
            format!(r#"<p:nvPicPr/>{blip_fill}<p:spPr/><p:style/>"#),
            // Arbitrary direct child inside blipFill.
            r#"<p:nvPicPr/><p:blipFill><a:blip r:embed="rIdImage"/><a:bad/><a:stretch><a:fillRect/></a:stretch></p:blipFill><p:spPr/>"#.to_owned(),
        ];
        for children in cases {
            let source = Arc::new(PictureCountingSource::new(picture_pptx(
                &picture_slide_with_pic_children(&children),
                br#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdImage" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image1.png"/></Relationships>"#,
                Some("image/png"),
                None,
            )));
            let presentation = SourceBackedPresentation::from_read_at(source).unwrap();
            assert!(presentation.slide(0).unwrap().images().is_err());
        }
    }

    #[test]
    fn selected_slide_root_is_validated_before_picture_scanning() {
        let wrong_root = String::from_utf8(picture_slide(r#"<a:blip r:embed="rIdImage"/>"#))
            .unwrap()
            .replace("<p:sld ", "<p:notSlide ")
            .replace("</p:sld>", "</p:notSlide>");
        let source = Arc::new(PictureCountingSource::new(picture_pptx(
            wrong_root.as_bytes(),
            br#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdImage" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image1.png"/></Relationships>"#,
            Some("image/png"),
            None,
        )));
        let presentation = SourceBackedPresentation::from_read_at(source).unwrap();
        assert!(matches!(
            presentation.slide(0).unwrap().images(),
            Err(Error::Invalid(message)) if message.contains("p:sld root")
        ));
    }

    #[test]
    fn paired_empty_blip_and_fill_rect_are_semantically_empty_but_content_is_refused() {
        let paired_empty =
            String::from_utf8(picture_slide(r#"<a:blip r:embed="rIdImage"></a:blip>"#))
                .unwrap()
                .replace("<a:fillRect/>", "<a:fillRect></a:fillRect>");
        let source = Arc::new(PictureCountingSource::new(picture_pptx(
            paired_empty.as_bytes(),
            br#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdImage" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image1.png"/></Relationships>"#,
            Some("image/png"),
            None,
        )));
        let presentation = SourceBackedPresentation::from_read_at(source.clone()).unwrap();
        assert_eq!(presentation.slide(0).unwrap().images().unwrap().len(), 1);
        assert_eq!(source.media_payload_reads.load(Ordering::SeqCst), 0);

        let nonempty_fill = paired_empty.replace(
            "<a:fillRect></a:fillRect>",
            "<a:fillRect>payload</a:fillRect>",
        );
        let source = Arc::new(PictureCountingSource::new(picture_pptx(
            nonempty_fill.as_bytes(),
            br#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdImage" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image1.png"/></Relationships>"#,
            Some("image/png"),
            None,
        )));
        let presentation = SourceBackedPresentation::from_read_at(source).unwrap();
        assert!(presentation.slide(0).unwrap().images().is_err());

        let nonempty = picture_slide(r#"<a:blip r:embed="rIdImage">payload</a:blip>"#);
        let source = Arc::new(PictureCountingSource::new(picture_pptx(
            &nonempty,
            br#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdImage" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image1.png"/></Relationships>"#,
            Some("image/png"),
            None,
        )));
        let presentation = SourceBackedPresentation::from_read_at(source).unwrap();
        assert!(presentation.slide(0).unwrap().images().is_err());
    }

    #[test]
    fn malformed_markup_compatibility_attribute_is_not_suppressed() {
        let malformed = String::from_utf8(picture_slide(r#"<a:blip r:embed="rIdImage"/>"#))
            .unwrap()
            .replace(
                "xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\"",
                "xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\"",
            )
            .replace(
                "<p:cSld>",
                "<p:cSld mc:Ignorable=\"a\" mc:Ignorable=\"p\">",
            );
        let source = Arc::new(PictureCountingSource::new(picture_pptx(
            malformed.as_bytes(),
            br#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdImage" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image1.png"/></Relationships>"#,
            Some("image/png"),
            None,
        )));
        let presentation = SourceBackedPresentation::from_read_at(source).unwrap();
        let result = presentation.slide(0).unwrap().images();
        assert!(matches!(result, Err(Error::Xml(_))));
    }

    #[test]
    fn markup_compatibility_and_media_outbound_edges_are_refused() {
        let mce_slide = br#"<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"><p:cSld><p:spTree><p:nvGrpSpPr/><p:grpSpPr/><mc:AlternateContent><mc:Fallback><p:pic><p:nvPicPr><p:cNvPr id="42" name="Photo"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr><p:blipFill><a:blip r:embed="rIdImage"/><a:stretch><a:fillRect/></a:stretch></p:blipFill><p:spPr><a:xfrm><a:off x="1" y="2"/><a:ext cx="3" cy="4"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic></mc:Fallback></mc:AlternateContent></p:spTree></p:cSld></p:sld>"#;
        let source = Arc::new(PictureCountingSource::new(picture_pptx(
            mce_slide,
            br#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdImage" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image1.png"/></Relationships>"#,
            Some("image/png"),
            None,
        )));
        let presentation = SourceBackedPresentation::from_read_at(source).unwrap();
        let result = presentation.slide(0).unwrap().images();
        assert!(matches!(result, Err(Error::UnsafeEdit { .. })));

        let source = Arc::new(PictureCountingSource::new(picture_pptx(
            &picture_slide(r#"<a:blip r:embed="rIdImage"/>"#),
            br#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdImage" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image1.png"/></Relationships>"#,
            Some("image/png"),
            Some(br#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="urn:test" Target="../custom.xml"/></Relationships>"#),
        )));
        let presentation = SourceBackedPresentation::from_read_at(source).unwrap();
        assert!(presentation.slide(0).unwrap().images().is_err());
    }

    #[test]
    fn picture_reads_preserve_source_version_and_limits() {
        let source = Arc::new(PictureCountingSource::new(embedded_picture_pptx()));
        let presentation = SourceBackedPresentation::from_read_at(source.clone()).unwrap();
        let slide = presentation.slide(0).unwrap();
        assert!(slide.images().is_ok());
        source.changed();
        assert!(matches!(
            slide.read_image(0),
            Err(Error::Opc(litchi_opc::OpcError::SourceChanged { .. }))
        ));

        let source = Arc::new(PictureCountingSource::new(embedded_picture_pptx()));
        let presentation = SourceBackedPresentation::from_read_at(source.clone()).unwrap();
        let slide = presentation.slide(0).unwrap();
        assert!(slide.read_image(0).is_ok());
        source.changed();
        assert!(matches!(
            slide.read_image(0),
            Err(Error::Opc(litchi_opc::OpcError::SourceChanged { .. }))
        ));

        let source = Arc::new(PictureCountingSource::new(embedded_picture_pptx()));
        let limits = ReadLimits::builder()
            .max_part_bytes(1024)
            .unwrap()
            .build()
            .unwrap();
        assert!(matches!(
            SourceBackedPresentation::from_read_at_with_limits(source.clone(), limits),
            Err(Error::Opc(litchi_opc::OpcError::ReadLimit { .. }))
        ));
        assert_eq!(source.media_payload_reads.load(Ordering::SeqCst), 0);
    }

    #[test]
    fn picture_source_open_honors_managed_cancellation_before_media_payload() {
        let source = Arc::new(PictureCountingSource::new(embedded_picture_pptx()));
        let budget = Budget::root(
            "pptx-picture-test",
            Limits::new(u64::MAX, u64::MAX, u64::MAX, u64::MAX, u64::MAX, u64::MAX),
        );
        let (cancellation_source, cancellation) = CancellationSource::pair();
        let execution_limits = ExecutionLimits::new(
            NonZeroUsize::new(1).unwrap(),
            NonZeroUsize::new(1).unwrap(),
            NonZeroU64::new(u64::MAX).unwrap(),
            0,
        )
        .unwrap();
        let context = ExecutionContext::new(budget.clone(), cancellation, execution_limits);
        cancellation_source.cancel();

        assert!(matches!(
            SourceBackedPackage::from_read_at_with_execution_context(
                source.clone(),
                ReadLimits::default(),
                context,
            ),
            Err(litchi_opc::OpcError::Cancelled)
        ));
        assert_eq!(source.media_payload_reads.load(Ordering::SeqCst), 0);
    }

    #[test]
    fn managed_picture_package_retains_budgeted_payload_handles() {
        let archive = embedded_picture_pptx();
        let source = Arc::new(PictureCountingSource::new(archive.clone()));
        let memory_limit = archive.len() as u64;
        let budget = Budget::root(
            "pptx-picture-managed-test",
            Limits::new(
                memory_limit,
                u64::MAX,
                u64::MAX,
                u64::MAX,
                u64::MAX,
                u64::MAX,
            ),
        );
        let (_cancellation_source, cancellation) = CancellationSource::pair();
        let execution_limits = ExecutionLimits::new(
            NonZeroUsize::new(1).unwrap(),
            NonZeroUsize::new(1).unwrap(),
            NonZeroU64::new(memory_limit.max(1)).unwrap(),
            0,
        )
        .unwrap();
        let context = ExecutionContext::new(budget.clone(), cancellation, execution_limits);
        let package = SourceBackedPackage::from_read_at_with_execution_context(
            source.clone(),
            ReadLimits::default(),
            context,
        )
        .unwrap();
        let presentation = SourceBackedPresentation::from_source_backed_package(package).unwrap();
        let slide = presentation.slide(0).unwrap();
        assert_eq!(slide.images().unwrap().len(), 1);
        assert!(budget.used(Resource::Memory) > 0);
        assert_eq!(source.media_payload_reads.load(Ordering::SeqCst), 0);

        let image = slide.read_image(0).unwrap();
        assert!(image.bytes().starts_with(MEDIA_MARKER));
        assert!(source.media_payload_reads.load(Ordering::SeqCst) > 0);
    }

    #[test]
    fn source_changes_are_returned_as_typed_opc_errors() {
        let source = Arc::new(CountingSource::new(source_backed_pptx()));
        let presentation = SourceBackedPresentation::from_read_at(source.clone()).unwrap();
        let slide = presentation.slide(0).unwrap();
        source.changed();

        assert!(matches!(
            presentation.check_source(),
            Err(Error::Opc(litchi_opc::OpcError::SourceChanged { .. }))
        ));
        assert!(matches!(
            slide.text(),
            Err(Error::Opc(litchi_opc::OpcError::SourceChanged { .. }))
        ));
    }

    #[test]
    fn opening_retains_caller_part_limits_without_reading_slide_payloads() {
        let source = Arc::new(CountingSource::new(source_backed_pptx()));
        let limits = ReadLimits::builder()
            .max_part_bytes(1024)
            .unwrap()
            .build()
            .unwrap();

        assert!(matches!(
            SourceBackedPresentation::from_read_at_with_limits(source.clone(), limits),
            Err(Error::Opc(litchi_opc::OpcError::ReadLimit {
                resource: ReadResource::PartBytes,
                maximum: 1024,
                ..
            }))
        ));
        assert_eq!(source.second_payload_reads.load(Ordering::SeqCst), 0);
    }

    #[cfg(any(unix, windows))]
    #[test]
    fn filesystem_path_constructors_keep_the_source_positional() {
        let archive = source_backed_pptx();
        let file = temporary_source(&archive);
        let cache = litchi_opc::SourceCacheLimits::new(archive.len(), 8).unwrap();

        assert_eq!(
            SourceBackedPresentation::from_path(file.path())
                .unwrap()
                .slide_count(),
            2
        );
        assert_eq!(
            SourceBackedPresentation::open(file.path())
                .unwrap()
                .slide_count(),
            2
        );
        assert_eq!(
            SourceBackedPresentation::from_path_with_limits(file.path(), ReadLimits::default())
                .unwrap()
                .slide_count(),
            2
        );
        assert_eq!(
            SourceBackedPresentation::from_path_with_cache_limits(file.path(), cache)
                .unwrap()
                .slide_count(),
            2
        );
        assert_eq!(
            SourceBackedPresentation::from_path_with_limits_and_cache_limits(
                file.path(),
                ReadLimits::default(),
                cache,
            )
            .unwrap()
            .slide_count(),
            2
        );
        assert_eq!(
            SourceBackedPresentation::open_with_limits_and_cache_limits(
                file.path(),
                ReadLimits::default(),
                cache,
            )
            .unwrap()
            .slide_count(),
            2
        );

        let (_budget, _cancellation_source, context) = source_backed_context(u64::MAX);
        assert_eq!(
            SourceBackedPresentation::from_path_with_execution_context(
                file.path(),
                ReadLimits::default(),
                context,
            )
            .unwrap()
            .slide_count(),
            2
        );

        let (_budget, _cancellation_source, context) = source_backed_context(u64::MAX);
        assert_eq!(
            SourceBackedPresentation::from_path_with_limits_and_execution_context(
                file.path(),
                ReadLimits::default(),
                context,
            )
            .unwrap()
            .slide_count(),
            2
        );

        let (_budget, _cancellation_source, context) = source_backed_context(u64::MAX);
        assert_eq!(
            SourceBackedPresentation::open_with_limits_and_cache_limits_and_execution_context(
                file.path(),
                ReadLimits::default(),
                cache,
                context,
            )
            .unwrap()
            .slide_count(),
            2
        );

        assert_eq!(
            SourceBackedPresentationEditor::from_path(file.path())
                .unwrap()
                .slide_count(),
            2
        );
        assert_eq!(
            SourceBackedPresentationEditor::open(file.path())
                .unwrap()
                .slide_count(),
            2
        );
        assert_eq!(
            SourceBackedPresentationEditor::from_path_with_limits_and_cache_limits(
                file.path(),
                ReadLimits::default(),
                cache,
            )
            .unwrap()
            .slide_count(),
            2
        );
        let (_budget, _cancellation_source, context) = source_backed_context(u64::MAX);
        assert_eq!(
            SourceBackedPresentationEditor::open_with_limits_and_cache_limits_and_execution_context(
                file.path(),
                ReadLimits::default(),
                cache,
                context,
            )
            .unwrap()
            .slide_count(),
            2
        );
    }

    #[cfg(any(unix, windows))]
    #[test]
    fn filesystem_listing_keeps_slide_payloads_cold() {
        let file = temporary_source(&source_backed_pptx());
        let presentation = SourceBackedPresentation::from_path(file.path()).unwrap();
        let opening = presentation.cache_diagnostics();
        assert_eq!(opening.retained_entries, 1);

        let slides = presentation.slides().collect::<Vec<_>>();
        assert_eq!(slides.len(), 2);
        assert_eq!(presentation.cache_diagnostics().retained_entries, 1);
        assert_eq!(presentation.cache_diagnostics().successful_loads, 1);

        assert_eq!(slides[0].text().unwrap(), "First slide");
        let selected = presentation.cache_diagnostics();
        assert_eq!(selected.retained_entries, 2);
        assert_eq!(selected.successful_loads, 2);
    }

    #[cfg(any(unix, windows))]
    #[test]
    fn filesystem_selected_image_loads_only_the_slide_and_media_closure() {
        let file = temporary_source(&embedded_picture_pptx());
        let presentation = SourceBackedPresentation::from_path(file.path()).unwrap();
        let slide = presentation.slide(0).unwrap();
        assert_eq!(presentation.cache_diagnostics().retained_entries, 1);

        let images = slide.images().unwrap();
        assert_eq!(images.len(), 1);
        let after_inventory = presentation.cache_diagnostics();
        assert_eq!(after_inventory.retained_entries, 2);
        assert_eq!(after_inventory.successful_loads, 2);

        let image = slide.read_image(0).unwrap();
        assert!(image.bytes().starts_with(MEDIA_MARKER));
        let after_read = presentation.cache_diagnostics();
        assert_eq!(after_read.retained_entries, 3);
        assert_eq!(after_read.successful_loads, 3);
    }

    #[cfg(any(unix, windows))]
    #[test]
    fn filesystem_managed_cache_releases_after_presentation_drop() {
        let archive = source_backed_pptx();
        let file = temporary_source(&archive);
        let (budget, _cancellation_source, context) = source_backed_context(u64::MAX);
        let cache = litchi_opc::SourceCacheLimits::new(archive.len(), 8).unwrap();
        let presentation =
            SourceBackedPresentation::from_path_with_limits_and_cache_limits_and_execution_context(
                file.path(),
                ReadLimits::default(),
                cache,
                context,
            )
            .unwrap();
        assert!(presentation.cache_diagnostics().budget_managed);
        assert!(budget.used(Resource::Memory) > 0);
        assert_eq!(
            presentation.slide(0).unwrap().text().unwrap(),
            "First slide"
        );
        assert!(budget.used(Resource::Memory) > 0);
        drop(presentation);
        assert_eq!(budget.used(Resource::Memory), 0);
    }

    #[cfg(any(unix, windows))]
    #[test]
    fn filesystem_editor_preserves_exact_noop_and_one_slide_overlay() {
        let archive = source_backed_pptx();
        let file = temporary_source(&archive);
        let editor = SourceBackedPresentationEditor::from_path(file.path()).unwrap();
        let mut noop = editor.edit_slide(0).unwrap();
        assert!(!noop.clear_transition().unwrap());
        let noop = noop.commit();
        let mut exact = Vec::new();
        editor
            .publish_slide_commit_to_stream(&mut exact, &noop)
            .unwrap();
        assert_eq!(exact, archive);

        let editor = SourceBackedPresentationEditor::open(file.path()).unwrap();
        let mut edit = editor.edit_slide(0).unwrap();
        assert!(
            edit.set_transition(&crate::transition::Transition::new(
                crate::transition::Kind::Fade { black: None }
            ))
            .unwrap()
        );
        let commit = edit.commit();
        let mut output = Vec::new();
        editor
            .publish_slide_commit_to_stream(&mut output, &commit)
            .unwrap();
        assert_ne!(output, archive);

        let source = litchi_opc::OpcPackage::from_bytes(&archive).unwrap();
        let candidate = litchi_opc::OpcPackage::from_bytes(&output).unwrap();
        for part in source.iter_parts() {
            let counterpart = candidate.get_part(part.partname()).unwrap();
            if part.partname().as_str() != "/ppt/slides/slide1.xml" {
                assert_eq!(counterpart.blob(), part.blob(), "{}", part.partname());
            }
        }
    }

    #[cfg(any(unix, windows))]
    #[test]
    fn filesystem_editor_reports_path_limits_and_version_conflicts() {
        let directory = tempfile::tempdir().unwrap();
        let missing = directory.path().join("missing.pptx");
        assert!(matches!(
            SourceBackedPresentation::from_path(&missing),
            Err(Error::Io(error)) if error.kind() == io::ErrorKind::NotFound
        ));
        assert!(matches!(
            SourceBackedPresentationEditor::from_path(directory.path()),
            Err(Error::Io(error))
                if matches!(
                    error.kind(),
                    io::ErrorKind::InvalidInput
                        | io::ErrorKind::PermissionDenied
                        | io::ErrorKind::IsADirectory
                )
        ));

        let archive = source_backed_pptx();
        let file = temporary_source(&archive);
        let limits = ReadLimits::builder()
            .max_part_bytes(1024)
            .unwrap()
            .build()
            .unwrap();
        assert!(matches!(
            SourceBackedPresentation::from_path_with_limits(file.path(), limits),
            Err(Error::Opc(litchi_opc::OpcError::ReadLimit { .. }))
        ));

        let editor = SourceBackedPresentationEditor::from_path(file.path()).unwrap();
        let mut edit = editor.edit_slide(0).unwrap();
        assert!(
            edit.set_transition(&crate::transition::Transition::new(
                crate::transition::Kind::Fade { black: None }
            ))
            .unwrap()
        );
        let commit = edit.commit();
        let mut replacement = archive.clone();
        replacement.extend_from_slice(b"replacement");
        fs::write(file.path(), replacement).unwrap();
        let mut output = Vec::new();
        assert!(matches!(
            editor.publish_slide_commit_to_stream(&mut output, &commit),
            Err(Error::Opc(litchi_opc::OpcError::SourceChanged { .. }))
        ));
        assert!(output.is_empty());
    }

    #[cfg(unix)]
    #[test]
    fn filesystem_path_replacement_does_not_retarget_and_pinned_change_is_detected() {
        let directory = tempfile::tempdir().unwrap();
        let path = directory.path().join("source.pptx");
        let pinned = directory.path().join("pinned-source.pptx");
        let replacement = directory.path().join("replacement.pptx");
        fs::write(&path, source_backed_pptx()).unwrap();
        fs::hard_link(&path, &pinned).unwrap();
        let presentation = SourceBackedPresentation::from_path(&path).unwrap();
        let first_slide = presentation.slide(0).unwrap();
        fs::write(&replacement, b"not a PPTX").unwrap();
        fs::rename(&replacement, &path).unwrap();
        let mut changed_pinned_source = source_backed_pptx();
        changed_pinned_source.extend_from_slice(b"changed pinned source");
        fs::write(&pinned, changed_pinned_source).unwrap();

        assert!(matches!(
            first_slide.text(),
            Err(Error::Opc(litchi_opc::OpcError::SourceChanged { .. }))
        ));
        assert!(SourceBackedPresentation::from_path(&path).is_err());
    }
}
