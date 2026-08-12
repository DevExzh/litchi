//! Immutable PresentationML reads backed by a caller-provided positional source.
//!
//! This facade keeps ordinary slide payloads deferred. Opening validates the
//! OPC catalog and mandatory presentation root, then resolves only slide
//! metadata. A slide body is loaded when a selected [`SourceSlide`] is read.

use std::io::Write;
use std::sync::{Arc, OnceLock};

use litchi_core::ReadAt;
use litchi_opc::constants::{content_type as ct, relationship_type as rt};
use litchi_opc::{
    BlobPart, PackURI, Part, PartData, PartView, ReadLimits, SourceBackedPackage, TargetMode,
};

use crate::parts::{PresentationPart, SlidePart, SlideReference};
use crate::{Error, Result};

/// Maximum number of existing slides in one source-backed batch edit.
pub const MAX_SOURCE_BACKED_SLIDE_BATCH: usize = 32;

struct SourceSlideData {
    position: usize,
    part_uri: PackURI,
    binding: SlideBinding,
    part: OnceLock<BlobPart>,
}

struct SourceInner {
    package: SourceBackedPackage,
    // Retain the pinned mandatory root without relying on the OPC payload cache.
    _presentation: BlobPart,
    slides: Box<[Arc<SourceSlideData>]>,
}

struct SourceCatalog {
    presentation: BlobPart,
    package_relationships: Box<[RelationshipBinding]>,
    presentation_binding: PartBinding,
    presentation_xml: Arc<Vec<u8>>,
    slides: Box<[Arc<SourceSlideData>]>,
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
struct SlideBinding {
    slide_reference_id: String,
    presentation_relationship: RelationshipBinding,
    slide: PartBinding,
}

#[derive(Clone)]
struct SlideClosure {
    package_relationships: Box<[RelationshipBinding]>,
    presentation: PartBinding,
    presentation_xml: Arc<Vec<u8>>,
    slide: SlideBinding,
}

/// Read-only PPTX catalog and selected-slide access over a positional source.
///
/// Opening validates the OPC catalog, package relationships, presentation
/// part, and presentation-to-slide graph. Slide payloads remain deferred until
/// [`SourceSlide::text`] selects one. The type has no edit or output APIs.
#[derive(Clone)]
pub struct SourceBackedPresentation {
    inner: Arc<SourceInner>,
}

/// A lifetime-free read-only slide handle from [`SourceBackedPresentation`].
///
/// Creating or listing handles does not read slide XML.
#[derive(Clone)]
pub struct SourceSlide {
    owner: Arc<SourceInner>,
    data: Arc<SourceSlideData>,
}

/// An owning, source-backed editor for one existing slide.
///
/// Unlike [`SourceBackedPresentation`], this type is intentionally not
/// cloneable: publishing consumes its deferred OPC source to ensure that the
/// exact source checked during editing is the source raw-copied to output.
/// It supports no package topology or relationship changes.
pub struct SourceBackedPresentationEditor {
    package: SourceBackedPackage,
    _presentation: BlobPart,
    slides: Box<[Arc<SourceSlideData>]>,
    limits: ReadLimits,
}

/// An immutable exact-source snapshot of one slide XML part.
#[derive(Clone)]
pub struct SourceBackedSlideSnapshot {
    position: usize,
    part_uri: PackURI,
    xml: Arc<Vec<u8>>,
    closure: SlideClosure,
    max_output_bytes: usize,
}

/// An isolated focused edit of one exact-source slide snapshot.
pub struct SourceBackedSlideEdit {
    source: SourceBackedSlideSnapshot,
    working: Arc<Vec<u8>>,
    operation_used: bool,
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

    /// Build the read-only PPTX facade from a validated deferred OPC package.
    ///
    /// Only the mandatory presentation payload is read by this constructor.
    ///
    /// # Errors
    ///
    /// Returns an error when the main part or its ordered slide graph is not a
    /// coherent PresentationML presentation.
    pub fn from_source_backed_package(package: SourceBackedPackage) -> Result<Self> {
        let catalog = source_catalog(&package)?;

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

    /// Iterate lightweight slide handles without reading slide payloads.
    #[must_use]
    pub fn slides(&self) -> impl ExactSizeIterator<Item = SourceSlide> + DoubleEndedIterator + '_ {
        self.inner.slides.iter().cloned().map(|data| SourceSlide {
            owner: Arc::clone(&self.inner),
            data,
        })
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
}

impl SourceBackedPresentationEditor {
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

    /// Build an owning editor from a validated deferred OPC package.
    fn from_source_backed_package(
        package: SourceBackedPackage,
        limits: ReadLimits,
    ) -> Result<Self> {
        let catalog = source_catalog(&package)?;
        Ok(Self {
            package,
            _presentation: catalog.presentation,
            slides: catalog.slides,
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
        self.slide_snapshot_for(position, "slide_snapshot")
    }

    /// Begin an isolated focused edit of one existing slide.
    pub fn edit_slide(&self, position: usize) -> Result<SourceBackedSlideEdit> {
        Ok(self.slide_snapshot_for(position, "edit_slide")?.edit())
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
        let current = self.slide_snapshot_for(
            commit.patch.source().position,
            "publish_slide_commit_to_stream",
        )?;
        let target = commit.patch.apply(&current)?;
        self.package.write_part_overlay_to_stream(
            writer,
            &current.part_uri,
            target.xml.as_ref().clone(),
        )?;
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
        let catalog = source_catalog(&self.package)?;
        let mut slides = Vec::new();
        slides
            .try_reserve_exact(commit.patch.source().slide_count())
            .map_err(|source| Error::Allocation {
                resource: "source-backed slide batch publication snapshot",
                source,
            })?;
        for source in commit.patch.source().slides() {
            slides.push(self.slide_snapshot_from_catalog(
                &catalog,
                source.position,
                "publish_slide_batch_commit_to_stream",
            )?);
        }
        let current = SourceBackedSlideBatchSnapshot {
            slides: slides.into_boxed_slice(),
        };
        let target = commit.patch.apply(&current)?;
        let mut replacements = Vec::new();
        replacements
            .try_reserve_exact(target.slide_count())
            .map_err(|source| Error::Allocation {
                resource: "source-backed slide batch publication payloads",
                source,
            })?;
        for slide in target.slides() {
            replacements.push((slide.part_uri.clone(), slide.xml.as_ref().clone()));
        }
        self.package
            .write_part_overlays_to_stream(writer, replacements)?;
        Ok(target)
    }

    fn slide_snapshot_for(
        &self,
        position: usize,
        operation: &'static str,
    ) -> Result<SourceBackedSlideSnapshot> {
        let catalog = source_catalog(&self.package)?;
        self.slide_snapshot_from_catalog(&catalog, position, operation)
    }

    fn slide_snapshot_from_catalog(
        &self,
        catalog: &SourceCatalog,
        position: usize,
        operation: &'static str,
    ) -> Result<SourceBackedSlideSnapshot> {
        let data = catalog
            .slides
            .get(position)
            .ok_or(Error::SlideIndexOutOfBounds {
                index: position,
                len: catalog.slides.len(),
            })?;
        let view = self.package.part(&data.part_uri)?;
        let raw = view.data()?.into_arc();
        let part = owned_part_shared(&view, Arc::clone(&raw))?;
        SlidePart::from_part(&part)?;
        let scene = crate::shape::Scene::read(raw.as_slice())?;
        if scene.is_rewritten() {
            return Err(Error::UnsafeEdit {
                operation,
                reason: "source-backed slide edits do not support markup-compatibility branch selection",
            });
        }
        Ok(SourceBackedSlideSnapshot {
            position,
            part_uri: data.part_uri.clone(),
            xml: raw,
            closure: SlideClosure {
                package_relationships: catalog.package_relationships.clone(),
                presentation: catalog.presentation_binding.clone(),
                presentation_xml: Arc::clone(&catalog.presentation_xml),
                slide: data.binding.clone(),
            },
            max_output_bytes: source_slide_output_limit(self.limits),
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
        crate::presentation::transition::read_direct(self.xml.as_slice())
    }

    /// Start an isolated focused edit from this exact source snapshot.
    #[must_use]
    pub fn edit(&self) -> SourceBackedSlideEdit {
        SourceBackedSlideEdit {
            source: self.clone(),
            working: Arc::clone(&self.xml),
            operation_used: false,
        }
    }

    fn same_source(&self, other: &Self) -> bool {
        self.position == other.position
            && self.part_uri == other.part_uri
            && self.xml == other.xml
            && self.closure.package_relationships == other.closure.package_relationships
            && self.closure.presentation == other.closure.presentation
            && self.closure.presentation_xml == other.closure.presentation_xml
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
        if self.operation_used {
            return Err(Error::UnsafeEdit {
                operation,
                reason: "source-backed slide edits support one atomic semantic operation",
            });
        }
        let (xml, changed) = crate::presentation::transition::stage(
            self.working.as_slice(),
            self.source.closure.presentation_xml.as_slice(),
            value,
            self.source.max_output_bytes,
            operation,
        )?;
        if let Some(xml) = xml {
            self.working = Arc::new(xml);
        }
        self.operation_used = true;
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
            self.working.as_slice(),
            replacements,
            crate::opened::Limits::default().max_text_bytes(),
            self.source.max_output_bytes,
        )?;
        if let Some(xml) = xml {
            self.working = Arc::new(xml);
        }
        self.operation_used = true;
        Ok(changed)
    }

    /// Validate and freeze this isolated edit for source-backed publication.
    #[must_use]
    pub fn commit(self) -> SourceBackedSlideCommit {
        let snapshot = SourceBackedSlideSnapshot {
            position: self.source.position,
            part_uri: self.source.part_uri.clone(),
            xml: self.working,
            closure: self.source.closure.clone(),
            max_output_bytes: self.source.max_output_bytes,
        };
        let patch = SourceBackedSlidePatch {
            before: self.source,
            after: snapshot.clone(),
        };
        SourceBackedSlideCommit { snapshot, patch }
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
        self.commits.push(edit.commit());
        self.text_bytes = text_bytes;
        Ok(changed)
    }

    /// Validate and freeze this selected slide set for publication.
    pub fn commit(mut self) -> Result<SourceBackedSlideBatchCommit> {
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

    /// Flatten DrawingML text runs from this selected slide in source order.
    ///
    /// The slide payload is loaded and retained on first use. No other slide
    /// payload is read.
    ///
    /// # Errors
    ///
    /// Returns an error if the source changed, the selected slide exceeds the
    /// retained OPC limits, or its PresentationML is malformed.
    pub fn text(&self) -> Result<String> {
        SlidePart::from_part(self.part()?)?.text()
    }

    fn part(&self) -> Result<&BlobPart> {
        // The metadata lookup keeps source-version checks active even after a
        // selected slide payload has entered the local cache.
        let view = self.owner.package.part(&self.data.part_uri)?;
        if let Some(part) = self.data.part.get() {
            return Ok(part);
        }

        let part = owned_part(&view, view.data()?)?;
        let _publish_result = self.data.part.set(part);
        self.data.part.get().ok_or_else(|| {
            Error::Invalid("source-backed slide cache did not publish a value".to_string())
        })
    }
}

fn validate_slide_graph(
    package: &SourceBackedPackage,
    presentation: &PartView<'_>,
    references: &[SlideReference],
) -> Result<Box<[Arc<SourceSlideData>]>> {
    let mut slides = Vec::new();
    slides
        .try_reserve_exact(references.len())
        .map_err(|source| Error::Allocation {
            resource: "source-backed presentation slide graph",
            source,
        })?;
    for (position, reference) in references.iter().enumerate() {
        let relationship = presentation
            .rels()
            .get(reference.relationship_id())
            .ok_or_else(|| {
                Error::Relationship(format!(
                    "presentation slide reference is missing relationship '{}'",
                    reference.relationship_id()
                ))
            })?;
        if relationship.is_external() {
            return Err(Error::Relationship(format!(
                "slide relationship '{}' must be internal",
                reference.relationship_id()
            )));
        }
        if !crate::parts::is_relationship_type(relationship.reltype(), rt::SLIDE, "slide") {
            return Err(Error::Relationship(format!(
                "relationship '{}' is not a slide relationship",
                reference.relationship_id()
            )));
        }
        let part_uri = relationship.target_partname()?;
        let slide = package.part(&part_uri)?;
        if slide.content_type() != ct::PML_SLIDE {
            return Err(Error::ContentType {
                expected: ct::PML_SLIDE.to_string(),
                actual: slide.content_type().to_string(),
            });
        }
        slides.push(Arc::new(SourceSlideData {
            position,
            part_uri: part_uri.clone(),
            binding: SlideBinding {
                slide_reference_id: reference.relationship_id().to_string(),
                presentation_relationship: relationship_binding(relationship),
                slide: part_binding(&slide),
            },
            part: OnceLock::new(),
        }));
    }
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
    let view = package.main_document_part()?;
    let presentation = owned_part(&view, view.data()?)?;
    let presentation_xml = presentation.blob_arc();
    let presentation_binding = part_binding(&view);
    let references = PresentationPart::from_part(&presentation)?.slide_references()?;
    let slides = validate_slide_graph(package, &view, &references)?;
    Ok(SourceCatalog {
        presentation,
        package_relationships: relationship_bindings(package.rels()),
        presentation_binding,
        presentation_xml,
        slides,
    })
}

fn owned_part(view: &PartView<'_>, data: PartData) -> Result<BlobPart> {
    owned_part_shared(view, data.into_arc())
}

fn owned_part_shared(view: &PartView<'_>, bytes: Arc<Vec<u8>>) -> Result<BlobPart> {
    let mut part = BlobPart::new_shared(
        view.partname().clone(),
        view.content_type().to_string(),
        bytes,
    );
    for relationship in view.rels().iter() {
        part.rels_mut().try_add_relationship(
            relationship.reltype().to_string(),
            relationship.target_ref().to_string(),
            relationship.r_id().to_string(),
            relationship.target_mode(),
        )?;
    }
    Ok(part)
}

fn part_binding(view: &PartView<'_>) -> PartBinding {
    PartBinding {
        uri: view.partname().clone(),
        content_type: view.content_type().to_string(),
        relationships: relationship_bindings(view.rels()),
    }
}

fn relationship_bindings(relationships: &litchi_opc::Relationships) -> Box<[RelationshipBinding]> {
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
    use std::io;
    use std::sync::Arc;
    use std::sync::atomic::{AtomicU64, AtomicUsize, Ordering};

    use litchi_core::{ReadAt, SourceVersion};
    use litchi_opc::{ReadLimits, ReadResource};
    use soapberry_zip::office::StreamingArchiveWriter;

    use super::SourceBackedPresentation;
    use crate::Error;

    const SECOND_MARKER: &[u8] = b"source-backed-unrequested-second-slide";

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

    #[test]
    fn catalog_and_selected_text_leave_unselected_slides_unread() {
        let source = Arc::new(CountingSource::new(source_backed_pptx()));
        let presentation = SourceBackedPresentation::from_read_at(source.clone()).unwrap();

        assert_eq!(presentation.slide_count(), 2);
        assert_eq!(
            presentation
                .slides()
                .map(|slide| slide.position())
                .collect::<Vec<_>>(),
            [0, 1]
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
    fn source_changes_are_returned_as_typed_opc_errors() {
        let source = Arc::new(CountingSource::new(source_backed_pptx()));
        let presentation = SourceBackedPresentation::from_read_at(source.clone()).unwrap();
        let slide = presentation.slide(0).unwrap();
        source.changed();

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
}
