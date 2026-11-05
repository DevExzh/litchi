//! Immutable PresentationML reads backed by a caller-provided positional source.
//!
//! This facade keeps ordinary slide payloads deferred. Opening validates the
//! OPC catalog and mandatory presentation root, then resolves only slide
//! metadata. A slide body is loaded when a selected [`SourceSlide`] is read.

pub mod primary;

use litchi_core::xml::ReaderOrigin;
use litchi_ooxml_common::xml::attributes::BytesStartExt as _;
use std::io::{Read, Write};
#[cfg(any(unix, windows))]
use std::path::Path;
use std::sync::{Arc, Mutex, OnceLock};

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
    PackURI, Part, PartData, PartView, Relationships, SourceBackedPackage, SourceCacheLimits,
    SourceLimits, SourceLineage, SourceTopologyPlan, TargetMode,
};
use quick_xml::events::Event;
use quick_xml::name::{Namespace, ResolveResult};
use quick_xml::reader::NsReader;

use crate::parts::{PresentationPart, SlidePart, SlideReference};
use crate::shape::{Bounds, Shape};
use crate::{Error, Result};

use litchi_drawingml::svg_blip::SvgBlip;

pub mod modern_comments;
pub mod raw;
pub mod related;
pub mod scenes;
pub mod sections;

mod svg_lifecycle;

pub(crate) use svg_lifecycle::SourceBackedSvgAttachmentPatch;
pub use svg_lifecycle::{AttachedSvg, SvgAttachment, SvgAttachmentReplacement};
#[cfg(test)]
mod engine_tests;

#[cfg(test)]
thread_local! {
    static SOURCE_CATALOG_BUILDS: Cell<usize> = const { Cell::new(0) };
    static SOURCE_PICTURE_QUERY_ALL: Cell<usize> = const { Cell::new(0) };
    static SOURCE_PICTURE_QUERY_SELECTED: Cell<usize> = const { Cell::new(0) };
    static SOURCE_PICTURE_DESCRIPTOR_RESERVE_CALLS: Cell<usize> = const { Cell::new(0) };
    static SOURCE_PICTURE_DESCRIPTOR_RESERVE_ITEMS: Cell<usize> = const { Cell::new(0) };
    static SOURCE_PICTURE_TARGET_RESOLUTIONS: Cell<usize> = const { Cell::new(0) };
    static SOURCE_LASER_SCANS: Cell<usize> = const { Cell::new(0) };
}

#[cfg(test)]
fn reset_source_catalog_builds() {
    SOURCE_CATALOG_BUILDS.set(0);
}

#[cfg(test)]
fn source_catalog_builds() -> usize {
    SOURCE_CATALOG_BUILDS.get()
}

#[cfg(test)]
fn reset_source_picture_query_counters() {
    SOURCE_PICTURE_QUERY_ALL.set(0);
    SOURCE_PICTURE_QUERY_SELECTED.set(0);
    SOURCE_PICTURE_DESCRIPTOR_RESERVE_CALLS.set(0);
    SOURCE_PICTURE_DESCRIPTOR_RESERVE_ITEMS.set(0);
    SOURCE_PICTURE_TARGET_RESOLUTIONS.set(0);
}

#[cfg(test)]
fn source_picture_query_all() -> usize {
    SOURCE_PICTURE_QUERY_ALL.get()
}

#[cfg(test)]
fn source_picture_query_selected() -> usize {
    SOURCE_PICTURE_QUERY_SELECTED.get()
}

#[cfg(test)]
fn source_picture_descriptor_reserve_calls() -> usize {
    SOURCE_PICTURE_DESCRIPTOR_RESERVE_CALLS.get()
}

#[cfg(test)]
fn source_picture_descriptor_reserve_items() -> usize {
    SOURCE_PICTURE_DESCRIPTOR_RESERVE_ITEMS.get()
}

#[cfg(test)]
fn source_picture_target_resolutions() -> usize {
    SOURCE_PICTURE_TARGET_RESOLUTIONS.get()
}

#[cfg(test)]
fn reset_source_laser_scans() {
    SOURCE_LASER_SCANS.set(0);
}

#[cfg(test)]
fn source_laser_scans() -> usize {
    SOURCE_LASER_SCANS.get()
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
    scenes: scenes::SceneCache,
    // Retain the pinned mandatory root without relying on the OPC payload cache.
    _presentation: SourcePart,
    pub(super) slides: Box<[Arc<SourceSlideData>]>,
    pub(super) limits: SourceLimits,
    /// Non-default semantic codec policies selected at ingress.  Keep this
    /// with the deferred owner so related-part reads and source-backed edits
    /// use the same policy as eager and owning projections.
    pub(super) semantic_limits: Option<crate::open_options::SemanticLimitsRef>,
    /// Every non-structural defect repaired while the source was opened, in
    /// discovery order: the OPC manifest repairs folded from `litchi_opc`,
    /// then the slide-list repairs of audit finding 6.7. Repairs a
    /// deferred slide read performs surface at that read's own entry point.
    /// See [`crate::leniency`].
    pub(super) tolerance: crate::leniency::ToleranceReport,
    /// Whether the slide list met the strict, pre-leniency contract at open:
    /// the tolerant scan repaired nothing, so the retained slide graph is the
    /// one [`SourceBackedPresentationEditor`] would build. Read by
    /// [`SourceBackedPresentation::try_into_editor`].
    strict_slide_list: bool,
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
    pub(super) fn from_view(view: &PartView<'_>, data: PartData) -> Result<Self> {
        let expected_len = view.rels().iter().count();
        let mut rels = Relationships::new(clone_relationship_text(
            view.partname().base_uri(),
            "source-backed relationship base URI",
        )?);
        for relationship in view.rels().iter() {
            let kind =
                clone_relationship_text(relationship.reltype(), "source-backed relationship type")?;
            let target = clone_relationship_text(
                relationship.target_ref(),
                "source-backed relationship target",
            )?;
            let id = clone_relationship_text(relationship.r_id(), "source-backed relationship ID")?;
            rels.try_add_relationship(kind, target, id, relationship.target_mode())?;
        }
        if rels.iter().count() != expected_len {
            return Err(Error::Invalid(
                "source-backed part relationship count changed during iteration".into(),
            ));
        }
        Ok(Self {
            partname: view.partname().clone(),
            content_type: clone_relationship_text(
                view.content_type(),
                "source-backed part content type",
            )?,
            data,
            replacement: None,
            rels,
        })
    }

    fn from_binding(binding: &PartBinding, data: PartData) -> Result<Self> {
        let expected_len = binding.relationships.len();
        let mut rels = Relationships::new(clone_relationship_text(
            binding.uri.base_uri(),
            "source-backed relationship base URI",
        )?);
        for relationship in binding.relationships.iter() {
            rels.try_add_relationship(
                clone_relationship_text(&relationship.kind, "source-backed relationship type")?,
                clone_relationship_text(&relationship.target, "source-backed relationship target")?,
                clone_relationship_text(&relationship.id, "source-backed relationship ID")?,
                relationship.mode,
            )?;
        }
        if rels.iter().count() != expected_len {
            return Err(Error::Invalid(
                "source-backed part relationship count changed during iteration".into(),
            ));
        }
        Ok(Self {
            partname: binding.uri.clone(),
            content_type: clone_relationship_text(
                &binding.content_type,
                "source-backed part content type",
            )?,
            data,
            replacement: None,
            rels,
        })
    }

    pub(super) fn source_data(&self) -> &PartData {
        &self.data
    }

    pub(super) fn source_partname(&self) -> &PackURI {
        &self.partname
    }

    pub(super) fn source_relationships(&self) -> &Relationships {
        &self.rels
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
    ReservedEdit {
        bytes: Arc<Vec<u8>>,
        reservation: Arc<litchi_core::Reservation>,
    },
}

impl SourcePayload {
    fn as_bytes(&self) -> &[u8] {
        match self {
            Self::Original(data) => data.as_bytes(),
            Self::Edited(data) | Self::ReservedEdit { bytes: data, .. } => data.as_slice(),
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
            Self::Edited(data) | Self::ReservedEdit { bytes: data, .. } => Some(Arc::clone(data)),
        }
    }

    #[cfg(test)]
    fn is_edited(&self) -> bool {
        matches!(self, Self::Edited(_) | Self::ReservedEdit { .. })
    }
}

/// Move an edit-owned payload into a replay recipient's execution scope.
///
/// The source patch keeps its original reservation until the patch itself is
/// dropped.  A replay therefore reserves the same transient envelope from
/// the recipient before replacing the copied source token.  If admission
/// fails, the copied token is dropped and the target is never published.
fn rebase_payload_reservation(
    payload: SourcePayload,
    context: Option<&ExecutionContext>,
) -> Result<SourcePayload> {
    let SourcePayload::ReservedEdit { bytes, reservation } = payload else {
        return Ok(payload);
    };
    let Some(context) = context else {
        return Ok(SourcePayload::Edited(bytes));
    };
    let replacement = context
        .reserve(litchi_core::Resource::Memory, reservation.amount())
        .map_err(|error| {
            Error::Opc(match error {
                ExecutionError::Cancelled => litchi_opc::OpcError::Cancelled,
                error => litchi_opc::OpcError::Execution(error),
            })
        })?;
    Ok(SourcePayload::ReservedEdit {
        bytes,
        reservation: Arc::new(replacement),
    })
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

fn capture_laser_usage(xml: &[u8], limits: crate::laser::Limits) -> Result<crate::laser::Usage> {
    #[cfg(test)]
    SOURCE_LASER_SCANS.set(SOURCE_LASER_SCANS.get().saturating_add(1));
    let mut usage = crate::laser::Usage::default();
    crate::laser::read(0, xml, &limits, &mut usage)?;
    Ok(usage)
}

fn merge_laser_usage(
    usage: &mut crate::laser::Usage,
    addition: crate::laser::Usage,
    limits: crate::laser::Limits,
) -> Result<()> {
    // Check each aggregate independently before using the compact ledger's
    // all-or-nothing addition.  `Usage::checked_add` intentionally returns
    // only `Option`, so a single generic overflow must not be misreported as
    // an XML-byte failure when the trace or point counter overflowed.
    if usage
        .total_slide_xml_bytes()
        .checked_add(addition.total_slide_xml_bytes())
        .is_none()
    {
        return Err(Error::Limit {
            kind: crate::opened::LimitKind::Codec(crate::opened::CodecLimitKind::XmlBytes),
            observed: usize::MAX,
            maximum: limits.max_total_slide_xml_bytes(),
        });
    }
    if usage
        .trace_count()
        .checked_add(addition.trace_count())
        .is_none()
    {
        return Err(Error::Limit {
            kind: crate::opened::LimitKind::Codec(crate::opened::CodecLimitKind::Traces),
            observed: usize::MAX,
            maximum: limits.max_trace_count(),
        });
    }
    if usage
        .point_count()
        .checked_add(addition.point_count())
        .is_none()
    {
        return Err(Error::Limit {
            kind: crate::opened::LimitKind::Codec(crate::opened::CodecLimitKind::TracePoints),
            observed: usize::MAX,
            maximum: limits.max_point_count(),
        });
    }
    let next = usage
        .checked_add(addition)
        .expect("aggregate counters were checked immediately above");
    if next.total_slide_xml_bytes() > limits.max_total_slide_xml_bytes() {
        return Err(Error::Limit {
            kind: crate::opened::LimitKind::Codec(crate::opened::CodecLimitKind::XmlBytes),
            observed: next.total_slide_xml_bytes(),
            maximum: limits.max_total_slide_xml_bytes(),
        });
    }
    if next.trace_count() > limits.max_trace_count() {
        return Err(Error::Limit {
            kind: crate::opened::LimitKind::Codec(crate::opened::CodecLimitKind::Traces),
            observed: next.trace_count(),
            maximum: limits.max_trace_count(),
        });
    }
    if next.point_count() > limits.max_point_count() {
        return Err(Error::Limit {
            kind: crate::opened::LimitKind::Codec(crate::opened::CodecLimitKind::TracePoints),
            observed: next.point_count(),
            maximum: limits.max_point_count(),
        });
    }
    *usage = next;
    Ok(())
}

/// Return the managed transient-memory ceiling without reusing the source
/// output ceiling. Serialized output and derived scene memory are independent
/// resource dimensions.
#[inline]
fn managed_memory_limit(context: &ExecutionContext) -> usize {
    usize::try_from(context.budget().limit(litchi_core::Resource::Memory)).unwrap_or(usize::MAX)
}

fn reserve_source_memory(
    context: &ExecutionContext,
    amount: usize,
    maximum: usize,
) -> Result<Option<Arc<litchi_core::Reservation>>> {
    let amount = u64::try_from(amount).map_err(|_| Error::Limit {
        kind: crate::opened::LimitKind::Codec(crate::opened::CodecLimitKind::MemoryBytes),
        observed: usize::MAX,
        maximum,
    })?;
    context
        .reserve(litchi_core::Resource::Memory, amount)
        .map(|reservation| Some(Arc::new(reservation)))
        .map_err(|error| {
            Error::Opc(match error {
                ExecutionError::Cancelled => litchi_opc::OpcError::Cancelled,
                error => litchi_opc::OpcError::Execution(error),
            })
        })
}

/// Reserve the full transient envelope of a source-backed text-box edit
/// before parsing or authoring. The scene is indexed once before authoring and
/// the candidate is indexed again during semantic readback, so both bounded
/// scene policies plus the escaped fragment and lexical output stay inside
/// the caller's execution-budget admission. The source output ceiling still
/// bounds the candidate bytes, while derived scene memory is governed by the
/// execution context rather than that semantic output limit.
fn reserve_source_text_box(
    context: Option<&ExecutionContext>,
    input_bytes: usize,
    text_bytes: usize,
    max_output_bytes: usize,
) -> Result<Option<Arc<litchi_core::Reservation>>> {
    // Physical output policy applies even without a managed context. Reject
    // an impossible output before charging the shared execution budget.
    let output_bound =
        crate::opened::text_box_output_bound(input_bytes, text_bytes, max_output_bytes)?;
    let Some(context) = context else {
        return Ok(None);
    };
    // The caller's semantic output ceiling and the managed execution budget
    // are independent policies.  Output admission above already ran first;
    // use the context's memory ceiling for the transient envelope so a
    // modest output limit does not reject a valid edit merely because its
    // scene indexes need more bytes than the serialized candidate.
    let maximum = managed_memory_limit(context);
    let input_scene = input_bytes
        .checked_mul(64)
        .and_then(|size| size.checked_add(1024))
        .ok_or(Error::Limit {
            kind: crate::opened::LimitKind::Codec(crate::opened::CodecLimitKind::MemoryBytes),
            observed: usize::MAX,
            maximum,
        })?;
    let output_scene = output_bound
        .checked_mul(64)
        .and_then(|size| size.checked_add(1024))
        .ok_or(Error::Limit {
            kind: crate::opened::LimitKind::Codec(crate::opened::CodecLimitKind::MemoryBytes),
            observed: usize::MAX,
            maximum,
        })?;
    let fragment_bound = output_bound.checked_sub(input_bytes).ok_or_else(|| {
        Error::internal(
            "reserve_source_text_box",
            "output bound is below the admitted input bytes",
        )
    })?;
    let transient = output_bound
        .checked_add(fragment_bound)
        .and_then(|size| size.checked_add(text_bytes.checked_mul(6)?))
        .ok_or(Error::Limit {
            kind: crate::opened::LimitKind::Codec(crate::opened::CodecLimitKind::MemoryBytes),
            observed: usize::MAX,
            maximum,
        })?;
    let amount = input_scene
        .checked_add(output_scene)
        .and_then(|size| size.checked_add(transient))
        .ok_or(Error::Limit {
            kind: crate::opened::LimitKind::Codec(crate::opened::CodecLimitKind::MemoryBytes),
            observed: usize::MAX,
            maximum,
        })?;
    if amount > maximum {
        return Err(Error::Limit {
            kind: crate::opened::LimitKind::Codec(crate::opened::CodecLimitKind::MemoryBytes),
            observed: amount,
            maximum,
        });
    }
    reserve_source_memory(context, amount, maximum)
}

/// Reserve the source shape-removal scene/index, candidate, and semantic
/// readback envelope before the native codec allocates either scene.
fn reserve_source_shape_removal(
    context: Option<&ExecutionContext>,
    input_bytes: usize,
    max_output_bytes: usize,
) -> Result<Option<Arc<litchi_core::Reservation>>> {
    let Some(context) = context else {
        return Ok(None);
    };
    let maximum = managed_memory_limit(context);
    // The exact retained length is checked by the native stage before it
    // allocates the candidate. This reservation only uses the output ceiling
    // to bound its transient envelope, never as the memory ceiling.
    let input_scene = input_bytes
        .checked_mul(64)
        .and_then(|size| size.checked_add(1024))
        .ok_or(Error::Limit {
            kind: crate::opened::LimitKind::Codec(crate::opened::CodecLimitKind::MemoryBytes),
            observed: usize::MAX,
            maximum,
        })?;
    let output_bound = input_bytes.min(max_output_bytes);
    let output_scene = output_bound
        .checked_mul(64)
        .and_then(|size| size.checked_add(1024))
        .ok_or(Error::Limit {
            kind: crate::opened::LimitKind::Codec(crate::opened::CodecLimitKind::MemoryBytes),
            observed: usize::MAX,
            maximum,
        })?;
    let amount = input_scene
        .checked_add(output_scene)
        .and_then(|size| size.checked_add(output_bound))
        .and_then(|size| size.checked_add(4096))
        .ok_or(Error::Limit {
            kind: crate::opened::LimitKind::Codec(crate::opened::CodecLimitKind::MemoryBytes),
            observed: usize::MAX,
            maximum,
        })?;
    if amount > maximum {
        return Err(Error::Limit {
            kind: crate::opened::LimitKind::Codec(crate::opened::CodecLimitKind::MemoryBytes),
            observed: amount,
            maximum,
        });
    }
    reserve_source_memory(context, amount, maximum)
}

/// Reserve the transient parse, lexical rewrite, and semantic read-back
/// envelope for one source-backed slide metadata edit before the shared
/// codec allocates its candidate XML.
fn reserve_source_slide_properties(
    context: Option<&ExecutionContext>,
    input_bytes: usize,
    name: Option<&str>,
    max_output_bytes: usize,
) -> Result<Option<Arc<litchi_core::Reservation>>> {
    let output_maximum = max_output_bytes;
    let name_bytes = name.map_or(0, str::len);
    // Attribute escaping can expand each input byte to a six-byte entity. A
    // small fixed charge covers a newly inserted attribute; the native stage
    // still applies the exact output-byte limit before allocating the result.
    let conservative_output_bound = input_bytes
        .checked_add(name_bytes.checked_mul(6).ok_or(Error::Limit {
            kind: crate::opened::LimitKind::Codec(crate::opened::CodecLimitKind::OutputBytes),
            observed: usize::MAX,
            maximum: output_maximum,
        })?)
        .and_then(|size| size.checked_add(16))
        .ok_or(Error::Limit {
            kind: crate::opened::LimitKind::Codec(crate::opened::CodecLimitKind::OutputBytes),
            observed: usize::MAX,
            maximum: output_maximum,
        })?;
    // The native stage applies the exact output limit after it knows whether
    // each requested attribute is a no-op, replacement, insertion, or
    // shortening.  Admission must not reject a valid shortening merely
    // because this conservative preallocation envelope exceeds that limit.
    let output_bound = conservative_output_bound.min(max_output_bytes);
    let Some(context) = context else {
        return Ok(None);
    };
    let maximum = managed_memory_limit(context);
    // The exact attribute splice can shorten or be a no-op, so its native
    // output check remains the authoritative output admission. This envelope
    // uses the source ceiling only to bound candidate-shaped memory.
    let input_scene = input_bytes
        .checked_mul(64)
        .and_then(|size| size.checked_add(1024))
        .ok_or(Error::Limit {
            kind: crate::opened::LimitKind::Codec(crate::opened::CodecLimitKind::MemoryBytes),
            observed: usize::MAX,
            maximum,
        })?;
    let output_scene = output_bound
        .checked_mul(64)
        .and_then(|size| size.checked_add(1024))
        .ok_or(Error::Limit {
            kind: crate::opened::LimitKind::Codec(crate::opened::CodecLimitKind::MemoryBytes),
            observed: usize::MAX,
            maximum,
        })?;
    let amount = input_scene
        .checked_add(output_scene)
        .and_then(|size| size.checked_add(output_bound))
        .and_then(|size| size.checked_add(name_bytes.checked_mul(6)?))
        .and_then(|size| size.checked_add(4096))
        .ok_or(Error::Limit {
            kind: crate::opened::LimitKind::Codec(crate::opened::CodecLimitKind::MemoryBytes),
            observed: usize::MAX,
            maximum,
        })?;
    if amount > maximum {
        return Err(Error::Limit {
            kind: crate::opened::LimitKind::Codec(crate::opened::CodecLimitKind::MemoryBytes),
            observed: amount,
            maximum,
        });
    }
    reserve_source_memory(context, amount, maximum)
}

fn reserve_source_shape_properties(
    context: Option<&ExecutionContext>,
    input_bytes: usize,
    replacements: &[crate::opened::ShapePropertiesReplacement<'_>],
    max_output_bytes: usize,
) -> Result<Option<Arc<litchi_core::Reservation>>> {
    if replacements.len() > crate::opened::MAX_SHAPE_PROPERTIES_REPLACEMENTS {
        return Err(Error::Limit {
            kind: crate::opened::LimitKind::Codec(crate::opened::CodecLimitKind::ShapeProperties),
            observed: replacements.len(),
            maximum: crate::opened::MAX_SHAPE_PROPERTIES_REPLACEMENTS,
        });
    }
    let text_maximum = crate::opened::Limits::default().max_text_bytes();
    let text_bytes = replacements.iter().try_fold(0usize, |sum, replacement| {
        sum.checked_add(replacement.text().map_or(0, str::len))
            .ok_or(Error::Limit {
                kind: crate::opened::LimitKind::Codec(crate::opened::CodecLimitKind::TextBytes),
                observed: usize::MAX,
                maximum: text_maximum,
            })
    })?;
    if text_bytes > text_maximum {
        return Err(Error::Limit {
            kind: crate::opened::LimitKind::Codec(crate::opened::CodecLimitKind::TextBytes),
            observed: text_bytes,
            maximum: text_maximum,
        });
    }
    let Some(context) = context else {
        return Ok(None);
    };
    let maximum = managed_memory_limit(context);
    // Three simultaneously live scene indexes, replacement plans, escaped
    // text, and intermediate/output buffers fit inside this conservative
    // envelope. Geometry stages release their prior candidate before the next
    // shape. The reservation follows the final immutable payload.
    let output_bound = text_bytes
        .checked_mul(6)
        .and_then(|size| size.checked_add(input_bytes))
        .and_then(|size| size.checked_add(replacements.len().checked_mul(512)?))
        .ok_or(Error::Limit {
            kind: crate::opened::LimitKind::Codec(crate::opened::CodecLimitKind::OutputBytes),
            observed: usize::MAX,
            maximum: max_output_bytes,
        })?
        .min(max_output_bytes);
    let amount = input_bytes
        .checked_mul(64)
        .and_then(|size| size.checked_add(output_bound.checked_mul(196)?))
        .and_then(|size| size.checked_add(text_bytes.checked_mul(6)?))
        .and_then(|size| size.checked_add(4096))
        .ok_or(Error::Limit {
            kind: crate::opened::LimitKind::Codec(crate::opened::CodecLimitKind::MemoryBytes),
            observed: usize::MAX,
            maximum,
        })?;
    if amount > maximum {
        return Err(Error::Limit {
            kind: crate::opened::LimitKind::Codec(crate::opened::CodecLimitKind::MemoryBytes),
            observed: amount,
            maximum,
        });
    }
    reserve_source_memory(context, amount, maximum)
}

fn reserve_source_shape_text_chunks(
    context: Option<&ExecutionContext>,
    input_bytes: usize,
    admission: crate::opened::ChunkAdmission,
    max_output_bytes: usize,
) -> Result<Option<Arc<litchi_core::Reservation>>> {
    let replacements = admission.selections();
    let text_bytes = admission.text_bytes();
    let Some(context) = context else {
        return Ok(None);
    };
    let maximum = managed_memory_limit(context);
    // Chunk callers perform output admission before this reservation. The
    // scalar output bound remains separate from the managed memory ceiling.
    let output_bound = input_bytes
        .checked_add(text_bytes.checked_mul(6).ok_or(Error::Limit {
            kind: crate::opened::LimitKind::Codec(crate::opened::CodecLimitKind::OutputBytes),
            observed: usize::MAX,
            maximum: max_output_bytes,
        })?)
        .and_then(|size| size.checked_add(replacements.checked_mul(256)?))
        .ok_or(Error::Limit {
            kind: crate::opened::LimitKind::Codec(crate::opened::CodecLimitKind::OutputBytes),
            observed: usize::MAX,
            maximum: max_output_bytes,
        })?
        .min(max_output_bytes);
    let input_scene = input_bytes
        .checked_mul(64)
        .and_then(|size| size.checked_add(1024))
        .ok_or(Error::Limit {
            kind: crate::opened::LimitKind::Codec(crate::opened::CodecLimitKind::MemoryBytes),
            observed: usize::MAX,
            maximum,
        })?;
    let output_scene = output_bound
        .checked_mul(64)
        .and_then(|size| size.checked_add(1024))
        .ok_or(Error::Limit {
            kind: crate::opened::LimitKind::Codec(crate::opened::CodecLimitKind::MemoryBytes),
            observed: usize::MAX,
            maximum,
        })?;
    let amount = input_scene
        .checked_add(output_scene)
        .and_then(|size| size.checked_add(output_bound))
        .and_then(|size| size.checked_add(text_bytes.checked_mul(6)?))
        .and_then(|size| size.checked_add(replacements.checked_mul(256)?))
        .and_then(|size| size.checked_add(4096))
        .ok_or(Error::Limit {
            kind: crate::opened::LimitKind::Codec(crate::opened::CodecLimitKind::MemoryBytes),
            observed: usize::MAX,
            maximum,
        })?;
    if amount > maximum {
        return Err(Error::Limit {
            kind: crate::opened::LimitKind::Codec(crate::opened::CodecLimitKind::MemoryBytes),
            observed: amount,
            maximum,
        });
    }
    reserve_source_memory(context, amount, maximum)
}

fn reserve_source_shape_property_chunks(
    context: Option<&ExecutionContext>,
    input_bytes: usize,
    admission: crate::opened::ChunkAdmission,
    max_output_bytes: usize,
) -> Result<Option<Arc<litchi_core::Reservation>>> {
    let replacements = admission.selections();
    let text_bytes = admission.text_bytes();
    let Some(context) = context else {
        return Ok(None);
    };
    let maximum = managed_memory_limit(context);
    let output_bound = input_bytes
        .checked_add(text_bytes.checked_mul(6).ok_or(Error::Limit {
            kind: crate::opened::LimitKind::Codec(crate::opened::CodecLimitKind::OutputBytes),
            observed: usize::MAX,
            maximum: max_output_bytes,
        })?)
        .and_then(|size| size.checked_add(replacements.checked_mul(512)?))
        .ok_or(Error::Limit {
            kind: crate::opened::LimitKind::Codec(crate::opened::CodecLimitKind::OutputBytes),
            observed: usize::MAX,
            maximum: max_output_bytes,
        })?
        .min(max_output_bytes);
    let amount = input_bytes
        .checked_mul(64)
        .and_then(|size| size.checked_add(output_bound.checked_mul(196)?))
        .and_then(|size| size.checked_add(text_bytes.checked_mul(6)?))
        .and_then(|size| size.checked_add(replacements.checked_mul(512)?))
        .and_then(|size| size.checked_add(4096))
        .ok_or(Error::Limit {
            kind: crate::opened::LimitKind::Codec(crate::opened::CodecLimitKind::MemoryBytes),
            observed: usize::MAX,
            maximum,
        })?;
    if amount > maximum {
        return Err(Error::Limit {
            kind: crate::opened::LimitKind::Codec(crate::opened::CodecLimitKind::MemoryBytes),
            observed: amount,
            maximum,
        });
    }
    reserve_source_memory(context, amount, maximum)
}

#[cfg(any(unix, windows))]
fn file_source(path: impl AsRef<Path>) -> Result<Arc<dyn ReadAt>> {
    Ok(Arc::new(FileSource::open(path)?))
}

/// Admit a deferred OPC package from `source` under `options`.
fn source_backed_package(
    source: Arc<dyn ReadAt>,
    options: crate::OpenOptions,
) -> Result<SourceBackedPackage> {
    let crate::OpenOptions {
        source_limits,
        cache_limits,
        execution_context,
        ..
    } = options;
    Ok(match (cache_limits, execution_context) {
        (None, None) => SourceBackedPackage::from_read_at(
            source,
            source_limits,
            SourceCacheLimits::default(),
            litchi_opc::SourceReadPolicy::exact(),
            None,
        )?,
        (Some(cache), None) => SourceBackedPackage::from_read_at(
            source,
            source_limits,
            cache,
            litchi_opc::SourceReadPolicy::exact(),
            None,
        )?,
        (None, Some(context)) => SourceBackedPackage::from_read_at(
            source,
            source_limits,
            SourceCacheLimits::default(),
            litchi_opc::SourceReadPolicy::exact(),
            Some(context),
        )?,
        (Some(cache), Some(context)) => SourceBackedPackage::from_read_at(
            source,
            source_limits,
            cache,
            litchi_opc::SourceReadPolicy::exact(),
            Some(context),
        )?,
    })
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

// Risk R3 of unified-ops record 0003: the unified `Pptx` handle shares these
// owners across threads, so they must stay `Send + Sync`.
const _: () = {
    const fn assert_send_sync<T: Send + Sync>() {}
    assert_send_sync::<SourceBackedPresentation>();
    assert_send_sync::<SourceSlide>();
    assert_send_sync::<SourceBackedPresentationEditor>();
    assert_send_sync::<SourceBackedIntoEditorError>();
};

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

/// Metadata for the optional DrawingML SVG resource paired with one raster
/// picture blip.  The raster target remains the compatibility image while the
/// SVG relationship is the semantic vector resource described by
/// `[MS-ODRAWXML]` section 2.26.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct SourceSvgDescriptor {
    relationship_id: String,
    target: SourceImageTarget,
    blip: SvgBlip,
}

impl SourceSvgDescriptor {
    /// Relationship ID carried by `asvg:svgBlip@r:embed` or `r:link`.
    #[must_use]
    pub fn relationship_id(&self) -> &str {
        &self.relationship_id
    }

    /// Validated embedded or inert external SVG target metadata.
    #[must_use]
    pub const fn target(&self) -> &SourceImageTarget {
        &self.target
    }

    /// Typed and lossless SVG blip metadata from the picture extension.
    pub const fn blip(&self) -> &SvgBlip {
        &self.blip
    }

    /// Whether this SVG descriptor points at an external, unreadable target.
    #[must_use]
    pub const fn is_external(&self) -> bool {
        self.target.is_external()
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
    svg: Option<Arc<SourceSvgDescriptor>>,
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

    /// Optional typed SVG resource paired with this raster compatibility
    /// image.  The SVG payload remains deferred until
    /// [`SourceSlide::read_svg_image`] is called.
    #[must_use]
    pub fn svg(&self) -> Option<&SourceSvgDescriptor> {
        self.svg.as_deref()
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

/// One selected embedded SVG resource and its bounded source-backed payload
/// handle.
#[derive(Debug, Clone)]
pub struct SourceSvgImage {
    descriptor: Arc<SourceSvgDescriptor>,
    data: PartData,
}

impl SourceSvgImage {
    /// Metadata descriptor used to select this SVG resource.
    #[must_use]
    pub fn descriptor(&self) -> &SourceSvgDescriptor {
        &self.descriptor
    }

    /// Borrow the exact embedded SVG payload bytes.
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

#[derive(Clone, Copy)]
enum PictureQueryMode {
    All,
    Selected(usize),
}

enum PictureQueryResult {
    All(Vec<SourceImageDescriptor>),
    Selected {
        descriptor: Option<SourceImageDescriptor>,
        len: usize,
    },
}

/// An owning, source-backed editor for one existing slide.
///
/// Unlike [`SourceBackedPresentation`], this type is intentionally not
/// cloneable: publishing consumes its deferred OPC source to ensure that the
/// exact source checked during editing is the source raw-copied to output.
/// Focused slide edits retain their existing package topology. The SVG picture
/// transaction is the bounded exception: it may replace the two media Parts
/// and their owning slide relationship targets as one checked closure.
#[derive(Clone)]
pub(crate) struct SourceBackedPresentationEditor {
    pub(super) package: SourceBackedPackage,
    // Retain the catalog validated at open so each selected slide does not
    // reparse the immutable presentation root and rebuild the complete slide
    // graph. Snapshot capture still checks execution, the selected Part, and
    // the exact source version before publishing anything.
    pub(super) _presentation: SourcePart,
    pub(super) slides: Box<[Arc<SourceSlideData>]>,
    package_relationships: Arc<[RelationshipBinding]>,
    presentation_binding: Arc<PartBinding>,
    pub(super) limits: SourceLimits,
    /// The ingress-selected semantic policies must stay attached to the
    /// deferred owner.  Source limits alone cannot recover content-family
    /// ceilings when a snapshot is projected or replayed.
    pub(super) semantic_limits: Option<crate::open_options::SemanticLimitsRef>,
}

/// An immutable exact-source snapshot of one slide XML part.
#[derive(Clone)]
pub(crate) struct SourceBackedSlideSnapshot {
    position: usize,
    part_uri: PackURI,
    xml: SourcePayload,
    closure: SlideClosure,
    max_output_bytes: usize,
    protection_limits: crate::presentation_properties::metadata::protection::Limits,
    /// Preserve the complete ingress semantic bundle with the exact source
    /// image.  A snapshot may outlive its presentation projection.
    semantic_limits: Option<crate::open_options::SemanticLimitsRef>,
    source_version: SourceVersion,
    lineage: SourceLineage,
    context: Option<ExecutionContext>,
    /// The source-selected laser policy. Usage is parsed only when a laser
    /// aggregate or replay boundary needs it, not for unrelated snapshots.
    laser_limits: crate::laser::Limits,
    /// Lazily cached laser usage for this exact XML under `laser_limits`.
    laser_usage: Arc<OnceLock<crate::laser::Usage>>,
    /// Markup compatibility enclosed in shape-tree children admits only
    /// whole top-level shape removal (record 0303).
    mce: crate::presentation::transition::SlideMce,
}

/// An isolated focused edit of one exact-source slide snapshot.
pub(crate) struct SourceBackedSlideEdit {
    source: SourceBackedSlideSnapshot,
    working: SourcePayload,
    operation_used: bool,
    context: Option<ExecutionContext>,
    /// The source policy is restored when the edit is frozen.  The focused
    /// editor may temporarily lower its output bound to a caller's exact
    /// retained-patch budget without changing patch identity.
    source_max_output_bytes: usize,
}

/// A reversible, exact-source-checked one-slide XML patch.
#[derive(Clone)]
pub(crate) struct SourceBackedSlidePatch {
    before: SourceBackedSlideSnapshot,
    after: SourceBackedSlideSnapshot,
}

/// A checked source-backed slide edit ready for publication.
pub(crate) struct SourceBackedSlideCommit {
    snapshot: SourceBackedSlideSnapshot,
    #[cfg_attr(
        not(test),
        expect(dead_code, reason = "read by the source-level engine tests")
    )]
    patch: SourceBackedSlidePatch,
}

/// Replacement payloads for one embedded PPTX SVG picture.
///
/// The raster payload is the compatibility fallback required by the native
/// DrawingML SVG extension.  Callers may optionally select new existing or
/// newly-added `/ppt/media/` Part URIs; the transaction keeps the relationship
/// IDs stable and updates the slide relationship closure atomically.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct SvgReplacement {
    svg: Vec<u8>,
    raster: Vec<u8>,
    svg_part_uri: Option<PackURI>,
    raster_part_uri: Option<PackURI>,
}

impl SvgReplacement {
    /// Create a replacement that retains both current media Part URIs.
    #[must_use]
    pub fn new(svg: impl Into<Vec<u8>>, raster: impl Into<Vec<u8>>) -> Self {
        Self {
            svg: svg.into(),
            raster: raster.into(),
            svg_part_uri: None,
            raster_part_uri: None,
        }
    }

    /// Select the internal SVG and raster Part URIs for this replacement.
    ///
    /// New names are added to the package when absent. Existing names must
    /// already be media Parts with compatible content types. Both names are
    /// validated against the source package before any output is written.
    #[must_use]
    pub fn with_part_uris(mut self, svg: PackURI, raster: PackURI) -> Self {
        self.svg_part_uri = Some(svg);
        self.raster_part_uri = Some(raster);
        self
    }

    /// SVG bytes supplied for the replacement.
    #[must_use]
    pub fn svg(&self) -> &[u8] {
        &self.svg
    }

    /// Raster fallback bytes supplied for the replacement.
    #[must_use]
    pub fn raster(&self) -> &[u8] {
        &self.raster
    }
}

#[derive(Clone)]
pub struct SvgImage {
    slide: SourceBackedSlideSnapshot,
    image_position: usize,
    svg_relationship_id: String,
    svg_relationship_type: String,
    raster_relationship_id: String,
    raster_relationship_type: String,
    svg_part_uri: PackURI,
    raster_part_uri: PackURI,
    svg: SourcePayload,
    raster: SourcePayload,
    limits: SourceLimits,
}

/// An exact-source transaction over one embedded PPTX SVG picture and its
/// raster compatibility fallback.
pub(crate) struct SourceBackedSvgEdit {
    source: SvgImage,
    svg_part_uri: PackURI,
    raster_part_uri: PackURI,
    svg: SourcePayload,
    raster: SourcePayload,
    operation_used: bool,
}

/// A reversible exact-source patch for one PPTX SVG picture closure.
#[derive(Clone)]
pub(crate) struct SourceBackedSvgPatch {
    before: SvgImage,
    after: SvgImage,
}

/// A checked source-backed SVG picture transaction ready for publication.
pub(crate) struct SourceBackedSvgCommit {
    snapshot: SvgImage,
    patch: SourceBackedSvgPatch,
}

/// A bounded multi-slide shape-text edit borrowing its deferred source.
#[cfg(test)]
pub(crate) struct SourceBackedSlideBatchEdit<'a> {
    editor: &'a SourceBackedPresentationEditor,
    commits: Vec<SourceBackedSlideCommit>,
    text_bytes: usize,
}

/// An immutable exact-source snapshot of a selected slide set.
#[derive(Clone)]
#[cfg(test)]
pub(crate) struct SourceBackedSlideBatchSnapshot {
    slides: Box<[SourceBackedSlideSnapshot]>,
}

/// A reversible, exact-source-checked multi-slide XML patch.
#[derive(Clone)]
#[cfg(test)]
pub(crate) struct SourceBackedSlideBatchPatch {
    before: SourceBackedSlideBatchSnapshot,
    after: SourceBackedSlideBatchSnapshot,
}

/// A checked source-backed multi-slide edit ready for publication.
#[cfg(test)]
pub(crate) struct SourceBackedSlideBatchCommit {
    snapshot: SourceBackedSlideBatchSnapshot,
    patch: SourceBackedSlideBatchPatch,
}

#[cfg(test)]
enum FastApply<T> {
    Applied(T),
    Recapture,
}

#[cfg(test)]
thread_local! {
    static TEST_SEMANTIC_SCENE_READS: Cell<usize> = const { Cell::new(0) };
    static TEST_SEMANTIC_SLIDE_PART_READS: Cell<usize> = const { Cell::new(0) };
}

#[cfg(test)]
fn reset_test_semantic_scene_reads() {
    TEST_SEMANTIC_SCENE_READS.with(|count| count.set(0));
    TEST_SEMANTIC_SLIDE_PART_READS.with(|count| count.set(0));
}

#[cfg(test)]
fn test_semantic_scene_reads() -> usize {
    TEST_SEMANTIC_SCENE_READS.with(Cell::get)
}

#[cfg(test)]
fn test_semantic_slide_part_reads() -> usize {
    TEST_SEMANTIC_SLIDE_PART_READS.with(Cell::get)
}

#[cfg(test)]
fn record_test_semantic_scene_read() {
    TEST_SEMANTIC_SCENE_READS.with(|count| count.set(count.get().saturating_add(1)));
}

#[cfg(test)]
fn record_test_semantic_slide_part_read() {
    TEST_SEMANTIC_SLIDE_PART_READS.with(|count| count.set(count.get().saturating_add(1)));
}

impl SourceBackedPresentation {
    /// Open a positional PPTX file under the default finite policies.
    ///
    /// # Errors
    ///
    /// Returns an error if the input cannot be read or is malformed.
    #[cfg(any(unix, windows))]
    pub fn open(path: impl AsRef<Path>) -> Result<Self> {
        Self::open_with_options(path, crate::OpenOptions::default())
    }

    /// Open a positional PPTX file under explicit ingress options: source
    /// limits, a payload-cache policy and an execution context.
    ///
    /// # Errors
    ///
    /// Returns an error if the input exceeds the policies, cannot be read or
    /// is malformed.
    #[cfg(any(unix, windows))]
    pub fn open_with_options(path: impl AsRef<Path>, options: crate::OpenOptions) -> Result<Self> {
        Self::from_read_at_with_options(file_source(path)?, options)
    }

    /// Admit a positional source under the default finite policies.
    ///
    /// # Errors
    ///
    /// Returns an error if the input cannot be read or is malformed.
    pub fn from_read_at(source: Arc<dyn ReadAt>) -> Result<Self> {
        Self::from_read_at_with_options(source, crate::OpenOptions::default())
    }

    /// Admit a positional source under explicit ingress options.
    ///
    /// # Errors
    ///
    /// Returns an error if the input exceeds the policies, cannot be read or
    /// is malformed.
    pub fn from_read_at_with_options(
        source: Arc<dyn ReadAt>,
        options: crate::OpenOptions,
    ) -> Result<Self> {
        let limits = options.source_limits;
        let semantic_limits = options.semantic_limits.clone();
        Self::from_source_backed_package_with_limits(
            source_backed_package(source, options)?,
            limits,
            semantic_limits,
        )
    }

    /// Read a PPTX stream into a retained source under the default finite
    /// policies.
    ///
    /// # Errors
    ///
    /// Returns an error if the input cannot be read or is malformed.
    pub fn from_reader<R: Read>(reader: R) -> Result<Self> {
        Self::from_reader_with_options(reader, crate::OpenOptions::default())
    }

    /// Read a PPTX stream into a retained source under explicit source
    /// limits.
    ///
    /// # Errors
    ///
    /// Returns [`crate::Error::UnsupportedOpenPolicy`] for cache or execution
    /// policies, which need a positional source, and an error if the input
    /// exceeds the source limits, cannot be read or is malformed.
    pub fn from_reader_with_options<R: Read>(
        reader: R,
        options: crate::OpenOptions,
    ) -> Result<Self> {
        options.check_owning()?;
        let limits = options.source_limits;
        let semantic_limits = options.semantic_limits.clone();
        Self::from_source_backed_package_with_limits(
            SourceBackedPackage::from_reader(reader, limits, SourceCacheLimits::default(), None)?,
            limits,
            semantic_limits,
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
        let limits = package.read_limits();
        Self::from_source_backed_package_with_limits(package, limits, None)
    }

    fn from_source_backed_package_with_limits(
        package: SourceBackedPackage,
        limits: SourceLimits,
        semantic_limits: Option<crate::open_options::SemanticLimitsRef>,
    ) -> Result<Self> {
        package.check_execution()?;
        let mut tolerance = crate::leniency::ToleranceReport::default();
        // OPC manifest repairs (audit findings 8.5-8.9) were applied while
        // the deferred package was opened; they come first.
        tolerance.append_package(package.tolerance_report());
        let repairs_before_catalog = tolerance.total();
        let catalog = source_catalog(&package, Some(&mut tolerance))?;
        package.check_execution()?;
        let strict_slide_list = tolerance.total() == repairs_before_catalog;

        Ok(Self {
            inner: Arc::new(SourceInner {
                scenes: scenes::SceneCache::for_package(&package),
                package,
                _presentation: catalog.presentation,
                slides: catalog.slides,
                limits,
                semantic_limits,
                tolerance,
                strict_slide_list,
            }),
        })
    }

    /// Borrow the validated deferred OPC package this presentation reads from.
    ///
    /// This is the package indexed when the presentation was opened: no ZIP
    /// directory, content-type catalog, or relationship graph is rebuilt.
    /// Borrowed package services, such as
    /// [`SourceBackedPackage::source_artifact`], read through the same
    /// positional source and observe the same source-version fence.
    #[must_use]
    pub fn source_backed_package(&self) -> &SourceBackedPackage {
        &self.inner.package
    }

    /// Convert this read-only presentation into a
    /// [`SourceBackedPresentationEditor`] over the same, already indexed
    /// source-backed package.
    ///
    /// The ZIP central directory, OPC catalog, presentation root, and slide
    /// graph validated at open are reused: no source byte is re-read for
    /// indexing and no part payload is loaded. Only the package and
    /// presentation relationship bindings the editor's publication planner
    /// needs are copied from the retained in-memory catalog. When the slide
    /// list needed a lenient repair at open, the editor's strict slide-list
    /// contract is applied again over the retained root, exactly as
    /// [`SourceBackedPresentationEditor::from_read_at`] would apply it, and
    /// its refusal is reported.
    ///
    /// The editor takes ownership of the package because publication consumes
    /// it, so the conversion succeeds only while this value is the sole owner
    /// of the opened source: no clone of it and no [`SourceSlide`] from it is
    /// alive.
    ///
    /// # Errors
    ///
    /// Returns the presentation unchanged inside [`SourceBackedIntoEditorError`]:
    /// - without an error when it is shared ([`SourceBackedIntoEditorError::is_shared`]);
    /// - with the typed cancellation, source-change, allocation, or strict
    ///   slide-list error that prevented the conversion.
    #[cfg(test)]
    pub(crate) fn try_into_editor(
        self,
    ) -> std::result::Result<SourceBackedPresentationEditor, SourceBackedIntoEditorError> {
        let seed = match self.editor_seed() {
            Ok(seed) => seed,
            Err(error) => {
                return Err(SourceBackedIntoEditorError {
                    presentation: self,
                    error: Some(error),
                });
            },
        };
        let inner = match Arc::try_unwrap(self.inner) {
            Ok(inner) => inner,
            Err(inner) => {
                return Err(SourceBackedIntoEditorError {
                    presentation: Self { inner },
                    error: None,
                });
            },
        };
        let SourceInner {
            scenes: _,
            package,
            _presentation,
            slides,
            limits,
            semantic_limits,
            tolerance: _,
            strict_slide_list: _,
        } = inner;
        let (presentation, slides) = match seed.strict_catalog {
            Some((presentation, slides)) => (presentation, slides),
            None => (_presentation, slides),
        };
        Ok(SourceBackedPresentationEditor {
            package,
            _presentation: presentation,
            slides,
            package_relationships: seed.package_relationships,
            presentation_binding: seed.presentation_binding,
            limits,
            semantic_limits,
        })
    }

    /// Everything fallible [`Self::try_into_editor`] needs, computed while
    /// the presentation is still borrowed so a failure can return it intact.
    fn editor_seed(&self) -> Result<EditorSeed> {
        let package = &self.inner.package;
        package.check_execution()?;
        package.source_version()?;
        if !self.inner.strict_slide_list {
            // The lenient open repaired the slide list; the editor refuses
            // such a list. Rebuild the strict catalog over the same indexed
            // package so the refusal (or, defensively, the strict graph) is
            // exactly the editor's own.
            let catalog = source_catalog(package, None)?;
            package.check_execution()?;
            package.source_version()?;
            return Ok(EditorSeed {
                package_relationships: Arc::from(catalog.package_relationships),
                presentation_binding: Arc::new(catalog.presentation_binding),
                strict_catalog: Some((catalog.presentation, catalog.slides)),
            });
        }
        let package_relationships = Arc::from(relationship_bindings(package, package.rels())?);
        let view = package.main_document_part()?;
        let presentation_binding = Arc::new(part_binding(package, &view)?);
        package.check_execution()?;
        package.source_version()?;
        Ok(EditorSeed {
            package_relationships,
            presentation_binding,
            strict_catalog: None,
        })
    }

    /// Every non-structural defect the reader repaired while opening this
    /// presentation, in discovery order.
    ///
    /// An empty report means the presentation conformed; see
    /// [`crate::ToleranceReport`]. Slide payloads stay deferred until a
    /// selected [`SourceSlide`] is read, so repairs a deferred read performs
    /// are reported by that read's own entry point rather than here; see
    /// [`crate::leniency`].
    #[must_use]
    pub fn tolerance_report(&self) -> &crate::leniency::ToleranceReport {
        &self.inner.tolerance
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
    /// any slide or media payload, appending every core-properties repair the
    /// read applied to `tolerance` as [`crate::Defect::Properties`].
    ///
    /// A missing core-properties relationship returns `None`; a present but
    /// empty properties part returns `Some(Props::new())`. The source version
    /// and caller execution policy are checked before and after the selected
    /// metadata payload read.
    ///
    /// Opening a source-backed presentation reads only the presentation
    /// root, so the core-properties repairs are not part of
    /// [`Self::tolerance_report`]; this is where a caller collects them.
    /// There is deliberately no report-less variant: repaired properties
    /// cannot be obtained without their report. See [`crate::leniency`].
    ///
    /// # Errors
    ///
    /// Returns a typed OPC cancellation or source-change error, or the core
    /// properties reader's error for a part it rejects.
    pub fn properties_with_tolerance(
        &self,
        tolerance: &mut crate::leniency::ToleranceReport,
    ) -> Result<Option<litchi_core::Metadata>> {
        self.inner.package.check_execution()?;
        let mut repairs = litchi_ooxml_common::leniency::ToleranceReport::default();
        let properties = litchi_ooxml_common::properties::read_source_backed_with_report(
            &self.inner.package,
            &mut repairs,
        )?;
        self.inner.package.check_execution()?;
        self.inner.package.source_version()?;
        tolerance.append_properties(&repairs);
        Ok(properties)
    }

    /// Presentation slide size in EMUs.
    ///
    /// The mandatory presentation root was loaded and retained while the
    /// source catalog was opened; this method parses only that pinned root.
    pub fn slide_size(&self) -> Result<(i64, i64)> {
        related::with_fences(&self.inner, || {
            let root = &self.inner._presentation;
            let part = related::load_part(
                &self.inner,
                root.partname(),
                root.content_type(),
                crate::notes::MAX_PRESENTATION_XML,
            )?;
            PresentationPart::from_part(&part)?.slide_size()
        })
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

    /// Return a fail-closed, content-free source-cache diagnostic snapshot.
    ///
    /// This is the instrumentation and telemetry entry point for callers that
    /// must not mistake a poisoned cache-state mutex or checked counter
    /// overflow for valid data. Like [`Self::cache_diagnostics`], observing the
    /// snapshot does not load a part. The infallible compatibility method above
    /// remains available for ordinary callers that accept its recovery
    /// behavior.
    pub fn try_cache_diagnostics(
        &self,
    ) -> std::result::Result<
        litchi_opc::SourceCacheDiagnostics,
        litchi_opc::SourceCacheDiagnosticsError,
    > {
        self.inner.package.try_cache_diagnostics()
    }
}

/// The editor-only catalog state [`SourceBackedPresentation::try_into_editor`]
/// derives from the retained catalog before taking ownership of the package.
struct EditorSeed {
    package_relationships: Arc<[RelationshipBinding]>,
    presentation_binding: Arc<PartBinding>,
    /// The strict presentation root and slide graph, present only when the
    /// open-time slide list needed a lenient repair.
    strict_catalog: Option<(SourcePart, Box<[Arc<SourceSlideData>]>)>,
}

/// Why [`SourceBackedPresentation::try_into_editor`] could not convert a
/// presentation. The presentation is returned unchanged and stays readable.
pub(crate) struct SourceBackedIntoEditorError {
    presentation: SourceBackedPresentation,
    error: Option<Error>,
}

impl SourceBackedIntoEditorError {
    /// Whether the conversion was refused only because another clone of the
    /// presentation, or a [`SourceSlide`] from it, still shares the source.
    #[must_use]
    #[cfg(test)]
    pub(crate) fn is_shared(&self) -> bool {
        self.error.is_none()
    }

    /// The typed error that prevented the conversion, or `None` when the
    /// presentation was merely shared.
    #[must_use]
    #[cfg(test)]
    pub(crate) fn error(&self) -> Option<&Error> {
        self.error.as_ref()
    }

    /// Split into the unchanged presentation and the optional typed error.
    #[must_use]
    #[cfg(test)]
    pub(crate) fn into_parts(self) -> (SourceBackedPresentation, Option<Error>) {
        (self.presentation, self.error)
    }
}

impl std::fmt::Debug for SourceBackedIntoEditorError {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        formatter
            .debug_struct("SourceBackedIntoEditorError")
            .field("slide_count", &self.presentation.slide_count())
            .field("error", &self.error)
            .finish()
    }
}

impl std::fmt::Display for SourceBackedIntoEditorError {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match &self.error {
            Some(error) => write!(
                formatter,
                "cannot convert the source-backed presentation into an editor: {error}"
            ),
            None => formatter
                .write_str("cannot convert a shared source-backed presentation into an editor"),
        }
    }
}

impl std::error::Error for SourceBackedIntoEditorError {
    fn source(&self) -> Option<&(dyn std::error::Error + 'static)> {
        let error: &(dyn std::error::Error + 'static) = self.error.as_ref()?;
        Some(error)
    }
}

impl SourceBackedPresentationEditor {
    /// Open a positional PPTX file under the default finite policies.
    ///
    /// # Errors
    ///
    /// Returns an error if the input cannot be read or is malformed.
    #[cfg(any(unix, windows))]
    #[cfg(test)]
    pub(crate) fn open(path: impl AsRef<Path>) -> Result<Self> {
        Self::open_with_options(path, crate::OpenOptions::default())
    }

    /// Open a positional PPTX file under explicit ingress options: source
    /// limits, a payload-cache policy and an execution context.
    ///
    /// # Errors
    ///
    /// Returns an error if the input exceeds the policies, cannot be read or
    /// is malformed.
    #[cfg(any(unix, windows))]
    #[cfg(test)]
    pub(crate) fn open_with_options(
        path: impl AsRef<Path>,
        options: crate::OpenOptions,
    ) -> Result<Self> {
        Self::from_read_at_with_options(file_source(path)?, options)
    }

    /// Admit a positional source under the default finite policies.
    ///
    /// # Errors
    ///
    /// Returns an error if the input cannot be read or is malformed.
    #[cfg(test)]
    pub(crate) fn from_read_at(source: Arc<dyn ReadAt>) -> Result<Self> {
        Self::from_read_at_with_options(source, crate::OpenOptions::default())
    }

    /// Admit a positional source under explicit ingress options.
    ///
    /// # Errors
    ///
    /// Returns an error if the input exceeds the policies, cannot be read or
    /// is malformed.
    #[cfg(test)]
    pub(crate) fn from_read_at_with_options(
        source: Arc<dyn ReadAt>,
        options: crate::OpenOptions,
    ) -> Result<Self> {
        let limits = options.source_limits;
        let semantic_limits = options.semantic_limits.clone();
        Self::from_source_backed_package_with_semantic_limits(
            source_backed_package(source, options)?,
            limits,
            semantic_limits,
        )
    }

    /// Build an owning editor from a validated deferred OPC package.
    #[cfg(test)]
    fn from_source_backed_package_with_semantic_limits(
        package: SourceBackedPackage,
        limits: SourceLimits,
        semantic_limits: Option<crate::open_options::SemanticLimitsRef>,
    ) -> Result<Self> {
        package.check_execution()?;
        // The lossless editor keeps the strict slide-list contract.
        let catalog = source_catalog(&package, None)?;
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
            semantic_limits,
        })
    }

    /// Number of logical slides in presentation order.
    #[must_use]
    #[cfg(test)]
    pub(crate) fn slide_count(&self) -> usize {
        self.slides.len()
    }

    /// Return content-free payload-cache activity for this deferred editor.
    ///
    /// Opening loads the presentation root; capturing one slide loads that
    /// selected slide only. The diagnostic exposes no member identities.
    #[must_use]
    #[cfg(test)]
    pub(crate) fn cache_diagnostics(&self) -> litchi_opc::SourceCacheDiagnostics {
        self.package.cache_diagnostics()
    }

    /// Return a fail-closed, content-free source-cache diagnostic snapshot.
    ///
    /// This is the instrumentation and telemetry entry point for callers that
    /// must reject a poisoned cache-state mutex or checked counter overflow
    /// instead of treating the recovered infallible snapshot as valid. Like
    /// [`Self::cache_diagnostics`], observing the snapshot does not load a
    /// part. The infallible method remains available for compatibility
    /// callers.
    #[cfg(test)]
    pub(crate) fn try_cache_diagnostics(
        &self,
    ) -> std::result::Result<
        litchi_opc::SourceCacheDiagnostics,
        litchi_opc::SourceCacheDiagnosticsError,
    > {
        self.package.try_cache_diagnostics()
    }

    /// Capture the exact raw XML of one existing slide.
    ///
    /// A direct transition's supported MCE owner is excluded while validating
    /// semantic shape offsets; unrelated markup-compatibility branch
    /// selection remains refused.
    pub(crate) fn slide_snapshot(&self, position: usize) -> Result<SourceBackedSlideSnapshot> {
        self.package.check_execution()?;
        self.slide_snapshot_for(position, "slide_snapshot")
    }

    /// Begin an isolated focused edit of one existing slide.
    #[cfg(test)]
    pub(crate) fn edit_slide(&self, position: usize) -> Result<SourceBackedSlideEdit> {
        self.package.check_execution()?;
        self.slide_snapshot_for(position, "edit_slide")?
            .edit_checked()
    }

    /// Begin an exact-source transaction over one direct embedded SVG picture.
    ///
    /// The selected picture must have an internal SVG relationship and an
    /// internal raster fallback. Linked SVGs, markup-compatibility picture
    /// branches, and ambiguous dependency closures are refused before an edit
    /// handle is returned. The slide XML remains source-backed; only the media
    /// Parts and their owning relationship targets belong to the transaction.
    #[cfg(test)]
    pub(crate) fn edit_svg_image(
        &self,
        slide_position: usize,
        image_position: usize,
    ) -> Result<SourceBackedSvgEdit> {
        self.svg_image_edit(slide_position, image_position)
    }

    /// The SVG picture transaction behind [`Self::edit_svg_image`].
    pub(crate) fn svg_image_edit(
        &self,
        slide_position: usize,
        image_position: usize,
    ) -> Result<SourceBackedSvgEdit> {
        self.package.check_execution()?;
        self.svg_snapshot_for(slide_position, image_position, "edit_svg_image")?
            .edit_checked()
    }

    /// Begin an atomic, bounded shape-text edit across existing slides.
    ///
    /// Each selected slide may receive one nonempty same-slide shape batch.
    /// The edit supports at most [`MAX_SOURCE_BACKED_SLIDE_BATCH`] distinct
    /// slides and the ordinary opened-presentation aggregate text budget.
    #[must_use]
    #[cfg(test)]
    pub(crate) fn edit_slides(&self) -> SourceBackedSlideBatchEdit<'_> {
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
    #[cfg(test)]
    pub(crate) fn publish_slide_commit_to_stream<W: Write>(
        self,
        writer: W,
        commit: &SourceBackedSlideCommit,
    ) -> Result<SourceBackedSlideSnapshot> {
        self.package.check_execution()?;
        let current = self
            .publication_raw_slide_snapshot_from_retained_catalog(commit.patch.source().position)?;
        let publication_uri = current.part_uri.clone();
        let target = match commit.patch.try_apply_fast(&current)? {
            FastApply::Applied(target) => {
                drop(current);
                target
            },
            FastApply::Recapture => {
                drop(current);
                let current = self.slide_snapshot_for(
                    commit.patch.source().position,
                    "publish_slide_commit_to_stream",
                )?;
                commit.patch.apply(&current)?
            },
        };
        self.package.check_execution()?;
        crate::source_limits::validate_xml_limits(target.xml.as_bytes(), self.limits)?;
        if let Some(replacement) = target.xml.edited_bytes() {
            self.package.write_part_overlay_shared_to_stream(
                writer,
                &publication_uri,
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

    /// Publish one exact-source-checked SVG picture transaction.
    ///
    /// The selected slide XML and every unrelated ZIP member are raw-copied.
    /// The owning slide relationship member, SVG Part, and raster fallback Part
    /// are changed together through one bounded OPC topology plan. New media
    /// Parts are added only when explicitly selected by the replacement; old
    /// unreferenced media Parts are removed as part of the same dependency
    /// closure. No-op commits use the exact source publication path.
    pub(crate) fn publish_svg_commit_to_stream<W: Write>(
        self,
        writer: W,
        commit: &SourceBackedSvgCommit,
    ) -> Result<SvgImage> {
        self.package.check_execution()?;
        let current = self.svg_snapshot_for(
            commit.patch.before.slide.position,
            commit.patch.before.image_position,
            "publish_svg_commit_to_stream",
        )?;
        if !current.same_source(&commit.patch.before) {
            return Err(Error::PatchConflict);
        }
        let target = commit.patch.apply(&current)?;
        if target.same_source(&current) {
            self.package
                .write_topology_to_stream(writer, SourceTopologyPlan::new())?;
            return Ok(target);
        }
        let plan = build_svg_topology_plan(&self.package, &current, &target)?;
        self.package.write_topology_to_stream(writer, plan)?;
        Ok(target)
    }

    /// Publish one exact-source-checked multi-slide commit to a stream.
    ///
    /// Only changed selected slide XML members are regenerated. Selected
    /// no-ops and every unselected member retain their raw ZIP records. An
    /// all-no-op commit reproduces the complete source artifact byte for byte.
    #[cfg(test)]
    pub(crate) fn publish_slide_batch_commit_to_stream<W: Write>(
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
            match self.publication_raw_slide_snapshot_from_retained_catalog(source.position) {
                Ok(snapshot) => slides.push(snapshot),
                Err(error) if slides.is_empty() => return Err(error),
                Err(error) => {
                    let raw_read_count = slides.len();
                    let original_error = error;
                    for snapshot in slides.drain(..).take(raw_read_count) {
                        self.validate_retained_raw_slide_snapshot(
                            snapshot,
                            "publish_slide_batch_commit_to_stream",
                        )?;
                    }
                    return Err(original_error);
                },
            }
        }
        let current = SourceBackedSlideBatchSnapshot {
            slides: slides.into_boxed_slice(),
        };
        let target = match commit.patch.try_apply_fast(&current)? {
            FastApply::Applied(target) => {
                drop(current);
                target
            },
            FastApply::Recapture => {
                drop(current);
                self.recapture_and_apply_slide_batch(&commit.patch)?
            },
        };
        self.package.check_execution()?;
        for slide in target.slides() {
            crate::source_limits::validate_xml_limits(slide.xml.as_bytes(), self.limits)?;
        }
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

    #[cfg(test)]
    fn recapture_and_apply_slide_batch(
        &self,
        patch: &SourceBackedSlideBatchPatch,
    ) -> Result<SourceBackedSlideBatchSnapshot> {
        let mut slides = Vec::new();
        slides
            .try_reserve_exact(patch.source().slide_count())
            .map_err(|source| Error::Allocation {
                resource: "source-backed slide batch publication snapshot",
                source,
            })?;
        for source in patch.source().slides() {
            slides.push(self.slide_snapshot_from_retained_catalog(
                source.position,
                "publish_slide_batch_commit_to_stream",
            )?);
        }
        let current = SourceBackedSlideBatchSnapshot {
            slides: slides.into_boxed_slice(),
        };
        patch.apply(&current)
    }

    #[cfg(test)]
    fn validate_retained_raw_slide_snapshot(
        &self,
        snapshot: SourceBackedSlideSnapshot,
        operation: &'static str,
    ) -> Result<()> {
        self.package.check_execution()?;
        let SourceBackedSlideSnapshot { closure, xml, .. } = snapshot;
        let SourcePayload::Original(raw) = xml else {
            return Err(Error::internal(
                operation,
                "publication raw snapshot must retain original source bytes",
            ));
        };
        let part = SourcePart::from_binding(&closure.slide.slide, raw)?;
        #[cfg(test)]
        record_test_semantic_slide_part_read();
        SlidePart::from_part(&part)?;
        #[cfg(test)]
        record_test_semantic_scene_read();
        let semantic = crate::open_options::selected(self.semantic_limits.as_ref());
        crate::presentation::transition::admit_source_with_limits(
            part.source_data().as_bytes(),
            operation,
            semantic.shape(),
            semantic.transition(),
        )?;
        self.package.check_execution()?;
        let _source_version = self.package.source_version()?;
        Ok(())
    }

    fn publication_raw_slide_snapshot_from_retained_catalog(
        &self,
        position: usize,
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
        let source_version = match self.package.source_version() {
            Ok(source_version) => source_version,
            Err(error) => {
                #[cfg(test)]
                record_test_semantic_scene_read();
                {
                    let _scene = crate::shape::Scene::read(
                        raw.as_bytes(),
                        crate::open_options::selected(self.semantic_limits.as_ref()).shape(),
                    )?;
                }
                let part = SourcePart::from_view(&view, raw)?;
                #[cfg(test)]
                record_test_semantic_slide_part_read();
                SlidePart::from_part(&part)?;
                self.package.check_execution()?;
                return Err(error.into());
            },
        };
        let semantic = crate::open_options::selected(self.semantic_limits.as_ref());
        let laser_limits = semantic.laser().constrained_by_source(self.limits);
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
            max_output_bytes: source_slide_output_limit(
                self.limits,
                crate::open_options::selected(self.semantic_limits.as_ref()).shape(),
            ),
            protection_limits:
                crate::presentation_properties::metadata::protection::Limits::from_source_limits(
                    self.limits,
                ),
            semantic_limits: self.semantic_limits.clone(),
            source_version,
            lineage,
            context,
            laser_limits,
            laser_usage: Arc::new(OnceLock::new()),
            // Read and publication snapshots are never edited.
            mce: crate::presentation::transition::SlideMce::Plain,
        })
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

    fn svg_snapshot_for(
        &self,
        slide_position: usize,
        image_position: usize,
        operation: &'static str,
    ) -> Result<SvgImage> {
        let slide = self.slide_snapshot_for(slide_position, operation)?;
        self.capture_svg_snapshot(slide, image_position, operation)
    }

    fn capture_svg_snapshot(
        &self,
        slide: SourceBackedSlideSnapshot,
        image_position: usize,
        operation: &'static str,
    ) -> Result<SvgImage> {
        self.package.check_execution()?;
        if slide.mce != crate::presentation::transition::SlideMce::Plain {
            return Err(crate::presentation::transition::enclosed_edit_refusal(
                operation,
            ));
        }
        let view = self.package.part(&slide.part_uri)?;
        let source_part = SourcePart::from_view(&view, view.data()?)?;
        validate_source_slide_root(&source_part)?;
        validate_full_slide_picture_relationships(&self.package, source_part.blob())?;
        let scene = crate::shape::Scene::read(
            source_part.blob(),
            crate::open_options::selected(self.semantic_limits.as_ref()).shape(),
        )?;
        let mut picture_xmls = Vec::new();
        picture_xmls
            .try_reserve(scene.len())
            .map_err(|source| Error::Allocation {
                resource: "source-backed picture XML fragments",
                source,
            })?;
        for shape in scene.iter() {
            if let Shape::Picture(picture) = shape {
                picture_xmls.push(picture.common().xml()?);
            }
        }
        let relationships = svg_lifecycle::parse_picture_relationships_with_source(
            source_part.blob(),
            picture_xmls.iter().copied(),
            self.limits.max_part_bytes(),
        )?;
        let mut relationships = relationships.into_iter();
        let mut picture_count = 0usize;
        let mut selected = None;
        for shape in scene.iter() {
            let Shape::Picture(_picture) = shape else {
                continue;
            };
            let position = picture_count;
            picture_count = picture_count
                .checked_add(1)
                .ok_or_else(|| Error::Invalid("source-backed picture count overflow".into()))?;
            let relationship = relationships.next().ok_or_else(|| {
                Error::Invalid("source-backed picture relationship inventory is truncated".into())
            })?;
            let raster_target = resolve_picture_target(&self.package, &view, &relationship)?;
            let svg_target = relationship
                .svg
                .as_ref()
                .map(|svg| resolve_svg_target(&self.package, &view, svg))
                .transpose()?;
            if position == image_position {
                selected = Some((relationship, raster_target, svg_target));
            }
        }
        let (picture, raster_target, svg_target) = selected.ok_or(Error::IndexOutOfBounds {
            index: image_position,
            len: picture_count,
        })?;
        let svg = picture.svg.as_ref().ok_or_else(|| {
            Error::Relationship(format!(
                "source-backed image {image_position} has no SVG relationship"
            ))
        })?;
        let svg_target = svg_target.ok_or_else(|| {
            Error::Relationship(format!(
                "source-backed image {image_position} has no SVG relationship"
            ))
        })?;
        let SourceImageTarget::Internal {
            part_uri: raster_part_uri,
            content_type: raster_content_type,
        } = raster_target
        else {
            return Err(Error::Relationship(
                "source-backed SVG transaction requires an internal raster fallback".into(),
            ));
        };
        let SourceImageTarget::Internal {
            part_uri: svg_part_uri,
            content_type: svg_content_type,
        } = svg_target.target.clone()
        else {
            return Err(Error::Relationship(
                "source-backed SVG transaction refuses linked SVG targets".into(),
            ));
        };
        if raster_part_uri == svg_part_uri {
            return Err(Error::Relationship(
                "source-backed SVG transaction requires distinct raster and SVG Parts".into(),
            ));
        }
        let raster_view = self.package.part(&raster_part_uri)?;
        let svg_view = self.package.part(&svg_part_uri)?;
        if !raster_view.rels().is_empty() || !svg_view.rels().is_empty() {
            return Err(Error::Relationship(
                "source-backed SVG media Parts must not have outbound relationships".into(),
            ));
        }
        if !is_svg_content_type(&svg_content_type) {
            return Err(Error::ContentType {
                expected: "image/svg+xml".into(),
                actual: svg_content_type,
            });
        }
        if !is_png_content_type(&raster_content_type) {
            return Err(Error::ContentType {
                expected: "image/png".into(),
                actual: raster_content_type,
            });
        }
        let raster_data = raster_view.data()?;
        let svg_data = svg_view.data()?;
        let raster_relationship = view.rels().get(&picture.id).ok_or_else(|| {
            Error::Relationship(format!(
                "picture raster relationship '{}' is missing",
                picture.id
            ))
        })?;
        let svg_relationship = view.rels().get(&svg.id).ok_or_else(|| {
            Error::Relationship(format!("picture SVG relationship '{}' is missing", svg.id))
        })?;
        if raster_relationship.target_mode() != TargetMode::Internal
            || svg_relationship.target_mode() != TargetMode::Internal
        {
            return Err(Error::Relationship(
                "source-backed SVG transaction requires internal media relationships".into(),
            ));
        }
        self.package.check_execution()?;
        self.package.source_version()?;
        Ok(SvgImage {
            slide,
            image_position,
            svg_relationship_id: svg.id.clone(),
            svg_relationship_type: svg_relationship.reltype().to_owned(),
            raster_relationship_id: picture.id,
            raster_relationship_type: raster_relationship.reltype().to_owned(),
            svg_part_uri,
            raster_part_uri,
            svg: SourcePayload::Original(svg_data),
            raster: SourcePayload::Original(raster_data),
            limits: self.limits,
        })
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
        #[cfg(test)]
        record_test_semantic_scene_read();
        crate::source_limits::validate_xml_limits(raw.as_bytes(), self.limits)?;
        let semantic = crate::open_options::selected(self.semantic_limits.as_ref());
        let mce = crate::presentation::transition::admit_source_with_limits(
            raw.as_bytes(),
            operation,
            semantic.shape(),
            semantic.transition(),
        )?;
        let part = SourcePart::from_view(&view, raw)?;
        #[cfg(test)]
        record_test_semantic_slide_part_read();
        if let Err(error) = SlidePart::from_part(&part) {
            // The SVG owner scanner preserves legal opaque extension markup
            // that the general SlidePart/MCE projection deliberately refuses
            // to interpret. Other malformed slide errors remain fatal.
            if !matches!(error, Error::MarkupCompatibility(_)) {
                return Err(error);
            }
        }
        let SourcePart { data: raw, .. } = part;
        self.package.check_execution()?;
        let source_version = self.package.source_version()?;
        let laser_limits = semantic.laser().constrained_by_source(self.limits);
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
            max_output_bytes: source_slide_output_limit(
                self.limits,
                crate::open_options::selected(self.semantic_limits.as_ref()).shape(),
            ),
            protection_limits:
                crate::presentation_properties::metadata::protection::Limits::from_source_limits(
                    self.limits,
                ),
            semantic_limits: self.semantic_limits.clone(),
            source_version,
            lineage,
            context,
            laser_limits,
            laser_usage: Arc::new(OnceLock::new()),
            mce,
        })
    }
}

impl SourceBackedSlideSnapshot {
    /// Checked zero-based slide position in the presentation catalog.
    #[must_use]
    pub(crate) const fn position(&self) -> usize {
        self.position
    }

    /// Read this slide's direct typed transition, if present.
    ///
    /// Repairs the read applied (for example a wheel spoke count kept as a
    /// raw effect) are listed, with this slide's part name, in
    /// [`crate::transition::Transition::tolerance_report`].
    ///
    /// Inherited layout/master transitions are intentionally outside this
    /// direct-slide capability. Supported PowerPoint extension transitions are
    /// read through their typed MCE choice and fallback without selecting any
    /// unrelated slide markup.
    pub(crate) fn transition(&self) -> Result<Option<crate::transition::Transition>> {
        self.check_execution()?;
        let mut transition = crate::presentation::transition::read_direct_with_limits(
            self.xml.as_bytes(),
            crate::open_options::selected(self.semantic_limits.as_ref()).transition(),
        )?;
        if let Some(transition) = transition.as_mut() {
            transition.locate_in_part(self.part_uri.as_str());
        }
        self.check_execution()?;
        Ok(transition)
    }

    /// Start an isolated focused edit from this exact source snapshot.
    #[must_use]
    pub(crate) fn edit(&self) -> SourceBackedSlideEdit {
        SourceBackedSlideEdit {
            source: self.clone(),
            working: self.xml.clone(),
            operation_used: false,
            context: self.context.clone(),
            source_max_output_bytes: self.max_output_bytes,
        }
    }

    /// Check the retained caller execution policy before creating an edit
    /// handle. The infallible [`Self::edit`] constructor remains available for
    /// compatibility because it performs no payload I/O.
    #[cfg(test)]
    pub(crate) fn edit_checked(&self) -> Result<SourceBackedSlideEdit> {
        self.check_execution()?;
        Ok(self.edit())
    }

    /// Start a checked focused edit under a caller-selected output bound.
    ///
    /// The bound controls every candidate XML builder and reservation during
    /// staging.  The immutable source identity keeps its original physical
    /// output policy when the edit is committed, so changing the transient
    /// admission ceiling cannot manufacture a patch conflict.
    pub(crate) fn edit_checked_with_max_patch_bytes(
        &self,
        maximum: usize,
    ) -> Result<SourceBackedSlideEdit> {
        self.check_execution()?;
        let mut edit = self.edit();
        edit.source.max_output_bytes = edit.source.max_output_bytes.min(maximum);
        edit.source_max_output_bytes = self.max_output_bytes;
        Ok(edit)
    }

    fn check_execution(&self) -> Result<()> {
        check_execution_context(self.context.as_ref())
    }

    #[cfg(test)]
    fn laser_usage(&self) -> Result<crate::laser::Usage> {
        if let Some(usage) = self.laser_usage.get() {
            return Ok(*usage);
        }
        let usage = capture_laser_usage(self.xml.as_bytes(), self.laser_limits)?;
        let _ = self.laser_usage.set(usage);
        Ok(usage)
    }

    fn same_source(&self, other: &Self) -> bool {
        self.position == other.position
            && self.part_uri == other.part_uri
            && self.xml.as_bytes() == other.xml.as_bytes()
            && self.closure.package_relationships == other.closure.package_relationships
            && self.closure.presentation == other.closure.presentation
            && self.closure.presentation_xml.as_bytes() == other.closure.presentation_xml.as_bytes()
            && self.closure.slide == other.closure.slide
    }

    /// Rebase an edited source image onto the recipient's retained policy.
    /// Equal source bytes are the patch scope; a profile difference is a
    /// destination admission question, never a source conflict.
    fn rebase_policy(&self, recipient: &Self) -> Result<Self> {
        let mut usage = crate::laser::Usage::default();
        self.rebase_policy_with_usage(recipient, &mut usage)
    }

    fn rebase_policy_with_usage(
        &self,
        recipient: &Self,
        usage: &mut crate::laser::Usage,
    ) -> Result<Self> {
        let mut target = self.clone();
        // A replayed candidate always belongs to the recipient's execution
        // scope, even when its semantic and output profiles compare equal.
        // A managed reservation belongs to the source edit's budget until the
        // recipient admits the replay. Rehome an equivalent token below so
        // the published candidate is charged to the recipient instead.
        target.context = recipient.context.clone();
        target.max_output_bytes = recipient.max_output_bytes;
        target.semantic_limits = recipient.semantic_limits.clone();
        let same_laser_policy = target.laser_limits == recipient.laser_limits;
        target.laser_limits = recipient.laser_limits;
        if !same_laser_policy {
            target.laser_usage = Arc::new(OnceLock::new());
        }
        target.xml = rebase_payload_reservation(target.xml, recipient.context.as_ref())?;
        if target.xml.as_bytes().len() > recipient.max_output_bytes {
            return Err(Error::Limit {
                kind: crate::opened::LimitKind::Codec(crate::opened::CodecLimitKind::OutputBytes),
                observed: target.xml.as_bytes().len(),
                maximum: recipient.max_output_bytes,
            });
        }
        let semantic = crate::open_options::selected(recipient.semantic_limits.as_ref());
        target.mce = crate::presentation::transition::admit_source_with_limits(
            target.xml.as_bytes(),
            "apply",
            semantic.shape(),
            semantic.transition(),
        )?;
        let slide_usage = if same_laser_policy {
            match self.laser_usage.get().copied() {
                Some(usage) => usage,
                None => capture_laser_usage(target.xml.as_bytes(), recipient.laser_limits)?,
            }
        } else {
            capture_laser_usage(target.xml.as_bytes(), recipient.laser_limits)?
        };
        merge_laser_usage(usage, slide_usage, recipient.laser_limits)?;
        if same_laser_policy && self.laser_usage.get().is_none() {
            let _ = self.laser_usage.set(slide_usage);
        }
        if !same_laser_policy {
            let _ = target.laser_usage.set(slide_usage);
        }
        Ok(target)
    }
}

impl SourceBackedSlideEdit {
    /// Exact immutable slide snapshot against which this edit was created.
    #[must_use]
    #[cfg(test)]
    pub(crate) const fn source(&self) -> &SourceBackedSlideSnapshot {
        &self.source
    }

    /// Replace the authored slide name and/or visibility attributes in one
    /// source-backed slide operation.
    ///
    /// `name` targets `p:cSld@name`; `hidden` targets the inverse of
    /// `p:sld@show`. Omitted fields are preserved byte-for-byte. The shared
    /// codec retains all unrelated root, shape, and opaque XML, and an equal
    /// value is an exact no-op.
    pub(crate) fn set_slide_properties(
        &mut self,
        name: Option<&str>,
        hidden: Option<bool>,
    ) -> Result<bool> {
        self.check_execution()?;
        self.refuse_enclosed("set_slide_properties")?;
        if name.is_none() && hidden.is_none() {
            return Ok(false);
        }
        if self.operation_used {
            return Err(Error::refused(
                litchi_core::edit::RefusalKind::UnsupportedOperation,
                "set_slide_properties",
                "source-backed slide edits support one atomic semantic operation",
            ));
        }
        let reservation = reserve_source_slide_properties(
            self.context.as_ref(),
            self.working.as_bytes().len(),
            name,
            self.source.max_output_bytes,
        )?;
        self.check_execution()?;
        let xml = crate::opened::stage_slide_properties(
            self.working.as_bytes(),
            name,
            hidden,
            self.source.max_output_bytes,
        )?;
        let changed = xml.is_some();
        self.check_execution()?;
        if let Some(xml) = xml {
            let bytes = Arc::new(xml);
            self.working = match reservation {
                Some(reservation) => SourcePayload::ReservedEdit { bytes, reservation },
                None => SourcePayload::Edited(bytes),
            };
        }
        self.operation_used = true;
        Ok(changed)
    }

    /// Set or replace this existing slide's direct typed transition.
    ///
    /// This operation never changes relationships or package topology. The
    /// supplied value may use the supported PowerPoint extension vocabulary
    /// and custom duration. Unknown transition effects are refused. An equal
    /// modeled value is an exact byte no-op.
    pub(crate) fn set_transition(&mut self, value: &crate::transition::Transition) -> Result<bool> {
        self.stage_transition(Some(value), "set_transition")
    }

    /// Clear this existing slide's direct transition.
    ///
    /// An absent direct transition is an exact byte no-op. Inherited
    /// layout/master state is not removed or rewritten.
    pub(crate) fn clear_transition(&mut self) -> Result<bool> {
        self.stage_transition(None, "clear_transition")
    }

    fn stage_transition(
        &mut self,
        value: Option<&crate::transition::Transition>,
        operation: &'static str,
    ) -> Result<bool> {
        self.check_execution()?;
        if self.operation_used {
            return Err(Error::refused(
                litchi_core::edit::RefusalKind::UnsupportedOperation,
                operation,
                "source-backed slide edits support one atomic semantic operation",
            ));
        }
        let (xml, changed) = crate::presentation::transition::stage_with_limits(
            self.working.as_bytes(),
            self.source.closure.presentation_xml.as_bytes(),
            value,
            self.source.max_output_bytes,
            crate::open_options::selected(self.source.semantic_limits.as_ref()).transition(),
            operation,
            self.source.protection_limits,
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
    #[cfg(test)]
    pub(crate) fn set_shape_text<'k>(
        &mut self,
        shape: impl Into<crate::shape::Key<'k>>,
        text: impl AsRef<str>,
    ) -> Result<bool> {
        let text = text.as_ref();
        let replacement = crate::opened::ShapeTextReplacement::new(shape.into(), text);
        Ok(self.set_shape_texts(std::slice::from_ref(&replacement))? != 0)
    }

    /// Remove one direct top-level shape from this exact source slide.
    ///
    /// A whole top-level group may be removed as one opaque root; selecting a
    /// grouped descendant is refused. Connector endpoints and animation
    /// targets that still point at any removed identity are also refused.
    /// Source-backed removal keeps the slide relationship table and any now
    /// unreachable Parts intact so the exact source topology remains
    /// reversible; the eager package editor may prune that closure separately.
    pub(crate) fn remove_shape<'k>(
        &mut self,
        shape: impl Into<crate::shape::Key<'k>>,
    ) -> Result<bool> {
        self.check_execution()?;
        if self.operation_used {
            return Err(Error::refused(
                litchi_core::edit::RefusalKind::UnsupportedOperation,
                "remove_shape",
                "source-backed slide edits support one atomic semantic operation",
            ));
        }
        let reservation = reserve_source_shape_removal(
            self.context.as_ref(),
            self.working.as_bytes().len(),
            self.source.max_output_bytes,
        )?;
        self.check_execution()?;
        // A slide that encloses markup compatibility in shape children is
        // spliced on its source bytes (record 0303).
        let (xml, changed) = if self.source.mce == crate::presentation::transition::SlideMce::Plain
        {
            crate::opened::stage_shape_removal_with_limits(
                self.working.as_bytes(),
                shape.into(),
                self.source.max_output_bytes,
                crate::open_options::selected(self.source.semantic_limits.as_ref()).shape(),
            )?
        } else {
            crate::opened::stage_enclosed_shape_removal_with_limits(
                self.working.as_bytes(),
                shape.into(),
                self.source.max_output_bytes,
                crate::open_options::selected(self.source.semantic_limits.as_ref()).shape(),
            )?
        };
        let Some(xml) = xml else {
            return Ok(false);
        };
        self.check_execution()?;
        let bytes = Arc::new(xml);
        self.working = match reservation {
            Some(reservation) => SourcePayload::ReservedEdit { bytes, reservation },
            None => SourcePayload::Edited(bytes),
        };
        self.operation_used = true;
        Ok(changed != 0)
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
    pub(crate) fn set_shape_texts(
        &mut self,
        replacements: &[crate::opened::ShapeTextReplacement<'_>],
    ) -> Result<usize> {
        self.check_execution()?;
        self.refuse_enclosed("set_shape_texts")?;
        if replacements.is_empty() {
            return Ok(0);
        }
        if self.operation_used {
            return Err(Error::refused(
                litchi_core::edit::RefusalKind::UnsupportedOperation,
                "set_shape_texts",
                "source-backed slide edits support one atomic shape text operation",
            ));
        }
        let (xml, changed) = crate::opened::stage_shape_texts_with_limits(
            self.working.as_bytes(),
            replacements,
            crate::opened::Limits::default().max_text_bytes(),
            self.source.max_output_bytes,
            crate::open_options::selected(self.source.semantic_limits.as_ref()).shape(),
        )?;
        if let Some(xml) = xml {
            self.working = SourcePayload::Edited(Arc::new(xml));
        }
        self.operation_used = true;
        self.check_execution()?;
        Ok(changed)
    }

    /// Atomically replace text from several already bounded same-slide
    /// batches in one scene read, rewrite, and read-back.
    ///
    /// Every inner slice retains the ordinary 256-selector and aggregate text
    /// limits of [`Self::set_shape_texts`]. The chunks must identify a
    /// globally disjoint set of shapes; the native codec validates that
    /// contract before publishing the single candidate payload. The number of
    /// chunks and the aggregate selection count are bounded by the native
    /// scene shape limit.
    /// Aggregate replacement text is bounded by the slide output-byte limit.
    pub(crate) fn set_shape_text_chunks(
        &mut self,
        chunks: &[&[crate::opened::ShapeTextReplacement<'_>]],
    ) -> Result<usize> {
        self.check_execution()?;
        self.refuse_enclosed("set_shape_text_chunks")?;
        let shape_limits =
            crate::open_options::selected(self.source.semantic_limits.as_ref()).shape();
        if chunks.len() > shape_limits.shapes() {
            return Err(Error::Limit {
                kind: crate::opened::LimitKind::Codec(
                    crate::opened::CodecLimitKind::TextReplacements,
                ),
                observed: chunks.len(),
                maximum: shape_limits.shapes(),
            });
        }
        if chunks.iter().all(|chunk| chunk.is_empty()) {
            return Ok(0);
        }
        if self.operation_used {
            return Err(Error::refused(
                litchi_core::edit::RefusalKind::UnsupportedOperation,
                "set_shape_text_chunks",
                "source-backed slide edits support one atomic shape text operation",
            ));
        }
        let admission = crate::opened::admit_shape_text_chunks(
            chunks,
            crate::opened::Limits::default().max_text_bytes(),
            self.source.max_output_bytes,
        )?;
        debug_assert!(admission.selections() != 0);
        let reservation = reserve_source_shape_text_chunks(
            self.context.as_ref(),
            self.working.as_bytes().len(),
            admission,
            self.source.max_output_bytes,
        )?;
        let (xml, changed) = crate::opened::stage_shape_text_chunks_admitted_with_limits(
            self.working.as_bytes(),
            chunks,
            crate::opened::Limits::default().max_text_bytes(),
            self.source.max_output_bytes,
            admission,
            shape_limits,
        )?;
        self.check_execution()?;
        if let Some(xml) = xml {
            let bytes = Arc::new(xml);
            self.working = match reservation {
                Some(reservation) => SourcePayload::ReservedEdit { bytes, reservation },
                None => SourcePayload::Edited(bytes),
            };
        }
        self.operation_used = true;
        Ok(changed)
    }

    /// Atomically update text and bounds on selected shapes of this slide.
    ///
    /// Only existing direct DrawingML transforms are edited. The operation
    /// retains opaque XML and applies all requested properties together.
    /// An empty batch does not consume this focused edit's operation allowance.
    pub(crate) fn set_shape_properties(
        &mut self,
        replacements: &[crate::opened::ShapePropertiesReplacement<'_>],
    ) -> Result<usize> {
        self.check_execution()?;
        self.refuse_enclosed("set_shape_properties")?;
        if replacements.is_empty() {
            return Ok(0);
        }
        if self.operation_used {
            return Err(Error::refused(
                litchi_core::edit::RefusalKind::UnsupportedOperation,
                "set_shape_properties",
                "source-backed slide edits support one atomic semantic operation",
            ));
        }
        let reservation = reserve_source_shape_properties(
            self.context.as_ref(),
            self.working.as_bytes().len(),
            replacements,
            self.source.max_output_bytes,
        )?;
        let (xml, changed) = crate::opened::stage_shape_properties_with_limits(
            self.working.as_bytes(),
            replacements,
            crate::opened::Limits::default().max_text_bytes(),
            self.source.max_output_bytes,
            crate::open_options::selected(self.source.semantic_limits.as_ref()).shape(),
        )?;
        self.check_execution()?;
        if let Some(xml) = xml {
            let bytes = Arc::new(xml);
            self.working = match reservation {
                Some(reservation) => SourcePayload::ReservedEdit { bytes, reservation },
                None => SourcePayload::Edited(bytes),
            };
        }
        self.operation_used = true;
        Ok(changed)
    }

    /// Atomically replace text and/or bounds from several already bounded
    /// same-slide property batches in one scene read, rewrite, and read-back.
    /// Every inner slice retains the ordinary per-call property limits.
    /// Aggregate replacement text is bounded by the slide output-byte limit.
    pub(crate) fn set_shape_property_chunks(
        &mut self,
        chunks: &[&[crate::opened::ShapePropertiesReplacement<'_>]],
    ) -> Result<usize> {
        self.check_execution()?;
        self.refuse_enclosed("set_shape_property_chunks")?;
        let shape_limits =
            crate::open_options::selected(self.source.semantic_limits.as_ref()).shape();
        if chunks.len() > shape_limits.shapes() {
            return Err(Error::Limit {
                kind: crate::opened::LimitKind::Codec(
                    crate::opened::CodecLimitKind::ShapeProperties,
                ),
                observed: chunks.len(),
                maximum: shape_limits.shapes(),
            });
        }
        if chunks.iter().all(|chunk| chunk.is_empty()) {
            return Ok(0);
        }
        if self.operation_used {
            return Err(Error::refused(
                litchi_core::edit::RefusalKind::UnsupportedOperation,
                "set_shape_property_chunks",
                "source-backed slide edits support one atomic semantic operation",
            ));
        }
        let admission = crate::opened::admit_shape_property_chunks(
            chunks,
            crate::opened::Limits::default().max_text_bytes(),
            self.source.max_output_bytes,
        )?;
        debug_assert!(admission.selections() != 0);
        let reservation = reserve_source_shape_property_chunks(
            self.context.as_ref(),
            self.working.as_bytes().len(),
            admission,
            self.source.max_output_bytes,
        )?;
        let (xml, changed) = crate::opened::stage_shape_property_chunks_admitted_with_limits(
            self.working.as_bytes(),
            chunks,
            crate::opened::Limits::default().max_text_bytes(),
            self.source.max_output_bytes,
            admission,
            shape_limits,
        )?;
        self.check_execution()?;
        if let Some(xml) = xml {
            let bytes = Arc::new(xml);
            self.working = match reservation {
                Some(reservation) => SourcePayload::ReservedEdit { bytes, reservation },
                None => SourcePayload::Edited(bytes),
            };
        }
        self.operation_used = true;
        Ok(changed)
    }

    /// Apply a bounded batch of rich text edits to this exact source slide.
    ///
    /// See [`crate::opened::Edit::set_text_properties`]. Only the selected
    /// slide payload changes, so the source-backed proof path still applies.
    pub(crate) fn set_text_properties(
        &mut self,
        edits: &[crate::opened::TextAssignment<'_>],
    ) -> Result<usize> {
        self.check_execution()?;
        self.refuse_enclosed("set_text_properties")?;
        if edits.is_empty() {
            return Ok(0);
        }
        if self.operation_used {
            return Err(Error::refused(
                litchi_core::edit::RefusalKind::UnsupportedOperation,
                "set_text_properties",
                "source-backed slide edits support one atomic semantic operation",
            ));
        }
        let (xml, changed) = crate::opened::stage_text_properties_with_limits(
            self.working.as_bytes(),
            edits,
            crate::opened::Limits::default().max_text_bytes(),
            self.source.max_output_bytes,
            crate::open_options::selected(self.source.semantic_limits.as_ref()).shape(),
        )?;
        self.check_execution()?;
        if let Some(xml) = xml {
            self.working = SourcePayload::Edited(Arc::new(xml));
        }
        self.operation_used = true;
        Ok(changed)
    }

    /// Insert or remove one paragraph or run of a text body of this exact
    /// source slide (record 0333); see
    /// [`crate::opened::Edit::insert_paragraph`]. Only the selected slide
    /// payload changes. Notes-page bodies live in another part and are
    /// refused.
    pub(crate) fn change_text_body(
        &mut self,
        body: crate::opened::TextBody,
        change: &crate::opened::TextChange<'_>,
    ) -> Result<usize> {
        self.check_execution()?;
        self.refuse_enclosed("text_structure")?;
        if matches!(body, crate::opened::TextBody::Notes { .. }) {
            return Err(Error::refused(
                litchi_core::edit::RefusalKind::UnsupportedOperation,
                "text_structure",
                "source-backed slide edits do not stage notes pages",
            ));
        }
        if self.operation_used {
            return Err(Error::refused(
                litchi_core::edit::RefusalKind::UnsupportedOperation,
                "text_structure",
                "source-backed slide edits support one atomic semantic operation",
            ));
        }
        let xml = crate::opened::stage_text_structure_with_limits(
            self.working.as_bytes(),
            body,
            change,
            crate::opened::Limits::default().max_text_bytes(),
            self.source.max_output_bytes,
            crate::open_options::selected(self.source.semantic_limits.as_ref()).shape(),
        )?;
        self.check_execution()?;
        self.working = SourcePayload::Edited(Arc::new(xml));
        self.operation_used = true;
        Ok(1)
    }

    /// Apply several bounded batches of rich text edits to this exact source
    /// slide in one pass (record 0333); see
    /// [`crate::opened::Edit::set_text_property_chunks`].
    pub(crate) fn set_text_property_chunks(
        &mut self,
        chunks: &[&[crate::opened::TextAssignment<'_>]],
    ) -> Result<usize> {
        self.text_property_chunks(chunks, None)
    }

    pub(crate) fn text_property_chunks(
        &mut self,
        chunks: &[&[crate::opened::TextAssignment<'_>]],
        mut changes: Option<&mut Vec<bool>>,
    ) -> Result<usize> {
        self.check_execution()?;
        self.refuse_enclosed("set_text_properties")?;
        if chunks.iter().all(|chunk| chunk.is_empty()) {
            if let Some(changes) = changes.as_deref_mut() {
                changes.clear();
            }
            return Ok(0);
        }
        if self.operation_used {
            return Err(Error::refused(
                litchi_core::edit::RefusalKind::UnsupportedOperation,
                "set_text_properties",
                "source-backed slide edits support one atomic semantic operation",
            ));
        }
        let (xml, changed) = crate::opened::stage_text_property_chunks_with_limits(
            self.working.as_bytes(),
            chunks,
            crate::opened::Limits::default().max_text_bytes(),
            self.source.max_output_bytes,
            changes,
            crate::open_options::selected(self.source.semantic_limits.as_ref()).shape(),
        )?;
        self.check_execution()?;
        if let Some(xml) = xml {
            self.working = SourcePayload::Edited(Arc::new(xml));
        }
        self.operation_used = true;
        Ok(changed)
    }

    /// Add one plain DrawingML text box to this exact source slide.
    ///
    /// The operation is staged against the retained slide bytes and then
    /// semantically read back before the edit is published. It does not build
    /// an eager package or rewrite any related Part, so the surrounding source
    /// XML remains eligible for the source-backed durable proof path.
    pub(crate) fn add_text_box(
        &mut self,
        text: impl AsRef<str>,
        bounds: (i64, i64, i64, i64),
    ) -> Result<u32> {
        self.check_execution()?;
        self.refuse_enclosed("add_text_box")?;
        if self.operation_used {
            return Err(Error::refused(
                litchi_core::edit::RefusalKind::UnsupportedOperation,
                "add_text_box",
                "source-backed slide edits support one atomic semantic operation",
            ));
        }
        let text = text.as_ref();
        crate::opened::validate_text_box_input(text, bounds)?;
        let reservation = reserve_source_text_box(
            self.context.as_ref(),
            self.working.as_bytes().len(),
            text.len(),
            self.source.max_output_bytes,
        )?;
        self.check_execution()?;
        let (xml, shape_id) = crate::opened::stage_text_box(
            self.working.as_bytes(),
            text,
            bounds,
            self.source.max_output_bytes,
            crate::open_options::selected(self.source.semantic_limits.as_ref()).shape(),
        )?;
        self.check_execution()?;
        let bytes = Arc::new(xml);
        self.working = match reservation {
            Some(reservation) => SourcePayload::ReservedEdit { bytes, reservation },
            None => SourcePayload::Edited(bytes),
        };
        self.operation_used = true;
        Ok(shape_id)
    }

    fn check_execution(&self) -> Result<()> {
        check_execution_context(self.context.as_ref())
    }

    /// Markup compatibility enclosed in shape children admits only whole
    /// top-level shape removal; every other edit keeps the guard
    /// (record 0303).
    fn refuse_enclosed(&self, operation: &'static str) -> Result<()> {
        if self.source.mce == crate::presentation::transition::SlideMce::Plain {
            Ok(())
        } else {
            Err(crate::presentation::transition::enclosed_edit_refusal(
                operation,
            ))
        }
    }

    fn into_commit(self) -> SourceBackedSlideCommit {
        let mut source = self.source;
        source.max_output_bytes = self.source_max_output_bytes;
        let working = self.working;
        let unchanged = working.as_bytes() == source.xml.as_bytes();
        let snapshot = SourceBackedSlideSnapshot {
            position: source.position,
            part_uri: source.part_uri.clone(),
            xml: working,
            closure: source.closure.clone(),
            max_output_bytes: self.source_max_output_bytes,
            protection_limits: source.protection_limits,
            semantic_limits: source.semantic_limits.clone(),
            source_version: source.source_version,
            lineage: source.lineage.clone(),
            context: source.context.clone(),
            laser_limits: source.laser_limits,
            laser_usage: if unchanged {
                Arc::clone(&source.laser_usage)
            } else {
                Arc::new(OnceLock::new())
            },
            mce: source.mce,
        };
        let patch = SourceBackedSlidePatch {
            before: source,
            after: snapshot.clone(),
        };
        SourceBackedSlideCommit { snapshot, patch }
    }

    /// Validate and freeze this isolated edit for source-backed publication.
    #[must_use]
    #[cfg(test)]
    pub(crate) fn commit(self) -> SourceBackedSlideCommit {
        self.into_commit()
    }

    /// Check cancellation and freeze this isolated edit for publication.
    pub(crate) fn commit_checked(self) -> Result<SourceBackedSlideCommit> {
        self.check_execution()?;
        let commit = self.into_commit();
        commit.snapshot.check_execution()?;
        Ok(commit)
    }
}

impl SourceBackedSlidePatch {
    /// Exact immutable source required by this patch.
    #[must_use]
    #[cfg(test)]
    pub(crate) const fn source(&self) -> &SourceBackedSlideSnapshot {
        &self.before
    }

    fn is_changed(&self) -> bool {
        !self.before.same_source(&self.after)
    }

    /// Whether this patch leaves its source unchanged.
    #[must_use]
    #[cfg(test)]
    pub(crate) fn is_noop(&self) -> bool {
        !self.is_changed()
    }

    /// Return the exact inverse slide patch.
    #[must_use]
    pub(crate) fn inverse(&self) -> Self {
        Self {
            before: self.after.clone(),
            after: self.before.clone(),
        }
    }

    #[cfg(test)]
    fn try_apply_fast(
        &self,
        source: &SourceBackedSlideSnapshot,
    ) -> Result<FastApply<SourceBackedSlideSnapshot>> {
        if !source.same_source(&self.before) {
            return Ok(FastApply::Recapture);
        }
        self.apply(source).map(FastApply::Applied)
    }

    /// Apply only to the exact raw slide source captured by this patch.
    pub(crate) fn apply(
        &self,
        source: &SourceBackedSlideSnapshot,
    ) -> Result<SourceBackedSlideSnapshot> {
        self.before.check_execution()?;
        source.check_execution()?;
        if !source.same_source(&self.before) {
            return Err(Error::PatchConflict);
        }
        Ok(if self.is_changed() {
            self.after.rebase_policy(source)?
        } else {
            source.clone()
        })
    }
}

impl SourceBackedSlideCommit {
    /// Candidate snapshot after this edit.
    #[must_use]
    #[cfg(test)]
    pub(crate) const fn snapshot(&self) -> &SourceBackedSlideSnapshot {
        &self.snapshot
    }

    /// Exact-source slide patch for this edit.
    #[must_use]
    #[cfg(test)]
    pub(crate) const fn patch(&self) -> &SourceBackedSlidePatch {
        &self.patch
    }

    #[cfg(test)]
    fn is_changed(&self) -> bool {
        self.patch.is_changed()
    }

    /// The outcome of this edit or commit.
    #[must_use]
    #[cfg(test)]
    pub(crate) fn diagnostics(&self) -> crate::diagnostics::Diagnostics {
        crate::diagnostics::Diagnostics::new(self.is_changed())
    }
}

impl SvgImage {
    /// Zero-based presentation position of the owning slide.
    #[must_use]
    pub const fn slide_position(&self) -> usize {
        self.slide.position
    }

    /// Zero-based direct-picture position in the owning slide.
    #[must_use]
    pub const fn image_position(&self) -> usize {
        self.image_position
    }

    /// Relationship ID retained by the native SVG `asvg:svgBlip`.
    #[must_use]
    pub fn svg_relationship_id(&self) -> &str {
        &self.svg_relationship_id
    }

    /// Relationship ID retained by the raster compatibility `a:blip`.
    #[must_use]
    pub fn raster_relationship_id(&self) -> &str {
        &self.raster_relationship_id
    }

    /// Internal SVG Part URI in this exact source snapshot.
    #[must_use]
    pub const fn svg_part_uri(&self) -> &PackURI {
        &self.svg_part_uri
    }

    /// Internal raster fallback Part URI in this exact source snapshot.
    #[must_use]
    pub const fn raster_part_uri(&self) -> &PackURI {
        &self.raster_part_uri
    }

    /// Borrow the source or candidate SVG payload bytes.
    #[must_use]
    pub fn svg_bytes(&self) -> &[u8] {
        self.svg.as_bytes()
    }

    /// Borrow the source or candidate raster fallback bytes.
    #[must_use]
    pub fn raster_bytes(&self) -> &[u8] {
        self.raster.as_bytes()
    }

    /// Start an isolated edit from this exact SVG source snapshot.
    #[must_use]
    pub(crate) fn edit(&self) -> SourceBackedSvgEdit {
        SourceBackedSvgEdit {
            source: self.clone(),
            svg_part_uri: self.svg_part_uri.clone(),
            raster_part_uri: self.raster_part_uri.clone(),
            svg: self.svg.clone(),
            raster: self.raster.clone(),
            operation_used: false,
        }
    }

    fn edit_checked(&self) -> Result<SourceBackedSvgEdit> {
        self.slide.check_execution()?;
        Ok(self.edit())
    }

    fn same_source(&self, other: &Self) -> bool {
        self.slide.same_source(&other.slide)
            && self.image_position == other.image_position
            && self.svg_relationship_id == other.svg_relationship_id
            && self.svg_relationship_type == other.svg_relationship_type
            && self.raster_relationship_id == other.raster_relationship_id
            && self.raster_relationship_type == other.raster_relationship_type
            && self.svg_part_uri == other.svg_part_uri
            && self.raster_part_uri == other.raster_part_uri
            && self.svg.as_bytes() == other.svg.as_bytes()
            && self.raster.as_bytes() == other.raster.as_bytes()
    }
}

impl SourceBackedSvgEdit {
    /// Exact immutable source snapshot against which this edit was created.
    #[must_use]
    pub(crate) const fn source(&self) -> &SvgImage {
        &self.source
    }

    /// Replace the SVG and required raster fallback payloads atomically.
    ///
    /// The source relationship IDs remain stable. Optional replacement Part
    /// URIs are applied as relationship-target changes during publication. An
    /// equal payload and equal target URI is an exact no-op, retaining the
    /// original managed source allocation.
    pub(crate) fn replace(&mut self, replacement: SvgReplacement) -> Result<bool> {
        if self.operation_used {
            return Err(Error::refused(
                litchi_core::edit::RefusalKind::UnsupportedOperation,
                "replace_svg_image",
                "source-backed SVG edits support one atomic replacement",
            ));
        }
        let svg_part_uri = replacement
            .svg_part_uri
            .unwrap_or_else(|| self.source.svg_part_uri.clone());
        let raster_part_uri = replacement
            .raster_part_uri
            .unwrap_or_else(|| self.source.raster_part_uri.clone());
        let svg_part_uri = if svg_part_uri.is_equivalent_to(&self.source.svg_part_uri) {
            self.source.svg_part_uri.clone()
        } else {
            svg_part_uri
        };
        let raster_part_uri = if raster_part_uri.is_equivalent_to(&self.source.raster_part_uri) {
            self.source.raster_part_uri.clone()
        } else {
            raster_part_uri
        };
        validate_svg_media_uri(&svg_part_uri)?;
        validate_svg_media_uri(&raster_part_uri)?;
        if svg_part_uri.is_equivalent_to(&raster_part_uri) {
            return Err(Error::Relationship(
                "SVG and raster fallback Part URIs must differ".into(),
            ));
        }
        check_svg_replacement_payload(&self.source.limits, replacement.svg.len(), true)?;
        check_svg_replacement_payload(&self.source.limits, replacement.raster.len(), false)?;
        if replacement.svg.is_empty() || replacement.raster.is_empty() {
            return Err(Error::Invalid(
                "SVG and raster fallback payloads cannot be empty".into(),
            ));
        }
        let svg = stage_svg_payload(&self.svg, replacement.svg);
        let raster = stage_svg_payload(&self.raster, replacement.raster);
        let changed = svg_part_uri != self.svg_part_uri
            || raster_part_uri != self.raster_part_uri
            || svg.as_bytes() != self.svg.as_bytes()
            || raster.as_bytes() != self.raster.as_bytes();
        self.svg_part_uri = svg_part_uri;
        self.raster_part_uri = raster_part_uri;
        self.svg = svg;
        self.raster = raster;
        self.operation_used = true;
        Ok(changed)
    }

    /// Validate and freeze this isolated SVG transaction.
    #[must_use]
    pub(crate) fn commit(self) -> SourceBackedSvgCommit {
        let snapshot = SvgImage {
            slide: self.source.slide.clone(),
            image_position: self.source.image_position,
            svg_relationship_id: self.source.svg_relationship_id.clone(),
            svg_relationship_type: self.source.svg_relationship_type.clone(),
            raster_relationship_id: self.source.raster_relationship_id.clone(),
            raster_relationship_type: self.source.raster_relationship_type.clone(),
            svg_part_uri: self.svg_part_uri,
            raster_part_uri: self.raster_part_uri,
            svg: self.svg,
            raster: self.raster,
            limits: self.source.limits,
        };
        let patch = SourceBackedSvgPatch {
            before: self.source,
            after: snapshot.clone(),
        };
        SourceBackedSvgCommit { snapshot, patch }
    }

    /// Check the retained execution policy and freeze this transaction.
    pub(crate) fn commit_checked(self) -> Result<SourceBackedSvgCommit> {
        self.source.slide.check_execution()?;
        let commit = self.commit();
        commit.snapshot.slide.check_execution()?;
        Ok(commit)
    }
}

impl SourceBackedSvgPatch {
    /// The slide and picture positions this patch addresses.
    pub(crate) fn location(&self) -> (usize, usize) {
        (self.before.slide.position, self.before.image_position)
    }

    /// The commit that publishes this patch.
    pub(crate) fn to_commit(&self) -> SourceBackedSvgCommit {
        SourceBackedSvgCommit {
            snapshot: self.after.clone(),
            patch: self.clone(),
        }
    }

    fn is_changed(&self) -> bool {
        !self.before.same_source(&self.after)
    }

    /// Whether this patch leaves its source unchanged.
    #[must_use]
    pub(crate) fn is_noop(&self) -> bool {
        !self.is_changed()
    }

    /// Return the exact inverse SVG patch.
    #[must_use]
    pub(crate) fn inverse(&self) -> Self {
        Self {
            before: self.after.clone(),
            after: self.before.clone(),
        }
    }

    /// Apply only to the exact source snapshot captured by this patch.
    pub(crate) fn apply(&self, source: &SvgImage) -> Result<SvgImage> {
        self.before.slide.check_execution()?;
        source.slide.check_execution()?;
        if !source.same_source(&self.before) {
            return Err(Error::PatchConflict);
        }
        if !self.is_changed() {
            return Ok(source.clone());
        }
        let mut target = self.after.clone();
        target.slide = if self.after.slide.same_source(&source.slide) {
            // An SVG/media-only change does not alter slide XML. Keep the
            // recipient snapshot verbatim so opaque slide markup is not
            // newly admitted under its semantic profile.
            source.slide.clone()
        } else {
            self.after.slide.rebase_policy(&source.slide)?
        };
        target.limits = source.limits;
        check_svg_replacement_payload(&target.limits, target.svg.as_bytes().len(), true)?;
        check_svg_replacement_payload(&target.limits, target.raster.as_bytes().len(), false)?;
        Ok(target)
    }
}

impl SourceBackedSvgCommit {
    /// Candidate SVG transaction snapshot after this edit.
    #[must_use]
    #[cfg(test)]
    pub(crate) const fn snapshot(&self) -> &SvgImage {
        &self.snapshot
    }

    /// Exact-source SVG patch for this edit.
    #[must_use]
    pub(crate) const fn patch(&self) -> &SourceBackedSvgPatch {
        &self.patch
    }

    #[cfg(test)]
    fn is_changed(&self) -> bool {
        self.patch.is_changed()
    }

    /// The outcome of this edit or commit.
    #[must_use]
    #[cfg(test)]
    pub(crate) fn diagnostics(&self) -> crate::diagnostics::Diagnostics {
        crate::diagnostics::Diagnostics::new(self.is_changed())
    }
}

#[cfg(test)]
impl SourceBackedSlideBatchEdit<'_> {
    /// Atomically replace visible text in one bounded shape set on one slide.
    ///
    /// A slide position may be selected only once in the outer batch. An empty
    /// shape batch is a no-op and does not select the slide. The return value
    /// is the number of shapes whose text bytes changed.
    #[cfg(test)]
    pub(crate) fn set_shape_texts(
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
            return Err(Error::refused(
                litchi_core::edit::RefusalKind::UnsupportedOperation,
                "set_shape_texts",
                "a source-backed slide batch may select each slide only once",
            ));
        }
        if self.commits.len() >= MAX_SOURCE_BACKED_SLIDE_BATCH {
            return Err(Error::Limit {
                kind: crate::opened::LimitKind::Codec(crate::opened::CodecLimitKind::Slides),
                observed: self.commits.len().checked_add(1).unwrap_or(usize::MAX),
                maximum: MAX_SOURCE_BACKED_SLIDE_BATCH,
            });
        }
        let added_text_bytes = replacements
            .iter()
            .try_fold(0_usize, |total, replacement| {
                total
                    .checked_add(replacement.text().len())
                    .ok_or(Error::Limit {
                        kind: crate::opened::LimitKind::TextBytes,
                        observed: usize::MAX,
                        maximum: crate::opened::Limits::default().max_text_bytes(),
                    })
            })?;
        let text_bytes = self
            .text_bytes
            .checked_add(added_text_bytes)
            .ok_or(Error::Limit {
                kind: crate::opened::LimitKind::TextBytes,
                observed: usize::MAX,
                maximum: crate::opened::Limits::default().max_text_bytes(),
            })?;
        if text_bytes > crate::opened::Limits::default().max_text_bytes() {
            return Err(Error::Limit {
                kind: crate::opened::LimitKind::TextBytes,
                observed: text_bytes,
                maximum: crate::opened::Limits::default().max_text_bytes(),
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
    #[cfg(test)]
    pub(crate) fn commit(mut self) -> Result<SourceBackedSlideBatchCommit> {
        self.editor.package.check_execution()?;
        if self.commits.is_empty() {
            return Err(Error::refused(
                litchi_core::edit::RefusalKind::UnsupportedOperation,
                "commit_slide_batch",
                "a source-backed slide batch requires at least one selected slide",
            ));
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

#[cfg(test)]
impl SourceBackedSlideBatchSnapshot {
    /// Number of selected slides in this exact-source snapshot.
    #[must_use]
    #[cfg(test)]
    pub(crate) const fn slide_count(&self) -> usize {
        self.slides.len()
    }

    /// Selected slide snapshots in ascending presentation order.
    #[cfg(test)]
    pub(crate) fn slides(&self) -> impl ExactSizeIterator<Item = &SourceBackedSlideSnapshot> {
        self.slides.iter()
    }

    #[cfg(test)]
    fn same_source(&self, other: &Self) -> bool {
        self.slides.len() == other.slides.len()
            && self
                .slides
                .iter()
                .zip(other.slides.iter())
                .all(|(left, right)| left.same_source(right))
    }

    #[cfg(test)]
    fn check_execution(&self) -> Result<()> {
        for slide in &self.slides {
            slide.check_execution()?;
        }
        Ok(())
    }
}

#[cfg(test)]
impl SourceBackedSlideBatchPatch {
    /// Exact immutable selected slide set required by this patch.
    #[must_use]
    #[cfg(test)]
    pub(crate) const fn source(&self) -> &SourceBackedSlideBatchSnapshot {
        &self.before
    }

    #[cfg(test)]
    fn is_changed(&self) -> bool {
        !self.before.same_source(&self.after)
    }

    /// Return the exact inverse multi-slide patch.
    #[must_use]
    #[cfg(test)]
    pub(crate) fn inverse(&self) -> Self {
        Self {
            before: self.after.clone(),
            after: self.before.clone(),
        }
    }

    #[cfg(test)]
    fn try_apply_fast(
        &self,
        source: &SourceBackedSlideBatchSnapshot,
    ) -> Result<FastApply<SourceBackedSlideBatchSnapshot>> {
        if !source.same_source(&self.before) {
            return Ok(FastApply::Recapture);
        }
        self.apply(source).map(FastApply::Applied)
    }

    /// Apply only to the exact selected slide set captured by this patch.
    #[cfg(test)]
    pub(crate) fn apply(
        &self,
        source: &SourceBackedSlideBatchSnapshot,
    ) -> Result<SourceBackedSlideBatchSnapshot> {
        self.before.check_execution()?;
        source.check_execution()?;
        if !source.same_source(&self.before) {
            return Err(Error::PatchConflict);
        }
        if !self.is_changed() {
            return Ok(source.clone());
        }
        let mut slides = Vec::new();
        slides
            .try_reserve_exact(self.after.slides.len())
            .map_err(|source| Error::Allocation {
                resource: "source-backed slide batch replay snapshots",
                source,
            })?;
        let mut usage = crate::laser::Usage::default();
        for (after, recipient) in self.after.slides.iter().zip(source.slides.iter()) {
            if after.same_source(recipient) {
                // The patch did not change this slide. Preserve the
                // recipient snapshot without re-admitting opaque XML under a
                // possibly tighter semantic profile, but charge its source
                // admission ledger to the same selected batch scope.
                merge_laser_usage(&mut usage, recipient.laser_usage()?, recipient.laser_limits)?;
                slides.push(recipient.clone());
            } else {
                // One usage ledger covers every selected slide in this batch;
                // changed slides are re-admitted under the recipient policy,
                // while unchanged slides contribute their captured source
                // usage above without a second XML scan.
                slides.push(after.rebase_policy_with_usage(recipient, &mut usage)?);
            }
        }
        Ok(SourceBackedSlideBatchSnapshot {
            slides: slides.into_boxed_slice(),
        })
    }
}

#[cfg(test)]
impl SourceBackedSlideBatchCommit {
    /// Candidate selected slide set after this edit.
    #[must_use]
    #[cfg(test)]
    pub(crate) const fn snapshot(&self) -> &SourceBackedSlideBatchSnapshot {
        &self.snapshot
    }

    /// Exact-source multi-slide patch for this edit.
    #[must_use]
    #[cfg(test)]
    pub(crate) const fn patch(&self) -> &SourceBackedSlideBatchPatch {
        &self.patch
    }

    #[cfg(test)]
    fn is_changed(&self) -> bool {
        self.patch.is_changed()
    }

    /// The outcome of this edit or commit.
    #[must_use]
    #[cfg(test)]
    pub(crate) fn diagnostics(&self) -> crate::diagnostics::Diagnostics {
        crate::diagnostics::Diagnostics::new(self.is_changed())
    }
}

fn stage_svg_payload(current: &SourcePayload, replacement: Vec<u8>) -> SourcePayload {
    if current.as_bytes() == replacement.as_slice() {
        return current.clone();
    }
    SourcePayload::Edited(Arc::new(replacement))
}

fn check_svg_replacement_payload(limits: &SourceLimits, bytes: usize, svg: bool) -> Result<()> {
    if bytes as u64 > limits.max_part_bytes() {
        return Err(Error::Limit {
            kind: crate::opened::LimitKind::Codec(crate::opened::CodecLimitKind::PartBytes),
            observed: bytes,
            maximum: usize::try_from(limits.max_part_bytes()).unwrap_or(usize::MAX),
        });
    }
    if svg && bytes == 0 {
        return Err(Error::Invalid("SVG payload cannot be empty".into()));
    }
    Ok(())
}

fn validate_svg_media_uri(uri: &PackURI) -> Result<()> {
    if !uri.as_str().starts_with("/ppt/media/") || uri.filename().is_empty() {
        return Err(Error::Relationship(format!(
            "SVG transaction target '{}' is outside /ppt/media",
            uri.as_str()
        )));
    }
    let lower = uri.as_str().to_ascii_lowercase();
    if lower.ends_with(".rels") || lower.ends_with("/[content_types].xml") {
        return Err(Error::Relationship(
            "SVG transaction target is a reserved OPC metadata member".into(),
        ));
    }
    Ok(())
}

fn build_svg_topology_plan(
    package: &SourceBackedPackage,
    current: &SvgImage,
    target: &SvgImage,
) -> Result<SourceTopologyPlan> {
    if !target.slide.same_source(&current.slide)
        || target.image_position != current.image_position
        || target.svg_relationship_id != current.svg_relationship_id
        || target.svg_relationship_type != current.svg_relationship_type
        || target.raster_relationship_id != current.raster_relationship_id
        || target.raster_relationship_type != current.raster_relationship_type
    {
        return Err(Error::PatchConflict);
    }
    let owner = current.slide.part_uri.clone();
    let mut plan = SourceTopologyPlan::new();
    stage_svg_media_part(
        package,
        &mut plan,
        &owner,
        &current.svg_relationship_id,
        &current.svg_part_uri,
        &target.svg_part_uri,
        target.svg.as_bytes(),
        true,
    )?;
    stage_svg_media_part(
        package,
        &mut plan,
        &owner,
        &current.raster_relationship_id,
        &current.raster_part_uri,
        &target.raster_part_uri,
        target.raster.as_bytes(),
        false,
    )?;
    if current.svg_part_uri != target.svg_part_uri {
        plan.try_replace_internal_relationship(
            owner.clone(),
            current.svg_relationship_id.clone(),
            current.svg_relationship_type.clone(),
            target.svg_part_uri.clone(),
        )?;
        remove_svg_media_if_unreferenced(
            package,
            &mut plan,
            &owner,
            &current.svg_relationship_id,
            &current.svg_part_uri,
        )?;
    }
    if current.raster_part_uri != target.raster_part_uri {
        plan.try_replace_internal_relationship(
            owner.clone(),
            current.raster_relationship_id.clone(),
            current.raster_relationship_type.clone(),
            target.raster_part_uri.clone(),
        )?;
        remove_svg_media_if_unreferenced(
            package,
            &mut plan,
            &owner,
            &current.raster_relationship_id,
            &current.raster_part_uri,
        )?;
    }
    Ok(plan)
}

fn stage_svg_media_part(
    package: &SourceBackedPackage,
    plan: &mut SourceTopologyPlan,
    owner: &PackURI,
    relationship_id: &str,
    current_uri: &PackURI,
    target_uri: &PackURI,
    payload: &[u8],
    svg: bool,
) -> Result<()> {
    validate_svg_media_uri(target_uri)?;
    let content_type = if svg { "image/svg+xml" } else { "image/png" };
    match package.part(target_uri) {
        Ok(view) => {
            if svg {
                if !is_svg_content_type(view.content_type()) {
                    return Err(Error::ContentType {
                        expected: "image/svg+xml".into(),
                        actual: view.content_type().to_owned(),
                    });
                }
            } else if !is_png_content_type(view.content_type()) {
                return Err(Error::ContentType {
                    expected: "image/png".into(),
                    actual: view.content_type().to_owned(),
                });
            }
            if !view.rels().is_empty() {
                return Err(Error::Relationship(format!(
                    "SVG transaction media Part '{}' has outbound relationships",
                    target_uri.as_str()
                )));
            }
            let existing = view.data()?;
            if existing.as_bytes() != payload {
                if has_other_inbound_media(package, target_uri, owner, relationship_id)? {
                    return Err(Error::Relationship(format!(
                        "SVG transaction refuses to mutate shared media Part '{}'",
                        target_uri.as_str()
                    )));
                }
                plan.try_replace_part(target_uri.clone(), payload.to_vec())?;
            }
        },
        Err(litchi_opc::OpcError::PartNotFound(_)) => {
            if target_uri == current_uri {
                return Err(Error::PartNotFound(target_uri.to_string()));
            }
            plan.try_add_part(target_uri.clone(), content_type, payload.to_vec())?;
        },
        Err(error) => return Err(error.into()),
    }
    Ok(())
}

fn remove_svg_media_if_unreferenced(
    package: &SourceBackedPackage,
    plan: &mut SourceTopologyPlan,
    owner: &PackURI,
    relationship_id: &str,
    uri: &PackURI,
) -> Result<()> {
    if has_unmodeled_relationship_members(package) {
        return Err(Error::Relationship(
            "SVG transaction refuses media removal while an opaque relationships member is retained"
                .into(),
        ));
    }
    if !has_other_inbound_media(package, uri, owner, relationship_id)? {
        plan.try_remove_part(uri.clone())?;
    }
    Ok(())
}

fn has_unmodeled_relationship_members(package: &SourceBackedPackage) -> bool {
    package.physical_member_names().any(|name| {
        let bytes = name.as_bytes();
        if bytes.len() < b".rels".len()
            || !bytes[bytes.len() - b".rels".len()..].eq_ignore_ascii_case(b".rels")
        {
            return false;
        }
        if name.eq_ignore_ascii_case("_rels/.rels") {
            return false;
        }
        package.iter_parts().all(|part| {
            part.partname()
                .rels_uri()
                .ok()
                .is_none_or(|rels_uri| !rels_uri.membername().eq_ignore_ascii_case(name))
        })
    })
}

fn has_other_inbound_media(
    package: &SourceBackedPackage,
    target: &PackURI,
    excluded_owner: &PackURI,
    excluded_relationship_id: &str,
) -> Result<bool> {
    let matches = |owner: Option<&PackURI>, relationship: &litchi_opc::Relationship| {
        if relationship.target_mode() != TargetMode::Internal {
            return false;
        }
        let Some(partname) = relationship.target_partname().ok() else {
            return false;
        };
        if !partname.is_equivalent_to(target) {
            return false;
        }
        owner != Some(excluded_owner) || relationship.r_id() != excluded_relationship_id
    };
    if package
        .rels()
        .iter()
        .any(|relationship| matches(None, relationship))
    {
        return Ok(true);
    }
    for part in package.iter_parts() {
        let owner = part.partname();
        if part
            .rels()
            .iter()
            .any(|relationship| matches(Some(owner), relationship))
        {
            return Ok(true);
        }
    }
    Ok(false)
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

    /// Return the authored `p:cSld@name`, without the physical slide-part
    /// URI fallback returned by [`Self::name`].
    ///
    /// Only this selected slide's XML is loaded. Source-version and caller
    /// execution checks surround the read, including when the payload is
    /// already resident in the bounded source cache.
    pub fn authored_name(&self) -> Result<Option<String>> {
        self.owner.package.check_execution()?;
        let part = self.part()?;
        let name = SlidePart::from_part(&part)?.authored_name()?;
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
                kind: crate::opened::LimitKind::Codec(crate::opened::CodecLimitKind::XmlBytes),
                observed: usize::try_from(declared).unwrap_or(usize::MAX),
                maximum: raw_xml_limit,
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
        self.text_with_tolerance(&mut crate::leniency::ToleranceReport::default())
    }

    /// [`Self::text`], appending the repairs the extraction applied to
    /// `tolerance`, stamped with the slide's part name (see
    /// [`crate::parts::SlidePart::text_with_tolerance`]).
    ///
    /// # Errors
    ///
    /// Returns the errors [`Self::text`] returns.
    pub fn text_with_tolerance(
        &self,
        tolerance: &mut crate::leniency::ToleranceReport,
    ) -> Result<String> {
        self.owner.package.check_execution()?;
        let part = self.part()?;
        let mut repairs = crate::leniency::ToleranceReport::default();
        let text = SlidePart::from_part(&part)?.text_with_tolerance(&mut repairs)?;
        self.owner.package.check_execution()?;
        self.owner.package.source_version()?;
        tolerance.append(&repairs);
        Ok(text)
    }

    /// List direct `p:pic` image descriptors in scene order.
    ///
    /// This reads only the selected slide XML and its relationship metadata.
    /// It never reads an embedded `/ppt/media/` payload. Backgrounds,
    /// inherited layout/master images, notes, charts, OLE previews, and
    /// non-picture graphic frames are outside this direct-picture closure.
    /// Markup-compatibility branches owned by a picture and ambiguous or
    /// malformed picture grammar are refused rather than projected into a
    /// lossy descriptor. MCE owned by another slide child remains outside
    /// this image inventory closure.
    pub fn images(&self) -> Result<Vec<SourceImageDescriptor>> {
        match self.query_images(PictureQueryMode::All)? {
            PictureQueryResult::All(descriptors) => Ok(descriptors),
            PictureQueryResult::Selected { .. } => unreachable!("all picture query result"),
        }
    }

    fn query_images(&self, mode: PictureQueryMode) -> Result<PictureQueryResult> {
        #[cfg(test)]
        match mode {
            PictureQueryMode::All => {
                SOURCE_PICTURE_QUERY_ALL.set(SOURCE_PICTURE_QUERY_ALL.get() + 1)
            },
            PictureQueryMode::Selected(_) => {
                SOURCE_PICTURE_QUERY_SELECTED.set(SOURCE_PICTURE_QUERY_SELECTED.get() + 1)
            },
        }

        self.owner.package.check_execution()?;
        let view = self.owner.package.part(&self.data.part_uri)?;
        let part = self.load_part(&view)?;
        validate_source_slide_root(&part)?;
        validate_full_slide_picture_relationships(&self.owner.package, part.blob())?;
        // Keep the normal borrowed part validation on the safe path as well;
        // the raw-root check above prevents MCE preprocessing from masking a
        // malformed selected slide before our refusal scan runs.
        if let Err(error) = SlidePart::from_part(&part) {
            // The SVG owner scanner preserves legal opaque extension markup
            // that the general SlidePart/MCE projection deliberately refuses
            // to interpret. Other malformed slide errors remain fatal.
            if !matches!(error, Error::MarkupCompatibility(_)) {
                return Err(error);
            }
        }
        let scene = match crate::shape::Scene::read(
            part.blob(),
            crate::open_options::selected(self.owner.semantic_limits.as_ref()).shape(),
        ) {
            Ok(scene) => scene,
            Err(Error::MarkupCompatibility(_)) => {
                return self.query_images_from_raw_owner(&part, &view, mode);
            },
            Err(error) => return Err(error),
        };
        let mut picture_xmls = Vec::new();
        picture_xmls
            .try_reserve(scene.len())
            .map_err(|source| Error::Allocation {
                resource: "source-backed picture XML fragments",
                source,
            })?;
        for shape in scene.iter() {
            if let Shape::Picture(picture) = shape {
                picture_xmls.push(picture.common().xml()?);
            }
        }
        let relationships = svg_lifecycle::parse_picture_relationships_with_source(
            part.blob(),
            picture_xmls.iter().copied(),
            self.owner.limits.max_part_bytes(),
        )?;
        let mut relationships = relationships.into_iter();

        let mut descriptors = match mode {
            PictureQueryMode::All => {
                let mut descriptors = Vec::new();
                #[cfg(test)]
                {
                    SOURCE_PICTURE_DESCRIPTOR_RESERVE_CALLS
                        .set(SOURCE_PICTURE_DESCRIPTOR_RESERVE_CALLS.get() + 1);
                    SOURCE_PICTURE_DESCRIPTOR_RESERVE_ITEMS
                        .set(SOURCE_PICTURE_DESCRIPTOR_RESERVE_ITEMS.get() + scene.len());
                }
                descriptors
                    .try_reserve_exact(scene.len())
                    .map_err(|source| Error::Allocation {
                        resource: "source-backed picture descriptors",
                        source,
                    })?;
                Some(descriptors)
            },
            PictureQueryMode::Selected(_) => None,
        };
        let mut selected_descriptor = None;
        let mut picture_count = 0usize;
        for (shape_position, shape) in scene.iter().enumerate() {
            let Shape::Picture(picture) = shape else {
                continue;
            };
            let position = picture_count;
            picture_count = picture_count
                .checked_add(1)
                .ok_or_else(|| Error::Invalid("source-backed picture count overflow".into()))?;
            let common = picture.common();
            let relationship = relationships.next().ok_or_else(|| {
                Error::Invalid("source-backed picture relationship inventory is truncated".into())
            })?;
            let target = resolve_picture_target(&self.owner.package, &view, &relationship)?;
            let svg = relationship
                .svg
                .as_ref()
                .map(|svg| resolve_svg_target(&self.owner.package, &view, svg).map(Arc::new))
                .transpose()?;
            let descriptor = match mode {
                PictureQueryMode::All => Some(SourceImageDescriptor {
                    position,
                    shape_position,
                    id: common.id(),
                    name: common.name().map(str::to_owned),
                    bounds: common.bounds(),
                    relationship_id: relationship.id,
                    target,
                    svg,
                }),
                PictureQueryMode::Selected(selected) if selected == position => {
                    Some(SourceImageDescriptor {
                        position,
                        shape_position,
                        id: common.id(),
                        name: common.name().map(str::to_owned),
                        bounds: common.bounds(),
                        relationship_id: relationship.id,
                        target,
                        svg,
                    })
                },
                PictureQueryMode::Selected(_) => None,
            };
            if let Some(descriptor) = descriptor {
                match mode {
                    PictureQueryMode::All => descriptors
                        .as_mut()
                        .expect("all picture query retains descriptors")
                        .push(descriptor),
                    PictureQueryMode::Selected(_) => selected_descriptor = Some(descriptor),
                }
            }
        }
        self.owner.package.check_execution()?;
        self.owner.package.source_version()?;
        match mode {
            PictureQueryMode::All => Ok(PictureQueryResult::All(
                descriptors
                    .take()
                    .expect("all picture query retains descriptors"),
            )),
            PictureQueryMode::Selected(_) => Ok(PictureQueryResult::Selected {
                descriptor: selected_descriptor,
                len: picture_count,
            }),
        }
    }

    /// Query the direct picture owner without asking the general Scene
    /// processor to interpret unrelated extension payload.  The Scene path
    /// remains the metadata-rich fast path; this bounded fallback is needed
    /// for source slides whose opaque extension subtree contains legal PI or
    /// unsupported MCE markup.  Relationship and target facts still come
    /// from the same raw owner scanner used by lifecycle edits.
    fn query_images_from_raw_owner(
        &self,
        part: &SourcePart,
        view: &PartView<'_>,
        mode: PictureQueryMode,
    ) -> Result<PictureQueryResult> {
        let relationships = svg_lifecycle::parse_picture_relationships_from_source(
            part.blob(),
            self.owner.limits.max_part_bytes(),
        )?;
        let picture_count = relationships.len();
        let mut descriptors = match mode {
            PictureQueryMode::All => {
                let mut descriptors = Vec::new();
                descriptors
                    .try_reserve_exact(picture_count)
                    .map_err(|source| Error::Allocation {
                        resource: "source-backed raw picture descriptors",
                        source,
                    })?;
                Some(descriptors)
            },
            PictureQueryMode::Selected(_) => None,
        };
        for (position, relationship) in relationships.into_iter().enumerate() {
            let target = resolve_picture_target(&self.owner.package, view, &relationship)?;
            let svg = relationship
                .svg
                .as_ref()
                .map(|svg| resolve_svg_target(&self.owner.package, view, svg).map(Arc::new))
                .transpose()?;
            let descriptor = match mode {
                PictureQueryMode::All => Some(SourceImageDescriptor {
                    position,
                    // The raw fallback intentionally does not invent a
                    // processed Scene pre-order index.  The direct picture
                    // order remains exact and is the only selector used by
                    // this source-backed lifecycle.
                    shape_position: position,
                    id: None,
                    name: None,
                    bounds: None,
                    relationship_id: relationship.id,
                    target,
                    svg,
                }),
                PictureQueryMode::Selected(selected) if selected == position => {
                    Some(SourceImageDescriptor {
                        position,
                        shape_position: position,
                        id: None,
                        name: None,
                        bounds: None,
                        relationship_id: relationship.id,
                        target,
                        svg,
                    })
                },
                PictureQueryMode::Selected(_) => None,
            };
            if let Some(descriptor) = descriptor {
                match mode {
                    PictureQueryMode::All => descriptors
                        .as_mut()
                        .expect("raw all picture query retains descriptors")
                        .push(descriptor),
                    PictureQueryMode::Selected(_) => {
                        return Ok(PictureQueryResult::Selected {
                            descriptor: Some(descriptor),
                            len: picture_count,
                        });
                    },
                }
            }
        }
        match mode {
            PictureQueryMode::All => Ok(PictureQueryResult::All(
                descriptors
                    .take()
                    .expect("raw all picture query retains descriptors"),
            )),
            PictureQueryMode::Selected(_) => Ok(PictureQueryResult::Selected {
                descriptor: None,
                len: picture_count,
            }),
        }
    }

    /// Select one direct image descriptor by exact zero-based image position.
    ///
    /// The selected embedded payload is still deferred; use
    /// [`Self::read_image`] when payload bytes are required.
    pub fn image(&self, position: usize) -> Result<SourceImageDescriptor> {
        let PictureQueryResult::Selected { descriptor, len } =
            self.query_images(PictureQueryMode::Selected(position))?
        else {
            unreachable!("selected picture query result")
        };
        descriptor.ok_or(Error::IndexOutOfBounds {
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
        let PictureQueryResult::Selected { descriptor, len } =
            self.query_images(PictureQueryMode::Selected(position))?
        else {
            unreachable!("selected picture query result")
        };
        let descriptor = descriptor.ok_or(Error::IndexOutOfBounds {
            index: position,
            len,
        })?;
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

    /// Read one embedded SVG resource paired with a direct picture by exact
    /// zero-based image position.  Linked SVG targets remain inert and are
    /// reported through [`SourceImageDescriptor::svg`].
    pub fn read_svg_image(&self, position: usize) -> Result<SourceSvgImage> {
        self.owner.package.check_execution()?;
        let PictureQueryResult::Selected { descriptor, len } =
            self.query_images(PictureQueryMode::Selected(position))?
        else {
            unreachable!("selected picture query result")
        };
        let descriptor = descriptor.ok_or(Error::IndexOutOfBounds {
            index: position,
            len,
        })?;
        let svg = descriptor.svg.clone().ok_or_else(|| {
            Error::Relationship(format!(
                "source-backed image {position} has no SVG relationship"
            ))
        })?;
        let SourceImageTarget::Internal { part_uri, .. } = &svg.target else {
            return Err(Error::Relationship(format!(
                "source-backed SVG image {position} has an external target and is inert"
            )));
        };
        let view = self.owner.package.part(part_uri)?;
        if !is_svg_content_type(view.content_type()) {
            return Err(Error::ContentType {
                expected: "image/svg+xml".into(),
                actual: view.content_type().to_owned(),
            });
        }
        if !view.rels().is_empty() {
            return Err(Error::Relationship(format!(
                "source-backed SVG image part '{}' has outbound relationships",
                part_uri.as_str()
            )));
        }
        let data = view.data()?;
        self.owner.package.check_execution()?;
        self.owner.package.source_version()?;
        Ok(SourceSvgImage {
            descriptor: svg,
            data,
        })
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
        let part = SourcePart::from_view(view, view.data()?)?;
        self.owner.package.check_execution()?;
        Ok(part)
    }
}

struct PictureRelationship {
    id: String,
    external: bool,
    svg: Option<SvgRelationship>,
}

struct SvgRelationship {
    id: String,
    external: bool,
    blip: SvgBlip,
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
    BlipExtensionList,
    BlipExtension,
    OpaqueBlipExtension,
    SvgBlipExtension,
    SvgBlip,
    Other,
}

const MCE_NAMESPACE: &[u8] = b"http://schemas.openxmlformats.org/markup-compatibility/2006";
const DRAWINGML_NAMESPACE: &[u8] = b"http://schemas.openxmlformats.org/drawingml/2006/main";
const STRICT_DRAWINGML_NAMESPACE: &[u8] = b"http://purl.oclc.org/ooxml/drawingml/main";
const SVG_BLIP_NAMESPACE: &[u8] = b"http://schemas.microsoft.com/office/drawing/2016/SVG/main";
const SVG_BLIP_EXTENSION_URI: &[u8] = b"{96DAC541-7B7A-43D3-8B79-37D633B846F1}";

fn parse_picture_relationship_with_limit(
    xml: &[u8],
    max_output_bytes: u64,
) -> Result<PictureRelationship> {
    parse_picture_relationship_inner(xml, false, max_output_bytes)
}

fn parse_picture_relationship_strict_with_limit(
    xml: &[u8],
    max_output_bytes: u64,
) -> Result<PictureRelationship> {
    parse_picture_relationship_inner(xml, true, max_output_bytes)
}

fn parse_picture_relationship_inner(
    xml: &[u8],
    reject_empty_svg_extension: bool,
    max_output_bytes: u64,
) -> Result<PictureRelationship> {
    let mut reader = NsReader::from_reader(xml);
    let origin = ReaderOrigin::of(xml);
    let mut stack: Vec<(PictureNode, Vec<u8>)> = Vec::new();
    let mut root_seen = false;
    let mut picture_stage = 0_u8;
    let mut blip_fill_stage = 0_u8;
    let mut saw_blip = false;
    let mut saw_stretch = false;
    let mut saw_tile = false;
    let mut saw_fill_rectangle = false;
    let mut saw_blip_extension_list = false;
    let mut saw_svg_blip = false;
    let mut current_svg_extension_has_child = false;
    let mut svg_start = None;
    let mut relationship = None;
    let mut root_namespace_declarations = Vec::<(Vec<u8>, Vec<u8>)>::new();
    loop {
        let event_start = position(&reader, origin)?;
        let event = reader
            .read_event()
            .map_err(|error| Error::Xml(error.to_string()))?;
        let event_end = position(&reader, origin)?;
        let (namespace, event) = reader.resolver().resolve_event(event);
        let is_start = matches!(&event, Event::Start(_));
        match event {
            Event::Start(element) | Event::Empty(element) => {
                let parent = stack.last().map(|(node, _)| *node);
                let raw_node = picture_node(&namespace, element.name());
                let opaque_parent = parent == Some(PictureNode::OpaqueBlipExtension);
                let mut node = if opaque_parent {
                    PictureNode::OpaqueBlipExtension
                } else {
                    raw_node
                };
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
                    root_namespace_declarations =
                        namespace_declarations(&element, reader.decoder())?;
                }
                if !opaque_parent && node == PictureNode::Picture && !stack.is_empty() {
                    return Err(Error::Invalid(
                        "picture descriptor contains a nested p:pic".into(),
                    ));
                }

                // `a:blip/extLst` can carry producer extensions unrelated to
                // the typed SVG descriptor.  Their payload is opaque to this
                // inventory, while the SVG URI selects the strict typed
                // child grammar below.
                if parent == Some(PictureNode::BlipExtensionList) {
                    if raw_node != PictureNode::BlipExtension {
                        return Err(Error::Invalid(
                            "picture a:blip extLst children must be a:ext".into(),
                        ));
                    }
                    node = if validate_blip_extension_attributes(&element, reader.decoder())? {
                        PictureNode::SvgBlipExtension
                    } else {
                        PictureNode::OpaqueBlipExtension
                    };
                }

                if reject_empty_svg_extension && node == PictureNode::SvgBlipExtension {
                    if !is_start {
                        return Err(Error::Invalid(
                            "recognized SVG extension must contain one svgBlip child".into(),
                        ));
                    }
                    current_svg_extension_has_child = false;
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
                            relationship = Some(PictureRelationship {
                                id,
                                external,
                                svg: None,
                            });
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
                    Some(PictureNode::Blip) => {
                        if node != PictureNode::BlipExtensionList || saw_blip_extension_list {
                            return Err(Error::Invalid(
                                "picture a:blip permits at most one direct extLst".into(),
                            ));
                        }
                        saw_blip_extension_list = true;
                    },
                    Some(PictureNode::BlipExtensionList) => {
                        // The extension node was validated and classified
                        // above.  Keep this arm explicit to document the
                        // finite parent grammar.
                    },
                    Some(PictureNode::SvgBlipExtension) => {
                        if node != PictureNode::SvgBlip || saw_svg_blip {
                            return Err(Error::Invalid(
                                "picture a:blip extension contains an unsupported or duplicate svgBlip".into(),
                            ));
                        }
                        current_svg_extension_has_child = true;
                        saw_svg_blip = true;
                        if is_start {
                            svg_start = Some(event_start);
                        } else {
                            let fragment = xml.get(event_start..event_end).ok_or_else(|| {
                                Error::Invalid("picture svgBlip range is outside its XML".into())
                            })?;
                            attach_svg_relationship(
                                &mut relationship,
                                fragment,
                                &root_namespace_declarations,
                                max_output_bytes,
                            )?;
                        }
                    },
                    Some(PictureNode::OpaqueBlipExtension) => {},
                    Some(PictureNode::BlipExtension) => {
                        return Err(Error::Invalid(
                            "picture a:blip extension classification is unavailable".into(),
                        ));
                    },
                    Some(PictureNode::SvgBlip) => {},
                    Some(
                        PictureNode::SourceRectangle
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
                let (closed_node, start_name) = stack.pop().ok_or_else(|| {
                    Error::Invalid("picture descriptor has an unmatched end element".into())
                })?;
                if start_name.as_slice() != element.name().as_ref() {
                    return Err(Error::Invalid(
                        "picture descriptor has a mismatched end element".into(),
                    ));
                }
                if reject_empty_svg_extension
                    && closed_node == PictureNode::SvgBlipExtension
                    && !current_svg_extension_has_child
                {
                    return Err(Error::Invalid(
                        "recognized SVG extension must contain one svgBlip child".into(),
                    ));
                }
                if closed_node == PictureNode::SvgBlip {
                    let start = svg_start.take().ok_or_else(|| {
                        Error::Invalid("picture svgBlip source range is unavailable".into())
                    })?;
                    let fragment = xml.get(start..event_end).ok_or_else(|| {
                        Error::Invalid("picture svgBlip range is outside its XML".into())
                    })?;
                    attach_svg_relationship(
                        &mut relationship,
                        fragment,
                        &root_namespace_declarations,
                        max_output_bytes,
                    )?;
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
                            | PictureNode::BlipExtensionList
                            | PictureNode::BlipExtension
                            | PictureNode::SvgBlipExtension
                    )
                );
                if finite_parent
                    && (parent == Some(PictureNode::Blip)
                        || parent == Some(PictureNode::BlipExtensionList)
                        || parent == Some(PictureNode::BlipExtension)
                        || parent == Some(PictureNode::SvgBlipExtension)
                        || !text.as_ref().iter().all(u8::is_ascii_whitespace))
                {
                    return Err(Error::Invalid(
                        "picture descriptor contains text in a finite picture grammar element"
                            .into(),
                    ));
                }
            },
            Event::CData(_) | Event::GeneralRef(_)
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
                            | PictureNode::BlipExtensionList
                            | PictureNode::BlipExtension
                            | PictureNode::SvgBlipExtension
                    )
                }) =>
            {
                return Err(Error::Invalid(
                    "picture descriptor contains unsupported content in a finite grammar element"
                        .into(),
                ));
            },
            Event::Comment(_) => {},
            Event::Eof => break,
            Event::PI(_)
                if stack
                    .last()
                    .is_some_and(|(node, _)| *node == PictureNode::OpaqueBlipExtension) => {},
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

fn validate_blip_extension_attributes(
    element: &quick_xml::events::BytesStart<'_>,
    decoder: quick_xml::encoding::Decoder,
) -> Result<bool> {
    let mut uri_seen = false;
    let mut svg_uri = false;
    for attribute in element.checked_attributes() {
        let attribute = attribute.map_err(|error| Error::Xml(error.to_string()))?;
        if attribute.key.as_namespace_binding().is_some() {
            continue;
        }
        if attribute.key.prefix().is_some() || attribute.key.as_ref() != b"uri" {
            return Err(Error::Invalid(
                "picture a:ext contains an unsupported attribute".into(),
            ));
        }
        if uri_seen {
            return Err(Error::Invalid(
                "picture a:ext contains duplicate uri attributes".into(),
            ));
        }
        uri_seen = true;
        let value = attribute
            .decoded_and_normalized_value(quick_xml::XmlVersion::Implicit1_0, decoder)
            .map_err(|error| Error::Xml(error.to_string()))?;
        if value.is_empty() {
            return Err(Error::Invalid(
                "picture a:ext uri attribute is empty".into(),
            ));
        }
        // CT_OfficeArtExtension@uri is xsd:token.  Its whitespace facet is
        // the XML four-character class; Rust's Unicode `split_whitespace`
        // would incorrectly classify NBSP and other Unicode separators.
        let normalized = value
            .split([' ', '\t', '\r', '\n'])
            .filter(|token| !token.is_empty())
            .collect::<Vec<_>>()
            .join(" ");
        svg_uri = normalized.as_bytes() == SVG_BLIP_EXTENSION_URI;
    }
    if !uri_seen {
        return Err(Error::Invalid(
            "picture a:ext is missing its uri attribute".into(),
        ));
    }
    Ok(svg_uri)
}

fn attach_svg_relationship(
    relationship: &mut Option<PictureRelationship>,
    fragment: &[u8],
    inherited_namespaces: &[(Vec<u8>, Vec<u8>)],
    max_output_bytes: u64,
) -> Result<()> {
    let fragment = svg_lifecycle::namespace_complete_element_fragment(
        fragment,
        inherited_namespaces,
        max_output_bytes,
    )?;
    let input_limit = usize::try_from(max_output_bytes)
        .unwrap_or(usize::MAX)
        .min(litchi_drawingml::svg_blip::MAX_XML_BYTES);
    let blip = litchi_drawingml::svg_blip::read(
        &fragment,
        litchi_drawingml::svg_blip::Limits::new(
            input_limit,
            litchi_drawingml::svg_blip::MAX_XML_BYTES,
        ),
    )?;
    let (id, external) = match (blip.embedded(), blip.linked()) {
        (Some(id), None) => (id.as_str().to_owned(), false),
        (None, Some(id)) => (id.as_str().to_owned(), true),
        (Some(_), Some(_)) => {
            return Err(Error::Relationship(
                "picture svgBlip cannot carry both r:embed and r:link".into(),
            ));
        },
        (None, None) => {
            return Err(Error::Relationship(
                "picture svgBlip is missing r:embed or r:link".into(),
            ));
        },
    };
    let relationship = relationship.as_mut().ok_or_else(|| {
        Error::Relationship("picture svgBlip appears before its a:blip relationship".into())
    })?;
    if relationship.svg.is_some() {
        return Err(Error::Relationship(
            "picture contains duplicate SVG relationships".into(),
        ));
    }
    relationship.svg = Some(SvgRelationship { id, external, blip });
    Ok(())
}

fn namespace_declarations(
    element: &quick_xml::events::BytesStart<'_>,
    decoder: quick_xml::encoding::Decoder,
) -> Result<Vec<(Vec<u8>, Vec<u8>)>> {
    let mut declarations = Vec::new();
    for attribute in element.checked_attributes() {
        let attribute = attribute.map_err(|error| Error::Xml(error.to_string()))?;
        let Some(prefix) = attribute.key.as_namespace_binding() else {
            continue;
        };
        let prefix = match prefix {
            quick_xml::name::PrefixDeclaration::Default => &[][..],
            quick_xml::name::PrefixDeclaration::Named(prefix) => prefix,
        };
        let uri = attribute
            .decoded_and_normalized_value(quick_xml::XmlVersion::Implicit1_0, decoder)
            .map_err(|error| Error::Xml(error.to_string()))?;
        declarations
            .try_reserve(1)
            .map_err(|source| Error::Allocation {
                resource: "SVG relationship namespace declarations",
                source,
            })?;
        let mut prefix_copy = Vec::new();
        prefix_copy
            .try_reserve_exact(prefix.len())
            .map_err(|source| Error::Allocation {
                resource: "SVG relationship namespace prefix",
                source,
            })?;
        prefix_copy.extend_from_slice(prefix);
        let mut uri_copy = Vec::new();
        uri_copy
            .try_reserve_exact(uri.len())
            .map_err(|source| Error::Allocation {
                resource: "SVG relationship namespace URI",
                source,
            })?;
        uri_copy.extend_from_slice(uri.as_bytes());
        declarations.push((prefix_copy, uri_copy));
    }
    Ok(declarations)
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
    } else if is_drawing_name(namespace, name, b"extLst") {
        PictureNode::BlipExtensionList
    } else if is_drawing_name(namespace, name, b"ext") {
        PictureNode::BlipExtension
    } else if is_svg_blip_name(namespace, name) {
        PictureNode::SvgBlip
    } else {
        PictureNode::Other
    }
}

fn is_svg_blip_name(namespace: &ResolveResult<'_>, name: quick_xml::name::QName<'_>) -> bool {
    name.local_name().as_ref() == b"svgBlip"
        && (matches!(
            namespace,
            ResolveResult::Bound(Namespace(value)) if *value == SVG_BLIP_NAMESPACE
        ) || matches!(namespace, ResolveResult::Unknown(prefix) if prefix.as_slice() == b"asvg"))
}

fn validate_blip_attributes(
    element: &quick_xml::events::BytesStart<'_>,
    decoder: quick_xml::encoding::Decoder,
    resolver: &quick_xml::name::NamespaceResolver,
) -> Result<()> {
    let mut cstate_seen = false;
    for attribute in element.checked_attributes() {
        let attribute = attribute.map_err(|error| Error::Xml(error.to_string()))?;
        if attribute.key.as_namespace_binding().is_some() {
            continue;
        }
        let (namespace, local) = resolver.resolve_attribute(attribute.key);
        if is_mce_namespace(&namespace) {
            return Err(Error::refused(
                litchi_core::edit::RefusalKind::UnsupportedContent,
                "picture relationship parsing",
                "markup-compatibility attributes are refused",
            ));
        }
        // Extracted p:pic fragments omit ancestor namespace declarations; the
        // full-package caller validates the inherited relationship binding.
        let is_relationship = matches!(
            namespace,
            ResolveResult::Bound(Namespace(value))
                if value == litchi_ooxml_common::relationships::TRANSITIONAL_NAMESPACE
                    || value == litchi_ooxml_common::relationships::STRICT_NAMESPACE
        ) || matches!(namespace, ResolveResult::Unknown(ref prefix) if prefix.as_slice() == b"r");
        if is_relationship {
            if !matches!(local.as_ref(), b"embed" | b"link") {
                return Err(Error::Invalid(
                    "picture a:blip contains an unsupported attribute".into(),
                ));
            }
            let _ = attribute
                .decoded_and_normalized_value(quick_xml::XmlVersion::Implicit1_0, decoder)
                .map_err(|error| Error::Xml(error.to_string()))?;
            continue;
        }
        if matches!(namespace, ResolveResult::Unbound) {
            if local.as_ref() != b"cstate" {
                return Err(Error::Invalid(
                    "picture a:blip contains an unsupported attribute".into(),
                ));
            }
            if cstate_seen {
                return Err(Error::Invalid(
                    "picture a:blip contains duplicate cstate attributes".into(),
                ));
            }
            let value = attribute
                .decoded_and_normalized_value(quick_xml::XmlVersion::Implicit1_0, decoder)
                .map_err(|error| Error::Xml(error.to_string()))?;
            let token = litchi_ooxml_common::xml::xsd_token_atom(value.as_ref());
            if !matches!(
                token,
                Some("email" | "screen" | "print" | "hqprint" | "none")
            ) {
                return Err(Error::Invalid(
                    "picture a:blip contains an invalid cstate value".into(),
                ));
            }
            cstate_seen = true;
            continue;
        }
        let is_drawingml = matches!(
            namespace,
            ResolveResult::Bound(Namespace(value))
                if value == DRAWINGML_NAMESPACE || value == STRICT_DRAWINGML_NAMESPACE
        ) || matches!(namespace, ResolveResult::Unknown(ref prefix) if prefix.as_slice() == b"a");
        if is_drawingml || matches!(namespace, ResolveResult::Unknown(_)) {
            return Err(Error::Invalid(
                "picture a:blip contains an unsupported attribute".into(),
            ));
        }
        // Attributes in a bound foreign namespace are preserved as opaque XML
        // and do not belong to the DrawingML blip semantic owner.
    }
    Ok(())
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

fn is_opaque_drawing_extension(
    element: &quick_xml::events::BytesStart<'_>,
    decoder: quick_xml::encoding::Decoder,
) -> Result<bool> {
    let mut uri_seen = false;
    let mut native_uri = false;
    for attribute in element.checked_attributes() {
        let attribute = attribute.map_err(|error| Error::Xml(error.to_string()))?;
        if attribute.key.as_namespace_binding().is_some() {
            continue;
        }
        if attribute.key.prefix().is_none() && attribute.key.as_ref() == b"uri" {
            if uri_seen {
                return Ok(true);
            }
            uri_seen = true;
            let uri = attribute
                .decoded_and_normalized_value(quick_xml::XmlVersion::Implicit1_0, decoder)
                .map_err(|error| Error::Xml(error.to_string()))?;
            // The admitted SVG GUID has no internal XML whitespace.  Trimming
            // the xsd:token edge class is therefore enough to recognize its
            // profile without allocating a collapsed-token buffer; any
            // internal whitespace remains a foreign URI.
            native_uri = uri
                .trim_matches(|character| matches!(character, ' ' | '\t' | '\r' | '\n'))
                .as_bytes()
                == SVG_BLIP_EXTENSION_URI;
        }
    }
    if !uri_seen {
        return Ok(true);
    }
    Ok(!native_uri)
}

fn validate_full_slide_picture_relationships(
    package: &SourceBackedPackage,
    xml: &[u8],
) -> Result<()> {
    let mut reader = NsReader::from_reader(xml);
    let mut stack = Vec::<Vec<u8>>::new();
    let mut picture_depth = None;
    let mut picture_blip_fill_depth = None;
    // A direct DrawingML extension with an unrecognized URI is opaque.  Keep
    // validating its XML stack, but do not infer a typed p:pic/a:blip owner
    // from namespace lookalikes nested inside that payload.
    let mut opaque_extension_depth = None;
    let mut root_seen = false;
    let mut root_closed = false;
    let mut expected_relationship_namespace = None;
    let mut buffer = Vec::new();

    loop {
        package.check_execution()?;
        let (_, event) = reader
            .read_resolved_event_into(&mut buffer)
            .map_err(|error| Error::Xml(error.to_string()))?;
        let event = event.into_owned();
        let resolver = reader.resolver().clone();
        let decoder = reader.decoder();
        let namespace = match &event {
            Event::Start(element) | Event::Empty(element) => {
                resolver.resolve_prefix(element.name().prefix(), true)
            },
            Event::End(element) => resolver.resolve_prefix(element.name().prefix(), true),
            _ => ResolveResult::Unbound,
        };
        let is_start = matches!(&event, Event::Start(_));
        match event {
            Event::Start(element) | Event::Empty(element) => {
                if root_closed {
                    return Err(Error::Invalid(
                        "source slide picture scan contains more than one root".into(),
                    ));
                }
                let depth = stack.len();
                let name = element.name();
                let opaque_ancestor = opaque_extension_depth.is_some();
                if !root_seen {
                    if !is_presentation_name(&namespace, name, b"sld") {
                        return Err(Error::Invalid(
                            "source slide picture scan does not have a p:sld root".into(),
                        ));
                    }
                    expected_relationship_namespace = match &namespace {
                        ResolveResult::Bound(Namespace(value))
                            if *value == crate::namespace::PRESENTATIONML_NAMESPACE =>
                        {
                            Some(litchi_ooxml_common::relationships::TRANSITIONAL_NAMESPACE)
                        },
                        ResolveResult::Bound(Namespace(value))
                            if *value == crate::namespace::STRICT_PRESENTATIONML_NAMESPACE =>
                        {
                            Some(litchi_ooxml_common::relationships::STRICT_NAMESPACE)
                        },
                        _ => {
                            return Err(Error::Invalid(
                                "source slide picture scan requires a bound p:sld namespace".into(),
                            ));
                        },
                    };
                    root_seen = true;
                }

                if !opaque_ancestor && is_presentation_name(&namespace, name, b"pic") {
                    if picture_depth.is_some() {
                        return Err(Error::Invalid(
                            "source slide picture scan contains a nested p:pic".into(),
                        ));
                    }
                    if is_start {
                        picture_depth = Some(depth.checked_add(1).ok_or_else(|| {
                            Error::Invalid("source slide XML depth overflow".into())
                        })?);
                    }
                } else if !opaque_ancestor
                    && is_presentation_name(&namespace, name, b"blipFill")
                    && picture_depth == Some(depth)
                {
                    if picture_blip_fill_depth.is_some() {
                        return Err(Error::Invalid(
                            "source slide picture scan contains multiple direct blipFill elements"
                                .into(),
                        ));
                    }
                    if is_start {
                        picture_blip_fill_depth = Some(depth.checked_add(1).ok_or_else(|| {
                            Error::Invalid("source slide XML depth overflow".into())
                        })?);
                    }
                } else if !opaque_ancestor
                    && is_drawing_name(&namespace, name, b"blip")
                    && picture_blip_fill_depth == Some(depth)
                {
                    validate_full_slide_blip_attributes(
                        package,
                        &element,
                        &resolver,
                        decoder,
                        expected_relationship_namespace.ok_or_else(|| {
                            Error::Invalid("source slide relationship namespace is missing".into())
                        })?,
                    )?;
                }

                if is_start {
                    stack.try_reserve(1).map_err(|source| Error::Allocation {
                        resource: "source-backed picture XML element stack",
                        source,
                    })?;
                    let mut start_name = Vec::new();
                    start_name
                        .try_reserve_exact(name.as_ref().len())
                        .map_err(|source| Error::Allocation {
                            resource: "source-backed picture XML element name",
                            source,
                        })?;
                    start_name.extend_from_slice(name.as_ref());
                    stack.push(start_name);
                    if opaque_extension_depth.is_none()
                        && is_drawing_name(&namespace, name, b"ext")
                        && is_opaque_drawing_extension(&element, decoder)?
                    {
                        opaque_extension_depth = Some(stack.len());
                    }
                } else if stack.is_empty() {
                    root_closed = true;
                }
            },
            Event::End(element) => {
                let start_name = stack.pop().ok_or_else(|| {
                    Error::Invalid("source slide picture scan has an unmatched end element".into())
                })?;
                if start_name.as_slice() != element.name().as_ref() {
                    return Err(Error::Invalid(
                        "source slide picture scan has a mismatched end element".into(),
                    ));
                }
                let closing_depth = stack.len() + 1;
                if opaque_extension_depth == Some(closing_depth) {
                    opaque_extension_depth = None;
                }
                if picture_blip_fill_depth == Some(closing_depth) {
                    picture_blip_fill_depth = None;
                }
                if picture_depth == Some(closing_depth) {
                    picture_depth = None;
                }
                if stack.is_empty() {
                    root_closed = true;
                }
            },
            Event::Text(text)
                if stack.is_empty() && !text.as_ref().iter().all(u8::is_ascii_whitespace) =>
            {
                return Err(Error::Invalid(
                    "source slide picture scan contains text outside its root".into(),
                ));
            },
            // Processing instructions are legal opaque content in an unknown
            // extension subtree.  The owner scanner applies the stricter
            // direct-owner/MCE rules; this relationship pass only rejects a
            // document type declaration, which cannot be preserved safely.
            Event::DocType(_) => {
                return Err(Error::refused(
                    litchi_core::edit::RefusalKind::UnsupportedContent,
                    "source-backed picture inventory",
                    "DOCTYPE content is refused",
                ));
            },
            Event::PI(_) => {},
            Event::Eof => break,
            _ => {},
        }
        buffer.clear();
    }
    if !root_seen || !root_closed || !stack.is_empty() {
        return Err(Error::Invalid(
            "source slide picture scan is not one closed root".into(),
        ));
    }
    Ok(())
}

fn validate_full_slide_blip_attributes(
    package: &SourceBackedPackage,
    element: &quick_xml::events::BytesStart<'_>,
    resolver: &quick_xml::name::NamespaceResolver,
    decoder: quick_xml::encoding::Decoder,
    expected_relationship_namespace: &'static [u8],
) -> Result<()> {
    let mut embed_seen = false;
    let mut link_seen = false;
    for attribute in element.checked_attributes() {
        package.check_execution()?;
        let attribute = attribute.map_err(|error| Error::Xml(error.to_string()))?;
        if attribute.key.as_namespace_binding().is_some() {
            continue;
        }
        let key = attribute.key.as_ref();
        let (namespace, local) = resolver.resolve_attribute(attribute.key);
        if is_mce_namespace(&namespace) {
            return Err(Error::refused(
                litchi_core::edit::RefusalKind::UnsupportedContent,
                "source-backed picture inventory",
                "markup-compatibility attributes are refused",
            ));
        }
        let relationship_namespace = match &namespace {
            ResolveResult::Bound(Namespace(value))
                if *value == litchi_ooxml_common::relationships::TRANSITIONAL_NAMESPACE =>
            {
                Some(litchi_ooxml_common::relationships::TRANSITIONAL_NAMESPACE)
            },
            ResolveResult::Bound(Namespace(value))
                if *value == litchi_ooxml_common::relationships::STRICT_NAMESPACE =>
            {
                Some(litchi_ooxml_common::relationships::STRICT_NAMESPACE)
            },
            _ => None,
        };
        if key.starts_with(b"r:") && relationship_namespace != Some(expected_relationship_namespace)
        {
            return Err(Error::Invalid(
                "picture a:blip r prefix is not bound to the slide relationship namespace".into(),
            ));
        }
        if let Some(namespace) = relationship_namespace {
            if namespace != expected_relationship_namespace {
                return Err(Error::Invalid(
                    "picture a:blip relationship namespace differs from the slide dialect".into(),
                ));
            }
            match local.as_ref() {
                b"embed" => {
                    if embed_seen || link_seen {
                        return Err(Error::Relationship(
                            "picture a:blip contains duplicate or mixed image relationships".into(),
                        ));
                    }
                    embed_seen = true;
                },
                b"link" => {
                    if embed_seen || link_seen {
                        return Err(Error::Relationship(
                            "picture a:blip contains duplicate or mixed image relationships".into(),
                        ));
                    }
                    link_seen = true;
                },
                _ => {
                    return Err(Error::Invalid(
                        "picture a:blip contains an unsupported relationship attribute".into(),
                    ));
                },
            }
            let _ = attribute
                .decoded_and_normalized_value(quick_xml::XmlVersion::Implicit1_0, decoder)
                .map_err(|error| Error::Xml(error.to_string()))?;
        } else if matches!(namespace, ResolveResult::Unknown(_)) {
            return Err(Error::Invalid(
                "picture a:blip contains an unresolved attribute prefix".into(),
            ));
        }
    }
    Ok(())
}

fn resolve_picture_target(
    package: &SourceBackedPackage,
    view: &PartView<'_>,
    picture: &PictureRelationship,
) -> Result<SourceImageTarget> {
    #[cfg(test)]
    {
        SOURCE_PICTURE_TARGET_RESOLUTIONS.set(SOURCE_PICTURE_TARGET_RESOLUTIONS.get() + 1);
    }
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

fn resolve_svg_target(
    package: &SourceBackedPackage,
    view: &PartView<'_>,
    svg: &SvgRelationship,
) -> Result<SourceSvgDescriptor> {
    let relationship = view.rels().get(&svg.id).ok_or_else(|| {
        Error::Relationship(format!("picture SVG relationship '{}' is missing", svg.id))
    })?;
    if relationship.is_external() != svg.external {
        return Err(Error::Relationship(format!(
            "picture SVG relationship '{}' has the wrong target mode",
            svg.id
        )));
    }
    if relationship.reltype() != rt::IMAGE && relationship.reltype() != rt::STRICT_IMAGE {
        return Err(Error::Relationship(format!(
            "picture SVG relationship '{}' has an unexpected type",
            svg.id
        )));
    }
    let target = if relationship.is_external() {
        if relationship.target_ref().is_empty() {
            return Err(Error::Relationship(
                "external picture SVG relationship has an empty target".into(),
            ));
        }
        SourceImageTarget::External {
            target: relationship.target_ref().to_owned(),
        }
    } else {
        let part_uri = relationship.target_partname()?;
        if !part_uri.as_str().starts_with("/ppt/media/") {
            return Err(Error::Relationship(format!(
                "picture SVG target '{}' is outside /ppt/media",
                part_uri.as_str()
            )));
        }
        let media = package.part(&part_uri)?;
        if !is_svg_content_type(media.content_type()) {
            return Err(Error::ContentType {
                expected: "image/svg+xml".into(),
                actual: media.content_type().to_owned(),
            });
        }
        if !media.rels().is_empty() {
            return Err(Error::Relationship(format!(
                "picture SVG image part '{}' has outbound relationships",
                part_uri.as_str()
            )));
        }
        SourceImageTarget::Internal {
            part_uri,
            content_type: media.content_type().to_owned(),
        }
    };
    Ok(SourceSvgDescriptor {
        relationship_id: svg.id.clone(),
        target,
        blip: svg.blip.clone(),
    })
}

fn is_image_content_type(content_type: &str) -> bool {
    content_type
        .split(';')
        .next()
        .is_some_and(|value| value.trim().to_ascii_lowercase().starts_with("image/"))
}

fn is_png_content_type(content_type: &str) -> bool {
    content_type
        .split(';')
        .next()
        .is_some_and(|value| value.trim().eq_ignore_ascii_case("image/png"))
}

fn is_svg_content_type(content_type: &str) -> bool {
    content_type
        .split(';')
        .next()
        .is_some_and(|value| value.trim().eq_ignore_ascii_case("image/svg+xml"))
}

fn position(reader: &NsReader<&[u8]>, origin: ReaderOrigin) -> Result<usize> {
    origin
        .offset(reader.buffer_position())
        .ok_or_else(|| Error::Invalid("source-backed picture XML position exceeds usize".into()))
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
                slide_reference_id: clone_relationship_text(
                    reference.relationship_id(),
                    "source-backed slide reference ID",
                )?,
                presentation_relationship: relationship_binding(relationship)?,
                slide: part_binding(package, &slide)?,
            }),
        }));
    }
    package.check_execution()?;
    Ok(slides.into_boxed_slice())
}

fn source_slide_output_limit(limits: SourceLimits, shape_limits: crate::shape::Limits) -> usize {
    let bounded = limits
        .max_part_bytes()
        .min(limits.max_entry_bytes())
        .min(limits.max_total_part_bytes())
        .min(limits.max_aggregate_bytes());
    usize::try_from(bounded)
        .unwrap_or(usize::MAX)
        .min(shape_limits.output_bytes())
}

/// Build the source catalog. With `tolerance` the slide list is read with
/// the audit-6.7 repairs recorded there; without it, strictly (editors).
fn source_catalog(
    package: &SourceBackedPackage,
    tolerance: Option<&mut crate::leniency::ToleranceReport>,
) -> Result<SourceCatalog> {
    #[cfg(test)]
    SOURCE_CATALOG_BUILDS.set(SOURCE_CATALOG_BUILDS.get() + 1);
    package.check_execution()?;
    let view = package.main_document_part()?;
    let presentation = SourcePart::from_view(&view, view.data()?)?;
    package.check_execution()?;
    let presentation_binding = part_binding(package, &view)?;
    let main = PresentationPart::from_part(&presentation)?;
    let references = match tolerance {
        Some(tolerance) => main.slide_references_with_tolerance(tolerance)?,
        None => main.slide_references_strict()?,
    };
    package.check_execution()?;
    let slides = validate_slide_graph(package, &view, &references)?;
    package.check_execution()?;
    Ok(SourceCatalog {
        presentation,
        package_relationships: relationship_bindings(package, package.rels())?,
        presentation_binding,
        slides,
    })
}

fn part_binding(package: &SourceBackedPackage, view: &PartView<'_>) -> Result<PartBinding> {
    Ok(PartBinding {
        uri: view.partname().clone(),
        content_type: clone_relationship_text(
            view.content_type(),
            "source-backed part content type",
        )?,
        relationships: relationship_bindings(package, view.rels())?,
    })
}

fn relationship_bindings(
    package: &SourceBackedPackage,
    relationships: &Relationships,
) -> Result<Box<[RelationshipBinding]>> {
    package.check_execution()?;
    let expected_len = relationships.len();
    let mut bindings = Vec::new();
    bindings
        .try_reserve_exact(expected_len)
        .map_err(|source| Error::Allocation {
            resource: "source-backed relationship bindings",
            source,
        })?;
    for relationship in relationships.iter() {
        package.check_execution()?;
        bindings.push(relationship_binding(relationship)?);
    }
    if bindings.len() != expected_len {
        return Err(Error::Invalid(
            "source-backed relationship catalog length changed during iteration".into(),
        ));
    }
    bindings.sort_unstable_by(|left, right| left.id.cmp(&right.id));
    Ok(bindings.into_boxed_slice())
}

fn relationship_binding(relationship: &litchi_opc::Relationship) -> Result<RelationshipBinding> {
    Ok(RelationshipBinding {
        id: clone_relationship_text(relationship.r_id(), "source-backed relationship ID")?,
        kind: clone_relationship_text(relationship.reltype(), "source-backed relationship type")?,
        target: clone_relationship_text(
            relationship.target_ref(),
            "source-backed relationship target",
        )?,
        mode: relationship.target_mode(),
    })
}

fn clone_relationship_text(value: &str, resource: &'static str) -> Result<String> {
    let mut text = String::new();
    text.try_reserve_exact(value.len())
        .map_err(|source| Error::Allocation { resource, source })?;
    text.push_str(value);
    Ok(text)
}

#[cfg(test)]
mod tests {
    mod primary;
    mod scenes;
    #[cfg(any(unix, windows))]
    use std::fs;
    use std::io;
    use std::num::{NonZeroU64, NonZeroUsize};
    use std::sync::atomic::{AtomicU64, AtomicUsize, Ordering};
    use std::sync::{Arc, OnceLock};

    use litchi_core::{
        Budget, CancellationSource, ExecutionContext, ExecutionError, ExecutionLimits, Limits,
        ReadAt, Resource, SourceVersion,
    };
    use litchi_opc::{
        PackURI, SourceBackedPackage, SourceCacheLimits, SourceLimitKind, SourceLimits,
    };
    use soapberry_zip::office::StreamingArchiveWriter;
    #[cfg(any(unix, windows))]
    use tempfile::NamedTempFile;

    use super::{
        SourceBackedPresentation, SourceBackedPresentationEditor, SourceBackedSlideBatchCommit,
        SourceBackedSlideBatchPatch, SourceBackedSlideBatchSnapshot, SourceBackedSlideCommit,
        SourceBackedSlidePatch, SourceBackedSlideSnapshot, SourceImageTarget, SourcePayload,
        reset_source_catalog_builds, reset_source_laser_scans, reset_source_picture_query_counters,
        reset_test_semantic_scene_reads, source_catalog_builds, source_laser_scans,
        source_picture_descriptor_reserve_calls, source_picture_descriptor_reserve_items,
        source_picture_query_all, source_picture_query_selected, source_picture_target_resolutions,
        test_semantic_scene_reads, test_semantic_slide_part_reads,
    };
    use crate::Error;

    const SECOND_MARKER: &[u8] = b"source-backed-unrequested-second-slide";
    const MEDIA_MARKER: &[u8] = b"source-backed-picture-media-payload";
    const CORE_MARKER: &[u8] = b"source-backed-core-properties-payload";
    const COLD_SLIDE_MARKER: &[u8] = b"source-backed-core-cold-slide-payload";
    const COLD_MEDIA_MARKER: &[u8] = b"source-backed-core-cold-media-payload";
    const NATIVE_SVG_EXTENSION_URI: &str = "{96DAC541-7B7A-43D3-8B79-37D633B846F1}";
    const SVG_NAMESPACE_URI: &str = "http://schemas.microsoft.com/office/drawing/2016/SVG/main";
    const OPAQUE_EXTENSION_URI: &str = "{28A0092B-C50C-407e-A947-70E740481C1C}";

    struct CountingSource {
        bytes: Vec<u8>,
        marker_offset: usize,
        read_calls: AtomicUsize,
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
                read_calls: AtomicUsize::new(0),
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
            self.read_calls.fetch_add(1, Ordering::SeqCst);
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
        source_backed_pptx_with_slide_list(
            r#"<p:sldId id="256" r:id="rId1"/><p:sldId id="257" r:id="rId2"/>"#,
        )
    }

    fn source_backed_pptx_with_slide_list(slide_list: &str) -> Vec<u8> {
        let presentation = format!(
            r#"<p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><p:sldIdLst>{slide_list}</p:sldIdLst><p:sldSz cx="9144000" cy="6858000"/></p:presentation>"#
        );
        let mut writer =
            StreamingArchiveWriter::new(soapberry_zip::office::StreamingArchiveLimits::default());
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
            .write_stored("ppt/presentation.xml", presentation.as_bytes())
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

    fn changed_slide_snapshots(
        editor: &SourceBackedPresentationEditor,
        position: usize,
        old: &str,
        new: &str,
    ) -> (SourceBackedSlideSnapshot, SourceBackedSlideSnapshot) {
        let before = editor.slide_snapshot(position).unwrap();
        let source = String::from_utf8(before.xml.as_bytes().to_vec()).unwrap();
        let after_xml = source.replacen(old, new, 1).into_bytes();
        assert_ne!(after_xml, before.xml.as_bytes());
        let after = SourceBackedSlideSnapshot {
            position: before.position,
            part_uri: before.part_uri.clone(),
            xml: SourcePayload::Edited(Arc::new(after_xml)),
            closure: before.closure.clone(),
            max_output_bytes: before.max_output_bytes,
            protection_limits: before.protection_limits,
            semantic_limits: before.semantic_limits.clone(),
            source_version: before.source_version,
            lineage: before.lineage.clone(),
            context: before.context.clone(),
            laser_limits: before.laser_limits,
            laser_usage: Arc::new(OnceLock::new()),
            mce: before.mce,
        };
        (before, after)
    }

    fn changed_slide_commit(
        editor: &SourceBackedPresentationEditor,
        position: usize,
        old: &str,
        new: &str,
    ) -> SourceBackedSlideCommit {
        let (before, after) = changed_slide_snapshots(editor, position, old, new);
        SourceBackedSlideCommit {
            snapshot: after.clone(),
            patch: SourceBackedSlidePatch { before, after },
        }
    }

    fn changed_slide_batch_commit(
        editor: &SourceBackedPresentationEditor,
    ) -> SourceBackedSlideBatchCommit {
        let (before_first, after_first) =
            changed_slide_snapshots(editor, 0, "First slide", "Updated slide");
        let (before_second, after_second) =
            changed_slide_snapshots(editor, 1, "Second slide", "Changed slide");
        let before = SourceBackedSlideBatchSnapshot {
            slides: vec![before_first, before_second].into_boxed_slice(),
        };
        let after = SourceBackedSlideBatchSnapshot {
            slides: vec![after_first, after_second].into_boxed_slice(),
        };
        SourceBackedSlideBatchCommit {
            snapshot: after.clone(),
            patch: SourceBackedSlideBatchPatch { before, after },
        }
    }

    fn source_backed_pptx_with_core_properties() -> Vec<u8> {
        let mut writer =
            StreamingArchiveWriter::new(soapberry_zip::office::StreamingArchiveLimits::default());
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

    fn multiple_picture_slide(blips: &[&str]) -> Vec<u8> {
        let pictures = blips
            .iter()
            .enumerate()
            .map(|(index, blip)| {
                format!(
                    r#"<p:pic><p:nvPicPr><p:cNvPr id="{}" name="Photo {}"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr><p:blipFill>{blip}<a:stretch><a:fillRect/></a:stretch></p:blipFill><p:spPr><a:xfrm><a:off x="{}" y="{}"/><a:ext cx="{}" cy="{}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic>"#,
                    42 + index,
                    index,
                    index + 1,
                    index + 2,
                    index + 3,
                    index + 4,
                )
            })
            .collect::<String>();
        format!(
            r#"<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><p:cSld><p:spTree><p:nvGrpSpPr/><p:grpSpPr/>{pictures}</p:spTree></p:cSld></p:sld>"#
        )
        .into_bytes()
    }

    fn picture_relationships(ids: &[&str]) -> Vec<u8> {
        let relationships = ids
            .iter()
            .map(|id| {
                format!(
                    r#"<Relationship Id="{id}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image1.png"/>"#
                )
            })
            .collect::<String>();
        format!(
            r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">{relationships}</Relationships>"#
        )
        .into_bytes()
    }

    fn picture_pptx(
        slide: &[u8],
        slide_relationships: &[u8],
        media_content_type: Option<&str>,
        media_relationships: Option<&[u8]>,
    ) -> Vec<u8> {
        let mut writer =
            StreamingArchiveWriter::new(soapberry_zip::office::StreamingArchiveLimits::default());
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

    pub(super) fn svg_picture_pptx() -> Vec<u8> {
        svg_picture_pptx_with_extension(&format!(
            r#"<a:ext uri="{NATIVE_SVG_EXTENSION_URI}"><asvg:svgBlip r:embed="rIdSvg"/></a:ext>"#
        ))
    }

    fn svg_picture_pptx_with_extension(extension_list: &str) -> Vec<u8> {
        const SVG_NAMESPACE: &str = "http://schemas.microsoft.com/office/drawing/2016/SVG/main";
        let slide = picture_slide(&format!(
            r#"<a:blip r:embed="rIdImage"><a:extLst>{extension_list}</a:extLst></a:blip>"#
        ));
        let slide = String::from_utf8(slide)
            .unwrap()
            .replace(
                "xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\"",
                &format!(
                    "xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" xmlns:asvg=\"{SVG_NAMESPACE}\""
                ),
            );
        let mut writer =
            StreamingArchiveWriter::new(soapberry_zip::office::StreamingArchiveLimits::default());
        writer
            .write_stored(
                "[Content_Types].xml",
                br#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Default Extension="png" ContentType="image/png"/><Default Extension="svg" ContentType="image/svg+xml"/><Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/><Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/></Types>"#,
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
            .write_stored("ppt/slides/slide1.xml", slide.as_bytes())
            .unwrap();
        writer
            .write_stored(
                "ppt/slides/_rels/slide1.xml.rels",
                br#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdImage" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image1.png"/><Relationship Id="rIdSvg" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image1.svg"/></Relationships>"#,
            )
            .unwrap();
        writer
            .write_stored("ppt/media/image1.png", &media_payload())
            .unwrap();
        writer
            .write_stored(
                "ppt/media/image1.svg",
                br#"<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 1 1"><path d="M0 0h1v1z"/></svg>"#,
            )
            .unwrap();
        writer.finish_to_bytes().unwrap()
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
        let package = SourceBackedPackage::from_read_at(
            Arc::new(CountingSource::new(source_backed_pptx())),
            SourceLimits::default(),
            SourceCacheLimits::default(),
            litchi_opc::SourceReadPolicy::exact(),
            None,
        )
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
        let package = SourceBackedPackage::from_read_at(
            source.clone(),
            SourceLimits::default(),
            SourceCacheLimits::default(),
            litchi_opc::SourceReadPolicy::exact(),
            None,
        )
        .unwrap();
        let presentation = SourceBackedPresentation::from_source_backed_package(package).unwrap();
        assert_eq!(source.core_payload_reads.load(Ordering::SeqCst), 0);
        assert_eq!(source.slide_payload_reads.load(Ordering::SeqCst), 0);
        assert_eq!(source.media_payload_reads.load(Ordering::SeqCst), 0);

        presentation.check_source().unwrap();
        assert_eq!(source.core_payload_reads.load(Ordering::SeqCst), 0);
        assert_eq!(source.slide_payload_reads.load(Ordering::SeqCst), 0);
        assert_eq!(source.media_payload_reads.load(Ordering::SeqCst), 0);

        let mut tolerance = crate::leniency::ToleranceReport::default();
        let properties = presentation
            .properties_with_tolerance(&mut tolerance)
            .unwrap()
            .unwrap();
        assert!(tolerance.is_clean(), "{tolerance:?}");
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
        let exact_editor = SourceBackedPresentationEditor::from_read_at_with_options(
            exact_source.clone(),
            crate::OpenOptions::default()
                .source_limits(SourceLimits::default())
                .execution_context(exact_context),
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
        let under_editor = SourceBackedPresentationEditor::from_read_at_with_options(
            under_source.clone(),
            crate::OpenOptions::default()
                .source_limits(SourceLimits::default())
                .execution_context(under_context),
        )
        .unwrap();
        assert!(under_editor.slide_snapshot(1).is_err());
        assert_eq!(under_source.second_payload_reads.load(Ordering::SeqCst), 0);
        drop(under_editor);
        assert_eq!(under_budget.used(Resource::Memory), 0);
    }

    #[test]
    fn selected_semantic_bundle_survives_source_editor_and_snapshot_projection() {
        let comments = crate::comments::Limits {
            max_nodes: crate::comments::Limits::DEFAULT.max_nodes - 1,
            ..crate::comments::Limits::DEFAULT
        };
        let options = crate::OpenOptions::default().comments_limits(comments);
        let presentation = SourceBackedPresentation::from_read_at_with_options(
            Arc::new(CountingSource::new(source_backed_pptx())),
            options.clone(),
        )
        .unwrap();
        assert_eq!(
            presentation
                .inner
                .semantic_limits
                .as_ref()
                .expect("selected semantic profile")
                .comments(),
            comments
        );
        let editor = SourceBackedPresentationEditor::from_read_at_with_options(
            Arc::new(CountingSource::new(source_backed_pptx())),
            options,
        )
        .unwrap();
        assert_eq!(
            editor
                .semantic_limits
                .as_ref()
                .expect("editor profile")
                .comments(),
            comments
        );
        let snapshot = editor.slide_snapshot(0).unwrap();
        assert_eq!(
            snapshot
                .semantic_limits
                .as_ref()
                .expect("snapshot profile")
                .comments(),
            comments
        );
    }

    #[test]
    fn managed_editor_exact_noop_keeps_source_bytes_and_budgeted_handles() {
        let archive = source_backed_pptx();
        let source = Arc::new(CountingSource::new(archive.clone()));
        let (budget, _cancel, context) = source_backed_context(archive.len() as u64);
        let editor = SourceBackedPresentationEditor::from_read_at_with_options(
            source,
            crate::OpenOptions::default()
                .source_limits(SourceLimits::default())
                .execution_context(context),
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
    fn replay_rehomes_edit_reservation_into_recipient_context() {
        let archive = source_backed_pptx();
        let (source_budget, _source_cancel, source_context) = source_backed_context(u64::MAX);
        let source = SourceBackedPresentationEditor::from_read_at_with_options(
            Arc::new(CountingSource::new(archive.clone())),
            crate::OpenOptions::default().execution_context(source_context),
        )
        .unwrap();
        let mut edit = source.edit_slide(0).unwrap();
        assert!(
            edit.set_slide_properties(Some("Rebased reservation"), None)
                .unwrap()
        );
        let commit = edit.commit_checked().unwrap();
        let patch = commit.patch().clone();

        let (recipient_budget, _recipient_cancel, recipient_context) =
            source_backed_context(u64::MAX);
        let recipient = SourceBackedPresentationEditor::from_read_at_with_options(
            Arc::new(CountingSource::new(archive.clone())),
            crate::OpenOptions::default().execution_context(recipient_context),
        )
        .unwrap();
        let recipient_slide = recipient.slide_snapshot(0).unwrap();
        let before_apply = recipient_budget.used(Resource::Memory);
        let applied = patch.apply(&recipient_slide).unwrap();

        assert!(matches!(&applied.xml, SourcePayload::ReservedEdit { .. }));
        assert!(recipient_budget.used(Resource::Memory) > before_apply);
        drop(applied);
        assert_eq!(recipient_budget.used(Resource::Memory), before_apply);

        let memory_limit = (archive.len() as u64).checked_mul(1024).unwrap();
        let (limited_budget, _limited_cancel, limited_context) =
            source_backed_context(memory_limit);
        let limited_recipient = SourceBackedPresentationEditor::from_read_at_with_options(
            Arc::new(CountingSource::new(archive.clone())),
            crate::OpenOptions::default().execution_context(limited_context.clone()),
        )
        .unwrap();
        let limited_slide = limited_recipient.slide_snapshot(0).unwrap();
        let required = match &patch.after.xml {
            SourcePayload::ReservedEdit { reservation, .. } => reservation.amount(),
            _ => panic!("the source property edit must retain its reservation"),
        };
        let used_before_blocker = limited_budget.used(Resource::Memory);
        let available = memory_limit.checked_sub(used_before_blocker).unwrap();
        assert!(available >= required);
        let blocker_amount = available - required + 1;
        let blocker = limited_context
            .reserve(Resource::Memory, blocker_amount)
            .unwrap();
        let source_before_failure = source_budget.used(Resource::Memory);
        let recipient_before_failure = limited_budget.used(Resource::Memory);
        assert!(matches!(
            patch.apply(&limited_slide),
            Err(Error::Opc(litchi_opc::OpcError::Execution(
                ExecutionError::ResourceLimit(_)
            )))
        ));
        assert_eq!(source_budget.used(Resource::Memory), source_before_failure);
        assert_eq!(
            limited_budget.used(Resource::Memory),
            recipient_before_failure
        );
        drop(blocker);
        let retried = patch.apply(&limited_slide).unwrap();
        drop(retried);
        drop(limited_slide);
        drop(limited_recipient);
        assert_eq!(limited_budget.used(Resource::Memory), 0);

        drop(recipient_slide);
        drop(recipient);
        drop(patch);
        drop(commit);
        drop(source);
        assert_eq!(recipient_budget.used(Resource::Memory), 0);
        assert_eq!(source_budget.used(Resource::Memory), 0);
    }

    #[test]
    fn source_edit_reservations_keep_output_and_memory_dimensions_independent() {
        const INPUT_BYTES: usize = 8192;
        const TIGHT_OUTPUT: usize = 128;
        const MEMORY: u64 = 4 * 1024 * 1024;
        let text_replacement = [crate::opened::ShapeTextReplacement::at(0, "x")];
        let text_chunks = [&text_replacement[..]];
        let admission = crate::opened::admit_shape_text_chunks(
            &text_chunks,
            crate::opened::Limits::default().max_text_bytes(),
            TIGHT_OUTPUT,
        )
        .unwrap();
        let replacements =
            [
                crate::opened::ShapePropertiesReplacement::new(litchi_core::Position::new(0))
                    .with_text("x"),
            ];
        let property_chunks = [&replacements[..]];
        let property_admission = crate::opened::admit_shape_property_chunks(
            &property_chunks,
            crate::opened::Limits::default().max_text_bytes(),
            TIGHT_OUTPUT,
        )
        .unwrap();
        let text_box_output = INPUT_BYTES + 4096 + 6;

        fn assert_admitted(result: crate::Result<Option<Arc<litchi_core::Reservation>>>) {
            assert!(result.unwrap().is_some());
        }

        let (budget, _cancel, context) = source_backed_context(MEMORY);
        // A small serialized-output ceiling must not become the managed
        // transient-memory ceiling. Every source-backed reservation keeps the
        // caller's output bound in its own arithmetic and charges the context
        // memory budget independently.
        assert_admitted(super::reserve_source_shape_removal(
            Some(&context),
            INPUT_BYTES,
            TIGHT_OUTPUT,
        ));
        assert_admitted(super::reserve_source_slide_properties(
            Some(&context),
            INPUT_BYTES,
            Some("name"),
            TIGHT_OUTPUT,
        ));
        assert_admitted(super::reserve_source_shape_properties(
            Some(&context),
            INPUT_BYTES,
            &replacements,
            TIGHT_OUTPUT,
        ));
        assert_admitted(super::reserve_source_shape_text_chunks(
            Some(&context),
            INPUT_BYTES,
            admission,
            TIGHT_OUTPUT,
        ));
        assert_admitted(super::reserve_source_shape_property_chunks(
            Some(&context),
            INPUT_BYTES,
            property_admission,
            TIGHT_OUTPUT,
        ));
        assert_admitted(super::reserve_source_text_box(
            Some(&context),
            INPUT_BYTES,
            1,
            text_box_output,
        ));
        assert_eq!(budget.used(Resource::Memory), 0);

        // Unmanaged callers do not reserve transient memory, while retaining
        // the same checked input/output admission path.
        assert!(
            super::reserve_source_shape_removal(None, INPUT_BYTES, TIGHT_OUTPUT)
                .unwrap()
                .is_none()
        );
        assert!(
            super::reserve_source_slide_properties(None, INPUT_BYTES, Some("name"), TIGHT_OUTPUT,)
                .unwrap()
                .is_none()
        );
        assert!(
            super::reserve_source_shape_properties(None, INPUT_BYTES, &replacements, TIGHT_OUTPUT,)
                .unwrap()
                .is_none()
        );
        assert!(
            super::reserve_source_shape_text_chunks(None, INPUT_BYTES, admission, TIGHT_OUTPUT,)
                .unwrap()
                .is_none()
        );
        assert!(
            super::reserve_source_shape_property_chunks(
                None,
                INPUT_BYTES,
                property_admission,
                TIGHT_OUTPUT,
            )
            .unwrap()
            .is_none()
        );
        assert!(
            super::reserve_source_text_box(None, INPUT_BYTES, 1, text_box_output)
                .unwrap()
                .is_none()
        );

        let root_budget = Budget::root(
            "pptx-source-parent-budget",
            Limits::new(1024, u64::MAX, u64::MAX, u64::MAX, u64::MAX, u64::MAX),
        );
        let child_budget = root_budget.child(
            "pptx-source-child-budget",
            Limits::new(MEMORY, u64::MAX, u64::MAX, u64::MAX, u64::MAX, u64::MAX),
        );
        let (_cancel, cancellation) = CancellationSource::pair();
        let execution_limits = ExecutionLimits::new(
            NonZeroUsize::MIN,
            NonZeroUsize::MIN,
            NonZeroU64::new(MEMORY).unwrap(),
            0,
        )
        .unwrap();
        let tight_context = ExecutionContext::new(child_budget, cancellation, execution_limits);

        fn assert_hierarchical_limit(result: crate::Result<Option<Arc<litchi_core::Reservation>>>) {
            assert!(matches!(
                result,
                Err(Error::Opc(litchi_opc::OpcError::Execution(
                    ExecutionError::ResourceLimit(_)
                )))
            ));
        }
        assert_hierarchical_limit(super::reserve_source_shape_removal(
            Some(&tight_context),
            INPUT_BYTES,
            TIGHT_OUTPUT,
        ));
        assert_hierarchical_limit(super::reserve_source_slide_properties(
            Some(&tight_context),
            INPUT_BYTES,
            Some("name"),
            TIGHT_OUTPUT,
        ));
        assert_hierarchical_limit(super::reserve_source_shape_properties(
            Some(&tight_context),
            INPUT_BYTES,
            &replacements,
            TIGHT_OUTPUT,
        ));
        assert_hierarchical_limit(super::reserve_source_shape_text_chunks(
            Some(&tight_context),
            INPUT_BYTES,
            admission,
            TIGHT_OUTPUT,
        ));
        assert_hierarchical_limit(super::reserve_source_shape_property_chunks(
            Some(&tight_context),
            INPUT_BYTES,
            property_admission,
            TIGHT_OUTPUT,
        ));
        assert_hierarchical_limit(super::reserve_source_text_box(
            Some(&tight_context),
            INPUT_BYTES,
            1,
            text_box_output,
        ));
        assert_eq!(root_budget.used(Resource::Memory), 0);
        assert_eq!(tight_context.budget().used(Resource::Memory), 0);
    }

    #[test]
    fn source_slide_property_edit_accepts_valid_tight_output_with_managed_memory() {
        const MEMORY: u64 = 4 * 1024 * 1024;
        let archive = source_backed_pptx();
        let baseline = SourceBackedPresentationEditor::from_read_at(Arc::new(CountingSource::new(
            archive.clone(),
        )))
        .unwrap();
        let input_bytes = baseline.slide_snapshot(0).unwrap().xml.as_bytes().len();
        drop(baseline);

        let defaults = crate::shape::Limits::DEFAULT;
        let tight_shape = crate::shape::Limits::new(
            input_bytes,
            input_bytes.checked_add(128).unwrap(),
            defaults.depth(),
            defaults.nodes(),
            defaults.shapes(),
            defaults.retained_text_bytes(),
        )
        .unwrap();
        let (budget, _cancel, context) = source_backed_context(MEMORY);
        let editor = SourceBackedPresentationEditor::from_read_at_with_options(
            Arc::new(CountingSource::new(archive)),
            crate::OpenOptions::default()
                .shape_limits(tight_shape)
                .execution_context(context),
        )
        .unwrap();
        let mut edit = editor.edit_slide(0).unwrap();
        assert!(edit.set_slide_properties(Some("x"), None).unwrap());
        let commit = edit.commit_checked().unwrap();
        assert!(commit.is_changed());
        drop(commit);
        drop(editor);
        assert_eq!(budget.used(Resource::Memory), 0);
    }

    fn publish_first_slide_update(editor: SourceBackedPresentationEditor) -> Vec<u8> {
        let commit = changed_slide_commit(&editor, 0, "First slide", "Updated slide");
        let mut output = Vec::new();
        editor
            .publish_slide_commit_to_stream(&mut output, &commit)
            .unwrap();
        output
    }

    #[test]
    fn presentation_converts_into_editor_without_reindexing() {
        reset_source_catalog_builds();
        let archive = source_backed_pptx();
        let source = Arc::new(CountingSource::new(archive.clone()));
        let shared: Arc<dyn ReadAt> = source.clone();
        let presentation = SourceBackedPresentation::from_read_at(shared).unwrap();
        assert!(presentation.tolerance_report().is_clean());
        assert_eq!(source_catalog_builds(), 1);
        let reads = source.read_calls.load(Ordering::SeqCst);

        let editor = presentation.try_into_editor().unwrap();
        assert_eq!(
            source_catalog_builds(),
            1,
            "the open-time catalog is reused"
        );
        assert_eq!(
            source.read_calls.load(Ordering::SeqCst),
            reads,
            "conversion reads no source byte"
        );
        assert_eq!(editor.slide_count(), 2);

        let converted = publish_first_slide_update(editor);
        let direct = SourceBackedPresentationEditor::from_read_at(Arc::new(CountingSource::new(
            archive.clone(),
        )))
        .unwrap();
        assert_eq!(converted, publish_first_slide_update(direct));
        assert_ne!(converted, archive);
    }

    #[test]
    fn converted_editor_reproduces_the_source_on_a_no_op() {
        let archive = source_backed_pptx();
        let presentation =
            SourceBackedPresentation::from_read_at(Arc::new(CountingSource::new(archive.clone())))
                .unwrap();
        let editor = presentation.try_into_editor().unwrap();
        let mut edit = editor.edit_slide(0).unwrap();
        assert!(!edit.clear_transition().unwrap());
        let commit = edit.commit();
        let mut output = Vec::new();
        editor
            .publish_slide_commit_to_stream(&mut output, &commit)
            .unwrap();
        assert_eq!(output, archive);
    }

    #[test]
    fn shared_presentation_is_returned_unchanged() {
        let presentation = SourceBackedPresentation::from_read_at(Arc::new(CountingSource::new(
            source_backed_pptx(),
        )))
        .unwrap();

        let clone = presentation.clone();
        let refused = presentation.try_into_editor().err().unwrap();
        assert!(refused.is_shared());
        assert!(refused.error().is_none());
        let (presentation, error) = refused.into_parts();
        assert!(error.is_none());
        assert_eq!(
            presentation.slide(1).unwrap().text().unwrap(),
            "Second slide"
        );
        drop(clone);

        let slide = presentation.slide(0).unwrap();
        let refused = presentation.try_into_editor().err().unwrap();
        assert!(refused.is_shared(), "a live slide handle shares the source");
        let presentation = refused.into_parts().0;
        drop(slide);

        assert!(presentation.try_into_editor().is_ok());
    }

    #[test]
    fn changed_source_refuses_conversion_and_keeps_the_presentation() {
        let source = Arc::new(CountingSource::new(source_backed_pptx()));
        let shared: Arc<dyn ReadAt> = source.clone();
        let presentation = SourceBackedPresentation::from_read_at(shared).unwrap();
        source.changed();

        let refused = presentation.try_into_editor().err().unwrap();
        assert!(!refused.is_shared());
        assert!(
            matches!(
                refused.error(),
                Some(Error::Opc(litchi_opc::OpcError::SourceChanged { .. }))
            ),
            "{refused:?}"
        );
        assert!(std::error::Error::source(&refused).is_some());
        let (presentation, _) = refused.into_parts();
        assert_eq!(presentation.slide_count(), 2);
    }

    #[test]
    fn repaired_slide_list_gets_the_editor_strict_refusal() {
        let archive = source_backed_pptx_with_slide_list(
            r#"<p:sldId id="256" r:id="rId1"/><p:sldId id="257" r:id="rId2"/><p:sldId id="258" r:id="rId1"/>"#,
        );
        let presentation =
            SourceBackedPresentation::from_read_at(Arc::new(CountingSource::new(archive.clone())))
                .unwrap();
        assert!(!presentation.tolerance_report().is_clean());
        assert_eq!(presentation.slide_count(), 2);

        let Err(direct) =
            SourceBackedPresentationEditor::from_read_at(Arc::new(CountingSource::new(archive)))
        else {
            panic!("the editor's strict slide-list contract refuses the repaired list");
        };
        let refused = presentation.try_into_editor().err().unwrap();
        assert!(!refused.is_shared());
        assert_eq!(refused.error().unwrap().to_string(), direct.to_string());
        assert_eq!(refused.into_parts().0.slide_count(), 2);
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
    fn changed_single_publication_reuses_semantic_snapshot() {
        reset_test_semantic_scene_reads();
        let editor = SourceBackedPresentationEditor::from_read_at(Arc::new(CountingSource::new(
            source_backed_pptx(),
        )))
        .unwrap();
        let commit = changed_slide_commit(&editor, 0, "First slide", "Updated slide");
        assert_eq!(test_semantic_scene_reads(), 1);
        assert_eq!(test_semantic_slide_part_reads(), 1);

        let mut output = Vec::new();
        editor
            .publish_slide_commit_to_stream(&mut output, &commit)
            .unwrap();
        assert_eq!(test_semantic_scene_reads(), 1);
        assert_eq!(test_semantic_slide_part_reads(), 1);
    }

    #[test]
    fn changed_multi_slide_batch_publication_reuses_semantic_snapshots() {
        reset_test_semantic_scene_reads();
        let editor = SourceBackedPresentationEditor::from_read_at(Arc::new(CountingSource::new(
            source_backed_pptx(),
        )))
        .unwrap();
        let commit = changed_slide_batch_commit(&editor);
        assert_eq!(test_semantic_scene_reads(), 2);
        assert_eq!(test_semantic_slide_part_reads(), 2);

        let mut output = Vec::new();
        editor
            .publish_slide_batch_commit_to_stream(&mut output, &commit)
            .unwrap();
        assert_eq!(test_semantic_scene_reads(), 2);
        assert_eq!(test_semantic_slide_part_reads(), 2);
    }

    #[test]
    fn batch_replay_preserves_unchanged_recipient_slide_policy() {
        let editor = SourceBackedPresentationEditor::from_read_at(Arc::new(CountingSource::new(
            source_backed_pptx(),
        )))
        .unwrap();
        let (before_first, after_first) =
            changed_slide_snapshots(&editor, 0, "First slide", "Updated slide");
        let before_second = editor.slide_snapshot(1).unwrap();
        let defaults = crate::shape::Limits::DEFAULT;
        let tight_shape = crate::shape::Limits::new(
            before_second.xml.as_bytes().len().saturating_sub(1),
            defaults.output_bytes(),
            defaults.depth(),
            defaults.nodes(),
            defaults.shapes(),
            defaults.retained_text_bytes(),
        )
        .unwrap();
        let tight_options = crate::OpenOptions::default().shape_limits(tight_shape);
        let mut tight_second = before_second.clone();
        tight_second.semantic_limits = tight_options.semantic_limits.clone();
        let unchanged_second = before_second.clone();
        let patch = SourceBackedSlideBatchPatch {
            before: SourceBackedSlideBatchSnapshot {
                slides: vec![before_first.clone(), before_second].into_boxed_slice(),
            },
            after: SourceBackedSlideBatchSnapshot {
                slides: vec![after_first, unchanged_second].into_boxed_slice(),
            },
        };
        let recipient = SourceBackedSlideBatchSnapshot {
            slides: vec![before_first, tight_second.clone()].into_boxed_slice(),
        };

        let applied = patch
            .apply(&recipient)
            .expect("unchanged recipient slides are not newly admitted");
        assert_eq!(
            applied.slides[1].semantic_limits,
            tight_second.semantic_limits
        );
        assert_eq!(
            applied.slides[0].xml.as_bytes(),
            patch.after.slides[0].xml.as_bytes()
        );
    }

    #[test]
    fn source_laser_usage_is_lazy_and_cached_for_same_policy_replay() {
        reset_source_laser_scans();
        let editor = SourceBackedPresentationEditor::from_read_at(Arc::new(CountingSource::new(
            source_backed_pptx(),
        )))
        .unwrap();
        let (before, after) = changed_slide_snapshots(&editor, 0, "First slide", "Updated slide");
        assert_eq!(source_laser_scans(), 0);

        let patch = SourceBackedSlidePatch {
            before: before.clone(),
            after,
        };
        let first = patch.apply(&before).unwrap();
        assert_eq!(source_laser_scans(), 1);
        drop(first);
        let second = patch.apply(&before).unwrap();
        assert_eq!(source_laser_scans(), 1);
        drop(second);
    }

    #[test]
    fn source_laser_usage_caches_different_policy_result_after_admission() {
        reset_source_laser_scans();
        let editor = SourceBackedPresentationEditor::from_read_at(Arc::new(CountingSource::new(
            source_backed_pptx(),
        )))
        .unwrap();
        let (before, after) = changed_slide_snapshots(&editor, 0, "First slide", "Updated slide");
        let defaults = crate::laser::Limits::DEFAULT;
        let profile = crate::laser::Limits::new(
            defaults.max_slide_xml_bytes(),
            defaults.max_total_slide_xml_bytes() - 1,
            defaults.max_trace_count(),
            defaults.max_point_count(),
            defaults.max_xml_nodes(),
            defaults.max_xml_depth(),
            defaults.max_output_bytes(),
        )
        .unwrap();
        let options = crate::OpenOptions::default().laser_limits(profile);
        let mut recipient = before.clone();
        recipient.semantic_limits = options.semantic_limits.clone();
        recipient.laser_limits = profile.constrained_by_source(SourceLimits::default());
        let patch = SourceBackedSlidePatch { before, after };

        let published = patch.apply(&recipient).unwrap();
        assert_eq!(source_laser_scans(), 1);
        assert_eq!(
            published.laser_usage().unwrap().total_slide_xml_bytes(),
            published.xml.as_bytes().len()
        );
        assert_eq!(source_laser_scans(), 1);
    }

    #[test]
    fn batch_replay_charges_unchanged_laser_usage_in_selected_scope() {
        reset_source_laser_scans();
        let editor = SourceBackedPresentationEditor::from_read_at(Arc::new(CountingSource::new(
            source_backed_pptx(),
        )))
        .unwrap();
        let (before_first, after_first) =
            changed_slide_snapshots(&editor, 0, "First slide", "Updated slide");
        let before_second = editor.slide_snapshot(1).unwrap();
        let total = before_first
            .xml
            .as_bytes()
            .len()
            .checked_add(before_second.xml.as_bytes().len())
            .unwrap();
        let defaults = crate::laser::Limits::DEFAULT;
        let profile = crate::laser::Limits::new(
            defaults.max_slide_xml_bytes(),
            total.saturating_sub(1),
            defaults.max_trace_count(),
            defaults.max_point_count(),
            defaults.max_xml_nodes(),
            defaults.max_xml_depth(),
            defaults.max_output_bytes(),
        )
        .unwrap();
        let options = crate::OpenOptions::default().laser_limits(profile);
        let retained_profile = profile.constrained_by_source(SourceLimits::default());
        let mut recipient_first = before_first.clone();
        recipient_first.semantic_limits = options.semantic_limits.clone();
        recipient_first.laser_limits = retained_profile;
        let mut recipient_second = before_second.clone();
        recipient_second.semantic_limits = options.semantic_limits.clone();
        recipient_second.laser_limits = retained_profile;
        let patch = SourceBackedSlideBatchPatch {
            before: SourceBackedSlideBatchSnapshot {
                slides: vec![before_first, before_second].into_boxed_slice(),
            },
            after: SourceBackedSlideBatchSnapshot {
                slides: vec![after_first, recipient_second.clone()].into_boxed_slice(),
            },
        };
        let expected_observed = patch.after.slides[0]
            .xml
            .as_bytes()
            .len()
            .checked_add(patch.after.slides[1].xml.as_bytes().len())
            .unwrap();
        let recipient = SourceBackedSlideBatchSnapshot {
            slides: vec![recipient_first, recipient_second].into_boxed_slice(),
        };

        assert!(matches!(
            patch.apply(&recipient),
            Err(Error::Limit {
                kind: crate::opened::LimitKind::Codec(crate::opened::CodecLimitKind::XmlBytes),
                observed,
                maximum,
            }) if observed == expected_observed && maximum == total - 1
        ));
        assert_eq!(source_laser_scans(), 2);
    }

    #[test]
    fn foreign_identical_source_replays_by_content_without_recapture() {
        reset_test_semantic_scene_reads();
        let archive = source_backed_pptx();
        let editor = SourceBackedPresentationEditor::from_read_at(Arc::new(CountingSource::new(
            archive.clone(),
        )))
        .unwrap();
        let commit = changed_slide_commit(&editor, 0, "First slide", "Updated slide");
        assert_eq!(test_semantic_scene_reads(), 1);
        assert_eq!(test_semantic_slide_part_reads(), 1);

        reset_test_semantic_scene_reads();
        let foreign =
            SourceBackedPresentationEditor::from_read_at(Arc::new(CountingSource::new(archive)))
                .unwrap();
        let mut output = Vec::new();
        assert!(
            foreign
                .publish_slide_commit_to_stream(&mut output, &commit)
                .is_ok()
        );
        assert_eq!(test_semantic_scene_reads(), 0);
        assert_eq!(test_semantic_slide_part_reads(), 0);
        assert!(!output.is_empty());
    }

    #[test]
    fn local_replay_matches_content_across_profiles_and_revalidates_the_recipient() {
        let editor = SourceBackedPresentationEditor::from_read_at(Arc::new(CountingSource::new(
            source_backed_pptx(),
        )))
        .unwrap();
        let source_len = editor.slide_snapshot(0).unwrap().xml.as_bytes().len();
        let replacement = "x".repeat(source_len);
        let (before, after) = changed_slide_snapshots(&editor, 0, "First slide", &replacement);
        let patch = SourceBackedSlidePatch {
            before: before.clone(),
            after: after.clone(),
        };

        let defaults = crate::shape::Limits::DEFAULT;
        let tight_shape = crate::shape::Limits::new(
            source_len,
            defaults.output_bytes(),
            defaults.depth(),
            defaults.nodes(),
            defaults.shapes(),
            defaults.retained_text_bytes(),
        )
        .unwrap();
        let tight_options = crate::OpenOptions::default().shape_limits(tight_shape);
        let mut tight_recipient = before.clone();
        tight_recipient.semantic_limits = tight_options.semantic_limits.clone();
        assert!(matches!(
            patch.apply(&tight_recipient),
            Err(Error::Limit {
                kind: crate::opened::LimitKind::Codec(crate::opened::CodecLimitKind::InputBytes),
                maximum,
                ..
            }) if maximum == source_len
        ));

        let loose_shape = crate::shape::Limits::new(
            source_len.saturating_mul(4),
            defaults.output_bytes(),
            defaults.depth(),
            defaults.nodes(),
            defaults.shapes(),
            defaults.retained_text_bytes(),
        )
        .unwrap();
        let loose_options = crate::OpenOptions::default().shape_limits(loose_shape);
        let mut loose_recipient = before.clone();
        loose_recipient.semantic_limits = loose_options.semantic_limits.clone();
        let applied = patch.apply(&loose_recipient).unwrap();
        assert_eq!(applied.xml.as_bytes(), after.xml.as_bytes());
        assert_eq!(applied.semantic_limits, loose_recipient.semantic_limits);

        let noop = SourceBackedSlidePatch {
            before: before.clone(),
            after: before,
        };
        let replayed_noop = noop.apply(&tight_recipient).unwrap();
        assert_eq!(replayed_noop.xml.as_bytes(), tight_recipient.xml.as_bytes());
        assert_eq!(
            replayed_noop.semantic_limits,
            tight_recipient.semantic_limits
        );
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
        let presentation = SourceBackedPresentation::from_read_at_with_options(
            source,
            crate::OpenOptions::default()
                .source_limits(SourceLimits::default())
                .execution_context(context),
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
        let presentation = SourceBackedPresentation::from_read_at_with_options(
            source,
            crate::OpenOptions::default()
                .source_limits(SourceLimits::default())
                .execution_context(context),
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
    fn try_cache_diagnostics_for_presentation_is_exact_and_payload_free() {
        let source = Arc::new(CountingSource::new(source_backed_pptx()));
        let (budget, _cancel, context) = source_backed_context(u64::MAX);
        let presentation = SourceBackedPresentation::from_read_at_with_options(
            source.clone(),
            crate::OpenOptions::default()
                .source_limits(SourceLimits::default())
                .execution_context(context),
        )
        .unwrap();
        let opening_reads = source.second_payload_reads.load(Ordering::SeqCst);
        let opening_io_reads = source.read_calls.load(Ordering::SeqCst);
        let opening = presentation.try_cache_diagnostics().unwrap();
        assert_eq!(
            opening,
            presentation.inner.package.try_cache_diagnostics().unwrap()
        );
        assert!(opening.budget_managed);
        assert_eq!(opening.retained_entries, 1);
        assert_eq!(
            source.second_payload_reads.load(Ordering::SeqCst),
            opening_reads
        );
        assert_eq!(source.read_calls.load(Ordering::SeqCst), opening_io_reads);

        let slide = presentation.slide(1).unwrap();
        assert_eq!(slide.text().unwrap(), "Second slide");
        let after_read = source.second_payload_reads.load(Ordering::SeqCst);
        let after_read_io = source.read_calls.load(Ordering::SeqCst);
        assert!(after_read > opening_reads);
        let selected = presentation.try_cache_diagnostics().unwrap();
        assert_eq!(
            selected,
            presentation.inner.package.try_cache_diagnostics().unwrap()
        );
        assert!(selected.successful_loads > opening.successful_loads);
        assert_eq!(
            source.second_payload_reads.load(Ordering::SeqCst),
            after_read
        );
        assert_eq!(source.read_calls.load(Ordering::SeqCst), after_read_io);

        drop(presentation);
        assert!(budget.used(Resource::Memory) > 0);
        drop(slide);
        assert_eq!(budget.used(Resource::Memory), 0);
        assert_eq!(budget.used(Resource::Objects), 0);
    }

    #[test]
    fn try_cache_diagnostics_for_editor_is_exact_and_releases_selected_budget() {
        let source = Arc::new(CountingSource::new(source_backed_pptx()));
        let (budget, _cancel, context) = source_backed_context(u64::MAX);
        let editor = SourceBackedPresentationEditor::from_read_at_with_options(
            source.clone(),
            crate::OpenOptions::default()
                .source_limits(SourceLimits::default())
                .execution_context(context),
        )
        .unwrap();
        let opening_reads = source.second_payload_reads.load(Ordering::SeqCst);
        let opening_io_reads = source.read_calls.load(Ordering::SeqCst);
        let opening = editor.try_cache_diagnostics().unwrap();
        assert_eq!(opening, editor.package.try_cache_diagnostics().unwrap());
        assert!(opening.budget_managed);
        assert_eq!(opening.retained_entries, 1);
        assert_eq!(
            source.second_payload_reads.load(Ordering::SeqCst),
            opening_reads
        );
        assert_eq!(source.read_calls.load(Ordering::SeqCst), opening_io_reads);

        let snapshot = editor.slide_snapshot(1).unwrap();
        let after_read = source.second_payload_reads.load(Ordering::SeqCst);
        let after_read_io = source.read_calls.load(Ordering::SeqCst);
        assert!(after_read > opening_reads);
        let selected = editor.try_cache_diagnostics().unwrap();
        assert_eq!(selected, editor.package.try_cache_diagnostics().unwrap());
        assert!(selected.successful_loads > opening.successful_loads);
        assert_eq!(
            source.second_payload_reads.load(Ordering::SeqCst),
            after_read
        );
        assert_eq!(source.read_calls.load(Ordering::SeqCst), after_read_io);

        drop(editor);
        drop(snapshot);
        assert_eq!(budget.used(Resource::Memory), 0);
        assert_eq!(budget.used(Resource::Objects), 0);
    }

    #[test]
    fn managed_selected_slide_payload_is_evictable_after_view_drop() {
        let source = Arc::new(CountingSource::new(source_backed_pptx()));
        let (_budget, _cancel, context) = source_backed_context(u64::MAX);
        let cache_limits = SourceCacheLimits::new(usize::MAX, 2).unwrap();
        let presentation = SourceBackedPresentation::from_read_at_with_options(
            source,
            crate::OpenOptions::default()
                .source_limits(SourceLimits::default())
                .cache_limits(cache_limits)
                .execution_context(context),
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
    fn identical_foreign_source_closure_replays_by_content() {
        let first = SourceBackedPresentationEditor::from_read_at(Arc::new(CountingSource::new(
            source_backed_pptx(),
        )))
        .unwrap();
        let second = SourceBackedPresentationEditor::from_read_at(Arc::new(CountingSource::new(
            source_backed_pptx(),
        )))
        .unwrap();
        let mut edit = first.edit_slide(0).unwrap();
        assert!(
            edit.set_slide_properties(Some("Updated slide"), None)
                .unwrap()
        );
        let patch = edit.commit_checked().unwrap().patch().clone();
        let foreign = second.slide_snapshot(0).unwrap();

        assert!(patch.apply(&foreign).is_ok());
    }

    #[test]
    fn managed_snapshot_edit_and_patch_checks_are_cooperatively_cancelled() {
        let (_budget, cancellation_source, context) = source_backed_context(u64::MAX);
        let editor = SourceBackedPresentationEditor::from_read_at_with_options(
            Arc::new(CountingSource::new(source_backed_pptx())),
            crate::OpenOptions::default()
                .source_limits(SourceLimits::default())
                .execution_context(context),
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
    fn native_guid_svg_picture_inventory_resolves_typed_extension_and_defers_payload() {
        let source = Arc::new(PictureCountingSource::new(svg_picture_pptx()));
        let presentation = SourceBackedPresentation::from_read_at(source.clone()).unwrap();
        let slide = presentation.slide(0).unwrap();
        let images = slide.images().unwrap();
        assert_eq!(images.len(), 1);
        let svg = images[0].svg().expect("SVG extension descriptor");
        assert_eq!(svg.relationship_id(), "rIdSvg");
        assert_eq!(
            svg.target().part_uri().unwrap().as_str(),
            "/ppt/media/image1.svg"
        );
        assert_eq!(svg.target().content_type(), Some("image/svg+xml"));
        assert_eq!(svg.blip().embedded().unwrap().as_str(), "rIdSvg");
        assert_eq!(source.media_payload_reads.load(Ordering::SeqCst), 0);

        let svg_image = slide.read_svg_image(0).unwrap();
        assert!(svg_image.bytes().starts_with(b"<svg "));
        assert_eq!(svg_image.descriptor().relationship_id(), "rIdSvg");
    }

    #[test]
    fn native_guid_svg_extension_rejects_malformed_svg_blip() {
        let source = Arc::new(PictureCountingSource::new(svg_picture_pptx_with_extension(
            &format!(r#"<a:ext uri="{NATIVE_SVG_EXTENSION_URI}"><asvg:svgBlip/></a:ext>"#),
        )));
        let presentation = SourceBackedPresentation::from_read_at(source).unwrap();
        let result = presentation.slide(0).unwrap().images();
        assert!(matches!(
            result,
            Err(Error::Relationship(message)) if message.contains("svgBlip is missing")
        ));
    }

    #[test]
    fn native_guid_svg_extension_rejects_duplicate_svg_blips() {
        let extension_list = format!(
            r#"<a:ext uri="{NATIVE_SVG_EXTENSION_URI}"><asvg:svgBlip r:embed="rIdSvg"/></a:ext><a:ext uri="{NATIVE_SVG_EXTENSION_URI}"><asvg:svgBlip r:embed="rIdSvg"/></a:ext>"#
        );
        let source = Arc::new(PictureCountingSource::new(svg_picture_pptx_with_extension(
            &extension_list,
        )));
        let presentation = SourceBackedPresentation::from_read_at(source).unwrap();
        let result = presentation.slide(0).unwrap().images();
        assert!(matches!(
            result,
            Err(Error::Invalid(message)) if message.contains("unsupported or duplicate svgBlip")
        ));
    }

    #[test]
    fn synthetic_svg_namespace_uri_is_opaque_without_native_guid() {
        let source = Arc::new(PictureCountingSource::new(svg_picture_pptx_with_extension(
            &format!(
                r#"<a:ext uri="{SVG_NAMESPACE_URI}"><asvg:svgBlip r:embed="rIdSvg"/></a:ext>"#
            ),
        )));
        let presentation = SourceBackedPresentation::from_read_at(source).unwrap();
        let images = presentation.slide(0).unwrap().images().unwrap();
        assert_eq!(images.len(), 1);
        assert!(images[0].svg().is_none());
    }

    #[test]
    fn opaque_extension_keeps_typed_looking_descendants_opaque() {
        let extension_list = format!(
            r#"<a:ext uri="{OPAQUE_EXTENSION_URI}"><asvg:svgBlip r:embed="rIdIgnored"><a:ext uri="{NATIVE_SVG_EXTENSION_URI}"><asvg:svgBlip r:embed="rIdDeep"/></a:ext></asvg:svgBlip></a:ext>"#
        );
        let source = Arc::new(PictureCountingSource::new(svg_picture_pptx_with_extension(
            &extension_list,
        )));
        let presentation = SourceBackedPresentation::from_read_at(source).unwrap();
        let images = presentation.slide(0).unwrap().images().unwrap();
        assert_eq!(images.len(), 1);
        assert!(images[0].svg().is_none());
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
    fn selected_picture_query_matches_all_without_descriptor_reservation() {
        let slide_xml = multiple_picture_slide(&[
            r#"<a:blip r:embed="rIdImage1"/>"#,
            r#"<a:blip r:embed="rIdImage2"/>"#,
            r#"<a:blip r:embed="rIdImage3"/>"#,
        ]);
        let relationships = picture_relationships(&["rIdImage1", "rIdImage2", "rIdImage3"]);
        let source = Arc::new(PictureCountingSource::new(picture_pptx(
            &slide_xml,
            &relationships,
            Some("image/png"),
            None,
        )));
        let presentation = SourceBackedPresentation::from_read_at(source.clone()).unwrap();
        let slide = presentation.slide(0).unwrap();

        reset_source_picture_query_counters();
        let images = slide.images().unwrap();
        assert_eq!(images.len(), 3);
        assert_eq!(source_picture_query_all(), 1);
        assert_eq!(source_picture_query_selected(), 0);
        assert_eq!(source_picture_descriptor_reserve_calls(), 1);
        assert!(source_picture_descriptor_reserve_items() >= images.len());
        assert_eq!(source_picture_target_resolutions(), 3);
        assert_eq!(images[0].shape_position(), 0);
        assert_eq!(images[1].shape_position(), 1);
        assert_eq!(images[2].shape_position(), 2);

        reset_source_picture_query_counters();
        let selected = slide.image(1).unwrap();
        assert_eq!(selected, images[1]);
        assert_eq!(source_picture_query_all(), 0);
        assert_eq!(source_picture_query_selected(), 1);
        assert_eq!(source_picture_descriptor_reserve_calls(), 0);
        assert_eq!(source_picture_descriptor_reserve_items(), 0);
        assert_eq!(source_picture_target_resolutions(), 3);

        reset_source_picture_query_counters();
        let media_reads_before = source.media_payload_reads.load(Ordering::SeqCst);
        let image = slide.read_image(1).unwrap();
        assert_eq!(image.descriptor(), &images[1]);
        assert_eq!(source_picture_query_all(), 0);
        assert_eq!(source_picture_query_selected(), 1);
        assert_eq!(source_picture_descriptor_reserve_calls(), 0);
        assert_eq!(source_picture_descriptor_reserve_items(), 0);
        assert_eq!(source_picture_target_resolutions(), 3);
        assert!(
            source.media_payload_reads.load(Ordering::SeqCst) > media_reads_before,
            "read_image must load the selected media payload after metadata validation"
        );
    }

    #[test]
    fn selected_picture_query_validates_later_grammar_and_targets() {
        let malformed_slide = multiple_picture_slide(&[
            r#"<a:blip r:embed="rIdImage1"/>"#,
            r#"<a:blip r:embed="rIdImage2"/><a:blip r:embed="rIdImage2"/>"#,
            r#"<a:blip r:embed="rIdImage3"/>"#,
        ]);
        let relationships = picture_relationships(&["rIdImage1", "rIdImage2", "rIdImage3"]);
        let source = Arc::new(PictureCountingSource::new(picture_pptx(
            &malformed_slide,
            &relationships,
            Some("image/png"),
            None,
        )));
        let presentation = SourceBackedPresentation::from_read_at(source).unwrap();
        let result = presentation.slide(0).unwrap().image(0);
        assert!(matches!(
            result,
            Err(Error::Invalid(message)) if message.contains("blipFill direct children")
        ));

        let missing_target_slide = multiple_picture_slide(&[
            r#"<a:blip r:embed="rIdImage1"/>"#,
            r#"<a:blip r:embed="rIdMissing"/>"#,
            r#"<a:blip r:embed="rIdImage3"/>"#,
        ]);
        let relationships = picture_relationships(&["rIdImage1", "rIdImage3"]);
        let source = Arc::new(PictureCountingSource::new(picture_pptx(
            &missing_target_slide,
            &relationships,
            Some("image/png"),
            None,
        )));
        let presentation = SourceBackedPresentation::from_read_at(source).unwrap();
        let result = presentation.slide(0).unwrap().image(0);
        assert!(matches!(
            result,
            Err(Error::Relationship(message)) if message.contains("'rIdMissing' is missing")
        ));
    }

    #[test]
    fn selected_picture_query_reports_final_length_and_malformed_before_oob() {
        let valid_slide = multiple_picture_slide(&[
            r#"<a:blip r:embed="rIdImage1"/>"#,
            r#"<a:blip r:embed="rIdImage2"/>"#,
            r#"<a:blip r:embed="rIdImage3"/>"#,
        ]);
        let relationships = picture_relationships(&["rIdImage1", "rIdImage2", "rIdImage3"]);
        let source = Arc::new(PictureCountingSource::new(picture_pptx(
            &valid_slide,
            &relationships,
            Some("image/png"),
            None,
        )));
        let presentation = SourceBackedPresentation::from_read_at(source).unwrap();
        assert!(matches!(
            presentation.slide(0).unwrap().image(99),
            Err(Error::IndexOutOfBounds { index: 99, len: 3 })
        ));

        let malformed_slide = multiple_picture_slide(&[
            r#"<a:blip r:embed="rIdImage1"/>"#,
            r#"<a:blip r:embed="rIdImage2"/><a:blip r:embed="rIdImage2"/>"#,
            r#"<a:blip r:embed="rIdImage3"/>"#,
        ]);
        let source = Arc::new(PictureCountingSource::new(picture_pptx(
            &malformed_slide,
            &relationships,
            Some("image/png"),
            None,
        )));
        let presentation = SourceBackedPresentation::from_read_at(source).unwrap();
        let result = presentation.slide(0).unwrap().image(99);
        assert!(matches!(
            result,
            Err(Error::Invalid(message)) if message.contains("blipFill direct children")
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
    fn foreign_markup_compatibility_attribute_is_preserved_on_blip_inventory() {
        let foreign = String::from_utf8(picture_slide(r#"<a:blip r:embed="rIdImage"/>"#))
            .unwrap()
            .replace(
                "xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\"",
                "xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" xmlns:foreign=\"http://purl.oclc.org/ooxml/markup-compatibility/2006\"",
            )
            .replace(
                r#"<a:blip r:embed="rIdImage"/>"#,
                r#"<a:blip foreign:Ignorable="keep" r:embed="rIdImage"/>"#,
            )
            .into_bytes();
        let source = Arc::new(PictureCountingSource::new(picture_pptx(
            &foreign,
            br#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdImage" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image1.png"/></Relationships>"#,
            Some("image/png"),
            None,
        )));
        let presentation = SourceBackedPresentation::from_read_at(source.clone()).unwrap();
        assert_eq!(presentation.slide(0).unwrap().images().unwrap().len(), 1);
        assert!(
            source
                .bytes
                .windows(b"foreign:Ignorable=\"keep\"".len())
                .any(|window| { window == b"foreign:Ignorable=\"keep\"" })
        );
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
        assert!(matches!(result, Err(Error::Refused(_))));
        let result = presentation.slide(0).unwrap().image(0);
        assert!(matches!(result, Err(Error::Refused(_))));

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
            slide.image(99),
            Err(Error::Opc(litchi_opc::OpcError::SourceChanged { .. }))
        ));
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
        let limits = SourceLimits::builder()
            .max_part_bytes(1024)
            .unwrap()
            .build()
            .unwrap();
        assert!(matches!(
            SourceBackedPresentation::from_read_at_with_options(
                source.clone(),
                crate::OpenOptions::default().source_limits(limits)
            ),
            Err(Error::Opc(litchi_opc::OpcError::Limit { .. }))
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
            SourceBackedPackage::from_read_at(
                source.clone(),
                SourceLimits::default(),
                SourceCacheLimits::default(),
                litchi_opc::SourceReadPolicy::exact(),
                Some(context)
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
        let package = SourceBackedPackage::from_read_at(
            source.clone(),
            SourceLimits::default(),
            SourceCacheLimits::default(),
            litchi_opc::SourceReadPolicy::exact(),
            Some(context),
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
        let limits = SourceLimits::builder()
            .max_part_bytes(1024)
            .unwrap()
            .build()
            .unwrap();

        assert!(matches!(
            SourceBackedPresentation::from_read_at_with_options(
                source.clone(),
                crate::OpenOptions::default().source_limits(limits)
            ),
            Err(Error::Opc(litchi_opc::OpcError::Limit {
                kind: SourceLimitKind::PartBytes,
                observed: 131326,
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
        let cache = SourceCacheLimits::new(archive.len(), 8).unwrap();

        assert_eq!(
            SourceBackedPresentation::open(file.path())
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
            SourceBackedPresentation::open_with_options(
                file.path(),
                crate::OpenOptions::default().source_limits(SourceLimits::default())
            )
            .unwrap()
            .slide_count(),
            2
        );
        assert_eq!(
            SourceBackedPresentation::open_with_options(
                file.path(),
                crate::OpenOptions::default().cache_limits(cache)
            )
            .unwrap()
            .slide_count(),
            2
        );
        assert_eq!(
            SourceBackedPresentation::open_with_options(
                file.path(),
                crate::OpenOptions::default()
                    .source_limits(SourceLimits::default())
                    .cache_limits(cache)
            )
            .unwrap()
            .slide_count(),
            2
        );
        assert_eq!(
            SourceBackedPresentation::open_with_options(
                file.path(),
                crate::OpenOptions::default()
                    .source_limits(SourceLimits::default())
                    .cache_limits(cache)
            )
            .unwrap()
            .slide_count(),
            2
        );

        let (_budget, _cancellation_source, context) = source_backed_context(u64::MAX);
        assert_eq!(
            SourceBackedPresentation::open_with_options(
                file.path(),
                crate::OpenOptions::default()
                    .source_limits(SourceLimits::default())
                    .execution_context(context)
            )
            .unwrap()
            .slide_count(),
            2
        );

        let (_budget, _cancellation_source, context) = source_backed_context(u64::MAX);
        assert_eq!(
            SourceBackedPresentation::open_with_options(
                file.path(),
                crate::OpenOptions::default()
                    .source_limits(SourceLimits::default())
                    .execution_context(context)
            )
            .unwrap()
            .slide_count(),
            2
        );

        let (_budget, _cancellation_source, context) = source_backed_context(u64::MAX);
        assert_eq!(
            SourceBackedPresentation::open_with_options(
                file.path(),
                crate::OpenOptions::default()
                    .source_limits(SourceLimits::default())
                    .cache_limits(cache)
                    .execution_context(context)
            )
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
            SourceBackedPresentationEditor::open(file.path())
                .unwrap()
                .slide_count(),
            2
        );
        assert_eq!(
            SourceBackedPresentationEditor::open_with_options(
                file.path(),
                crate::OpenOptions::default()
                    .source_limits(SourceLimits::default())
                    .cache_limits(cache)
            )
            .unwrap()
            .slide_count(),
            2
        );
        let (_budget, _cancellation_source, context) = source_backed_context(u64::MAX);
        assert_eq!(
            SourceBackedPresentationEditor::open_with_options(
                file.path(),
                crate::OpenOptions::default()
                    .source_limits(SourceLimits::default())
                    .cache_limits(cache)
                    .execution_context(context)
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
        let presentation = SourceBackedPresentation::open(file.path()).unwrap();
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
        let presentation = SourceBackedPresentation::open(file.path()).unwrap();
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
        let cache = SourceCacheLimits::new(archive.len(), 8).unwrap();
        let presentation = SourceBackedPresentation::open_with_options(
            file.path(),
            crate::OpenOptions::default()
                .source_limits(SourceLimits::default())
                .cache_limits(cache)
                .execution_context(context),
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
        let editor = SourceBackedPresentationEditor::open(file.path()).unwrap();
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

        let source = litchi_opc::OpcPackage::from_bytes(&archive, SourceLimits::default()).unwrap();
        let candidate =
            litchi_opc::OpcPackage::from_bytes(&output, SourceLimits::default()).unwrap();
        for part in source
            .try_iter_parts()
            .map(|part| part.expect("part payload decodes"))
        {
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
            SourceBackedPresentation::open(&missing),
            Err(Error::Io(error)) if error.kind() == io::ErrorKind::NotFound
        ));
        assert!(matches!(
            SourceBackedPresentationEditor::open(directory.path()),
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
        let limits = SourceLimits::builder()
            .max_part_bytes(1024)
            .unwrap()
            .build()
            .unwrap();
        assert!(matches!(
            SourceBackedPresentation::open_with_options(
                file.path(),
                crate::OpenOptions::default().source_limits(limits)
            ),
            Err(Error::Opc(litchi_opc::OpcError::Limit { .. }))
        ));

        let editor = SourceBackedPresentationEditor::open(file.path()).unwrap();
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
        let presentation = SourceBackedPresentation::open(&path).unwrap();
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
        assert!(SourceBackedPresentation::open(&path).is_err());
    }
}

#[cfg(test)]
#[allow(
    clippy::unwrap_used,
    clippy::expect_used,
    reason = "test assertions panic on failure by design"
)]
mod picture_fragment_tests {
    use crate::Error;
    use crate::shape::{Scene, Shape};
    use litchi_opc::SourceLimits;

    /// Parse one standalone picture descriptor under the output bound a
    /// source-backed editor passes: the part-byte ceiling of the default read
    /// policy.
    fn parse_picture_relationship(xml: &[u8]) -> crate::Result<super::PictureRelationship> {
        super::parse_picture_relationship_with_limit(xml, SourceLimits::default().max_part_bytes())
    }

    const PML: &str = "http://schemas.openxmlformats.org/presentationml/2006/main";
    const DML: &str = "http://schemas.openxmlformats.org/drawingml/2006/main";
    const REL: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    const MCE: &str = "http://schemas.openxmlformats.org/markup-compatibility/2006";
    const P14: &str = "http://schemas.microsoft.com/office/powerpoint/2010/main";

    /// One markup-compatibility slide holding a single picture, with the
    /// producer's namespace aliases chosen by the caller so that a deck
    /// outside the conventional `p`/`a`/`r` spelling is covered too.
    fn slide(p: &str, a: &str, r: &str) -> String {
        format!(
            r#"<{p}:sld xmlns:{a}="{DML}" xmlns:{r}="{REL}" xmlns:{p}="{PML}" xmlns:mc="{MCE}" xmlns:p14="{P14}" mc:Ignorable="p14">
  <{p}:cSld><{p}:spTree>
    <{p}:nvGrpSpPr/><{p}:grpSpPr/>
    <{p}:pic>
      <{p}:nvPicPr><{p}:cNvPr id="4" name="Photo" p14:dummy="retained"/></{p}:nvPicPr>
      <{p}:blipFill><{a}:blip {r}:embed="rId3"/><{a}:stretch><{a}:fillRect/></{a}:stretch></{p}:blipFill>
      <{p}:spPr/>
    </{p}:pic>
  </{p}:spTree></{p}:cSld>
</{p}:sld>"#
        )
    }

    #[test]
    fn self_contained_picture_descriptors_parse_for_any_producer_prefix() {
        for (p, a, r) in [("p", "a", "r"), ("q", "d", "rel")] {
            let xml = slide(p, a, r);
            let scene = Scene::read(xml.as_bytes(), crate::shape::Limits::DEFAULT).unwrap();
            assert!(scene.is_rewritten(), "{p}: the MCE input must be rewritten");
            let shape = scene.at(0).unwrap();
            assert!(matches!(shape, Shape::Picture(_)));
            let descriptor = shape.self_contained_xml().unwrap();
            let relationship = parse_picture_relationship(descriptor.as_ref()).unwrap();
            assert_eq!(relationship.id, "rId3", "{p}: embedded relationship");
            assert!(!relationship.external, "{p}: embedded, not linked");
        }
    }

    #[test]
    fn a_borrowed_span_does_not_re_declare_an_unconventional_ancestor_prefix() {
        let xml = slide("q", "d", "rel");
        let scene = Scene::read(xml.as_bytes(), crate::shape::Limits::DEFAULT).unwrap();
        let borrowed = scene.at(0).unwrap().xml().unwrap();
        assert!(
            !borrowed
                .windows(PML.len())
                .any(|window| window == PML.as_bytes()),
            "the borrowed span inherits its bindings instead of repeating them"
        );
        match parse_picture_relationship(borrowed) {
            Err(Error::Invalid(message)) => assert_eq!(
                message, "picture descriptor XML does not have a p:pic root",
                "unexpected refusal text"
            ),
            Err(other) => panic!("unexpected refusal: {other}"),
            Ok(_) => panic!("a span without its bindings must not parse standalone"),
        }
    }
}
