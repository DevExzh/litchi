//! Source-backed attach/detach of SVG resources on existing main-story
//! WordprocessingML pictures.
//!
//! The lifecycle is deliberately small.  It selects a direct picture by
//! source order, keeps the raster fallback byte-for-byte, and changes only the
//! admitted SVG extension plus its OPC dependency closure.  The drawing source
//! scanner owns expanded-name, namespace, and MCE handling; this module owns
//! the story transaction and package graph.

use std::{mem::size_of, sync::Arc};

use litchi_core::{Reservation, Resource, SourceVersion};
use litchi_drawingml::svg_blip;
use litchi_opc::constants::{content_type as ct, relationship_type as rt};
use litchi_opc::{
    AuthoredXmlFragment, EffectiveTopology, PackURI, PartData, PartView, SourceArtifact,
    SourceArtifactFingerprint, SourceTopologyPlan, TargetMode,
};
use sha2::{Digest as _, Sha256};

use crate::drawing::source::{
    ByteRange, DrawingPlacement, PictureSource, RelationshipDialect, ScanLimits, SourceDrawing,
    SvgOwnerState,
};
use crate::error::{Error, Result};
use crate::source_backed::StorySelector;

const TRANSITIONAL_RELATIONSHIP_NAMESPACE: &str = svg_blip::RELATIONSHIP_NAMESPACE;
const MAX_GENERATED_NAME_ATTEMPTS: usize = 100_000;

/// A semantic source-order selector for one direct main-story picture.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct PictureSelector {
    /// The single source drawing inventory used by the main story.
    pub drawing: usize,
    /// Zero-based direct-picture ordinal in the story.
    pub picture: usize,
}

impl PictureSelector {
    /// Construct a selector.  `drawing` is retained in the public shape so
    /// the selector can be extended to subsidiary story inventories later.
    #[must_use]
    pub const fn new(drawing: usize, picture: usize) -> Self {
        Self { drawing, picture }
    }
}

impl From<(usize, usize)> for PictureSelector {
    fn from((drawing, picture): (usize, usize)) -> Self {
        Self::new(drawing, picture)
    }
}

impl From<usize> for PictureSelector {
    fn from(picture: usize) -> Self {
        Self::new(0, picture)
    }
}

/// Borrowed opaque SVG media input.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct SvgInput<'a>(&'a [u8]);

impl<'a> SvgInput<'a> {
    /// Borrow an opaque SVG payload for preflight and attachment.
    #[must_use]
    pub const fn borrowed(bytes: &'a [u8]) -> Self {
        Self(bytes)
    }

    /// Borrow the input bytes.
    #[must_use]
    pub const fn as_bytes(self) -> &'a [u8] {
        self.0
    }
}

impl<'a> From<&'a [u8]> for SvgInput<'a> {
    fn from(bytes: &'a [u8]) -> Self {
        Self::borrowed(bytes)
    }
}

impl<'a> From<&'a Vec<u8>> for SvgInput<'a> {
    fn from(bytes: &'a Vec<u8>) -> Self {
        Self::borrowed(bytes.as_slice())
    }
}

impl AsRef<[u8]> for SvgInput<'_> {
    fn as_ref(&self) -> &[u8] {
        self.0
    }
}

/// A caller-owned advanced attach request.  Ordinary callers should use
/// [`SvgInput`] so relationship IDs and Part names remain package-owned.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct SourceBackedSvgAttachmentReplacement {
    svg: Arc<Vec<u8>>,
    part_uri: Option<PackURI>,
    relationship_id: Option<String>,
}

impl SourceBackedSvgAttachmentReplacement {
    /// Construct an ordinary request with package-allocated identities.
    #[must_use]
    pub fn new(bytes: impl Into<Vec<u8>>) -> Self {
        Self {
            svg: Arc::new(bytes.into()),
            part_uri: None,
            relationship_id: None,
        }
    }

    /// Construct an ordinary request after rejecting an empty payload.
    ///
    /// `new` remains the infallible compatibility constructor; callers that
    /// want ownership-consuming validation before retaining their bytes can
    /// use this fallible form. Package-specific size limits are still checked
    /// by the edit before the request is staged.
    pub fn try_new(bytes: impl Into<Vec<u8>>) -> Result<Self> {
        let request = Self::new(bytes);
        validate_svg_payload_length(request.svg.len())?;
        Ok(request)
    }

    /// Set an explicit SVG media Part URI for advanced callers.
    #[must_use]
    pub fn with_part_uri(mut self, part_uri: PackURI) -> Self {
        self.part_uri = Some(part_uri);
        self
    }

    /// Fallible explicit-Part setter for callers that want URI validation
    /// before retaining an advanced request.
    pub fn try_with_part_uri(mut self, part_uri: PackURI) -> Result<Self> {
        validate_svg_part_uri(&part_uri)?;
        self.part_uri = Some(part_uri);
        Ok(self)
    }

    /// Set an explicit owning-story relationship ID for advanced callers.
    #[must_use]
    pub fn with_relationship_id(mut self, relationship_id: impl Into<String>) -> Self {
        self.relationship_id = Some(relationship_id.into());
        self
    }

    /// Fallible explicit-relationship setter for callers that want NCName
    /// validation before retaining an advanced request.
    pub fn try_with_relationship_id(mut self, relationship_id: impl Into<String>) -> Result<Self> {
        let relationship_id = relationship_id.into();
        validate_relationship_id(&relationship_id)?;
        self.relationship_id = Some(relationship_id);
        Ok(self)
    }

    /// Borrow the opaque SVG bytes.
    #[must_use]
    pub fn svg(&self) -> &[u8] {
        self.svg.as_slice()
    }

    /// Borrow the requested Part URI, if any.
    #[must_use]
    pub const fn part_uri(&self) -> Option<&PackURI> {
        self.part_uri.as_ref()
    }

    /// Borrow the requested relationship ID, if any.
    #[must_use]
    pub fn relationship_id(&self) -> Option<&str> {
        self.relationship_id.as_deref()
    }
}

/// An embedded SVG resource in a source-backed view or snapshot.
#[derive(Clone)]
pub struct SourceSvgAttachment {
    relationship_id: String,
    relationship_type: String,
    part_uri: PackURI,
    payload: SvgPayload,
}

impl SourceSvgAttachment {
    /// The owning-story relationship ID.
    #[must_use]
    pub fn relationship_id(&self) -> &str {
        &self.relationship_id
    }

    /// The physical image relationship type.
    #[must_use]
    pub fn relationship_type(&self) -> &str {
        &self.relationship_type
    }

    /// The internal SVG media Part URI.
    #[must_use]
    pub const fn part_uri(&self) -> &PackURI {
        &self.part_uri
    }

    /// Borrow the exact source or staged payload.
    #[must_use]
    pub fn bytes(&self) -> &[u8] {
        self.payload.as_bytes()
    }
}

/// A read-only SVG resource projection.
#[derive(Clone)]
pub struct SvgResourceView {
    attachment: SourceSvgAttachment,
}

/// The raster fallback paired with one source-backed picture view.
///
/// The payload remains a managed [`PartData`] handle, so callers can borrow
/// the exact PNG bytes without copying the media member into the view.
#[derive(Clone)]
pub struct RasterResourceView {
    relationship_id: String,
    relationship_type: String,
    part_uri: PackURI,
    content_type: String,
    payload: PartData,
}

impl RasterResourceView {
    /// The owning-story raster relationship ID.
    #[must_use]
    pub fn relationship_id(&self) -> &str {
        &self.relationship_id
    }

    /// The physical raster relationship type.
    #[must_use]
    pub fn relationship_type(&self) -> &str {
        &self.relationship_type
    }

    /// The internal raster media Part URI.
    #[must_use]
    pub const fn part_uri(&self) -> &PackURI {
        &self.part_uri
    }

    /// The declared raster content type.
    #[must_use]
    pub fn content_type(&self) -> &str {
        &self.content_type
    }

    /// Borrow the exact source raster bytes.
    #[must_use]
    pub fn bytes(&self) -> &[u8] {
        self.payload.as_bytes()
    }
}

impl SvgResourceView {
    /// Relationship ID carried by the `svgBlip`.
    #[must_use]
    pub fn relationship_id(&self) -> &str {
        self.attachment.relationship_id()
    }

    /// Physical image relationship type.
    #[must_use]
    pub fn relationship_type(&self) -> &str {
        self.attachment.relationship_type()
    }

    /// Internal SVG media Part URI.
    #[must_use]
    pub const fn part_uri(&self) -> &PackURI {
        self.attachment.part_uri()
    }

    /// Borrow the opaque SVG payload.
    #[must_use]
    pub fn bytes(&self) -> &[u8] {
        self.attachment.bytes()
    }
}

/// A metadata-only SVG media projection.  The payload is read only when
/// [`Self::data`] is requested, so enumerating picture owners does not load
/// unrelated or unselected media members.
#[derive(Clone, Copy)]
pub struct SvgResourceSourceView<'a> {
    relationship_id: &'a str,
    relationship_type: &'a str,
    part_uri: &'a PackURI,
    part: PartView<'a>,
}

impl SvgResourceSourceView<'_> {
    /// The owning-story relationship ID.
    #[must_use]
    pub fn relationship_id(&self) -> &str {
        self.relationship_id
    }

    /// The physical image relationship type.
    #[must_use]
    pub fn relationship_type(&self) -> &str {
        self.relationship_type
    }

    /// The internal SVG media Part URI.
    #[must_use]
    pub const fn part_uri(&self) -> &PackURI {
        self.part_uri
    }

    /// Lazily read the exact SVG media payload.
    pub fn data(&self) -> Result<PartData> {
        self.part.data().map_err(Error::from)
    }
}

/// A metadata-only raster fallback projection paired with one source picture.
#[derive(Clone, Copy)]
pub struct RasterResourceSourceView<'a> {
    relationship_id: &'a str,
    relationship_type: &'a str,
    part_uri: &'a PackURI,
    content_type: &'a str,
    part: PartView<'a>,
}

impl RasterResourceSourceView<'_> {
    /// The owning-story raster relationship ID.
    #[must_use]
    pub fn relationship_id(&self) -> &str {
        self.relationship_id
    }

    /// The physical raster relationship type.
    #[must_use]
    pub fn relationship_type(&self) -> &str {
        self.relationship_type
    }

    /// The internal raster media Part URI.
    #[must_use]
    pub const fn part_uri(&self) -> &PackURI {
        self.part_uri
    }

    /// The declared raster content type.
    #[must_use]
    pub fn content_type(&self) -> &str {
        self.content_type
    }

    /// Lazily read the exact raster payload.
    pub fn data(&self) -> Result<PartData> {
        self.part.data().map_err(Error::from)
    }
}

/// Borrowed metadata view of one direct main-story picture.
///
/// This view keeps source-backed Part handles rather than retaining decoded
/// media bytes. Use [`Self::raster`] and [`Self::svg`] for graph metadata, and
/// their `data` methods only when the corresponding payload is needed.
#[derive(Clone, Copy)]
pub struct SourceSvgPictureSourceView<'a> {
    selector: PictureSelector,
    placement: DrawingPlacement,
    owner_state: SvgPictureOwnerState,
    raster: RasterResourceSourceView<'a>,
    svg: Option<SvgResourceSourceView<'a>>,
}

impl SourceSvgPictureSourceView<'_> {
    /// Semantic source selector.
    #[must_use]
    pub const fn selector(&self) -> PictureSelector {
        self.selector
    }

    /// Inline or floating placement.
    pub const fn placement(&self) -> DrawingPlacement {
        self.placement
    }

    /// Typed recognized-owner status.
    #[must_use]
    pub const fn owner_state(&self) -> SvgPictureOwnerState {
        self.owner_state
    }

    /// Metadata-only raster fallback resource.
    #[must_use]
    pub const fn raster(&self) -> &RasterResourceSourceView<'_> {
        &self.raster
    }

    /// Metadata-only SVG resource, if one is present and graph-valid.
    #[must_use]
    pub const fn svg(&self) -> Option<&SvgResourceSourceView<'_>> {
        self.svg.as_ref()
    }
}

/// Typed owner status for a source picture.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[non_exhaustive]
pub enum SvgPictureOwnerState {
    /// No recognized SVG extension exists.
    None,
    /// One embedded owner exists and is graph-valid.
    Embedded,
    /// A linked owner exists and is inert.
    Linked,
    /// More than one recognized owner exists.
    Ambiguous,
    /// The recognized owner is malformed or otherwise refused.
    Refused,
    /// Unknown extension data is present but remains opaque.
    Opaque,
}

/// One direct main-story picture projection.
#[derive(Clone)]
pub struct SourceSvgPictureView {
    selector: PictureSelector,
    placement: DrawingPlacement,
    raster_relationship_id: String,
    raster_part_uri: PackURI,
    owner_state: SvgPictureOwnerState,
    raster: RasterResourceView,
    svg: Option<SvgResourceView>,
}

/// Short alias matching the design vocabulary for a picture view.
pub type SvgPictureView = SourceSvgPictureView;

impl SourceSvgPictureView {
    /// Semantic source selector.
    #[must_use]
    pub const fn selector(&self) -> PictureSelector {
        self.selector
    }

    /// Whether the picture is inline or floating.
    pub const fn placement(&self) -> DrawingPlacement {
        self.placement
    }

    /// Raster fallback relationship ID.
    #[must_use]
    pub fn raster_relationship_id(&self) -> &str {
        &self.raster_relationship_id
    }

    /// Raster fallback Part URI.
    #[must_use]
    pub const fn raster_part_uri(&self) -> &PackURI {
        &self.raster_part_uri
    }

    /// Raster fallback content type.
    #[must_use]
    pub fn raster_content_type(&self) -> &str {
        self.raster.content_type()
    }

    /// The borrowed raster fallback resource and its exact payload.
    #[must_use]
    pub const fn raster(&self) -> &RasterResourceView {
        &self.raster
    }

    /// Typed recognized-owner status.
    #[must_use]
    pub const fn owner_state(&self) -> SvgPictureOwnerState {
        self.owner_state
    }

    /// Embedded SVG resource, if one is present and graph-valid.
    #[must_use]
    pub const fn svg(&self) -> Option<&SvgResourceView> {
        self.svg.as_ref()
    }
}

#[derive(Clone)]
enum StoryPayload {
    Original(litchi_opc::SourceXmlPart),
    Edited(litchi_opc::SourceXmlPart),
    /// A projected intermediate story.  It is intentionally kept as raw
    /// bytes: OPC source-publication provenance permits only one source
    /// splice, so intermediate batch projections must not try to edit a
    /// derived `SourceXmlPart` a second time.
    Projected(Arc<Vec<u8>>),
}

impl StoryPayload {
    fn as_bytes(&self) -> &[u8] {
        match self {
            Self::Original(value) | Self::Edited(value) => value.bytes(),
            Self::Projected(value) => value.as_slice(),
        }
    }

    fn source_xml(&self) -> &litchi_opc::SourceXmlPart {
        match self {
            Self::Original(value) | Self::Edited(value) => value,
            Self::Projected(_) => {
                unreachable!("projected story bytes have no OPC source proof")
            },
        }
    }
}

#[derive(Clone)]
enum SvgPayload {
    Original(PartData),
    Edited(Arc<Vec<u8>>),
}

impl SvgPayload {
    fn as_bytes(&self) -> &[u8] {
        match self {
            Self::Original(value) => value.as_bytes(),
            Self::Edited(value) => value.as_slice(),
        }
    }
}

#[derive(Clone)]
struct AttachmentState {
    relationship_id: String,
    relationship_type: String,
    part_uri: PackURI,
    payload: SvgPayload,
    // Pins the staged payload's budget for its lifetime; not semantic state.
    _reservation: Option<Arc<Reservation>>,
}

impl AttachmentState {
    fn public(&self) -> SourceSvgAttachment {
        SourceSvgAttachment {
            relationship_id: self.relationship_id.clone(),
            relationship_type: self.relationship_type.clone(),
            part_uri: self.part_uri.clone(),
            payload: self.payload.clone(),
        }
    }
}

#[derive(Clone)]
struct SnapshotState {
    part_uri: PackURI,
    xml: StoryPayload,
    selector: PictureSelector,
    placement: DrawingPlacement,
    raster_relationship_id: String,
    raster_relationship_type: String,
    raster_part_uri: PackURI,
    raster_content_type: String,
    raster_payload: PartData,
    svg: Option<AttachmentState>,
    owner_state: SvgPictureOwnerState,
    relationship_dialect: RelationshipDialect,
    existing_relationship_ids: Arc<[String]>,
    physical_member_names: Arc<[String]>,
    relationship_fingerprint: Arc<[RelationshipFingerprint]>,
    content_type_fingerprint: Arc<[ContentTypeFingerprint]>,
    package_budget: Arc<PackageBudget>,
    limits: litchi_opc::ReadLimits,
    lineage: litchi_opc::SourceLineage,
    source_version: SourceVersion,
    /// Exact immutable package source retained for physical inverse tooling.
    /// The handle is O(1); its relationship and content-type members are
    /// copied only if a caller explicitly writes the artifact.
    source_artifact: SourceArtifact,
}

/// Metadata-only package capacities retained by every selected state.
///
/// Keeping this inventory separate from payload reads lets a batch reject an
/// attachment against aggregate package limits before it copies caller bytes.
#[derive(Clone)]
struct PackageBudget {
    part_count: usize,
    total_part_bytes: u64,
    relationship_count: usize,
    physical_member_count: usize,
    content_type_mapping_lower_bound: usize,
}

#[derive(Clone, PartialEq, Eq)]
struct RelationshipFingerprint {
    owner: PackURI,
    id: String,
    relationship_type: String,
    target: Option<PackURI>,
    target_mode: TargetMode,
}

#[derive(Clone, PartialEq, Eq)]
struct ContentTypeFingerprint {
    part_uri: PackURI,
    content_type: String,
}

/// The immutable byte layout needed to project one picture operation.
///
/// This is deliberately detached from `PictureSource<'a>`: the scanner's
/// borrowed records cannot outlive the source slice returned by a scan, while
/// a batch edit must reuse the same ranges for every staged operation.  The
/// namespace prefixes are the only source bytes copied into this compact
/// layout; the picture and extension payloads remain in `source_xml`.
#[derive(Clone)]
struct PictureLayout {
    blip_range: ByteRange,
    blip_prefix: Box<[u8]>,
    blip_close_start: Option<usize>,
    ext_list_range: Option<ByteRange>,
    ext_list_prefix: Option<Box<[u8]>>,
    ext_list_close_start: Option<usize>,
    svg_extension_range: Option<ByteRange>,
    owner_state: SvgPictureOwnerState,
    post_detach_owner_state: SvgPictureOwnerState,
}

/// Shared source facts for the lifetime of a batch edit.
///
/// Every selected picture points into one immutable source scan.  Staging
/// looks up these compact layouts instead of reparsing the whole projected
/// story for each operation.
#[derive(Clone)]
struct BatchSourceContext {
    source_xml: litchi_opc::SourceXmlPart,
    source_version: SourceVersion,
    lineage: litchi_opc::SourceLineage,
    layouts: Arc<[PictureLayout]>,
}

#[derive(Clone)]
struct ProjectionEdit {
    range: ByteRange,
    replacement_len: usize,
}

/// Immutable source-bound state for one selected picture.
#[derive(Clone)]
pub struct SourceBackedSvgAttachmentSnapshot {
    state: SnapshotState,
}

impl SourceBackedSvgAttachmentSnapshot {
    /// Selected semantic picture.
    #[must_use]
    pub const fn selector(&self) -> PictureSelector {
        self.state.selector
    }

    /// Inline or floating placement.
    pub const fn placement(&self) -> DrawingPlacement {
        self.state.placement
    }

    /// Existing raster relationship ID.
    #[must_use]
    pub fn raster_relationship_id(&self) -> &str {
        &self.state.raster_relationship_id
    }

    /// Existing raster relationship type.
    #[must_use]
    pub fn raster_relationship_type(&self) -> &str {
        &self.state.raster_relationship_type
    }

    /// Existing raster fallback Part URI.
    #[must_use]
    pub const fn raster_part_uri(&self) -> &PackURI {
        &self.state.raster_part_uri
    }

    /// Return the exact raster fallback payload retained by this snapshot.
    #[must_use]
    pub fn raster(&self) -> RasterResourceView {
        RasterResourceView {
            relationship_id: self.state.raster_relationship_id.clone(),
            relationship_type: self.state.raster_relationship_type.clone(),
            part_uri: self.state.raster_part_uri.clone(),
            content_type: self.state.raster_content_type.clone(),
            payload: self.state.raster_payload.clone(),
        }
    }

    /// Current optional SVG attachment.
    #[must_use]
    pub fn svg(&self) -> Option<SourceSvgAttachment> {
        self.state.svg.as_ref().map(AttachmentState::public)
    }

    /// Current typed owner status.
    #[must_use]
    pub const fn owner_state(&self) -> SvgPictureOwnerState {
        self.state.owner_state
    }

    /// Exact source XML bytes for the story state.
    #[must_use]
    pub fn story_xml(&self) -> &[u8] {
        self.state.xml.as_bytes()
    }

    /// Retain the exact source package, including raw relationships and
    /// `[Content_Types].xml`, for an explicit physical inverse operation.
    #[must_use]
    pub fn source_artifact(&self) -> SourceArtifact {
        self.state.source_artifact.clone()
    }
}

/// A source-backed batch of selected main-story SVG picture edits.
///
/// The story is scanned once when the batch is captured. Each staged
/// operation is projected against the previous source-checked projection, so
/// a batch does not replay earlier intents or retain a rewritten copy for
/// every intermediate state. All selected pictures are published through one
/// XML replacement and one graph closure. A failed operation leaves the edit
/// unchanged.
pub struct SourceBackedSvgAttachmentBatchEdit<'a> {
    package: &'a super::Package,
    source: SourceBackedSvgAttachmentBatchSnapshot,
    current: SourceBackedSvgAttachmentBatchSnapshot,
    context: Arc<BatchSourceContext>,
    operations: Vec<BatchOperation>,
    projection_edits: Vec<ProjectionEdit>,
    staged_attach_count: usize,
    staged_payload_bytes: u64,
    next_relationship_index: usize,
    next_part_index: usize,
}

/// An exact source-bound snapshot of a selected picture batch.
#[derive(Clone)]
pub struct SourceBackedSvgAttachmentBatchSnapshot {
    states: Vec<SnapshotState>,
    story: StoryPayload,
    context: Option<Arc<BatchSourceContext>>,
}

impl SourceBackedSvgAttachmentBatchSnapshot {
    /// Number of selected pictures in source order.
    #[must_use]
    pub fn len(&self) -> usize {
        self.states.len()
    }

    /// Whether this batch has no selected pictures.
    #[must_use]
    pub fn is_empty(&self) -> bool {
        self.states.is_empty()
    }

    /// Borrow one selected picture snapshot by its semantic selector.
    #[must_use]
    pub fn picture(
        &self,
        selector: impl Into<PictureSelector>,
    ) -> Option<SourceBackedSvgAttachmentSnapshot> {
        let selector = selector.into();
        self.states
            .iter()
            .find(|state| state.selector == selector)
            .cloned()
            .map(|mut state| {
                state.xml = self.story.clone();
                SourceBackedSvgAttachmentSnapshot { state }
            })
    }

    /// Exact staged story bytes shared by every selected picture.
    #[must_use]
    pub fn story_xml(&self) -> Option<&[u8]> {
        self.states.first().map(|_| self.story.as_bytes())
    }

    fn state(&self, selector: PictureSelector) -> Result<&SnapshotState> {
        self.states
            .iter()
            .find(|state| state.selector == selector)
            .ok_or_else(|| Error::Invalid("picture is not selected in this SVG batch".into()))
    }

    fn state_mut(&mut self, selector: PictureSelector) -> Result<&mut SnapshotState> {
        self.states
            .iter_mut()
            .find(|state| state.selector == selector)
            .ok_or_else(|| Error::Invalid("picture is not selected in this SVG batch".into()))
    }
}

/// A source-backed one-picture edit. It is a compatibility wrapper over the
/// batch implementation and therefore retains the same atomic projection.
pub struct SourceBackedSvgAttachmentEdit<'a> {
    batch: SourceBackedSvgAttachmentBatchEdit<'a>,
    selector: PictureSelector,
}

#[derive(Clone)]
enum BatchOperation {
    Attach {
        selector: PictureSelector,
        relationship_id: String,
        part_uri: PackURI,
        relationship_type: String,
        payload: Arc<Vec<u8>>,
        reservation: Option<Arc<Reservation>>,
    },
    Detach {
        selector: PictureSelector,
    },
}

impl BatchOperation {
    fn selector(&self) -> PictureSelector {
        match self {
            Self::Attach { selector, .. } | Self::Detach { selector, .. } => *selector,
        }
    }
}

/// Exact-source reversible attach/detach patch for a selected batch.
#[derive(Clone)]
pub struct SourceBackedSvgAttachmentBatchPatch {
    before: SourceBackedSvgAttachmentBatchSnapshot,
    after: SourceBackedSvgAttachmentBatchSnapshot,
}

/// Content-free facts about a successful SVG lifecycle commit.
///
/// The diagnostics deliberately contain no story, relationship, URI, or
/// media bytes.  They remain useful for a no-op commit: the selected-picture
/// count still describes the edit scope even when no operation was staged.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct SvgAttachmentCommitDiagnostics {
    operations: usize,
    selected_pictures: usize,
    changed: bool,
}

impl SvgAttachmentCommitDiagnostics {
    const fn new(operations: usize, selected_pictures: usize, changed: bool) -> Self {
        Self {
            operations,
            selected_pictures,
            changed,
        }
    }

    /// Number of staged attach or detach operations in the commit.
    #[must_use]
    pub const fn operations(self) -> usize {
        self.operations
    }

    /// Alias spelling for callers that prefer an explicit count name.
    #[must_use]
    pub const fn operation_count(self) -> usize {
        self.operations()
    }

    /// Number of selected pictures covered by the commit.
    #[must_use]
    pub const fn selected_pictures(self) -> usize {
        self.selected_pictures
    }

    /// Alias spelling for callers that prefer an explicit count name.
    #[must_use]
    pub const fn selected_picture_count(self) -> usize {
        self.selected_pictures()
    }

    /// Whether the commit changes story or dependency state.
    #[must_use]
    pub const fn changed(self) -> bool {
        self.changed
    }
}

/// A checked SVG lifecycle batch commit ready for package publication.
pub struct SourceBackedSvgAttachmentBatchCommit {
    snapshot: SourceBackedSvgAttachmentBatchSnapshot,
    patch: SourceBackedSvgAttachmentBatchPatch,
    diagnostics: SvgAttachmentCommitDiagnostics,
}

/// Physical publication evidence for a changed SVG batch.
///
/// The publication retains the exact pre-edit source artifact and the
/// fingerprint of the emitted artifact.  A reopened package must match both
/// the fingerprint and the target semantic batch before the original bytes
/// can be restored.
pub struct SourceBackedSvgAttachmentBatchPublication {
    snapshot: SourceBackedSvgAttachmentBatchSnapshot,
    original_snapshot: SourceBackedSvgAttachmentBatchSnapshot,
    original_artifact: SourceArtifact,
    published_fingerprint: SourceArtifactFingerprint,
}

/// Exact-source reversible attach/detach patch.
#[derive(Clone)]
pub struct SourceBackedSvgAttachmentPatch {
    before: SourceBackedSvgAttachmentSnapshot,
    after: SourceBackedSvgAttachmentSnapshot,
}

/// A checked SVG lifecycle commit ready for package publication.
pub struct SourceBackedSvgAttachmentCommit {
    snapshot: SourceBackedSvgAttachmentSnapshot,
    patch: SourceBackedSvgAttachmentPatch,
    diagnostics: SvgAttachmentCommitDiagnostics,
}

/// Physical publication evidence for a changed single-picture SVG edit.
pub struct SourceBackedSvgAttachmentPublication {
    snapshot: SourceBackedSvgAttachmentSnapshot,
    original_snapshot: SourceBackedSvgAttachmentSnapshot,
    original_artifact: SourceArtifact,
    published_fingerprint: SourceArtifactFingerprint,
}

impl SourceBackedSvgAttachmentEdit<'_> {
    /// Exact immutable source snapshot used by this edit.
    #[must_use]
    pub fn source(&self) -> SourceBackedSvgAttachmentSnapshot {
        self.batch
            .source
            .picture(self.selector)
            .expect("single-picture batch always contains its selector")
    }

    /// Attach one borrowed opaque SVG payload using package-allocated IDs.
    pub fn attach_svg<'a>(&mut self, input: impl Into<SvgInput<'a>>) -> Result<bool> {
        self.batch.attach_svg(self.selector, input)
    }

    /// Attach one advanced request after all identity and payload checks.
    pub fn attach(&mut self, request: &SourceBackedSvgAttachmentReplacement) -> Result<bool> {
        self.batch.attach(self.selector, request)
    }

    /// Detach the selected embedded SVG owner while retaining the raster.
    pub fn detach_svg(&mut self) -> Result<bool> {
        self.batch.detach_svg(self.selector)
    }

    /// Validate the changed closure and freeze this edit into a commit.
    pub fn commit(self) -> Result<SourceBackedSvgAttachmentCommit> {
        let batch_commit = self.batch.commit()?;
        let before = batch_commit
            .patch
            .before
            .picture(self.selector)
            .ok_or_else(|| Error::Invalid("single-picture batch selector disappeared".into()))?;
        let after = batch_commit
            .patch
            .after
            .picture(self.selector)
            .ok_or_else(|| Error::Invalid("single-picture batch target disappeared".into()))?;
        let patch = SourceBackedSvgAttachmentPatch { before, after };
        Ok(SourceBackedSvgAttachmentCommit {
            snapshot: patch.after.clone(),
            patch,
            diagnostics: batch_commit.diagnostics(),
        })
    }

    /// Explicit alias for callers that want the fallible commit spelling.
    pub fn commit_checked(self) -> Result<SourceBackedSvgAttachmentCommit> {
        self.commit()
    }
}

impl SourceBackedSvgAttachmentBatchEdit<'_> {
    /// Borrow the immutable batch source snapshot.
    #[must_use]
    pub const fn source(&self) -> &SourceBackedSvgAttachmentBatchSnapshot {
        &self.source
    }

    /// Return the current projected batch after staged operations.
    #[must_use]
    pub const fn projected(&self) -> &SourceBackedSvgAttachmentBatchSnapshot {
        &self.current
    }

    /// Attach one borrowed SVG payload to a selected picture.
    pub fn attach_svg<'a>(
        &mut self,
        selector: impl Into<PictureSelector>,
        input: impl Into<SvgInput<'a>>,
    ) -> Result<bool> {
        let input = input.into();
        self.attach_bytes(selector.into(), input.as_bytes(), None, None, None)
    }

    /// Attach an advanced caller-owned request to a selected picture.
    pub fn attach(
        &mut self,
        selector: impl Into<PictureSelector>,
        request: &SourceBackedSvgAttachmentReplacement,
    ) -> Result<bool> {
        self.attach_bytes(
            selector.into(),
            request.svg.as_slice(),
            request.relationship_id.as_deref(),
            request.part_uri.as_ref(),
            Some(Arc::clone(&request.svg)),
        )
    }

    fn attach_bytes(
        &mut self,
        selector: PictureSelector,
        bytes: &[u8],
        requested_relationship_id: Option<&str>,
        requested_part_uri: Option<&PackURI>,
        requested_payload: Option<Arc<Vec<u8>>>,
    ) -> Result<bool> {
        self.ensure_source_current()?;
        self.check_selector_operation(selector, "attach_svg")?;
        let current = self.current.state(selector)?.clone();
        if current.svg.is_some() {
            return Err(Error::Invalid(
                "selected picture already has an embedded SVG owner".into(),
            ));
        }
        if matches!(
            current.owner_state,
            SvgPictureOwnerState::Linked
                | SvgPictureOwnerState::Ambiguous
                | SvgPictureOwnerState::Refused
        ) {
            return Err(unsafe_edit(
                "attach_svg",
                "selected picture has a refused SVG owner state",
            ));
        }
        validate_svg_payload(&current.limits, bytes.len())?;
        // Validate caller-supplied identities while they are still borrowed.
        // Allocation of the owned selector strings/URI is intentionally after
        // these checks so rejected advanced requests do not copy unadmitted
        // identifiers into the staged operation.
        if let Some(relationship_id) = requested_relationship_id {
            validate_relationship_id(relationship_id)?;
        }
        if let Some(part_uri) = requested_part_uri {
            validate_svg_part_uri(part_uri)?;
            part_uri
                .rels_uri()
                .map_err(|error| Error::Uri(error.to_string()))?;
        }
        let relationship_id = self.allocate_relationship_id(requested_relationship_id)?;
        validate_relationship_id(&relationship_id)?;
        if self.relationship_id_in_use(&relationship_id) {
            return Err(Error::Invalid(format!(
                "SVG relationship ID '{relationship_id}' is already in use"
            )));
        }
        let part_uri = self.allocate_part_uri(requested_part_uri)?;
        validate_svg_part_uri(&part_uri)?;
        let rels_uri = part_uri
            .rels_uri()
            .map_err(|error| Error::Uri(error.to_string()))?;
        if self.physical_name_in_use(part_uri.membername())
            || self.physical_name_in_use(rels_uri.membername())
        {
            return Err(Error::Invalid(format!(
                "SVG Part '{}' or its relationship member already exists",
                part_uri.as_str()
            )));
        }
        self.preflight_attach(&current, bytes.len())?;
        let reservation = reserve_payload_memory(self.package, bytes.len())?;
        let relationship_type = image_relationship_type(&current).to_owned();
        // XML projection is performed before copying a borrowed payload. A
        // rejected operation therefore cannot retain or partially stage the
        // caller's bytes.
        let mut operation = BatchOperation::Attach {
            selector,
            relationship_id,
            part_uri,
            relationship_type,
            payload: Arc::new(Vec::new()),
            reservation,
        };
        let (xml, projection) = self.stage_operation(&operation)?;
        let owned_payload = if let Some(payload) = requested_payload {
            payload
        } else {
            let mut owned_payload = Vec::new();
            owned_payload
                .try_reserve_exact(bytes.len())
                .map_err(|source| Error::Allocation {
                    resource: "DOCX SVG attachment payload",
                    source,
                })?;
            owned_payload.extend_from_slice(bytes);
            Arc::new(owned_payload)
        };
        if let BatchOperation::Attach { payload, .. } = &mut operation {
            *payload = owned_payload;
        }
        self.finish_operation(operation, xml, projection)?;
        Ok(true)
    }

    /// Detach one selected embedded SVG owner while retaining its raster.
    pub fn detach_svg(&mut self, selector: impl Into<PictureSelector>) -> Result<bool> {
        let selector = selector.into();
        self.ensure_source_current()?;
        self.check_selector_operation(selector, "detach_svg")?;
        let current = self.current.state(selector)?.clone();
        if current.svg.is_none() {
            if matches!(
                current.owner_state,
                SvgPictureOwnerState::Linked
                    | SvgPictureOwnerState::Ambiguous
                    | SvgPictureOwnerState::Refused
            ) {
                return Err(unsafe_edit(
                    "detach_svg",
                    "selected picture has a refused SVG owner state",
                ));
            }
            return Ok(false);
        }
        let operation = BatchOperation::Detach { selector };
        let (xml, projection) = self.stage_operation(&operation)?;
        self.finish_operation(operation, xml, projection)?;
        Ok(true)
    }

    /// Validate the complete changed story/dependency closure and freeze it.
    pub fn commit(mut self) -> Result<SourceBackedSvgAttachmentBatchCommit> {
        self.ensure_source_current()?;
        if !self.operations.is_empty() {
            // Intermediate projections are raw source bytes because an OPC
            // `SourceXmlPart` may issue only one source splice. Apply the
            // complete intent list once to the immutable original source for
            // the final provenance-bearing replacement.
            let final_xml = rewrite_story_xml_batch(
                self.original_xml_part()?,
                &self.operations,
                self.limits(),
            )?;
            self.current.story = StoryPayload::Edited(final_xml);
        }
        let patch = SourceBackedSvgAttachmentBatchPatch {
            before: self.source,
            after: self.current,
        };
        let diagnostics = SvgAttachmentCommitDiagnostics::new(
            self.operations.len(),
            patch.before.len(),
            patch.is_changed(),
        );
        if patch.is_changed() {
            let plan =
                build_batch_topology_plan(&self.package.package, &patch.before, &patch.after)?;
            // Validate the complete effective relationship/content-type/Part
            // closure before returning the commit. Publication repeats this
            // callback on its fresh source boundary, but a caller must never
            // receive a frozen commit whose prepared candidate was not
            // checked against the same typed graph.
            self.package
                .package
                .with_prepared_topology(plan, |candidate| {
                    Ok(validate_effective_candidate(
                        candidate,
                        &patch.before,
                        &patch.after,
                    ))
                })??;
        }
        self.package.package.check_execution()?;
        Ok(SourceBackedSvgAttachmentBatchCommit {
            snapshot: patch.after.clone(),
            patch,
            diagnostics,
        })
    }

    /// Explicit alias for callers that want the fallible commit spelling.
    pub fn commit_checked(self) -> Result<SourceBackedSvgAttachmentBatchCommit> {
        self.commit()
    }

    fn check_selector_operation(
        &self,
        selector: PictureSelector,
        operation: &'static str,
    ) -> Result<()> {
        self.current.state(selector)?;
        if self
            .operations
            .iter()
            .any(|item| item.selector() == selector)
        {
            return Err(unsafe_edit(
                operation,
                "picture was already staged in this SVG batch",
            ));
        }
        Ok(())
    }

    fn relationship_id_in_use(&self, candidate: &str) -> bool {
        self.current.states.first().is_some_and(|state| {
            state
                .existing_relationship_ids
                .iter()
                .any(|value| value == candidate)
        }) || self.operations.iter().any(|operation| {
            matches!(
                operation,
                BatchOperation::Attach {
                    relationship_id,
                    ..
                } if relationship_id == candidate
            )
        })
    }

    fn physical_name_in_use(&self, candidate: &str) -> bool {
        self.current.states.first().is_some_and(|state| {
            state
                .physical_member_names
                .iter()
                .any(|value| value.eq_ignore_ascii_case(candidate))
        }) || self.operations.iter().any(|operation| {
            matches!(operation, BatchOperation::Attach { part_uri, .. }
                    if part_uri.membername().eq_ignore_ascii_case(candidate)
                        || part_uri
                            .rels_uri()
                            .is_ok_and(|uri| uri.membername().eq_ignore_ascii_case(candidate)))
        })
    }

    fn allocate_relationship_id(&mut self, requested: Option<&str>) -> Result<String> {
        if let Some(requested) = requested {
            return Ok(requested.to_owned());
        }
        while self.next_relationship_index < MAX_GENERATED_NAME_ATTEMPTS {
            let index = self.next_relationship_index;
            self.next_relationship_index += 1;
            let candidate = if index == 0 {
                "rIdSvg".to_owned()
            } else {
                format!("rIdSvg{index}")
            };
            if !self.relationship_id_in_use(&candidate) {
                return Ok(candidate);
            }
        }
        Err(Error::Invalid(
            "could not allocate a collision-free SVG relationship ID".into(),
        ))
    }

    fn allocate_part_uri(&mut self, requested: Option<&PackURI>) -> Result<PackURI> {
        if let Some(requested) = requested {
            return Ok(requested.clone());
        }
        while self.next_part_index < MAX_GENERATED_NAME_ATTEMPTS {
            let index = self.next_part_index;
            self.next_part_index += 1;
            let name = if index == 0 {
                "image.svg".to_owned()
            } else {
                format!("image{index}.svg")
            };
            let Ok(candidate) = PackURI::new(format!("/word/media/{name}")) else {
                continue;
            };
            let Ok(rels) = candidate.rels_uri() else {
                continue;
            };
            if !self.physical_name_in_use(candidate.membername())
                && !self.physical_name_in_use(rels.membername())
            {
                return Ok(candidate);
            }
        }
        Err(Error::Invalid(
            "could not allocate a collision-free SVG media Part URI".into(),
        ))
    }

    fn original_xml_part(&self) -> Result<&litchi_opc::SourceXmlPart> {
        Ok(&self.context.source_xml)
    }

    fn projected_bytes(&self) -> Result<&[u8]> {
        Ok(self.current.story.as_bytes())
    }

    fn ensure_source_current(&self) -> Result<()> {
        self.package.package.check_execution()?;
        if self.package.package.source_version()? != self.context.source_version
            || self.package.package.source_lineage() != self.context.lineage
        {
            return Err(Error::Invalid(
                "DOCX SVG batch source changed during staging".into(),
            ));
        }
        Ok(())
    }

    fn layout(&self, selector: PictureSelector) -> Result<&PictureLayout> {
        if selector.drawing != 0 {
            return Err(Error::OutOfBounds {
                object: "DOCX main-story drawing",
                index: selector.drawing,
                len: 1,
            });
        }
        self.context
            .layouts
            .get(selector.picture)
            .ok_or_else(|| Error::OutOfBounds {
                object: "DOCX main-story picture",
                index: selector.picture,
                len: self.context.layouts.len(),
            })
    }

    fn shifted_offset(&self, offset: usize) -> Result<usize> {
        let mut shifted = offset;
        for edit in &self.projection_edits {
            if edit.range.end > offset || (edit.range.is_empty() && edit.range.start > offset) {
                continue;
            }
            let removed = edit.range.len()?;
            if edit.replacement_len >= removed {
                shifted = shifted
                    .checked_add(edit.replacement_len - removed)
                    .ok_or_else(|| {
                        Error::Invalid("DOCX projected story offset overflows".into())
                    })?;
            } else {
                shifted = shifted
                    .checked_sub(removed - edit.replacement_len)
                    .ok_or_else(|| {
                        Error::Invalid("DOCX projected story offset underflows".into())
                    })?;
            }
        }
        Ok(shifted)
    }

    fn shifted_range(&self, range: ByteRange) -> Result<ByteRange> {
        let start = self.shifted_offset(range.start)?;
        let end = self.shifted_offset(range.end)?;
        if end < start {
            return Err(Error::Invalid(
                "DOCX projected SVG range is no longer ordered".into(),
            ));
        }
        Ok(ByteRange::new(start, end))
    }

    fn stage_operation(&self, operation: &BatchOperation) -> Result<(Vec<u8>, ProjectionEdit)> {
        // The layout and namespace context are captured once from the
        // immutable source story.  Staging therefore avoids both the scan of
        // the projected story and the second scan previously used to recover
        // the selected owner state.
        let layout = self.layout(operation.selector())?;
        let (original_range, replacement) = match operation {
            BatchOperation::Attach {
                relationship_id, ..
            } => {
                if matches!(
                    layout.owner_state,
                    SvgPictureOwnerState::Linked
                        | SvgPictureOwnerState::Ambiguous
                        | SvgPictureOwnerState::Refused
                        | SvgPictureOwnerState::Embedded
                ) {
                    return Err(unsafe_edit(
                        "attach_svg",
                        "selected picture has an existing recognized SVG owner",
                    ));
                }
                generated_attachment_replacement_from_layout(
                    layout,
                    relationship_id,
                    self.context.source_xml.bytes(),
                    self.limits(),
                )?
            },
            BatchOperation::Detach { .. } => {
                let range = layout.svg_extension_range.ok_or_else(|| {
                    Error::Invalid("selected picture has no admitted embedded SVG owner".into())
                })?;
                (range, Vec::new())
            },
        };
        let projected_range = self.shifted_range(original_range)?;
        let projected = self.projected_bytes()?;
        if projected_range.end > projected.len() || projected_range.start > projected_range.end {
            return Err(Error::Invalid(
                "DOCX projected SVG replacement range is outside the story".into(),
            ));
        }
        let output_len = projected
            .len()
            .checked_sub(projected_range.len()?)
            .and_then(|value| value.checked_add(replacement.len()))
            .ok_or_else(|| Error::Invalid("DOCX story output size overflows".into()))?;
        if output_len as u64 > self.limits().max_part_bytes() {
            return Err(Error::Invalid(
                "DOCX story output exceeds the Part limit".into(),
            ));
        }
        let mut output = Vec::new();
        output
            .try_reserve_exact(output_len)
            .map_err(|source| Error::Allocation {
                resource: "DOCX projected SVG story",
                source,
            })?;
        output.extend_from_slice(&projected[..projected_range.start]);
        output.extend_from_slice(&replacement);
        output.extend_from_slice(&projected[projected_range.end..]);
        Ok((
            output,
            ProjectionEdit {
                range: original_range,
                replacement_len: replacement.len(),
            },
        ))
    }

    fn finish_operation(
        &mut self,
        operation: BatchOperation,
        xml: Vec<u8>,
        projection: ProjectionEdit,
    ) -> Result<()> {
        let selector = operation.selector();
        self.operations
            .try_reserve_exact(1)
            .map_err(|source| Error::Allocation {
                resource: "DOCX SVG batch operations",
                source,
            })?;
        self.projection_edits
            .try_reserve_exact(1)
            .map_err(|source| Error::Allocation {
                resource: "DOCX SVG batch projections",
                source,
            })?;
        let next_attach_count = self
            .staged_attach_count
            .checked_add(usize::from(matches!(
                &operation,
                BatchOperation::Attach { .. }
            )))
            .ok_or_else(|| Error::Invalid("DOCX SVG attachment count overflows".into()))?;
        let next_payload_bytes = self
            .staged_payload_bytes
            .checked_add(match &operation {
                BatchOperation::Attach { payload, .. } => payload.len() as u64,
                BatchOperation::Detach { .. } => 0,
            })
            .ok_or_else(|| Error::Invalid("DOCX SVG staged payload bytes overflow".into()))?;
        let post_detach_owner_state = matches!(&operation, BatchOperation::Detach { .. })
            .then(|| {
                self.layout(selector)
                    .map(|layout| layout.post_detach_owner_state)
            })
            .transpose()?
            .unwrap_or(SvgPictureOwnerState::None);
        let projected_xml = Arc::new(xml);
        let state = self.current.state_mut(selector)?;
        match &operation {
            BatchOperation::Attach {
                relationship_id,
                part_uri,
                relationship_type,
                payload,
                reservation,
                ..
            } => {
                state.svg = Some(AttachmentState {
                    relationship_id: relationship_id.clone(),
                    relationship_type: relationship_type.clone(),
                    part_uri: part_uri.clone(),
                    payload: SvgPayload::Edited(Arc::clone(payload)),
                    _reservation: reservation.clone(),
                });
                state.owner_state = SvgPictureOwnerState::Embedded;
            },
            BatchOperation::Detach { .. } => {
                state.svg = None;
                state.owner_state = post_detach_owner_state;
            },
        }
        self.current.story = StoryPayload::Projected(projected_xml);
        self.staged_attach_count = next_attach_count;
        self.staged_payload_bytes = next_payload_bytes;
        self.projection_edits.push(projection);
        self.operations.push(operation);
        Ok(())
    }

    fn limits(&self) -> litchi_opc::ReadLimits {
        self.current
            .states
            .first()
            .map(|state| state.limits)
            .unwrap_or_default()
    }

    fn preflight_attach(&self, current: &SnapshotState, payload_len: usize) -> Result<()> {
        let budget = &current.package_budget;
        let payload_len_u64 = u64::try_from(payload_len)
            .map_err(|_| Error::Invalid("DOCX SVG payload length exceeds u64".into()))?;
        let attach_count = self
            .staged_attach_count
            .checked_add(1)
            .ok_or_else(|| Error::Invalid("DOCX SVG attachment count overflows".into()))?;
        let operation_count = self
            .operations
            .len()
            .checked_add(1)
            .ok_or_else(|| Error::Invalid("DOCX SVG operation count overflows".into()))?;
        let limits = current.limits;
        if operation_count > limits.max_total_relationships() {
            return Err(Error::Invalid(
                "DOCX SVG batch operation count exceeds the relationship limit".into(),
            ));
        }
        if budget
            .part_count
            .checked_add(attach_count)
            .is_none_or(|count| count > limits.max_parts())
        {
            return Err(Error::Invalid(
                "DOCX SVG attachment count exceeds the Part limit".into(),
            ));
        }
        let total_part_bytes = budget
            .total_part_bytes
            .checked_add(self.staged_payload_bytes)
            .and_then(|value| value.checked_add(payload_len_u64))
            .ok_or_else(|| Error::Invalid("DOCX SVG total Part bytes overflow".into()))?;
        if total_part_bytes > limits.max_total_part_bytes() {
            return Err(Error::Invalid(
                "DOCX SVG attachment bytes exceed the total Part limit".into(),
            ));
        }
        if budget
            .relationship_count
            .checked_add(attach_count)
            .is_none_or(|count| count > limits.max_total_relationships())
        {
            return Err(Error::Invalid(
                "DOCX SVG attachment relationships exceed the package limit".into(),
            ));
        }
        let member_additions = attach_count
            .checked_mul(2)
            .ok_or_else(|| Error::Invalid("DOCX SVG physical member count overflows".into()))?;
        if budget
            .physical_member_count
            .checked_add(member_additions)
            .is_none_or(|count| count > limits.max_archive_members())
        {
            return Err(Error::Invalid(
                "DOCX SVG attachment members exceed the archive limit".into(),
            ));
        }
        if budget
            .content_type_mapping_lower_bound
            .checked_add(attach_count)
            .is_none_or(|count| count > limits.max_content_type_mappings())
        {
            return Err(Error::Invalid(
                "DOCX SVG attachment content-type mappings exceed the limit".into(),
            ));
        }
        Ok(())
    }
}

fn reserve_payload_memory(
    package: &super::Package,
    payload_len: usize,
) -> Result<Option<Arc<Reservation>>> {
    let Some(context) = package.package.execution_context() else {
        return Ok(None);
    };
    let bytes = (payload_len as u64)
        .checked_add(size_of::<Vec<u8>>() as u64)
        .ok_or_else(|| Error::Invalid("DOCX SVG payload reservation overflows".into()))?;
    context
        .reserve(Resource::Memory, bytes)
        .map(|reservation| Some(Arc::new(reservation)))
        .map_err(|error| Error::from(litchi_opc::OpcError::Execution(error)))
}

impl SourceBackedSvgAttachmentBatchPatch {
    /// Exact source snapshot required by this patch.
    #[must_use]
    pub const fn source(&self) -> &SourceBackedSvgAttachmentBatchSnapshot {
        &self.before
    }

    /// Target snapshot produced by this patch.
    #[must_use]
    pub const fn target(&self) -> &SourceBackedSvgAttachmentBatchSnapshot {
        &self.after
    }

    /// Whether this patch changes story or graph state.
    #[must_use]
    pub fn is_changed(&self) -> bool {
        !same_batch_state(&self.before, &self.after)
    }

    /// Return the exact in-memory inverse.
    #[must_use]
    pub fn inverse(&self) -> Self {
        Self {
            before: self.after.clone(),
            after: self.before.clone(),
        }
    }

    /// Apply this patch only to the exact captured source batch.
    pub fn apply(
        &self,
        source: &SourceBackedSvgAttachmentBatchSnapshot,
    ) -> Result<SourceBackedSvgAttachmentBatchSnapshot> {
        if !same_batch_source(source, &self.before) {
            return Err(Error::Invalid(
                "SVG lifecycle batch patch source is stale or foreign".into(),
            ));
        }
        Ok(if self.is_changed() {
            self.after.clone()
        } else {
            source.clone()
        })
    }
}

impl SourceBackedSvgAttachmentBatchCommit {
    /// Candidate target snapshot.
    #[must_use]
    pub const fn snapshot(&self) -> &SourceBackedSvgAttachmentBatchSnapshot {
        &self.snapshot
    }

    /// Exact-source reversible batch patch.
    #[must_use]
    pub const fn patch(&self) -> &SourceBackedSvgAttachmentBatchPatch {
        &self.patch
    }

    /// Return content-free facts about this commit.
    #[must_use]
    pub const fn diagnostics(&self) -> SvgAttachmentCommitDiagnostics {
        self.diagnostics
    }

    /// Whether this commit changes the package.
    #[must_use]
    pub fn is_changed(&self) -> bool {
        self.patch.is_changed()
    }
}

impl SourceBackedSvgAttachmentBatchPublication {
    /// Borrow the semantic batch represented by the emitted artifact.
    #[must_use]
    pub const fn snapshot(&self) -> &SourceBackedSvgAttachmentBatchSnapshot {
        &self.snapshot
    }

    pub(crate) const fn original_snapshot(&self) -> &SourceBackedSvgAttachmentBatchSnapshot {
        &self.original_snapshot
    }

    pub(crate) const fn original_artifact(&self) -> &SourceArtifact {
        &self.original_artifact
    }

    pub(crate) const fn published_fingerprint(&self) -> SourceArtifactFingerprint {
        self.published_fingerprint
    }
}

impl SourceBackedSvgAttachmentPublication {
    /// Borrow the semantic picture represented by the emitted artifact.
    #[must_use]
    pub const fn snapshot(&self) -> &SourceBackedSvgAttachmentSnapshot {
        &self.snapshot
    }

    pub(crate) const fn original_snapshot(&self) -> &SourceBackedSvgAttachmentSnapshot {
        &self.original_snapshot
    }

    pub(crate) const fn original_artifact(&self) -> &SourceArtifact {
        &self.original_artifact
    }

    pub(crate) const fn published_fingerprint(&self) -> SourceArtifactFingerprint {
        self.published_fingerprint
    }
}

impl SourceBackedSvgAttachmentPatch {
    /// Source snapshot required by this patch.
    #[must_use]
    pub const fn source(&self) -> &SourceBackedSvgAttachmentSnapshot {
        &self.before
    }

    /// Target snapshot produced by this patch.
    #[must_use]
    pub const fn target(&self) -> &SourceBackedSvgAttachmentSnapshot {
        &self.after
    }

    /// Whether this patch changes story or graph state.
    #[must_use]
    pub fn is_changed(&self) -> bool {
        !same_state(&self.before.state, &self.after.state)
    }

    /// Return the exact in-memory inverse.
    #[must_use]
    pub fn inverse(&self) -> Self {
        Self {
            before: self.after.clone(),
            after: self.before.clone(),
        }
    }

    /// Apply this patch only to the exact captured source state.
    pub fn apply(
        &self,
        source: &SourceBackedSvgAttachmentSnapshot,
    ) -> Result<SourceBackedSvgAttachmentSnapshot> {
        if !same_source(&source.state, &self.before.state) {
            return Err(Error::Invalid(
                "SVG lifecycle patch source is stale or foreign".into(),
            ));
        }
        Ok(if self.is_changed() {
            self.after.clone()
        } else {
            source.clone()
        })
    }
}

impl SourceBackedSvgAttachmentCommit {
    /// Candidate target snapshot.
    #[must_use]
    pub const fn snapshot(&self) -> &SourceBackedSvgAttachmentSnapshot {
        &self.snapshot
    }

    /// Exact-source reversible patch.
    #[must_use]
    pub const fn patch(&self) -> &SourceBackedSvgAttachmentPatch {
        &self.patch
    }

    /// Return content-free facts about this commit.
    #[must_use]
    pub const fn diagnostics(&self) -> SvgAttachmentCommitDiagnostics {
        self.diagnostics
    }

    /// Whether this commit changes the package.
    #[must_use]
    pub fn is_changed(&self) -> bool {
        self.patch.is_changed()
    }
}

impl super::Package {
    /// Inspect direct main-story pictures without reading raster or SVG media
    /// payloads. The returned source views retain lazy Part handles; callers
    /// opt into a payload read through `raster().data()` or `svg().data()`.
    pub fn svg_picture_sources(
        &self,
        story: StorySelector,
    ) -> Result<Vec<SourceSvgPictureSourceView<'_>>> {
        if !matches!(story, StorySelector::Main) {
            return Err(unsafe_edit(
                "svg_picture_sources",
                "only the main Word story is in this lifecycle batch",
            ));
        }
        self.package.check_execution()?;
        let capture_version = self.package.source_version()?;
        let main = self.package.main_document_part()?;
        let source = main.source_xml()?;
        let drawing = scan_story(source.bytes(), self.package.read_limits())?;
        let mut views = Vec::new();
        views
            .try_reserve_exact(drawing.pictures().len())
            .map_err(|source| Error::Allocation {
                resource: "DOCX SVG source picture view",
                source,
            })?;
        for picture in drawing.pictures() {
            views.push(source_view_from_picture(
                self,
                &main,
                picture,
                PictureSelector::new(0, picture.picture_ordinal()),
            )?);
        }
        self.package.check_execution()?;
        if self.package.source_version()? != capture_version {
            return Err(Error::Invalid(
                "DOCX SVG picture source changed during capture".into(),
            ));
        }
        Ok(views)
    }

    /// Inspect direct main-story pictures and their SVG owner state.
    ///
    /// This compatibility method returns payload-bearing views and therefore
    /// intentionally reads the selected media. Use [`Self::svg_picture_sources`]
    /// when only metadata or a lazily loaded resource handle is needed.
    pub fn svg_pictures(&self, story: StorySelector) -> Result<Vec<SourceSvgPictureView>> {
        if !matches!(story, StorySelector::Main) {
            return Err(unsafe_edit(
                "svg_pictures",
                "only the main Word story is in this lifecycle batch",
            ));
        }
        let snapshot = capture_inventory(self)?;
        let mut views = Vec::new();
        views
            .try_reserve_exact(snapshot.len())
            .map_err(|source| Error::Allocation {
                resource: "DOCX SVG picture view",
                source,
            })?;
        for state in snapshot {
            views.push(view_from_snapshot(state));
        }
        Ok(views)
    }

    /// Begin an exact-source attach/detach edit on a direct main-story
    /// picture.
    pub fn edit_svg_attachment(
        &self,
        selector: impl Into<PictureSelector>,
    ) -> Result<SourceBackedSvgAttachmentEdit<'_>> {
        self.package.check_execution()?;
        let selector = selector.into();
        let batch = self.edit_svg_attachments([selector])?;
        Ok(SourceBackedSvgAttachmentEdit { batch, selector })
    }

    /// Begin one atomic source-backed edit spanning selected direct
    /// main-story pictures. The source drawing and relationship inventory are
    /// captured once, even when many selectors are supplied.
    pub fn edit_svg_attachments<I, S>(
        &self,
        selectors: I,
    ) -> Result<SourceBackedSvgAttachmentBatchEdit<'_>>
    where
        I: IntoIterator<Item = S>,
        S: Into<PictureSelector>,
    {
        self.package.check_execution()?;
        let selector_limit = self.package.read_limits().max_total_relationships();
        let mut selected = Vec::new();
        for selector in selectors {
            if selected.len() >= selector_limit {
                return Err(Error::Invalid(
                    "SVG lifecycle selector count exceeds the bounded batch limit".into(),
                ));
            }
            selected
                .try_reserve_exact(1)
                .map_err(|source| Error::Allocation {
                    resource: "DOCX SVG lifecycle selectors",
                    source,
                })?;
            selected.push(selector.into());
        }
        let selectors = selected;
        let snapshot = capture_batch(self, selectors)?;
        if snapshot.states.iter().any(|state| {
            matches!(
                state.owner_state,
                SvgPictureOwnerState::Linked
                    | SvgPictureOwnerState::Ambiguous
                    | SvgPictureOwnerState::Refused
            )
        }) {
            return Err(unsafe_edit(
                "edit_svg_attachment",
                "selected picture has a refused SVG or raster owner state",
            ));
        }
        let context = snapshot
            .context
            .clone()
            .ok_or_else(|| Error::Invalid("SVG batch has no source layout context".into()))?;
        Ok(SourceBackedSvgAttachmentBatchEdit {
            package: self,
            source: snapshot.clone(),
            current: snapshot,
            context,
            operations: Vec::new(),
            projection_edits: Vec::new(),
            staged_attach_count: 0,
            staged_payload_bytes: 0,
            next_relationship_index: 0,
            next_part_index: 0,
        })
    }

    /// Alias using the explicit batch noun for downstream callers.
    pub fn edit_svg_attachment_batch<I, S>(
        &self,
        selectors: I,
    ) -> Result<SourceBackedSvgAttachmentBatchEdit<'_>>
    where
        I: IntoIterator<Item = S>,
        S: Into<PictureSelector>,
    {
        self.edit_svg_attachments(selectors)
    }

    /// Compatibility alias for the picture-specific operation name.
    pub fn edit_svg_picture(
        &self,
        selector: impl Into<PictureSelector>,
    ) -> Result<SourceBackedSvgAttachmentEdit<'_>> {
        self.edit_svg_attachment(selector)
    }

    /// Publish one checked SVG attach/detach commit atomically.
    pub fn publish_svg_attachment_commit_to_stream<W: std::io::Write>(
        self,
        writer: W,
        commit: &SourceBackedSvgAttachmentCommit,
    ) -> Result<SourceBackedSvgAttachmentSnapshot> {
        Ok(self
            .publish_svg_attachment_commit_with_publication_to_stream(writer, commit)?
            .snapshot
            .clone())
    }

    /// Publish one checked SVG commit and retain exact physical inverse
    /// authorization for the emitted artifact.
    pub fn publish_svg_attachment_commit_with_publication_to_stream<W: std::io::Write>(
        self,
        writer: W,
        commit: &SourceBackedSvgAttachmentCommit,
    ) -> Result<SourceBackedSvgAttachmentPublication> {
        self.package.check_execution()?;
        let current = capture_picture(&self, commit.patch.before.state.selector)?;
        if !same_source(&current, &commit.patch.before.state) {
            return Err(Error::Invalid(
                "SVG lifecycle commit source is stale or foreign".into(),
            ));
        }
        let target = commit.patch.apply(&current_snapshot(&current))?;
        let original_snapshot = current_snapshot(&current);
        let original_artifact = self.package.source_artifact();
        let before = SourceBackedSvgAttachmentBatchSnapshot {
            states: vec![current.clone()],
            story: current.xml.clone(),
            context: None,
        };
        let expected = SourceBackedSvgAttachmentBatchSnapshot {
            states: vec![target.state.clone()],
            story: target.state.xml.clone(),
            context: None,
        };
        let published_fingerprint = if target_state_changed(&current, &target.state) {
            let plan = build_batch_topology_plan(&self.package, &before, &expected)?;
            publish_checked_topology(self.package, writer, plan, &before, &expected)?
        } else {
            publish_exact_source_with_fingerprint(self.package, writer)?
        };
        Ok(SourceBackedSvgAttachmentPublication {
            snapshot: target,
            original_snapshot,
            original_artifact,
            published_fingerprint,
        })
    }

    /// Publish one checked multi-picture SVG batch atomically.
    pub fn publish_svg_attachment_batch_commit_to_stream<W: std::io::Write>(
        self,
        writer: W,
        commit: &SourceBackedSvgAttachmentBatchCommit,
    ) -> Result<SourceBackedSvgAttachmentBatchSnapshot> {
        Ok(self
            .publish_svg_attachment_batch_commit_with_publication_to_stream(writer, commit)?
            .snapshot
            .clone())
    }

    /// Publish one checked SVG batch and retain exact physical inverse
    /// authorization for the emitted artifact.
    pub fn publish_svg_attachment_batch_commit_with_publication_to_stream<W: std::io::Write>(
        self,
        writer: W,
        commit: &SourceBackedSvgAttachmentBatchCommit,
    ) -> Result<SourceBackedSvgAttachmentBatchPublication> {
        self.package.check_execution()?;
        let selectors = commit
            .patch
            .before
            .states
            .iter()
            .map(|state| state.selector)
            .collect::<Vec<_>>();
        let current = capture_batch(&self, selectors)?;
        if !same_batch_source(&current, &commit.patch.before) {
            return Err(Error::Invalid(
                "SVG lifecycle batch commit source is stale or foreign".into(),
            ));
        }
        let target = commit.patch.apply(&current)?;
        let original_snapshot = current.clone();
        let original_artifact = self.package.source_artifact();
        let published_fingerprint = if !same_batch_state(&current, &target) {
            let plan = build_batch_topology_plan(&self.package, &current, &target)?;
            publish_checked_topology(self.package, writer, plan, &current, &target)?
        } else {
            publish_exact_source_with_fingerprint(self.package, writer)?
        };
        Ok(SourceBackedSvgAttachmentBatchPublication {
            snapshot: target,
            original_snapshot,
            original_artifact,
            published_fingerprint,
        })
    }

    /// Restore the exact source artifact retained by a published single edit.
    /// The reopened package must match both its physical fingerprint and the
    /// published semantic target before any output is accepted.
    pub fn publish_svg_attachment_inverse_to_stream<W: std::io::Write>(
        self,
        writer: W,
        publication: &SourceBackedSvgAttachmentPublication,
    ) -> Result<SourceBackedSvgAttachmentSnapshot> {
        self.package.check_execution()?;
        let current = capture_picture(&self, publication.snapshot.selector())?;
        if !same_state(&current, &publication.snapshot.state)
            || self.package.source_artifact().fingerprint()? != publication.published_fingerprint()
        {
            return Err(Error::Invalid(
                "SVG lifecycle inverse publication source is stale or foreign".into(),
            ));
        }
        publication.original_artifact().write_to_stream(writer)?;
        Ok(publication.original_snapshot().clone())
    }

    /// Restore the exact source artifact retained by a published SVG batch.
    /// The reopened package must match both its physical fingerprint and the
    /// published semantic target before any output is accepted.
    pub fn publish_svg_attachment_batch_inverse_to_stream<W: std::io::Write>(
        self,
        writer: W,
        publication: &SourceBackedSvgAttachmentBatchPublication,
    ) -> Result<SourceBackedSvgAttachmentBatchSnapshot> {
        self.package.check_execution()?;
        let selectors = publication
            .snapshot
            .states
            .iter()
            .map(|state| state.selector)
            .collect::<Vec<_>>();
        let current = capture_batch(&self, selectors)?;
        if !same_batch_state(&current, &publication.snapshot)
            || self.package.source_artifact().fingerprint()? != publication.published_fingerprint()
        {
            return Err(Error::Invalid(
                "SVG lifecycle batch inverse publication source is stale or foreign".into(),
            ));
        }
        publication.original_artifact().write_to_stream(writer)?;
        Ok(publication.original_snapshot().clone())
    }

    /// Alias using the explicit batch noun for downstream callers.
    pub fn publish_svg_lifecycle_batch_commit_to_stream<W: std::io::Write>(
        self,
        writer: W,
        commit: &SourceBackedSvgAttachmentBatchCommit,
    ) -> Result<SourceBackedSvgAttachmentBatchSnapshot> {
        self.publish_svg_attachment_batch_commit_to_stream(writer, commit)
    }

    /// Compatibility alias for publication callers using lifecycle wording.
    pub fn publish_svg_lifecycle_commit_to_stream<W: std::io::Write>(
        self,
        writer: W,
        commit: &SourceBackedSvgAttachmentCommit,
    ) -> Result<SourceBackedSvgAttachmentSnapshot> {
        self.publish_svg_attachment_commit_to_stream(writer, commit)
    }
}

/// Prepare the source-preserving topology, validate its complete typed
/// effective candidate, and only then release it to the external sink.  The
/// OPC prepared seam retains lazy source payloads and all graph/content-type
/// overlays; this function never stages a complete ZIP archive in memory.
fn publish_checked_topology<W: std::io::Write>(
    package: litchi_opc::SourceBackedPackage,
    writer: W,
    plan: SourceTopologyPlan,
    before: &SourceBackedSvgAttachmentBatchSnapshot,
    expected: &SourceBackedSvgAttachmentBatchSnapshot,
) -> Result<SourceArtifactFingerprint> {
    let prepared = package.prepare_topology(plan)?;
    prepared.with_candidate(|candidate| {
        Ok(validate_effective_candidate(candidate, before, expected))
    })??;
    let mut output = super::FingerprintingWriter {
        inner: writer,
        hasher: Sha256::new(),
    };
    prepared.publish_to_stream(&mut output)?;
    Ok(SourceArtifactFingerprint::from_sha256(
        output.hasher.finalize().into(),
    ))
}

fn publish_exact_source_with_fingerprint<W: std::io::Write>(
    package: litchi_opc::SourceBackedPackage,
    writer: W,
) -> Result<SourceArtifactFingerprint> {
    let mut output = super::FingerprintingWriter {
        inner: writer,
        hasher: Sha256::new(),
    };
    package.write_topology_to_stream(&mut output, SourceTopologyPlan::new())?;
    Ok(SourceArtifactFingerprint::from_sha256(
        output.hasher.finalize().into(),
    ))
}

/// Validate the complete typed candidate exposed by OPC preparation.  The
/// selected story/picture checks are paired with package-wide relationship and
/// content-type fingerprints so a successful owner readback cannot hide an
/// unrelated dangling edge or retained media mapping.
fn validate_effective_candidate(
    candidate: &EffectiveTopology<'_>,
    before: &SourceBackedSvgAttachmentBatchSnapshot,
    expected: &SourceBackedSvgAttachmentBatchSnapshot,
) -> Result<()> {
    if before.states.len() != expected.states.len() || expected.states.is_empty() {
        return Err(Error::Invalid(
            "SVG lifecycle candidate state shape differs from its source".into(),
        ));
    }
    validate_effective_graph(candidate, before)?;
    let (expected_relationships, expected_content_types) =
        expected_transition_fingerprints(candidate, before, expected)?;
    let mut actual_relationships = effective_relationship_fingerprints(candidate)?;
    let mut actual_content_types = effective_content_type_fingerprints(candidate)?;
    let mut expected_relationships = expected_relationships;
    let mut expected_content_types = expected_content_types;
    sort_relationship_fingerprints(&mut actual_relationships);
    sort_relationship_fingerprints(&mut expected_relationships);
    sort_content_type_fingerprints(&mut actual_content_types);
    sort_content_type_fingerprints(&mut expected_content_types);
    if actual_relationships != expected_relationships
        || actual_content_types != expected_content_types
    {
        return Err(unsafe_edit(
            "SVG lifecycle publication",
            "prepared candidate relationship or content-type closure differs from the staged transition",
        ));
    }

    let story_uri = expected
        .states
        .first()
        .map(|state| state.part_uri.clone())
        .ok_or_else(|| Error::Invalid("SVG lifecycle candidate has no story state".into()))?;
    if expected
        .states
        .iter()
        .any(|state| state.part_uri != story_uri)
    {
        return Err(Error::Invalid(
            "SVG lifecycle candidate spans multiple story parts".into(),
        ));
    }
    // All selected pictures in this batch belong to the same main-story
    // Part. Read and scan that effective Part once, then validate each
    // selected ordinal against the shared inventory. This avoids repeating a
    // potentially large source read and XML scan for every selected picture.
    let story_part = candidate.part(&story_uri)?;
    let story_data = story_part.data()?;
    if !same_bytes(story_data.as_bytes(), expected.story.as_bytes()) {
        return Err(unsafe_edit(
            "SVG lifecycle publication",
            "prepared candidate story bytes differ from the staged source projection",
        ));
    }
    let drawing = scan_story(story_data.as_bytes(), candidate.read_limits())?;
    let relationships = candidate.relationships(&story_uri)?;
    for state in &expected.states {
        let picture = drawing.picture(state.selector.picture)?;
        if picture.placement() != state.placement
            || picture.raster_relationship_id() != Some(state.raster_relationship_id.as_str())
        {
            return Err(unsafe_edit(
                "SVG lifecycle publication",
                "prepared candidate does not resolve the selected raster picture",
            ));
        }
        let raster = relationships
            .get(&state.raster_relationship_id)
            .ok_or_else(|| {
                Error::Invalid(format!(
                    "prepared candidate is missing raster relationship '{}'",
                    state.raster_relationship_id
                ))
            })?;
        if raster.is_external()
            || raster.target_mode() != TargetMode::Internal
            || raster.reltype() != state.raster_relationship_type
            || !raster
                .target_partname()?
                .is_equivalent_to(&state.raster_part_uri)
        {
            return Err(unsafe_edit(
                "SVG lifecycle publication",
                "prepared candidate raster relationship changed",
            ));
        }
        let raster_part = candidate.part(&state.raster_part_uri)?;
        let raster_data = raster_part.data()?;
        if raster_part.content_type().as_str() != state.raster_content_type
            || raster_data.as_bytes() != state.raster_payload.as_bytes()
        {
            return Err(unsafe_edit(
                "SVG lifecycle publication",
                "prepared candidate raster fallback changed",
            ));
        }
        if map_owner_state(picture.svg_owner()) != state.owner_state {
            return Err(unsafe_edit(
                "SVG lifecycle publication",
                "prepared candidate SVG owner state differs from the staged source projection",
            ));
        }
        match (&state.svg, picture.svg_owner()) {
            (Some(svg), SvgOwnerState::Embedded(owner)) => {
                if owner.embedded_relationship_id() != Some(svg.relationship_id.as_str()) {
                    return Err(unsafe_edit(
                        "SVG lifecycle publication",
                        "prepared candidate resolves a different SVG relationship ID",
                    ));
                }
                let relation = relationships.get(&svg.relationship_id).ok_or_else(|| {
                    Error::Invalid(format!(
                        "prepared candidate is missing SVG relationship '{}'",
                        svg.relationship_id
                    ))
                })?;
                if relation.is_external()
                    || relation.target_mode() != TargetMode::Internal
                    || relation.reltype() != svg.relationship_type
                    || !relation.target_partname()?.is_equivalent_to(&svg.part_uri)
                {
                    return Err(unsafe_edit(
                        "SVG lifecycle publication",
                        "prepared candidate SVG relationship changed",
                    ));
                }
                let svg_part = candidate.part(&svg.part_uri)?;
                let svg_data = svg_part.data()?;
                if !is_svg_content_type(svg_part.content_type().as_str())
                    || svg_data.as_bytes() != svg.payload.as_bytes()
                    || !svg_part.relationships().is_empty()
                    || !candidate.has_physical_member(svg.part_uri.membername())?
                {
                    return Err(unsafe_edit(
                        "SVG lifecycle publication",
                        "prepared candidate SVG media closure changed",
                    ));
                }
            },
            (None, SvgOwnerState::None | SvgOwnerState::Opaque) => {},
            (None, _) => {
                return Err(unsafe_edit(
                    "SVG lifecycle publication",
                    "prepared candidate retained a selected SVG owner",
                ));
            },
            (Some(_), _) => {
                return Err(unsafe_edit(
                    "SVG lifecycle publication",
                    "prepared candidate lost the selected embedded SVG owner",
                ));
            },
        }
    }
    Ok(())
}

fn validate_effective_graph(
    candidate: &EffectiveTopology<'_>,
    before: &SourceBackedSvgAttachmentBatchSnapshot,
) -> Result<()> {
    let source_relationships = before
        .states
        .first()
        .map(|state| state.relationship_fingerprint.as_ref())
        .unwrap_or(&[]);
    let package_owner = PackURI::new("/").map_err(|error| Error::Uri(error.to_string()))?;
    validate_effective_relationship_targets(
        candidate,
        &package_owner,
        candidate.package_relationships(),
        source_relationships,
    )?;
    for part in candidate.parts() {
        let owner = part.partname().clone();
        validate_effective_relationship_targets(
            candidate,
            &owner,
            part.relationships(),
            source_relationships,
        )?;
        // Every effective Part must have a resolved content-type binding.
        if candidate.content_type(&owner)?.as_str() != part.content_type().as_str() {
            return Err(unsafe_edit(
                "SVG lifecycle publication",
                "prepared candidate content-type binding differs from its Part catalog",
            ));
        }
    }
    Ok(())
}

fn validate_effective_relationship_targets(
    candidate: &EffectiveTopology<'_>,
    owner: &PackURI,
    relationships: &litchi_opc::Relationships,
    source_relationships: &[RelationshipFingerprint],
) -> Result<()> {
    for relationship in relationships.iter() {
        if relationship.is_external() {
            continue;
        }
        let target = relationship.target_partname()?;
        if candidate.part(&target).is_err() {
            let was_source_edge = source_relationships.iter().any(|source| {
                source.owner.is_equivalent_to(owner)
                    && source.id == relationship.r_id()
                    && source.relationship_type == relationship.reltype()
                    && source.target_mode == relationship.target_mode()
                    && source
                        .target
                        .as_ref()
                        .is_some_and(|source_target| source_target.is_equivalent_to(&target))
            });
            if was_source_edge {
                continue;
            }
            return Err(unsafe_edit(
                "SVG lifecycle publication",
                "prepared candidate contains a dangling internal relationship",
            ));
        }
    }
    Ok(())
}

fn effective_relationship_fingerprints(
    candidate: &EffectiveTopology<'_>,
) -> Result<Vec<RelationshipFingerprint>> {
    let root = PackURI::new("/").map_err(|error| Error::Uri(error.to_string()))?;
    let mut result = Vec::new();
    result
        .try_reserve(candidate.package_relationships().len())
        .map_err(|source| Error::Allocation {
            resource: "DOCX SVG effective relationship fingerprint",
            source,
        })?;
    append_effective_relationships(&mut result, &root, candidate.package_relationships())?;
    for part in candidate.parts() {
        let owner = part.partname().clone();
        result
            .try_reserve(part.relationships().len())
            .map_err(|source| Error::Allocation {
                resource: "DOCX SVG effective relationship fingerprint",
                source,
            })?;
        append_effective_relationships(&mut result, &owner, part.relationships())?;
    }
    Ok(result)
}

fn append_effective_relationships(
    output: &mut Vec<RelationshipFingerprint>,
    owner: &PackURI,
    relationships: &litchi_opc::Relationships,
) -> Result<()> {
    for relationship in relationships.iter() {
        output.push(RelationshipFingerprint {
            owner: owner.clone(),
            id: relationship.r_id().to_owned(),
            relationship_type: relationship.reltype().to_owned(),
            target: (!relationship.is_external())
                .then(|| relationship.target_partname())
                .transpose()?,
            target_mode: relationship.target_mode(),
        });
    }
    Ok(())
}

fn effective_content_type_fingerprints(
    candidate: &EffectiveTopology<'_>,
) -> Result<Vec<ContentTypeFingerprint>> {
    let mut result = Vec::new();
    result
        .try_reserve(candidate.parts().len())
        .map_err(|source| Error::Allocation {
            resource: "DOCX SVG effective content-type fingerprint",
            source,
        })?;
    for part in candidate.parts() {
        result.push(ContentTypeFingerprint {
            part_uri: part.partname().clone(),
            content_type: part.content_type().as_str().to_owned(),
        });
    }
    Ok(result)
}

fn expected_transition_fingerprints(
    candidate: &EffectiveTopology<'_>,
    before: &SourceBackedSvgAttachmentBatchSnapshot,
    after: &SourceBackedSvgAttachmentBatchSnapshot,
) -> Result<(Vec<RelationshipFingerprint>, Vec<ContentTypeFingerprint>)> {
    let source = before
        .states
        .first()
        .ok_or_else(|| Error::Invalid("SVG lifecycle source batch is empty".into()))?;
    let mut relationships = source.relationship_fingerprint.as_ref().to_vec();
    let mut content_types = source.content_type_fingerprint.as_ref().to_vec();
    let mut detached = Vec::<(String, PackURI)>::new();
    for (old, new) in before.states.iter().zip(&after.states) {
        match (&old.svg, &new.svg) {
            (None, Some(svg)) => {
                relationships.push(RelationshipFingerprint {
                    owner: old.part_uri.clone(),
                    id: svg.relationship_id.clone(),
                    relationship_type: svg.relationship_type.clone(),
                    target: Some(svg.part_uri.clone()),
                    target_mode: TargetMode::Internal,
                });
                content_types.push(ContentTypeFingerprint {
                    part_uri: svg.part_uri.clone(),
                    content_type: "image/svg+xml".to_owned(),
                });
            },
            (Some(svg), None) => detached.push((svg.relationship_id.clone(), svg.part_uri.clone())),
            (Some(left), Some(right))
                if !same_attachment(&Some(left.clone()), &Some(right.clone())) =>
            {
                return Err(unsafe_edit(
                    "SVG lifecycle publication",
                    "SVG owner replacement is not supported",
                ));
            },
            _ => {},
        }
    }
    if !detached.is_empty() {
        let drawing = scan_story(before.story.as_bytes(), candidate.read_limits())?;
        for (id, target) in &detached {
            let total_references = drawing
                .relationship_references()
                .iter()
                .filter(|reference| reference.id() == id)
                .count();
            let detached_references = detached.iter().filter(|(value, _)| value == id).count();
            if total_references <= detached_references {
                relationships.retain(|relationship| {
                    !(relationship.owner.is_equivalent_to(&source.part_uri)
                        && relationship.id == *id)
                });
                if !relationships.iter().any(|relationship| {
                    relationship
                        .target
                        .as_ref()
                        .is_some_and(|value| value.is_equivalent_to(target))
                }) {
                    content_types
                        .retain(|content_type| !content_type.part_uri.is_equivalent_to(target));
                }
            }
        }
    }
    Ok((relationships, content_types))
}

fn sort_relationship_fingerprints(values: &mut [RelationshipFingerprint]) {
    values.sort_unstable_by(|left, right| {
        left.owner
            .as_str()
            .cmp(right.owner.as_str())
            .then_with(|| left.id.cmp(&right.id))
            .then_with(|| left.relationship_type.cmp(&right.relationship_type))
            .then_with(|| {
                (left.target_mode == TargetMode::External)
                    .cmp(&(right.target_mode == TargetMode::External))
            })
            .then_with(|| {
                left.target
                    .as_ref()
                    .map(PackURI::as_str)
                    .cmp(&right.target.as_ref().map(PackURI::as_str))
            })
    });
}

fn sort_content_type_fingerprints(values: &mut [ContentTypeFingerprint]) {
    values.sort_unstable_by(|left, right| {
        left.part_uri
            .as_str()
            .cmp(right.part_uri.as_str())
            .then_with(|| left.content_type.cmp(&right.content_type))
    });
}

fn current_snapshot(state: &SnapshotState) -> SourceBackedSvgAttachmentSnapshot {
    SourceBackedSvgAttachmentSnapshot {
        state: state.clone(),
    }
}

fn capture_inventory(package: &super::Package) -> Result<Vec<SnapshotState>> {
    let inventory_version = package.package.source_version()?;
    let main = package.package.main_document_part()?;
    let source = main.source_xml()?;
    let source_bytes = source.bytes();
    let drawing = scan_story(source_bytes, package.package.read_limits())?;
    let identities = identity_inventory(package, &main)?;
    let closure = closure_inventory(package)?;
    let budget = package_budget(package)?;
    let mut result = Vec::new();
    result
        .try_reserve_exact(drawing.pictures().len())
        .map_err(|source| Error::Allocation {
            resource: "DOCX SVG picture inventory",
            source,
        })?;
    for picture in drawing.pictures() {
        result.push(capture_from_picture(
            package,
            main.partname().clone(),
            source.clone(),
            picture,
            PictureSelector::new(0, picture.picture_ordinal()),
            identities.0.clone(),
            identities.1.clone(),
            closure.0.clone(),
            closure.1.clone(),
            Arc::clone(&budget),
            drawing.relationship_dialect(),
        )?);
    }
    if package.package.source_version()? != inventory_version {
        return Err(Error::Invalid(
            "DOCX SVG picture inventory source changed during capture".into(),
        ));
    }
    Ok(result)
}

fn capture_batch(
    package: &super::Package,
    mut selectors: Vec<PictureSelector>,
) -> Result<SourceBackedSvgAttachmentBatchSnapshot> {
    if selectors.is_empty() {
        return Err(Error::Invalid(
            "SVG lifecycle batch must select at least one picture".into(),
        ));
    }
    if selectors.iter().any(|selector| selector.drawing != 0) {
        let selector = selectors
            .iter()
            .find(|selector| selector.drawing != 0)
            .copied()
            .unwrap_or(PictureSelector::new(usize::MAX, 0));
        return Err(Error::OutOfBounds {
            object: "DOCX main-story drawing",
            index: selector.drawing,
            len: 1,
        });
    }
    selectors.sort_unstable_by_key(|selector| selector.picture);
    for pair in selectors.windows(2) {
        if pair[0] == pair[1] {
            return Err(Error::Invalid(
                "SVG lifecycle batch selects one picture more than once".into(),
            ));
        }
    }
    let inventory_version = package.package.source_version()?;
    let main = package.package.main_document_part()?;
    let source = main.source_xml()?;
    let drawing = scan_story(source.bytes(), package.package.read_limits())?;
    let identities = identity_inventory(package, &main)?;
    let closure = closure_inventory(package)?;
    let budget = package_budget(package)?;
    let mut states = Vec::new();
    states
        .try_reserve_exact(selectors.len())
        .map_err(|source| Error::Allocation {
            resource: "DOCX SVG selected picture batch",
            source,
        })?;
    for selector in selectors {
        let picture = drawing.picture(selector.picture)?;
        let state = capture_from_picture(
            package,
            main.partname().clone(),
            source.clone(),
            picture,
            selector,
            Arc::clone(&identities.0),
            Arc::clone(&identities.1),
            Arc::clone(&closure.0),
            Arc::clone(&closure.1),
            Arc::clone(&budget),
            drawing.relationship_dialect(),
        )?;
        states.push(state);
    }
    let mut layouts = Vec::new();
    layouts
        .try_reserve_exact(drawing.pictures().len())
        .map_err(|source| Error::Allocation {
            resource: "DOCX SVG batch picture layouts",
            source,
        })?;
    for picture in drawing.pictures() {
        layouts.push(layout_from_picture(picture)?);
    }
    if package.package.source_version()? != inventory_version {
        return Err(Error::Invalid(
            "DOCX SVG selected-picture capture source changed during capture".into(),
        ));
    }
    let context = Arc::new(BatchSourceContext {
        source_xml: source.clone(),
        source_version: inventory_version,
        lineage: package.package.source_lineage(),
        layouts: Arc::from(layouts.into_boxed_slice()),
    });
    Ok(SourceBackedSvgAttachmentBatchSnapshot {
        states,
        story: StoryPayload::Original(source),
        context: Some(context),
    })
}

fn capture_picture(package: &super::Package, selector: PictureSelector) -> Result<SnapshotState> {
    let capture_version = package.package.source_version()?;
    if selector.drawing != 0 {
        return Err(Error::OutOfBounds {
            object: "DOCX main-story drawing",
            index: selector.drawing,
            len: 1,
        });
    }
    let main = package.package.main_document_part()?;
    let source = main.source_xml()?;
    let scan_source = source.clone();
    let drawing = scan_story(scan_source.bytes(), package.package.read_limits())?;
    let picture = drawing.picture(selector.picture)?;
    let identities = identity_inventory(package, &main)?;
    let closure = closure_inventory(package)?;
    let budget = package_budget(package)?;
    let state = capture_from_picture(
        package,
        main.partname().clone(),
        source,
        picture,
        selector,
        identities.0,
        identities.1,
        closure.0,
        closure.1,
        budget,
        drawing.relationship_dialect(),
    )?;
    if package.package.source_version()? != capture_version {
        return Err(Error::Invalid(
            "DOCX SVG picture source changed during capture".into(),
        ));
    }
    Ok(state)
}

fn capture_from_picture(
    package: &super::Package,
    part_uri: PackURI,
    source: litchi_opc::SourceXmlPart,
    picture: &PictureSource<'_>,
    selector: PictureSelector,
    relationship_ids: Arc<[String]>,
    physical_names: Arc<[String]>,
    relationship_fingerprint: Arc<[RelationshipFingerprint]>,
    content_type_fingerprint: Arc<[ContentTypeFingerprint]>,
    package_budget: Arc<PackageBudget>,
    relationship_dialect: RelationshipDialect,
) -> Result<SnapshotState> {
    let raster_id = picture
        .raster_relationship_id()
        .ok_or_else(|| Error::Invalid("direct DOCX picture has no raster relationship".into()))?;
    let main = package.package.part(&part_uri)?;
    let raster = main
        .rels()
        .get(raster_id)
        .ok_or_else(|| Error::Invalid(format!("raster relationship '{raster_id}' is missing")))?;
    if raster.is_external() || raster.target_mode() != TargetMode::Internal {
        return Err(Error::Invalid(
            "DOCX SVG lifecycle requires an internal raster relationship".into(),
        ));
    }
    let raster_uri = raster.target_partname()?;
    if !raster_uri.as_str().starts_with("/word/media/")
        || (raster.reltype() != rt::IMAGE && raster.reltype() != rt::STRICT_IMAGE)
    {
        return Err(Error::Invalid(
            "selected raster fallback is outside /word/media or has a non-image relationship"
                .into(),
        ));
    }
    let raster_part = package.package.part(&raster_uri)?;
    if raster_part.content_type() != ct::PNG {
        return Err(Error::Invalid(format!(
            "selected raster fallback has content type '{}', expected image/png",
            raster_part.content_type()
        )));
    }
    let raster_payload = raster_part.data()?;
    let owner_state = map_owner_state(picture.svg_owner());
    let svg = match picture.svg_owner() {
        SvgOwnerState::Embedded(owner) => {
            let id = owner.embedded_relationship_id().ok_or_else(|| {
                Error::Invalid("embedded SVG owner has no embedded relationship".into())
            })?;
            let relation = main
                .rels()
                .get(id)
                .ok_or_else(|| Error::Invalid(format!("SVG relationship '{id}' is missing")))?;
            if relation.is_external() || relation.target_mode() != TargetMode::Internal {
                return Err(Error::Invalid(
                    "SVG lifecycle requires an internal SVG relationship".into(),
                ));
            }
            let svg_uri = relation.target_partname()?;
            if !svg_uri.as_str().starts_with("/word/media/") {
                return Err(Error::Invalid(
                    "SVG lifecycle requires SVG media below /word/media/".into(),
                ));
            }
            if relation.reltype() != rt::IMAGE && relation.reltype() != rt::STRICT_IMAGE {
                return Err(Error::Invalid(
                    "SVG relationship does not use an image relationship type".into(),
                ));
            }
            let svg_part = package.package.part(&svg_uri)?;
            if !is_svg_content_type(svg_part.content_type()) {
                return Err(Error::Invalid(format!(
                    "SVG relationship target has content type '{}'",
                    svg_part.content_type()
                )));
            }
            if !svg_part.rels().is_empty() {
                return Err(Error::Invalid(
                    "SVG media Part has outbound relationships".into(),
                ));
            }
            Some(AttachmentState {
                relationship_id: id.to_owned(),
                relationship_type: relation.reltype().to_owned(),
                part_uri: svg_uri,
                payload: SvgPayload::Original(svg_part.data()?),
                _reservation: None,
            })
        },
        SvgOwnerState::Linked(_) | SvgOwnerState::Ambiguous | SvgOwnerState::Refused => None,
        SvgOwnerState::None | SvgOwnerState::Opaque => None,
    };
    Ok(SnapshotState {
        part_uri,
        xml: StoryPayload::Original(source),
        selector,
        placement: picture.placement(),
        raster_relationship_id: raster_id.to_owned(),
        raster_relationship_type: raster.reltype().to_owned(),
        raster_part_uri: raster_uri,
        raster_content_type: raster_part.content_type().to_owned(),
        raster_payload,
        owner_state,
        relationship_dialect,
        svg,
        existing_relationship_ids: relationship_ids,
        physical_member_names: physical_names,
        relationship_fingerprint,
        content_type_fingerprint,
        package_budget,
        limits: package.package.read_limits(),
        lineage: package.package.source_lineage(),
        source_version: package.package.source_version()?,
        source_artifact: package.package.source_artifact(),
    })
}

fn view_from_snapshot(state: SnapshotState) -> SourceSvgPictureView {
    let raster_relationship_id = state.raster_relationship_id.clone();
    let raster_relationship_type = state.raster_relationship_type.clone();
    let raster_part_uri = state.raster_part_uri.clone();
    let raster_content_type = state.raster_content_type.clone();
    SourceSvgPictureView {
        selector: state.selector,
        placement: state.placement,
        raster_relationship_id: state.raster_relationship_id,
        raster_part_uri: state.raster_part_uri,
        owner_state: state.owner_state,
        raster: RasterResourceView {
            relationship_id: raster_relationship_id,
            relationship_type: raster_relationship_type,
            part_uri: raster_part_uri,
            content_type: raster_content_type,
            payload: state.raster_payload.clone(),
        },
        svg: state.svg.map(|attachment| SvgResourceView {
            attachment: attachment.public(),
        }),
    }
}

fn source_view_from_picture<'a>(
    package: &'a super::Package,
    main: &PartView<'a>,
    picture: &PictureSource<'_>,
    selector: PictureSelector,
) -> Result<SourceSvgPictureSourceView<'a>> {
    let raster_id = picture
        .raster_relationship_id()
        .ok_or_else(|| Error::Invalid("direct DOCX picture has no raster relationship".into()))?;
    let raster = main
        .rels()
        .get(raster_id)
        .ok_or_else(|| Error::Invalid(format!("raster relationship '{raster_id}' is missing")))?;
    if raster.is_external() || raster.target_mode() != TargetMode::Internal {
        return Err(Error::Invalid(
            "DOCX SVG lifecycle requires an internal raster relationship".into(),
        ));
    }
    let raster_uri = raster.target_partname()?;
    let raster_part = package.package.part(&raster_uri)?;
    if !raster_uri.as_str().starts_with("/word/media/")
        || (raster.reltype() != rt::IMAGE && raster.reltype() != rt::STRICT_IMAGE)
        || raster_part.content_type() != ct::PNG
    {
        return Err(Error::Invalid(
            "selected raster fallback is not an internal PNG media Part".into(),
        ));
    }
    let svg = match picture.svg_owner() {
        SvgOwnerState::Embedded(owner) => {
            let id = owner.embedded_relationship_id().ok_or_else(|| {
                Error::Invalid("embedded SVG owner has no embedded relationship".into())
            })?;
            let relation = main
                .rels()
                .get(id)
                .ok_or_else(|| Error::Invalid(format!("SVG relationship '{id}' is missing")))?;
            if relation.is_external() || relation.target_mode() != TargetMode::Internal {
                return Err(Error::Invalid(
                    "SVG lifecycle requires an internal SVG relationship".into(),
                ));
            }
            let svg_uri = relation.target_partname()?;
            if !svg_uri.as_str().starts_with("/word/media/")
                || (relation.reltype() != rt::IMAGE && relation.reltype() != rt::STRICT_IMAGE)
            {
                return Err(Error::Invalid(
                    "SVG relationship does not target an image media Part".into(),
                ));
            }
            let svg_part = package.package.part(&svg_uri)?;
            if !is_svg_content_type(svg_part.content_type()) || !svg_part.rels().is_empty() {
                return Err(Error::Invalid(
                    "SVG media Part is not a relationship-free image/svg+xml leaf".into(),
                ));
            }
            Some(SvgResourceSourceView {
                relationship_id: relation.r_id(),
                relationship_type: relation.reltype(),
                part_uri: svg_part.partname(),
                part: svg_part,
            })
        },
        SvgOwnerState::None
        | SvgOwnerState::Linked(_)
        | SvgOwnerState::Ambiguous
        | SvgOwnerState::Refused
        | SvgOwnerState::Opaque => None,
    };
    Ok(SourceSvgPictureSourceView {
        selector,
        placement: picture.placement(),
        owner_state: map_owner_state(picture.svg_owner()),
        raster: RasterResourceSourceView {
            relationship_id: raster.r_id(),
            relationship_type: raster.reltype(),
            part_uri: raster_part.partname(),
            content_type: raster_part.content_type(),
            part: raster_part,
        },
        svg,
    })
}

fn map_owner_state(state: &SvgOwnerState<'_>) -> SvgPictureOwnerState {
    match state {
        SvgOwnerState::None => SvgPictureOwnerState::None,
        SvgOwnerState::Embedded(_) => SvgPictureOwnerState::Embedded,
        SvgOwnerState::Linked(_) => SvgPictureOwnerState::Linked,
        SvgOwnerState::Ambiguous => SvgPictureOwnerState::Ambiguous,
        SvgOwnerState::Refused => SvgPictureOwnerState::Refused,
        SvgOwnerState::Opaque => SvgPictureOwnerState::Opaque,
    }
}

fn copy_layout_prefix(prefix: &[u8]) -> Result<Box<[u8]>> {
    let mut owned = Vec::new();
    owned
        .try_reserve_exact(prefix.len())
        .map_err(|source| Error::Allocation {
            resource: "DOCX SVG picture namespace prefix",
            source,
        })?;
    owned.extend_from_slice(prefix);
    Ok(owned.into_boxed_slice())
}

fn layout_from_picture(picture: &PictureSource<'_>) -> Result<PictureLayout> {
    let blip = picture
        .blip_range()
        .ok_or_else(|| Error::Invalid("direct picture has no raster blip range".into()))?;
    let ext_list = picture.ext_list_range();
    let owner_state = map_owner_state(picture.svg_owner());
    let svg_extension_range = match picture.svg_owner() {
        SvgOwnerState::Embedded(owner) => Some(owner.extension_range()),
        SvgOwnerState::None
        | SvgOwnerState::Linked(_)
        | SvgOwnerState::Ambiguous
        | SvgOwnerState::Refused
        | SvgOwnerState::Opaque => None,
    };
    let post_detach_owner_state = if picture.has_opaque_svg_extension() {
        SvgPictureOwnerState::Opaque
    } else {
        SvgPictureOwnerState::None
    };
    Ok(PictureLayout {
        blip_range: blip.range(),
        blip_prefix: copy_layout_prefix(blip.prefix())?,
        blip_close_start: blip.close_start(),
        ext_list_range: ext_list.map(|range| range.range()),
        ext_list_prefix: ext_list
            .map(|range| copy_layout_prefix(range.prefix()))
            .transpose()?,
        ext_list_close_start: ext_list.and_then(|range| range.close_start()),
        svg_extension_range,
        owner_state,
        post_detach_owner_state,
    })
}

fn scan_story<'a>(bytes: &'a [u8], limits: litchi_opc::ReadLimits) -> Result<SourceDrawing<'a>> {
    let max = usize::try_from(limits.max_part_bytes()).unwrap_or(usize::MAX);
    let mut scan = ScanLimits::default();
    scan.max_xml_bytes = scan.max_xml_bytes.min(max);
    scan.max_fragment_bytes = scan.max_fragment_bytes.min(max);
    SourceDrawing::scan_with_limits(bytes, 0, scan)
}

fn identity_inventory(
    package: &super::Package,
    main: &PartView<'_>,
) -> Result<(Arc<[String]>, Arc<[String]>)> {
    let mut relationship_ids = Vec::new();
    relationship_ids
        .try_reserve_exact(main.rels().len())
        .map_err(|source| Error::Allocation {
            resource: "DOCX SVG relationship identity inventory",
            source,
        })?;
    for relation in main.rels().iter() {
        relationship_ids.push(relation.r_id().to_owned());
    }
    let mut physical_names = Vec::new();
    physical_names
        .try_reserve_exact(package.package.physical_member_names().len())
        .map_err(|source| Error::Allocation {
            resource: "DOCX SVG physical identity inventory",
            source,
        })?;
    for name in package.package.physical_member_names() {
        physical_names.push(name.to_owned());
    }
    Ok((
        Arc::from(relationship_ids.into_boxed_slice()),
        Arc::from(physical_names.into_boxed_slice()),
    ))
}

fn package_budget(package: &super::Package) -> Result<Arc<PackageBudget>> {
    let mut part_count = 0usize;
    let mut total_part_bytes = 0u64;
    let mut relationship_count = package.package.rels().len();
    for part in package.package.iter_parts() {
        part_count = part_count
            .checked_add(1)
            .ok_or_else(|| Error::Invalid("DOCX package Part count overflows".into()))?;
        total_part_bytes = total_part_bytes
            .checked_add(part.declared_uncompressed_size()?)
            .ok_or_else(|| Error::Invalid("DOCX package Part byte count overflows".into()))?;
        relationship_count = relationship_count
            .checked_add(part.rels().len())
            .ok_or_else(|| Error::Invalid("DOCX package relationship count overflows".into()))?;
    }
    let physical_member_count = package.package.physical_member_names().len();
    Ok(Arc::new(PackageBudget {
        part_count,
        total_part_bytes,
        relationship_count,
        physical_member_count,
        // Every indexed Part needs a content-type mapping or default. This is
        // a lower-bound precharge; the publisher independently parses the
        // manifest before applying the final override.
        content_type_mapping_lower_bound: part_count,
    }))
}

fn closure_inventory(
    package: &super::Package,
) -> Result<(
    Arc<[RelationshipFingerprint]>,
    Arc<[ContentTypeFingerprint]>,
)> {
    let mut relationships = Vec::new();
    let root = PackURI::new("/").map_err(|error| Error::Uri(error.to_string()))?;
    relationships
        .try_reserve(package.package.rels().len())
        .map_err(|source| Error::Allocation {
            resource: "DOCX SVG package relationship fingerprint",
            source,
        })?;
    for relation in package.package.rels().iter() {
        relationships.push(RelationshipFingerprint {
            owner: root.clone(),
            id: relation.r_id().to_owned(),
            relationship_type: relation.reltype().to_owned(),
            target: (!relation.is_external())
                .then(|| relation.target_partname())
                .transpose()?,
            target_mode: relation.target_mode(),
        });
    }
    let mut content_types = Vec::new();
    content_types
        .try_reserve_exact(package.package.iter_parts().count())
        .map_err(|source| Error::Allocation {
            resource: "DOCX SVG content-type fingerprint",
            source,
        })?;
    for part in package.package.iter_parts() {
        content_types.push(ContentTypeFingerprint {
            part_uri: part.partname().clone(),
            content_type: part.content_type().to_owned(),
        });
        relationships
            .try_reserve(part.rels().len())
            .map_err(|source| Error::Allocation {
                resource: "DOCX SVG relationship fingerprint",
                source,
            })?;
        for relation in part.rels().iter() {
            relationships.push(RelationshipFingerprint {
                owner: part.partname().clone(),
                id: relation.r_id().to_owned(),
                relationship_type: relation.reltype().to_owned(),
                target: (!relation.is_external())
                    .then(|| relation.target_partname())
                    .transpose()?,
                target_mode: relation.target_mode(),
            });
        }
    }
    Ok((
        Arc::from(relationships.into_boxed_slice()),
        Arc::from(content_types.into_boxed_slice()),
    ))
}

/// Apply all staged operations once to the immutable original source and
/// retain the resulting OPC source-splice proof for publication.
fn rewrite_story_xml_batch(
    source_xml: &litchi_opc::SourceXmlPart,
    operations: &[BatchOperation],
    limits: litchi_opc::ReadLimits,
) -> Result<litchi_opc::SourceXmlPart> {
    let bytes = source_xml.bytes();
    let drawing = scan_story(bytes, limits)?;
    if operations.is_empty() {
        return Ok(source_xml.clone());
    }
    let mut prospective_len = bytes.len();
    for operation in operations {
        let picture = drawing.picture(operation.selector().picture)?;
        let (range_len, replacement_len) = match operation {
            BatchOperation::Attach {
                relationship_id, ..
            } => {
                let (range, length) =
                    generated_attachment_replacement_bounds(picture, relationship_id)?;
                (range.len()?, length)
            },
            BatchOperation::Detach { .. } => {
                let owner = match picture.svg_owner() {
                    SvgOwnerState::Embedded(owner) => owner,
                    SvgOwnerState::Linked(_)
                    | SvgOwnerState::Ambiguous
                    | SvgOwnerState::Refused => {
                        return Err(unsafe_edit(
                            "detach_svg",
                            "selected picture has a refused SVG owner state",
                        ));
                    },
                    SvgOwnerState::None | SvgOwnerState::Opaque => {
                        return Err(Error::Invalid(
                            "selected picture has no admitted embedded SVG owner".into(),
                        ));
                    },
                };
                (owner.extension_range().len()?, 0)
            },
        };
        if replacement_len as u64 > limits.max_part_bytes() {
            return Err(Error::Invalid(
                "DOCX SVG replacement exceeds the Part limit".into(),
            ));
        }
        prospective_len = prospective_len
            .checked_sub(range_len)
            .and_then(|value| value.checked_add(replacement_len))
            .ok_or_else(|| Error::Invalid("DOCX story output size overflows".into()))?;
    }
    if prospective_len as u64 > limits.max_part_bytes() {
        return Err(Error::Invalid(
            "DOCX story output exceeds the Part limit".into(),
        ));
    }
    let mut publication = source_xml.clone().into_publication()?;
    for operation in operations {
        let picture = drawing.picture(operation.selector().picture)?;
        let (range, fragment) = match operation {
            BatchOperation::Attach {
                relationship_id, ..
            } => {
                if picture.svg_owner().is_refused()
                    || matches!(picture.svg_owner(), SvgOwnerState::Embedded(_))
                {
                    return Err(unsafe_edit(
                        "attach_svg",
                        "selected picture has an existing recognized SVG owner",
                    ));
                }
                generated_attachment_replacement(picture, relationship_id, bytes, limits)?
            },
            BatchOperation::Detach { .. } => {
                let owner = match picture.svg_owner() {
                    SvgOwnerState::Embedded(owner) => owner,
                    SvgOwnerState::Linked(_)
                    | SvgOwnerState::Ambiguous
                    | SvgOwnerState::Refused => {
                        return Err(unsafe_edit(
                            "detach_svg",
                            "selected picture has a refused SVG owner state",
                        ));
                    },
                    SvgOwnerState::None | SvgOwnerState::Opaque => {
                        return Err(Error::Invalid(
                            "selected picture has no admitted embedded SVG owner".into(),
                        ));
                    },
                };
                (owner.extension_range(), Vec::new())
            },
        };
        let expected = range.slice(bytes)?;
        let proof = source_xml.checked_range(range.start..range.end, expected)?;
        let fragment = if fragment.is_empty() {
            AuthoredXmlFragment::empty()
        } else {
            AuthoredXmlFragment::markup(fragment).map_err(Error::from)?
        };
        publication.replace(proof, fragment)?;
    }
    Ok(publication.finish()?)
}

fn generated_attachment_replacement(
    picture: &PictureSource<'_>,
    relationship_id: &str,
    source: &[u8],
    limits: litchi_opc::ReadLimits,
) -> Result<(ByteRange, Vec<u8>)> {
    let (range, replacement_len) =
        generated_attachment_replacement_bounds(picture, relationship_id)?;
    let blip = picture
        .blip_range()
        .ok_or_else(|| Error::Invalid("direct picture has no raster blip range".into()))?;
    let drawing_prefix = blip.prefix();
    let ext_prefix = picture
        .ext_list_range()
        .map_or(drawing_prefix, |range| range.prefix());
    let ext = generated_svg_extension(ext_prefix, relationship_id)?;
    let replacement = if let Some(ext_list) = picture.ext_list_range() {
        if ext_list.close_start().is_some() {
            ext
        } else {
            let old = ext_list.bytes()?;
            generated_ext_list_from_empty(old, ext_prefix, ext)?
        }
    } else if blip.close_start().is_some() {
        generated_ext_list(drawing_prefix, ext)?
    } else {
        let old = blip.bytes()?;
        generated_nonempty_blip(old, drawing_prefix, ext)?
    };
    if replacement.len() != replacement_len {
        return Err(Error::Invalid(
            "DOCX generated SVG replacement length drifted during construction".into(),
        ));
    }
    if source
        .len()
        .checked_sub(range.len()?)
        .and_then(|value| value.checked_add(replacement.len()))
        .is_none_or(|length| length as u64 > limits.max_part_bytes())
    {
        return Err(Error::Invalid(
            "DOCX story output exceeds the Part limit".into(),
        ));
    }
    Ok((range, replacement))
}

fn generated_attachment_replacement_from_layout(
    layout: &PictureLayout,
    relationship_id: &str,
    source: &[u8],
    limits: litchi_opc::ReadLimits,
) -> Result<(ByteRange, Vec<u8>)> {
    let (range, replacement_len) =
        generated_attachment_replacement_bounds_from_layout(layout, relationship_id, source)?;
    let drawing_prefix = layout.blip_prefix.as_ref();
    let ext_prefix = layout.ext_list_prefix.as_deref().unwrap_or(drawing_prefix);
    let ext = generated_svg_extension(ext_prefix, relationship_id)?;
    let replacement = if let Some(ext_list_range) = layout.ext_list_range {
        if layout.ext_list_close_start.is_some() {
            ext
        } else {
            let old = ext_list_range.slice(source)?;
            generated_ext_list_from_empty(old, ext_prefix, ext)?
        }
    } else if layout.blip_close_start.is_some() {
        generated_ext_list(drawing_prefix, ext)?
    } else {
        let old = layout.blip_range.slice(source)?;
        generated_nonempty_blip(old, drawing_prefix, ext)?
    };
    if replacement.len() != replacement_len {
        return Err(Error::Invalid(
            "DOCX generated SVG replacement length drifted during construction".into(),
        ));
    }
    if source
        .len()
        .checked_sub(range.len()?)
        .and_then(|value| value.checked_add(replacement.len()))
        .is_none_or(|length| length as u64 > limits.max_part_bytes())
    {
        return Err(Error::Invalid(
            "DOCX story output exceeds the Part limit".into(),
        ));
    }
    Ok((range, replacement))
}

fn generated_attachment_replacement_bounds_from_layout(
    layout: &PictureLayout,
    relationship_id: &str,
    source: &[u8],
) -> Result<(ByteRange, usize)> {
    validate_relationship_id(relationship_id)?;
    let drawing_prefix = layout.blip_prefix.as_ref();
    let ext_prefix = layout.ext_list_prefix.as_deref().unwrap_or(drawing_prefix);
    let extension_len = generated_svg_extension_len(ext_prefix, relationship_id)?;
    if let Some(ext_list_range) = layout.ext_list_range {
        if let Some(close_start) = layout.ext_list_close_start {
            return Ok((ByteRange::new(close_start, close_start), extension_len));
        }
        let old = ext_list_range.slice(source)?;
        return Ok((
            ext_list_range,
            generated_ext_list_from_empty_len(old, ext_prefix, extension_len)?,
        ));
    }
    let ext_list_len = generated_ext_list_len(drawing_prefix, extension_len)?;
    if let Some(close_start) = layout.blip_close_start {
        return Ok((ByteRange::new(close_start, close_start), ext_list_len));
    }
    let old = layout.blip_range.slice(source)?;
    Ok((
        layout.blip_range,
        generated_nonempty_blip_len(old, drawing_prefix, ext_list_len)?,
    ))
}

fn generated_attachment_replacement_bounds(
    picture: &PictureSource<'_>,
    relationship_id: &str,
) -> Result<(ByteRange, usize)> {
    let blip = picture
        .blip_range()
        .ok_or_else(|| Error::Invalid("direct picture has no raster blip range".into()))?;
    let drawing_prefix = blip.prefix();
    let ext_prefix = picture
        .ext_list_range()
        .map_or(drawing_prefix, |range| range.prefix());
    let extension_len = generated_svg_extension_len(ext_prefix, relationship_id)?;
    if let Some(ext_list) = picture.ext_list_range() {
        if let Some(close_start) = ext_list.close_start() {
            return Ok((ByteRange::new(close_start, close_start), extension_len));
        }
        let old = ext_list.bytes()?;
        return Ok((
            ext_list.range(),
            generated_ext_list_from_empty_len(old, ext_prefix, extension_len)?,
        ));
    }
    let ext_list_len = generated_ext_list_len(drawing_prefix, extension_len)?;
    if let Some(close_start) = blip.close_start() {
        return Ok((ByteRange::new(close_start, close_start), ext_list_len));
    }
    let old = blip.bytes()?;
    Ok((
        blip.range(),
        generated_nonempty_blip_len(old, drawing_prefix, ext_list_len)?,
    ))
}

fn qname_len(prefix: &[u8], local: &[u8]) -> Result<usize> {
    prefix
        .len()
        .checked_add(local.len())
        .and_then(|value| value.checked_add(usize::from(!prefix.is_empty())))
        .ok_or_else(|| Error::Invalid("DOCX generated QName size overflows".into()))
}

fn generated_svg_extension_len(prefix: &[u8], relationship_id: &str) -> Result<usize> {
    validate_relationship_id(relationship_id)?;
    let ext_name = qname_len(prefix, b"ext")?;
    let fixed = b" uri=\"{96DAC541-7B7A-43D3-8B79-37D633B846F1}\"><asvg:svgBlip xmlns:asvg=\"http://schemas.microsoft.com/office/drawing/2016/SVG/main\" xmlns:r=\"";
    1usize
        .checked_add(ext_name)
        .and_then(|value| value.checked_add(fixed.len()))
        .and_then(|value| value.checked_add(TRANSITIONAL_RELATIONSHIP_NAMESPACE.len()))
        .and_then(|value| value.checked_add(b"\" r:embed=\"".len()))
        .and_then(|value| value.checked_add(relationship_id.len()))
        .and_then(|value| value.checked_add(b"\"/></".len()))
        .and_then(|value| value.checked_add(ext_name))
        .and_then(|value| value.checked_add(1))
        .ok_or_else(|| Error::Invalid("DOCX generated SVG extension size overflows".into()))
}

fn generated_ext_list_len(prefix: &[u8], extension_len: usize) -> Result<usize> {
    qname_len(prefix, b"extLst")?
        .checked_mul(2)
        .and_then(|value| value.checked_add(extension_len))
        .and_then(|value| value.checked_add(5))
        .ok_or_else(|| Error::Invalid("DOCX generated extLst size overflows".into()))
}

fn generated_ext_list_from_empty_len(
    old: &[u8],
    prefix: &[u8],
    extension_len: usize,
) -> Result<usize> {
    if !old.ends_with(b"/>") {
        return Err(Error::Invalid("empty extLst is not self-closing".into()));
    }
    let close_len = 3usize
        .checked_add(qname_len(prefix, b"extLst")?)
        .ok_or_else(|| Error::Invalid("DOCX generated extLst close size overflows".into()))?;
    old.len()
        .checked_sub(1)
        .and_then(|value| value.checked_add(extension_len))
        .and_then(|value| value.checked_add(close_len))
        .ok_or_else(|| Error::Invalid("DOCX generated extLst size overflows".into()))
}

fn generated_nonempty_blip_len(old: &[u8], prefix: &[u8], ext_list_len: usize) -> Result<usize> {
    if !old.ends_with(b"/>") {
        return Err(Error::Invalid("empty blip is not self-closing".into()));
    }
    let close_len = 3usize
        .checked_add(qname_len(prefix, b"blip")?)
        .ok_or_else(|| Error::Invalid("DOCX generated blip close size overflows".into()))?;
    old.len()
        .checked_sub(1)
        .and_then(|value| value.checked_add(ext_list_len))
        .and_then(|value| value.checked_add(close_len))
        .ok_or_else(|| Error::Invalid("DOCX generated blip size overflows".into()))
}

fn generated_svg_extension(prefix: &[u8], relationship_id: &str) -> Result<Vec<u8>> {
    let capacity = generated_svg_extension_len(prefix, relationship_id)?;
    let mut output = Vec::new();
    output
        .try_reserve_exact(capacity)
        .map_err(|source| Error::Allocation {
            resource: "DOCX generated SVG extension",
            source,
        })?;
    output.push(b'<');
    push_qname(&mut output, prefix, b"ext");
    output.extend_from_slice(
        b" uri=\"{96DAC541-7B7A-43D3-8B79-37D633B846F1}\"><asvg:svgBlip xmlns:asvg=\"http://schemas.microsoft.com/office/drawing/2016/SVG/main\" xmlns:r=\"",
    );
    output.extend_from_slice(TRANSITIONAL_RELATIONSHIP_NAMESPACE.as_bytes());
    output.extend_from_slice(b"\" r:embed=\"");
    output.extend_from_slice(relationship_id.as_bytes());
    output.extend_from_slice(b"\"/></");
    push_qname(&mut output, prefix, b"ext");
    output.push(b'>');
    Ok(output)
}

fn generated_ext_list(prefix: &[u8], extension: Vec<u8>) -> Result<Vec<u8>> {
    let capacity = generated_ext_list_len(prefix, extension.len())?;
    let mut output = Vec::new();
    output
        .try_reserve_exact(capacity)
        .map_err(|source| Error::Allocation {
            resource: "DOCX generated SVG extension list",
            source,
        })?;
    output.push(b'<');
    push_qname(&mut output, prefix, b"extLst");
    output.push(b'>');
    output.extend_from_slice(&extension);
    output.extend_from_slice(b"</");
    push_qname(&mut output, prefix, b"extLst");
    output.push(b'>');
    Ok(output)
}

fn generated_ext_list_from_empty(old: &[u8], prefix: &[u8], extension: Vec<u8>) -> Result<Vec<u8>> {
    if !old.ends_with(b"/>") {
        return Err(Error::Invalid("empty extLst is not self-closing".into()));
    }
    let capacity = generated_ext_list_from_empty_len(old, prefix, extension.len())?;
    let slash = old
        .iter()
        .rposition(|byte| *byte == b'/')
        .ok_or_else(|| Error::Invalid("empty extLst has no close slash".into()))?;
    let mut output = Vec::new();
    output
        .try_reserve_exact(capacity)
        .map_err(|source| Error::Allocation {
            resource: "DOCX generated expanded extLst",
            source,
        })?;
    output.extend_from_slice(&old[..slash]);
    output.push(b'>');
    output.extend_from_slice(&extension);
    output.extend_from_slice(b"</");
    push_qname(&mut output, prefix, b"extLst");
    output.push(b'>');
    Ok(output)
}

fn generated_nonempty_blip(old: &[u8], prefix: &[u8], extension: Vec<u8>) -> Result<Vec<u8>> {
    if !old.ends_with(b"/>") {
        return Err(Error::Invalid("empty blip is not self-closing".into()));
    }
    let ext_list = generated_ext_list(prefix, extension)?;
    let capacity = generated_nonempty_blip_len(old, prefix, ext_list.len())?;
    let slash = old
        .iter()
        .rposition(|byte| *byte == b'/')
        .ok_or_else(|| Error::Invalid("empty blip has no close slash".into()))?;
    let mut output = Vec::new();
    output
        .try_reserve_exact(capacity)
        .map_err(|source| Error::Allocation {
            resource: "DOCX generated expanded blip",
            source,
        })?;
    output.extend_from_slice(&old[..slash]);
    output.push(b'>');
    output.extend_from_slice(&ext_list);
    output.extend_from_slice(b"</");
    push_qname(&mut output, prefix, b"blip");
    output.push(b'>');
    Ok(output)
}

fn push_qname(output: &mut Vec<u8>, prefix: &[u8], local: &[u8]) {
    if !prefix.is_empty() {
        output.extend_from_slice(prefix);
        output.push(b':');
    }
    output.extend_from_slice(local);
}

fn same_source(left: &SnapshotState, right: &SnapshotState) -> bool {
    same_source_core(left, right)
        && same_bytes(left.xml.as_bytes(), right.xml.as_bytes())
        && (Arc::ptr_eq(
            &left.relationship_fingerprint,
            &right.relationship_fingerprint,
        ) || left.relationship_fingerprint == right.relationship_fingerprint)
        && (Arc::ptr_eq(
            &left.content_type_fingerprint,
            &right.content_type_fingerprint,
        ) || left.content_type_fingerprint == right.content_type_fingerprint)
}

fn same_source_core(left: &SnapshotState, right: &SnapshotState) -> bool {
    same_identity(left, right)
}

fn same_identity(left: &SnapshotState, right: &SnapshotState) -> bool {
    left.part_uri == right.part_uri
        && left.selector == right.selector
        && left.lineage == right.lineage
        && left.source_version == right.source_version
}

fn same_state(left: &SnapshotState, right: &SnapshotState) -> bool {
    same_state_core(left, right) && same_bytes(left.xml.as_bytes(), right.xml.as_bytes())
}

fn same_state_core(left: &SnapshotState, right: &SnapshotState) -> bool {
    left.selector == right.selector
        && left.placement == right.placement
        && left.raster_relationship_id == right.raster_relationship_id
        && left.raster_relationship_type == right.raster_relationship_type
        && left.raster_part_uri == right.raster_part_uri
        && left.raster_content_type == right.raster_content_type
        && same_bytes(
            left.raster_payload.as_bytes(),
            right.raster_payload.as_bytes(),
        )
        && left.owner_state == right.owner_state
        && same_attachment(&left.svg, &right.svg)
}

fn same_batch_source(
    left_batch: &SourceBackedSvgAttachmentBatchSnapshot,
    right_batch: &SourceBackedSvgAttachmentBatchSnapshot,
) -> bool {
    if left_batch.states.len() != right_batch.states.len() {
        return false;
    }
    if !same_bytes(left_batch.story.as_bytes(), right_batch.story.as_bytes()) {
        return false;
    }
    left_batch
        .states
        .iter()
        .zip(right_batch.states.iter())
        .all(|(left, right)| {
            same_source_core(left, right)
                && (Arc::ptr_eq(
                    &left.relationship_fingerprint,
                    &right.relationship_fingerprint,
                ) || left.relationship_fingerprint == right.relationship_fingerprint)
                && (Arc::ptr_eq(
                    &left.content_type_fingerprint,
                    &right.content_type_fingerprint,
                ) || left.content_type_fingerprint == right.content_type_fingerprint)
        })
}

fn same_batch_state(
    left_batch: &SourceBackedSvgAttachmentBatchSnapshot,
    right_batch: &SourceBackedSvgAttachmentBatchSnapshot,
) -> bool {
    if left_batch.states.len() != right_batch.states.len() {
        return false;
    }
    same_bytes(left_batch.story.as_bytes(), right_batch.story.as_bytes())
        && left_batch
            .states
            .iter()
            .zip(right_batch.states.iter())
            .all(|(left, right)| same_state_core(left, right))
}

fn target_state_changed(left: &SnapshotState, right: &SnapshotState) -> bool {
    !same_state(left, right)
}

fn same_attachment(left: &Option<AttachmentState>, right: &Option<AttachmentState>) -> bool {
    match (left, right) {
        (None, None) => true,
        (Some(left), Some(right)) => {
            left.relationship_id == right.relationship_id
                && left.relationship_type == right.relationship_type
                && left.part_uri == right.part_uri
                && same_svg_payload(&left.payload, &right.payload)
        },
        _ => false,
    }
}

fn same_svg_payload(left: &SvgPayload, right: &SvgPayload) -> bool {
    match (left, right) {
        (SvgPayload::Edited(left), SvgPayload::Edited(right)) if Arc::ptr_eq(left, right) => true,
        (SvgPayload::Original(left), SvgPayload::Original(right))
            if same_bytes(left.as_bytes(), right.as_bytes()) =>
        {
            true
        },
        _ => same_bytes(left.as_bytes(), right.as_bytes()),
    }
}

fn same_bytes(left: &[u8], right: &[u8]) -> bool {
    std::ptr::eq(left, right) || left == right
}

fn build_batch_topology_plan(
    package: &litchi_opc::SourceBackedPackage,
    before: &SourceBackedSvgAttachmentBatchSnapshot,
    after: &SourceBackedSvgAttachmentBatchSnapshot,
) -> Result<SourceTopologyPlan> {
    if before.states.len() != after.states.len() || before.states.is_empty() {
        return Err(Error::Invalid(
            "SVG lifecycle batch source shape differs".into(),
        ));
    }
    package.check_execution()?;
    let source_version = package.source_version()?;
    if before
        .states
        .iter()
        .any(|state| state.source_version != source_version)
    {
        return Err(Error::Invalid(
            "SVG lifecycle source changed since the edit was captured".into(),
        ));
    }
    if before
        .states
        .iter()
        .zip(&after.states)
        .any(|(left, right)| !same_identity(left, right))
    {
        return Err(Error::Invalid("SVG lifecycle source state differs".into()));
    }
    let owner = before.states[0].part_uri.clone();
    if after.states.iter().any(|state| state.part_uri != owner) {
        return Err(Error::Invalid(
            "SVG lifecycle batch spans multiple story parts".into(),
        ));
    }
    let story_bytes = after.story.as_bytes();
    let drawing = scan_story(story_bytes, before.states[0].limits)?;
    let mut plan = SourceTopologyPlan::new();
    plan.try_replace_source_xml_part(owner.clone(), after.story.source_xml().clone())?;

    let mut removed_relationships: Vec<(String, PackURI)> = Vec::new();
    for (before, after) in before.states.iter().zip(&after.states) {
        let picture = drawing.picture(after.selector.picture)?;
        if picture.placement() != after.placement
            || picture.raster_relationship_id() != Some(after.raster_relationship_id.as_str())
        {
            return Err(unsafe_edit(
                "SVG lifecycle publication",
                "staged story no longer resolves the selected raster picture",
            ));
        }
        validate_raster_projection(package, &owner, before, after)?;
        let projected_owner = map_owner_state(picture.svg_owner());
        if projected_owner != after.owner_state {
            return Err(unsafe_edit(
                "SVG lifecycle publication",
                "staged story owner state differs from the projected batch",
            ));
        }
        match (&before.svg, &after.svg) {
            (None, Some(svg)) => {
                if !matches!(projected_owner, SvgPictureOwnerState::Embedded) {
                    return Err(unsafe_edit(
                        "SVG lifecycle publication",
                        "staged story does not resolve the newly attached SVG owner",
                    ));
                }
                if picture
                    .svg_owner()
                    .owner()
                    .and_then(|owner| owner.embedded_relationship_id())
                    != Some(svg.relationship_id.as_str())
                {
                    return Err(unsafe_edit(
                        "SVG lifecycle publication",
                        "staged story resolves a different SVG relationship ID",
                    ));
                }
                if !drawing
                    .relationship_references()
                    .iter()
                    .any(|reference| reference.id() == svg.relationship_id)
                {
                    return Err(unsafe_edit(
                        "SVG lifecycle publication",
                        "staged story is missing the new SVG relationship reference",
                    ));
                }
                let payload = match &svg.payload {
                    SvgPayload::Edited(bytes) => Arc::clone(bytes),
                    SvgPayload::Original(_) => {
                        return Err(Error::Invalid(
                            "new SVG attachment does not own a staged payload".into(),
                        ));
                    },
                };
                plan.try_add_part_shared(
                    svg.part_uri.clone(),
                    "image/svg+xml".to_owned(),
                    payload,
                )?;
                plan.try_add_internal_relationship(
                    owner.clone(),
                    svg.relationship_id.clone(),
                    svg.relationship_type.clone(),
                    svg.part_uri.clone(),
                )?;
            },
            (Some(svg), None) => {
                if let Some(id) = picture
                    .svg_owner()
                    .owner()
                    .and_then(|owner| owner.embedded_relationship_id())
                {
                    let reason = if id == svg.relationship_id {
                        "staged story still contains the detached SVG owner"
                    } else {
                        "staged story resolves a different SVG owner"
                    };
                    return Err(unsafe_edit("detach_svg", reason));
                }
                removed_relationships.push((svg.relationship_id.clone(), svg.part_uri.clone()));
            },
            (Some(left), Some(right)) => {
                if !same_attachment(&Some(left.clone()), &Some(right.clone())) {
                    return Err(unsafe_edit(
                        "SVG lifecycle publication",
                        "attach/detach cannot replace an existing SVG owner",
                    ));
                }
                if picture
                    .svg_owner()
                    .owner()
                    .and_then(|owner| owner.embedded_relationship_id())
                    != Some(right.relationship_id.as_str())
                {
                    return Err(unsafe_edit(
                        "SVG lifecycle publication",
                        "staged story resolves a different existing SVG relationship ID",
                    ));
                }
            },
            (None, None) => {},
        }
        if let Some(svg) = &after.svg
            && before.svg.is_some()
        {
            let relation = package
                .part(&owner)?
                .rels()
                .get(&svg.relationship_id)
                .ok_or_else(|| {
                    Error::Invalid(format!(
                        "SVG relationship '{}' is missing from the source graph",
                        svg.relationship_id
                    ))
                })?;
            if relation.is_external()
                || relation.target_mode() != TargetMode::Internal
                || (relation.reltype() != rt::IMAGE && relation.reltype() != rt::STRICT_IMAGE)
                || !relation.target_partname()?.is_equivalent_to(&svg.part_uri)
            {
                return Err(unsafe_edit(
                    "SVG lifecycle publication",
                    "existing SVG relationship closure changed",
                ));
            }
            let svg_part = package.part(&svg.part_uri)?;
            if !is_svg_content_type(svg_part.content_type()) || !svg_part.rels().is_empty() {
                return Err(unsafe_edit(
                    "SVG lifecycle publication",
                    "existing SVG media closure is not an image/svg+xml leaf",
                ));
            }
        }
        if package.source_version()? != source_version {
            return Err(Error::Invalid(
                "SVG lifecycle source changed during validation".into(),
            ));
        }
    }

    let mut seen_removed = Vec::<String>::new();
    let mut scheduled_removals = Vec::<(String, PackURI)>::new();
    for (relationship_id, target) in removed_relationships {
        if seen_removed.iter().any(|value| value == &relationship_id) {
            continue;
        }
        let remaining = drawing
            .relationship_references()
            .iter()
            .any(|reference| reference.id() == relationship_id);
        if remaining {
            continue;
        }
        plan.try_remove_relationship(owner.clone(), relationship_id.clone())?;
        seen_removed.push(relationship_id.clone());
        scheduled_removals.push((relationship_id, target));
    }
    let mut targets = Vec::<PackURI>::new();
    targets
        .try_reserve_exact(scheduled_removals.len())
        .map_err(|source| Error::Allocation {
            resource: "DOCX SVG detached target census",
            source,
        })?;
    for (_, target) in &scheduled_removals {
        if !targets
            .iter()
            .any(|candidate| candidate.is_equivalent_to(target))
        {
            targets.push(target.clone());
        }
    }
    let retained_targets =
        package_relationship_targets_many(package, &owner, &targets, &seen_removed)?;
    let mut removed_parts = Vec::<PackURI>::new();
    for (_, target) in scheduled_removals {
        let target_index = targets
            .iter()
            .position(|candidate| candidate.is_equivalent_to(&target))
            .ok_or_else(|| Error::Invalid("detached SVG target census lost a target".into()))?;
        if retained_targets[target_index] {
            continue;
        }
        if !removed_parts
            .iter()
            .any(|candidate| candidate.is_equivalent_to(&target))
        {
            removed_parts.push(target.clone());
            plan.try_remove_part(target)?;
        }
    }
    package.check_execution()?;
    if package.source_version()? != source_version {
        return Err(Error::Invalid(
            "SVG lifecycle source changed during validation".into(),
        ));
    }
    Ok(plan)
}

fn package_relationship_targets_many(
    package: &litchi_opc::SourceBackedPackage,
    owner: &PackURI,
    targets: &[PackURI],
    removed_ids: &[String],
) -> Result<Vec<bool>> {
    let mut retained = Vec::new();
    retained
        .try_reserve_exact(targets.len())
        .map_err(|source| Error::Allocation {
            resource: "DOCX SVG incoming target census",
            source,
        })?;
    retained.resize(targets.len(), false);
    for relation in package.rels().iter() {
        if relation.is_external() {
            continue;
        }
        let relation_target = relation.target_partname()?;
        for (index, target) in targets.iter().enumerate() {
            if relation_target.is_equivalent_to(target) {
                retained[index] = true;
            }
        }
    }
    for part in package.iter_parts() {
        for relation in part.rels().iter() {
            if owner.is_equivalent_to(part.partname())
                && removed_ids.iter().any(|id| id == relation.r_id())
            {
                continue;
            }
            if relation.is_external() {
                continue;
            }
            let relation_target = relation.target_partname()?;
            for (index, target) in targets.iter().enumerate() {
                if relation_target.is_equivalent_to(target) {
                    retained[index] = true;
                }
            }
        }
    }
    Ok(retained)
}

fn validate_raster_projection(
    package: &litchi_opc::SourceBackedPackage,
    owner: &PackURI,
    before: &SnapshotState,
    after: &SnapshotState,
) -> Result<()> {
    if before.raster_relationship_id != after.raster_relationship_id
        || before.raster_relationship_type != after.raster_relationship_type
        || before.raster_part_uri != after.raster_part_uri
        || before.raster_content_type != after.raster_content_type
        || !same_bytes(
            before.raster_payload.as_bytes(),
            after.raster_payload.as_bytes(),
        )
    {
        return Err(unsafe_edit(
            "SVG lifecycle publication",
            "staged story changed the selected raster fallback",
        ));
    }
    let owner_part = package.part(owner)?;
    let relation = owner_part
        .rels()
        .get(&before.raster_relationship_id)
        .ok_or_else(|| {
            Error::Invalid(format!(
                "raster relationship '{}' is missing from the source graph",
                before.raster_relationship_id
            ))
        })?;
    if relation.is_external()
        || relation.target_mode() != TargetMode::Internal
        || (relation.reltype() != rt::IMAGE && relation.reltype() != rt::STRICT_IMAGE)
        || !relation
            .target_partname()?
            .is_equivalent_to(&before.raster_part_uri)
    {
        return Err(unsafe_edit(
            "SVG lifecycle publication",
            "selected raster relationship closure changed",
        ));
    }
    let raster_part = package.part(&before.raster_part_uri)?;
    if raster_part.content_type() != before.raster_content_type
        || raster_part.content_type() != ct::PNG
        || !same_bytes(
            raster_part.data()?.as_bytes(),
            before.raster_payload.as_bytes(),
        )
    {
        return Err(unsafe_edit(
            "SVG lifecycle publication",
            "selected raster Part closure changed",
        ));
    }
    Ok(())
}

fn validate_svg_payload(limits: &litchi_opc::ReadLimits, length: usize) -> Result<()> {
    validate_svg_payload_length(length)?;
    if length as u64 > limits.max_part_bytes() {
        return Err(Error::Invalid(
            "SVG payload exceeds the package Part limit".into(),
        ));
    }
    Ok(())
}

fn validate_svg_payload_length(length: usize) -> Result<()> {
    if length == 0 {
        return Err(Error::Invalid("SVG payload cannot be empty".into()));
    }
    Ok(())
}

fn validate_svg_part_uri(uri: &PackURI) -> Result<()> {
    if !uri.as_str().starts_with("/word/media/")
        || !uri.as_str().to_ascii_lowercase().ends_with(".svg")
    {
        return Err(Error::Uri(format!(
            "SVG media URI '{}' is outside /word/media or lacks .svg",
            uri.as_str()
        )));
    }
    Ok(())
}

fn validate_relationship_id(id: &str) -> Result<()> {
    if id.is_empty()
        || id.len() > svg_blip::MAX_RELATIONSHIP_ID_BYTES
        || !litchi_ooxml_common::xml::is_ncname(id)
    {
        return Err(Error::Invalid(format!(
            "invalid SVG relationship ID '{id}'"
        )));
    }
    Ok(())
}

fn image_relationship_type(state: &SnapshotState) -> &'static str {
    match state.relationship_dialect {
        RelationshipDialect::Transitional => rt::IMAGE,
        RelationshipDialect::Strict => rt::STRICT_IMAGE,
    }
}

fn is_svg_content_type(content_type: &str) -> bool {
    content_type.eq_ignore_ascii_case("image/svg+xml")
}

fn unsafe_edit(operation: &'static str, reason: &'static str) -> Error {
    Error::UnsafeEdit {
        format: "DOCX",
        operation,
        reason,
    }
}
